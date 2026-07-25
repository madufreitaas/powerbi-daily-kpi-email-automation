<#
.SYNOPSIS
    Consulta um dataset do Power BI Service via API REST, monta um resumo
    gerencial diario em HTML (compativel com Outlook) e envia por e-mail
    via Microsoft Graph.

.DESCRIPTION
    Pipeline pensado para rodar todo dia via Agendador de Tarefas do
    Windows (ou qualquer scheduler). Nao depende do Power BI Desktop nem
    de nenhuma automacao dentro do proprio relatorio - so consulta o
    modelo semantico ja publicado, calcula os indicadores e envia.

    Configuracao via variaveis de ambiente (nada de credencial no codigo):
      PBI_TENANT_ID       - Tenant ID do Azure AD
      PBI_CLIENT_ID       - Client ID do App Registration
      PBI_WORKSPACE_ID    - ID do workspace do Power BI
      PBI_DATASET_ID      - ID do semantic model / dataset
      PBI_SENDER_UPN      - caixa de e-mail que envia o relatorio

    Segredos (Client Secret do Graph, Refresh Token do Power BI) ficam em
    arquivos fora do repositorio, em $env:USERPROFILE\.secrets\, nunca
    versionados. Veja o README para o passo a passo de configuracao.

.NOTES
    Este script foi anonimizado para fins de portfolio. Nomes de empresa,
    e-mails e IDs sao ficticios/placeholders - substitua pelos seus.
#>

param(
    [string[]]$Destinatarios = @("destinatario@suaempresa.com.br")
)

$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURACAO (via variaveis de ambiente - ver README)
# ============================================================
$tenantId    = $env:PBI_TENANT_ID
$clientId    = $env:PBI_CLIENT_ID
$workspaceId = $env:PBI_WORKSPACE_ID
$datasetId   = $env:PBI_DATASET_ID
$senderUpn   = $env:PBI_SENDER_UPN

foreach ($v in @("tenantId","clientId","workspaceId","datasetId","senderUpn")) {
    if ([string]::IsNullOrWhiteSpace((Get-Variable $v -ValueOnly))) {
        throw "Variavel de ambiente nao configurada para '$v'. Veja o README."
    }
}

$secretsDir = "$env:USERPROFILE\.secrets"
function Read-SecretValue($path, $prefix) {
    $raw = (Get-Content $path -Raw).Trim()
    $line = ($raw -split "`n") | Where-Object { $_ -match "^$prefix=" } | Select-Object -First 1
    if ($line) { return ($line -replace "^$prefix=", '').Trim() }
    return $raw
}

# ============================================================
# 1. REGRA DE DATA D-1 UTIL
# ============================================================
# O relatorio nunca usa o dia corrente (dados normalmente incompletos ate
# o fechamento do dia anterior). Se hoje for segunda-feira, usa a
# sexta-feira anterior em vez do domingo, para manter comparacoes justas
# de dias uteis entre o mes atual e o mes anterior.
$today = Get-Date
$cutoff = if ($today.DayOfWeek -eq [DayOfWeek]::Monday) { $today.AddDays(-3) } else { $today.AddDays(-1) }
$cutoffStr = "$($cutoff.Year),$($cutoff.Month),$($cutoff.Day)"
Write-Output "Data de corte (D-1 util): $($cutoff.ToString('yyyy-MM-dd'))"

# ============================================================
# 2. TOKEN POWER BI (autenticacao delegada via refresh token)
# ============================================================
# Por que autenticacao delegada (usuario) em vez de client-credentials
# (service principal) puro: datasets com Row-Level Security (RLS)
# configurada podem negar consultas via API para uma identidade de
# aplicativo puro, mesmo com papel de Admin no workspace - a API de
# Execute Queries aplica RLS de um jeito diferente da experiencia normal
# do Power BI Desktop/Service. A solucao mais simples e confiavel foi
# autenticar como um usuario real (com acesso irrestrito ao dataset),
# via OAuth Device Code Flow, guardando o refresh token para renovar o
# access token sozinho a cada execucao (sem precisar logar de novo).
$refreshToken = Read-SecretValue "$secretsDir\pbi-delegated.env" "REFRESH_TOKEN"
$pbiTokenResp = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body @{
    grant_type    = "refresh_token"
    client_id     = $clientId
    refresh_token = $refreshToken
    scope         = "https://analysis.windows.net/powerbi/api/Dataset.Read.All offline_access"
}
$pbiToken = $pbiTokenResp.access_token
if ($pbiTokenResp.refresh_token) {
    # O refresh token roda em rotacao: cada uso emite um novo, que precisa
    # ser salvo para a proxima execucao.
    "REFRESH_TOKEN=$($pbiTokenResp.refresh_token)" | Set-Content "$secretsDir\pbi-delegated.env"
}
Write-Output "Token Power BI obtido."

function Invoke-PbiQuery($dax) {
    $payload = @{ queries = @(@{ query = $dax }); serializerSettings = @{ includeNulls = $true } } | ConvertTo-Json -Depth 8
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    try {
        # Usamos HttpWebRequest/Response diretamente (nao Invoke-RestMethod) para
        # controlar explicitamente a decodificacao UTF-8 da resposta. O Windows
        # PowerShell 5.1 decodifica tanto o corpo enviado quanto o recebido
        # usando o codepage ANSI do sistema por padrao (mesmo com
        # Content-Type: application/json), corrompendo qualquer acento nos
        # nomes de medidas DAX ou nos dados retornados. Forcar UTF-8 dos dois
        # lados foi a unica forma de resolver isso de vez.
        $req = [System.Net.HttpWebRequest]::Create("https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/executeQueries")
        $req.Method = "POST"
        $req.ContentType = "application/json; charset=utf-8"
        $req.Headers.Add("Authorization", "Bearer $pbiToken")
        $req.ContentLength = $bodyBytes.Length
        $reqStream = $req.GetRequestStream()
        $reqStream.Write($bodyBytes, 0, $bodyBytes.Length)
        $reqStream.Close()

        $resp = $req.GetResponse()
        $respStream = $resp.GetResponseStream()
        $utf8Reader = New-Object System.IO.StreamReader($respStream, [System.Text.Encoding]::UTF8)
        $jsonText = $utf8Reader.ReadToEnd()
        $resp.Close()

        $result = $jsonText | ConvertFrom-Json
        return $result.results[0].tables[0].rows
    }
    catch {
        Write-Host "ERRO na consulta DAX:" -ForegroundColor Red
        Write-Host $dax
        if ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $utf8Reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
                Write-Host $utf8Reader.ReadToEnd() -ForegroundColor Red
            } catch {}
        }
        exit 1
    }
}

# ============================================================
# 3. CONSULTAS DAX
# ============================================================
# Reaproveita as mesmas medidas ja publicadas no modelo semantico (nao
# recria a logica de negocio em codigo) - so define o recorte de datas
# (mes atual ate a data de corte vs. mesmo periodo do mes anterior) e
# agrega por algumas dimensoes de negocio para o detalhamento.
$defineBlock = @"
DEFINE
    VAR CutoffDate = DATE($cutoffStr)
    VAR CurStart = DATE(YEAR(CutoffDate), MONTH(CutoffDate), 1)
    VAR CurEnd = CutoffDate
    VAR PrevStart = EDATE(CurStart, -1)
    VAR PrevEnd = EDATE(CurEnd, -1)
    VAR MonthEnd = EOMONTH(CurStart, 0)
"@

$summaryQuery = @"
$defineBlock
EVALUATE
UNION(
    ROW(
        "Periodo", "Atual",
        "ReceitaBruta", CALCULATE([Receita Bruta (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], CurStart, CurEnd)),
        "Impostos", CALCULATE([Impostos (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], CurStart, CurEnd)),
        "ReceitaLiquida", CALCULATE([Receita Líquida (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], CurStart, CurEnd)),
        "CMV", CALCULATE([CMV (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], CurStart, CurEnd)),
        "LucroBruto", CALCULATE([Lucro Bruto (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], CurStart, CurEnd)),
        "MargemBruta", CALCULATE([Margem Bruta (Realizado)], DATESBETWEEN(Dim_Calendario[Data], CurStart, CurEnd)),
        "MetaMes", CALCULATE([Receita Bruta (Orçado) 2], DATESBETWEEN(Dim_Calendario[Data], CurStart, MonthEnd)),
        "Tendencia", CALCULATE([Tendência 2], DATESBETWEEN(Dim_Calendario[Data], CurStart, CurEnd))
    ),
    ROW(
        "Periodo", "Anterior",
        "ReceitaBruta", CALCULATE([Receita Bruta (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], PrevStart, PrevEnd)),
        "Impostos", CALCULATE([Impostos (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], PrevStart, PrevEnd)),
        "ReceitaLiquida", CALCULATE([Receita Líquida (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], PrevStart, PrevEnd)),
        "CMV", CALCULATE([CMV (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], PrevStart, PrevEnd)),
        "LucroBruto", CALCULATE([Lucro Bruto (Realizado) 2], DATESBETWEEN(Dim_Calendario[Data], PrevStart, PrevEnd)),
        "MargemBruta", CALCULATE([Margem Bruta (Realizado)], DATESBETWEEN(Dim_Calendario[Data], PrevStart, PrevEnd)),
        "MetaMes", CALCULATE([Receita Bruta (Orçado) 2], DATESBETWEEN(Dim_Calendario[Data], PrevStart, EOMONTH(PrevStart, 0))),
        "Tendencia", BLANK()
    )
)
"@

function New-BreakdownQuery($column) {
    return @"
$defineBlock
EVALUATE
CALCULATETABLE(
    SUMMARIZECOLUMNS(
        $column,
        "ReceitaBruta", CALCULATE([Receita Bruta (Realizado) 2]),
        "MargemBruta", CALCULATE([Margem Bruta (Realizado)])
    ),
    DATESBETWEEN(Dim_Calendario[Data], CurStart, CurEnd)
)
ORDER BY [ReceitaBruta] DESC
"@
}

Write-Output "Executando consultas..."
$summaryRows    = Invoke-PbiQuery $summaryQuery
$gerenteRows    = Invoke-PbiQuery (New-BreakdownQuery "Fato_1[Gerente]")
$ufRows         = Invoke-PbiQuery (New-BreakdownQuery "Fato_1[UF]")
$unidadeRows    = Invoke-PbiQuery (New-BreakdownQuery "Fato_1[Unidade]")
$tipoFatRows    = Invoke-PbiQuery (New-BreakdownQuery "Fato_1[Tipo Faturamento]")
Write-Output "Consultas concluidas."

$atual = $summaryRows | Where-Object { $_.'[Periodo]' -eq 'Atual' }
$anterior = $summaryRows | Where-Object { $_.'[Periodo]' -eq 'Anterior' }

# ============================================================
# 4. FUNCOES DE FORMATACAO
# ============================================================
function Format-Reais($valor) {
    $abs = [Math]::Abs($valor)
    return "R$ " + $abs.ToString("N0", [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
}
function Format-ReaisMi($valor) {
    $mi = [Math]::Abs($valor) / 1000000
    return "R$&nbsp;" + $mi.ToString("N2", [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")) + "&nbsp;mi"
}
function Format-Pct($valor, $casas = 1) {
    return ($valor * 100).ToString("N$casas", [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")) + "%"
}
function Get-DeltaPct($atualVal, $anteriorVal) {
    if ($anteriorVal -eq 0) { return 0 }
    return ($atualVal - $anteriorVal) / [Math]::Abs($anteriorVal)
}

# ============================================================
# 5. CALCULO DOS KPIs
# ============================================================
$receitaBrutaAtual = [double]$atual.'[ReceitaBruta]'
$receitaBrutaAnt    = [double]$anterior.'[ReceitaBruta]'
$impostosAtual      = [Math]::Abs([double]$atual.'[Impostos]')
$impostosAnt        = [Math]::Abs([double]$anterior.'[Impostos]')
$receitaLiqAtual    = [double]$atual.'[ReceitaLiquida]'
$receitaLiqAnt      = [double]$anterior.'[ReceitaLiquida]'
$cmvAtual           = [Math]::Abs([double]$atual.'[CMV]')
$cmvAnt             = [Math]::Abs([double]$anterior.'[CMV]')
$lucroBrutoAtual    = [double]$atual.'[LucroBruto]'
$lucroBrutoAnt      = [double]$anterior.'[LucroBruto]'
$margemBrutaAtual   = [double]$atual.'[MargemBruta]'
$margemBrutaAnt     = [double]$anterior.'[MargemBruta]'
$tendencia          = [double]$atual.'[Tendencia]'
$orcadoMesAtual     = [double]$atual.'[MetaMes]'
$orcadoMesAnt       = [double]$anterior.'[MetaMes]'

# Atingimento de Meta = Receita Bruta realizada ate a data de corte / Meta
# orcada do mes inteiro (nao a mesma medida pronta de "Atingimento" do
# modelo, que somava indicadores diferentes e distorcia o percentual -
# replicamos aqui a mesma logica usada no card oficial do relatorio).
$atingimentoAtual = if ($orcadoMesAtual -ne 0) { $receitaBrutaAtual / $orcadoMesAtual } else { 0 }
$atingimentoAnt   = if ($orcadoMesAnt -ne 0) { $receitaBrutaAnt / $orcadoMesAnt } else { 0 }

$impostosPctAtual = if ($receitaBrutaAtual -ne 0) { $impostosAtual / $receitaBrutaAtual } else { 0 }
$impostosPctAnt   = if ($receitaBrutaAnt -ne 0) { $impostosAnt / $receitaBrutaAnt } else { 0 }

$deltaReceitaBruta = Get-DeltaPct $receitaBrutaAtual $receitaBrutaAnt
$deltaImpostosPP    = $impostosPctAtual - $impostosPctAnt
$deltaReceitaLiq   = Get-DeltaPct $receitaLiqAtual $receitaLiqAnt
$deltaCmv          = Get-DeltaPct $cmvAtual $cmvAnt
$deltaLucroBruto   = Get-DeltaPct $lucroBrutoAtual $lucroBrutoAnt
$deltaMargemPP     = $margemBrutaAtual - $margemBrutaAnt
$deltaAtingimentoPP = $atingimentoAtual - $atingimentoAnt

$tendenciaVsOrcado = if ($orcadoMesAtual -ne 0) { ($tendencia - $orcadoMesAtual) / $orcadoMesAtual } else { 0 }

Write-Output "KPIs calculados. Receita Bruta atual: $(Format-ReaisMi $receitaBrutaAtual)"

# ============================================================
# 6. TOP 5 + OUTROS PARA CADA QUEBRA
# ============================================================
function Get-Top5MaisOutros($rows, $colName) {
    $items = @()
    foreach ($r in $rows) {
        $nome = $r.$colName
        if ([string]::IsNullOrWhiteSpace($nome)) { $nome = "Não informado" }
        $items += [PSCustomObject]@{
            Nome = $nome
            Receita = [double]$r.'[ReceitaBruta]'
            Margem  = [double]$r.'[MargemBruta]'
        }
    }
    $top5 = $items | Select-Object -First 5
    $outros = $items | Select-Object -Skip 5
    $outrosReceita = ($outros | Measure-Object -Property Receita -Sum).Sum
    $outrosMargemPonderada = 0
    if ($outrosReceita -ne 0 -and $outros.Count -gt 0) {
        $outrosMargemPonderada = (($outros | ForEach-Object { $_.Receita * $_.Margem } | Measure-Object -Sum).Sum) / $outrosReceita
    }
    return @{ Top5 = $top5; OutrosReceita = $outrosReceita; OutrosMargem = $outrosMargemPonderada; OutrosCount = $outros.Count }
}

$gerenteBd  = Get-Top5MaisOutros $gerenteRows  'Fato_1[Gerente]'
$ufBd       = Get-Top5MaisOutros $ufRows       'Fato_1[UF]'
$unidadeBd  = Get-Top5MaisOutros $unidadeRows  'Fato_1[Unidade]'
$tipoFatBd  = Get-Top5MaisOutros $tipoFatRows  'Fato_1[Tipo Faturamento]'

Write-Output "Quebras por dimensao calculadas."

# ============================================================
# 7. MONTAGEM DO HTML (compativel com Outlook - baseado em tabelas)
# ============================================================
$corBoa  = "#1AAB40"
$corRuim = "#AF2A34"
$corNav  = "#0B4F8A"
$corBg   = "#FBF9F9"
$corBorda = "#DDE3EA"
$corTextoMudo = "#5B6B7C"

function Kpi-Cor($higherIsBetter, $delta) {
    $isUp = $delta -ge 0
    if ($higherIsBetter) { $isGood = $isUp } else { $isGood = -not $isUp }
    if ($isGood) { $cor = $corBoa } else { $cor = $corRuim }
    if ($isUp) { $seta = "&#9650;" } else { $seta = "&#9660;" }
    return @{ cor = $cor; seta = $seta }
}

function Kpi-RowHtml($label, $atualTxt, $antTxt, $deltaTxt, $higherIsBetter, $deltaValor) {
    $c = Kpi-Cor $higherIsBetter $deltaValor
    return @"
<tr>
  <td style="padding:10px 8px;border-bottom:1px solid $corBorda;font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;font-weight:600;color:#142437;">$label</td>
  <td align="right" style="padding:10px 8px;border-bottom:1px solid $corBorda;font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;color:#142437;">$atualTxt</td>
  <td align="right" style="padding:10px 8px;border-bottom:1px solid $corBorda;font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;color:$corTextoMudo;">$antTxt</td>
  <td align="right" style="padding:10px 8px;border-bottom:1px solid $corBorda;font-family:Segoe UI,Arial,sans-serif;font-size:13.5px;font-weight:600;color:$($c.cor);">$($c.seta) $deltaTxt</td>
</tr>
"@
}

function Breakdown-TableHtml($titulo, $bd) {
    $rows = ""
    foreach ($item in $bd.Top5) {
        $rows += @"
<tr>
  <td style="padding:7px 6px;border-bottom:1px solid #EEE;font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#142437;">$($item.Nome)</td>
  <td align="right" style="text-align:right;padding:7px 6px;border-bottom:1px solid #EEE;font-family:Segoe UI,Arial,sans-serif;font-size:13px;font-weight:600;color:#142437;white-space:nowrap;"><div style="width:100%;text-align:right;">$(Format-ReaisMi $item.Receita)</div></td>
  <td align="right" style="text-align:right;padding:7px 6px;border-bottom:1px solid #EEE;font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:$corTextoMudo;"><div style="width:100%;text-align:right;">$(Format-Pct $item.Margem)</div></td>
</tr>
"@
    }
    if ($bd.OutrosCount -gt 0) {
        $rows += @"
<tr>
  <td style="padding:7px 6px;font-family:Segoe UI,Arial,sans-serif;font-size:13px;font-style:italic;color:$corTextoMudo;">Outros ($($bd.OutrosCount))</td>
  <td align="right" style="padding:7px 6px;font-family:Segoe UI,Arial,sans-serif;font-size:13px;font-style:italic;color:$corTextoMudo;white-space:nowrap;">$(Format-ReaisMi $bd.OutrosReceita)</td>
  <td align="right" style="padding:7px 6px;font-family:Segoe UI,Arial,sans-serif;font-size:13px;font-style:italic;color:$corTextoMudo;">$(Format-Pct $bd.OutrosMargem)</td>
</tr>
"@
    }
    return @"
<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
  <tr>
    <td colspan="3" style="font-family:Segoe UI,Arial,sans-serif;font-size:11.5px;letter-spacing:0.03em;text-transform:uppercase;color:$corNav;font-weight:600;padding-bottom:6px;">$titulo</td>
  </tr>
  <tr>
    <td style="text-align:left;font-family:Segoe UI,Arial,sans-serif;font-size:10.5px;text-transform:uppercase;color:$corTextoMudo;border-bottom:2px solid $corBorda;padding:0 6px 6px;">Nome</td>
    <td align="right" style="text-align:right;font-family:Segoe UI,Arial,sans-serif;font-size:10.5px;text-transform:uppercase;color:$corTextoMudo;border-bottom:2px solid $corBorda;padding:0 14px 6px 6px;"><div style="width:100%;text-align:right;">Receita</div></td>
    <td align="right" style="text-align:right;font-family:Segoe UI,Arial,sans-serif;font-size:10.5px;text-transform:uppercase;color:$corTextoMudo;border-bottom:2px solid $corBorda;padding:0 6px 6px;"><div style="width:100%;text-align:right;">Margem</div></td>
  </tr>
  $rows
</table>
"@
}

$ptBr = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")
$mesAtualNome = $ptBr.TextInfo.ToTitleCase(([datetime]$cutoff).ToString("MMMM/yyyy", $ptBr))
$mesAnteriorNome = $ptBr.TextInfo.ToTitleCase(([datetime]$cutoff).AddMonths(-1).ToString("MMMM/yyyy", $ptBr))
$diaCorte = $cutoff.Day

$cReceita = Kpi-Cor $true $deltaReceitaBruta
$cLucro   = Kpi-Cor $true $deltaLucroBruto
$cAting   = Kpi-Cor $true $deltaAtingimentoPP
$cTend    = Kpi-Cor $true $tendenciaVsOrcado

$comparativoRows = ""
$comparativoRows += Kpi-RowHtml "Receita Bruta" (Format-Reais $receitaBrutaAtual) (Format-Reais $receitaBrutaAnt) (Format-Pct $deltaReceitaBruta) $true $deltaReceitaBruta
$comparativoRows += Kpi-RowHtml "Impostos" (Format-Pct $impostosPctAtual) (Format-Pct $impostosPctAnt) ("$(Format-Pct ([Math]::Abs($deltaImpostosPP))) p.p.") $false $deltaImpostosPP
$comparativoRows += Kpi-RowHtml "Receita Líquida" (Format-Reais $receitaLiqAtual) (Format-Reais $receitaLiqAnt) (Format-Pct $deltaReceitaLiq) $true $deltaReceitaLiq
$comparativoRows += Kpi-RowHtml "CMV" (Format-Reais $cmvAtual) (Format-Reais $cmvAnt) (Format-Pct $deltaCmv) $false $deltaCmv
$comparativoRows += Kpi-RowHtml "Lucro Bruto" (Format-Reais $lucroBrutoAtual) (Format-Reais $lucroBrutoAnt) (Format-Pct $deltaLucroBruto) $true $deltaLucroBruto
$comparativoRows += Kpi-RowHtml "Margem Bruta" (Format-Pct $margemBrutaAtual) (Format-Pct $margemBrutaAnt) ("$(Format-Pct ([Math]::Abs($deltaMargemPP))) p.p.") $true $deltaMargemPP
$comparativoRows += Kpi-RowHtml "Atingimento de Meta" (Format-Pct $atingimentoAtual) (Format-Pct $atingimentoAnt) ("$(Format-Pct ([Math]::Abs($deltaAtingimentoPP))) p.p.") $true $deltaAtingimentoPP

$html = @"
<div style="background:$corBg;padding:24px 12px;font-family:Segoe UI,Arial,sans-serif;">
<table width="680" align="center" cellpadding="0" cellspacing="0" style="border-collapse:collapse;background:#ffffff;">

<tr><td style="background:$corNav;padding:24px 28px;">
  <table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-bottom:14px;">
    <tr>
      <td valign="middle" style="padding-right:9px;">
        <table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
          <tr><td width="7" height="7" style="font-size:0;line-height:0;">&nbsp;</td><td width="7" height="7" bgcolor="#ffffff" style="font-size:0;line-height:0;">&nbsp;</td><td width="7" height="7" style="font-size:0;line-height:0;">&nbsp;</td></tr>
          <tr><td width="7" height="7" bgcolor="#ffffff" style="font-size:0;line-height:0;">&nbsp;</td><td width="7" height="7" style="font-size:0;line-height:0;">&nbsp;</td><td width="7" height="7" bgcolor="#ffffff" style="font-size:0;line-height:0;">&nbsp;</td></tr>
          <tr><td width="7" height="7" style="font-size:0;line-height:0;">&nbsp;</td><td width="7" height="7" bgcolor="#ffffff" style="font-size:0;line-height:0;">&nbsp;</td><td width="7" height="7" style="font-size:0;line-height:0;">&nbsp;</td></tr>
        </table>
      </td>
      <td valign="middle">
        <div style="font-family:Segoe UI,Arial,sans-serif;color:#ffffff;font-size:21px;font-weight:700;">Acme</div>
      </td>
    </tr>
  </table>
  <div style="font-family:Segoe UI,Arial,sans-serif;color:#ffffff;font-size:24px;font-weight:700;margin-bottom:6px;">Resumo Gerencial Diário — $mesAtualNome</div>
  <div style="font-family:Segoe UI,Arial,sans-serif;color:#ffffff;font-size:13.5px;opacity:0.85;">Acumulado até $($diaCorte)/$($cutoff.ToString('MM')) (D-1), comparado ao mesmo período de $mesAnteriorNome &middot; Base: Realizado</div>
</td></tr>

<tr><td style="padding:20px 28px 4px;">
<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
<tr>
  <td width="33%" style="padding:14px 12px;border:1px solid $corBorda;">
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;text-transform:uppercase;color:$corTextoMudo;margin-bottom:4px;">Receita Bruta</div>
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:26px;font-weight:700;color:#142437;margin-bottom:6px;">$(Format-ReaisMi $receitaBrutaAtual)</div>
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:12px;font-weight:600;color:$($cReceita.cor);">$($cReceita.seta) $(Format-Pct ([Math]::Abs($deltaReceitaBruta))) vs mês anterior</div>
  </td>
  <td width="2%"></td>
  <td width="33%" style="padding:14px 12px;border:1px solid $corBorda;">
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;text-transform:uppercase;color:$corTextoMudo;margin-bottom:4px;">Lucro Bruto</div>
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:26px;font-weight:700;color:#142437;margin-bottom:6px;">$(Format-ReaisMi $lucroBrutoAtual)</div>
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:12px;font-weight:600;color:$($cLucro.cor);">$($cLucro.seta) $(Format-Pct ([Math]::Abs($deltaLucroBruto))) vs mês anterior</div>
  </td>
  <td width="2%"></td>
  <td width="30%" style="padding:14px 12px;border:1px solid $corBorda;">
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;text-transform:uppercase;color:$corTextoMudo;margin-bottom:4px;">Atingimento de Meta</div>
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:26px;font-weight:700;color:#142437;margin-bottom:6px;">$(Format-Pct $atingimentoAtual)</div>
    <div style="font-family:Segoe UI,Arial,sans-serif;font-size:12px;font-weight:600;color:$($cAting.cor);">$($cAting.seta) $(Format-Pct ([Math]::Abs($deltaAtingimentoPP))) p.p. vs mês anterior</div>
  </td>
</tr>
</table>
</td></tr>

<tr><td style="padding:16px 28px 4px;font-family:Segoe UI,Arial,sans-serif;font-size:15px;color:#142437;">
Bom dia. Segue o resumo do mês corrente, com os principais indicadores acumulados até hoje e a comparação com o mesmo período do mês anterior.
</td></tr>

<tr><td style="padding:16px 28px 4px;">
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:15.5px;font-weight:600;color:#142437;margin-bottom:2px;">Comparativo de indicadores</div>
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:$corTextoMudo;margin-bottom:10px;">Acumulado no mês &middot; base Realizado</div>
<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
<tr>
  <td style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;text-transform:uppercase;color:$corTextoMudo;border-bottom:2px solid $corBorda;padding:0 8px 6px;">Indicador</td>
  <td align="right" style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;text-transform:uppercase;color:$corTextoMudo;border-bottom:2px solid $corBorda;padding:0 8px 6px;">$mesAtualNome (até $diaCorte)</td>
  <td align="right" style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;text-transform:uppercase;color:$corTextoMudo;border-bottom:2px solid $corBorda;padding:0 8px 6px;">$mesAnteriorNome (até $diaCorte)</td>
  <td align="right" style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;text-transform:uppercase;color:$corTextoMudo;border-bottom:2px solid $corBorda;padding:0 8px 6px;">Variação</td>
</tr>
$comparativoRows
</table>
</td></tr>

<tr><td style="padding:16px 28px;">
<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;background:#F2F6FA;">
<tr><td style="padding:16px 18px;border-left:4px solid $corNav;">
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;text-transform:uppercase;color:$corNav;font-weight:700;margin-bottom:6px;">Tendência de fechamento do mês</div>
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14.5px;color:#142437;">No ritmo atual de faturamento por dia útil, a projeção de fechamento de $mesAtualNome é de $(Format-ReaisMi $tendencia), <span style="color:$($cTend.cor);font-weight:600;">$($cTend.seta) $(Format-Pct ([Math]::Abs($tendenciaVsOrcado)))</span> em relação à meta orçada de $(Format-ReaisMi $orcadoMesAtual).</div>
</td></tr>
</table>
</td></tr>

<tr><td style="padding:8px 28px 4px;">
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:17px;font-weight:600;color:#142437;">Detalhamento — $mesAtualNome <span style="font-size:14px;font-weight:400;color:$corTextoMudo;">(até dia $diaCorte)</span></div>
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:$corTextoMudo;margin-bottom:10px;">Receita Bruta e Margem Bruta por Gerente, UF, Unidade e Tipo de Faturamento &middot; top 5 de cada, ordenado por Receita</div>
</td></tr>

<tr><td style="padding:0 28px 20px;">
<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
<tr>
  <td width="48%" valign="top" style="padding:10px;border:1px solid $corBorda;">$(Breakdown-TableHtml "Por Gerente" $gerenteBd)</td>
  <td width="4%"></td>
  <td width="48%" valign="top" style="padding:10px;border:1px solid $corBorda;">$(Breakdown-TableHtml "Por UF" $ufBd)</td>
</tr>
<tr><td colspan="3" style="height:12px;"></td></tr>
<tr>
  <td width="48%" valign="top" style="padding:10px;border:1px solid $corBorda;">$(Breakdown-TableHtml "Por Unidade" $unidadeBd)</td>
  <td width="4%"></td>
  <td width="48%" valign="top" style="padding:10px;border:1px solid $corBorda;">$(Breakdown-TableHtml "Por Tipo de Faturamento" $tipoFatBd)</td>
</tr>
</table>
</td></tr>

<tr><td style="padding:16px 28px 22px;border-top:1px solid $corBorda;font-family:Segoe UI,Arial,sans-serif;font-size:11.5px;color:$corTextoMudo;">
Gerado automaticamente às 08:00, sempre com dados até D-1 (ou até a última sexta-feira, se hoje for segunda-feira).
</td></tr>

</table>
</div>
"@

Write-Output "HTML montado (tamanho: $($html.Length) caracteres)."

$htmlPath = "$env:TEMP\relatorio-gerencial-preview.html"
$html | Set-Content -Path $htmlPath -Encoding UTF8
Write-Output "Preview salvo em: $htmlPath"

# ============================================================
# 8. ENVIO VIA MICROSOFT GRAPH
# ============================================================
$clientSecret = Read-SecretValue "$secretsDir\pbi.env" "CLIENT_SECRET"
$graphTokenResp = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = "https://graph.microsoft.com/.default"
}
$graphToken = $graphTokenResp.access_token
Write-Output "Token Microsoft Graph obtido."

$mailPayload = @{
    message = @{
        subject = "Resumo Gerencial Diário — $mesAtualNome (até dia $diaCorte)"
        body = @{ contentType = "HTML"; content = $html }
        toRecipients = @($Destinatarios | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    }
    saveToSentItems = "true"
} | ConvertTo-Json -Depth 6

Write-Output "Enviando e-mail para: $($Destinatarios -join ', ')"
$mailBodyBytes = [System.Text.Encoding]::UTF8.GetBytes($mailPayload)
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$senderUpn/sendMail" `
    -Method Post -Headers @{ Authorization = "Bearer $graphToken" } -Body $mailBodyBytes -ContentType "application/json; charset=utf-8"
Write-Output "E-mail enviado com sucesso!"
