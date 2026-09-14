#Requires -Version 7.0
<#
.SYNOPSIS
    Audita caixas de e-mail no Exchange Online.
.DESCRIPTION
    Identifica caixas de usuario que NAO estao na lista de colaboradores ativos
    (provaveis demitidos nao convertidos) e caixas compartilhadas fora do padrao.

    Fontes de colaboradores ativos:
      1. colaboradores.csv  (gerado pelo script da Atualizacao de LD)
      2. funcionarios PJs.xlsx  (prestadores de servico)
      3. Funcionarios DT Novo.xlsx  (aba com Nome + Situacao, apenas ativos)

    Cruzamento por e-mail (CSV + DT EMAIL_CORPORATIVO) e por nome normalizado.
    Usa AccountDisabled do Azure AD para classificar contas desabilitadas.
    Inclui WhenCreated nas caixas para priorizar SharedMailboxes fora de padrao.

    MODOS DE AUTENTICACAO:
      1. Interativo : -AdminUPN admin@contoso.com  (abre MFA no browser)
      2. Automatico : sem parametros, le AppId/CertThumbprint/Organization do config.json
                      (requer Setup-Automacao.ps1 executado previamente)

.PARAMETER AdminUPN
    UPN do administrador para autenticacao interativa (abre MFA no browser).
    Opcional se AppId + CertThumbprint + Organization estiverem no config.json.
.PARAMETER CSVPath
    Caminho do CSV com colaboradores ativos. Se nao informado, busca automaticamente.
.EXAMPLE
    .\Auditar-Caixas.ps1 -AdminUPN admin@contoso.com
.EXAMPLE
    .\Auditar-Caixas.ps1
    (modo automatico — le credenciais do config.json)
#>

param(
    [string]$AdminUPN      = "",
    [string]$CSVPath       = "",
    [string]$PrefixoShared = "share.",
    [string]$PlanilhaPJs   = "",
    [string]$PlanilhaDT    = "",
    # Parametros de autenticacao app-only (podem vir do config.json)
    [string]$AppId         = "",
    [string]$CertThumbprint = "",
    [string]$Organization  = ""
)

$ErrorActionPreference = "Continue"

function Wait-ForKey {
    # Em execucao nao-interativa (tarefa agendada/SYSTEM) nao ha console -> nao pausar.
    if (-not [Environment]::UserInteractive) { return }
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  Pressione qualquer tecla para fechar a janela" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    try {
        $null = [System.Console]::ReadKey($true)
    } catch {
        try { Read-Host "Pressione Enter" | Out-Null } catch { }
    }
}

try {

# ===== RESOLVER CSV =====
$pasta = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($CSVPath)) {
    $local = Join-Path $pasta "colaboradores.csv"
    if (Test-Path $local) {
        $CSVPath = $local
    } else {
        $parent = Split-Path $pasta -Parent
        $found  = @(Get-ChildItem -Path $parent -Directory -Filter "Atualiza*" -ErrorAction SilentlyContinue |
                    ForEach-Object { Join-Path $_.FullName "colaboradores.csv" } |
                    Where-Object { Test-Path $_ })
        if ($found.Count -gt 0) { $CSVPath = $found[0] }
    }
}

if ([string]::IsNullOrWhiteSpace($CSVPath) -or -not (Test-Path $CSVPath)) {
    Write-Host "[ERRO] colaboradores.csv nao encontrado." -ForegroundColor Red
    Write-Host ""
    Write-Host "Procurado em:" -ForegroundColor Yellow
    Write-Host "  - $pasta\colaboradores.csv"
    Write-Host "  - Qualquer pasta irma que comece com 'Atualiza'"
    Write-Host ""
    Write-Host "Use: .\Auditar-Caixas.ps1 -AdminUPN ... -CSVPath 'C:\caminho\colaboradores.csv'" -ForegroundColor Yellow
    Wait-ForKey
    exit 1
}

Write-Host "CSV: $CSVPath" -ForegroundColor Gray

# ===== CONFIGURACAO =====
$emailsProtegidos  = @()
$caixasServico     = @()
$padroesServico    = @()
$enviarEmail       = $false
$emailRemetente    = ""
$emailDestinatario = ""
$tenantId          = ""
$validarLogon      = $true
$diasLogonRecente  = 45
$diasTolerAdmissao = 90
$configPath = Join-Path $pasta "config.json"
if (Test-Path $configPath) {
    try {
        $cfg = Get-Content $configPath -Encoding UTF8 -Raw | ConvertFrom-Json
        if ($cfg.EmailsProtegidos)    { $emailsProtegidos = @($cfg.EmailsProtegidos    | ForEach-Object { $_.ToString().ToLower() }) }
        if ($cfg.CaixasServico)       { $caixasServico    = @($cfg.CaixasServico       | ForEach-Object { $_.ToString().ToLower() }) }
        if ($cfg.PadroesCaixaServico) { $padroesServico   = @($cfg.PadroesCaixaServico | ForEach-Object { $_.ToString() }) }
        if ($cfg.PrefixoShared)       { $PrefixoShared    = $cfg.PrefixoShared }
        # Credenciais de automacao (podem ser sobrepostas por parametros da linha de comando)
        if ([string]::IsNullOrWhiteSpace($AppId)          -and $cfg.AppId)          { $AppId          = $cfg.AppId }
        if ([string]::IsNullOrWhiteSpace($CertThumbprint) -and $cfg.CertThumbprint) { $CertThumbprint = $cfg.CertThumbprint }
        if ([string]::IsNullOrWhiteSpace($Organization)   -and $cfg.Organization)   { $Organization   = $cfg.Organization }
        if (-not [string]::IsNullOrWhiteSpace($cfg.TenantId))                       { $tenantId       = $cfg.TenantId }
        # Validacao de ultimo logon (Microsoft Graph)
        if ($null -ne $cfg.ValidarUltimoLogon)     { $validarLogon      = [bool]$cfg.ValidarUltimoLogon }
        if ($cfg.DiasLogonRecente -as [int])       { $diasLogonRecente  = [int]$cfg.DiasLogonRecente }
        if ($cfg.DiasToleranciaAdmissao -as [int]) { $diasTolerAdmissao = [int]$cfg.DiasToleranciaAdmissao }
        # Configuracao de envio de e-mail
        if ($cfg.EnviarEmail -eq $true)                                             { $enviarEmail       = $true }
        if (-not [string]::IsNullOrWhiteSpace($cfg.EmailRemetente))                 { $emailRemetente    = $cfg.EmailRemetente.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($cfg.EmailDestinatario))              { $emailDestinatario = $cfg.EmailDestinatario.Trim() }
    } catch {
        Write-Host "[AVISO] Nao consegui ler config.json: $_" -ForegroundColor Yellow
    }
}

# ===== MODULO IMPORTEXCEL (leitura nativa de .xlsx sem Python) =====
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Instalando modulo ImportExcel..." -ForegroundColor Yellow
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    }
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    } catch {
        Write-Host "[ERRO] Nao foi possivel instalar o modulo ImportExcel automaticamente." -ForegroundColor Red
        Write-Host "       Execute manualmente em um terminal com internet:" -ForegroundColor Yellow
        Write-Host "       [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" -ForegroundColor White
        Write-Host "       Install-Module ImportExcel -Scope CurrentUser -Force" -ForegroundColor White
        Wait-ForKey
        exit 1
    }
}
Import-Module ImportExcel -Force -ErrorAction Stop

# Helper: busca valor de coluna por nome normalizado (case-insensitive, espacos->_)
function Get-ColVal {
    param($Row, [string]$ColNorm)
    foreach ($p in $Row.PSObject.Properties) {
        if (($p.Name.Trim().ToLower() -replace '\s+', '_') -eq $ColNorm) { return [string]$p.Value }
    }
    return ""
}

# Helper: encontra o nome da aba que contem todas as colunas obrigatorias
function Buscar-Aba {
    param([string]$Arquivo, [string[]]$ColunasObrig)
    foreach ($sheet in (Get-ExcelSheetInfo -Path $Arquivo -ErrorAction SilentlyContinue)) {
        try {
            $amostra = @(Import-Excel -Path $Arquivo -WorksheetName $sheet.Name -EndRow 3 -ErrorAction Stop)
            if ($amostra.Count -gt 0) {
                $cols = $amostra[0].PSObject.Properties.Name |
                    ForEach-Object { $_.Trim().ToLower() -replace '\s+', '_' }
                $ok = $true
                foreach ($c in $ColunasObrig) { if ($c -notin $cols) { $ok = $false; break } }
                if ($ok) { return $sheet.Name }
            }
        } catch {}
    }
    return $null
}

# ===== GRAPH API — TOKEN (JWT client assertion, sem modulos extras) =====
function Get-GraphToken {
    param(
        [string]$Tenant,
        [string]$AppId,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert
    )
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    # x5t: thumbprint SHA-1 em base64url
    $thumbBytes = for ($i = 0; $i -lt $Cert.Thumbprint.Length; $i += 2) {
        [byte][Convert]::ToUInt32($Cert.Thumbprint.Substring($i, 2), 16)
    }
    $x5t = ([Convert]::ToBase64String($thumbBytes)) -replace '\+','-' -replace '/','_' -replace '='

    $Encode = {
        param($obj)
        $j = $obj | ConvertTo-Json -Compress
        ([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($j))) -replace '\+','-' -replace '/','_' -replace '='
    }

    $hdr = & $Encode ([ordered]@{ alg = "RS256"; typ = "JWT"; x5t = $x5t })
    $pay = & $Encode ([ordered]@{
        aud = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token"
        iss = $AppId; sub = $AppId
        jti = [Guid]::NewGuid().ToString()
        nbf = $now; exp = $now + 600
    })

    $toSign   = "$hdr.$pay"

    # GetRSAPrivateKey() e metodo de instancia so no .NET Core (PS7). No Windows
    # PowerShell 5.1 e extension method e nao resolve — chamar a classe estatica
    # funciona nos dois.
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Cert)
    if (-not $rsa) {
        throw "Certificado $($Cert.Thumbprint) nao expos a chave privada RSA. Verifique se foi importado com a chave privada e se a conta atual tem permissao de leitura nela."
    }

    $sigBytes = $rsa.SignData(
        [System.Text.Encoding]::UTF8.GetBytes($toSign),
        [Security.Cryptography.HashAlgorithmName]::SHA256,
        [Security.Cryptography.RSASignaturePadding]::Pkcs1
    )
    $sig = ([Convert]::ToBase64String($sigBytes)) -replace '\+','-' -replace '/','_' -replace '='

    $r = Invoke-RestMethod `
        -Method      Post `
        -Uri         "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
            client_id             = $AppId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = "$toSign.$sig"
            scope                 = "https://graph.microsoft.com/.default"
            grant_type            = "client_credentials"
        }
    return $r.access_token
}

# ===== GRAPH API — ENVIO DE E-MAIL (sendMail + CSV anexo) =====
function Enviar-RelatorioEmail {
    param(
        [string]$Token,
        [string]$De,
        [string]$Para,
        [string]$CsvPath,
        [int]$TotalCaixas,
        [int]$TotalAlta,
        [int]$TotalMedia,
        [int]$TotalBaixa,
        [string]$DataExecucao
    )

    $corAlerta  = if ($TotalAlta -gt 0) { "#c0392b" } else { "#27ae60" }
    $iconAlerta = if ($TotalAlta -gt 0) { "⚠️" }      else { "✅" }

    $htmlBody = @"
<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222;margin:20px;">
<h2 style="color:#2c3e50;margin-bottom:4px;">📋 Auditoria de Caixas Exchange</h2>
<p style="color:#666;margin-top:0;">Executada em $DataExecucao</p>
<table style="border-collapse:collapse;min-width:340px;margin-top:12px;">
  <tr>
    <td style="padding:8px 14px;background:#ecf0f1;font-weight:bold;">Total de caixas analisadas</td>
    <td style="padding:8px 16px;">$TotalCaixas</td>
  </tr>
  <tr>
    <td style="padding:8px 14px;background:#fadbd8;color:$corAlerta;font-weight:bold;">$iconAlerta Prioridade Alta</td>
    <td style="padding:8px 16px;color:$corAlerta;font-weight:bold;">$TotalAlta</td>
  </tr>
  <tr>
    <td style="padding:8px 14px;background:#fef9e7;font-weight:bold;">⚡ Prioridade Média</td>
    <td style="padding:8px 16px;">$TotalMedia</td>
  </tr>
  <tr>
    <td style="padding:8px 14px;background:#eafaf1;font-weight:bold;">✔ Prioridade Baixa</td>
    <td style="padding:8px 16px;">$TotalBaixa</td>
  </tr>
</table>
<br>
<p>O relatório completo está anexado neste e-mail (CSV).</p>
<hr style="border:none;border-top:1px solid #ddd;margin-top:24px;">
<p style="font-size:11px;color:#aaa;">Enviado automaticamente — Auditoria de Caixas Exchange</p>
</body></html>
"@

    $csvBytes = [System.IO.File]::ReadAllBytes($CsvPath)
    $csvB64   = [Convert]::ToBase64String($csvBytes)
    $csvNome  = [System.IO.Path]::GetFileName($CsvPath)

    $payload = @{
        message = @{
            subject      = "Auditoria de Caixas — $DataExecucao — Alta: $TotalAlta | Media: $TotalMedia"
            body         = @{ contentType = "HTML"; content = $htmlBody }
            toRecipients = @(@{ emailAddress = @{ address = $Para } })
            attachments  = @(@{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name          = $csvNome
                contentType   = "text/csv"
                contentBytes  = $csvB64
            })
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 10

    Invoke-RestMethod `
        -Method  Post `
        -Uri     "https://graph.microsoft.com/v1.0/users/$De/sendMail" `
        -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" } `
        -Body    $payload `
        -ErrorAction Stop | Out-Null
}

function Test-CaixaServico {
    param([string]$Email, [string]$Prefixo)
    if ($Email -in $caixasServico) { return $true }
    foreach ($p in $padroesServico) { if ($Prefixo -match $p) { return $true } }
    return $false
}

function Normalizar-Nome {
    param([string]$Texto)
    if ([string]::IsNullOrWhiteSpace($Texto)) { return "" }
    $t  = $Texto.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = [System.Text.StringBuilder]::new()
    foreach ($c in $t.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne
            [System.Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
    }
    $r = $sb.ToString().ToUpper().Trim()
    $r = $r -replace '^\s*SHARE\s*[-:]\s*', ''
    $r = $r -replace '^\s*VISOTECH\s*[-:]\s*', ''
    $r = $r -replace '\s+', ' '
    return $r.Trim()
}

# Retorna dois indices do historico DT (nativo PS, sem Python):
#   porNome  : nome_upper  -> {situacao, dataDemissao, causaDemissao}
#   porEmail : email_lower -> {situacao, nome, dataDemissao, causaDemissao}
# Quando nome/email aparece mais de uma vez, prioriza o registro ativo (recontratacao).
function Ler-HistoricoDT {
    param([string]$Arquivo)
    if (-not (Test-Path $Arquivo)) { return @{ porNome = @{}; porEmail = @{} } }

    $aba = Buscar-Aba -Arquivo $Arquivo -ColunasObrig @('nome', 'situacao')
    if (-not $aba) {
        Write-Host "[AVISO] Nenhuma aba com 'nome'+'situacao' encontrada em: $Arquivo" -ForegroundColor Yellow
        return @{ porNome = @{}; porEmail = @{} }
    }

    $porNome  = @{}
    $porEmail = @{}
    try {
        foreach ($r in @(Import-Excel -Path $Arquivo -WorksheetName $aba -ErrorAction Stop)) {
            $nome  = (Get-ColVal $r 'nome').Trim()
            $sit   = (Get-ColVal $r 'situacao').Trim()
            $dtDem = (Get-ColVal $r 'data_demissao').Trim()
            $causa = (Get-ColVal $r 'causa_demissao').Trim()
            $emC   = (Get-ColVal $r 'email_corporativo').Trim()
            $emP   = (Get-ColVal $r 'email').Trim()

            if (-not $nome -or $nome.ToLower() -eq 'nan') { continue }
            if (-not $sit)  { $sit = 'Desconhecido' }
            if ($dtDem -eq 'nan' -or $dtDem -match '^0\d/\d') { $dtDem = $dtDem -replace '^nan$','' }
            if ($causa -eq 'nan') { $causa = '' }

            $key    = $nome.ToUpper()
            $entryN = @{ situacao = $sit; dataDemissao = $dtDem; causaDemissao = $causa }
            $ex     = $porNome[$key]
            if (-not $ex -or ($ex.situacao -like 'Demit*' -and $sit -like 'Trabalh*')) {
                $porNome[$key] = $entryN
            }

            foreach ($em in @($emC, $emP)) {
                if ($em -and $em.ToLower() -ne 'nan' -and $em -like '*@*') {
                    $emL    = $em.ToLower()
                    $entryE = @{ situacao = $sit; nome = $nome; dataDemissao = $dtDem; causaDemissao = $causa }
                    $exE    = $porEmail[$emL]
                    if (-not $exE -or ($exE.situacao -like 'Demit*' -and $sit -like 'Trabalh*')) {
                        $porEmail[$emL] = $entryE
                    }
                }
            }
        }
    } catch {
        Write-Host "[AVISO] Erro ao ler historico DT: $_" -ForegroundColor Yellow
    }
    return @{ porNome = $porNome; porEmail = $porEmail }
}

# Retorna nomes e e-mails dos funcionarios com situacao ATIVA no DT (nativo PS).
function Ler-AtivosDT {
    param([string]$Arquivo)
    if (-not (Test-Path $Arquivo)) { return @{ nomes = @(); emails = @() } }

    $ATIVOS  = @('trabalh', 'em admiss', 'afastad', 'em licen')
    $aba     = Buscar-Aba -Arquivo $Arquivo -ColunasObrig @('nome', 'situacao')
    if (-not $aba) { return @{ nomes = @(); emails = @() } }

    $nomes  = [System.Collections.Generic.List[string]]::new()
    $emails = [System.Collections.Generic.List[string]]::new()
    try {
        foreach ($r in @(Import-Excel -Path $Arquivo -WorksheetName $aba -ErrorAction Stop)) {
            $sit = (Get-ColVal $r 'situacao').Trim().ToLower()
            if (-not ($ATIVOS | Where-Object { $sit.StartsWith($_) })) { continue }

            $nome = (Get-ColVal $r 'nome').Trim()
            if ($nome -and $nome.ToLower() -ne 'nan') { $nomes.Add($nome) }

            foreach ($col in @('email_corporativo', 'email')) {
                $em = (Get-ColVal $r $col).Trim()
                if ($em -and $em.ToLower() -ne 'nan' -and $em -like '*@*') { $emails.Add($em.ToLower()) }
            }
        }
    } catch {
        Write-Host "[AVISO] Erro ao ler ativos DT: $_" -ForegroundColor Yellow
    }
    return @{ nomes = @($nomes); emails = @($emails) }
}

# Le nomes de uma coluna de planilha xlsx (para PJs) — nativo PS.
function Ler-NomesPlanilha {
    param([string]$Arquivo, [string]$ColunaNome, [int]$LinhaHeader = 0)
    if (-not (Test-Path $Arquivo)) { return @() }

    $nomes = [System.Collections.Generic.List[string]]::new()
    try {
        # StartRow: LinhaHeader 0 → header na linha 1 → StartRow 1
        $startRow = $LinhaHeader + 1
        $linhas   = @(Import-Excel -Path $Arquivo -StartRow $startRow -ErrorAction Stop)
        if ($linhas.Count -eq 0) { return @() }

        # Encontrar coluna case-insensitive (match exato, depois prefixo)
        $colKey = $linhas[0].PSObject.Properties.Name |
            Where-Object { $_.Trim().ToLower() -eq $ColunaNome.ToLower() } |
            Select-Object -First 1
        if (-not $colKey) {
            $colKey = $linhas[0].PSObject.Properties.Name |
                Where-Object { $_.Trim().ToLower().StartsWith($ColunaNome.ToLower()) } |
                Select-Object -First 1
        }
        if (-not $colKey) {
            Write-Host "[AVISO] Coluna '$ColunaNome' nao encontrada em: $Arquivo" -ForegroundColor Yellow
            return @()
        }

        foreach ($r in $linhas) {
            $v = ([string]$r.$colKey).Trim()
            if ($v -and $v.ToLower() -ne 'nan' -and $v.ToLower() -ne $ColunaNome.ToLower()) {
                $nomes.Add($v)
            }
        }
    } catch {
        Write-Host "[AVISO] Erro ao ler '$Arquivo': $_" -ForegroundColor Yellow
    }
    return @($nomes)
}

# ===== MODULO EXCHANGE =====
$eomMinVersion = [Version]'3.0.0'
$eomInstalado  = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
    Sort-Object Version -Descending | Select-Object -First 1
if (-not $eomInstalado -or $eomInstalado.Version -lt $eomMinVersion) {
    Write-Host "Instalando/atualizando modulo ExchangeOnlineManagement (>= $eomMinVersion)..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -MinimumVersion $eomMinVersion -Scope CurrentUser -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement -Force

# ===== CONECTAR =====
Write-Host ""
Write-Host "=== AUDITORIA DE CAIXAS ===" -ForegroundColor Cyan
Write-Host ""

$modoAutomatico = $AppId -and $CertThumbprint -and $Organization

if ($modoAutomatico) {
    Write-Host "Modo: AUTOMATICO (app-only, sem MFA)" -ForegroundColor Green
    Write-Host "App: $AppId | Org: $Organization" -ForegroundColor Gray
    Write-Host ""
    try {
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertThumbprint `
            -Organization $Organization -ShowBanner:$false -ErrorAction Stop
    } catch {
        Write-Host "[ERRO] Falha ao conectar (modo automatico):" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host ""
        Write-Host "Verifique: AppId, CertThumbprint e Organization no config.json" -ForegroundColor Yellow
        Write-Host "Execute Setup-Automacao.ps1 para reconfigurar." -ForegroundColor Yellow
        Wait-ForKey
        exit 1
    }
} elseif (-not [string]::IsNullOrWhiteSpace($AdminUPN)) {
    Write-Host "Modo: INTERATIVO (MFA)" -ForegroundColor Cyan
    Write-Host "Conectando como $AdminUPN..." -ForegroundColor Gray
    Write-Host ""
    try {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -DisableWAM -ShowBanner:$false -ErrorAction Stop
    } catch {
        Write-Host "[ERRO] Falha ao conectar ao Exchange Online:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Wait-ForKey
        exit 1
    }
} else {
    Write-Host "[ERRO] Nenhuma credencial configurada." -ForegroundColor Red
    Write-Host ""
    Write-Host "Use uma das opcoes abaixo:" -ForegroundColor Yellow
    Write-Host "  1. Interativo : .\Auditar-Caixas.ps1 -AdminUPN admin@contoso.com" -ForegroundColor Yellow
    Write-Host "  2. Automatico : Execute .\Setup-Automacao.ps1 para configurar" -ForegroundColor Yellow
    Wait-ForKey
    exit 1
}

# ===== CARREGAR COLABORADORES (CSV) =====
$csvDados = Import-Csv -Path $CSVPath -Encoding UTF8 |
    Where-Object { $_.Status -eq 'Ativo' -and -not [string]::IsNullOrWhiteSpace($_.Email) }

# Validar schema do CSV
$colsCSV    = @(if ($csvDados.Count -gt 0) { $csvDados[0].PSObject.Properties.Name } else { @() })
$colsObrig  = @('Status', 'Email', 'Nome')
$faltandoCSV = $colsObrig | Where-Object { $_ -notin $colsCSV }
if ($faltandoCSV.Count -gt 0) {
    Write-Host "[ERRO] colaboradores.csv faltando colunas: $($faltandoCSV -join ', ')" -ForegroundColor Red
    Write-Host "       Colunas encontradas: $($colsCSV -join ', ')" -ForegroundColor Yellow
    Wait-ForKey; exit 1
}

$colaboradoresAtivos = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]($csvDados | Select-Object -ExpandProperty Email | ForEach-Object { $_.ToString().ToLower() })
)
$nomesAtivosCSV = @($csvDados | Select-Object -ExpandProperty Nome |
    ForEach-Object { Normalizar-Nome $_ } | Where-Object { $_ })

Write-Host "Colaboradores ativos no CSV:    $($colaboradoresAtivos.Count)" -ForegroundColor White

# ===== CARREGAR PJs =====
if ([string]::IsNullOrWhiteSpace($PlanilhaPJs)) { $PlanilhaPJs = Join-Path $pasta "funcionarios PJs.xlsx" }
$nomesPJs = @()
if (Test-Path $PlanilhaPJs) {
    $nomesPJs = @(Ler-NomesPlanilha -Arquivo $PlanilhaPJs -ColunaNome "Prestador" -LinhaHeader 0 |
                    ForEach-Object { Normalizar-Nome $_ } | Where-Object { $_ })
    Write-Host "PJs ativos na planilha:         $($nomesPJs.Count)" -ForegroundColor White
} else {
    Write-Host "[AVISO] funcionarios PJs.xlsx nao encontrado: $PlanilhaPJs" -ForegroundColor Yellow
}

# ===== CARREGAR DT (ativos) =====
if ([string]::IsNullOrWhiteSpace($PlanilhaDT)) { $PlanilhaDT = Join-Path $pasta "Funcionarios DT Novo.xlsx" }

$nomesDT   = @()
$emailsDT  = [System.Collections.Generic.HashSet[string]]::new()
if (Test-Path $PlanilhaDT) {
    Write-Host "Lendo ativos do DT..." -ForegroundColor Gray
    $ativosDT = Ler-AtivosDT -Arquivo $PlanilhaDT
    $nomesDT  = @($ativosDT.nomes | ForEach-Object { Normalizar-Nome $_ } | Where-Object { $_ })
    foreach ($em in $ativosDT.emails) { if ($em) { [void]$emailsDT.Add($em.ToLower()) } }
    Write-Host "Funcionarios DT ativos (nomes):  $($nomesDT.Count)" -ForegroundColor White
    Write-Host "Funcionarios DT ativos (e-mails): $($emailsDT.Count)" -ForegroundColor White
} else {
    Write-Host "[AVISO] Funcionarios DT Novo.xlsx nao encontrado: $PlanilhaDT" -ForegroundColor Yellow
}

# Conjunto unificado de nomes ativos (HashSet para lookup O(1))
$nomesAtivos = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]($nomesAtivosCSV + $nomesPJs + $nomesDT)
)
# Sets separados por fonte — usados para registrar FonteConfirmacao
$nomesCSVSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$nomesAtivosCSV)
$nomesPJsSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$nomesPJs)
$nomesDTSet  = [System.Collections.Generic.HashSet[string]]::new([string[]]$nomesDT)
Write-Host "Total nomes ativos unificados:   $($nomesAtivos.Count)" -ForegroundColor White

# ===== HISTORICO DT =====
$historicoDT_porNome  = @{}
$historicoDT_porEmail = @{}
if (Test-Path $PlanilhaDT) {
    Write-Host "Carregando historico DT completo..." -ForegroundColor Gray
    $histRaw = Ler-HistoricoDT -Arquivo $PlanilhaDT

    foreach ($k in $histRaw.porNome.Keys) {
        $kn = Normalizar-Nome $k
        if (-not $kn) { continue }
        $v  = $histRaw.porNome[$k]
        $ex = $historicoDT_porNome[$kn]
        if (-not $ex -or ($ex.situacao -like 'Demit*' -and $v.situacao -like 'Trabalh*')) {
            $historicoDT_porNome[$kn] = $v
        }
    }
    foreach ($k in $histRaw.porEmail.Keys) {
        $v  = $histRaw.porEmail[$k]
        $ex = $historicoDT_porEmail[$k]
        if (-not $ex -or ($ex.situacao -like 'Demit*' -and $v.situacao -like 'Trabalh*')) {
            $historicoDT_porEmail[$k] = $v
        }
    }
    Write-Host "Historico DT: $($historicoDT_porNome.Count) nomes, $($historicoDT_porEmail.Count) e-mails" -ForegroundColor White
    if ($historicoDT_porNome.Count -eq 0) {
        Write-Host ""
        Write-Host "================================================" -ForegroundColor Red
        Write-Host "  [AVISO] Historico DT carregou 0 registros!" -ForegroundColor Red
        Write-Host "  Verifique se a planilha '$PlanilhaDT'" -ForegroundColor Red
        Write-Host "  existe e tem aba com colunas 'nome' e 'situacao'." -ForegroundColor Red
        Write-Host "  Demitidos NAO serao confirmados automaticamente." -ForegroundColor Red
        Write-Host "================================================" -ForegroundColor Red
        Write-Host ""
    }
}

# ===== BUSCAR CAIXAS =====
Write-Host "Buscando caixas de usuario..." -ForegroundColor Gray
$caixasUsuario = @(Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited |
    Select-Object DisplayName, PrimarySmtpAddress, WhenCreated)

Write-Host "Buscando caixas compartilhadas..." -ForegroundColor Gray
$caixasShared = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
    Select-Object DisplayName, PrimarySmtpAddress, WhenCreated)

Write-Host "UserMailbox:    $($caixasUsuario.Count)"
Write-Host "SharedMailbox:  $($caixasShared.Count)"

# ===== LICENCAS + CONTAS DESABILITADAS (Get-User traz AccountDisabled) =====
Write-Host "Buscando usuarios Azure AD (licencas + status de conta)..." -ForegroundColor Gray
$licencas      = @{}
$desabilitados = [System.Collections.Generic.HashSet[string]]::new()
try {
    $users = @(Get-User -ResultSize Unlimited -ErrorAction Stop |
        Select-Object UserPrincipalName, WindowsEmailAddress,
            @{N='SKU';E={[bool]$_.SKUAssigned}},
            AccountDisabled)
    foreach ($u in $users) {
        if ($u.UserPrincipalName) {
            $k = $u.UserPrincipalName.ToString().ToLower()
            $licencas[$k] = $u.SKU
            if ($u.AccountDisabled) { [void]$desabilitados.Add($k) }
        }
        if ($u.WindowsEmailAddress) {
            $k = $u.WindowsEmailAddress.ToString().ToLower()
            $licencas[$k] = $u.SKU
            if ($u.AccountDisabled) { [void]$desabilitados.Add($k) }
        }
    }
    Write-Host "Usuarios Azure AD: $($users.Count)  |  Contas desabilitadas: $($desabilitados.Count)" -ForegroundColor White
} catch {
    Write-Host "[AVISO] Nao consegui ler usuarios AD: $($_.Exception.Message)" -ForegroundColor Yellow
}
Write-Host ""

# ===== ULTIMO LOGON (Microsoft Graph / signInActivity) =====
# Sem este dado, qualquer conta ausente do CSV do RH virava prioridade Alta — inclusive
# gente trabalhando normalmente (ex.: CLT que virou PJ). O logon e o desempate.
$ultimoLogon = @{}
$logonOK     = $false
if ($validarLogon) {
    if (-not $modoAutomatico) {
        Write-Host "[AVISO] Validacao de logon exige modo automatico (app-only). Pulando." -ForegroundColor Yellow
    } elseif (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        Write-Host "[AVISO] Modulo Microsoft.Graph.Authentication ausente. Pulando validacao de logon." -ForegroundColor Yellow
        Write-Host "        Instale com: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -ForegroundColor Gray
    } else {
        Write-Host "Buscando ultimo logon (Microsoft Graph)..." -ForegroundColor Gray
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            $tenantGraph = if (-not [string]::IsNullOrWhiteSpace($tenantId)) { $tenantId } else { $Organization }
            Connect-MgGraph -ClientId $AppId -TenantId $tenantGraph `
                -CertificateThumbprint $CertThumbprint -NoWelcome -ErrorAction Stop

            # NAO usar GET por chave com UPN: o Graph exige GUID e responde 400.
            # O endpoint de lista aceita UPN e e o unico que devolve signInActivity em lote.
            $uri = 'https://graph.microsoft.com/v1.0/users?$select=userPrincipalName,signInActivity&$top=999'
            $paginas = 0
            while ($uri) {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                foreach ($u in $resp.value) {
                    if (-not $u.userPrincipalName) { continue }
                    $sa = $u.signInActivity
                    if (-not $sa) { continue }
                    $datas = @(
                        $sa.lastSignInDateTime
                        $sa.lastNonInteractiveSignInDateTime
                        $sa.lastSuccessfulSignInDateTime
                    ) | Where-Object { $_ } | ForEach-Object { [datetime]$_ }
                    if ($datas) {
                        $ultimoLogon[$u.userPrincipalName.ToString().ToLower()] =
                            ($datas | Sort-Object -Descending | Select-Object -First 1)
                    }
                }
                $paginas++
                $uri = $resp.'@odata.nextLink'
            }
            $logonOK = $true
            Write-Host "Logons coletados: $($ultimoLogon.Count) contas ($paginas paginas)" -ForegroundColor White
        } catch {
            Write-Host "[AVISO] Nao consegui ler signInActivity: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "        Confirme a permissao AuditLog.Read.All no App do Azure AD." -ForegroundColor Gray
            Write-Host "        A auditoria continua, mas SEM desempate por logon." -ForegroundColor Gray
        } finally {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        }
    }
}
Write-Host ""

# Busca info do historico DT: tenta pelo e-mail primeiro, depois pelo nome.
function Get-InfoHistDT {
    param([string]$Email, [string]$NomeNorm)
    if ($Email -and $historicoDT_porEmail.ContainsKey($Email)) { return $historicoDT_porEmail[$Email] }
    if ($NomeNorm -and $historicoDT_porNome.ContainsKey($NomeNorm)) { return $historicoDT_porNome[$NomeNorm] }
    return $null
}

# ===== ANALISE =====
$relatorio = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($cx in $caixasUsuario) {
    $email    = $cx.PrimarySmtpAddress.ToString().ToLower()
    $prefixo  = $email.Split('@')[0].ToLower()
    $display  = [string]$cx.DisplayName
    $nomeNorm = Normalizar-Nome $display
    $ehShareNoNome = $display -match '^\s*SHARE\s*[-:]'

    $dtLogon    = if ($ultimoLogon.ContainsKey($email)) { $ultimoLogon[$email] } else { $null }
    $diasLogon  = if ($dtLogon) { [int]((Get-Date) - $dtLogon).TotalDays } else { $null }
    $diasDesdeCriacao = if ($cx.WhenCreated) { [int]((Get-Date) - $cx.WhenCreated).TotalDays } else { $null }

    $fonte = ""
    if ($email -in $emailsProtegidos) {
        $situacao = "Protegido (OK)"; $prioridade = "Baixa"; $fonte = "Protegido"

    } elseif ($colaboradoresAtivos.Contains($email)) {
        $situacao = "Ativo (e-mail no CSV)"; $prioridade = "Baixa"; $fonte = "CSV-email"

    } elseif (Test-CaixaServico -Email $email -Prefixo $prefixo) {
        $situacao = "Caixa de servico/departamento (OK)"; $prioridade = "Baixa"; $fonte = "Config"

    } elseif ($emailsDT.Contains($email)) {
        $situacao = "Ativo (e-mail no DT)"; $prioridade = "Baixa"; $fonte = "DT-email"

    } elseif ($nomeNorm -and $nomesAtivos.Contains($nomeNorm)) {
        if     ($nomesCSVSet.Contains($nomeNorm)) { $fonte = "CSV-nome";  $src = "CSV" }
        elseif ($nomesPJsSet.Contains($nomeNorm)) { $fonte = "PJ-nome";   $src = "PJs" }
        else                                      { $fonte = "DT-nome";   $src = "DT" }
        $situacao = "Ativo (nome em $src — confirmar e-mail)"; $prioridade = "Baixa"

    } elseif ($ehShareNoNome) {
        $situacao = "UserMailbox com nome SHARE- (converter para SharedMailbox)"; $prioridade = "Media"; $fonte = ""

    } elseif ($prefixo.StartsWith($PrefixoShared)) {
        $situacao = "UserMailbox com prefixo share. (converter para SharedMailbox)"; $prioridade = "Media"; $fonte = ""

    } elseif ($desabilitados.Contains($email)) {
        $situacao = "NAO ESTA NO CSV + conta desabilitada (provavelmente demitido)"
        $prioridade = "Alta"; $fonte = ""

    } elseif ($null -ne $diasLogon -and $diasLogon -le $diasLogonRecente) {
        # Conta EM USO. O erro nao esta na conta, esta no cadastro do RH.
        # Nao e candidata a bloqueio — e divergencia cadastral para o RH resolver.
        $situacao   = "EM USO (logou ha $diasLogon dias) mas fora do CSV - divergencia cadastral, tratar com RH"
        $prioridade = "Media"
        $fonte      = "Logon"

    } elseif ($null -ne $diasDesdeCriacao -and $diasDesdeCriacao -le $diasTolerAdmissao) {
        # Admissao recente: o CSV do RH ainda nao alcancou. Nao e demissao.
        $situacao   = "Admissao recente (caixa criada ha $diasDesdeCriacao dias) - aguardando CSV do RH"
        $prioridade = "Baixa"
        $fonte      = "Admissao recente"

    } elseif ($logonOK -and $null -eq $dtLogon) {
        $situacao   = "NAO ESTA NO CSV + sem logon nos ultimos 180 dias - candidata a bloqueio"
        $prioridade = "Alta"; $fonte = "Logon"

    } elseif ($null -ne $diasLogon) {
        $situacao   = "NAO ESTA NO CSV + inativa ha $diasLogon dias - candidata a bloqueio"
        $prioridade = "Alta"; $fonte = "Logon"

    } else {
        # Sem dado de logon (Graph indisponivel): mantem o comportamento antigo,
        # mas deixa explicito que a classificacao nao foi validada.
        $situacao   = "NAO ESTA NO CSV - verificar se foi demitido (SEM validacao de logon)"
        $prioridade = "Alta"; $fonte = ""
    }

    $hist     = Get-InfoHistDT -Email $email -NomeNorm $nomeNorm
    $sitDT    = if ($hist) { $hist.situacao }      else { "" }
    $datDem   = if ($hist) { $hist.dataDemissao }  else { "" }
    $causaDem = if ($hist) { $hist.causaDemissao } else { "" }
    $lic      = if ($licencas.ContainsKey($email)) { if ($licencas[$email]) { "SIM" } else { "NAO" } } else { "" }
    $desab    = if ($desabilitados.Contains($email)) { "SIM" } else { "NAO" }
    $criada   = if ($cx.WhenCreated) { $cx.WhenCreated.ToString("yyyy-MM-dd") } else { "" }

    # Logon POSTERIOR a demissao = conta de desligado sendo usada.
    # Isso e achado de seguranca por si so — nao se resolve apenas desativando.
    $alerta = ""
    if ($dtLogon -and $sitDT -match 'Demitid' -and $datDem) {
        [datetime]$dtDem = [datetime]::MinValue
        $parsed = $false
        if ($datDem -is [datetime]) { $dtDem = $datDem; $parsed = $true }
        else { $parsed = [datetime]::TryParse([string]$datDem, [ref]$dtDem) }
        if ($parsed -and $dtLogon -gt $dtDem.AddDays(7)) {
            $alerta = "LOGON APOS DEMISSAO ($($dtLogon.ToString('yyyy-MM-dd')) vs demissao $($dtDem.ToString('yyyy-MM-dd')))"
        }
    }

    $relatorio.Add([PSCustomObject]@{
        TipoCaixa            = "UserMailbox"
        Email                = $cx.PrimarySmtpAddress
        Nome                 = $cx.DisplayName
        Licenciado           = $lic
        ContaDesabilitada    = $desab
        ContaCriada          = $criada
        UltimoLogon          = if ($dtLogon) { $dtLogon.ToString("yyyy-MM-dd") } elseif ($logonOK) { "sem registro" } else { "" }
        DiasSemLogar         = if ($null -ne $diasLogon) { $diasLogon } else { "" }
        AlertaSeguranca      = $alerta
        Situacao             = $situacao
        FonteConfirmacao     = $fonte
        SituacaoDT           = $sitDT
        DataDemissaoDT       = $datDem
        CausaDemissaoDT      = $causaDem
        Prioridade           = $prioridade
    })
}

foreach ($cx in $caixasShared) {
    $email    = $cx.PrimarySmtpAddress.ToString().ToLower()
    $prefixo  = $email.Split('@')[0].ToLower()
    $display  = [string]$cx.DisplayName
    $nomeNorm = Normalizar-Nome $display
    $ehShareNoNome = $display -match '^\s*SHARE\s*[-:]'

    $fonte = ""
    if (Test-CaixaServico -Email $email -Prefixo $prefixo) {
        $situacao = "SharedMailbox de servico/departamento (OK)"; $prioridade = "Baixa"; $fonte = "Config"

    } elseif ($ehShareNoNome) {
        $situacao = "SharedMailbox de demitido convertido - nome SHARE- (OK - padrao correto)"; $prioridade = "Baixa"; $fonte = "Padrao"

    } elseif ($prefixo.StartsWith($PrefixoShared)) {
        $situacao = "SharedMailbox com prefixo share. (OK)"; $prioridade = "Baixa"; $fonte = "Padrao"

    } elseif ($nomeNorm -and $nomesAtivos.Contains($nomeNorm)) {
        if     ($nomesCSVSet.Contains($nomeNorm)) { $fonte = "CSV-nome" }
        elseif ($nomesPJsSet.Contains($nomeNorm)) { $fonte = "PJ-nome" }
        else                                      { $fonte = "DT-nome" }
        $situacao = "SharedMailbox com nome de colaborador ATIVO (verificar - deveria ser UserMailbox?)"; $prioridade = "Media"

    } else {
        $situacao = "SharedMailbox SEM prefixo share. e SEM nome SHARE- (fora do padrao)"; $prioridade = "Media"; $fonte = ""
    }

    $hist     = Get-InfoHistDT -Email $email -NomeNorm $nomeNorm
    $sitDT    = if ($hist) { $hist.situacao }      else { "" }
    $datDem   = if ($hist) { $hist.dataDemissao }  else { "" }
    $causaDem = if ($hist) { $hist.causaDemissao } else { "" }
    $lic      = if ($licencas.ContainsKey($email)) { if ($licencas[$email]) { "SIM" } else { "NAO" } } else { "" }
    $desab    = if ($desabilitados.Contains($email)) { "SIM" } else { "NAO" }
    $criada   = if ($cx.WhenCreated) { $cx.WhenCreated.ToString("yyyy-MM-dd") } else { "" }

    # SharedMailbox nao deveria receber logon interativo. Se recebeu, alguem entra nela direto.
    $dtLogonSh   = if ($ultimoLogon.ContainsKey($email)) { $ultimoLogon[$email] } else { $null }
    $diasLogonSh = if ($dtLogonSh) { [int]((Get-Date) - $dtLogonSh).TotalDays } else { $null }
    $alertaSh    = if ($dtLogonSh -and $lic -eq "SIM") { "SharedMailbox LICENCIADA com logon direto em $($dtLogonSh.ToString('yyyy-MM-dd'))" } else { "" }

    $relatorio.Add([PSCustomObject]@{
        TipoCaixa            = "SharedMailbox"
        Email                = $cx.PrimarySmtpAddress
        Nome                 = $cx.DisplayName
        Licenciado           = $lic
        ContaDesabilitada    = $desab
        ContaCriada          = $criada
        UltimoLogon          = if ($dtLogonSh) { $dtLogonSh.ToString("yyyy-MM-dd") } else { "" }
        DiasSemLogar         = if ($null -ne $diasLogonSh) { $diasLogonSh } else { "" }
        AlertaSeguranca      = $alertaSh
        Situacao             = $situacao
        FonteConfirmacao     = $fonte
        SituacaoDT           = $sitDT
        DataDemissaoDT       = $datDem
        CausaDemissaoDT      = $causaDem
        Prioridade           = $prioridade
    })
}

# ===== RESUMO =====
$alta  = @($relatorio | Where-Object { $_.Prioridade -eq 'Alta' })
$media = @($relatorio | Where-Object { $_.Prioridade -eq 'Media' })
$baixa = @($relatorio | Where-Object { $_.Prioridade -eq 'Baixa' })

Write-Host "=== RESUMO ===" -ForegroundColor Cyan
Write-Host ""

$altaDemitidosDT   = @($alta | Where-Object { $_.SituacaoDT -like 'Demit*' -or $_.SituacaoDT -like 'Em demiss*' })
$altaTransferidos  = @($alta | Where-Object { $_.SituacaoDT -like 'Transfer*' })
$altaTrabalhando   = @($alta | Where-Object { $_.SituacaoDT -like 'Trabalh*' -or $_.SituacaoDT -like 'Afastad*' -or $_.SituacaoDT -like 'Em admiss*' })
$altaDesabSemDT    = @($alta | Where-Object { $_.ContaDesabilitada -eq 'SIM' -and [string]::IsNullOrWhiteSpace($_.SituacaoDT) })
$altaSemInfoAtiva  = @($alta | Where-Object { [string]::IsNullOrWhiteSpace($_.SituacaoDT) -and $_.ContaDesabilitada -ne 'SIM' })

Write-Host "[ALTA]  Total fora do CSV/PJs/DT: $($alta.Count)" -ForegroundColor Red
Write-Host ""

if ($altaDemitidosDT.Count -gt 0) {
    Write-Host ">>> DEMITIDOS CONFIRMADOS no historico DT: $($altaDemitidosDT.Count) <<<" -ForegroundColor Red
    Write-Host "    ACAO: converter para SharedMailbox (prefixo 'share.', nome 'SHARE - Nome')" -ForegroundColor Yellow
    $altaDemitidosDT | Sort-Object ContaDesabilitada, Email |
        Format-Table TipoCaixa, Licenciado, ContaDesabilitada, Email, Nome, SituacaoDT, DataDemissaoDT -AutoSize
    Write-Host ""
}

if ($altaDesabSemDT.Count -gt 0) {
    Write-Host ">>> CONTA DESABILITADA no Azure AD (sem historico DT): $($altaDesabSemDT.Count) <<<" -ForegroundColor Red
    Write-Host "    ACAO: confirmar demissao com o RH e converter para SharedMailbox" -ForegroundColor Yellow
    $altaDesabSemDT | Sort-Object Email |
        Format-Table TipoCaixa, Licenciado, Email, Nome -AutoSize
    Write-Host ""
}

if ($altaTransferidos.Count -gt 0) {
    Write-Host ">>> TRANSFERIDOS DE EMPRESA: $($altaTransferidos.Count) (verificar se ainda trabalham)" -ForegroundColor Yellow
    $altaTransferidos | Sort-Object Email |
        Format-Table TipoCaixa, Licenciado, ContaDesabilitada, Email, Nome, SituacaoDT -AutoSize
    Write-Host ""
}

if ($altaTrabalhando.Count -gt 0) {
    Write-Host ">>> ATIVOS NO HISTORICO MAS NAO NO CSV: $($altaTrabalhando.Count)" -ForegroundColor Cyan
    Write-Host "    (e-mail do Exchange diferente do cadastrado, ou trocaram de empresa)" -ForegroundColor Gray
    $altaTrabalhando | Sort-Object Email |
        Format-Table TipoCaixa, Licenciado, ContaDesabilitada, Email, Nome, SituacaoDT -AutoSize
    Write-Host ""
}

if ($altaSemInfoAtiva.Count -gt 0) {
    Write-Host ">>> SEM INFORMACAO - conta ativa, nao esta em nenhuma lista: $($altaSemInfoAtiva.Count)" -ForegroundColor DarkYellow
    Write-Host "    (verificar manualmente no RH)" -ForegroundColor Gray
    $altaSemInfoAtiva | Sort-Object Email |
        Format-Table TipoCaixa, Licenciado, Email, Nome -AutoSize
    Write-Host ""
}

Write-Host "[MEDIA] Fora do padrao (verificar): $($media.Count)" -ForegroundColor Yellow
$media | Sort-Object TipoCaixa, ContaCriada |
    Format-Table TipoCaixa, Licenciado, ContaDesabilitada, Email, Nome, ContaCriada, Situacao, SituacaoDT -AutoSize

# ===== ALERTA: SharedMailbox com licenca (desperdicio de custo) =====
$sharedComLicenca = @($relatorio | Where-Object { $_.TipoCaixa -eq 'SharedMailbox' -and $_.Licenciado -eq 'SIM' })
Write-Host ""
Write-Host "================================================" -ForegroundColor Magenta
Write-Host "  [DESPERDICIO] SharedMailbox COM licenca: $($sharedComLicenca.Count)" -ForegroundColor Magenta
Write-Host "  ACAO: remover licenca (SharedMailbox < 50GB nao precisa de licenca)" -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Magenta
if ($sharedComLicenca.Count -gt 0) {
    $sharedComLicenca | Sort-Object Email | Format-Table Email, Nome, Situacao -AutoSize
}

# ===== SALVAR =====
$pastaLogs = Join-Path $pasta "logs"
New-Item -ItemType Directory -Force -Path $pastaLogs | Out-Null

# Manter apenas os ultimos 12 relatorios (rotacao automatica)
$logsAntigos = @(Get-ChildItem -Path $pastaLogs -Filter "auditoria_*.csv" |
    Sort-Object LastWriteTime -Descending | Select-Object -Skip 12)
foreach ($f in $logsAntigos) {
    Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
    Write-Host "Log antigo removido: $($f.Name)" -ForegroundColor DarkGray
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$csvSaida  = Join-Path $pastaLogs "auditoria_$timestamp.csv"
$relatorio | Sort-Object Prioridade, TipoCaixa, Email |
    Export-Csv -Path $csvSaida -NoTypeInformation -Encoding UTF8

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

# ===== LOG DE EXECUCAO =====
$logExec = Join-Path $pasta "execution.log"
$errosDT  = if ($historicoDT_porNome.Count -eq 0 -and (Test-Path $PlanilhaDT)) { "AVISO:DT=0registros" } else { "OK" }
$authUsed = if ($modoAutomatico) { "app-only:$AppId" } else { "interativo:$AdminUPN" }
"$(Get-Date -Format 'yyyy-MM-dd HH:mm') | $authUsed | $($relatorio.Count) caixas | Alta:$($alta.Count) Media:$($media.Count) Baixa:$($baixa.Count) | $errosDT" |
    Add-Content $logExec -Encoding UTF8

# ===== DASHBOARD =====
$dashPy = Join-Path $pasta "Gerar-Dashboard.py"
if (Test-Path $dashPy) {
    Write-Host ""
    Write-Host "Gerando dashboard..." -ForegroundColor Gray
    try {
        & python $dashPy
    } catch {
        Write-Host "[AVISO] Nao foi possivel gerar o dashboard: $_" -ForegroundColor Yellow
    }
}

# ===== ENVIO DE E-MAIL =====
if ($enviarEmail) {
    if (-not $modoAutomatico) {
        Write-Host ""
        Write-Host "[INFO] EnviarEmail=true mas modo interativo esta ativo." -ForegroundColor DarkYellow
        Write-Host "       O envio de e-mail requer modo automatico (app-only)." -ForegroundColor DarkYellow
        Write-Host "       Configure via Setup-Automacao.ps1 e execute sem -AdminUPN." -ForegroundColor DarkYellow
    } elseif ([string]::IsNullOrWhiteSpace($emailRemetente) -or [string]::IsNullOrWhiteSpace($emailDestinatario)) {
        Write-Host ""
        Write-Host "[AVISO] EnviarEmail=true mas EmailRemetente ou EmailDestinatario nao configurados no config.json." -ForegroundColor Yellow
    } else {
        Write-Host ""
        Write-Host "Enviando relatorio por e-mail para $emailDestinatario..." -ForegroundColor Gray
        try {
            $certEmail    = Get-Item "Cert:\CurrentUser\My\$CertThumbprint" -ErrorAction Stop
            $tenantParam  = if (-not [string]::IsNullOrWhiteSpace($tenantId)) { $tenantId } else { $Organization }
            $token        = Get-GraphToken -Tenant $tenantParam -AppId $AppId -Cert $certEmail
            Enviar-RelatorioEmail `
                -Token       $token `
                -De          $emailRemetente `
                -Para        $emailDestinatario `
                -CsvPath     $csvSaida `
                -TotalCaixas $relatorio.Count `
                -TotalAlta   $alta.Count `
                -TotalMedia  $media.Count `
                -TotalBaixa  $baixa.Count `
                -DataExecucao (Get-Date -Format 'dd/MM/yyyy HH:mm')
            Write-Host "E-mail enviado com sucesso!" -ForegroundColor Green
        } catch {
            Write-Host "[AVISO] Falha ao enviar e-mail: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "Relatorio completo salvo em:" -ForegroundColor Green
Write-Host "  $csvSaida" -ForegroundColor Green
Write-Host ""
Write-Host "Concluido com sucesso." -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Red
    Write-Host "[ERRO FATAL]" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Linha: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
    Write-Host "Comando: $($_.InvocationInfo.Line.Trim())" -ForegroundColor Yellow
    Write-Host "================================================" -ForegroundColor Red
} finally {
    Wait-ForKey
}
