#Requires -Version 7.0
<#
.SYNOPSIS
    Varre caixas do Exchange Online procurando lixo de sincronizacao
    (pastas "Problemas de Sincronizacao/Conflitos" PT e "Sync Issues/Conflicts" EN)
    e calcula o espaco recuperavel por caixa.

.DESCRIPTION
    Coletor de lixo de sincronizacao para painel de inventario de caixas.
    - Conecta app-only (cert) usando o config.json desta pasta (mesmo do Auditar.ps1).
    - Para cada caixa no escopo, soma itens e tamanho das pastas do tipo
      SyncIssues / Conflicts / LocalFailures / ServerFailures.
    - Enriquece com uso total, quota e % de ocupacao -> classifica criticidade.
    - Exporta 2 CSVs: conflitos_bruto.csv (por pasta) e conflitos_resumo.csv (por caixa).

    NAO apaga nada. Somente leitura. A limpeza e feita em outro processo.

.PARAMETER Escopo
    Compartilhadas (padrao) | Usuarios | Todas

.PARAMETER MinGB
    Ignora caixas cujo lixo recuperavel seja menor que este valor (default 0 = todas).

.PARAMETER OutDir
    Pasta de saida dos CSVs (default: .\dados).

.PARAMETER AdminUPN
    Forca login interativo com este UPN (fallback se nao houver cert app-only).

.EXAMPLE
    .\Scan-Conflitos.ps1
.EXAMPLE
    .\Scan-Conflitos.ps1 -Escopo Todas -MinGB 0.5
#>
param(
    [ValidateSet('Compartilhadas','Usuarios','Todas')]
    [string]$Escopo = 'Compartilhadas',
    [double]$MinGB   = 0,
    [string]$OutDir  = "$PSScriptRoot\dados",
    [string]$AdminUPN
)

$ErrorActionPreference = 'Stop'
$pasta = $PSScriptRoot

# Forca ponto decimal nos CSVs (locale pt-BR escreveria virgula, que quebra o
# Measure-Object do orquestrador e o parseFloat do dashboard JS).
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture

function Write-Step($m){ Write-Host "`n>> $m" -ForegroundColor Cyan }
function Write-OK($m)  { Write-Host "   v $m"  -ForegroundColor Green }
function Write-Warn($m){ Write-Host "   ! $m"  -ForegroundColor Yellow }
function Write-Fail($m){ Write-Host "   x $m"  -ForegroundColor Red }

# ── Converte string de tamanho EXO ("1.23 GB (1,234,567 bytes)") em bytes ──
function ConvertTo-Bytes($valor) {
    if ($null -eq $valor) { return 0 }
    $s = "$valor"
    if ($s -match '\(([\d.,]+)\s*bytes\)') {
        return [double]($matches[1] -replace '[.,]', '')
    }
    if ($s -match '([\d.,]+)\s*(B|KB|MB|GB|TB)') {
        $n = [double](($matches[1] -replace '\.', '') -replace ',', '.')
        switch ($matches[2]) {
            'B'  { return $n }
            'KB' { return $n * 1KB }
            'MB' { return $n * 1MB }
            'GB' { return $n * 1GB }
            'TB' { return $n * 1TB }
        }
    }
    return 0
}

function To-GB($bytes){ [math]::Round(($bytes / 1GB), 2) }

# ===== 1. Conexao =====
Write-Step "Conectando ao Exchange Online..."
$cfgPath = Join-Path $pasta 'config.json'
if (-not (Test-Path $cfgPath)) { Write-Fail "config.json nao encontrado em $pasta"; exit 1 }
$cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json

if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
    Write-Warn "Instalando modulo ExchangeOnlineManagement..."
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement -Force

try {
    if ($AdminUPN) {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
        Write-OK "Conectado (interativo) como $AdminUPN"
    } else {
        Connect-ExchangeOnline -AppId $cfg.AppId -CertificateThumbprint $cfg.CertThumbprint `
            -Organization $cfg.Organization -ShowBanner:$false -ErrorAction Stop
        Write-OK "Conectado app-only ($($cfg.Organization))"
    }
} catch {
    Write-Fail "Falha na conexao: $($_.Exception.Message)"
    exit 1
}

# ===== 2. Lista de caixas =====
Write-Step "Listando caixas (escopo: $Escopo)..."
$filtroTipo = switch ($Escopo) {
    'Compartilhadas' { 'SharedMailbox' }
    'Usuarios'       { 'UserMailbox' }
    'Todas'          { $null }
}
$paramMb = @{ ResultSize = 'Unlimited' }
if ($filtroTipo) { $paramMb['RecipientTypeDetails'] = $filtroTipo }
$caixas = Get-EXOMailbox @paramMb -Properties RecipientTypeDetails,DisplayName,PrimarySmtpAddress,ProhibitSendReceiveQuota
Write-OK "$($caixas.Count) caixas no escopo."

# ===== 3. Varredura =====
Write-Step "Varrendo pastas de conflito/sincronizacao..."
$bruto   = [System.Collections.Generic.List[object]]::new()
$resumo  = [System.Collections.Generic.List[object]]::new()

$i = 0
foreach ($mb in $caixas) {
    $i++
    if ($i % 50 -eq 0) { Write-Host "   ... $i/$($caixas.Count)" -ForegroundColor DarkGray }
    $email = $mb.PrimarySmtpAddress

  try {
    # Pastas de lixo desta caixa.
    # IMPORTANTE: -FolderScope SyncIssues e obrigatorio — as pastas de conflito ficam
    # na subarvore nao-IPM e NAO aparecem no Get-EXOMailboxFolderStatistics padrao.
    try {
        $folders = Get-EXOMailboxFolderStatistics -Identity $email -FolderScope SyncIssues -ErrorAction Stop |
            Where-Object { [int]$_.ItemsInFolder -gt 0 }
    } catch { continue }
    if (-not $folders) { continue }

    $itensLixo = 0
    [double]$bytesLixo = 0
    foreach ($f in $folders) {
        $fb = ConvertTo-Bytes $f.FolderSize
        $itensLixo += [int]$f.ItemsInFolder
        $bytesLixo += $fb
        $bruto.Add([PSCustomObject]@{
            Email = $email; Nome = $mb.DisplayName
            Pasta = $f.Name; FolderType = $f.FolderType
            Itens = [int]$f.ItemsInFolder; GB = (To-GB $fb)
        })
    }
    $gbLixo = To-GB $bytesLixo
    if ($gbLixo -lt $MinGB) { continue }

    # Uso total + quota -> %
    [double]$bytesUso = 0; [double]$bytesQuota = 0
    try {
        $st = Get-EXOMailboxStatistics -Identity $email -ErrorAction Stop
        $bytesUso = ConvertTo-Bytes $st.TotalItemSize
    } catch {}
    $bytesQuota = ConvertTo-Bytes $mb.ProhibitSendReceiveQuota
    $pct = if ($bytesQuota -gt 0) { [math]::Round(($bytesUso / $bytesQuota) * 100, 1) } else { 0 }

    $criticidade = if ($pct -ge 90) { 'Critica' }
                   elseif ($pct -ge 70) { 'Alta' }
                   elseif ($pct -ge 40) { 'Media' }
                   else { 'Baixa' }

    $licenciado = $false
    # Heuristica de licenca: shared sem licenca; user assume licenciado.
    if ($mb.RecipientTypeDetails -eq 'SharedMailbox') { $licenciado = $false } else { $licenciado = $true }

    $resumo.Add([PSCustomObject]@{
        Email           = $email
        Nome            = $mb.DisplayName
        TipoCaixa       = $mb.RecipientTypeDetails
        Licenciado      = $licenciado
        UsoTotalGB      = (To-GB $bytesUso)
        QuotaGB         = (To-GB $bytesQuota)
        PercentUso      = $pct
        Criticidade     = $criticidade
        PastasConflito  = $folders.Count
        ItensLixo       = $itensLixo
        GBLixo          = $gbLixo
        UsoAposLimpezaGB = (To-GB ([math]::Max([double]0, ($bytesUso - $bytesLixo))))
    })
  } catch {
    Write-Warn "Erro na caixa $email : $($_.Exception.Message)"
    continue
  }
}

Write-OK "$($resumo.Count) caixas com lixo (acima de ${MinGB} GB)."

# ===== 4. Exportar =====
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$pBruto  = Join-Path $OutDir 'conflitos_bruto.csv'
$pResumo = Join-Path $OutDir 'conflitos_resumo.csv'
$bruto  | Sort-Object GB -Descending     | Export-Csv $pBruto  -NoTypeInformation -Encoding UTF8
$resumo | Sort-Object GBLixo -Descending | Export-Csv $pResumo -NoTypeInformation -Encoding UTF8
Write-OK "Exportado: $pResumo"
Write-OK "Exportado: $pBruto"

$gbTotal = [math]::Round((($resumo | Measure-Object GBLixo -Sum).Sum), 1)
$itensTotal = ($resumo | Measure-Object ItensLixo -Sum).Sum
Write-Host ""
Write-Host "  RESUMO: $($resumo.Count) caixas | $gbTotal GB recuperaveis | $itensTotal itens de lixo" -ForegroundColor White

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
