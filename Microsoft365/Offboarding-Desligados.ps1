#Requires -Version 5.1
<#
.SYNOPSIS
    Offboarding de colaboradores desligados: converte caixa para compartilhada, renomeia com
    prefixo SHARE, remove licencas (grupo E direta) e desabilita a conta.

.DESCRIPTION
    ORDEM DAS ETAPAS (nao inverter):
      1. CONVERTER  - UserMailbox -> SharedMailbox   (ANTES de tirar a licenca!)
      2. RENOMEAR   - displayName ganha prefixo "SHARE - "
      3. LICENCA    - remove do grupo de licenca E remove atribuicao direta
      4. DESABILITAR- bloqueia a conta

    POR QUE ESSA ORDEM: remover a licenca de uma UserMailbox inicia o prazo de retencao e a
    caixa e excluida em ~30 dias. Convertida para SharedMailbox ela nao precisa de licenca e o
    historico fica preservado (limite de 50 GB).

    TRES CASOS ESPECIAIS TRATADOS (levantados em 13/08/2026):
      a) 28 de 29 contas sao SINCRONIZADAS do AD -> o rename precisa ser feito no AD.
         Renomear no Entra e inutil: o proximo ciclo de sync sobrescreve.
      b) contas CLOUD-ONLY -> desabilitar no AD nao bloqueia o M365.
         Para ela, rename e disable vao direto no Entra.
      c) 12 contas tem licenca por GRUPO **e** DIRETA. Remover so do grupo deixa a direta ativa
         e o custo continua. O script trata as duas sempre.

.PARAMETER ArquivoContas
    CSV com as contas a tratar. Precisa da coluna UserPrincipalName; demais
    colunas sao ignoradas. Obrigatorio.

.PARAMETER Lote
    Rotulo livre do lote, usado no cabecalho e no nome do log. Padrao: "lote".

.PARAMETER Etapas
    Quais etapas executar. Padrao: todas. Permite a execucao em passadas, ex.:
      -Etapas Converter,Renomear,Licenca     (captura economia sem bloquear acesso)
      -Etapas Desabilitar                    (depois da confirmacao do gestor)

.PARAMETER Executar
    Sem este switch o script roda em DRY-RUN e nao altera nada.

.PARAMETER ConfirmarGestor
    Exige confirmacao digitada antes de desabilitar. Use em lotes de contas com
    uso recente, onde bloquear interrompe trabalho em andamento.

.EXAMPLE
    # 1) Sempre comece em dry-run
    .\Offboarding-Desligados.ps1 -ArquivoContas .\sem-uso.csv -Lote sem-uso

.EXAMPLE
    # 2) Lote de baixo risco: contas sem uso recente
    .\Offboarding-Desligados.ps1 -ArquivoContas .\sem-uso.csv -Lote sem-uso -Executar

.EXAMPLE
    # 3) Lote com uso recente: captura a economia sem cortar acesso ainda
    .\Offboarding-Desligados.ps1 -ArquivoContas .\com-uso.csv -Lote com-uso `
        -Etapas Converter,Renomear,Licenca -Executar

.EXAMPLE
    # 4) So depois da confirmacao do gestor, corta o acesso
    .\Offboarding-Desligados.ps1 -ArquivoContas .\com-uso.csv -Lote com-uso `
        -Etapas Desabilitar -ConfirmarGestor -Executar

.NOTES
    PRE-REQUISITOS de sessao (o app de auditoria NAO tem permissao de escrita):
      Connect-MgGraph -Scopes User.ReadWrite.All,GroupMember.ReadWrite.All,Directory.Read.All
      Connect-ExchangeOnline -UserPrincipalName <seu-admin>
      + a sessao do Windows precisa poder escrever nos objetos do AD.
#>
[CmdletBinding()]
param(
    # CSV com as contas a tratar. Coluna obrigatoria: UserPrincipalName.
    [Parameter(Mandatory)][string]$ArquivoContas,

    # Rotulo do lote - usado no cabecalho e no nome do arquivo de log.
    [string]$Lote = 'lote',

    [ValidateSet('Converter','Renomear','Licenca','Desabilitar')]
    [string[]]$Etapas = @('Converter','Renomear','Licenca','Desabilitar'),

    [switch]$Executar,

    # Exige confirmacao digitada antes de desabilitar. Use em lotes de contas
    # COM uso recente, onde bloquear interrompe trabalho em andamento.
    [switch]$ConfirmarGestor,

    [string]$PrefixoShare = 'SHARE - '
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

function Passo($m) { Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function OK($m)    { Write-Host "  [ok]   $m" -ForegroundColor Green }
function Skip($m)  { Write-Host "  [--]   $m" -ForegroundColor DarkGray }
function Warn($m)  { Write-Host "  [!]    $m" -ForegroundColor Yellow }
function Falha($m) { Write-Host "  [ERRO] $m" -ForegroundColor Red }
function Dry($m)   { Write-Host "  [dry]  $m" -ForegroundColor Gray }

# ─────────────────────────────────────────────────────────────
# Pre-requisitos
# ─────────────────────────────────────────────────────────────
Passo "Verificando pre-requisitos"


$precisaGraph    = ($Etapas -contains 'Licenca') -or ($Etapas -contains 'Renomear') -or ($Etapas -contains 'Desabilitar')
$precisaExchange = ($Etapas -contains 'Converter')

if ($precisaGraph) {
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) { Falha "Modulo Microsoft.Graph.Authentication ausente."; exit 1 }
    Import-Module Microsoft.Graph.Authentication
    $ctx = Get-MgContext
    if (-not $ctx) {
        Falha "Sem sessao Graph. Rode antes:"
        Write-Host "  Connect-MgGraph -Scopes User.ReadWrite.All,GroupMember.ReadWrite.All,Directory.Read.All" -ForegroundColor Yellow
        exit 1
    }
    $temEscrita = ($ctx.Scopes -match 'User\.ReadWrite\.All|Directory\.ReadWrite\.All') -and
                  ($ctx.Scopes -match 'GroupMember\.ReadWrite\.All|Group\.ReadWrite\.All|Directory\.ReadWrite\.All')
    OK "Graph conectado ($($ctx.ClientId))"
    if ($Executar -and -not $temEscrita) {
        Falha "A sessao Graph nao tem permissao de escrita (User.ReadWrite.All + GroupMember.ReadWrite.All). Abortando."
        exit 1
    }
}

if ($precisaExchange) {
    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Falha "Sem sessao Exchange Online. Rode antes:"
        Write-Host "  Connect-ExchangeOnline -UserPrincipalName <seu-admin>" -ForegroundColor Yellow
        exit 1
    }
    OK "Exchange Online conectado"
}

# ─────────────────────────────────────────────────────────────
# Alvos
# ─────────────────────────────────────────────────────────────
if (-not (Test-Path $ArquivoContas)) { Falha "Arquivo de contas nao encontrado: $ArquivoContas"; exit 1 }
$alvos = @(Import-Csv -Path $ArquivoContas)

if ($alvos.Count -eq 0) { Falha "Nenhuma conta no arquivo: $ArquivoContas"; exit 1 }
if (-not ($alvos[0].PSObject.Properties.Name -contains 'UserPrincipalName')) {
    Falha "O CSV precisa da coluna 'UserPrincipalName'. Colunas encontradas: $($alvos[0].PSObject.Properties.Name -join ', ')"
    exit 1
}

Passo "Lote $Lote — $($alvos.Count) contas — etapas: $($Etapas -join ', ')"
Write-Host ("Modo: {0}" -f $(if ($Executar) { 'EXECUCAO' } else { 'DRY-RUN (nada sera alterado)' })) -ForegroundColor $(if ($Executar) { 'Red' } else { 'Cyan' })

if ($ConfirmarGestor -and $Etapas -contains 'Desabilitar' -and $Executar) {
    Warn "Este lote foi marcado como contas COM USO RECENTE."
    Warn "Desabilitar interrompe o acesso imediatamente - confirme com o gestor da area."
    $r = Read-Host "Digite CONFIRMO para continuar"
    if ($r -ne 'CONFIRMO') { Write-Host "Abortado." -ForegroundColor Yellow; exit 0 }
}

# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────
function Get-UsuarioGraph {
    param([string]$Email)
    $sel = 'id,displayName,accountEnabled,onPremisesSyncEnabled,assignedLicenses,licenseAssignmentStates'
    return (Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$Email'&`$select=$sel").value |
        Select-Object -First 1
}

function Get-ObjetoAD {
    param([string]$Email)
    $ds = New-Object DirectoryServices.DirectorySearcher
    $ds.Filter = "(&(objectCategory=person)(objectClass=user)(|(mail=$Email)(userPrincipalName=$Email)(proxyAddresses=smtp:$Email)))"
    [void]$ds.PropertiesToLoad.Add('distinguishedname')
    $r = $ds.FindOne()
    if (-not $r) { return $null }
    return $r.GetDirectoryEntry()
}

$resultado = [System.Collections.Generic.List[PSCustomObject]]::new()

# ─────────────────────────────────────────────────────────────
# Processamento
# ─────────────────────────────────────────────────────────────
foreach ($a in $alvos) {
    $email = [string]$a.Email
    $nome  = [string]$a.Nome
    Write-Host "`n--- $nome <$email>" -ForegroundColor White

    $st = [ordered]@{ Nome = $nome; Email = $email; Converter = ''; Renomear = ''; Licenca = ''; Desabilitar = '' }

    $u = $null
    if ($precisaGraph) {
        $u = Get-UsuarioGraph -Email $email
        if (-not $u) { Falha "nao encontrado no Entra — pulando"; $st.Converter='nao encontrado'; $resultado.Add([PSCustomObject]$st); continue }
    }
    $sincronizada = [bool]$u.onPremisesSyncEnabled

    # ---------- ETAPA 1: CONVERTER PARA SHARED ----------
    if ($Etapas -contains 'Converter') {
        try {
            # Conta sem caixa e situacao NORMAL (ex.: so tem licenca gratuita Flow/Power BI,
            # que nao provisiona Exchange). Nao e erro — apenas nao ha o que converter.
            $mb = Get-Mailbox -Identity $email -ErrorAction SilentlyContinue
            if (-not $mb) {
                Skip "sem caixa no Exchange (licenca sem Exchange) — nada a converter"
                $st.Converter = 'sem caixa'
            }
            elseif ($mb.RecipientTypeDetails -eq 'SharedMailbox') {
                Skip "caixa ja e SharedMailbox"; $st.Converter = 'ja era shared'
            } else {
                # pre-flight de tamanho: shared > 50 GB precisa de licenca
                $stat = Get-MailboxStatistics -Identity $email -ErrorAction SilentlyContinue
                $gb = $null
                if ($stat.TotalItemSize -and ([string]$stat.TotalItemSize) -match '\(([\d,\.]+) bytes\)') {
                    $gb = [math]::Round([double]($matches[1] -replace '[,\.]','') / 1GB, 2)
                }
                if ($gb -ne $null -and $gb -gt 50) {
                    Warn "caixa tem $gb GB (> 50 GB) — NAO convertida. Trate o arquivo morto antes."
                    $st.Converter = "BLOQUEADO ${gb}GB"
                } elseif ($Executar) {
                    Set-Mailbox -Identity $email -Type Shared -ErrorAction Stop
                    OK "convertida para SharedMailbox ($gb GB)"; $st.Converter = 'convertida'
                } else {
                    Dry "converteria para SharedMailbox ($gb GB)"; $st.Converter = 'dry'
                }
            }
        } catch { Falha "converter: $($_.Exception.Message)"; $st.Converter = 'ERRO' }
    }

    # ---------- ETAPA 2: RENOMEAR COM PREFIXO SHARE ----------
    if ($Etapas -contains 'Renomear') {
        try {
            $atual = [string]$u.displayName
            if ($atual -match '^\s*SHARE\s*[-:]') {
                Skip "displayName ja tem prefixo SHARE"; $st.Renomear = 'ja tinha'
            } else {
                $novo = "$PrefixoShare$atual"
                if ($sincronizada) {
                    # AD e a fonte de autoridade — renomear no Entra seria sobrescrito no proximo sync
                    $ad = Get-ObjetoAD -Email $email
                    if (-not $ad) {
                        Warn "sincronizada mas objeto AD nao localizado — rename ignorado"; $st.Renomear = 'AD nao achado'
                    } elseif ($Executar) {
                        $ad.Put('displayName', $novo); $ad.SetInfo()
                        OK "AD displayName -> '$novo'"; $st.Renomear = 'renomeada (AD)'
                    } else {
                        Dry "renomearia no AD para '$novo'"; $st.Renomear = 'dry (AD)'
                    }
                } else {
                    if ($Executar) {
                        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($u.id)" `
                            -Body (@{ displayName = $novo } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop
                        OK "Entra displayName -> '$novo' (cloud-only)"; $st.Renomear = 'renomeada (Entra)'
                    } else {
                        Dry "renomearia no Entra para '$novo' (cloud-only)"; $st.Renomear = 'dry (Entra)'
                    }
                }
            }
        } catch { Falha "renomear: $($_.Exception.Message)"; $st.Renomear = 'ERRO' }
    }

    # ---------- ETAPA 3: REMOVER LICENCAS (GRUPO + DIRETA) ----------
    if ($Etapas -contains 'Licenca') {
        try {
            # so remove licenca se a caixa ja for shared (ou nao existir caixa)
            $podeRemover = $true
            $mbAtual = Get-Mailbox -Identity $email -ErrorAction SilentlyContinue
            if ($mbAtual -and $mbAtual.RecipientTypeDetails -ne 'SharedMailbox') {
                if ($Executar) {
                    Warn "caixa ainda NAO e shared — licenca NAO sera removida (evita exclusao da caixa em 30d)"
                    $podeRemover = $false; $st.Licenca = 'adiado: caixa nao shared'
                }
            }

            if ($podeRemover) {
                $estados = @($u.licenseAssignmentStates)
                $grupos  = @($estados | Where-Object { $_.assignedByGroup } | Select-Object -ExpandProperty assignedByGroup -Unique)
                $diretas = @($estados | Where-Object { -not $_.assignedByGroup } | Select-Object -ExpandProperty skuId -Unique)

                if ($grupos.Count -eq 0 -and $diretas.Count -eq 0) {
                    Skip "sem licenca atribuida"; $st.Licenca = 'ja sem licenca'
                } else {
                    $feito = @()
                    foreach ($g in $grupos) {
                        if ($Executar) {
                            try {
                                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$g/members/$($u.id)/`$ref" -ErrorAction Stop
                                OK "removida do grupo de licenca $g"; $feito += 'grupo'
                            } catch { Falha "remover do grupo ${g}: $($_.Exception.Message)" }
                        } else { Dry "removeria do grupo de licenca $g"; $feito += 'grupo' }
                    }
                    if ($diretas.Count -gt 0) {
                        if ($Executar) {
                            try {
                                $body = @{ addLicenses = @(); removeLicenses = @($diretas) } | ConvertTo-Json -Depth 4
                                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($u.id)/assignLicense" `
                                    -Body $body -ContentType 'application/json' -ErrorAction Stop
                                OK "removida(s) $($diretas.Count) licenca(s) DIRETA(s)"; $feito += 'direta'
                            } catch { Falha "remover licenca direta: $($_.Exception.Message)" }
                        } else { Dry "removeria $($diretas.Count) licenca(s) DIRETA(s)"; $feito += 'direta' }
                    }
                    $st.Licenca = ($feito | Select-Object -Unique) -join '+'
                }
            }
        } catch { Falha "licenca: $($_.Exception.Message)"; $st.Licenca = 'ERRO' }
    }

    # ---------- ETAPA 4: DESABILITAR CONTA ----------
    if ($Etapas -contains 'Desabilitar') {
        try {
            if ($sincronizada) {
                $ad = Get-ObjetoAD -Email $email
                if (-not $ad) {
                    Warn "objeto AD nao localizado — conta NAO desabilitada"; $st.Desabilitar = 'AD nao achado'
                } else {
                    $uac = [int]$ad.userAccountControl.Value
                    if ($uac -band 2) {
                        Skip "conta AD ja estava desabilitada"; $st.Desabilitar = 'ja estava'
                    } elseif ($Executar) {
                        $ad.userAccountControl = $uac -bor 2
                        $ad.SetInfo()
                        OK "conta AD desabilitada (propaga ao M365 no proximo sync)"; $st.Desabilitar = 'desabilitada (AD)'
                    } else {
                        Dry "desabilitaria a conta no AD"; $st.Desabilitar = 'dry (AD)'
                    }
                }
            } else {
                # cloud-only: AD nao resolve, tem que ser no Entra
                if (-not $u.accountEnabled) {
                    Skip "conta Entra ja estava desabilitada"; $st.Desabilitar = 'ja estava'
                } elseif ($Executar) {
                    Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($u.id)" `
                        -Body (@{ accountEnabled = $false } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop
                    OK "conta Entra desabilitada (cloud-only)"; $st.Desabilitar = 'desabilitada (Entra)'
                } else {
                    Dry "desabilitaria no Entra (cloud-only)"; $st.Desabilitar = 'dry (Entra)'
                }
            }
        } catch { Falha "desabilitar: $($_.Exception.Message)"; $st.Desabilitar = 'ERRO' }
    }

    $resultado.Add([PSCustomObject]$st)
}

# ─────────────────────────────────────────────────────────────
# Resumo + log
# ─────────────────────────────────────────────────────────────
Passo "Resumo — lote $Lote"
$resultado | Format-Table -AutoSize | Out-String -Width 200 | Write-Host

$erros = @($resultado | Where-Object { $_.Converter -eq 'ERRO' -or $_.Renomear -eq 'ERRO' -or $_.Licenca -eq 'ERRO' -or $_.Desabilitar -eq 'ERRO' })
if ($erros.Count -gt 0) { Warn "$($erros.Count) conta(s) com erro — revise acima antes de repetir." }

$bloq = @($resultado | Where-Object { $_.Converter -like 'BLOQUEADO*' })
if ($bloq.Count -gt 0) { Warn "$($bloq.Count) caixa(s) acima de 50 GB nao convertidas — tratar arquivo morto e repetir." }

if ($Executar) {
    $dirLog = Join-Path $PSScriptRoot 'logs'
    if (-not (Test-Path $dirLog)) { New-Item -ItemType Directory -Path $dirLog | Out-Null }
    $log = Join-Path $dirLog ("offboarding_{0}_{1}.csv" -f $Lote, (Get-Date -Format 'yyyyMMdd_HHmm'))
    $resultado | Export-Csv -Path $log -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    OK "log gravado: $log"
} else {
    Write-Host "`nDRY-RUN concluido. Nada foi alterado. Repita com -Executar para aplicar." -ForegroundColor Cyan
}
