#requires -version 5.1
<#
.SYNOPSIS
    Assistente de instalação e reinstalação do NotaJá.

.DESCRIPTION
    - Nova instalação com seleção entre:
        mysqlnuvem  = banco de dados em nuvem
        semmysql    = estação de trabalho local
        instalarmysql = banco de dados local (MySQL 5.6)
    - Reinstalação:
        1. Encerra processos do NotaJá
        2. Protege C:\NFE temporariamente como C:\NFEBKP para o desinstalador não tocar em seu conteúdo
        3. Executa o desinstalador Inno Setup e restaura C:\NFEBKP para C:\NFE
        4. Confirma que C:\NFE e C:\DPCOMPV foram preservadas
        5. Cria um backup completo da base em C:\NotaJa-Suporte\Backups\Base\DPCOMPV-<data-hora>
        5. Renomeia C:\DPCOMPV para C:\DPCOMPVBKP
        6. Faz backup e remove somente os componentes listados de System32/SysWOW64
        7. Usa o instalador já escolhido/baixado e validado no início da sessão
        8. Para o serviço MySQL
        9. Renomeia a nova C:\DPCOMPV para C:\DPCOMPVAZIO, DPCOMPVAZIO2 etc.
       10. Restaura C:\DPCOMPVBKP como C:\DPCOMPV e inicia o MySQL

    O script nunca remove C:\NFE, C:\DPCOMPV ou C:\DPCOMPVBKP.
    Logs, backups permanentes e arquivos temporários ficam organizados em C:\NotaJa-Suporte.

.NOTES
    Requer Windows PowerShell 5.1 e privilégios de administrador.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ============================================================
# CONFIGURAÇÃO — ALTERE ESTES VALORES ANTES DE PUBLICAR
# ============================================================

# URL RAW deste próprio script no GitHub.
# Ela permite que o script se eleve novamente quando executado com:
# irm "URL_RAW" | iex
$RemoteScriptUrl = 'https://raw.githubusercontent.com/fiuzadan/nota/main/NotaJa-Suporte.ps1'

# Download direto da versão mais atual publicada no site oficial do NotaJá.
$InstallerUrl = 'https://cliente.notaja.com.br/emissor-nota-fiscal-eletronica/download/NotaJa.exe'

# Se este arquivo existir, o script pergunta se deve usá-lo ou baixar do site.
# Se não existir, o download do site é iniciado automaticamente.
$LocalInstallerPath = 'C:\NotaJa.exe'

# SHA-256 esperado da versão atual.
# Gere com: Get-FileHash C:\NotaJa.exe -Algorithm SHA256
$ExpectedInstallerSha256 = 'E3CB6B76ED9AB4AC5B715AD9059637DF50A0EF8A8CD186F5BF47ABFE907EB51A'

# Mantém uma cópia dos componentes removidos em C:\NotaJa-Suporte\Backups\Componentes.
$BackupComponentsBeforeRemoval = $true

# Quando falso, o instalador baixado é mantido durante toda a sessão e removido ao encerrar o assistente.
# C:\NotaJa.exe nunca é removido pelo script.
$KeepDownloadedInstaller = $false

# ============================================================
# CAMINHOS E LISTAS
# ============================================================

$TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'

# Toda a estrutura de suporte fica centralizada aqui.
$SupportRoot = 'C:\NotaJa-Suporte'
$LogDirectory = Join-Path $SupportRoot 'Logs'
$BackupDirectoryRoot = Join-Path $SupportRoot 'Backups'
$DatabaseSafetyBackupRoot = Join-Path $BackupDirectoryRoot 'Base'
$ComponentBackupRoot = Join-Path $BackupDirectoryRoot 'Componentes'
$BackupRoot = Join-Path $ComponentBackupRoot $TimeStamp
$WorkDirectory = Join-Path $SupportRoot "Temp\$TimeStamp"

$InstallerPath = Join-Path $WorkDirectory 'NotaJa.exe'
$SessionLog = Join-Path $LogDirectory "NotaJa-Suporte-$TimeStamp.log"

# Instalador escolhido/baixado logo no início da sessão.
$script:PreparedInstaller = $null

# Pastas críticas que jamais devem ser excluídas pelo script.
$NfeDirectory = 'C:\NFE'
$NfeBackupDirectory = 'C:\NFEBKP'
$DatabaseDirectory = 'C:\DPCOMPV'
$DatabaseBackupDirectory = 'C:\DPCOMPVBKP'
$EmptyDatabaseBaseDirectory = 'C:\DPCOMPVAZIO'

# Nomes mais comuns do serviço instalado pelo MySQL 5.6.
$PreferredMySqlServiceNames = @(
    'MySQL',
    'MySQL56'
)

$ComponentFiles = @(
    'CFeSatDataSetX.dll',
    'CFeSatGovX.ocx',
    'CFeSatX.ocx',
    'GNReGovX.ocx',
    'GnreX.ocx',
    'licensex.lic',
    'licensex.ocx',
    'MDFeGovX.ocx',
    'MDFeX.lic',
    'MDFeX.ocx',
    'NFCeDataSetX.dll',
    'NFCeGovX.ocx',
    'NFCeX.ocx',
    'NFeDataSetX.dll',
    'NFeGovX.ocx',
    'NFeX.dll',
    'NFSeBaseX.ocx',
    'NFSeConverterX.ocx',
    'NFSeConverterX.lic',
    'NFSeDataSetX.dll',
    'NFSeImpressaoRB.dll',
    'NFSeImpressaoRBUnicode.dll',
    'NFSeNacionalConverterX.lic',
    'NFSeNacionalConverterX.ocx',
    'NFSeNacionalDataSetX.dll',
    'NFSeNacionalGovX.ocx',
    'NFSeNacionalImpressaoRB.dll',
    'NFSeNacionalImpressaoRBUnicode.dll',
    'NFSeNacionalRESTX.dll',
    'NFSeNacionalSignerX.dll',
    'NFSeNacionalX.dll',
    'NFSeRESTX.dll',
    'NFSeSignerX.dll',
    'NFSeX.dll',
    'spdEmail.dll',
    'spdMdfeDatasetX.ocx',
    'spdNFELib.dll',
    'spdNFeLibUNICODE.dll',
    'spdNFSeConverterLib.dll',
    'spdNFSeConverterLibUnicode.dll',
    'spdNotaSeguraX.lic',
    'spdNotaSeguraX.ocx',
    'spdpresenterdialogx.lic',
    'spdpresenterdialogx.ocx',
    'spdValidadorClientX.lic',
    'spdValidadorClientX.ocx',
    'validador_events.dll',
    'XSDDataSetX.dll',
    'spdNFSeLib.dll',
    'spdNFSeLibUnicode.dll'
)

$NotaJaProcesses = @(
    'NotaJa',
    'NotaJaMenu',
    'NotaJaServico',
    'NotaJaCupom',
    'NotaJaSAT',
    'DPUpdate',
    'DPBackupNuvem',
    'ImportadorSefaz',
    'DPCloudEmissor',
    'DpConsultoria'
)

# ============================================================
# FUNÇÕES BÁSICAS
# ============================================================

function Initialize-Environment {
    # Cria toda a estrutura centralizada do suporte.
    New-Item -ItemType Directory -Path $SupportRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $BackupDirectoryRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $DatabaseSafetyBackupRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $ComponentBackupRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $WorkDirectory -Force | Out-Null

    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor
            [Net.SecurityProtocolType]::Tls12
    }
    catch {
        # O download ainda será tentado com a configuração disponível.
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'OK', 'AVISO', 'ERRO')]
        [string]$Level = 'INFO'
    )

    $Line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $SessionLog -Value $Line -Encoding UTF8

    switch ($Level) {
        'OK'    { Write-Host $Line -ForegroundColor Green }
        'AVISO' { Write-Host $Line -ForegroundColor Yellow }
        'ERRO'  { Write-Host $Line -ForegroundColor Red }
        default { Write-Host $Line }
    }
}

function Test-Administrator {
    $Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object Security.Principal.WindowsPrincipal($Identity)

    return $Principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}

function Restart-Elevated {
    if (Test-Administrator) {
        return
    }

    Write-Host 'Solicitando privilégios de administrador...' -ForegroundColor Yellow

    if ($PSCommandPath) {
        $CommandToRun = "& '$($PSCommandPath.Replace("'", "''"))'"
    }
    elseif (
        -not [string]::IsNullOrWhiteSpace($RemoteScriptUrl) -and
        $RemoteScriptUrl -notmatch 'SEU_USUARIO|SEU_REPOSITORIO'
    ) {
        $SafeUrl = $RemoteScriptUrl.Replace("'", "''")
        $CommandToRun = @"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-RestMethod -Uri '$SafeUrl' | Invoke-Expression
"@
    }
    else {
        throw 'O script não está elevado e a variável $RemoteScriptUrl ainda não foi configurada.'
    }

    $EncodedCommand = [Convert]::ToBase64String(
        [Text.Encoding]::Unicode.GetBytes($CommandToRun)
    )

    Start-Process -FilePath 'powershell.exe' `
        -Verb RunAs `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $EncodedCommand"

    exit
}

function Read-ValidChoice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [string[]]$ValidChoices
    )

    while ($true) {
        $Choice = (Read-Host $Prompt).Trim()

        if ($ValidChoices -contains $Choice) {
            return $Choice
        }

        Write-Host "Opção inválida. Use: $($ValidChoices -join ', ')." -ForegroundColor Yellow
    }
}

function Confirm-ExactText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedText
    )

    $TypedText = Read-Host $Prompt
    return ($TypedText -ceq $ExpectedText)
}

function Pause-Menu {
    [void](Read-Host 'Pressione ENTER para continuar')
}

# ============================================================
# INSTALAÇÃO
# ============================================================

function Select-InstallationComponent {
    Clear-Host
    Write-Host '=============================================' -ForegroundColor Cyan
    Write-Host '      TIPO DE INSTALAÇÃO DO NOTAJÁ' -ForegroundColor Cyan
    Write-Host '=============================================' -ForegroundColor Cyan
    Write-Host
    Write-Host '1 - Utilizar banco de dados em nuvem'
    Write-Host '2 - Esta máquina é uma estação de trabalho local'
    Write-Host '3 - Instalar banco de dados local (MySQL 5.6)'
    Write-Host '0 - Voltar'
    Write-Host

    switch (Read-ValidChoice -Prompt 'Escolha uma opção' -ValidChoices @('0', '1', '2', '3')) {
        '1' {
            return [PSCustomObject]@{
                Label     = 'Banco de dados em nuvem'
                Component = 'mysqlnuvem'
            }
        }
        '2' {
            return [PSCustomObject]@{
                Label     = 'Estação de trabalho local'
                Component = 'semmysql'
            }
        }
        '3' {
            return [PSCustomObject]@{
                Label     = 'Banco de dados local (MySQL 5.6)'
                Component = 'instalarmysql'
            }
        }
        default {
            return $null
        }
    }
}

function Test-Configuration {
    if (
        [string]::IsNullOrWhiteSpace($InstallerUrl) -or
        $InstallerUrl -match 'SEU_USUARIO|SEU_REPOSITORIO'
    ) {
        throw 'Configure a variável $InstallerUrl no início do script antes de utilizá-lo.'
    }

    if ([string]::IsNullOrWhiteSpace($LocalInstallerPath)) {
        throw 'Configure a variável $LocalInstallerPath no início do script antes de utilizá-lo.'
    }
}

function Test-NotaJaInstallerFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$SourceDescription,

        [bool]$DeleteOnHashFailure = $false
    )

    if (
        -not (Test-Path -LiteralPath $Path -PathType Leaf) -or
        (Get-Item -LiteralPath $Path).Length -le 0
    ) {
        throw "O instalador de $SourceDescription não é um arquivo válido: $Path"
    }

    $FileLengthMb = [Math]::Round(
        (Get-Item -LiteralPath $Path).Length / 1MB,
        2
    )

    Write-Log "Instalador selecionado: $SourceDescription ($FileLengthMb MB)."

    if (-not [string]::IsNullOrWhiteSpace($ExpectedInstallerSha256)) {
        $ExpectedHash = $ExpectedInstallerSha256.Replace(' ', '').ToUpperInvariant()
        $ActualHash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()

        if ($ActualHash -ne $ExpectedHash) {
            if ($DeleteOnHashFailure) {
                Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
            }

            throw "O SHA-256 do instalador não confere. Esperado: $ExpectedHash | Obtido: $ActualHash"
        }

        Write-Log 'SHA-256 do instalador conferido com sucesso.' 'OK'
    }
    else {
        Write-Log 'SHA-256 não configurado. O instalador será executado sem validação de hash.' 'AVISO'
    }

    try {
        $Signature = Get-AuthenticodeSignature -LiteralPath $Path

        if ($Signature.Status -eq 'Valid') {
            Write-Log "Assinatura digital válida: $($Signature.SignerCertificate.Subject)" 'OK'
        }
        else {
            Write-Log "Assinatura digital do instalador: $($Signature.Status)." 'AVISO'
        }
    }
    catch {
        Write-Log "Não foi possível consultar a assinatura digital: $($_.Exception.Message)" 'AVISO'
    }
}

function Get-NotaJaInstaller {
    Test-Configuration

    # Se houver um instalador em C:\NotaJa.exe, deixa o técnico escolher.
    if (Test-Path -LiteralPath $LocalInstallerPath -PathType Leaf) {
        Write-Host ''
        Write-Host "Instalador local encontrado em: $LocalInstallerPath" -ForegroundColor Yellow
        Write-Host ''
        Write-Host '1 - Usar o instalador local'
        Write-Host '2 - Baixar a versão mais atual do site do NotaJá'
        Write-Host ''

        do {
            $InstallerChoice = (Read-Host 'Escolha 1 ou 2').Trim()
        } until ($InstallerChoice -in @('1', '2'))

        if ($InstallerChoice -eq '1') {
            Write-Log "Utilizando o instalador local: $LocalInstallerPath"

            Test-NotaJaInstallerFile `
                -Path $LocalInstallerPath `
                -SourceDescription 'arquivo local' `
                -DeleteOnHashFailure $false

            return [PSCustomObject]@{
                Path        = $LocalInstallerPath
                IsTemporary = $false
                Source      = 'Arquivo local'
            }
        }

        Write-Log 'O arquivo local foi ignorado por escolha do usuário. Será feito download do site oficial.'
    }
    else {
        Write-Log "Nenhum instalador local encontrado em $LocalInstallerPath. O download será iniciado automaticamente."
    }

    # Download para a pasta temporária. O arquivo local, se existir, nunca é alterado.
    if (Test-Path -LiteralPath $InstallerPath) {
        Remove-Item -LiteralPath $InstallerPath -Force
    }

    Write-Log "Baixando o instalador do site oficial: $InstallerUrl"

    try {
        Invoke-WebRequest `
            -Uri $InstallerUrl `
            -OutFile $InstallerPath `
            -UseBasicParsing
    }
    catch {
        Write-Log "Invoke-WebRequest falhou: $($_.Exception.Message). Tentando WebClient." 'AVISO'

        if (Test-Path -LiteralPath $InstallerPath) {
            Remove-Item -LiteralPath $InstallerPath -Force -ErrorAction SilentlyContinue
        }

        $WebClient = New-Object Net.WebClient

        try {
            $WebClient.DownloadFile($InstallerUrl, $InstallerPath)
        }
        finally {
            $WebClient.Dispose()
        }
    }

    if (
        -not (Test-Path -LiteralPath $InstallerPath -PathType Leaf) -or
        (Get-Item -LiteralPath $InstallerPath).Length -le 0
    ) {
        throw 'O instalador não foi baixado corretamente.'
    }

    $FileLengthMb = [Math]::Round(
        (Get-Item -LiteralPath $InstallerPath).Length / 1MB,
        2
    )

    Write-Log "Download concluído: $FileLengthMb MB." 'OK'

    Test-NotaJaInstallerFile `
        -Path $InstallerPath `
        -SourceDescription 'download do site oficial' `
        -DeleteOnHashFailure $true

    return [PSCustomObject]@{
        Path        = $InstallerPath
        IsTemporary = $true
        Source      = 'Site oficial'
    }
}

function Initialize-InstallerSource {
    Clear-Host
    Write-Host '=============================================' -ForegroundColor Cyan
    Write-Host '         INSTALADOR DO NOTAJÁ' -ForegroundColor Cyan
    Write-Host '=============================================' -ForegroundColor Cyan
    Write-Host

    Write-Log "Verificando se existe instalador local em $LocalInstallerPath."

    # Esta decisão acontece uma única vez, antes do menu principal e antes
    # da escolha do tipo de instalação/reinstalação.
    $script:PreparedInstaller = Get-NotaJaInstaller

    Write-Host
    Write-Host "Instalador preparado: $($script:PreparedInstaller.Source)" -ForegroundColor Green
    Write-Log "Instalador preparado no início da sessão. Origem: $($script:PreparedInstaller.Source). Caminho: $($script:PreparedInstaller.Path)." 'OK'
}

function Remove-PreparedInstaller {
    if ($null -eq $script:PreparedInstaller) {
        return
    }

    # O arquivo local C:\NotaJa.exe nunca é removido.
    if ($script:PreparedInstaller.IsTemporary -and -not $KeepDownloadedInstaller) {
        $PreparedPath = $script:PreparedInstaller.Path

        if (Test-Path -LiteralPath $PreparedPath -PathType Leaf) {
            Remove-Item -LiteralPath $PreparedPath -Force -ErrorAction SilentlyContinue
            Write-Log "Instalador temporário removido ao encerrar a sessão: $PreparedPath"
        }
    }

    $script:PreparedInstaller = $null

    # Remove apenas a pasta temporária desta sessão se estiver vazia.
    # Backups e logs nunca são removidos automaticamente.
    if (Test-Path -LiteralPath $WorkDirectory -PathType Container) {
        try {
            $RemainingItems = @(Get-ChildItem -LiteralPath $WorkDirectory -Force -ErrorAction Stop)

            if ($RemainingItems.Count -eq 0) {
                Remove-Item -LiteralPath $WorkDirectory -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log "Não foi possível limpar a pasta temporária da sessão: $($_.Exception.Message)" 'AVISO'
        }
    }
}

function Invoke-NotaJaInstallation {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Selection
    )

    if ($null -eq $script:PreparedInstaller) {
        throw 'O instalador não foi preparado no início da sessão.'
    }

    $Installer = $script:PreparedInstaller
    $SetupPath = $Installer.Path

    if (-not (Test-Path -LiteralPath $SetupPath -PathType Leaf)) {
        throw "O instalador selecionado no início da sessão não está mais disponível em: $SetupPath"
    }
    $InstallLog = Join-Path $LogDirectory "Instalacao-$($Selection.Component)-$TimeStamp.log"

    $Arguments = @(
        '/SP-',
        '/SILENT',
        '/SUPPRESSMSGBOXES',
        '/NORESTART',
        '/CLOSEAPPLICATIONS',
        '/TYPE="custom"',
        "/COMPONENTS=`"$($Selection.Component)`"",
        "/LOG=`"$InstallLog`""
    ) -join ' '

    Write-Log "Iniciando instalação: $($Selection.Label)."
    Write-Log "Componente interno: $($Selection.Component)."
    Write-Log "Origem do instalador: $($Installer.Source)."
    Write-Log 'A janela nativa de progresso do instalador será exibida durante a instalação.'
    Write-Host
    Write-Host 'Instalação em andamento. A janela de progresso do NotaJá será exibida.' -ForegroundColor Cyan

    $Process = Start-Process `
        -FilePath $SetupPath `
        -ArgumentList $Arguments `
        -Wait `
        -PassThru

    $SetupExitCode = $Process.ExitCode

    if ($SetupExitCode -ne 0) {
        # Este instalador específico já foi observado concluindo todas as etapas e, ainda assim,
        # fazendo o processo pai retornar código 1. Para não confundir isso com sucesso real,
        # somente o código 1 pode ser aceito mediante validações pós-instalação rigorosas.
        $LogLooksComplete = $false

        if (Test-Path -LiteralPath $InstallLog -PathType Leaf) {
            try {
                $InstallLogContent = Get-Content -LiteralPath $InstallLog -Raw -ErrorAction Stop
                $LogLooksComplete = (
                    $InstallLogContent -match 'Installation process succeeded\.' -and
                    $InstallLogContent -match 'Deinitializing Setup\.'
                )
            }
            catch {
                Write-Log "Não foi possível analisar o log do instalador após o código $SetupExitCode: $($_.Exception.Message)" 'AVISO'
            }
        }

        $UninstallerPresent = [bool](Get-UninstallerPath)
        $ComponentPostCheck = $true

        if ($Selection.Component -eq 'instalarmysql') {
            $ComponentPostCheck = Test-Path -LiteralPath $DatabaseDirectory -PathType Container
        }

        if (
            $SetupExitCode -eq 1 -and
            $LogLooksComplete -and
            $UninstallerPresent -and
            $ComponentPostCheck
        ) {
            Write-Log "O processo do instalador retornou código 1, porém o log fechou normalmente e as validações pós-instalação foram aprovadas. O fluxo continuará e o MySQL será validado na próxima etapa." 'AVISO'
        }
        else {
            throw "O instalador terminou com o código $SetupExitCode e as validações pós-instalação não confirmaram uma instalação íntegra. Consulte: $InstallLog"
        }
    }
    else {
        Write-Log 'Instalação concluída com código 0.' 'OK'
    }

    Write-Log "Log do instalador: $InstallLog"

    # O instalador preparado no início é mantido durante toda a sessão.
    # Se for temporário, será removido somente ao encerrar o assistente.
}

# ============================================================
# PROTEÇÃO E TROCA DA BASE LOCAL
# ============================================================

function Test-ProtectedFoldersBeforeReinstallation {
    $NfeExists = Test-Path -LiteralPath $NfeDirectory -PathType Container
    $NfeBackupExists = Test-Path -LiteralPath $NfeBackupDirectory -PathType Container

    if ($NfeExists -and $NfeBackupExists) {
        throw "As pastas $NfeDirectory e $NfeBackupDirectory existem ao mesmo tempo. A reinstalação foi bloqueada para não sobrescrever nenhum conteúdo."
    }

    if (-not $NfeExists -and $NfeBackupExists) {
        Write-Log "Foi encontrada $NfeBackupDirectory sem $NfeDirectory. Restaurando a NFE antes de continuar." 'AVISO'
        Move-Item -LiteralPath $NfeBackupDirectory -Destination $NfeDirectory -ErrorAction Stop
        $NfeExists = $true
    }

    if (-not $NfeExists) {
        throw "A pasta protegida $NfeDirectory não foi encontrada. A reinstalação foi bloqueada para evitar perda de dados."
    }

    $DatabaseExists = Test-Path -LiteralPath $DatabaseDirectory -PathType Container
    $BackupExists = Test-Path -LiteralPath $DatabaseBackupDirectory -PathType Container

    if ($DatabaseExists -and $BackupExists) {
        throw "As pastas $DatabaseDirectory e $DatabaseBackupDirectory já existem ao mesmo tempo. A reinstalação foi bloqueada para não sobrescrever nenhuma base."
    }

    if (-not $DatabaseExists -and -not $BackupExists) {
        throw "Nem $DatabaseDirectory nem $DatabaseBackupDirectory foram encontradas. A reinstalação foi bloqueada para evitar perda da base local."
    }

    if ($BackupExists -and -not $DatabaseExists) {
        Write-Log "Foi encontrada $DatabaseBackupDirectory sem $DatabaseDirectory. O script tratará isso como uma reinstalação anterior interrompida." 'AVISO'
    }

    Write-Log "Pasta protegida confirmada: $NfeDirectory" 'OK'
}

function Protect-NfeBeforeUninstall {
    if (-not (Test-Path -LiteralPath $NfeDirectory -PathType Container)) {
        throw "A pasta $NfeDirectory não existe e não pode ser protegida antes da desinstalação."
    }

    if (Test-Path -LiteralPath $NfeBackupDirectory) {
        throw "A pasta temporária de proteção $NfeBackupDirectory já existe. O processo foi bloqueado para evitar sobrescrita."
    }

    Write-Log "Protegendo a NFE contra o desinstalador: $NfeDirectory -> $NfeBackupDirectory."
    Move-Item -LiteralPath $NfeDirectory -Destination $NfeBackupDirectory -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $NfeBackupDirectory -PathType Container)) {
        throw "Não foi possível confirmar a proteção da pasta NFE em $NfeBackupDirectory."
    }

    Write-Log "Pasta NFE protegida temporariamente em: $NfeBackupDirectory" 'OK'
}

function Restore-NfeAfterUninstall {
    if (-not (Test-Path -LiteralPath $NfeBackupDirectory -PathType Container)) {
        if (Test-Path -LiteralPath $NfeDirectory -PathType Container) {
            return
        }

        throw "Nem $NfeDirectory nem $NfeBackupDirectory foram encontradas durante a restauração da NFE."
    }

    if (Test-Path -LiteralPath $NfeDirectory -PathType Container) {
        $UnexpectedNfe = Join-Path $WorkDirectory 'NFE-Gerada-Durante-Desinstalacao'
        $Suffix = 2
        while (Test-Path -LiteralPath $UnexpectedNfe) {
            $UnexpectedNfe = Join-Path $WorkDirectory "NFE-Gerada-Durante-Desinstalacao$Suffix"
            $Suffix++
        }

        Write-Log "Uma nova $NfeDirectory apareceu durante a desinstalação. Ela será preservada em $UnexpectedNfe antes de restaurar a original." 'AVISO'
        Move-Item -LiteralPath $NfeDirectory -Destination $UnexpectedNfe -ErrorAction Stop
    }

    Write-Log "Restaurando a pasta NFE original: $NfeBackupDirectory -> $NfeDirectory."
    Move-Item -LiteralPath $NfeBackupDirectory -Destination $NfeDirectory -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $NfeDirectory -PathType Container)) {
        throw "A pasta NFE original não pôde ser restaurada para $NfeDirectory."
    }

    Write-Log 'Pasta NFE original restaurada após a desinstalação.' 'OK'
}

function Test-ProtectedFoldersAfterUninstall {
    if (-not (Test-Path -LiteralPath $NfeDirectory -PathType Container)) {
        throw "A pasta protegida $NfeDirectory não foi encontrada após a desinstalação. O processo foi interrompido."
    }

    if (
        -not (Test-Path -LiteralPath $DatabaseDirectory -PathType Container) -and
        -not (Test-Path -LiteralPath $DatabaseBackupDirectory -PathType Container)
    ) {
        throw "A base local não foi encontrada após a desinstalação: $DatabaseDirectory / $DatabaseBackupDirectory. O processo foi interrompido."
    }

    Write-Log 'As pastas protegidas permaneceram disponíveis após a desinstalação.' 'OK'
}

function Get-UniqueDatabaseSafetyBackupDirectory {
    $BaseName = "DPCOMPV-$TimeStamp"
    $Candidate = Join-Path $DatabaseSafetyBackupRoot $BaseName

    if (-not (Test-Path -LiteralPath $Candidate)) {
        return $Candidate
    }

    $Index = 2

    while ($true) {
        $Candidate = Join-Path $DatabaseSafetyBackupRoot ("{0}-{1}" -f $BaseName, $Index)

        if (-not (Test-Path -LiteralPath $Candidate)) {
            return $Candidate
        }

        $Index++
    }
}

function Backup-DatabaseBeforeReinstallation {
    $SourceDirectory = $null

    if (Test-Path -LiteralPath $DatabaseDirectory -PathType Container) {
        $SourceDirectory = $DatabaseDirectory
    }
    elseif (Test-Path -LiteralPath $DatabaseBackupDirectory -PathType Container) {
        # Permite recuperar com segurança uma reinstalação anterior interrompida.
        $SourceDirectory = $DatabaseBackupDirectory
        Write-Log "$DatabaseDirectory não existe. O backup de segurança será feito a partir de $DatabaseBackupDirectory." 'AVISO'
    }
    else {
        throw 'Não existe uma pasta de base local disponível para gerar o backup de segurança.'
    }

    New-Item -ItemType Directory -Path $DatabaseSafetyBackupRoot -Force | Out-Null
    $DestinationDirectory = Get-UniqueDatabaseSafetyBackupDirectory

    Write-Log "Criando backup completo da base antes de continuar: $SourceDirectory -> $DestinationDirectory"

    $RoboCopyPath = Join-Path $env:SystemRoot 'System32\robocopy.exe'

    if (Test-Path -LiteralPath $RoboCopyPath) {
        New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null

        $RoboArguments = @(
            "`"$SourceDirectory`"",
            "`"$DestinationDirectory`"",
            '/E',
            '/COPY:DAT',
            '/DCOPY:T',
            '/R:2',
            '/W:2',
            '/XJ',
            '/NP',
            '/NFL',
            '/NDL'
        ) -join ' '

        $Process = Start-Process `
            -FilePath $RoboCopyPath `
            -ArgumentList $RoboArguments `
            -Wait `
            -PassThru `
            -WindowStyle Hidden

        # Códigos 0 a 7 do Robocopy indicam sucesso ou sucesso com diferenças/cópias.
        if ($Process.ExitCode -ge 8) {
            throw "O backup da base falhou. Robocopy retornou código $($Process.ExitCode). A reinstalação foi interrompida antes de alterar a pasta original."
        }
    }
    else {
        Write-Log 'Robocopy não encontrado. Utilizando Copy-Item para o backup da base.' 'AVISO'
        Copy-Item `
            -LiteralPath $SourceDirectory `
            -Destination $DestinationDirectory `
            -Recurse `
            -Force `
            -ErrorAction Stop
    }

    if (-not (Test-Path -LiteralPath $DestinationDirectory -PathType Container)) {
        throw 'Não foi possível confirmar a criação do backup da base. A reinstalação foi interrompida.'
    }

    $BackupHasContent = @(Get-ChildItem -LiteralPath $DestinationDirectory -Force -ErrorAction Stop).Count -gt 0

    if (-not $BackupHasContent) {
        throw "O backup foi criado em $DestinationDirectory, mas está vazio. A reinstalação foi interrompida antes de alterar a base original."
    }

    Write-Log "Backup de segurança da base concluído: $DestinationDirectory" 'OK'
    return $DestinationDirectory
}

function Move-DatabaseToBackup {
    $DatabaseExists = Test-Path -LiteralPath $DatabaseDirectory -PathType Container
    $BackupExists = Test-Path -LiteralPath $DatabaseBackupDirectory -PathType Container

    if ($BackupExists -and -not $DatabaseExists) {
        Write-Log "$DatabaseBackupDirectory já existe e $DatabaseDirectory não existe. Mantendo o backup existente." 'AVISO'
        return
    }

    if ($BackupExists) {
        throw "Não é possível renomear a base: $DatabaseBackupDirectory já existe. Nenhuma pasta será sobrescrita."
    }

    if (-not $DatabaseExists) {
        throw "A pasta $DatabaseDirectory não foi encontrada para ser preservada."
    }

    Write-Log "Renomeando $DatabaseDirectory para $DatabaseBackupDirectory."
    Move-Item -LiteralPath $DatabaseDirectory -Destination $DatabaseBackupDirectory -ErrorAction Stop

    if (
        (Test-Path -LiteralPath $DatabaseDirectory) -or
        -not (Test-Path -LiteralPath $DatabaseBackupDirectory -PathType Container)
    ) {
        throw 'Não foi possível confirmar a criação do backup da pasta DPCOMPV.'
    }

    Write-Log "Base local preservada em: $DatabaseBackupDirectory" 'OK'
}

function Get-UniqueEmptyDatabaseDirectory {
    if (-not (Test-Path -LiteralPath $EmptyDatabaseBaseDirectory)) {
        return $EmptyDatabaseBaseDirectory
    }

    $Index = 2

    while ($true) {
        $Candidate = "${EmptyDatabaseBaseDirectory}$Index"

        if (-not (Test-Path -LiteralPath $Candidate)) {
            return $Candidate
        }

        $Index++
    }
}

function Get-NotaJaMySqlService {
    foreach ($ServiceName in $PreferredMySqlServiceNames) {
        $Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

        if ($Service) {
            return $Service
        }
    }

    $ServiceDetails = @()

    try {
        $ServiceDetails = Get-CimInstance -ClassName Win32_Service -ErrorAction Stop |
            Where-Object {
                $_.Name -match '(?i)mysql' -or
                $_.DisplayName -match '(?i)mysql'
            }
    }
    catch {
        try {
            $ServiceDetails = Get-WmiObject -Class Win32_Service -ErrorAction Stop |
                Where-Object {
                    $_.Name -match '(?i)mysql' -or
                    $_.DisplayName -match '(?i)mysql'
                }
        }
        catch {
            Write-Log "Não foi possível consultar os detalhes dos serviços MySQL: $($_.Exception.Message)" 'AVISO'
        }
    }

    $DpcompvServices = @(
        $ServiceDetails | Where-Object { $_.PathName -match '(?i)DPCOMPV' }
    )

    if ($DpcompvServices.Count -eq 1) {
        return Get-Service -Name $DpcompvServices[0].Name -ErrorAction Stop
    }

    $AllMySqlServices = @(
        Get-Service -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -match '(?i)mysql' -or
                $_.DisplayName -match '(?i)mysql'
            }
    )

    if ($AllMySqlServices.Count -eq 1) {
        return $AllMySqlServices[0]
    }

    if ($AllMySqlServices.Count -gt 1) {
        $Names = ($AllMySqlServices | ForEach-Object { $_.Name }) -join ', '
        throw "Foram encontrados vários serviços MySQL ($Names), mas não foi possível identificar com segurança qual pertence ao NotaJá."
    }

    throw 'O serviço MySQL do NotaJá não foi encontrado após a instalação.'
}

function Stop-NotaJaMySqlService {
    $Service = Get-NotaJaMySqlService
    $ServiceName = $Service.Name

    if ($Service.Status -ne 'Stopped') {
        Write-Log "Parando o serviço MySQL: $ServiceName."
        Stop-Service -Name $ServiceName -Force -ErrorAction Stop

        $Service = Get-Service -Name $ServiceName -ErrorAction Stop
        $Service.WaitForStatus(
            [ServiceProcess.ServiceControllerStatus]::Stopped,
            (New-TimeSpan -Seconds 60)
        )
    }

    $Service = Get-Service -Name $ServiceName -ErrorAction Stop

    if ($Service.Status -ne 'Stopped') {
        throw "O serviço $ServiceName não parou dentro do tempo esperado."
    }

    Write-Log "Serviço MySQL parado: $ServiceName." 'OK'
    return $ServiceName
}

function Start-NotaJaMySqlService {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName
    )

    $Service = Get-Service -Name $ServiceName -ErrorAction Stop

    if ($Service.Status -ne 'Running') {
        Write-Log "Iniciando o serviço MySQL: $ServiceName."
        Start-Service -Name $ServiceName -ErrorAction Stop

        $Service = Get-Service -Name $ServiceName -ErrorAction Stop
        $Service.WaitForStatus(
            [ServiceProcess.ServiceControllerStatus]::Running,
            (New-TimeSpan -Seconds 60)
        )
    }

    $Service = Get-Service -Name $ServiceName -ErrorAction Stop

    if ($Service.Status -ne 'Running') {
        throw "O serviço $ServiceName não iniciou dentro do tempo esperado."
    }

    Write-Log "Serviço MySQL iniciado: $ServiceName." 'OK'
}

function Restore-OriginalDatabase {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName
    )

    if (-not (Test-Path -LiteralPath $DatabaseBackupDirectory -PathType Container)) {
        throw "A pasta de backup $DatabaseBackupDirectory não foi encontrada."
    }

    if (-not (Test-Path -LiteralPath $DatabaseDirectory -PathType Container)) {
        throw "A nova pasta $DatabaseDirectory não foi criada pela instalação."
    }

    $EmptyDirectory = Get-UniqueEmptyDatabaseDirectory

    Write-Log "Renomeando a nova pasta $DatabaseDirectory para $EmptyDirectory."
    Move-Item -LiteralPath $DatabaseDirectory -Destination $EmptyDirectory -ErrorAction Stop

    try {
        Write-Log "Restaurando $DatabaseBackupDirectory como $DatabaseDirectory."
        Move-Item -LiteralPath $DatabaseBackupDirectory -Destination $DatabaseDirectory -ErrorAction Stop
    }
    catch {
        Write-Log 'Falha ao restaurar a base original. Tentando recolocar a pasta nova no caminho DPCOMPV.' 'ERRO'

        if (
            -not (Test-Path -LiteralPath $DatabaseDirectory) -and
            (Test-Path -LiteralPath $EmptyDirectory -PathType Container)
        ) {
            Move-Item -LiteralPath $EmptyDirectory -Destination $DatabaseDirectory -ErrorAction SilentlyContinue
        }

        throw
    }

    Write-Log "Nova base vazia preservada em: $EmptyDirectory" 'OK'
    Write-Log "Base original restaurada em: $DatabaseDirectory" 'OK'

    Start-NotaJaMySqlService -ServiceName $ServiceName
}

function Repair-DatabaseFoldersAfterFailure {
    param(
        [string]$ServiceName
    )

    if (-not (Test-Path -LiteralPath $DatabaseBackupDirectory -PathType Container)) {
        if (
            $ServiceName -and
            (Test-Path -LiteralPath $DatabaseDirectory -PathType Container)
        ) {
            try {
                Start-NotaJaMySqlService -ServiceName $ServiceName
            }
            catch {
                Write-Log "A base original está em $DatabaseDirectory, mas o serviço MySQL não iniciou: $($_.Exception.Message)" 'ERRO'
            }
        }

        return
    }

    Write-Log 'Foi detectada uma falha com a base preservada em DPCOMPVBKP. Iniciando tentativa de recuperação.' 'AVISO'

    if (-not $ServiceName) {
        try {
            $DetectedService = Get-NotaJaMySqlService
            $ServiceName = $DetectedService.Name
            Write-Log "Serviço MySQL identificado durante a recuperação: $ServiceName." 'AVISO'
        }
        catch {
            Write-Log "Não foi possível identificar o serviço MySQL durante a recuperação: $($_.Exception.Message)" 'AVISO'
        }
    }

    if ($ServiceName) {
        try {
            $Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

            if ($Service -and $Service.Status -ne 'Stopped') {
                Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
            }
        }
        catch {
            Write-Log "Não foi possível parar o serviço durante a recuperação: $($_.Exception.Message)" 'AVISO'
        }
    }

    if (Test-Path -LiteralPath $DatabaseDirectory -PathType Container) {
        $RecoveryEmptyDirectory = Get-UniqueEmptyDatabaseDirectory

        try {
            Move-Item -LiteralPath $DatabaseDirectory -Destination $RecoveryEmptyDirectory -ErrorAction Stop
            Write-Log "Pasta DPCOMPV gerada durante a tentativa foi preservada em: $RecoveryEmptyDirectory" 'AVISO'
        }
        catch {
            Write-Log "Não foi possível liberar $DatabaseDirectory durante a recuperação: $($_.Exception.Message)" 'ERRO'
            return
        }
    }

    try {
        Move-Item -LiteralPath $DatabaseBackupDirectory -Destination $DatabaseDirectory -ErrorAction Stop
        Write-Log 'A base original foi restaurada automaticamente após a falha.' 'OK'
    }
    catch {
        Write-Log "Não foi possível restaurar automaticamente a base original: $($_.Exception.Message)" 'ERRO'
        return
    }

    if ($ServiceName) {
        try {
            Start-NotaJaMySqlService -ServiceName $ServiceName
        }
        catch {
            Write-Log "A base foi restaurada, mas o serviço MySQL não iniciou: $($_.Exception.Message)" 'ERRO'
        }
    }
}

# ============================================================
# REINSTALAÇÃO
# ============================================================

function Stop-NotaJaApplications {
    foreach ($ProcessName in $NotaJaProcesses) {
        $Processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue

        foreach ($Process in $Processes) {
            try {
                Write-Log "Encerrando processo: $($Process.ProcessName) (PID $($Process.Id))."
                Stop-Process -Id $Process.Id -Force -ErrorAction Stop
            }
            catch {
                Write-Log "Falha ao encerrar $($Process.ProcessName): $($_.Exception.Message)" 'AVISO'
            }
        }
    }
}

function Get-UninstallerPath {
    $PossiblePaths = @()

    if (${env:ProgramFiles(x86)}) {
        $PossiblePaths += Join-Path ${env:ProgramFiles(x86)} 'Dpcompg\NotaJa\unins000.exe'
    }

    if ($env:ProgramFiles) {
        $PossiblePaths += Join-Path $env:ProgramFiles 'Dpcompg\NotaJa\unins000.exe'
    }

    foreach ($Path in ($PossiblePaths | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $Path) {
            return $Path
        }
    }

    return $null
}

function Invoke-NotaJaUninstall {
    $UninstallerPath = Get-UninstallerPath

    if (-not $UninstallerPath) {
        throw 'O desinstalador unins000.exe não foi encontrado na pasta padrão do NotaJá.'
    }

    $UninstallLog = Join-Path $LogDirectory "Desinstalacao-$TimeStamp.log"
    $Arguments = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /LOG=`"$UninstallLog`""

    Write-Log "Executando desinstalador: $UninstallerPath"

    $Process = Start-Process `
        -FilePath $UninstallerPath `
        -ArgumentList $Arguments `
        -Wait `
        -PassThru

    if ($Process.ExitCode -ne 0) {
        throw "O desinstalador terminou com o código $($Process.ExitCode). Consulte: $UninstallLog"
    }

    Write-Log 'Desinstalação concluída com código 0.' 'OK'
    Write-Log "Log do desinstalador: $UninstallLog"
}

function Unregister-OcxFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory
    )

    if ([IO.Path]::GetExtension($FilePath) -ine '.ocx') {
        return
    }

    $RegSvr32 = Join-Path $SourceDirectory 'regsvr32.exe'

    if (-not (Test-Path -LiteralPath $RegSvr32)) {
        Write-Log "regsvr32.exe não encontrado para desregistrar: $FilePath" 'AVISO'
        return
    }

    try {
        $Process = Start-Process `
            -FilePath $RegSvr32 `
            -ArgumentList "/u /s `"$FilePath`"" `
            -WindowStyle Hidden `
            -Wait `
            -PassThru

        Write-Log "Desregistro tentado para $FilePath. Código: $($Process.ExitCode)."
    }
    catch {
        Write-Log "Falha ao desregistrar ${FilePath}: $($_.Exception.Message)" 'AVISO'
    }
}

function Remove-LegacyComponents {
    $SystemDirectories = @(
        (Join-Path $env:WINDIR 'System32')
    )

    if ([Environment]::Is64BitOperatingSystem) {
        $SystemDirectories += Join-Path $env:WINDIR 'SysWOW64'
    }

    $Removed = 0
    $NotFound = 0
    $Failed = 0

    foreach ($Directory in $SystemDirectories) {
        $DirectoryName = Split-Path -Path $Directory -Leaf

        foreach ($FileName in $ComponentFiles) {
            $FullPath = Join-Path $Directory $FileName

            if (-not (Test-Path -LiteralPath $FullPath)) {
                $NotFound++
                continue
            }

            try {
                if ($BackupComponentsBeforeRemoval) {
                    $BackupDirectory = Join-Path $BackupRoot $DirectoryName
                    New-Item -ItemType Directory -Path $BackupDirectory -Force | Out-Null
                    Copy-Item -LiteralPath $FullPath -Destination $BackupDirectory -Force
                }

                Unregister-OcxFile -FilePath $FullPath -SourceDirectory $Directory
                Remove-Item -LiteralPath $FullPath -Force -ErrorAction Stop

                Write-Log "Removido: $FullPath" 'OK'
                $Removed++
            }
            catch {
                Write-Log "Não foi possível remover ${FullPath}: $($_.Exception.Message)" 'ERRO'
                $Failed++
            }
        }
    }

    Write-Log "Limpeza concluída. Removidos: $Removed | Não encontrados: $NotFound | Falhas: $Failed."

    if ($BackupComponentsBeforeRemoval -and $Removed -gt 0) {
        Write-Log "Backup dos componentes removidos: $BackupRoot"
    }

    if ($Failed -gt 0) {
        $Continue = Read-ValidChoice `
            -Prompt 'Ocorreram falhas. Continuar para a instalação mesmo assim? [S/N]' `
            -ValidChoices @('S', 's', 'N', 'n')

        if ($Continue -notin @('S', 's')) {
            throw 'Reinstalação cancelada devido a falhas na remoção dos componentes.'
        }
    }
}

function Invoke-Reinstallation {
    Clear-Host
    Write-Host '=============================================' -ForegroundColor Red
    Write-Host '             REINSTALAÇÃO DO NOTAJÁ' -ForegroundColor Red
    Write-Host '=============================================' -ForegroundColor Red
    Write-Host
    Write-Host 'Este procedimento irá:' -ForegroundColor Yellow
    Write-Host ' - Preservar obrigatoriamente C:\NFE e a base C:\DPCOMPV'
    Write-Host ' - Proteger temporariamente C:\NFE para que o desinstalador não altere seu conteúdo'
    Write-Host ' - Desinstalar o NotaJá silenciosamente e restaurar C:\NFE em seguida'
    Write-Host ' - Fazer um backup completo de C:\DPCOMPV em C:\NotaJa-Suporte\Backups\Base'
    Write-Host ' - Renomear a base atual para C:\DPCOMPVBKP'
    Write-Host ' - Remover somente os componentes informados de System32/SysWOW64'
    Write-Host ' - Usar o instalador preparado no início e executar a instalação novamente'
    Write-Host ' - Parar o MySQL, guardar a base nova como DPCOMPVAZIO e restaurar a base original'
    Write-Host
    Write-Host 'O script NÃO exclui C:\NFE, C:\DPCOMPV ou C:\DPCOMPVBKP.' -ForegroundColor Green
    Write-Host 'Se já existir DPCOMPVBKP, nenhuma pasta será sobrescrita.' -ForegroundColor Green
    Write-Host 'Os componentes removidos podem ser compartilhados por outros sistemas da TecnoSpeed.' -ForegroundColor Yellow
    Write-Host

    if (-not (Confirm-ExactText -Prompt 'Digite REINSTALAR para confirmar' -ExpectedText 'REINSTALAR')) {
        Write-Log 'Reinstalação cancelada pelo usuário.' 'AVISO'
        return
    }

    $Selection = Select-InstallationComponent

    if (-not $Selection) {
        Write-Log 'Reinstalação cancelada na seleção do componente.' 'AVISO'
        return
    }

    $MySqlServiceName = $null
    $DatabaseWasMoved = $false
    $NfeWasProtected = $false

    Test-ProtectedFoldersBeforeReinstallation
    Stop-NotaJaApplications

    try {
        # O desinstalador do NotaJá possui entradas que removem arquivos dentro de C:\NFE.
        # Para garantir que nenhum conteúdo original seja tocado, a pasta é ocultada dele
        # por meio de um rename temporário e restaurada imediatamente após a desinstalação.
        Protect-NfeBeforeUninstall
        $NfeWasProtected = $true

        Invoke-NotaJaUninstall

        Restore-NfeAfterUninstall
        $NfeWasProtected = $false

        Test-ProtectedFoldersAfterUninstall

        $DatabaseSafetyBackupPath = Backup-DatabaseBeforeReinstallation
        Write-Log "Backup independente confirmado antes da troca da base: $DatabaseSafetyBackupPath" 'OK'

        Move-DatabaseToBackup
        $DatabaseWasMoved = $true

        Remove-LegacyComponents
        Invoke-NotaJaInstallation -Selection $Selection

        $MySqlServiceName = Stop-NotaJaMySqlService
        Restore-OriginalDatabase -ServiceName $MySqlServiceName
        $DatabaseWasMoved = $false

        Write-Log 'Reinstalação e restauração da base concluídas com sucesso.' 'OK'
    }
    catch {
        $OriginalError = $_

        if ($NfeWasProtected -or (Test-Path -LiteralPath $NfeBackupDirectory -PathType Container)) {
            try {
                Restore-NfeAfterUninstall
                $NfeWasProtected = $false
            }
            catch {
                Write-Log "Falha ao restaurar a NFE após erro no fluxo: $($_.Exception.Message)" 'ERRO'
            }
        }

        if ($DatabaseWasMoved -or (Test-Path -LiteralPath $DatabaseBackupDirectory -PathType Container)) {
            Repair-DatabaseFoldersAfterFailure -ServiceName $MySqlServiceName
        }

        throw $OriginalError
    }
}

# ============================================================
# MENU PRINCIPAL
# ============================================================

function Show-MainMenu {
    while ($true) {
        Clear-Host
        Write-Host '=============================================' -ForegroundColor Cyan
        Write-Host '          ASSISTENTE DE SUPORTE NOTAJÁ' -ForegroundColor Cyan
        Write-Host '=============================================' -ForegroundColor Cyan
        Write-Host
        Write-Host '1 - Nova instalação'
        Write-Host '2 - Reinstalação'
        Write-Host '0 - Sair'
        Write-Host

        $Choice = Read-ValidChoice -Prompt 'Escolha uma opção' -ValidChoices @('0', '1', '2')

        try {
            switch ($Choice) {
                '1' {
                    $Selection = Select-InstallationComponent

                    if ($Selection) {
                        Invoke-NotaJaInstallation -Selection $Selection
                        Write-Host
                        Write-Host 'Operação finalizada com sucesso.' -ForegroundColor Green
                        Pause-Menu
                    }
                }

                '2' {
                    Invoke-Reinstallation
                    Write-Host
                    Write-Host 'Operação finalizada.' -ForegroundColor Green
                    Pause-Menu
                }

                '0' {
                    Write-Log 'Assistente encerrado pelo usuário.'
                    return
                }
            }
        }
        catch {
            Write-Log $_.Exception.Message 'ERRO'
            Write-Host
            Write-Host "Consulte o log: $SessionLog" -ForegroundColor Yellow
            Pause-Menu
        }
    }
}

# ============================================================
# INÍCIO
# ============================================================

Initialize-Environment
Restart-Elevated

Write-Log 'Assistente de suporte iniciado.'
Write-Log "Usuário: $env:USERDOMAIN\$env:USERNAME"
Write-Log "Computador: $env:COMPUTERNAME"
Write-Log "PowerShell: $($PSVersionTable.PSVersion)"
Write-Log "Pasta central do suporte: $SupportRoot"
Write-Log "Log da sessão: $SessionLog"

try {
    # Primeira decisão operacional da sessão: localizar C:\NotaJa.exe.
    # Se existir, pergunta se deve usá-lo ou baixar do site.
    # Se não existir, baixa automaticamente. Só depois o menu principal é exibido.
    Initialize-InstallerSource
    Show-MainMenu
}
finally {
    Remove-PreparedInstaller
}
