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
        2. Executa o desinstalador Inno Setup
        3. Faz backup e remove somente os componentes listados de System32/SysWOW64
        4. Baixa e executa novamente o instalador

    O script não remove C:\NFE nem apaga intencionalmente os dados do MySQL.

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
$RemoteScriptUrl = 'https://raw.githubusercontent.com/SEU_USUARIO/SEU_REPOSITORIO/main/NotaJa-Suporte.ps1'

# Recomenda-se hospedar o EXE como ativo de uma GitHub Release.
$InstallerUrl = 'https://github.com/SEU_USUARIO/SEU_REPOSITORIO/releases/download/v5.10c/NotaJa.exe'

# Opcional, mas altamente recomendado.
# Gere com: Get-FileHash .\NotaJa.exe -Algorithm SHA256
# Deixe vazio para apenas avisar e continuar.
$ExpectedInstallerSha256 = ''

# Mantém uma cópia dos componentes removidos em ProgramData.
$BackupComponentsBeforeRemoval = $true

# Exclui o instalador temporário depois de uma instalação bem-sucedida.
$KeepDownloadedInstaller = $false

# ============================================================
# CAMINHOS E LISTAS
# ============================================================

$TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$WorkDirectory = Join-Path $env:TEMP 'NotaJa-Suporte'
$LogDirectory = Join-Path $env:ProgramData 'NotaJa-Suporte\Logs'
$BackupRoot = Join-Path $env:ProgramData "NotaJa-Suporte\BackupComponentes\$TimeStamp"
$InstallerPath = Join-Path $WorkDirectory 'NotaJa.exe'
$SessionLog = Join-Path $LogDirectory "NotaJa-Suporte-$TimeStamp.log"

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
    New-Item -ItemType Directory -Path $WorkDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null

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
}

function Get-NotaJaInstaller {
    Test-Configuration

    if (Test-Path -LiteralPath $InstallerPath) {
        Remove-Item -LiteralPath $InstallerPath -Force
    }

    Write-Log "Baixando o instalador: $InstallerUrl"

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
        -not (Test-Path -LiteralPath $InstallerPath) -or
        (Get-Item -LiteralPath $InstallerPath).Length -le 0
    ) {
        throw 'O instalador não foi baixado corretamente.'
    }

    $FileLengthMb = [Math]::Round(
        (Get-Item -LiteralPath $InstallerPath).Length / 1MB,
        2
    )

    Write-Log "Download concluído: $FileLengthMb MB." 'OK'

    if (-not [string]::IsNullOrWhiteSpace($ExpectedInstallerSha256)) {
        $ExpectedHash = $ExpectedInstallerSha256.Replace(' ', '').ToUpperInvariant()
        $ActualHash = (Get-FileHash -LiteralPath $InstallerPath -Algorithm SHA256).Hash.ToUpperInvariant()

        if ($ActualHash -ne $ExpectedHash) {
            Remove-Item -LiteralPath $InstallerPath -Force -ErrorAction SilentlyContinue
            throw "O SHA-256 do instalador não confere. Esperado: $ExpectedHash | Obtido: $ActualHash"
        }

        Write-Log 'SHA-256 do instalador conferido com sucesso.' 'OK'
    }
    else {
        Write-Log 'SHA-256 não configurado. O instalador será executado sem validação de hash.' 'AVISO'
    }

    try {
        $Signature = Get-AuthenticodeSignature -LiteralPath $InstallerPath

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

    return $InstallerPath
}

function Invoke-NotaJaInstallation {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Selection
    )

    $SetupPath = Get-NotaJaInstaller
    $InstallLog = Join-Path $LogDirectory "Instalacao-$($Selection.Component)-$TimeStamp.log"

    $Arguments = @(
        '/SP-',
        '/VERYSILENT',
        '/SUPPRESSMSGBOXES',
        '/NORESTART',
        '/CLOSEAPPLICATIONS',
        '/TYPE="custom"',
        "/COMPONENTS=`"$($Selection.Component)`"",
        "/LOG=`"$InstallLog`""
    ) -join ' '

    Write-Log "Iniciando instalação: $($Selection.Label)."
    Write-Log "Componente interno: $($Selection.Component)."

    $Process = Start-Process `
        -FilePath $SetupPath `
        -ArgumentList $Arguments `
        -Wait `
        -PassThru

    if ($Process.ExitCode -ne 0) {
        throw "O instalador terminou com o código $($Process.ExitCode). Consulte: $InstallLog"
    }

    Write-Log 'Instalação concluída com código 0.' 'OK'
    Write-Log "Log do instalador: $InstallLog"

    if (-not $KeepDownloadedInstaller) {
        Remove-Item -LiteralPath $SetupPath -Force -ErrorAction SilentlyContinue
        Write-Log 'Instalador temporário removido.'
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
        Write-Log "Falha ao desregistrar $FilePath: $($_.Exception.Message)" 'AVISO'
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
                Write-Log "Não foi possível remover $FullPath: $($_.Exception.Message)" 'ERRO'
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
    Write-Host ' - Desinstalar o NotaJá silenciosamente'
    Write-Host ' - Remover somente os componentes informados de System32/SysWOW64'
    Write-Host ' - Baixar e executar novamente o instalador'
    Write-Host
    Write-Host 'O script NÃO remove C:\NFE nem apaga intencionalmente os dados do MySQL.' -ForegroundColor Green
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

    Stop-NotaJaApplications
    Invoke-NotaJaUninstall
    Remove-LegacyComponents
    Invoke-NotaJaInstallation -Selection $Selection
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
Write-Log "Log da sessão: $SessionLog"

Show-MainMenu
