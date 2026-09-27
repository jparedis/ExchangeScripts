#Requires -Version 5.1

<#
.SYNOPSIS
    Downloads and installs the Windows prerequisites for Exchange Server Subscription Edition on Windows Server 2025.

.DESCRIPTION
    Covers the Mailbox server role prerequisites as listed by Microsoft (Exchange Server prerequisites, SE RTM):

      Windows features      the Install-WindowsFeature list from the Microsoft docs, Server Core aware
      .NET Framework 4.8.1  already part of Windows Server 2025, only installed when the registry says it is missing
      Visual C++ 2012 x64   redistributable
      Visual C++ 2013 x64   redistributable (the KB4032938 build)
      UCMA 4.0              Unified Communications Managed API runtime (the same package as \UCMARedist on the ISO)
      IIS URL Rewrite 2.1   x64 module
      Remote Registry       service startup type set to Automatic

    Three modes:

      Download            only download the packages to -Path. Run this on a machine with internet access and copy
                          the folder to the Exchange server when that server has no internet.
      Install             only install, using the packages already present in -Path. No internet needed.
      DownloadAndInstall  both, on the same machine. This is the default.

    Every item ends in one of these states: Downloaded, AlreadyDownloaded, Installed, AlreadyInstalled, Set,
    Failed or Skipped. A failing item does not stop the script, the remaining items are still processed so one run
    shows the full picture. Everything is written to the console and to a log file in -Path
    (Install-ExchangeSEPrerequisites_<timestamp>.log).

    After installing, the script checks every requirement again (read only) and ends with a verdict:

      SERVER READY FOR EXCHANGE SETUP                 exit code 0
      SERVER READY FOR EXCHANGE SETUP AFTER REBOOT    exit code 3010, an installer or feature asked for a reboot
      SERVER NOT READY FOR EXCHANGE SETUP             exit code 1, the missing items are listed

    In Download mode the verdict is ALL PACKAGES DOWNLOADED (0) or DOWNLOAD INCOMPLETE (1).

    The script is idempotent: packages that are already downloaded are not downloaded again and products that are
    already installed are skipped. Run it with -WhatIf to get the readiness check without changing anything,
    the log file is still written.

    Note: Exchange SE CU1 drops UCMA and moves to the Visual C++ 2022 runtime. This script follows the SE RTM
    prerequisites. Check the Microsoft docs again once CU1 is out.

.PARAMETER Mode
    Download, Install or DownloadAndInstall. Defaults to DownloadAndInstall.

.PARAMETER Path
    Folder where the packages are downloaded to or read from, and where the log file is written.
    Defaults to C:\ExchangePrereqs.

.PARAMETER Restart
    Reboot the server at the end when the server is ready and a reboot is required.

.EXAMPLE
    .\Install-ExchangeSEPrerequisites.ps1 -Mode Download -Path D:\ExchangePrereqs

    Download everything to D:\ExchangePrereqs on a machine with internet access.

.EXAMPLE
    .\Install-ExchangeSEPrerequisites.ps1 -Mode Install -Path D:\ExchangePrereqs -Restart

    Install from the copied folder on the Exchange server and reboot when done.

.EXAMPLE
    .\Install-ExchangeSEPrerequisites.ps1 -Mode Install -WhatIf

    Show what would be installed and report whether the server is ready, without changing anything.

.NOTES
    jentech consulting
    Run in an elevated PowerShell session for Install and DownloadAndInstall.
    Sources: https://learn.microsoft.com/exchange/plan-and-deploy/prerequisites
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('Download', 'Install', 'DownloadAndInstall')]
    [string]$Mode = 'DownloadAndInstall',

    [string]$Path = 'C:\ExchangePrereqs',

    [switch]$Restart
)

$ErrorActionPreference = 'Stop'

# Package list. Detection is a DisplayName pattern in the Uninstall registry keys, except for .NET which uses
# the Release value. Arguments are the silent install switches per installer type.
$Packages = @(
    @{
        Name      = '.NET Framework 4.8.1'
        FileName  = 'NDP481-x86-x64-AllOS-ENU.exe'
        Url       = 'https://go.microsoft.com/fwlink/?linkid=2203305'
        Arguments = '/q /norestart'
        Detect    = 'DotNet481'
    },
    @{
        Name      = 'Visual C++ 2012 Redistributable x64'
        FileName  = 'vcredist2012_x64.exe'
        Url       = 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe'
        Arguments = '/install /quiet /norestart'
        Detect    = 'Microsoft Visual C++ 2012 Redistributable (x64)*'
    },
    @{
        Name      = 'Visual C++ 2013 Redistributable x64'
        FileName  = 'vcredist2013_x64.exe'
        Url       = 'https://aka.ms/highdpimfc2013x64enu'
        Arguments = '/install /quiet /norestart'
        Detect    = 'Microsoft Visual C++ 2013 Redistributable (x64)*'
    },
    @{
        Name      = 'Unified Communications Managed API 4.0'
        FileName  = 'UcmaRuntimeSetup.exe'
        Url       = 'https://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe'
        Arguments = '-q'
        Detect    = 'Microsoft Unified Communications Managed API 4.0, Runtime*'
    },
    @{
        Name      = 'IIS URL Rewrite Module 2.1 x64'
        FileName  = 'rewrite_amd64_en-US.msi'
        Url       = 'https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi'
        Arguments = '/quiet /norestart'
        Detect    = 'IIS URL Rewrite Module 2*'
    }
)

# Windows features for the Mailbox role (Desktop Experience list from the Microsoft docs).
$WindowsFeatures = @(
    'Server-Media-Foundation', 'NET-Framework-45-Core', 'NET-Framework-45-ASPNET', 'NET-WCF-HTTP-Activation45',
    'NET-WCF-Pipe-Activation45', 'NET-WCF-TCP-Activation45', 'NET-WCF-TCP-PortSharing45', 'RPC-over-HTTP-proxy',
    'RSAT-Clustering', 'RSAT-Clustering-CmdInterface', 'RSAT-Clustering-Mgmt', 'RSAT-Clustering-PowerShell',
    'WAS-Process-Model', 'Web-Asp-Net45', 'Web-Basic-Auth', 'Web-Client-Auth', 'Web-Digest-Auth',
    'Web-Dir-Browsing', 'Web-Dyn-Compression', 'Web-Http-Errors', 'Web-Http-Logging', 'Web-Http-Redirect',
    'Web-Http-Tracing', 'Web-ISAPI-Ext', 'Web-ISAPI-Filter', 'Web-Metabase', 'Web-Mgmt-Console',
    'Web-Mgmt-Service', 'Web-Net-Ext45', 'Web-Request-Monitor', 'Web-Server', 'Web-Stat-Compression',
    'Web-Static-Content', 'Web-Windows-Auth', 'Web-WMI', 'Windows-Identity-Foundation', 'RSAT-ADDS'
)

# These three do not exist on Server Core and are left out there, matching the Server Core list in the docs.
$DesktopOnlyFeatures = @('RSAT-Clustering-Mgmt', 'Web-Mgmt-Console', 'Windows-Identity-Foundation')

# Logging and result tracking

$Script:LogFile = $null
$Script:Results = New-Object System.Collections.Generic.List[object]

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'FAIL')]
        [string]$Level = 'INFO'
    )

    $Color = switch ($Level) {
        'OK'   { 'Green' }
        'WARN' { 'Yellow' }
        'FAIL' { 'Red' }
        default { 'Gray' }
    }
    Write-Host $Message -ForegroundColor $Color
    if ($Script:LogFile) {
        $Line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.PadRight(4), $Message
        # WhatIf must not suppress the log itself.
        Add-Content -Path $Script:LogFile -Value $Line -WhatIf:$false
    }
}

function Add-Result {
    param(
        [string]$Item,
        [ValidateSet('Downloaded', 'AlreadyDownloaded', 'Installed', 'AlreadyInstalled', 'Set', 'Failed', 'Skipped')]
        [string]$Status,
        [string]$Detail = ''
    )

    $Script:Results.Add([pscustomobject]@{ Item = $Item; Status = $Status; Detail = $Detail })

    $Level = switch ($Status) {
        'Failed'  { 'FAIL' }
        'Skipped' { 'WARN' }
        default   { 'OK' }
    }
    $Text = '  {0,-18} {1}' -f $Status, $Item
    if ($Detail) { $Text += " ($Detail)" }
    Write-Log -Message $Text -Level $Level
}

# Detection

function Test-ProductInstalled {
    param([string]$Detect)

    if ($Detect -eq 'DotNet481') {
        # Release 533320 is .NET Framework 4.8.1. Windows Server 2025 ships with it.
        $ReleaseValue = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
        return ($ReleaseValue -ge 533320)
    }

    $UninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $Match = Get-ItemProperty $UninstallKeys -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like $Detect }
    return ($null -ne $Match)
}

function Get-RequiredFeatures {
    # Server Core does not know the three desktop only features.
    $InstallationType = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').InstallationType
    if ($InstallationType -eq 'Server Core') {
        return ($WindowsFeatures | Where-Object { $_ -notin $DesktopOnlyFeatures })
    }
    return $WindowsFeatures
}

# Download and install

function Invoke-PackageDownload {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([hashtable]$Package, [string]$Folder)

    $Target = Join-Path $Folder $Package.FileName
    if (Test-Path $Target) {
        Add-Result -Item $Package.Name -Status AlreadyDownloaded
        return
    }
    if (-not $PSCmdlet.ShouldProcess($Target, "Download $($Package.Name)")) {
        Add-Result -Item $Package.Name -Status Skipped -Detail 'WhatIf'
        return
    }

    try {
        Write-Log "  downloading $($Package.Name) from $($Package.Url)"
        Invoke-WebRequest -Uri $Package.Url -OutFile $Target -UseBasicParsing
        $SizeMb = [math]::Round((Get-Item $Target).Length / 1MB, 1)
        Add-Result -Item $Package.Name -Status Downloaded -Detail "$SizeMb MB"
    }
    catch {
        # A partial file would be seen as AlreadyDownloaded on the next run, so remove it.
        Remove-Item $Target -ErrorAction SilentlyContinue
        Add-Result -Item $Package.Name -Status Failed -Detail $_.Exception.Message
    }
}

function Invoke-PackageInstall {
    # Returns $true when the installer asked for a reboot.
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([hashtable]$Package, [string]$Folder)

    if (Test-ProductInstalled -Detect $Package.Detect) {
        Add-Result -Item $Package.Name -Status AlreadyInstalled
        return $false
    }

    $Installer = Join-Path $Folder $Package.FileName
    if (-not (Test-Path $Installer)) {
        Add-Result -Item $Package.Name -Status Failed -Detail "installer not found: $Installer, run with -Mode Download first"
        return $false
    }

    if (-not $PSCmdlet.ShouldProcess($Installer, "Install $($Package.Name)")) {
        Add-Result -Item $Package.Name -Status Skipped -Detail 'WhatIf'
        return $false
    }

    try {
        Write-Log "  installing $($Package.Name)"
        if ($Installer -like '*.msi') {
            $Process = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/i `"$Installer`" $($Package.Arguments)" -Wait -PassThru
        }
        else {
            $Process = Start-Process -FilePath $Installer -ArgumentList $Package.Arguments -Wait -PassThru
        }
    }
    catch {
        Add-Result -Item $Package.Name -Status Failed -Detail $_.Exception.Message
        return $false
    }

    # 3010 and 1641 mean success with reboot required.
    switch ($Process.ExitCode) {
        0 {
            Add-Result -Item $Package.Name -Status Installed
            return $false
        }
        { $_ -in 3010, 1641 } {
            Add-Result -Item $Package.Name -Status Installed -Detail 'reboot required'
            return $true
        }
        default {
            Add-Result -Item $Package.Name -Status Failed -Detail "exit code $($Process.ExitCode)"
            return $false
        }
    }
}

function Install-RequiredFeatures {
    # Returns $true when a reboot is needed.
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $Required = Get-RequiredFeatures
    $Missing = Get-WindowsFeature -Name $Required | Where-Object { -not $_.Installed }
    if (-not $Missing) {
        Add-Result -Item 'Windows features' -Status AlreadyInstalled -Detail "$($Required.Count) features"
        return $false
    }

    $MissingNames = @($Missing.Name)
    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Install $($MissingNames.Count) Windows features")) {
        Add-Result -Item 'Windows features' -Status Skipped -Detail "WhatIf, missing: $($MissingNames -join ', ')"
        return $false
    }

    try {
        Write-Log "  installing features: $($MissingNames -join ', ')"
        $FeatureResult = Install-WindowsFeature -Name $MissingNames
    }
    catch {
        Add-Result -Item 'Windows features' -Status Failed -Detail $_.Exception.Message
        return $false
    }

    if (-not $FeatureResult.Success) {
        Add-Result -Item 'Windows features' -Status Failed -Detail "Install-WindowsFeature exit code $($FeatureResult.ExitCode)"
        return $false
    }

    $RebootNeeded = ($FeatureResult.RestartNeeded -eq 'Yes')
    $Detail = "$($MissingNames.Count) features added"
    if ($RebootNeeded) { $Detail += ', reboot required' }
    Add-Result -Item 'Windows features' -Status Installed -Detail $Detail
    return $RebootNeeded
}

function Set-RemoteRegistryAutomatic {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $Service = Get-Service -Name RemoteRegistry
    if ($Service.StartType -eq 'Automatic') {
        Add-Result -Item 'Remote Registry service' -Status AlreadyInstalled -Detail 'startup type Automatic'
        return
    }
    if (-not $PSCmdlet.ShouldProcess('RemoteRegistry', 'Set startup type to Automatic')) {
        Add-Result -Item 'Remote Registry service' -Status Skipped -Detail "WhatIf, currently $($Service.StartType)"
        return
    }
    try {
        Set-Service -Name RemoteRegistry -StartupType Automatic
        Add-Result -Item 'Remote Registry service' -Status Set -Detail 'startup type Automatic'
    }
    catch {
        Add-Result -Item 'Remote Registry service' -Status Failed -Detail $_.Exception.Message
    }
}

# Final readiness check, read only. Returns the list of missing items.

function Get-MissingRequirements {
    $Missing = New-Object System.Collections.Generic.List[string]

    $MissingFeatures = Get-WindowsFeature -Name (Get-RequiredFeatures) | Where-Object { -not $_.Installed }
    foreach ($Feature in $MissingFeatures) {
        $Missing.Add("Windows feature $($Feature.Name)")
    }

    foreach ($Package in $Packages) {
        if (-not (Test-ProductInstalled -Detect $Package.Detect)) {
            $Missing.Add($Package.Name)
        }
    }

    if ((Get-Service -Name RemoteRegistry).StartType -ne 'Automatic') {
        $Missing.Add('Remote Registry service not set to Automatic')
    }

    return $Missing
}

# Main

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'   # Invoke-WebRequest is very slow on PS 5.1 with the progress bar on
$RebootRequired = $false
$ExitCode = 0

if ($Mode -ne 'Download') {
    $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $IsAdmin) {
        throw 'Run this script in an elevated PowerShell session to install.'
    }
}

if (-not (Test-Path $Path)) {
    New-Item -Path $Path -ItemType Directory -Force -WhatIf:$false | Out-Null
}

$Script:LogFile = Join-Path $Path ('Install-ExchangeSEPrerequisites_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
Write-Log "Exchange SE prerequisites, mode $Mode, server $env:COMPUTERNAME, folder $Path"
Write-Log "Log file: $Script:LogFile"

# Download
if ($Mode -in 'Download', 'DownloadAndInstall') {
    Write-Log ''
    Write-Log 'Packages: download'
    foreach ($Package in $Packages) {
        Invoke-PackageDownload -Package $Package -Folder $Path
    }
}

if ($Mode -eq 'Download') {
    $FailedDownloads = $Script:Results | Where-Object { $_.Status -in 'Failed', 'Skipped' }
    Write-Log ''
    if ($FailedDownloads) {
        Write-Log 'DOWNLOAD INCOMPLETE' -Level FAIL
        foreach ($Item in $FailedDownloads) { Write-Log "  $($Item.Item): $($Item.Detail)" -Level FAIL }
        $ExitCode = 1
    }
    else {
        Write-Log "ALL PACKAGES DOWNLOADED. Copy $Path to the Exchange server and run with -Mode Install -Path <folder>." -Level OK
    }
    exit $ExitCode
}

# Windows features go first: UCMA needs Server-Media-Foundation and URL Rewrite needs IIS.
Write-Log ''
Write-Log 'Windows features'
if (Install-RequiredFeatures) { $RebootRequired = $true }

Write-Log ''
Write-Log "Packages: install from $Path"
foreach ($Package in $Packages) {
    if (Invoke-PackageInstall -Package $Package -Folder $Path) { $RebootRequired = $true }
}

Write-Log ''
Write-Log 'Remote Registry service'
Set-RemoteRegistryAutomatic

# Summary and verdict
Write-Log ''
Write-Log 'Summary'
$Summary = $Script:Results | Group-Object Status | ForEach-Object { '{0} {1}' -f $_.Count, $_.Name }
Write-Log "  $($Summary -join ', ')"

$MissingItems = Get-MissingRequirements
Write-Log ''
if ($MissingItems.Count -gt 0) {
    Write-Log 'SERVER NOT READY FOR EXCHANGE SETUP' -Level FAIL
    Write-Log 'Missing:' -Level FAIL
    foreach ($Item in $MissingItems) { Write-Log "  $Item" -Level FAIL }
    $ExitCode = 1
}
elseif ($RebootRequired) {
    Write-Log 'SERVER READY FOR EXCHANGE SETUP AFTER REBOOT' -Level WARN
    $ExitCode = 3010
    if ($Restart -and $PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Restart computer')) {
        Write-Log 'Rebooting now.' -Level WARN
        Restart-Computer -Force
    }
}
else {
    Write-Log 'SERVER READY FOR EXCHANGE SETUP' -Level OK
}

Write-Log "Log file: $Script:LogFile"
exit $ExitCode
