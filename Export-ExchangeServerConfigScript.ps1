#Requires -Version 5.1

<#
.SYNOPSIS
    Reads the client access and transport configuration of an Exchange server and writes it out as a plain
    PowerShell script of filled in Set commands for a new server.

.DESCRIPTION
    Alternative to Copy-ExchangeServerConfig.ps1 for people who want to see and edit every command before it
    runs. This script only reads. It produces a second script that contains no functions and no logic: just the
    Exchange cmdlets, one per object, with every settable parameter filled in with the value found on the source
    server. Open the generated script, review or trim it, and run it in the Exchange Management Shell against the
    new server.

    The generated script starts with three variables ($TargetServer, $TargetFqdn, $WhatIf). Every command carries
    -WhatIf:$WhatIf, so setting $WhatIf = $true at the top gives a dry run of the whole file. Change $TargetServer
    and $TargetFqdn to reuse the same file for the second DAG member.

    Sections in the generated script, in this order:

      Certificates        Export from the source, import on the target and Enable-ExchangeCertificate with the same
                          services, for every certificate that is not self signed plus the OAuth certificate.
                          Non exportable keys are listed as comments.
      Set-ExchangeServer, Set-ClientAccessService, Set-MailboxServer, Set-MalwareFilteringServer
      Virtual directories OWA, ECP, EWS, ActiveSync, OAB, MAPI, PowerShell, Autodiscover (frontend and backend)
      Set-OutlookAnywhere
      Set-TransportService, Set-FrontendTransportService, Set-MailboxTransportService
      Set-ImapSettings, Set-PopSettings and Set-Service for the IMAP and POP service startup type
      Set-EventLogLevel for every level raised above Lowest
      Set-SettingOverride for overrides scoped to the source server
      Receive connectors  New-ReceiveConnector for connectors that are not the default ones, Set-ReceiveConnector
                          for all of them, Add-ADPermission for every explicit ACE (anonymous relay)
      Send connectors     Set-SendConnector -SourceTransportServers with the target added. Written as comments,
                          remove the leading # once the target is validated because this changes live routing.
      Transport agents and hybrid configuration as comments, they need manual work.

    Values are converted to what the Set cmdlets accept: sizes as byte counts, collections as arrays, enums as
    strings. Strings that contain the source server name or FQDN are written with $TargetServer or $TargetFqdn so
    the generated file adapts to the target. Bindings on a specific IP address are translated with -IpAddressMap,
    unmapped ones are kept as is and flagged with a TODO comment.

    Parameters with an empty value on the source are left out by default, because a fresh server already has them
    empty and some parameter types refuse $null. Use -IncludeNullValues to write them anyway.

    Requirements: Exchange Management Shell with at least View-Only Organization Management on the source. The
    generated script needs Organization Management on the target. Windows PowerShell 5.1.

.PARAMETER SourceServer
    Name of the Exchange server to read.

.PARAMETER TargetServer
    Name of the server the generated script is meant for. Written into the $TargetServer variable at the top of
    the generated script. When the server already exists in the organization its FQDN is read from AD, otherwise
    the DNS suffix of the source is reused.

.PARAMETER OutputPath
    Path of the generated script. Defaults to Set-ExchangeServerConfig_from_<source>_<timestamp>.ps1 in the
    current directory.

.PARAMETER IpAddressMap
    Hashtable that maps IP addresses of the source to those of the target for connector and protocol bindings.
    Example: @{ '10.0.0.11' = '10.0.0.12' }

.PARAMETER IncludePaths
    Also write log, queue and pickup directory paths. Only when the target has the same disk layout.

.PARAMETER IncludeNullValues
    Also write parameters whose value is empty on the source, as $null.

.PARAMETER ExcludeParameter
    Parameter names to leave out of every command, for example deprecated parameters that the target rejects.

.PARAMETER NoServerNameSubstitution
    Write the source server name literally instead of replacing it by $TargetServer and $TargetFqdn.

.EXAMPLE
    .\Export-ExchangeServerConfigScript.ps1 -SourceServer EX01 -TargetServer EX02

    Writes Set-ExchangeServerConfig_from_EX01_<timestamp>.ps1 in the current folder.

.EXAMPLE
    .\Export-ExchangeServerConfigScript.ps1 -SourceServer EX01 -TargetServer EX02 -IpAddressMap @{ '10.0.0.11' = '10.0.0.12' } -OutputPath C:\Temp\EX02-config.ps1

    Relay connector bound to 10.0.0.11 on the source is written with 10.0.0.12 for the target.

.EXAMPLE
    .\Export-ExchangeServerConfigScript.ps1 -SourceServer EX01 -TargetServer EX02 -ExcludeParameter CalendarRepairWorkCycle, SharingPolicyWorkCycle

    Leaves two deprecated Set-MailboxServer parameters out of the generated script.

.NOTES
    jentech consulting
    The generated script is a snapshot. Regenerate it after configuration changes on the source instead of
    editing it by hand, so the file in your documentation always matches the source.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string] $SourceServer,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string] $TargetServer,

    [string] $OutputPath,

    [hashtable] $IpAddressMap = @{},

    [switch] $IncludePaths,

    [switch] $IncludeNullValues,

    [string[]] $ExcludeParameter = @(),

    [switch] $NoServerNameSubstitution
)

# Parameters that are never written, whatever the cmdlet.
$script:NeverCopy = @('Identity', 'Server', 'DomainController', 'Force', 'Confirm', 'WhatIf', 'AsJob', 'Name')

$script:PathExclusions = @('*Path', '*Location', '*Directory')
if ($IncludePaths) { $script:PathExclusions = @() }

# Placeholders that survive quoting and are turned into $TargetServer and $TargetFqdn in the generated text.
$script:ShortToken = '__TARGET_SERVER__'
$script:FqdnToken  = '__TARGET_FQDN__'

$script:Lines = New-Object System.Collections.Generic.List[string]
$script:Warnings = New-Object System.Collections.Generic.List[string]

# ---------------------------------------------------------------------------
# Helpers for the generator itself (the generated script contains none of these)
# ---------------------------------------------------------------------------

function Add-Line {
    param([string] $Text = '')
    $script:Lines.Add($Text)
}

function Add-Section {
    param([string] $Title)
    Add-Line
    Add-Line ('# ' + ('-' * 96))
    Add-Line ('# {0}' -f $Title)
    Add-Line ('# ' + ('-' * 96))
}

function Add-Warning {
    param([string] $Text)
    $script:Warnings.Add($Text)
    Write-Warning $Text
}

function Convert-ServerName {
    # Replaces the source FQDN and NetBIOS name inside a string by placeholders. The placeholders become
    # $TargetFqdn and $TargetServer when the value is quoted. NetBIOS match is bounded so EX01 never hits EX010.
    param($Value)
    if ($NoServerNameSubstitution -or $null -eq $Value) { return $Value }
    if ($Value -is [string]) {
        $result = $Value
        if ($script:SourceFqdn) {
            $result = [regex]::Replace($result, [regex]::Escape($script:SourceFqdn), $script:FqdnToken, 'IgnoreCase')
        }
        $pattern = '(?<![\w-])' + [regex]::Escape($script:SourceShort) + '(?![\w-])'
        return [regex]::Replace($result, $pattern, $script:ShortToken, 'IgnoreCase')
    }
    if ($Value -is [array]) {
        $converted = @(foreach ($item in $Value) { Convert-ServerName -Value $item })
        return ,[string[]]$converted
    }
    return $Value
}

function ConvertFrom-DisplayString {
    # Deserialized sizes look like "35 MB (36,700,160 bytes)", the cmdlets want the byte count.
    param([string] $Text)
    if ($Text -match '^\s*unlimited\s*$') { return 'Unlimited' }
    if ($Text -match '\(([\d,\.\s]+)\s*bytes\)') { return ('{0}B' -f ($Matches[1] -replace '[^\d]', '')) }
    if ($Text -eq '') { return $null }
    return $Text
}

function ConvertTo-SettableValue {
    # Scalars stay scalars, collections become string arrays, empty becomes $null.
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return $Value
    }
    if ($Value -is [string]) { return (ConvertFrom-DisplayString -Text $Value) }
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @(foreach ($item in $Value) {
            $converted = ConvertTo-SettableValue -Value $item
            if ($null -ne $converted -and "$converted" -ne '') { [string]$converted }
        })
        if ($items.Count -eq 0) { return $null }
        return ,[string[]]$items
    }
    $memberNames = @($Value.PSObject.Properties | Select-Object -ExpandProperty Name)
    if ($memberNames -contains 'IsUnlimited') {
        if ($Value.IsUnlimited) { return 'Unlimited' }
        return (ConvertTo-SettableValue -Value $Value.Value)
    }
    if (@($Value.PSObject.Methods | Select-Object -ExpandProperty Name) -contains 'ToBytes') {
        return ('{0}B' -f $Value.ToBytes())
    }
    return (ConvertFrom-DisplayString -Text $Value.ToString())
}

function Get-ValueList {
    # Always returns a flat array. Wrapping the output of ConvertTo-SettableValue in @() would nest the string
    # array it returns, and -contains would then compare against the array instead of its elements.
    param($Value)
    if ($null -eq $Value) { return ,@() }
    $flat = @()
    foreach ($item in @($Value)) { $flat += $item }
    return ,$flat
}

function Format-ErrorText {
    # Error messages can span several lines; inside a generated comment every line must start with #.
    param([string] $Text)
    return (($Text -split "`r?`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ' | '
}

function ConvertTo-PsString {
    # Quotes a string for the generated script. Plain strings get single quotes. Strings that carry a server
    # placeholder get double quotes with $TargetServer or $TargetFqdn inside, everything else escaped.
    param([string] $Text)
    if ($Text -notmatch '__TARGET_(SERVER|FQDN)__') {
        return ("'{0}'" -f ($Text -replace "'", "''"))
    }
    $escaped = $Text -replace '`', '``' -replace '"', '`"' -replace '\$', '`$'
    $escaped = $escaped -replace $script:FqdnToken, '$TargetFqdn' -replace $script:ShortToken, '$TargetServer'
    return ('"{0}"' -f $escaped)
}

function ConvertTo-PsLiteral {
    # Turns a normalised value into PowerShell source text.
    param($Value)
    if ($null -eq $Value) { return '$null' }
    if ($Value -is [bool]) { if ($Value) { return '$true' } else { return '$false' } }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [string]$Value }
    if ($Value -is [array]) {
        return ('@({0})' -f ((@($Value) | ForEach-Object { ConvertTo-PsString -Text ([string]$_) }) -join ', '))
    }
    return (ConvertTo-PsString -Text ([string]$Value))
}

function ConvertTo-PermissionGroupString {
    # "Custom" is reported when explicit ACEs exist but Set-ReceiveConnector refuses it as input.
    param($Value)
    if ($null -eq $Value) { return $null }
    $groups = @(("$Value" -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne 'Custom' })
    if ($groups.Count -eq 0) { return 'None' }
    return ($groups -join ', ')
}

function Convert-Binding {
    # Wildcards pass through, specific addresses are mapped with -IpAddressMap. Returns the bindings plus a list
    # of addresses that could not be mapped so the caller can flag them.
    param($Binding)
    $result = @()
    $unmapped = @()
    foreach ($entry in @($Binding)) {
        $ip = $null
        $port = $null
        if ($entry -match '^\[(.+)\]:(\d+)$') { $ip = $Matches[1]; $port = $Matches[2] }
        elseif ($entry -match '^(.+):(\d+)$') { $ip = $Matches[1]; $port = $Matches[2] }
        else { $result += [string]$entry; continue }

        if ($ip -eq '0.0.0.0' -or $ip -eq '::') { $result += [string]$entry; continue }
        if ($IpAddressMap.ContainsKey($ip)) {
            $newIp = [string]$IpAddressMap[$ip]
            if ($newIp -match ':') { $result += ('[{0}]:{1}' -f $newIp, $port) } else { $result += ('{0}:{1}' -f $newIp, $port) }
            continue
        }
        $result += [string]$entry
        $unmapped += $ip
    }
    return [pscustomobject]@{ Bindings = [string[]]$result; Unmapped = [string[]]$unmapped }
}

function Get-SettableParameterNames {
    # Parameters of the Set cmdlet minus common, identity and excluded ones, read from the cmdlet metadata.
    param([string] $CmdletName, [string[]] $Exclude = @())
    $command = Get-Command -Name $CmdletName -ErrorAction Stop
    $common = @([System.Management.Automation.PSCmdlet]::CommonParameters) +
              @([System.Management.Automation.PSCmdlet]::OptionalCommonParameters)
    $names = @()
    foreach ($parameterName in $command.Parameters.Keys) {
        if ($common -contains $parameterName -or $script:NeverCopy -contains $parameterName) { continue }
        if ($ExcludeParameter -contains $parameterName) { continue }
        $excluded = $false
        foreach ($pattern in $Exclude) {
            if ($parameterName -like $pattern) { $excluded = $true; break }
        }
        if (-not $excluded) { $names += $parameterName }
    }
    return ($names | Sort-Object)
}

function Add-SetCommand {
    <#
        Writes one Set command for a source object: the cmdlet, the identity, then every settable parameter with
        the source value, one per line with a backtick continuation, closed by -WhatIf:$WhatIf.
    #>
    param(
        [string] $Cmdlet,
        [string] $IdentityParameterName = 'Identity',
        [string] $IdentityText,
        $SourceObject,
        [string[]] $Exclude = @(),
        [string[]] $BindingProperties = @(),
        [hashtable] $ExtraArguments = @{}
    )
    if ($SourceObject -is [array]) { $SourceObject = $SourceObject[0] }
    $parameterNames = Get-SettableParameterNames -CmdletName $Cmdlet -Exclude $Exclude
    $propertyNames = @($SourceObject.PSObject.Properties | Select-Object -ExpandProperty Name)
    $argumentLines = @()
    $todo = @()

    foreach ($name in $parameterNames) {
        if ($propertyNames -notcontains $name) { continue }
        $value = ConvertTo-SettableValue -Value $SourceObject.$name
        if ($name -eq 'PermissionGroups') { $value = ConvertTo-PermissionGroupString -Value $value }
        if ($BindingProperties -contains $name -and $null -ne $value) {
            $mapped = Convert-Binding -Binding $value
            $value = $mapped.Bindings
            if ($mapped.Unmapped.Count -gt 0) {
                $todo += ('TODO {0}: bound to source address {1}, no entry in -IpAddressMap' -f $name, ($mapped.Unmapped -join ', '))
            }
        }
        if ($null -eq $value -and -not $IncludeNullValues) { continue }
        $value = Convert-ServerName -Value $value
        $argumentLines += ('    -{0} {1}' -f $name, (ConvertTo-PsLiteral -Value $value))
    }
    foreach ($key in @($ExtraArguments.Keys | Sort-Object)) {
        $argumentLines += ('    -{0} {1}' -f $key, (ConvertTo-PsLiteral -Value $ExtraArguments[$key]))
    }

    foreach ($entry in $todo) {
        Add-Line ('# {0}' -f $entry)
        Add-Warning ('{0} {1}: {2}' -f $Cmdlet, $IdentityText, $entry)
    }
    Add-Line ('{0} -{1} {2} `' -f $Cmdlet, $IdentityParameterName, $IdentityText)
    foreach ($argumentLine in $argumentLines) { Add-Line ($argumentLine + ' `') }
    Add-Line '    -WhatIf:$WhatIf'
    Add-Line
}

function Get-TargetIdentityText {
    # "EX01\owa (Default Web Site)" becomes "$TargetServer\owa (Default Web Site)" in the generated text.
    param([string] $SourceIdentity)
    return (ConvertTo-PsString -Text (Convert-ServerName -Value $SourceIdentity))
}

# ---------------------------------------------------------------------------
# Read the source and resolve the target names
# ---------------------------------------------------------------------------

if (-not (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    throw 'Exchange cmdlets are not available. Run this script from the Exchange Management Shell or import a remote Exchange session first.'
}

$sourceExchangeServer = Get-ExchangeServer -Identity $SourceServer -ErrorAction Stop
$script:SourceShort = [string]$sourceExchangeServer.Name
$script:SourceFqdn  = [string]$sourceExchangeServer.Fqdn

$targetShort = $TargetServer.Split('.')[0].ToUpperInvariant()
$targetFqdn = $null
$targetExchangeServer = Get-ExchangeServer -Identity $TargetServer -ErrorAction SilentlyContinue
if ($null -ne $targetExchangeServer) {
    $targetShort = [string]$targetExchangeServer.Name
    $targetFqdn = [string]$targetExchangeServer.Fqdn
}
elseif ($TargetServer -match '\.') { $targetFqdn = $TargetServer.ToLowerInvariant() }
elseif ($script:SourceFqdn -match '^[^.]+\.(.+)$') { $targetFqdn = ('{0}.{1}' -f $targetShort, $Matches[1]).ToLowerInvariant() }
else { $targetFqdn = $targetShort }

if ($targetShort -eq $script:SourceShort) { throw 'Source and target are the same server.' }

if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location).Path ('Set-ExchangeServerConfig_from_{0}_{1}.ps1' -f $script:SourceShort, (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

Write-Host ('Reading {0}, generating script for {1} ({2})' -f $script:SourceFqdn, $targetShort, $targetFqdn)

# ---------------------------------------------------------------------------
# Header of the generated script
# ---------------------------------------------------------------------------

Add-Line '<#'
Add-Line ('    Exchange server configuration of {0} ({1})' -f $script:SourceShort, $script:SourceFqdn)
Add-Line ('    Exported {0} with Export-ExchangeServerConfigScript.ps1 (jentech consulting)' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'))
Add-Line ('    Source build: {0}' -f $sourceExchangeServer.AdminDisplayVersion)
Add-Line
Add-Line '    Plain list of Exchange commands that reproduce the client access and transport configuration of the'
Add-Line '    source on the server named in $TargetServer. Run it in the Exchange Management Shell with Organization'
Add-Line '    Management rights. Review it first, remove what does not apply, and start with $WhatIf = $true.'
Add-Line
Add-Line '    Not included: DAG, databases, organization wide objects, config file customisations (web.config,'
Add-Line '    EdgeTransport.exe.config, OWA logon files), DNS and load balancer changes.'
Add-Line '#>'
Add-Line
Add-Line ("`$TargetServer = '{0}'" -f $targetShort)
Add-Line ("`$TargetFqdn   = '{0}'" -f $targetFqdn)
Add-Line '$WhatIf       = $true      # set to $false to apply'
Add-Line '$ErrorActionPreference = ''Continue'''
Add-Line
Add-Line 'if (-not (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue)) { throw ''Run this script in the Exchange Management Shell.'' }'
Add-Line 'Get-ExchangeServer -Identity $TargetServer -ErrorAction Stop | Out-Null'

# ---------------------------------------------------------------------------
# Certificates
# ---------------------------------------------------------------------------

Add-Section 'Certificates: export from the source, import on the target, enable the same services'
Add-Line '# Non exportable and self signed certificates are listed as comments. Enable-ExchangeCertificate -Force replaces'
Add-Line '# the default SMTP certificate on the target without prompting, which is the intent.'
Add-Line '$pfxPassword = Read-Host -AsSecureString -Prompt ''Password for the exported PFX files'''
Add-Line '$pfxFolder   = Join-Path $env:TEMP ''ExchangeCertificateCopy'''
Add-Line 'New-Item -Path $pfxFolder -ItemType Directory -Force | Out-Null'
Add-Line

try {
    $authThumbprints = @()
    try {
        $authConfig = Get-AuthConfig -ErrorAction Stop
        $authThumbprints = @($authConfig.CurrentCertificateThumbprint, $authConfig.NextCertificateThumbprint) | Where-Object { $_ }
    }
    catch { Add-Warning ('Get-AuthConfig failed: {0}' -f $_.Exception.Message) }

    foreach ($cert in @(Get-ExchangeCertificate -Server $script:SourceShort -ErrorAction Stop)) {
        $thumbprint = [string]$cert.Thumbprint
        $subject = [string]$cert.Subject
        $isAuthCert = $authThumbprints -contains $thumbprint
        $services = @(("$($cert.Services)" -split ',') | ForEach-Object { $_.Trim() } | Where-Object { @('IIS', 'SMTP', 'IMAP', 'POP') -contains $_ })

        if ([bool]$cert.IsSelfSigned -and -not $isAuthCert) {
            Add-Line ('# skipped, self signed: {0} {1} (services: {2})' -f $thumbprint, $subject, $cert.Services)
            continue
        }
        if ($cert.NotAfter -and ([datetime]$cert.NotAfter) -lt (Get-Date)) {
            Add-Line ('# skipped, expired {0}: {1} {2}' -f $cert.NotAfter, $thumbprint, $subject)
            continue
        }
        if (-not [bool]$cert.PrivateKeyExportable) {
            Add-Line ('# TODO private key not exportable, import by hand: {0} {1} (services: {2})' -f $thumbprint, $subject, ($services -join ','))
            Add-Warning ('Certificate {0} ({1}) has a non exportable private key.' -f $thumbprint, $subject)
            continue
        }

        Add-Line ('# {0} (expires {1}, domains: {2})' -f $subject, $cert.NotAfter, ((ConvertTo-SettableValue -Value $cert.CertificateDomains) -join ', '))
        Add-Line ("`$export = Export-ExchangeCertificate -Server '{0}' -Thumbprint '{1}' -BinaryEncoded -Password `$pfxPassword" -f $script:SourceShort, $thumbprint)
        Add-Line ('[System.IO.File]::WriteAllBytes((Join-Path $pfxFolder ''{0}.pfx''), $export.FileData)' -f $thumbprint)
        Add-Line ('Import-ExchangeCertificate -Server $TargetServer -FileData ([System.IO.File]::ReadAllBytes((Join-Path $pfxFolder ''{0}.pfx''))) -Password $pfxPassword -PrivateKeyExportable $true -WhatIf:$WhatIf' -f $thumbprint)
        if ($isAuthCert) {
            Add-Line '# OAuth certificate: imported only. Setup flags it for SMTP itself and it must never become the default SMTP certificate.'
        }
        elseif ($services.Count -gt 0) {
            Add-Line ('Enable-ExchangeCertificate -Server $TargetServer -Thumbprint ''{0}'' -Services ''{1}'' -Force -Confirm:$false -WhatIf:$WhatIf' -f $thumbprint, ($services -join ','))
        }
        Add-Line
    }
}
catch { Add-Line ('# ERROR reading certificates: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

# ---------------------------------------------------------------------------
# Server level settings
# ---------------------------------------------------------------------------

Add-Section 'Set-ExchangeServer (product key and static domain controllers are never exported)'
try {
    $server = Get-ExchangeServer -Identity $script:SourceShort -ErrorAction Stop
    if (-not [bool]$server.IsExchangeTrialEdition) {
        Add-Line ('# Source edition: {0}. License the target with Set-ExchangeServer -ProductKey before going live.' -f $server.Edition)
    }
    Add-SetCommand -Cmdlet 'Set-ExchangeServer' -IdentityText '$TargetServer' -SourceObject $server -Exclude @('ProductKey', 'Static*')
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

Add-Section 'Set-ClientAccessService'
try {
    $clientAccess = Get-ClientAccessService -Identity $script:SourceShort -IncludeAlternateServiceAccountCredentialStatus -ErrorAction Stop
    Add-SetCommand -Cmdlet 'Set-ClientAccessService' -IdentityText '$TargetServer' -SourceObject $clientAccess `
        -Exclude @('AlternateServiceAccountCredential', 'RemoveAlternateServiceAccountCredentials', 'CleanUpInvalidAlternateServiceAccountCredentials', 'Array')
    $asa = $clientAccess.AlternateServiceAccountConfiguration
    if ($null -ne $asa -and $null -ne $asa.EffectiveCredentials -and @($asa.EffectiveCredentials).Count -gt 0) {
        Add-Line ('# TODO alternate service account is configured on the source. Run from the Exchange Scripts folder:')
        Add-Line ('#   .\RollAlternateServiceAccountPassword.ps1 -ToSpecificServer $TargetServer -CopyFrom ''{0}''' -f $script:SourceShort)
        Add-Line
    }
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

Add-Section 'Set-MailboxServer'
try {
    $mailboxServer = Get-MailboxServer -Identity $script:SourceShort -ErrorAction Stop
    Add-SetCommand -Cmdlet 'Set-MailboxServer' -IdentityText '$TargetServer' -SourceObject $mailboxServer `
        -Exclude ($script:PathExclusions + @('DatabaseCopyActivationDisabledAndMoveNow', 'WorkloadManagementPolicy'))
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

Add-Section 'Set-MalwareFilteringServer'
try {
    $malware = Get-MalwareFilteringServer -Identity $script:SourceShort -ErrorAction Stop
    Add-SetCommand -Cmdlet 'Set-MalwareFilteringServer' -IdentityText '$TargetServer' -SourceObject $malware -Exclude @('ForceRescan')
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

# ---------------------------------------------------------------------------
# Virtual directories and Outlook Anywhere
# ---------------------------------------------------------------------------

$directoryTypes = @(
    @{ Get = 'Get-OwaVirtualDirectory';          Set = 'Set-OwaVirtualDirectory';          Exclude = @() },
    @{ Get = 'Get-EcpVirtualDirectory';          Set = 'Set-EcpVirtualDirectory';          Exclude = @() },
    @{ Get = 'Get-WebServicesVirtualDirectory';  Set = 'Set-WebServicesVirtualDirectory';  Exclude = @() },
    @{ Get = 'Get-ActiveSyncVirtualDirectory';   Set = 'Set-ActiveSyncVirtualDirectory';   Exclude = @('ActiveSyncServer') },
    @{ Get = 'Get-OabVirtualDirectory';          Set = 'Set-OabVirtualDirectory';          Exclude = @() },
    @{ Get = 'Get-MapiVirtualDirectory';         Set = 'Set-MapiVirtualDirectory';         Exclude = @() },
    @{ Get = 'Get-PowerShellVirtualDirectory';   Set = 'Set-PowerShellVirtualDirectory';   Exclude = @() },
    @{ Get = 'Get-AutodiscoverVirtualDirectory'; Set = 'Set-AutodiscoverVirtualDirectory'; Exclude = @() }
)
foreach ($type in $directoryTypes) {
    Add-Section $type.Set
    try {
        foreach ($directory in @(& $type.Get -Server $script:SourceShort -ErrorAction Stop)) {
            Add-SetCommand -Cmdlet $type.Set -IdentityText (Get-TargetIdentityText -SourceIdentity ([string]$directory.Identity)) -SourceObject $directory -Exclude $type.Exclude
        }
    }
    catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }
}

Add-Section 'Set-OutlookAnywhere'
try {
    foreach ($item in @(Get-OutlookAnywhere -Server $script:SourceShort -ErrorAction Stop)) {
        Add-SetCommand -Cmdlet 'Set-OutlookAnywhere' -IdentityText (Get-TargetIdentityText -SourceIdentity ([string]$item.Identity)) -SourceObject $item `
            -Exclude @('ClientAuthenticationMethod', 'DefaultAuthenticationMethod')
    }
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

# ---------------------------------------------------------------------------
# Transport services, IMAP and POP
# ---------------------------------------------------------------------------

foreach ($pair in @(
    @{ Get = 'Get-TransportService';         Set = 'Set-TransportService' },
    @{ Get = 'Get-FrontendTransportService'; Set = 'Set-FrontendTransportService' },
    @{ Get = 'Get-MailboxTransportService';  Set = 'Set-MailboxTransportService' })) {
    Add-Section ('{0} (paths {1})' -f $pair.Set, $(if ($IncludePaths) { 'included' } else { 'left out, use -IncludePaths' }))
    try {
        $transport = & $pair.Get -Identity $script:SourceShort -ErrorAction Stop
        Add-SetCommand -Cmdlet $pair.Set -IdentityText '$TargetServer' -SourceObject $transport `
            -Exclude ($script:PathExclusions + @('InternalDNSAdapterGuid', 'ExternalDNSAdapterGuid'))
    }
    catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }
}

foreach ($protocol in @(
    @{ Get = 'Get-ImapSettings'; Set = 'Set-ImapSettings'; Services = @('MSExchangeIMAP4', 'MSExchangeIMAP4BE') },
    @{ Get = 'Get-PopSettings';  Set = 'Set-PopSettings';  Services = @('MSExchangePOP3', 'MSExchangePOP3BE') })) {
    Add-Section ('{0} and service startup type' -f $protocol.Set)
    try {
        $settings = & $protocol.Get -Server $script:SourceShort -ErrorAction Stop
        Add-SetCommand -Cmdlet $protocol.Set -IdentityParameterName 'Server' -IdentityText '$TargetServer' -SourceObject $settings `
            -Exclude $script:PathExclusions -BindingProperties @('SSLBindings', 'UnencryptedOrTLSBindings')
        foreach ($serviceName in $protocol.Services) {
            try {
                $service = Get-CimInstance -ClassName Win32_Service -ComputerName $script:SourceFqdn -Filter ("Name='{0}'" -f $serviceName) -ErrorAction Stop
                if ($null -eq $service) { continue }
                $startupType = switch ($service.StartMode) { 'Auto' { 'Automatic' } 'Manual' { 'Manual' } default { 'Disabled' } }
                Add-Line ('Set-Service -ComputerName $TargetFqdn -Name ''{0}'' -StartupType {1} -WhatIf:$WhatIf' -f $serviceName, $startupType)
                if ($startupType -eq 'Automatic') { Add-Line ('if (-not $WhatIf) {{ Get-Service -ComputerName $TargetFqdn -Name ''{0}'' | Start-Service }}' -f $serviceName) }
            }
            catch { Add-Line ('# TODO could not read the startup type of {0} on the source: {1}' -f $serviceName, (Format-ErrorText -Text $_.Exception.Message)) }
        }
        Add-Line
    }
    catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }
}

# ---------------------------------------------------------------------------
# Event log levels and setting overrides
# ---------------------------------------------------------------------------

Add-Section 'Set-EventLogLevel (levels raised above Lowest on the source)'
try {
    $raised = @(Get-EventLogLevel -Server $script:SourceShort -ErrorAction Stop | Where-Object { "$($_.EventLevel)" -ne 'Lowest' })
    if ($raised.Count -eq 0) { Add-Line '# none' }
    foreach ($level in $raised) {
        $parts = ([string]$level.Identity) -split '\\'
        $category = if ($parts.Count -ge 3) { ($parts | Select-Object -Skip 1) -join '\' } else { [string]$level.Identity }
        Add-Line ('Set-EventLogLevel -Identity "$TargetServer\{0}" -Level {1} -WhatIf:$WhatIf' -f $category, $level.EventLevel)
    }
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

Add-Section 'Set-SettingOverride (overrides scoped to the source server get the target added)'
try {
    $found = 0
    foreach ($override in @(Get-SettingOverride -ErrorAction Stop)) {
        $servers = Get-ValueList -Value (ConvertTo-SettableValue -Value $override.Server)
        if ($servers.Count -eq 0 -or $servers -notcontains $script:SourceShort) { continue }
        $found++
        $serverList = (@($servers | ForEach-Object { "'{0}'" -f $_ }) + '$TargetServer') -join ', '
        Add-Line ('# {0}: {1}' -f $override.Name, ((ConvertTo-SettableValue -Value $override.Parameters) -join '; '))
        Add-Line ('Set-SettingOverride -Identity ''{0}'' -Server @({1}) -Confirm:$false -WhatIf:$WhatIf' -f (([string]$override.Identity) -replace "'", "''"), $serverList)
    }
    if ($found -eq 0) { Add-Line '# none' }
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

# ---------------------------------------------------------------------------
# Receive connectors
# ---------------------------------------------------------------------------

Add-Section 'Receive connectors'
Add-Line '# Default connectors already exist on the target under the target name, only Set-ReceiveConnector is written'
Add-Line '# for them. Custom connectors get New-ReceiveConnector first. Add-ADPermission lines reproduce the explicit'
Add-Line '# ACEs of the source; most are granted by the permission groups anyway and adding them twice is harmless.'
Add-Line
try {
    foreach ($connector in @(Get-ReceiveConnector -Server $script:SourceShort -ErrorAction Stop)) {
        $connectorName = [string]$connector.Name
        $identityText = Get-TargetIdentityText -SourceIdentity ([string]$connector.Identity)
        $isDefault = $connectorName -match ('(?<![\w-])' + [regex]::Escape($script:SourceShort) + '(?![\w-])')
        Add-Line ('# --- {0} ({1}) ---' -f $connectorName, $connector.TransportRole)

        if (-not $isDefault) {
            $mapped = Convert-Binding -Binding (Get-ValueList -Value (ConvertTo-SettableValue -Value $connector.Bindings))
            if ($mapped.Unmapped.Count -gt 0) {
                Add-Line ('# TODO Bindings use source address {0}, no entry in -IpAddressMap' -f ($mapped.Unmapped -join ', '))
                Add-Warning ('Receive connector {0} is bound to {1}.' -f $connectorName, ($mapped.Unmapped -join ', '))
            }
            $remoteRanges = [string[]](Get-ValueList -Value (ConvertTo-SettableValue -Value $connector.RemoteIPRanges))
            Add-Line ('New-ReceiveConnector -Name {0} -Server $TargetServer -TransportRole {1} -Custom `' -f (ConvertTo-PsString -Text (Convert-ServerName -Value $connectorName)), $connector.TransportRole)
            Add-Line ('    -Bindings {0} `' -f (ConvertTo-PsLiteral -Value $mapped.Bindings))
            Add-Line ('    -RemoteIPRanges {0} `' -f (ConvertTo-PsLiteral -Value $remoteRanges))
            Add-Line '    -Confirm:$false -WhatIf:$WhatIf | Out-Null'
            Add-Line
        }

        Add-SetCommand -Cmdlet 'Set-ReceiveConnector' -IdentityText $identityText -SourceObject $connector `
            -Exclude @('TransportRole', 'Usage') -BindingProperties @('Bindings')

        try {
            # The distinguished name is used on purpose: Get-ADPermission does not resolve the Server\Name form reliably.
            foreach ($ace in @(Get-ADPermission -Identity ([string]$connector.DistinguishedName) -ErrorAction Stop | Where-Object { -not [bool]$_.IsInherited })) {
                $arguments = @()
                $arguments += ('-User {0}' -f (ConvertTo-PsString -Text (Convert-ServerName -Value ([string]$ace.User))))
                $extendedRights = ConvertTo-SettableValue -Value $ace.ExtendedRights
                $accessRights = @()
                foreach ($right in (Get-ValueList -Value (ConvertTo-SettableValue -Value $ace.AccessRights))) {
                    foreach ($part in ($right -split ',')) { if ($part.Trim()) { $accessRights += $part.Trim() } }
                }
                $accessRights = @($accessRights | Where-Object { $_ -ne 'ExtendedRight' -or $null -eq $extendedRights })
                if ($null -ne $extendedRights) { $arguments += ('-ExtendedRights {0}' -f (ConvertTo-PsLiteral -Value ([string[]]$extendedRights))) }
                if ($accessRights.Count -gt 0) { $arguments += ('-AccessRights {0}' -f (ConvertTo-PsLiteral -Value ([string[]]$accessRights))) }
                if ([bool]$ace.Deny) { $arguments += '-Deny' }
                $properties = ConvertTo-SettableValue -Value $ace.Properties
                if ($null -ne $properties) { $arguments += ('-Properties {0}' -f (ConvertTo-PsLiteral -Value ([string[]]$properties))) }
                $inheritance = ConvertTo-SettableValue -Value $ace.InheritanceType
                if ($null -ne $inheritance -and "$inheritance" -ne 'None') { $arguments += ('-InheritanceType {0}' -f $inheritance) }
                Add-Line ('Add-ADPermission -Identity (Get-ReceiveConnector -Identity {0}).DistinguishedName {1} -Confirm:$false -WhatIf:$WhatIf | Out-Null' -f $identityText, ($arguments -join ' '))
            }
        }
        catch { Add-Line ('# ERROR reading permissions: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }
        Add-Line
    }
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

# ---------------------------------------------------------------------------
# Send connectors, transport agents, hybrid (manual steps)
# ---------------------------------------------------------------------------

Add-Section 'Send connectors: LIVE ROUTING CHANGE, remove the leading # once the target is validated'
try {
    $found = 0
    foreach ($connector in @(Get-SendConnector -ErrorAction Stop)) {
        $servers = Get-ValueList -Value (ConvertTo-SettableValue -Value $connector.SourceTransportServers)
        if ($servers -notcontains $script:SourceShort) { continue }
        $found++
        $serverList = (@($servers | ForEach-Object { "'{0}'" -f $_ }) + '$TargetServer') -join ', '
        Add-Line ('# Set-SendConnector -Identity ''{0}'' -SourceTransportServers @({1}) -Confirm:$false -WhatIf:$WhatIf' -f (([string]$connector.Identity) -replace "'", "''"), $serverList)
    }
    if ($found -eq 0) { Add-Line '# the source is not a source transport server on any send connector' }
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

Add-Section 'Transport agents on the source (manual: copy the assembly, then run the Install line)'
try {
    $found = 0
    foreach ($transportService in @('Hub', 'FrontEnd')) {
        foreach ($agent in @(Get-TransportAgent -Server $script:SourceShort -TransportService $transportService -ErrorAction Stop)) {
            $found++
            $detail = $agent
            try { $detail = Get-TransportAgent -Identity ([string]$agent.Identity) -Server $script:SourceShort -TransportService $transportService -ErrorAction Stop | Select-Object -First 1 } catch { }
            Add-Line ('# {0} | {1} | enabled {2} | priority {3} | {4}' -f $transportService, $agent.Identity, $agent.Enabled, $agent.Priority, $detail.AssemblyPath)
            if ("$($detail.AssemblyPath)" -notmatch '\\Microsoft\\Exchange Server\\V15\\(Bin|TransportRoles)\\' -and "$($detail.AssemblyPath)" -ne '') {
                Add-Line ('#   Install-TransportAgent -Name ''{0}'' -TransportService {1} -TransportAgentFactory ''{2}'' -AssemblyPath ''{3}''; Enable-TransportAgent -Identity ''{0}'' -TransportService {1}; Set-TransportAgent -Identity ''{0}'' -TransportService {1} -Priority {4}' -f $agent.Identity, $transportService, $detail.TransportAgentFactory, $detail.AssemblyPath, $agent.Priority)
            }
        }
    }
    if ($found -eq 0) { Add-Line '# none' }
}
catch { Add-Line ('# ERROR: {0}' -f (Format-ErrorText -Text $_.Exception.Message)); Add-Warning $_.Exception.Message }

Add-Section 'Hybrid configuration (manual: rerun the Hybrid Configuration Wizard and select the target as well)'
try {
    $hybrid = Get-HybridConfiguration -ErrorAction Stop
    foreach ($property in @('SendingTransportServers', 'ReceivingTransportServers', 'EdgeTransportServers')) {
        $servers = Get-ValueList -Value (ConvertTo-SettableValue -Value $hybrid.$property)
        if ($servers -contains $script:SourceShort) {
            Add-Line ('# {0} currently: {1}. Add $TargetServer via the wizard or Set-HybridConfiguration -{0}.' -f $property, ($servers -join ', '))
        }
    }
}
catch { Add-Line '# no hybrid configuration found' }

Add-Section 'Config files to compare by hand (not scriptable, a cumulative update overwrites them)'
foreach ($file in @('Bin\EdgeTransport.exe.config', 'Bin\MSExchangeMailboxReplication.exe.config', 'TransportRoles\Shared\agents.config',
                    'FrontEnd\HttpProxy\SharedWebConfig.config', 'FrontEnd\HttpProxy\owa\web.config', 'FrontEnd\HttpProxy\owa\auth\logon.aspx',
                    'FrontEnd\HttpProxy\ews\web.config', 'FrontEnd\HttpProxy\rpc\web.config', 'FrontEnd\HttpProxy\mapi\web.config',
                    'ClientAccess\SharedWebConfig.config', 'ClientAccess\Owa\web.config', 'ClientAccess\exchweb\ews\web.config')) {
    Add-Line ('# {0}' -f $file)
}

Add-Section 'After the run'
Add-Line '# iisreset on the target, Restart-Service MSExchangeTransport and MSExchangeFrontEndTransport, restart the IMAP and'
Add-Line '# POP services when they are used. Then Test-OutlookConnectivity, Test-OwaConnectivity, Test-ActiveSyncConnectivity'
Add-Line '# before the target joins the load balancer pool.'

# ---------------------------------------------------------------------------
# Write the generated script
# ---------------------------------------------------------------------------

$script:Lines | Set-Content -Path $OutputPath -Encoding UTF8
$resolved = (Resolve-Path $OutputPath).Path
Write-Host ('Generated script: {0} ({1} lines)' -f $resolved, $script:Lines.Count)
if ($script:Warnings.Count -gt 0) {
    Write-Host ('{0} item(s) need attention, search the generated script for TODO:' -f $script:Warnings.Count) -ForegroundColor Yellow
    foreach ($warning in $script:Warnings) { Write-Host ('  - {0}' -f $warning) -ForegroundColor Yellow }
}
