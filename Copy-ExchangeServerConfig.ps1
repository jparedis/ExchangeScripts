#Requires -Version 5.1

<#
.SYNOPSIS
    Copies the client access and transport configuration of an existing Exchange server to a newly installed one.

.DESCRIPTION
    Built for the Exchange Server Subscription Edition migration scenario: the existing server was upgraded in place,
    new servers are added to the organization (for a new DAG) and they must end up with exactly the same client
    access and transport configuration as the existing server.

    The script reads the relevant settings from the source server, compares them with the target server and only
    applies what differs. Every planned or applied change is recorded and written to a CSV file in the output
    folder, together with a JSON snapshot of the source and target configuration taken before any change.
    Run it with -WhatIf first to get a full report without touching the target. The script is idempotent: a second
    run reports "in sync" for everything that was applied.

    Areas (parameter -Area, default is all of them, executed in this order):

      Certificates        Exports every certificate on the source that is not self signed (plus the OAuth
                          certificate referenced by Get-AuthConfig when it is missing on the target), imports it on
                          the target and enables the same services (IIS, SMTP, IMAP, POP). Certificates whose
                          private key is not exportable are reported so you can handle them manually.
      ExchangeServer      Set-ExchangeServer settings (internet web proxy, error reporting, monitoring group, ...).
                          The product key and the static domain controller settings are never copied.
      ClientAccess        Set-ClientAccessService: Autodiscover internal URI and site scope. An alternate service
                          account (Kerberos) on the source is reported, it cannot be copied by this script.
      VirtualDirectories  OWA, ECP, EWS, ActiveSync, OAB, MAPI, PowerShell and Autodiscover virtual directories:
                          URLs, authentication methods and every other property the Set cmdlet accepts.
      OutlookAnywhere     Hostnames, authentication methods and SSL settings.
      Transport           Set-TransportService, Set-FrontendTransportService and Set-MailboxTransportService: DNS
                          settings, limits, logging retention, message tracking and so on. Log and queue paths are
                          skipped unless -IncludePaths is given, because the disk layout may differ.
      PopImap             Set-ImapSettings and Set-PopSettings, plus the startup type of the IMAP and POP services.
      MailboxServer       Set-MailboxServer settings (activation policy, maximum active databases, schedules, ...).
      Malware             Set-MalwareFilteringServer.
      EventLogLevels      Diagnostic logging levels that are raised above Lowest on the source.
      SettingOverrides    Setting overrides scoped to the source server get the target server added.
      ReceiveConnectors   Every receive connector on the source. Default connectors are matched by name (source
                          server name replaced by the target server name) and their settings converged. Custom
                          connectors (relay connectors and so on) are created on the target. Explicit AD permissions
                          on the connectors (for example ms-Exch-SMTP-Accept-Any-Recipient for anonymous relay) are
                          copied as well.
      SendConnectors      The target server is added as source transport server to every send connector that lists
                          the source server. This is a live routing change, so it is deliberately the last write.
      TransportAgents     Report only: transport agents present on the source but not on the target.
      ConfigFiles         Report only: compares web.config, EdgeTransport.exe.config, SharedWebConfig.config,
                          agents.config and other files that administrators commonly customise, plus the OWA logon
                          customisation folder and the OWA theme folders. Differences are saved with a line diff in
                          the output folder. These files are never written to the target automatically, because a
                          cumulative update replaces them anyway and the customisation must be documented.
      Iis                 HTTP redirect and SSL settings on the root of the Default Web Site (the classic redirect
                          to /owa) are compared and applied. Site bindings are compared and reported.
      Hybrid              Report only: whether the source is listed as sending, receiving or edge transport server
                          in the hybrid configuration.

    Server specific values are translated: the source server FQDN and NetBIOS name inside string values are
    replaced by those of the target (connector names, banners, InternalNLBBypassUrl, SPN lists, ...). Use
    -NoServerNameSubstitution to disable this. IP addresses in connector or protocol bindings are translated with
    the -IpAddressMap hashtable; a binding on a specific IP address without a mapping is reported and skipped.

    Requirements: run from the Exchange Management Shell (or a remote PowerShell session to an Exchange server)
    with an account that holds Organization Management. The config file, IIS and service checks additionally need
    access to the admin share (C$) and WinRM of both servers. Windows PowerShell 5.1.

    Not in scope: DAG creation, mailbox databases and database copies, organization wide objects (accepted
    domains, email address policies, transport rules, OWA mailbox policies), DNS records and load balancer pools.

.PARAMETER SourceServer
    Name of the existing Exchange server whose configuration is the reference.

.PARAMETER TargetServer
    Name of the newly installed Exchange server that must receive the configuration.

.PARAMETER Area
    One or more areas to process. Defaults to all areas. See the description for the list and the order.

.PARAMETER SkipArea
    One or more areas to leave out. Useful to postpone SendConnectors until the target is fully tested.

.PARAMETER IpAddressMap
    Hashtable that maps IP addresses of the source server to those of the target server, used for receive connector
    bindings and IMAP or POP bindings that are bound to a specific address. Wildcard bindings (0.0.0.0 and ::)
    need no mapping. Example: @{ '10.0.0.11' = '10.0.0.12'; '10.0.0.111' = '10.0.0.112' }

.PARAMETER CertificatePassword
    Password used to protect the exported certificates. When given, every exported certificate is also saved as a
    PFX file in the output folder. When omitted, a random password is used for the in memory transfer only and no
    PFX files are written to disk.

.PARAMETER OutputPath
    Folder for the transcript, the change report (changes.csv), the configuration snapshots and the config file
    diffs. Defaults to a timestamped folder in the current directory.

.PARAMETER IncludePaths
    Also copy log, queue and pickup directory paths (Transport, PopImap and MailboxServer areas). Only do this when
    the target has the same disk layout as the source.

.PARAMETER NoServerNameSubstitution
    Do not replace the source server name by the target server name inside string values.

.PARAMETER AdditionalConfigFile
    Extra files (relative to the Exchange installation folder) to include in the ConfigFiles comparison.

.EXAMPLE
    .\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -WhatIf

    Full report of every difference between EX01 and EX02 without changing anything.

.EXAMPLE
    .\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -SkipArea SendConnectors -CertificatePassword (Read-Host -AsSecureString 'PFX password')

    Copies everything except the send connector membership, and keeps a PFX copy of every transferred certificate
    in the output folder.

.EXAMPLE
    .\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -Area ReceiveConnectors -IpAddressMap @{ '10.0.0.11' = '10.0.0.12' }

    Only the receive connectors, with the relay connector that is bound to 10.0.0.11 on EX01 recreated on
    10.0.0.12 on EX02.

.EXAMPLE
    .\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -Area SendConnectors

    Final step once EX02 is tested: add it as source transport server to the send connectors.

.NOTES
    jentech consulting
    After the run: iisreset on the target when virtual directory or Outlook Anywhere settings changed, restart the
    transport services when connector settings changed, restart the IMAP and POP services when their settings
    changed. The script prints these hints at the end based on what it applied.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string] $SourceServer,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string] $TargetServer,

    [ValidateSet('Certificates', 'ExchangeServer', 'ClientAccess', 'VirtualDirectories', 'OutlookAnywhere', 'Transport',
                 'PopImap', 'MailboxServer', 'Malware', 'EventLogLevels', 'SettingOverrides', 'ReceiveConnectors',
                 'SendConnectors', 'TransportAgents', 'ConfigFiles', 'Iis', 'Hybrid')]
    [string[]] $Area,

    [ValidateSet('Certificates', 'ExchangeServer', 'ClientAccess', 'VirtualDirectories', 'OutlookAnywhere', 'Transport',
                 'PopImap', 'MailboxServer', 'Malware', 'EventLogLevels', 'SettingOverrides', 'ReceiveConnectors',
                 'SendConnectors', 'TransportAgents', 'ConfigFiles', 'Iis', 'Hybrid')]
    [string[]] $SkipArea = @(),

    [hashtable] $IpAddressMap = @{},

    [System.Security.SecureString] $CertificatePassword,

    [string] $OutputPath,

    [switch] $IncludePaths,

    [switch] $NoServerNameSubstitution,

    [string[]] $AdditionalConfigFile = @()
)

# The areas in execution order. Certificates come first so that connector and protocol settings that reference a
# certificate resolve on the target. Send connectors come last because they change live mail routing.
$script:AllAreas = @('Certificates', 'ExchangeServer', 'ClientAccess', 'VirtualDirectories', 'OutlookAnywhere',
                     'Transport', 'PopImap', 'MailboxServer', 'Malware', 'EventLogLevels', 'SettingOverrides',
                     'ReceiveConnectors', 'SendConnectors', 'TransportAgents', 'ConfigFiles', 'Iis', 'Hybrid')

$script:Changes      = New-Object System.Collections.Generic.List[object]
$script:Snapshots    = [ordered]@{}
$script:RestartHints = New-Object System.Collections.Generic.List[string]

# Parameters that are never copied, whatever the cmdlet: identity, remoting and confirmation parameters.
$script:NeverCopy = @('Identity', 'Server', 'DomainController', 'Force', 'Confirm', 'WhatIf', 'AsJob', 'Name')

# Path like parameters are only copied with -IncludePaths, the target may have another disk layout.
$script:PathExclusions = @('*Path', '*Location', '*Directory')
if ($IncludePaths) { $script:PathExclusions = @() }

# ---------------------------------------------------------------------------
# Logging and change tracking
# ---------------------------------------------------------------------------

function Write-Log {
    param(
        [string] $Message,
        [ValidateSet('Info', 'Change', 'Warn', 'Error', 'Section')]
        [string] $Level = 'Info'
    )
    $timestamp = Get-Date -Format 'HH:mm:ss'
    switch ($Level) {
        'Section' { Write-Host ''; Write-Host ('=== {0} ===' -f $Message) -ForegroundColor Cyan }
        'Change'  { Write-Host ('[{0}] {1}' -f $timestamp, $Message) -ForegroundColor Yellow }
        'Warn'    { Write-Host ('[{0}] WARNING: {1}' -f $timestamp, $Message) -ForegroundColor Magenta }
        'Error'   { Write-Host ('[{0}] ERROR: {1}' -f $timestamp, $Message) -ForegroundColor Red }
        default   { Write-Host ('[{0}] {1}' -f $timestamp, $Message) }
    }
}

function Format-ReportValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [array]) { return (($Value | ForEach-Object { [string]$_ }) -join '; ') }
    return [string]$Value
}

function Add-Change {
    # Every decision the script takes ends up here, so changes.csv is the complete audit trail of the run.
    param(
        [string] $Area,
        [string] $Identity,
        [string] $Property,
        $SourceValue,
        $TargetValue,
        [ValidateSet('Applied', 'WhatIf', 'Failed', 'Skipped', 'ReviewRequired', 'Info')]
        [string] $Status,
        [string] $Message = ''
    )
    $record = [pscustomobject]@{
        Area              = $Area
        Identity          = $Identity
        Property          = $Property
        SourceValue       = Format-ReportValue $SourceValue
        TargetValueBefore = Format-ReportValue $TargetValue
        Status            = $Status
        Message           = $Message
    }
    $script:Changes.Add($record)

    $line = '{0} | {1} | {2}' -f $Identity, $Property, $Status
    if ($Status -ne 'Info') {
        $line += (' | source: {0} | target: {1}' -f $record.SourceValue, $record.TargetValueBefore)
    }
    if ($Message) { $line += (' | {0}' -f $Message) }

    switch ($Status) {
        'Applied'        { Write-Log -Level Change -Message $line }
        'WhatIf'         { Write-Log -Level Change -Message $line }
        'Failed'         { Write-Log -Level Error -Message $line }
        'ReviewRequired' { Write-Log -Level Warn -Message $line }
        default          { Write-Log -Level Info -Message $line }
    }
}

function Get-PendingStatus {
    # Status to record when ShouldProcess returned false: WhatIf run, or the operator declined a confirmation.
    if ($WhatIfPreference) { return 'WhatIf' }
    return 'Skipped'
}

function Save-Snapshot {
    param([string] $Area, $Source, $Target)
    if (-not $script:Snapshots.Contains($Area)) {
        $script:Snapshots[$Area] = [ordered]@{ Source = @(); Target = @() }
    }
    if ($null -ne $Source) { $script:Snapshots[$Area].Source += @($Source) }
    if ($null -ne $Target) { $script:Snapshots[$Area].Target += @($Target) }
}

# ---------------------------------------------------------------------------
# Value conversion
# ---------------------------------------------------------------------------

function Convert-ServerName {
    # Replaces the source server FQDN and NetBIOS name inside a string (or each string of an array) by the target
    # names. The NetBIOS match is bounded so that EX01 never matches inside EX010.
    param($Value)
    if ($NoServerNameSubstitution -or $null -eq $Value) { return $Value }
    if ($Value -is [string]) {
        $result = $Value
        if ($script:SourceFqdn -and $script:TargetFqdn) {
            $result = [regex]::Replace($result, [regex]::Escape($script:SourceFqdn), $script:TargetFqdn, 'IgnoreCase')
        }
        $pattern = '(?<![\w-])' + [regex]::Escape($script:SourceShort) + '(?![\w-])'
        $result = [regex]::Replace($result, $pattern, $script:TargetShort, 'IgnoreCase')
        return $result
    }
    if ($Value -is [array]) {
        $converted = @(foreach ($item in $Value) { Convert-ServerName -Value $item })
        return ,[string[]]$converted
    }
    return $Value
}

function ConvertFrom-DisplayString {
    # Exchange objects arrive deserialized in the management shell, so sizes look like "35 MB (36,700,160 bytes)".
    # The Set cmdlets do not parse that display format, the byte count with a B suffix is what they accept.
    param([string] $Text)
    if ($Text -match '^\s*unlimited\s*$') { return 'Unlimited' }
    if ($Text -match '\(([\d,\.\s]+)\s*bytes\)') { return ('{0}B' -f ($Matches[1] -replace '[^\d]', '')) }
    if ($Text -eq '') { return $null }
    return $Text
}

function ConvertTo-SettableValue {
    # Normalises a property value (live or deserialized) into something a Set cmdlet accepts and that can be compared
    # as a string: scalars stay scalars, collections become string arrays, empty becomes $null.
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

function Get-CompareKey {
    # Order and case insensitive string key, used to decide whether source and target differ.
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [array]) {
        return ((@($Value | ForEach-Object { ([string]$_).Trim() }) | Sort-Object) -join '|').ToLowerInvariant()
    }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function ConvertTo-PermissionGroupString {
    # PermissionGroups reports "Custom" when explicit ACEs exist on the connector, but Set-ReceiveConnector refuses
    # that value. The explicit ACEs themselves are copied separately by Sync-ReceiveConnectorPermission.
    param($Value)
    if ($null -eq $Value) { return $null }
    $groups = @(("$Value" -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne 'Custom' })
    if ($groups.Count -eq 0) { return 'None' }
    return ($groups -join ', ')
}

function Convert-Binding {
    # Translates "ip:port" bindings. Wildcards pass through, specific addresses need an entry in -IpAddressMap.
    # Returns $null when one of the addresses cannot be mapped, the caller reports and skips in that case.
    param($Binding)
    if ($null -eq $Binding) { return $null }
    $result = @()
    foreach ($entry in @($Binding)) {
        $ip = $null
        $port = $null
        if ($entry -match '^\[(.+)\]:(\d+)$') { $ip = $Matches[1]; $port = $Matches[2] }
        elseif ($entry -match '^(.+):(\d+)$') { $ip = $Matches[1]; $port = $Matches[2] }
        else { $result += [string]$entry; continue }

        if ($ip -eq '0.0.0.0' -or $ip -eq '::') { $result += [string]$entry; continue }
        if ($IpAddressMap.ContainsKey($ip)) {
            $newIp = [string]$IpAddressMap[$ip]
            if ($newIp -match ':') { $result += ('[{0}]:{1}' -f $newIp, $port) }
            else { $result += ('{0}:{1}' -f $newIp, $port) }
            continue
        }
        return $null
    }
    return ,[string[]]$result
}

function Get-SettableParameterNames {
    # Every parameter of the Set cmdlet that is not a common, identity or explicitly excluded parameter. Working from
    # the cmdlet metadata keeps the script complete when a cumulative update adds parameters.
    param([string] $CmdletName, [string[]] $Exclude = @())
    $command = Get-Command -Name $CmdletName -ErrorAction Stop
    $common = @([System.Management.Automation.PSCmdlet]::CommonParameters) +
              @([System.Management.Automation.PSCmdlet]::OptionalCommonParameters)
    $names = @()
    foreach ($parameterName in $command.Parameters.Keys) {
        if ($common -contains $parameterName -or $script:NeverCopy -contains $parameterName) { continue }
        $excluded = $false
        foreach ($pattern in $Exclude) {
            if ($parameterName -like $pattern) { $excluded = $true; break }
        }
        if (-not $excluded) { $names += $parameterName }
    }
    return $names
}

# ---------------------------------------------------------------------------
# Generic settings synchronisation
# ---------------------------------------------------------------------------

function Sync-ObjectSettings {
    <#
        Compares a source object with its counterpart on the target and calls the Set cmdlet with only the
        properties that differ. When the bulk call fails (one deprecated or rejected parameter is enough for
        Exchange to reject the whole call) it retries property by property so that one failure does not block the
        rest. TargetObject may be $null when the object does not exist yet on the target, every property is then a
        difference.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string] $Area,
        [Parameter(Mandatory = $true)] $SourceObject,
        $TargetObject,
        [string] $SetCmdlet,
        [string] $TargetIdentity,
        [string] $IdentityParameterName = 'Identity',
        [string[]] $ExcludeProperties = @(),
        [hashtable] $CoupledProperties = @{},
        [string[]] $BindingProperties = @(),
        [string] $RestartHint
    )

    if ($SourceObject -is [array]) { $SourceObject = $SourceObject[0] }
    if ($TargetObject -is [array]) { $TargetObject = $TargetObject[0] }
    $parameterNames = Get-SettableParameterNames -CmdletName $SetCmdlet -Exclude $ExcludeProperties
    $sourcePropertyNames = @($SourceObject.PSObject.Properties | Select-Object -ExpandProperty Name)
    $desired = @{}
    $before = @{}

    foreach ($name in $parameterNames) {
        if ($sourcePropertyNames -notcontains $name) { continue }

        $sourceValue = ConvertTo-SettableValue -Value $SourceObject.$name
        $sourceValue = Convert-ServerName -Value $sourceValue
        $targetValue = $null
        if ($null -ne $TargetObject) { $targetValue = ConvertTo-SettableValue -Value $TargetObject.$name }

        if ($name -eq 'PermissionGroups') {
            $sourceValue = ConvertTo-PermissionGroupString -Value $sourceValue
            $targetValue = ConvertTo-PermissionGroupString -Value $targetValue
        }

        if ($BindingProperties -contains $name -and $null -ne $sourceValue) {
            $mapped = Convert-Binding -Binding $sourceValue
            if ($null -eq $mapped) {
                Add-Change -Area $Area -Identity $TargetIdentity -Property $name -SourceValue $sourceValue -TargetValue $targetValue `
                    -Status ReviewRequired -Message 'Binding on a specific IP address without an entry in -IpAddressMap, not copied.'
                continue
            }
            $sourceValue = $mapped
        }

        if ((Get-CompareKey $sourceValue) -eq (Get-CompareKey $targetValue)) { continue }
        $desired[$name] = $sourceValue
        $before[$name] = $targetValue
    }

    # Some parameters only validate in combination (ExternalHostname needs ExternalClientsRequireSsl, LogonFormat
    # needs DefaultDomain). Pull the partner in even when it is already in sync.
    foreach ($key in @($CoupledProperties.Keys)) {
        if (-not $desired.ContainsKey($key)) { continue }
        foreach ($partner in @($CoupledProperties[$key])) {
            if ($desired.ContainsKey($partner)) { continue }
            if ($sourcePropertyNames -notcontains $partner -or $parameterNames -notcontains $partner) { continue }
            $desired[$partner] = Convert-ServerName -Value (ConvertTo-SettableValue -Value $SourceObject.$partner)
            if ($null -ne $TargetObject) { $before[$partner] = ConvertTo-SettableValue -Value $TargetObject.$partner }
            else { $before[$partner] = $null }
        }
    }

    if ($desired.Count -eq 0) {
        Write-Log -Message ('  {0}: in sync' -f $TargetIdentity)
        return
    }

    $propertyList = (@($desired.Keys) | Sort-Object) -join ', '
    if (-not $PSCmdlet.ShouldProcess($TargetIdentity, ('{0} ({1})' -f $SetCmdlet, $propertyList))) {
        foreach ($name in @($desired.Keys | Sort-Object)) {
            Add-Change -Area $Area -Identity $TargetIdentity -Property $name -SourceValue $desired[$name] -TargetValue $before[$name] -Status (Get-PendingStatus)
        }
        return
    }

    $arguments = @{ $IdentityParameterName = $TargetIdentity; ErrorAction = 'Stop'; Confirm = $false }
    foreach ($name in $desired.Keys) { $arguments[$name] = $desired[$name] }

    try {
        & $SetCmdlet @arguments
        foreach ($name in @($desired.Keys | Sort-Object)) {
            Add-Change -Area $Area -Identity $TargetIdentity -Property $name -SourceValue $desired[$name] -TargetValue $before[$name] -Status Applied
        }
        if ($RestartHint -and -not $script:RestartHints.Contains($RestartHint)) { $script:RestartHints.Add($RestartHint) }
        return
    }
    catch {
        Write-Log -Level Warn -Message ('  {0}: bulk call failed ({1}), retrying property by property' -f $TargetIdentity, $_.Exception.Message)
    }

    $anyApplied = $false
    foreach ($name in @($desired.Keys | Sort-Object)) {
        $single = @{ $IdentityParameterName = $TargetIdentity; ErrorAction = 'Stop'; Confirm = $false }
        $single[$name] = $desired[$name]
        if ($CoupledProperties.ContainsKey($name)) {
            foreach ($partner in @($CoupledProperties[$name])) {
                if ($desired.ContainsKey($partner)) { $single[$partner] = $desired[$partner] }
            }
        }
        try {
            & $SetCmdlet @single
            Add-Change -Area $Area -Identity $TargetIdentity -Property $name -SourceValue $desired[$name] -TargetValue $before[$name] -Status Applied
            $anyApplied = $true
        }
        catch {
            Add-Change -Area $Area -Identity $TargetIdentity -Property $name -SourceValue $desired[$name] -TargetValue $before[$name] `
                -Status Failed -Message $_.Exception.Message
        }
    }
    if ($anyApplied -and $RestartHint -and -not $script:RestartHints.Contains($RestartHint)) { $script:RestartHints.Add($RestartHint) }
}

function Get-TargetCounterpart {
    # Finds the object on the target with the same Name as the source object (virtual directories, connectors).
    param($SourceObject, [array] $TargetObjects)
    $wantedName = Convert-ServerName -Value ([string]$SourceObject.Name)
    foreach ($candidate in $TargetObjects) {
        if ([string]$candidate.Name -eq $wantedName) { return $candidate }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Area: Certificates
# ---------------------------------------------------------------------------

function Get-CertificateServiceList {
    # "IMAP, POP, IIS, SMTP" as reported by Get-ExchangeCertificate, reduced to what Enable-ExchangeCertificate
    # can enable. Federation is managed by the federation trust and cannot be enabled by hand.
    param($Services)
    $allowed = @('IIS', 'SMTP', 'IMAP', 'POP')
    return @(("$Services" -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $allowed -contains $_ })
}

function Invoke-AreaCertificates {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $areaName = 'Certificates'
    $sourceCerts = @(Get-ExchangeCertificate -Server $script:SourceShort -ErrorAction Stop)
    $targetCerts = @(Get-ExchangeCertificate -Server $script:TargetShort -ErrorAction Stop)
    Save-Snapshot -Area $areaName -Source ($sourceCerts | Select-Object Thumbprint, Subject, Issuer, Services, NotAfter, IsSelfSigned, PrivateKeyExportable, CertificateDomains) `
                                   -Target ($targetCerts | Select-Object Thumbprint, Subject, Issuer, Services, NotAfter, IsSelfSigned, PrivateKeyExportable, CertificateDomains)

    # The OAuth certificate is self signed but organization wide, every server must hold it.
    $authThumbprints = @()
    try {
        $authConfig = Get-AuthConfig -ErrorAction Stop
        $authThumbprints = @($authConfig.CurrentCertificateThumbprint, $authConfig.NextCertificateThumbprint) | Where-Object { $_ }
    }
    catch { Write-Log -Level Warn -Message ('Get-AuthConfig failed: {0}' -f $_.Exception.Message) }

    $certificateFolder = Join-Path $script:OutputFolder 'Certificates'

    foreach ($cert in $sourceCerts) {
        $thumbprint = [string]$cert.Thumbprint
        $subject = [string]$cert.Subject
        $label = '{0} ({1})' -f $thumbprint, $subject
        $isAuthCert = $authThumbprints -contains $thumbprint

        if ([bool]$cert.IsSelfSigned -and -not $isAuthCert) {
            Add-Change -Area $areaName -Identity $label -Property 'Certificate' -Status Skipped -Message 'Self signed, server specific certificate.'
            continue
        }
        if ($cert.NotAfter -and ([datetime]$cert.NotAfter) -lt (Get-Date)) {
            Add-Change -Area $areaName -Identity $label -Property 'Certificate' -Status Skipped -Message ('Expired on {0}.' -f $cert.NotAfter)
            continue
        }

        $services = Get-CertificateServiceList -Services $cert.Services
        $existing = $targetCerts | Where-Object { [string]$_.Thumbprint -eq $thumbprint } | Select-Object -First 1

        if ($null -eq $existing) {
            if (-not [bool]$cert.PrivateKeyExportable) {
                Add-Change -Area $areaName -Identity $label -Property 'Certificate' -SourceValue $subject -Status ReviewRequired `
                    -Message 'Private key is not exportable. Export it with the original PFX or request the certificate again, then import it on the target.'
                continue
            }
            if ($PSCmdlet.ShouldProcess($script:TargetShort, ('Import certificate {0}' -f $label))) {
                try {
                    $exported = Export-ExchangeCertificate -Server $script:SourceShort -Thumbprint $thumbprint -BinaryEncoded -Password $script:CertificateSecret -ErrorAction Stop
                    $fileData = [byte[]]$exported.FileData
                    if ($CertificatePassword) {
                        if (-not (Test-Path $certificateFolder)) { New-Item -Path $certificateFolder -ItemType Directory | Out-Null }
                        [System.IO.File]::WriteAllBytes((Join-Path $certificateFolder ('{0}.pfx' -f $thumbprint)), $fileData)
                    }
                    Import-ExchangeCertificate -Server $script:TargetShort -FileData $fileData -Password $script:CertificateSecret -PrivateKeyExportable $true -ErrorAction Stop | Out-Null
                    Add-Change -Area $areaName -Identity $label -Property 'Certificate' -SourceValue $subject -Status Applied -Message 'Imported.'
                }
                catch {
                    Add-Change -Area $areaName -Identity $label -Property 'Certificate' -SourceValue $subject -Status Failed -Message $_.Exception.Message
                    continue
                }
            }
            else {
                Add-Change -Area $areaName -Identity $label -Property 'Certificate' -SourceValue $subject -Status (Get-PendingStatus) -Message 'Would import.'
            }
        }
        else {
            Write-Log -Message ('  {0}: already present on target' -f $label)
        }

        # The OAuth certificate must exist on every server but must never become the default SMTP certificate,
        # which is what Enable-ExchangeCertificate -Services SMTP -Force would do. Setup flags it itself.
        if ($isAuthCert) { continue }
        if ($services.Count -eq 0) { continue }
        $targetServices = @()
        if ($null -ne $existing) { $targetServices = Get-CertificateServiceList -Services $existing.Services }
        $missingServices = @($services | Where-Object { $targetServices -notcontains $_ })
        if ($missingServices.Count -eq 0) { continue }

        $serviceString = $services -join ','
        if ($PSCmdlet.ShouldProcess($script:TargetShort, ('Enable-ExchangeCertificate {0} -Services {1}' -f $thumbprint, $serviceString))) {
            try {
                # Force suppresses the "overwrite the default SMTP certificate" prompt, which is exactly the intent.
                Enable-ExchangeCertificate -Server $script:TargetShort -Thumbprint $thumbprint -Services $serviceString -Force -Confirm:$false -ErrorAction Stop
                Add-Change -Area $areaName -Identity $label -Property 'Services' -SourceValue $serviceString -TargetValue ($targetServices -join ',') -Status Applied
                if (-not $script:RestartHints.Contains('iisreset')) { $script:RestartHints.Add('iisreset') }
            }
            catch {
                Add-Change -Area $areaName -Identity $label -Property 'Services' -SourceValue $serviceString -TargetValue ($targetServices -join ',') -Status Failed -Message $_.Exception.Message
            }
        }
        else {
            Add-Change -Area $areaName -Identity $label -Property 'Services' -SourceValue $serviceString -TargetValue ($targetServices -join ',') -Status (Get-PendingStatus)
        }
    }
}

# ---------------------------------------------------------------------------
# Area: ExchangeServer, ClientAccess, MailboxServer, Malware
# ---------------------------------------------------------------------------

function Invoke-AreaExchangeServer {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $source = Get-ExchangeServer -Identity $script:SourceShort -ErrorAction Stop
    $target = Get-ExchangeServer -Identity $script:TargetShort -ErrorAction Stop
    Save-Snapshot -Area 'ExchangeServer' -Source $source -Target $target

    if ([bool]$target.IsExchangeTrialEdition -and -not [bool]$source.IsExchangeTrialEdition) {
        Add-Change -Area 'ExchangeServer' -Identity $script:TargetShort -Property 'Edition' -SourceValue $source.Edition -TargetValue 'Trial' `
            -Status ReviewRequired -Message 'Target runs unlicensed. License it with Set-ExchangeServer -ProductKey, the key is never copied by this script.'
    }

    # Product key is a secret and the static domain controller pins belong to the site the server lives in.
    Sync-ObjectSettings -Area 'ExchangeServer' -SourceObject $source -TargetObject $target -SetCmdlet 'Set-ExchangeServer' `
        -TargetIdentity $script:TargetShort -ExcludeProperties @('ProductKey', 'Static*')
}

function Invoke-AreaClientAccess {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $source = Get-ClientAccessService -Identity $script:SourceShort -IncludeAlternateServiceAccountCredentialStatus -ErrorAction Stop
    $target = Get-ClientAccessService -Identity $script:TargetShort -IncludeAlternateServiceAccountCredentialStatus -ErrorAction Stop
    Save-Snapshot -Area 'ClientAccess' -Source $source -Target $target

    Sync-ObjectSettings -Area 'ClientAccess' -SourceObject $source -TargetObject $target -SetCmdlet 'Set-ClientAccessService' `
        -TargetIdentity $script:TargetShort `
        -ExcludeProperties @('AlternateServiceAccountCredential', 'RemoveAlternateServiceAccountCredentials', 'CleanUpInvalidAlternateServiceAccountCredentials', 'Array')

    # The alternate service account (Kerberos for MAPI and Outlook Anywhere) holds a password, so it is reported
    # rather than copied.
    $asa = $source.AlternateServiceAccountConfiguration
    $hasAsa = $false
    if ($null -ne $asa -and $null -ne $asa.EffectiveCredentials) { $hasAsa = (@($asa.EffectiveCredentials).Count -gt 0) }
    if ($hasAsa) {
        $targetHasAsa = $false
        $targetAsa = $target.AlternateServiceAccountConfiguration
        if ($null -ne $targetAsa -and $null -ne $targetAsa.EffectiveCredentials) { $targetHasAsa = (@($targetAsa.EffectiveCredentials).Count -gt 0) }
        if (-not $targetHasAsa) {
            Add-Change -Area 'ClientAccess' -Identity $script:TargetShort -Property 'AlternateServiceAccountCredential' -SourceValue 'configured' -TargetValue 'not configured' `
                -Status ReviewRequired -Message ('Run .\RollAlternateServiceAccountPassword.ps1 -ToSpecificServer {0} -CopyFrom {1} from the Exchange Scripts folder.' -f $script:TargetShort, $script:SourceShort)
        }
    }
}

function Invoke-AreaMailboxServer {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $source = Get-MailboxServer -Identity $script:SourceShort -ErrorAction Stop
    $target = Get-MailboxServer -Identity $script:TargetShort -ErrorAction Stop
    Save-Snapshot -Area 'MailboxServer' -Source $source -Target $target

    # DatabaseCopyActivationDisabledAndMoveNow triggers moves, that is an operational state and not configuration.
    Sync-ObjectSettings -Area 'MailboxServer' -SourceObject $source -TargetObject $target -SetCmdlet 'Set-MailboxServer' `
        -TargetIdentity $script:TargetShort `
        -ExcludeProperties ($script:PathExclusions + @('DatabaseCopyActivationDisabledAndMoveNow', 'WorkloadManagementPolicy'))
}

function Invoke-AreaMalware {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $source = Get-MalwareFilteringServer -Identity $script:SourceShort -ErrorAction Stop
    $target = Get-MalwareFilteringServer -Identity $script:TargetShort -ErrorAction Stop
    Save-Snapshot -Area 'Malware' -Source $source -Target $target
    Sync-ObjectSettings -Area 'Malware' -SourceObject $source -TargetObject $target -SetCmdlet 'Set-MalwareFilteringServer' `
        -TargetIdentity $script:TargetShort -ExcludeProperties @('ForceRescan')
}

# ---------------------------------------------------------------------------
# Area: VirtualDirectories, OutlookAnywhere
# ---------------------------------------------------------------------------

function Invoke-AreaVirtualDirectories {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $areaName = 'VirtualDirectories'
    $directoryTypes = @(
        @{ Get = 'Get-OwaVirtualDirectory';          Set = 'Set-OwaVirtualDirectory';          Coupled = @{ LogonFormat = @('DefaultDomain') }; Exclude = @() },
        @{ Get = 'Get-EcpVirtualDirectory';          Set = 'Set-EcpVirtualDirectory';          Coupled = @{}; Exclude = @() },
        @{ Get = 'Get-WebServicesVirtualDirectory';  Set = 'Set-WebServicesVirtualDirectory';  Coupled = @{}; Exclude = @() },
        # ActiveSyncServer is an alias attribute of ExternalUrl, passing both makes Exchange complain.
        @{ Get = 'Get-ActiveSyncVirtualDirectory';   Set = 'Set-ActiveSyncVirtualDirectory';   Coupled = @{}; Exclude = @('ActiveSyncServer') },
        @{ Get = 'Get-OabVirtualDirectory';          Set = 'Set-OabVirtualDirectory';          Coupled = @{}; Exclude = @() },
        @{ Get = 'Get-MapiVirtualDirectory';         Set = 'Set-MapiVirtualDirectory';         Coupled = @{}; Exclude = @() },
        @{ Get = 'Get-PowerShellVirtualDirectory';   Set = 'Set-PowerShellVirtualDirectory';   Coupled = @{}; Exclude = @() },
        @{ Get = 'Get-AutodiscoverVirtualDirectory'; Set = 'Set-AutodiscoverVirtualDirectory'; Coupled = @{}; Exclude = @() }
    )

    foreach ($type in $directoryTypes) {
        Write-Log -Message ('-- {0}' -f $type.Get)
        try {
            # Without -ADPropertiesOnly the IIS metabase is read too, which is where the authentication settings live.
            $sourceDirectories = @(& $type.Get -Server $script:SourceShort -ErrorAction Stop)
            $targetDirectories = @(& $type.Get -Server $script:TargetShort -ErrorAction Stop)
        }
        catch {
            Add-Change -Area $areaName -Identity $type.Get -Property '(read)' -Status Failed -Message $_.Exception.Message
            continue
        }
        Save-Snapshot -Area $areaName -Source $sourceDirectories -Target $targetDirectories

        foreach ($sourceDirectory in $sourceDirectories) {
            $targetDirectory = Get-TargetCounterpart -SourceObject $sourceDirectory -TargetObjects $targetDirectories
            if ($null -eq $targetDirectory) {
                Add-Change -Area $areaName -Identity ('{0}\{1}' -f $script:TargetShort, $sourceDirectory.Name) -Property '(exists)' -SourceValue 'present' -TargetValue 'missing' `
                    -Status ReviewRequired -Message 'Virtual directory does not exist on the target.'
                continue
            }
            Sync-ObjectSettings -Area $areaName -SourceObject $sourceDirectory -TargetObject $targetDirectory -SetCmdlet $type.Set `
                -TargetIdentity ([string]$targetDirectory.Identity) -ExcludeProperties $type.Exclude -CoupledProperties $type.Coupled `
                -RestartHint 'iisreset'
        }
    }
}

function Invoke-AreaOutlookAnywhere {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $sourceItems = @(Get-OutlookAnywhere -Server $script:SourceShort -ErrorAction Stop)
    $targetItems = @(Get-OutlookAnywhere -Server $script:TargetShort -ErrorAction Stop)
    Save-Snapshot -Area 'OutlookAnywhere' -Source $sourceItems -Target $targetItems

    foreach ($sourceItem in $sourceItems) {
        $targetItem = Get-TargetCounterpart -SourceObject $sourceItem -TargetObjects $targetItems
        if ($null -eq $targetItem) {
            Add-Change -Area 'OutlookAnywhere' -Identity ('{0}\{1}' -f $script:TargetShort, $sourceItem.Name) -Property '(exists)' -SourceValue 'present' -TargetValue 'missing' `
                -Status ReviewRequired -Message 'Outlook Anywhere virtual directory does not exist on the target.'
            continue
        }
        # ClientAuthenticationMethod and DefaultAuthenticationMethod are legacy parameters that conflict with the
        # explicit internal and external ones. A hostname must be set together with its RequireSsl and auth method.
        Sync-ObjectSettings -Area 'OutlookAnywhere' -SourceObject $sourceItem -TargetObject $targetItem -SetCmdlet 'Set-OutlookAnywhere' `
            -TargetIdentity ([string]$targetItem.Identity) `
            -ExcludeProperties @('ClientAuthenticationMethod', 'DefaultAuthenticationMethod') `
            -CoupledProperties @{
                ExternalHostname = @('ExternalClientsRequireSsl', 'ExternalClientAuthenticationMethod')
                InternalHostname = @('InternalClientsRequireSsl', 'InternalClientAuthenticationMethod')
            } -RestartHint 'iisreset'
    }
}

# ---------------------------------------------------------------------------
# Area: Transport, PopImap, EventLogLevels, SettingOverrides
# ---------------------------------------------------------------------------

function Invoke-AreaTransport {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $services = @(
        @{ Get = 'Get-TransportService';         Set = 'Set-TransportService';         Hint = 'Restart-Service MSExchangeTransport' },
        @{ Get = 'Get-FrontendTransportService'; Set = 'Set-FrontendTransportService'; Hint = 'Restart-Service MSExchangeFrontEndTransport' },
        @{ Get = 'Get-MailboxTransportService';  Set = 'Set-MailboxTransportService';  Hint = 'Restart-Service MSExchangeDelivery, MSExchangeSubmission' }
    )
    foreach ($service in $services) {
        Write-Log -Message ('-- {0}' -f $service.Get)
        $source = & $service.Get -Identity $script:SourceShort -ErrorAction Stop
        $target = & $service.Get -Identity $script:TargetShort -ErrorAction Stop
        Save-Snapshot -Area 'Transport' -Source $source -Target $target
        # The DNS adapter GUIDs identify a network adapter of the source machine and never match on the target.
        Sync-ObjectSettings -Area 'Transport' -SourceObject $source -TargetObject $target -SetCmdlet $service.Set `
            -TargetIdentity $script:TargetShort `
            -ExcludeProperties ($script:PathExclusions + @('InternalDNSAdapterGuid', 'ExternalDNSAdapterGuid')) `
            -RestartHint $service.Hint
    }
}

function Invoke-AreaPopImap {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $areaName = 'PopImap'
    $protocols = @(
        @{ Get = 'Get-ImapSettings'; Set = 'Set-ImapSettings'; Services = @('MSExchangeIMAP4', 'MSExchangeIMAP4BE'); Hint = 'Restart-Service MSExchangeIMAP4, MSExchangeIMAP4BE' },
        @{ Get = 'Get-PopSettings';  Set = 'Set-PopSettings';  Services = @('MSExchangePOP3', 'MSExchangePOP3BE');   Hint = 'Restart-Service MSExchangePOP3, MSExchangePOP3BE' }
    )
    foreach ($protocol in $protocols) {
        Write-Log -Message ('-- {0}' -f $protocol.Get)
        $source = & $protocol.Get -Server $script:SourceShort -ErrorAction Stop
        $target = & $protocol.Get -Server $script:TargetShort -ErrorAction Stop
        Save-Snapshot -Area $areaName -Source $source -Target $target
        Sync-ObjectSettings -Area $areaName -SourceObject $source -TargetObject $target -SetCmdlet $protocol.Set `
            -TargetIdentity $script:TargetShort -IdentityParameterName 'Server' -ExcludeProperties $script:PathExclusions `
            -BindingProperties @('SSLBindings', 'UnencryptedOrTLSBindings') -RestartHint $protocol.Hint

        # IMAP and POP are disabled by default. When the source runs them, the target must start them too.
        foreach ($serviceName in $protocol.Services) {
            try {
                $sourceService = Get-CimInstance -ClassName Win32_Service -ComputerName $script:SourceFqdn -Filter ("Name='{0}'" -f $serviceName) -ErrorAction Stop
                $targetService = Get-CimInstance -ClassName Win32_Service -ComputerName $script:TargetFqdn -Filter ("Name='{0}'" -f $serviceName) -ErrorAction Stop
            }
            catch {
                Add-Change -Area $areaName -Identity ('{0}\{1}' -f $script:TargetShort, $serviceName) -Property 'StartMode' -Status ReviewRequired `
                    -Message ('Could not read the service state remotely: {0}' -f $_.Exception.Message)
                continue
            }
            if ($null -eq $sourceService -or $null -eq $targetService) { continue }
            if ($sourceService.StartMode -eq $targetService.StartMode) { continue }

            $startupType = switch ($sourceService.StartMode) { 'Auto' { 'Automatic' } 'Manual' { 'Manual' } default { 'Disabled' } }
            $identity = '{0}\{1}' -f $script:TargetShort, $serviceName
            if ($PSCmdlet.ShouldProcess($identity, ('Set-Service -StartupType {0}' -f $startupType))) {
                try {
                    Set-Service -ComputerName $script:TargetFqdn -Name $serviceName -StartupType $startupType -ErrorAction Stop
                    Add-Change -Area $areaName -Identity $identity -Property 'StartMode' -SourceValue $sourceService.StartMode -TargetValue $targetService.StartMode -Status Applied
                    if (-not $script:RestartHints.Contains($protocol.Hint)) { $script:RestartHints.Add($protocol.Hint) }
                }
                catch {
                    Add-Change -Area $areaName -Identity $identity -Property 'StartMode' -SourceValue $sourceService.StartMode -TargetValue $targetService.StartMode -Status Failed -Message $_.Exception.Message
                }
            }
            else {
                Add-Change -Area $areaName -Identity $identity -Property 'StartMode' -SourceValue $sourceService.StartMode -TargetValue $targetService.StartMode -Status (Get-PendingStatus)
            }
        }
    }
}

function Get-EventLogCategory {
    # Identity looks like "EX01\MSExchangeTransport\SmtpReceive" or "MSExchangeTransport\SmtpReceive" depending on
    # how it was queried. Return the part without the server prefix.
    param([string] $Identity)
    $parts = $Identity -split '\\'
    if ($parts.Count -ge 3) { return (($parts | Select-Object -Skip 1) -join '\') }
    return $Identity
}

function Invoke-AreaEventLogLevels {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $areaName = 'EventLogLevels'
    $sourceLevels = @(Get-EventLogLevel -Server $script:SourceShort -ErrorAction Stop | Where-Object { "$($_.EventLevel)" -ne 'Lowest' })
    $targetLevels = @(Get-EventLogLevel -Server $script:TargetShort -ErrorAction Stop)
    Save-Snapshot -Area $areaName -Source $sourceLevels -Target ($targetLevels | Where-Object { "$($_.EventLevel)" -ne 'Lowest' })

    if ($sourceLevels.Count -eq 0) { Write-Log -Message '  No raised diagnostic levels on the source.'; return }

    $targetByCategory = @{}
    foreach ($level in $targetLevels) { $targetByCategory[(Get-EventLogCategory -Identity ([string]$level.Identity))] = [string]$level.EventLevel }

    foreach ($level in $sourceLevels) {
        $category = Get-EventLogCategory -Identity ([string]$level.Identity)
        $wanted = [string]$level.EventLevel
        $current = $targetByCategory[$category]
        if ($current -eq $wanted) { continue }
        $identity = '{0}\{1}' -f $script:TargetShort, $category
        if ($PSCmdlet.ShouldProcess($identity, ('Set-EventLogLevel -Level {0}' -f $wanted))) {
            try {
                Set-EventLogLevel -Identity $identity -Level $wanted -ErrorAction Stop
                Add-Change -Area $areaName -Identity $identity -Property 'EventLevel' -SourceValue $wanted -TargetValue $current -Status Applied
            }
            catch {
                Add-Change -Area $areaName -Identity $identity -Property 'EventLevel' -SourceValue $wanted -TargetValue $current -Status Failed -Message $_.Exception.Message
            }
        }
        else {
            Add-Change -Area $areaName -Identity $identity -Property 'EventLevel' -SourceValue $wanted -TargetValue $current -Status (Get-PendingStatus)
        }
    }
}

function Invoke-AreaSettingOverrides {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $areaName = 'SettingOverrides'
    $overrides = @(Get-SettingOverride -ErrorAction Stop)
    $relevant = @()
    foreach ($override in $overrides) {
        $servers = Get-ValueList -Value (ConvertTo-SettableValue -Value $override.Server)
        if ($servers.Count -eq 0) { continue }                       # organization wide, applies to the target already
        if ($servers -notcontains $script:SourceShort) { continue }
        $relevant += $override
        if ($servers -contains $script:TargetShort) { Write-Log -Message ('  {0}: in sync' -f $override.Name); continue }

        $newServers = [string[]]($servers + $script:TargetShort)
        if ($PSCmdlet.ShouldProcess([string]$override.Identity, ('Set-SettingOverride -Server {0}' -f ($newServers -join ', ')))) {
            try {
                Set-SettingOverride -Identity ([string]$override.Identity) -Server $newServers -Confirm:$false -ErrorAction Stop
                Add-Change -Area $areaName -Identity ([string]$override.Name) -Property 'Server' -SourceValue $newServers -TargetValue $servers -Status Applied
            }
            catch {
                Add-Change -Area $areaName -Identity ([string]$override.Name) -Property 'Server' -SourceValue $newServers -TargetValue $servers -Status Failed -Message $_.Exception.Message
            }
        }
        else {
            Add-Change -Area $areaName -Identity ([string]$override.Name) -Property 'Server' -SourceValue $newServers -TargetValue $servers -Status (Get-PendingStatus)
        }
    }
    Save-Snapshot -Area $areaName -Source $relevant
    if ($relevant.Count -eq 0) { Write-Log -Message '  No setting overrides scoped to the source server.' }
}

# ---------------------------------------------------------------------------
# Area: ReceiveConnectors, SendConnectors
# ---------------------------------------------------------------------------

function Get-AcePermissionKey {
    # Key that identifies an explicit ACE regardless of the object it sits on, used to find missing ACEs.
    param($Ace)
    $user = Convert-ServerName -Value ([string]$Ace.User)
    $extended = Get-CompareKey (ConvertTo-SettableValue -Value $Ace.ExtendedRights)
    $access = Get-CompareKey (ConvertTo-SettableValue -Value $Ace.AccessRights)
    return ('{0}|{1}|{2}|{3}' -f $user.ToLowerInvariant(), $extended, $access, [bool]$Ace.Deny)
}

function Sync-ReceiveConnectorPermission {
    <#
        Copies the explicit (not inherited) AD permissions of a receive connector. This is where anonymous relay
        lives: ms-Exch-SMTP-Accept-Any-Recipient granted to NT AUTHORITY\ANONYMOUS LOGON is not part of any
        permission group and is lost when a connector is rebuilt by hand.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string] $SourceIdentity, [string] $TargetIdentity, [string] $TargetLabel, [bool] $TargetExists)
    $areaName = 'ReceiveConnectors'
    if (-not $TargetLabel) { $TargetLabel = $TargetIdentity }

    $sourceAces = @(Get-ADPermission -Identity $SourceIdentity -ErrorAction Stop | Where-Object { -not [bool]$_.IsInherited })
    if ($sourceAces.Count -eq 0) { return }

    $targetKeys = @{}
    if ($TargetExists) {
        foreach ($ace in @(Get-ADPermission -Identity $TargetIdentity -ErrorAction Stop | Where-Object { -not [bool]$_.IsInherited })) {
            $targetKeys[(Get-AcePermissionKey -Ace $ace)] = $true
        }
    }

    foreach ($ace in $sourceAces) {
        $key = Get-AcePermissionKey -Ace $ace
        if ($targetKeys.ContainsKey($key)) { continue }

        $user = Convert-ServerName -Value ([string]$ace.User)
        $extendedRights = ConvertTo-SettableValue -Value $ace.ExtendedRights
        $accessRights = @()
        foreach ($right in (Get-ValueList -Value (ConvertTo-SettableValue -Value $ace.AccessRights))) {
            foreach ($part in ($right -split ',')) { if ($part.Trim()) { $accessRights += $part.Trim() } }
        }
        # ExtendedRight as access right is implied by -ExtendedRights, passing both is rejected.
        $accessRights = @($accessRights | Where-Object { $_ -ne 'ExtendedRight' -or $null -eq $extendedRights })

        $arguments = @{ Identity = $TargetIdentity; User = $user; Confirm = $false; ErrorAction = 'Stop' }
        if ($null -ne $extendedRights) { $arguments['ExtendedRights'] = [string[]]$extendedRights }
        if ($accessRights.Count -gt 0) { $arguments['AccessRights'] = [string[]]$accessRights }
        if ([bool]$ace.Deny) { $arguments['Deny'] = $true }
        $properties = ConvertTo-SettableValue -Value $ace.Properties
        if ($null -ne $properties) { $arguments['Properties'] = [string[]]$properties }
        $childTypes = ConvertTo-SettableValue -Value $ace.ChildObjectTypes
        if ($null -ne $childTypes) { $arguments['ChildObjectTypes'] = [string[]]$childTypes }
        $inheritedType = ConvertTo-SettableValue -Value $ace.InheritedObjectType
        if ($null -ne $inheritedType) { $arguments['InheritedObjectType'] = [string]$inheritedType }
        $inheritance = ConvertTo-SettableValue -Value $ace.InheritanceType
        if ($null -ne $inheritance -and "$inheritance" -ne 'None') { $arguments['InheritanceType'] = [string]$inheritance }

        $description = '{0}: {1}{2}' -f $user, (($extendedRights + $accessRights | Where-Object { $_ }) -join ','), $(if ([bool]$ace.Deny) { ' (Deny)' } else { '' })
        if ($PSCmdlet.ShouldProcess($TargetLabel, ('Add-ADPermission {0}' -f $description))) {
            try {
                Add-ADPermission @arguments | Out-Null
                Add-Change -Area $areaName -Identity $TargetLabel -Property 'ADPermission' -SourceValue $description -TargetValue 'missing' -Status Applied
            }
            catch {
                Add-Change -Area $areaName -Identity $TargetLabel -Property 'ADPermission' -SourceValue $description -TargetValue 'missing' -Status Failed -Message $_.Exception.Message
            }
        }
        else {
            Add-Change -Area $areaName -Identity $TargetLabel -Property 'ADPermission' -SourceValue $description -TargetValue 'missing' -Status (Get-PendingStatus)
        }
    }
}

function Invoke-AreaReceiveConnectors {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $areaName = 'ReceiveConnectors'
    $sourceConnectors = @(Get-ReceiveConnector -Server $script:SourceShort -ErrorAction Stop)
    $targetConnectors = @(Get-ReceiveConnector -Server $script:TargetShort -ErrorAction Stop)
    Save-Snapshot -Area $areaName -Source $sourceConnectors -Target $targetConnectors

    foreach ($connector in $sourceConnectors) {
        $targetName = Convert-ServerName -Value ([string]$connector.Name)
        $targetIdentity = '{0}\{1}' -f $script:TargetShort, $targetName
        $targetConnector = Get-TargetCounterpart -SourceObject $connector -TargetObjects $targetConnectors
        Write-Log -Message ('-- {0}' -f $targetIdentity)

        if ($null -eq $targetConnector) {
            $bindings = Convert-Binding -Binding (ConvertTo-SettableValue -Value $connector.Bindings)
            if ($null -eq $bindings) {
                Add-Change -Area $areaName -Identity $targetIdentity -Property '(create)' -SourceValue (ConvertTo-SettableValue -Value $connector.Bindings) -TargetValue 'missing' `
                    -Status ReviewRequired -Message 'Connector is bound to a specific IP address, supply -IpAddressMap to create it on the target.'
                continue
            }
            $remoteRanges = ConvertTo-SettableValue -Value $connector.RemoteIPRanges
            $role = [string]$connector.TransportRole
            if ($PSCmdlet.ShouldProcess($script:TargetShort, ('New-ReceiveConnector {0} ({1}, {2})' -f $targetName, $role, ($bindings -join ', ')))) {
                try {
                    # Created as a custom connector with only the mandatory values, Sync-ObjectSettings converges the rest.
                    New-ReceiveConnector -Name $targetName -Server $script:TargetShort -TransportRole $role -Custom -Bindings $bindings `
                        -RemoteIPRanges $remoteRanges -Confirm:$false -ErrorAction Stop | Out-Null
                    $targetConnector = Get-ReceiveConnector -Identity $targetIdentity -ErrorAction Stop
                    Add-Change -Area $areaName -Identity $targetIdentity -Property '(create)' -SourceValue ($bindings -join ', ') -TargetValue 'missing' -Status Applied
                }
                catch {
                    Add-Change -Area $areaName -Identity $targetIdentity -Property '(create)' -SourceValue ($bindings -join ', ') -TargetValue 'missing' -Status Failed -Message $_.Exception.Message
                    continue
                }
            }
            else {
                Add-Change -Area $areaName -Identity $targetIdentity -Property '(create)' -SourceValue ($bindings -join ', ') -TargetValue 'missing' -Status (Get-PendingStatus)
            }
        }

        Sync-ObjectSettings -Area $areaName -SourceObject $connector -TargetObject $targetConnector -SetCmdlet 'Set-ReceiveConnector' `
            -TargetIdentity $targetIdentity -ExcludeProperties @('TransportRole', 'Usage', 'Custom', 'Internal', 'Internet', 'Client', 'Partner') `
            -BindingProperties @('Bindings') -RestartHint 'Restart-Service MSExchangeTransport, MSExchangeFrontEndTransport'

        # Distinguished names on purpose: Get-ADPermission and Add-ADPermission do not resolve the Server\Name form
        # reliably. A connector that only exists in the WhatIf log keeps the Server\Name text for the report.
        $targetPermissionIdentity = if ($null -ne $targetConnector) { [string]$targetConnector.DistinguishedName } else { $targetIdentity }
        Sync-ReceiveConnectorPermission -SourceIdentity ([string]$connector.DistinguishedName) -TargetIdentity $targetPermissionIdentity -TargetLabel $targetIdentity -TargetExists ($null -ne $targetConnector)
    }

    # Connectors that only exist on the target are not touched, the operator decides.
    foreach ($targetConnector in $targetConnectors) {
        $match = $false
        foreach ($connector in $sourceConnectors) {
            if ((Convert-ServerName -Value ([string]$connector.Name)) -eq [string]$targetConnector.Name) { $match = $true; break }
        }
        if (-not $match) {
            Add-Change -Area $areaName -Identity ([string]$targetConnector.Identity) -Property '(exists)' -SourceValue 'missing' -TargetValue 'present' `
                -Status Info -Message 'Connector exists on the target only, left untouched.'
        }
    }
}

function Invoke-AreaSendConnectors {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $areaName = 'SendConnectors'
    $connectors = @(Get-SendConnector -ErrorAction Stop)
    Save-Snapshot -Area $areaName -Source $connectors

    foreach ($connector in $connectors) {
        $sourceServers = Get-ValueList -Value (ConvertTo-SettableValue -Value $connector.SourceTransportServers)
        $name = [string]$connector.Name
        if ($sourceServers -notcontains $script:SourceShort) {
            Add-Change -Area $areaName -Identity $name -Property 'SourceTransportServers' -SourceValue $sourceServers -Status Info -Message 'Source server is not a source transport server, nothing to do.'
            continue
        }
        if ($sourceServers -contains $script:TargetShort) { Write-Log -Message ('  {0}: in sync' -f $name); continue }

        $newServers = [string[]]($sourceServers + $script:TargetShort)
        if ($PSCmdlet.ShouldProcess($name, ('Set-SendConnector -SourceTransportServers {0}' -f ($newServers -join ', ')))) {
            try {
                # From this moment on the target takes part in outbound routing for this connector.
                Set-SendConnector -Identity ([string]$connector.Identity) -SourceTransportServers $newServers -Confirm:$false -ErrorAction Stop
                Add-Change -Area $areaName -Identity $name -Property 'SourceTransportServers' -SourceValue $newServers -TargetValue $sourceServers -Status Applied
            }
            catch {
                Add-Change -Area $areaName -Identity $name -Property 'SourceTransportServers' -SourceValue $newServers -TargetValue $sourceServers -Status Failed -Message $_.Exception.Message
            }
        }
        else {
            Add-Change -Area $areaName -Identity $name -Property 'SourceTransportServers' -SourceValue $newServers -TargetValue $sourceServers -Status (Get-PendingStatus)
        }
    }
}

# ---------------------------------------------------------------------------
# Area: TransportAgents, Hybrid (report only)
# ---------------------------------------------------------------------------

function Get-TransportAgentFromConfig {
    # Get-TransportAgent has no Server parameter and only reads the server the shell is connected to, so the agent
    # list is read from agents.config (Hub) or fetagents.config (FrontEnd) over the admin share instead. The file
    # lists the agents in priority order with the same values Get-TransportAgent shows.
    param($ExchangeServer, [string] $TransportService)
    $fileName = if ($TransportService -eq 'FrontEnd') { 'fetagents.config' } else { 'agents.config' }
    $path = Join-Path (Get-ExchangeInstallShare -ExchangeServer $ExchangeServer) ('TransportRoles\Shared\{0}' -f $fileName)
    [xml]$xml = Get-Content -Path $path -ErrorAction Stop
    $priority = 0
    foreach ($agent in @($xml.configuration.mexRuntime.agentList.agent)) {
        if ($null -eq $agent) { continue }
        $priority++
        [pscustomobject]@{
            Identity              = [string]$agent.name
            Enabled               = ([string]$agent.enabled -eq 'true')
            Priority              = $priority
            TransportAgentFactory = [string]$agent.classFactory
            AssemblyPath          = [string]$agent.assemblyPath
        }
    }
}

function Invoke-AreaTransportAgents {
    [CmdletBinding()]
    param()
    $areaName = 'TransportAgents'
    foreach ($transportService in @('Hub', 'FrontEnd')) {
        try {
            $sourceAgents = @(Get-TransportAgentFromConfig -ExchangeServer $script:SourceExchangeServer -TransportService $transportService)
            $targetAgents = @(Get-TransportAgentFromConfig -ExchangeServer $script:TargetExchangeServer -TransportService $transportService)
        }
        catch {
            Add-Change -Area $areaName -Identity $transportService -Property '(read)' -Status Failed -Message $_.Exception.Message
            continue
        }
        Save-Snapshot -Area $areaName -Source $sourceAgents -Target $targetAgents

        foreach ($agent in $sourceAgents) {
            $agentName = [string]$agent.Identity
            $counterpart = $targetAgents | Where-Object { [string]$_.Identity -eq $agentName } | Select-Object -First 1
            $label = '{0}\{1}\{2}' -f $script:TargetShort, $transportService, $agentName
            if ($null -eq $counterpart) {
                $detail = $agent
                Add-Change -Area $areaName -Identity $label -Property '(exists)' -SourceValue 'present' -TargetValue 'missing' -Status ReviewRequired `
                    -Message ('Copy the assembly and run: Install-TransportAgent -Name "{0}" -TransportService {1} -TransportAgentFactory "{2}" -AssemblyPath "{3}"; Enable-TransportAgent; Set-TransportAgent -Priority {4}' -f $agentName, $transportService, $detail.TransportAgentFactory, $detail.AssemblyPath, $agent.Priority)
                continue
            }
            if ([bool]$agent.Enabled -ne [bool]$counterpart.Enabled) {
                Add-Change -Area $areaName -Identity $label -Property 'Enabled' -SourceValue $agent.Enabled -TargetValue $counterpart.Enabled -Status ReviewRequired `
                    -Message ('Use Enable-TransportAgent or Disable-TransportAgent -TransportService {0} on the target.' -f $transportService)
            }
            if ([string]$agent.Priority -ne [string]$counterpart.Priority) {
                Add-Change -Area $areaName -Identity $label -Property 'Priority' -SourceValue $agent.Priority -TargetValue $counterpart.Priority -Status ReviewRequired `
                    -Message ('Set-TransportAgent -Identity "{0}" -TransportService {1} -Priority {2}' -f $agentName, $transportService, $agent.Priority)
            }
        }
    }
}

function Invoke-AreaHybrid {
    [CmdletBinding()]
    param()
    $areaName = 'Hybrid'
    $hybrid = $null
    try { $hybrid = Get-HybridConfiguration -ErrorAction Stop } catch { }
    if ($null -eq $hybrid) { Write-Log -Message '  No hybrid configuration in this organization.'; return }
    Save-Snapshot -Area $areaName -Source $hybrid

    foreach ($property in @('SendingTransportServers', 'ReceivingTransportServers', 'EdgeTransportServers')) {
        $servers = Get-ValueList -Value (ConvertTo-SettableValue -Value $hybrid.$property)
        if ($servers -notcontains $script:SourceShort) { continue }
        if ($servers -contains $script:TargetShort) { continue }
        Add-Change -Area $areaName -Identity 'HybridConfiguration' -Property $property -SourceValue $servers -TargetValue $servers -Status ReviewRequired `
            -Message ('Rerun the Hybrid Configuration Wizard and select {0} as well, or run Set-HybridConfiguration -{1} {2}.' -f $script:TargetShort, $property, (($servers + $script:TargetShort) -join ','))
    }
}

# ---------------------------------------------------------------------------
# Area: ConfigFiles (report only)
# ---------------------------------------------------------------------------

function Get-ExchangeInstallShare {
    # UNC path to the Exchange installation folder via the admin share. DataPath of Get-ExchangeServer points to the
    # Mailbox folder under the installation root, so its parent is the root.
    param($ExchangeServer)
    $dataPath = [string]$ExchangeServer.DataPath
    $installPath = 'C:\Program Files\Microsoft\Exchange Server\V15'
    if ($dataPath) { $installPath = Split-Path -Path $dataPath -Parent }
    return ('\\{0}\{1}' -f $ExchangeServer.Fqdn, ($installPath -replace '^([A-Za-z]):', '$1$$'))
}

function Save-ConfigFileDiff {
    # Saves both versions and a line diff under ConfigFiles\<relative path> in the output folder.
    param([string] $RelativePath, [string] $SourcePath, [string] $TargetPath)
    $folder = Join-Path (Join-Path $script:OutputFolder 'ConfigFiles') (Split-Path -Path $RelativePath -Parent)
    if (-not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    $fileName = Split-Path -Path $RelativePath -Leaf
    Copy-Item -Path $SourcePath -Destination (Join-Path $folder ('{0}.{1}' -f $fileName, $script:SourceShort)) -Force
    $diffPath = Join-Path $folder ('{0}.diff.txt' -f $fileName)
    if (Test-Path $TargetPath) {
        Copy-Item -Path $TargetPath -Destination (Join-Path $folder ('{0}.{1}' -f $fileName, $script:TargetShort)) -Force
        $sourceLines = @(Get-Content -Path $SourcePath)
        $targetLines = @(Get-Content -Path $TargetPath)
        $diff = @(Compare-Object -ReferenceObject $sourceLines -DifferenceObject $targetLines)
        $lines = foreach ($entry in $diff) {
            if ($entry.SideIndicator -eq '<=') { 'SOURCE ONLY : {0}' -f $entry.InputObject } else { 'TARGET ONLY : {0}' -f $entry.InputObject }
        }
        @(('Diff of {0}' -f $RelativePath), ('SOURCE = {0}, TARGET = {1}' -f $script:SourceShort, $script:TargetShort), '') + @($lines) | Set-Content -Path $diffPath
    }
    else {
        @(('{0} exists on {1} only.' -f $RelativePath, $script:SourceShort)) | Set-Content -Path $diffPath
    }
    return $diffPath
}

function Invoke-AreaConfigFiles {
    [CmdletBinding()]
    param()
    $areaName = 'ConfigFiles'
    $sourceRoot = Get-ExchangeInstallShare -ExchangeServer $script:SourceExchangeServer
    $targetRoot = Get-ExchangeInstallShare -ExchangeServer $script:TargetExchangeServer
    Write-Log -Message ('  Source: {0}' -f $sourceRoot)
    Write-Log -Message ('  Target: {0}' -f $targetRoot)

    if (-not (Test-Path $sourceRoot)) {
        Add-Change -Area $areaName -Identity $sourceRoot -Property '(access)' -Status Failed -Message 'Admin share of the source is not reachable, config files not compared.'
        return
    }
    if (-not (Test-Path $targetRoot)) {
        Add-Change -Area $areaName -Identity $targetRoot -Property '(access)' -Status Failed -Message 'Admin share of the target is not reachable, config files not compared.'
        return
    }

    # Files that hold the customisations found in the field: transport tuning, MRS throttling, timeouts in the
    # proxy web.config files, transport agent registrations and the OWA logon page.
    $files = @(
        'Bin\EdgeTransport.exe.config',
        'Bin\MSExchangeFrontendTransport.exe.config',
        'Bin\MSExchangeDelivery.exe.config',
        'Bin\MSExchangeSubmission.exe.config',
        'Bin\MSExchangeMailboxReplication.exe.config',
        'Bin\Microsoft.Exchange.Store.Worker.exe.config',
        'Bin\Microsoft.Exchange.Imap4.exe.config',
        'Bin\Microsoft.Exchange.Imap4Service.exe.config',
        'Bin\Microsoft.Exchange.Pop3.exe.config',
        'Bin\Microsoft.Exchange.Pop3Service.exe.config',
        'Bin\Microsoft.Exchange.Directory.TopologyService.exe.config',
        'Bin\Microsoft.Exchange.ServiceHost.exe.config',
        'TransportRoles\Shared\agents.config',
        'TransportRoles\Shared\fetagents.config',
        'FrontEnd\HttpProxy\SharedWebConfig.config',
        'FrontEnd\HttpProxy\web.config',
        'FrontEnd\HttpProxy\owa\web.config',
        'FrontEnd\HttpProxy\owa\auth\logon.aspx',
        'FrontEnd\HttpProxy\ecp\web.config',
        'FrontEnd\HttpProxy\ews\web.config',
        'FrontEnd\HttpProxy\rpc\web.config',
        'FrontEnd\HttpProxy\mapi\web.config',
        'FrontEnd\HttpProxy\sync\web.config',
        'FrontEnd\HttpProxy\oab\web.config',
        'FrontEnd\HttpProxy\Autodiscover\web.config',
        'FrontEnd\HttpProxy\PowerShell\web.config',
        'ClientAccess\SharedWebConfig.config',
        'ClientAccess\Owa\web.config',
        'ClientAccess\ecp\web.config',
        'ClientAccess\exchweb\ews\web.config',
        'ClientAccess\Sync\web.config',
        'ClientAccess\Autodiscover\web.config',
        'ClientAccess\mapi\emsmdb\web.config',
        'ClientAccess\mapi\nspi\web.config',
        'ClientAccess\PowerShell\web.config',
        'ClientAccess\rpc\web.config',
        'ClientAccess\OAB\web.config'
    ) + @($AdditionalConfigFile)

    $identical = 0
    foreach ($relativePath in $files) {
        $sourcePath = Join-Path $sourceRoot $relativePath
        $targetPath = Join-Path $targetRoot $relativePath
        if (-not (Test-Path $sourcePath)) { continue }
        if (-not (Test-Path $targetPath)) {
            $diffPath = Save-ConfigFileDiff -RelativePath $relativePath -SourcePath $sourcePath -TargetPath $targetPath
            Add-Change -Area $areaName -Identity $relativePath -Property '(file)' -SourceValue 'present' -TargetValue 'missing' -Status ReviewRequired -Message ('Copy saved: {0}' -f $diffPath)
            continue
        }
        $sourceHash = (Get-FileHash -Path $sourcePath -Algorithm SHA256).Hash
        $targetHash = (Get-FileHash -Path $targetPath -Algorithm SHA256).Hash
        if ($sourceHash -eq $targetHash) { $identical++; continue }
        $diffPath = Save-ConfigFileDiff -RelativePath $relativePath -SourcePath $sourcePath -TargetPath $targetPath
        Add-Change -Area $areaName -Identity $relativePath -Property '(content)' -SourceValue $sourceHash.Substring(0, 12) -TargetValue $targetHash.Substring(0, 12) -Status ReviewRequired `
            -Message ('Files differ, review the diff and apply the customisation by hand: {0}' -f $diffPath)
    }
    Write-Log -Message ('  {0} compared files are identical.' -f $identical)

    # OWA logon customisations (logo, background, css) live under owa\auth and its version folders.
    $authFolder = 'FrontEnd\HttpProxy\owa\auth'
    $sourceAuth = Join-Path $sourceRoot $authFolder
    $targetAuth = Join-Path $targetRoot $authFolder
    if ((Test-Path $sourceAuth) -and (Test-Path $targetAuth)) {
        $targetFiles = @{}
        foreach ($file in @(Get-ChildItem -Path $targetAuth -Recurse -File -ErrorAction SilentlyContinue)) {
            $targetFiles[$file.FullName.Substring($targetAuth.Length).TrimStart('\')] = $file
        }
        $differences = 0
        foreach ($file in @(Get-ChildItem -Path $sourceAuth -Recurse -File -ErrorAction SilentlyContinue)) {
            $relative = $file.FullName.Substring($sourceAuth.Length).TrimStart('\')
            $counterpart = $targetFiles[$relative]
            $reason = $null
            if ($null -eq $counterpart) { $reason = 'file exists on the source only' }
            elseif ($counterpart.Length -ne $file.Length) { $reason = 'file size differs' }
            elseif ($file.Length -lt 5MB -and (Get-FileHash -Path $file.FullName -Algorithm SHA256).Hash -ne (Get-FileHash -Path $counterpart.FullName -Algorithm SHA256).Hash) { $reason = 'content differs' }
            if ($null -eq $reason) { continue }
            $differences++
            $copyFolder = Join-Path (Join-Path $script:OutputFolder 'ConfigFiles') (Join-Path $authFolder (Split-Path -Path $relative -Parent))
            if (-not (Test-Path $copyFolder)) { New-Item -Path $copyFolder -ItemType Directory -Force | Out-Null }
            Copy-Item -Path $file.FullName -Destination (Join-Path $copyFolder $file.Name) -Force
            Add-Change -Area $areaName -Identity (Join-Path $authFolder $relative) -Property '(owa logon)' -SourceValue 'present' -TargetValue $(if ($counterpart) { 'different' } else { 'missing' }) `
                -Status ReviewRequired -Message ('OWA logon customisation, {0}. Source copy saved under ConfigFiles in the output folder.' -f $reason)
        }
        if ($differences -eq 0) { Write-Log -Message '  OWA logon folder is identical.' }
    }

    # Custom OWA themes are extra folders under prem\<version>\resources\themes.
    $sourcePrem = Join-Path $sourceRoot 'ClientAccess\Owa\prem'
    $targetPrem = Join-Path $targetRoot 'ClientAccess\Owa\prem'
    if ((Test-Path $sourcePrem) -and (Test-Path $targetPrem)) {
        $sourceThemes = @(Get-ChildItem -Path (Join-Path $sourcePrem '*\resources\themes') -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object -Unique)
        $targetThemes = @(Get-ChildItem -Path (Join-Path $targetPrem '*\resources\themes') -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object -Unique)
        foreach ($theme in $sourceThemes) {
            if ($targetThemes -contains $theme) { continue }
            Add-Change -Area $areaName -Identity ('ClientAccess\Owa\prem\*\resources\themes\{0}' -f $theme) -Property '(owa theme)' -SourceValue 'present' -TargetValue 'missing' `
                -Status ReviewRequired -Message 'Custom OWA theme folder, copy it into the same version folder on the target.'
        }
    }
}

# ---------------------------------------------------------------------------
# Area: Iis
# ---------------------------------------------------------------------------

$script:IisReadScript = {
    Import-Module WebAdministration -ErrorAction Stop
    $site = 'IIS:\Sites\Default Web Site'
    $redirect = Get-WebConfiguration -Filter 'system.webServer/httpRedirect' -PSPath $site
    $sslRaw = Get-WebConfigurationProperty -Filter 'system.webServer/security/access' -Name sslFlags -PSPath $site
    $sslValue = $sslRaw
    if ($null -ne $sslRaw -and $null -ne $sslRaw.PSObject.Properties['Value']) { $sslValue = $sslRaw.Value }
    [pscustomobject]@{
        RedirectEnabled          = [bool]$redirect.enabled
        RedirectDestination      = [string]$redirect.destination
        RedirectExactDestination = [bool]$redirect.exactDestination
        RedirectChildOnly        = [bool]$redirect.childOnly
        RedirectStatus           = [string]$redirect.httpResponseStatus
        SslFlags                 = [string]$sslValue
        Bindings                 = @(Get-WebBinding -Name 'Default Web Site' | ForEach-Object { '{0} {1}' -f $_.protocol, $_.bindingInformation })
    }
}

$script:IisApplyScript = {
    param($Settings, [bool] $ApplyRedirect, [bool] $ApplySsl)
    Import-Module WebAdministration -ErrorAction Stop
    $site = 'IIS:\Sites\Default Web Site'
    if ($ApplyRedirect) {
        Set-WebConfigurationProperty -Filter 'system.webServer/httpRedirect' -PSPath $site -Name enabled -Value $Settings.RedirectEnabled
        Set-WebConfigurationProperty -Filter 'system.webServer/httpRedirect' -PSPath $site -Name destination -Value $Settings.RedirectDestination
        Set-WebConfigurationProperty -Filter 'system.webServer/httpRedirect' -PSPath $site -Name exactDestination -Value $Settings.RedirectExactDestination
        Set-WebConfigurationProperty -Filter 'system.webServer/httpRedirect' -PSPath $site -Name childOnly -Value $Settings.RedirectChildOnly
        Set-WebConfigurationProperty -Filter 'system.webServer/httpRedirect' -PSPath $site -Name httpResponseStatus -Value $Settings.RedirectStatus
    }
    if ($ApplySsl) {
        Set-WebConfigurationProperty -Filter 'system.webServer/security/access' -PSPath $site -Name sslFlags -Value $Settings.SslFlags
    }
}

function Invoke-AreaIis {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $areaName = 'Iis'
    try {
        $sourceIis = Invoke-Command -ComputerName $script:SourceFqdn -ScriptBlock $script:IisReadScript -ErrorAction Stop
        $targetIis = Invoke-Command -ComputerName $script:TargetFqdn -ScriptBlock $script:IisReadScript -ErrorAction Stop
    }
    catch {
        Add-Change -Area $areaName -Identity 'Default Web Site' -Property '(read)' -Status Failed -Message ('WinRM read failed, compare the Default Web Site root redirect and SSL settings by hand: {0}' -f $_.Exception.Message)
        return
    }
    Save-Snapshot -Area $areaName -Source $sourceIis -Target $targetIis

    $redirectProperties = @('RedirectEnabled', 'RedirectDestination', 'RedirectExactDestination', 'RedirectChildOnly', 'RedirectStatus')
    $redirectDiffers = $false
    foreach ($property in $redirectProperties) {
        if ([string]$sourceIis.$property -ne [string]$targetIis.$property) { $redirectDiffers = $true }
    }
    $sslDiffers = ([string]$sourceIis.SslFlags -ne [string]$targetIis.SslFlags)

    if (-not $redirectDiffers -and -not $sslDiffers) {
        Write-Log -Message '  Default Web Site root redirect and SSL settings are in sync.'
    }
    else {
        $identity = '{0}\Default Web Site' -f $script:TargetShort
        $description = @()
        if ($redirectDiffers) { $description += 'httpRedirect' }
        if ($sslDiffers) { $description += 'sslFlags' }
        if ($PSCmdlet.ShouldProcess($identity, ('Set-WebConfigurationProperty ({0})' -f ($description -join ', ')))) {
            try {
                Invoke-Command -ComputerName $script:TargetFqdn -ScriptBlock $script:IisApplyScript -ArgumentList $sourceIis, $redirectDiffers, $sslDiffers -ErrorAction Stop
                $status = 'Applied'
            }
            catch {
                $status = 'Failed'
                $failure = $_.Exception.Message
            }
        }
        else { $status = Get-PendingStatus }

        foreach ($property in $redirectProperties) {
            if ([string]$sourceIis.$property -eq [string]$targetIis.$property) { continue }
            Add-Change -Area $areaName -Identity $identity -Property $property -SourceValue $sourceIis.$property -TargetValue $targetIis.$property -Status $status -Message $failure
        }
        if ($sslDiffers) {
            Add-Change -Area $areaName -Identity $identity -Property 'SslFlags' -SourceValue $sourceIis.SslFlags -TargetValue $targetIis.SslFlags -Status $status -Message $failure
        }
    }

    # Bindings usually carry the same wildcard on both servers, a difference points at a host header or extra port.
    $sourceBindings = @($sourceIis.Bindings | Sort-Object)
    $targetBindings = @($targetIis.Bindings | Sort-Object)
    if ((Get-CompareKey $sourceBindings) -ne (Get-CompareKey $targetBindings)) {
        Add-Change -Area $areaName -Identity ('{0}\Default Web Site' -f $script:TargetShort) -Property 'Bindings' -SourceValue $sourceBindings -TargetValue $targetBindings `
            -Status ReviewRequired -Message 'Site bindings differ, align them in IIS Manager if the difference is intentional on the source.'
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if (-not (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    throw 'Exchange cmdlets are not available. Run this script from the Exchange Management Shell or import a remote Exchange session first.'
}

$script:SourceExchangeServer = Get-ExchangeServer -Identity $SourceServer -ErrorAction Stop
$script:TargetExchangeServer = Get-ExchangeServer -Identity $TargetServer -ErrorAction Stop
$script:SourceShort = [string]$script:SourceExchangeServer.Name
$script:TargetShort = [string]$script:TargetExchangeServer.Name
$script:SourceFqdn  = [string]$script:SourceExchangeServer.Fqdn
$script:TargetFqdn  = [string]$script:TargetExchangeServer.Fqdn

if ($script:SourceShort -eq $script:TargetShort) { throw 'Source and target are the same server.' }

if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location).Path ('ExchangeServerConfigCopy_{0}_to_{1}_{2}' -f $script:SourceShort, $script:TargetShort, (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
if (-not (Test-Path $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
$script:OutputFolder = (Resolve-Path $OutputPath).Path

$transcriptStarted = $false
try { Start-Transcript -Path (Join-Path $script:OutputFolder 'transcript.log') -Append | Out-Null; $transcriptStarted = $true } catch { }

# Random password for the in memory certificate transfer when the operator did not supply one. It is never printed
# and no PFX file is written in that case.
if ($CertificatePassword) { $script:CertificateSecret = $CertificatePassword }
else {
    $characters = @([char[]]'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789')
    $random = -join (1..32 | ForEach-Object { $characters | Get-Random })
    $script:CertificateSecret = ConvertTo-SecureString -String $random -AsPlainText -Force
}

Write-Log -Level Section -Message ('Copy Exchange server configuration: {0} to {1}' -f $script:SourceFqdn, $script:TargetFqdn)
Write-Log -Message ('Source version: {0}' -f $script:SourceExchangeServer.AdminDisplayVersion)
Write-Log -Message ('Target version: {0}' -f $script:TargetExchangeServer.AdminDisplayVersion)
Write-Log -Message ('Output folder : {0}' -f $script:OutputFolder)
if ($WhatIfPreference) { Write-Log -Level Warn -Message 'WhatIf mode, nothing will be changed on the target.' }
if ([string]$script:SourceExchangeServer.AdminDisplayVersion -ne [string]$script:TargetExchangeServer.AdminDisplayVersion) {
    Write-Log -Level Warn -Message 'Source and target run a different build. Settings that only exist in one build are reported as failed, review those by hand.'
}

$areaFunctions = @{
    Certificates       = 'Invoke-AreaCertificates'
    ExchangeServer     = 'Invoke-AreaExchangeServer'
    ClientAccess       = 'Invoke-AreaClientAccess'
    VirtualDirectories = 'Invoke-AreaVirtualDirectories'
    OutlookAnywhere    = 'Invoke-AreaOutlookAnywhere'
    Transport          = 'Invoke-AreaTransport'
    PopImap            = 'Invoke-AreaPopImap'
    MailboxServer      = 'Invoke-AreaMailboxServer'
    Malware            = 'Invoke-AreaMalware'
    EventLogLevels     = 'Invoke-AreaEventLogLevels'
    SettingOverrides   = 'Invoke-AreaSettingOverrides'
    ReceiveConnectors  = 'Invoke-AreaReceiveConnectors'
    SendConnectors     = 'Invoke-AreaSendConnectors'
    TransportAgents    = 'Invoke-AreaTransportAgents'
    ConfigFiles        = 'Invoke-AreaConfigFiles'
    Iis                = 'Invoke-AreaIis'
    Hybrid             = 'Invoke-AreaHybrid'
}

foreach ($areaName in $script:AllAreas) {
    if ($Area -and $Area -notcontains $areaName) { continue }
    if ($SkipArea -contains $areaName) { continue }
    Write-Log -Level Section -Message $areaName
    try {
        & $areaFunctions[$areaName]
    }
    catch {
        Write-Log -Level Error -Message ('{0}: {1}' -f $areaName, $_.Exception.Message)
        Add-Change -Area $areaName -Identity $script:TargetShort -Property '(area)' -Status Failed -Message $_.Exception.Message
    }
}

# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------

Write-Log -Level Section -Message 'Summary'

$changesPath = Join-Path $script:OutputFolder 'changes.csv'
$script:Changes | Export-Csv -Path $changesPath -NoTypeInformation -Encoding UTF8

try {
    $snapshotSource = [ordered]@{}
    $snapshotTarget = [ordered]@{}
    foreach ($key in $script:Snapshots.Keys) {
        $snapshotSource[$key] = $script:Snapshots[$key].Source
        $snapshotTarget[$key] = $script:Snapshots[$key].Target
    }
    $snapshotSource | ConvertTo-Json -Depth 3 -WarningAction SilentlyContinue | Set-Content -Path (Join-Path $script:OutputFolder ('{0}-config.json' -f $script:SourceShort)) -Encoding UTF8
    $snapshotTarget | ConvertTo-Json -Depth 3 -WarningAction SilentlyContinue | Set-Content -Path (Join-Path $script:OutputFolder ('{0}-config-before.json' -f $script:TargetShort)) -Encoding UTF8
}
catch { Write-Log -Level Warn -Message ('Snapshot export failed: {0}' -f $_.Exception.Message) }

$summary = foreach ($group in ($script:Changes | Group-Object -Property Area)) {
    [pscustomobject]@{
        Area           = $group.Name
        Applied        = @($group.Group | Where-Object Status -eq 'Applied').Count
        WhatIf         = @($group.Group | Where-Object Status -eq 'WhatIf').Count
        Failed         = @($group.Group | Where-Object Status -eq 'Failed').Count
        ReviewRequired = @($group.Group | Where-Object Status -eq 'ReviewRequired').Count
        Skipped        = @($group.Group | Where-Object Status -eq 'Skipped').Count
        Info           = @($group.Group | Where-Object Status -eq 'Info').Count
    }
}
if ($summary) { $summary | Format-Table -AutoSize | Out-String | Write-Host }
else { Write-Log -Message 'No differences found, the servers are in sync for the selected areas.' }

Write-Log -Message ('Change report : {0}' -f $changesPath)
Write-Log -Message ('Output folder : {0}' -f $script:OutputFolder)

$reviewCount = @($script:Changes | Where-Object Status -eq 'ReviewRequired').Count
$failedCount = @($script:Changes | Where-Object Status -eq 'Failed').Count
if ($reviewCount -gt 0) { Write-Log -Level Warn -Message ('{0} item(s) need a manual follow up, filter changes.csv on Status = ReviewRequired.' -f $reviewCount) }
if ($failedCount -gt 0) { Write-Log -Level Error -Message ('{0} item(s) failed, filter changes.csv on Status = Failed.' -f $failedCount) }

if ($script:RestartHints.Count -gt 0) {
    Write-Log -Level Section -Message ('Follow up on {0}' -f $script:TargetShort)
    foreach ($hint in $script:RestartHints) { Write-Host ('  - {0}' -f $hint) }
}
if ($WhatIfPreference) {
    Write-Log -Message 'Run again without -WhatIf to apply. Consider -SkipArea SendConnectors until the target is tested.'
}
else {
    Write-Log -Message 'Run again with -WhatIf to confirm both servers are in sync. Test with Test-OutlookConnectivity, Test-OwaConnectivity and Test-ActiveSyncConnectivity before adding the target to the load balancer.'
}

if ($transcriptStarted) { try { Stop-Transcript | Out-Null } catch { } }
