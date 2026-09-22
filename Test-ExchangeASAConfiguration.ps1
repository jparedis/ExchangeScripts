<#
.SYNOPSIS
    Validates the Alternate Service Account (ASA) configuration used for Kerberos authentication
    on Exchange Server, and reports every finding as a pass, warning or failure.

.DESCRIPTION
    Kerberos authentication for load balanced Exchange namespaces depends on a single Alternate
    Service Account (ASA), shared by every Client Access endpoint in the site. The configuration
    breaks silently in a number of well known ways: the credential is deployed on some servers but
    not on all of them, the password was rolled in Active Directory but never pushed to every
    server, a required SPN is missing, an SPN ended up as a duplicate on a computer account, or
    Negotiate was never enabled on the virtual directories.

    This script is read only. It collects the current state and validates it:

      1. Exchange management shell and the correct CAS cmdlet are available, and the ASA
         credential can be read on every server at all
      2. Every Client Access service has an ASA credential deployed
      3. All servers use the same ASA account (mismatch breaks Kerberos for part of the farm)
      4. All servers carry the same credential generation (a rolled password reached every server)
      5. The credential age stays below the rotation threshold
      6. The ASA account exists in Active Directory, is enabled, and its pwdLastSet matches the
         credential that is actually deployed on the servers
      7. Every client namespace (Outlook Anywhere, MAPI, EWS, OAB, Autodiscover) has a matching
         http/ SPN on the ASA account
      8. None of those SPNs is registered on a second account (duplicate SPN kills Kerberos)
      9. Outlook Anywhere and the MAPI virtual directories offer Negotiate

    Active Directory is queried through System.DirectoryServices, so the ActiveDirectory module
    and RSAT are not required. SPN duplicates are searched forest wide through the global catalog.

.PARAMETER Server
    One or more Exchange servers to validate. Default is every Client Access service in the
    organization. The credential and Negotiate checks honour this scope, the namespace and SPN
    checks stay organization wide because a shared ASA account is by definition shared.

.PARAMETER ExpectedASAAccount
    The ASA account you expect to find, for example CONTOSO\EXCHANGE$. When omitted, the account
    is taken from the servers themselves and only checked for consistency between them.

.PARAMETER PasswordMaximumAgeDays
    Age in days above which the deployed credential is flagged. Default 90, matching the common
    rotation schedule of RollAlternateServiceAccountPassword.ps1.

.PARAMETER AdditionalNamespace
    Extra namespaces that must have an http/ SPN on the ASA account, for example a legacy or
    migration namespace that is not returned by the virtual directory cmdlets.

.PARAMETER SkipSPNCheck
    Skips every Active Directory lookup. Use this when the account running the script has no
    read access to the directory.

.PARAMETER ReportPath
    Writes the findings to this path. A .csv extension produces CSV, anything else produces HTML.

.EXAMPLE
    .\Test-ExchangeASAConfiguration.ps1

    Validates the whole organization and prints the findings.

.EXAMPLE
    .\Test-ExchangeASAConfiguration.ps1 -ExpectedASAAccount 'CONTOSO\EXCHANGE$' -ReportPath C:\Temp\asa.html

    Validates against a known ASA account and writes an HTML report.

.EXAMPLE
    .\Test-ExchangeASAConfiguration.ps1 -Server EX01,EX02 -PasswordMaximumAgeDays 30

    Validates two servers with a stricter rotation threshold.

.OUTPUTS
    PSCustomObject per finding: Category, Target, Check, Status, Details.

.NOTES
    Author: Jente Paredis - jente@jentech.be

    The ASA credential is stored in the registry of every server, not in Active Directory, so
    reading it reaches out to each server in turn. A server that is offline, that has the
    RemoteRegistry service stopped, or on which you have no rights, is reported as a single
    finding while the rest of the organization is still validated.

    Requirements:
    - Run inside the Exchange Management Shell, or in a session with the Exchange cmdlets
      imported through implicit remoting
    - Rights to read the registry of every Exchange server, otherwise those servers come back
      as a readable finding instead of a credential state
    - View only Exchange permissions are sufficient, the script never writes
    - Directory read access for the SPN checks, otherwise use -SkipSPNCheck
    - Windows PowerShell 5.1

    Tested against Exchange Server 2016 and 2019. Exchange 2013 is supported through the
    Get-ClientAccessServer fallback.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$Server,

    [Parameter()]
    [string]$ExpectedASAAccount,

    [Parameter()]
    [ValidateRange(1, 3650)]
    [int]$PasswordMaximumAgeDays = 90,

    [Parameter()]
    [string[]]$AdditionalNamespace,

    [Parameter()]
    [switch]$SkipSPNCheck,

    [Parameter()]
    [string]$ReportPath
)

#region helpers

# Every check funnels through this function so the output stays one single shape.
function New-ASAFinding {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$Check,
        [Parameter(Mandatory)][ValidateSet('Pass', 'Warning', 'Fail', 'Info')][string]$Status,
        [string]$Details
    )

    [pscustomobject]@{
        Category = $Category
        Target   = $Target
        Check    = $Check
        Status   = $Status
        Details  = $Details
    }
}

# Exchange 2016 and later use Get-ClientAccessService, Exchange 2013 uses Get-ClientAccessServer.
function Resolve-ClientAccessCmdlet {
    foreach ($candidate in @('Get-ClientAccessService', 'Get-ClientAccessServer')) {
        $command = Get-Command -Name $candidate -ErrorAction SilentlyContinue
        if ($command) { return $command.Name }
    }
    return $null
}

# The ASA configuration object exposes its credentials as a collection. Older builds only render
# a multi line string, so the string form is parsed as a fallback.
function ConvertTo-ASACredentialList {
    param($Configuration)

    $credentialList = @()
    if (-not $Configuration) { return $credentialList }

    $effective = $null
    if ($Configuration.PSObject.Properties.Match('EffectiveCredentials').Count -gt 0) {
        $effective = $Configuration.EffectiveCredentials
    }

    if ($effective) {
        foreach ($entry in $effective) {
            $userName = $null
            if ($entry.PSObject.Properties.Match('Credential').Count -gt 0 -and $entry.Credential) {
                $userName = $entry.Credential.UserName
            }

            $whenAdded = $null
            if ($entry.PSObject.Properties.Match('WhenAddedUTC').Count -gt 0 -and $entry.WhenAddedUTC) {
                $whenAdded = [datetime]$entry.WhenAddedUTC
            }

            $credentialList += [pscustomobject]@{
                UserName     = $userName
                WhenAddedUtc = $whenAdded
            }
        }

        return $credentialList
    }

    # Fallback: parse lines such as "Latest: 3/12/2025 5:20:15 PM, CONTOSO\EXCHANGE$".
    foreach ($line in ($Configuration.ToString() -split "`r?`n")) {
        $match = [regex]::Match($line, '^\s*(Latest|Previous)\s*:\s*(?<stamp>.+?)\s*,\s*(?<user>\S+)\s*$')
        if (-not $match.Success) { continue }
        if ($match.Groups['user'].Value -match '^<') { continue }

        $whenAdded = $null
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($match.Groups['stamp'].Value, [ref]$parsed)) {
            $whenAdded = $parsed.ToUniversalTime()
        }

        $credentialList += [pscustomobject]@{
            UserName     = $match.Groups['user'].Value
            WhenAddedUtc = $whenAdded
        }
    }

    return $credentialList
}

# Returns the sAMAccountName part of DOMAIN\account, account@domain or a bare account name.
function ConvertTo-SamAccountName {
    param([string]$UserName)

    if ([string]::IsNullOrWhiteSpace($UserName)) { return $null }
    if ($UserName -match '\\') { return $UserName.Split('\')[-1] }
    if ($UserName -match '@')  { return $UserName.Split('@')[0] }
    return $UserName
}

# Directory lookup through System.DirectoryServices, so RSAT is not a requirement.
function Get-ASADirectoryAccount {
    param([Parameter(Mandatory)][string]$SamAccountName)

    $rootDse = New-Object System.DirectoryServices.DirectoryEntry('LDAP://RootDSE')
    $searchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($rootDse.defaultNamingContext)")

    $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
    $searcher.Filter = "(sAMAccountName=$SamAccountName)"
    $searcher.PageSize = 100
    foreach ($property in @('samaccountname', 'distinguishedname', 'useraccountcontrol', 'pwdlastset', 'serviceprincipalname', 'objectclass')) {
        $null = $searcher.PropertiesToLoad.Add($property)
    }

    $result = $searcher.FindOne()
    if (-not $result) { return $null }

    $userAccountControl = 0
    if ($result.Properties['useraccountcontrol'].Count -gt 0) {
        $userAccountControl = [int]$result.Properties['useraccountcontrol'][0]
    }

    $passwordLastSet = $null
    if ($result.Properties['pwdlastset'].Count -gt 0) {
        $rawValue = [int64]$result.Properties['pwdlastset'][0]
        if ($rawValue -gt 0) { $passwordLastSet = [datetime]::FromFileTimeUtc($rawValue) }
    }

    $servicePrincipalNames = @()
    if ($result.Properties['serviceprincipalname'].Count -gt 0) {
        $servicePrincipalNames = @($result.Properties['serviceprincipalname'])
    }

    [pscustomobject]@{
        SamAccountName        = [string]$result.Properties['samaccountname'][0]
        DistinguishedName     = [string]$result.Properties['distinguishedname'][0]
        # bit 2 of userAccountControl is ACCOUNTDISABLE
        Enabled               = -not ($userAccountControl -band 0x2)
        PasswordLastSetUtc    = $passwordLastSet
        ServicePrincipalNames = $servicePrincipalNames
        ObjectClasses         = @($result.Properties['objectclass'])
    }
}

# Forest wide SPN lookup through the global catalog, the equivalent of setspn -F -Q.
function Get-SPNRegistrationOwner {
    param([Parameter(Mandatory)][string]$ServicePrincipalName)

    $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $searchRoot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($forest.Name)")

    $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
    $searcher.Filter = "(servicePrincipalName=$ServicePrincipalName)"
    $searcher.PageSize = 100
    $null = $searcher.PropertiesToLoad.Add('samaccountname')
    $null = $searcher.PropertiesToLoad.Add('distinguishedname')

    $owners = @()
    foreach ($result in $searcher.FindAll()) {
        $owners += [pscustomobject]@{
            SamAccountName    = [string]$result.Properties['samaccountname'][0]
            DistinguishedName = [string]$result.Properties['distinguishedname'][0]
        }
    }

    return $owners
}

# Pulls the host out of a URL and keeps a bare hostname as is.
function ConvertTo-HostName {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $candidate = $Value.ToString().Trim()
    if ($candidate -match '^[a-z]+://') {
        try { return ([uri]$candidate).Host } catch { return $null }
    }
    return $candidate
}

# The ASA credential itself is not stored in Active Directory but in the registry of every
# server, under MSExchangeServiceHost\ServiceAccounts. When Exchange cannot read that subtree it
# returns one single error, without telling you whether the server is unreachable or whether no
# credential was ever deployed. This read only probe answers that question.
function Test-ASARegistryState {
    param([Parameter(Mandatory)][string]$ComputerName)

    $state = [pscustomobject]@{
        RegistryReachable         = $false
        ServiceAccountsKeyPresent = $false
        AccountKeyCount           = 0
        ProbeError                = $null
    }

    $baseKey = $null
    $serviceAccountsKey = $null

    try {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName)
        $state.RegistryReachable = $true

        $serviceAccountsKey = $baseKey.OpenSubKey('SYSTEM\CurrentControlSet\Services\MSExchangeServiceHost\ServiceAccounts')
        if ($serviceAccountsKey) {
            $state.ServiceAccountsKeyPresent = $true
            $state.AccountKeyCount = @($serviceAccountsKey.GetSubKeyNames()).Count
        }
    }
    catch {
        $state.ProbeError = $_.Exception.Message
    }
    finally {
        if ($serviceAccountsKey) { $serviceAccountsKey.Close() }
        if ($baseKey) { $baseKey.Close() }
    }

    return $state
}

# Limits a per server check to the servers passed through -Server, matching on the short name.
function Test-ServerInScope {
    param(
        [string]$TargetServer,
        [string[]]$Scope
    )

    if (-not $Scope -or $Scope.Count -eq 0) { return $true }
    if ([string]::IsNullOrWhiteSpace($TargetServer)) { return $true }

    $targetShortName = $TargetServer.Split('.')[0]
    foreach ($scopeEntry in $Scope) {
        if ($scopeEntry.Split('.')[0] -eq $targetShortName) { return $true }
    }

    return $false
}

function Write-ASAFindingToHost {
    param([Parameter(Mandatory)]$Finding)

    $color = 'Gray'
    switch ($Finding.Status) {
        'Pass'    { $color = 'Green' }
        'Warning' { $color = 'Yellow' }
        'Fail'    { $color = 'Red' }
        'Info'    { $color = 'Cyan' }
    }

    Write-Host ('{0,-7} {1,-22} {2,-44} {3}' -f $Finding.Status, $Finding.Target, $Finding.Check, $Finding.Details) -ForegroundColor $color
}

#endregion helpers

$findings = @()
$scriptStart = Get-Date

Write-Host ''
Write-Host 'Exchange Alternate Service Account validation' -ForegroundColor White
Write-Host ('Started {0}' -f $scriptStart) -ForegroundColor DarkGray
Write-Host ''

#region prerequisites

$clientAccessCmdlet = Resolve-ClientAccessCmdlet
if (-not $clientAccessCmdlet) {
    throw 'The Exchange cmdlets are not available in this session. Run this script in the Exchange Management Shell.'
}

$findings += New-ASAFinding -Category 'Prerequisites' -Target 'Session' -Check 'Exchange cmdlets available' -Status 'Pass' -Details "Using $clientAccessCmdlet"

#endregion prerequisites

#region collect

# The server list comes from Active Directory only. The ASA credential is read per server further
# down, because that read touches the registry of every single server and one unreachable server
# would otherwise hide the state of the whole organization.
$clientAccessServices = @()
try {
    if ($Server) {
        foreach ($serverName in $Server) {
            $clientAccessServices += & $clientAccessCmdlet -Identity $serverName -ErrorAction Stop
        }
    }
    else {
        $clientAccessServices = @(& $clientAccessCmdlet -ErrorAction Stop)
    }
}
catch {
    throw "Unable to enumerate the Client Access services: $($_.Exception.Message)"
}

if (-not $clientAccessServices -or $clientAccessServices.Count -eq 0) {
    throw 'No Client Access services were returned. Check the -Server value or your Exchange permissions.'
}

$findings += New-ASAFinding -Category 'Prerequisites' -Target 'Organization' -Check 'Client Access services collected' -Status 'Info' -Details ('{0} server(s): {1}' -f $clientAccessServices.Count, (($clientAccessServices | ForEach-Object { $_.Name }) -join ', '))

#endregion collect

#region per server credential state

# Deployed credentials are kept per server so the consistency checks can compare them afterwards.
$deployedState = @()

foreach ($clientAccessService in $clientAccessServices) {
    $serverName = [string]$clientAccessService.Name

    $asaConfiguration = $null
    $credentialReadError = $null

    try {
        $detailedService = & $clientAccessCmdlet -Identity $serverName -IncludeAlternateServiceAccountCredentialStatus -ErrorAction Stop
        $asaConfiguration = $detailedService.AlternateServiceAccountConfiguration
    }
    catch {
        $credentialReadError = $_.Exception.Message
    }

    # Exchange failed the registry read, so work out why instead of reporting one opaque error.
    if ($credentialReadError) {
        $registryState = Test-ASARegistryState -ComputerName $serverName

        if (-not $registryState.RegistryReachable) {
            $diagnosis = 'The registry of this server cannot be read from this session. Check that the server is online, that the RemoteRegistry service runs, that the firewall allows remote registry, and that your account has rights on it.'
            if ($registryState.ProbeError) { $diagnosis += " Probe error: $($registryState.ProbeError)" }
            $status = 'Fail'
        }
        elseif (-not $registryState.ServiceAccountsKeyPresent) {
            $diagnosis = 'The registry is reachable but the ServiceAccounts subtree does not exist, so no ASA credential was ever deployed on this server.'
            $status = 'Fail'
        }
        else {
            $diagnosis = 'The registry is reachable and the ServiceAccounts subtree exists with {0} account key(s), so this is a rights issue on that subtree rather than a missing credential.' -f $registryState.AccountKeyCount
            $status = 'Warning'
        }

        $findings += New-ASAFinding -Category 'Credential' -Target $serverName -Check 'ASA credential readable' -Status $status -Details ('{0} Exchange returned: {1}' -f $diagnosis, $credentialReadError)

        $deployedState += [pscustomobject]@{
            Server       = $serverName
            UserName     = $null
            WhenAddedUtc = $null
        }

        continue
    }

    $credentialList = @(ConvertTo-ASACredentialList -Configuration $asaConfiguration)

    if ($credentialList.Count -eq 0) {
        $findings += New-ASAFinding -Category 'Credential' -Target $serverName -Check 'ASA credential deployed' -Status 'Fail' -Details 'No alternate service account credential is configured on this server.'
        $deployedState += [pscustomobject]@{
            Server       = $serverName
            UserName     = $null
            WhenAddedUtc = $null
        }
        continue
    }

    # The most recently added credential is the one Kerberos uses.
    $currentCredential = $credentialList |
        Sort-Object -Property @{ Expression = { if ($_.WhenAddedUtc) { $_.WhenAddedUtc } else { [datetime]::MinValue } } } -Descending |
        Select-Object -First 1

    $deployedState += [pscustomobject]@{
        Server       = $serverName
        UserName     = $currentCredential.UserName
        WhenAddedUtc = $currentCredential.WhenAddedUtc
    }

    $findings += New-ASAFinding -Category 'Credential' -Target $serverName -Check 'ASA credential deployed' -Status 'Pass' -Details ('Account {0}, {1} credential(s) present' -f $currentCredential.UserName, $credentialList.Count)

    if ($currentCredential.WhenAddedUtc) {
        $credentialAgeDays = [int]((Get-Date).ToUniversalTime() - $currentCredential.WhenAddedUtc).TotalDays
        $ageDetails = 'Deployed {0} UTC, {1} day(s) old, threshold {2}' -f $currentCredential.WhenAddedUtc, $credentialAgeDays, $PasswordMaximumAgeDays

        if ($credentialAgeDays -gt $PasswordMaximumAgeDays) {
            $findings += New-ASAFinding -Category 'Credential' -Target $serverName -Check 'Credential age within threshold' -Status 'Warning' -Details $ageDetails
        }
        else {
            $findings += New-ASAFinding -Category 'Credential' -Target $serverName -Check 'Credential age within threshold' -Status 'Pass' -Details $ageDetails
        }
    }
    else {
        $findings += New-ASAFinding -Category 'Credential' -Target $serverName -Check 'Credential age within threshold' -Status 'Warning' -Details 'The deployment timestamp could not be read, so the age could not be validated.'
    }

    if ($ExpectedASAAccount) {
        $expectedSam = ConvertTo-SamAccountName -UserName $ExpectedASAAccount
        $actualSam = ConvertTo-SamAccountName -UserName $currentCredential.UserName

        if ($actualSam -and $expectedSam -and $actualSam -eq $expectedSam) {
            $findings += New-ASAFinding -Category 'Credential' -Target $serverName -Check 'Matches expected ASA account' -Status 'Pass' -Details ('Expected and deployed account are both {0}' -f $expectedSam)
        }
        else {
            $findings += New-ASAFinding -Category 'Credential' -Target $serverName -Check 'Matches expected ASA account' -Status 'Fail' -Details ('Expected {0} but found {1}' -f $ExpectedASAAccount, $currentCredential.UserName)
        }
    }
}

#endregion per server credential state

#region organization consistency

$serversWithCredential = @($deployedState | Where-Object { $_.UserName })

if ($serversWithCredential.Count -eq 0) {
    $findings += New-ASAFinding -Category 'Consistency' -Target 'Organization' -Check 'ASA account identical on all servers' -Status 'Fail' -Details 'Not a single server carries an ASA credential, Kerberos authentication cannot work.'
}
else {
    $distinctAccounts = @($serversWithCredential | ForEach-Object { ConvertTo-SamAccountName -UserName $_.UserName } | Sort-Object -Unique)

    if ($distinctAccounts.Count -eq 1) {
        $findings += New-ASAFinding -Category 'Consistency' -Target 'Organization' -Check 'ASA account identical on all servers' -Status 'Pass' -Details ('All servers use {0}' -f $distinctAccounts[0])
    }
    else {
        $accountOverview = ($serversWithCredential | ForEach-Object { '{0}={1}' -f $_.Server, $_.UserName }) -join '; '
        $findings += New-ASAFinding -Category 'Consistency' -Target 'Organization' -Check 'ASA account identical on all servers' -Status 'Fail' -Details ('Multiple accounts in use: {0}' -f $accountOverview)
    }

    # A server that missed the last password roll keeps an older credential and starts failing
    # Kerberos as soon as the previous password leaves the account.
    if ($serversWithCredential.Count -gt 1) {
        $missingTimestamp = @($serversWithCredential | Where-Object { -not $_.WhenAddedUtc })

        if ($missingTimestamp.Count -gt 0) {
            $findings += New-ASAFinding -Category 'Consistency' -Target 'Organization' -Check 'Same credential generation on all servers' -Status 'Warning' -Details ('No timestamp available for: {0}' -f (($missingTimestamp | ForEach-Object { $_.Server }) -join ', '))
        }
        else {
            $newestDeployment = ($serversWithCredential | Sort-Object -Property WhenAddedUtc -Descending | Select-Object -First 1).WhenAddedUtc
            $staleServers = @($serversWithCredential | Where-Object { ($newestDeployment - $_.WhenAddedUtc).TotalHours -gt 1 })

            if ($staleServers.Count -eq 0) {
                $findings += New-ASAFinding -Category 'Consistency' -Target 'Organization' -Check 'Same credential generation on all servers' -Status 'Pass' -Details ('All servers carry the credential deployed on {0} UTC' -f $newestDeployment)
            }
            else {
                $staleOverview = ($staleServers | ForEach-Object { '{0}={1} UTC' -f $_.Server, $_.WhenAddedUtc }) -join '; '
                $findings += New-ASAFinding -Category 'Consistency' -Target 'Organization' -Check 'Same credential generation on all servers' -Status 'Fail' -Details ('Newest credential is from {0} UTC, behind: {1}' -f $newestDeployment, $staleOverview)
            }
        }
    }
}

#endregion organization consistency

#region namespaces

# Kerberos only works for namespaces that carry an http/ SPN on the ASA account, so the list of
# namespaces is built from the virtual directories the clients actually connect to.
$namespaceList = @()

try {
    foreach ($outlookAnywhere in @(Get-OutlookAnywhere -ErrorAction Stop)) {
        $namespaceList += ConvertTo-HostName -Value $outlookAnywhere.InternalHostname
        $namespaceList += ConvertTo-HostName -Value $outlookAnywhere.ExternalHostname
    }
}
catch {
    $findings += New-ASAFinding -Category 'Namespace' -Target 'Organization' -Check 'Outlook Anywhere namespaces collected' -Status 'Warning' -Details "Get-OutlookAnywhere failed: $($_.Exception.Message)"
}

foreach ($virtualDirectoryCmdlet in @('Get-MapiVirtualDirectory', 'Get-WebServicesVirtualDirectory', 'Get-OabVirtualDirectory')) {
    if (-not (Get-Command -Name $virtualDirectoryCmdlet -ErrorAction SilentlyContinue)) { continue }

    try {
        foreach ($virtualDirectory in @(& $virtualDirectoryCmdlet -ErrorAction Stop)) {
            $namespaceList += ConvertTo-HostName -Value $virtualDirectory.InternalUrl
            $namespaceList += ConvertTo-HostName -Value $virtualDirectory.ExternalUrl
        }
    }
    catch {
        $findings += New-ASAFinding -Category 'Namespace' -Target 'Organization' -Check "$virtualDirectoryCmdlet namespaces collected" -Status 'Warning' -Details "$virtualDirectoryCmdlet failed: $($_.Exception.Message)"
    }
}

foreach ($clientAccessService in $clientAccessServices) {
    $namespaceList += ConvertTo-HostName -Value $clientAccessService.AutoDiscoverServiceInternalUri
}

if ($AdditionalNamespace) {
    foreach ($additional in $AdditionalNamespace) {
        $namespaceList += ConvertTo-HostName -Value $additional
    }
}

$namespaceList = @($namespaceList | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)

$findings += New-ASAFinding -Category 'Namespace' -Target 'Organization' -Check 'Client namespaces in use' -Status 'Info' -Details ($namespaceList -join ', ')

#endregion namespaces

#region active directory and SPN

$asaSamAccountName = $null
if ($ExpectedASAAccount) {
    $asaSamAccountName = ConvertTo-SamAccountName -UserName $ExpectedASAAccount
}
elseif ($serversWithCredential.Count -gt 0) {
    $asaSamAccountName = ConvertTo-SamAccountName -UserName $serversWithCredential[0].UserName
}

if ($SkipSPNCheck) {
    $findings += New-ASAFinding -Category 'ActiveDirectory' -Target 'Organization' -Check 'Directory validation' -Status 'Info' -Details 'Skipped on request (-SkipSPNCheck).'
}
elseif (-not $asaSamAccountName) {
    $findings += New-ASAFinding -Category 'ActiveDirectory' -Target 'Organization' -Check 'Directory validation' -Status 'Fail' -Details 'No ASA account name is known, so the directory checks were skipped.'
}
else {
    $directoryAccount = $null
    try {
        $directoryAccount = Get-ASADirectoryAccount -SamAccountName $asaSamAccountName
    }
    catch {
        $findings += New-ASAFinding -Category 'ActiveDirectory' -Target $asaSamAccountName -Check 'ASA account found in Active Directory' -Status 'Warning' -Details "Directory lookup failed: $($_.Exception.Message)"
    }

    if (-not $directoryAccount) {
        if (-not ($findings | Where-Object { $_.Check -eq 'ASA account found in Active Directory' })) {
            $findings += New-ASAFinding -Category 'ActiveDirectory' -Target $asaSamAccountName -Check 'ASA account found in Active Directory' -Status 'Fail' -Details 'The account deployed on the Exchange servers does not exist in this domain.'
        }
    }
    else {
        $findings += New-ASAFinding -Category 'ActiveDirectory' -Target $asaSamAccountName -Check 'ASA account found in Active Directory' -Status 'Pass' -Details $directoryAccount.DistinguishedName

        if ($directoryAccount.Enabled) {
            $findings += New-ASAFinding -Category 'ActiveDirectory' -Target $asaSamAccountName -Check 'ASA account enabled' -Status 'Pass' -Details 'The account is enabled.'
        }
        else {
            $findings += New-ASAFinding -Category 'ActiveDirectory' -Target $asaSamAccountName -Check 'ASA account enabled' -Status 'Fail' -Details 'The account is disabled, Kerberos ticket decryption will fail.'
        }

        # A password that is newer in Active Directory than on the servers means the roll script
        # changed the password but never pushed it to Exchange.
        $newestDeployedCredential = $null
        if ($serversWithCredential.Count -gt 0) {
            $newestDeployedCredential = ($serversWithCredential | Sort-Object -Property WhenAddedUtc -Descending | Select-Object -First 1).WhenAddedUtc
        }

        if ($directoryAccount.PasswordLastSetUtc -and $newestDeployedCredential) {
            $driftHours = ($directoryAccount.PasswordLastSetUtc - $newestDeployedCredential).TotalHours
            $driftDetails = 'pwdLastSet {0} UTC, newest deployed credential {1} UTC' -f $directoryAccount.PasswordLastSetUtc, $newestDeployedCredential

            if ($driftHours -gt 1) {
                $findings += New-ASAFinding -Category 'ActiveDirectory' -Target $asaSamAccountName -Check 'Directory password matches deployed credential' -Status 'Fail' -Details ('The password was changed after the last deployment. ' + $driftDetails)
            }
            else {
                $findings += New-ASAFinding -Category 'ActiveDirectory' -Target $asaSamAccountName -Check 'Directory password matches deployed credential' -Status 'Pass' -Details $driftDetails
            }
        }

        if ($directoryAccount.PasswordLastSetUtc) {
            $directoryPasswordAgeDays = [int]((Get-Date).ToUniversalTime() - $directoryAccount.PasswordLastSetUtc).TotalDays
            $findings += New-ASAFinding -Category 'ActiveDirectory' -Target $asaSamAccountName -Check 'Directory password age' -Status 'Info' -Details ('{0} day(s) since the last password change' -f $directoryPasswordAgeDays)
        }

        # SPN coverage: every client namespace needs http/<namespace> on the ASA account.
        $registeredSpn = @($directoryAccount.ServicePrincipalNames | ForEach-Object { $_.ToString().ToLowerInvariant() })

        $findings += New-ASAFinding -Category 'SPN' -Target $asaSamAccountName -Check 'SPNs registered on the ASA account' -Status 'Info' -Details (($directoryAccount.ServicePrincipalNames | Sort-Object) -join ', ')

        foreach ($namespace in $namespaceList) {
            $expectedSpn = "http/$namespace"

            if ($registeredSpn -contains $expectedSpn) {
                $findings += New-ASAFinding -Category 'SPN' -Target $namespace -Check 'http SPN present on ASA account' -Status 'Pass' -Details "$expectedSpn is registered on $asaSamAccountName"
            }
            else {
                $findings += New-ASAFinding -Category 'SPN' -Target $namespace -Check 'http SPN present on ASA account' -Status 'Fail' -Details "$expectedSpn is missing, clients fall back to NTLM for this namespace"
            }

            # A second registration of the same SPN makes the KDC refuse to issue tickets.
            try {
                $spnOwners = @(Get-SPNRegistrationOwner -ServicePrincipalName $expectedSpn)
                $foreignOwners = @($spnOwners | Where-Object { $_.SamAccountName -and $_.SamAccountName -ne $asaSamAccountName })

                if ($foreignOwners.Count -eq 0) {
                    $findings += New-ASAFinding -Category 'SPN' -Target $namespace -Check 'No duplicate SPN registration' -Status 'Pass' -Details "$expectedSpn is registered once"
                }
                else {
                    $ownerOverview = ($foreignOwners | ForEach-Object { $_.DistinguishedName }) -join '; '
                    $findings += New-ASAFinding -Category 'SPN' -Target $namespace -Check 'No duplicate SPN registration' -Status 'Fail' -Details ("$expectedSpn is also registered on: " + $ownerOverview)
                }
            }
            catch {
                $findings += New-ASAFinding -Category 'SPN' -Target $namespace -Check 'No duplicate SPN registration' -Status 'Warning' -Details "Global catalog lookup failed: $($_.Exception.Message)"
            }
        }
    }
}

#endregion active directory and SPN

#region authentication methods

# Without Negotiate in IISAuthenticationMethods the client never asks for a Kerberos ticket,
# no matter how correct the ASA and the SPNs are.
try {
    foreach ($outlookAnywhere in @(Get-OutlookAnywhere -ErrorAction Stop)) {
        $target = [string]$outlookAnywhere.Server
        if (-not (Test-ServerInScope -TargetServer $target -Scope $Server)) { continue }

        $authenticationMethods = @($outlookAnywhere.IISAuthenticationMethods | ForEach-Object { $_.ToString() })

        if ($authenticationMethods -contains 'Negotiate') {
            $findings += New-ASAFinding -Category 'Authentication' -Target $target -Check 'Outlook Anywhere offers Negotiate' -Status 'Pass' -Details ('IISAuthenticationMethods: {0}' -f ($authenticationMethods -join ', '))
        }
        else {
            $findings += New-ASAFinding -Category 'Authentication' -Target $target -Check 'Outlook Anywhere offers Negotiate' -Status 'Fail' -Details ('IISAuthenticationMethods: {0}' -f ($authenticationMethods -join ', '))
        }
    }
}
catch {
    $findings += New-ASAFinding -Category 'Authentication' -Target 'Organization' -Check 'Outlook Anywhere offers Negotiate' -Status 'Warning' -Details "Get-OutlookAnywhere failed: $($_.Exception.Message)"
}

if (Get-Command -Name 'Get-MapiVirtualDirectory' -ErrorAction SilentlyContinue) {
    try {
        foreach ($mapiVirtualDirectory in @(Get-MapiVirtualDirectory -ErrorAction Stop)) {
            $target = [string]$mapiVirtualDirectory.Server
            if (-not (Test-ServerInScope -TargetServer $target -Scope $Server)) { continue }

            $authenticationMethods = @($mapiVirtualDirectory.IISAuthenticationMethods | ForEach-Object { $_.ToString() })

            if ($authenticationMethods -contains 'Negotiate') {
                $findings += New-ASAFinding -Category 'Authentication' -Target $target -Check 'MAPI virtual directory offers Negotiate' -Status 'Pass' -Details ('IISAuthenticationMethods: {0}' -f ($authenticationMethods -join ', '))
            }
            else {
                $findings += New-ASAFinding -Category 'Authentication' -Target $target -Check 'MAPI virtual directory offers Negotiate' -Status 'Fail' -Details ('IISAuthenticationMethods: {0}' -f ($authenticationMethods -join ', '))
            }
        }
    }
    catch {
        $findings += New-ASAFinding -Category 'Authentication' -Target 'Organization' -Check 'MAPI virtual directory offers Negotiate' -Status 'Warning' -Details "Get-MapiVirtualDirectory failed: $($_.Exception.Message)"
    }
}

#endregion authentication methods

#region output

foreach ($finding in $findings) {
    Write-ASAFindingToHost -Finding $finding
}

$failCount = @($findings | Where-Object { $_.Status -eq 'Fail' }).Count
$warningCount = @($findings | Where-Object { $_.Status -eq 'Warning' }).Count
$passCount = @($findings | Where-Object { $_.Status -eq 'Pass' }).Count

Write-Host ''
Write-Host ('Result: {0} pass, {1} warning, {2} fail' -f $passCount, $warningCount, $failCount) -ForegroundColor White

if ($failCount -gt 0) {
    Write-Host 'The ASA configuration is not valid, clients will fall back to NTLM or fail to authenticate.' -ForegroundColor Red
}
elseif ($warningCount -gt 0) {
    Write-Host 'The ASA configuration works but needs attention.' -ForegroundColor Yellow
}
else {
    Write-Host 'The ASA configuration is valid.' -ForegroundColor Green
}

if ($ReportPath) {
    try {
        if ([System.IO.Path]::GetExtension($ReportPath) -eq '.csv') {
            $findings | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        }
        else {
            $head = @'
<style>
body { background-color:white; font-family:Calibri; font-size:12pt; }
th { border-bottom:1px solid black; background-color:#00004d; color:white; text-align:left; }
td { color:black; text-align:left; }
table, tr, td, th { padding:2px; margin:0px }
table { margin-left:50px; }
h1 { text-align:center; color:#00004d; }
</style>
'@
            $findings |
                ConvertTo-Html -Head $head -Title 'Exchange ASA validation' -PreContent ('<h1>Exchange ASA validation</h1><p>Generated {0}</p>' -f $scriptStart) |
                Out-File -FilePath $ReportPath -Encoding UTF8
        }

        Write-Host ('Report written to {0}' -f $ReportPath) -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Writing the report failed: $($_.Exception.Message)"
    }
}

Write-Host ''

# The findings go to the pipeline so the script can be used inside a larger health check.
$findings

#endregion output
