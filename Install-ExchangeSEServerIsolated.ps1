#Requires -Version 5.1

<#
.SYNOPSIS
    Installs an Exchange SE Mailbox server, as the first server of a new organization or as an extra server in an
    existing one, and keeps it out of client access and mail transport until the configuration is done.

.DESCRIPTION
    Quick and dirty runbook in three steps. Adjust the variables at the top, then run each step in order.

      Setup     runs Exchange Setup unattended with /DoNotStartTransport so the transport services stay stopped.
                Set $OrganizationName for a new organization, Setup then also prepares the schema, AD and all
                domains, which needs an account in Schema Admins and Enterprise Admins. Leave it empty to add a
                server to an existing organization.
      Isolate   run right after Setup finishes, before the reboot. Puts every server component in Inactive
                (maintenance mode) so the server refuses SMTP and proxies no client traffic even after the
                services start on the next reboot, and points the Autodiscover SCP at the shared namespace so
                Outlook clients never pick this server directly.
      Release   run when the certificate, virtual directories and transport configuration are in place
                (Copy-ExchangeServerConfig.ps1 or the generated script from Export-ExchangeServerConfigScript.ps1,
                checked with Compare-ExchangeServerConfig.ps1). Sets the components back to Active and restarts
                transport.

    Isolate and Release load the local Exchange snapin, so run them on the new server itself in an elevated
    PowerShell session. Setup expects the Exchange ISO to be mounted.

    Order: Setup, Isolate, reboot, configure, Release, add to load balancer, add to DAG.

.PARAMETER Step
    Setup, Isolate or Release.

.EXAMPLE
    .\Install-ExchangeSEServerIsolated.ps1 -Step Setup
    .\Install-ExchangeSEServerIsolated.ps1 -Step Isolate
    Restart-Computer
    # configure the server, then
    .\Install-ExchangeSEServerIsolated.ps1 -Step Release

.NOTES
    jentech consulting
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Setup', 'Isolate', 'Release')]
    [string]$Step
)

# Variables, adjust before running
$SetupPath       = 'D:\Setup.exe'                              # mounted Exchange SE ISO
$TargetDir       = 'C:\Program Files\Microsoft\Exchange Server\V15'
$OrganizationName = ''                                          # new organization: set the name, existing: leave empty
$AutodiscoverUri = 'https://autodiscover.contoso.com/Autodiscover/Autodiscover.xml'   # the shared namespace
$Requester       = 'Maintenance'                               # must be the same for Inactive and Active

$ErrorActionPreference = 'Stop'
$ServerName = $env:COMPUTERNAME

function Import-ExchangeSnapin {
    if (-not (Get-Command Get-ServerComponentState -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
    }
}

switch ($Step) {

    'Setup' {
        # /InstallWindowsComponents adds the Windows features, /DoNotStartTransport leaves MSExchangeTransport and
        # MSExchangeFrontEndTransport stopped after Setup. /OrganizationName only for the first server of a new
        # organization, Setup then runs PrepareSchema, PrepareAD and PrepareAllDomains itself.
        $Arguments = @(
            '/Mode:Install',
            '/Roles:Mailbox',
            '/IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF',
            '/InstallWindowsComponents',
            '/DoNotStartTransport',
            "/TargetDir:`"$TargetDir`""
        )
        if ($OrganizationName) {
            $Arguments += "/OrganizationName:`"$OrganizationName`""
        }
        Write-Host "Running: $SetupPath $($Arguments -join ' ')"
        $Process = Start-Process -FilePath $SetupPath -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        if ($Process.ExitCode -ne 0) {
            throw "Setup ended with exit code $($Process.ExitCode). Check C:\ExchangeSetupLogs\ExchangeSetup.log"
        }
        Write-Host 'Setup done. Run -Step Isolate now, before the reboot.'
    }

    'Isolate' {
        Import-ExchangeSnapin

        # ServerWideOffline covers every component: HubTransport, FrontendTransport and all the client access
        # proxies (OWA, ECP, EWS, ActiveSync, OAB, MAPI, Autodiscover, PowerShell). Other servers stop routing
        # mail to this one and the HTTP proxy answers 503, which makes the load balancer health probe fail too.
        Set-ServerComponentState -Identity $ServerName -Component ServerWideOffline -State Inactive -Requester $Requester

        # Setup created the SCP with this server's own FQDN and the self signed certificate. Point it at the
        # shared namespace right away so Outlook in this AD site never contacts this server directly.
        Set-ClientAccessService -Identity $ServerName -AutoDiscoverServiceInternalUri $AutodiscoverUri

        # Keep transport stopped until Release, even if someone starts it by hand.
        Stop-Service MSExchangeTransport, MSExchangeFrontEndTransport -ErrorAction SilentlyContinue

        Get-ServerComponentState -Identity $ServerName | Where-Object { $_.State -ne 'Active' } |
            Format-Table Component, State -AutoSize
        Write-Host "$ServerName is isolated. Reboot, configure the server, then run -Step Release."
    }

    'Release' {
        Import-ExchangeSnapin

        Set-ServerComponentState -Identity $ServerName -Component ServerWideOffline -State Active -Requester $Requester

        Restart-Service MSExchangeTransport
        Restart-Service MSExchangeFrontEndTransport

        Get-ServerComponentState -Identity $ServerName | Where-Object { $_.State -ne 'Active' } |
            Format-Table Component, State -AutoSize
        Test-ServiceHealth -Server $ServerName | Format-Table Role, RequiredServicesRunning -AutoSize
        Write-Host "$ServerName is live. Add it to the load balancer and the DAG."
    }
}
