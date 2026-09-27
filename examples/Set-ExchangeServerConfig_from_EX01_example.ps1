# EXAMPLE OUTPUT of Export-ExchangeServerConfigScript.ps1, generated against a fictional contoso lab (EX01 to EX02).
# Server names, URLs, IP addresses and thumbprints are made up. Do not run this file, generate your own.

<#
    Exchange server configuration of EX01 (ex01.contoso.local)
    Exported 2026-09-25 20:49 with Export-ExchangeServerConfigScript.ps1 (jentech consulting)
    Source build: Version 15.2 (Build 2562.17)

    Plain list of Exchange commands that reproduce the client access and transport configuration of the
    source on the server named in $TargetServer. Run it in the Exchange Management Shell with Organization
    Management rights. Review it first, remove what does not apply, and start with $WhatIf = $true.

    Not included: DAG, databases, organization wide objects, config file customisations (web.config,
    EdgeTransport.exe.config, OWA logon files), DNS and load balancer changes.
#>

$TargetServer = 'EX02'
$TargetFqdn   = 'ex02.contoso.local'
$WhatIf       = $true      # set to $false to apply
$ErrorActionPreference = 'Continue'

if (-not (Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue)) { throw 'Run this script in the Exchange Management Shell.' }
Get-ExchangeServer -Identity $TargetServer -ErrorAction Stop | Out-Null

# ------------------------------------------------------------------------------------------------
# Certificates: export from the source, import on the target, enable the same services
# ------------------------------------------------------------------------------------------------
# Non exportable and self signed certificates are listed as comments. Enable-ExchangeCertificate -Force replaces
# the default SMTP certificate on the target without prompting, which is the intent.
$pfxPassword = Read-Host -AsSecureString -Prompt 'Password for the exported PFX files'
$pfxFolder   = Join-Path $env:TEMP 'ExchangeCertificateCopy'
New-Item -Path $pfxFolder -ItemType Directory -Force | Out-Null

# CN=mail.contoso.com (expires 25/09/2027 20:49:21, domains: mail.contoso.com, autodiscover.contoso.com)
$export = Export-ExchangeCertificate -Server 'EX01' -Thumbprint 'CERT1' -BinaryEncoded -Password $pfxPassword
[System.IO.File]::WriteAllBytes((Join-Path $pfxFolder 'CERT1.pfx'), $export.FileData)
Import-ExchangeCertificate -Server $TargetServer -FileData ([System.IO.File]::ReadAllBytes((Join-Path $pfxFolder 'CERT1.pfx'))) -Password $pfxPassword -PrivateKeyExportable $true -WhatIf:$WhatIf
Enable-ExchangeCertificate -Server $TargetServer -Thumbprint 'CERT1' -Services 'IMAP,POP,IIS,SMTP' -Force -Confirm:$false -WhatIf:$WhatIf

# CN=Microsoft Exchange Server Auth Certificate (expires 25/09/2029 20:49:21, domains: )
$export = Export-ExchangeCertificate -Server 'EX01' -Thumbprint 'AUTH1' -BinaryEncoded -Password $pfxPassword
[System.IO.File]::WriteAllBytes((Join-Path $pfxFolder 'AUTH1.pfx'), $export.FileData)
Import-ExchangeCertificate -Server $TargetServer -FileData ([System.IO.File]::ReadAllBytes((Join-Path $pfxFolder 'AUTH1.pfx'))) -Password $pfxPassword -PrivateKeyExportable $true -WhatIf:$WhatIf
# OAuth certificate: imported only. Setup flags it for SMTP itself and it must never become the default SMTP certificate.

# skipped, self signed: SELF1 CN=EX01 (services: IIS, SMTP)
# TODO private key not exportable, import by hand: NOEXP CN=legacy.contoso.com (services: )

# ------------------------------------------------------------------------------------------------
# Set-ExchangeServer (product key and static domain controllers are never exported)
# ------------------------------------------------------------------------------------------------
# Source edition: Enterprise. License the target with Set-ExchangeServer -ProductKey before going live.
Set-ExchangeServer -Identity $TargetServer `
    -ErrorReportingEnabled $false `
    -InternetWebProxy 'http://proxy.contoso.local:8080' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-ClientAccessService
# ------------------------------------------------------------------------------------------------
Set-ClientAccessService -Identity $TargetServer `
    -AutoDiscoverServiceInternalUri 'https://autodiscover.contoso.com/Autodiscover/Autodiscover.xml' `
    -AutoDiscoverSiteScope @('Default-First-Site-Name') `
    -WhatIf:$WhatIf

# TODO alternate service account is configured on the source. Run from the Exchange Scripts folder:
#   .\RollAlternateServiceAccountPassword.ps1 -ToSpecificServer $TargetServer -CopyFrom 'EX01'


# ------------------------------------------------------------------------------------------------
# Set-MailboxServer
# ------------------------------------------------------------------------------------------------
Set-MailboxServer -Identity $TargetServer `
    -AutoDatabaseMountDial 'GoodAvailability' `
    -DatabaseCopyAutoActivationPolicy 'Unrestricted' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-MalwareFilteringServer
# ------------------------------------------------------------------------------------------------
Set-MalwareFilteringServer -Identity $TargetServer `
    -BypassFiltering $false `
    -UpdateFrequency 60 `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-OwaVirtualDirectory
# ------------------------------------------------------------------------------------------------
Set-OwaVirtualDirectory -Identity "$TargetServer\owa (Default Web Site)" `
    -BasicAuthentication $true `
    -DefaultDomain 'contoso.com' `
    -ExternalAuthenticationMethods @('Fba') `
    -ExternalUrl 'https://mail.contoso.com/owa' `
    -FormsAuthentication $true `
    -InternalUrl 'https://mail.contoso.com/owa' `
    -LogonFormat 'UserName' `
    -WhatIf:$WhatIf

Set-OwaVirtualDirectory -Identity "$TargetServer\owa (Exchange Back End)" `
    -BasicAuthentication $false `
    -FormsAuthentication $false `
    -InternalUrl "https://$TargetFqdn:444/owa" `
    -LogonFormat 'FullDomain' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-EcpVirtualDirectory
# ------------------------------------------------------------------------------------------------
Set-EcpVirtualDirectory -Identity "$TargetServer\ecp (Default Web Site)" `
    -AdminEnabled $true `
    -BasicAuthentication $true `
    -ExternalUrl 'https://mail.contoso.com/ecp' `
    -FormsAuthentication $true `
    -InternalUrl 'https://mail.contoso.com/ecp' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-WebServicesVirtualDirectory
# ------------------------------------------------------------------------------------------------
Set-WebServicesVirtualDirectory -Identity "$TargetServer\EWS (Default Web Site)" `
    -ExternalUrl 'https://mail.contoso.com/EWS/Exchange.asmx' `
    -InternalNLBBypassUrl "https://$TargetFqdn/ews/exchange.asmx" `
    -InternalUrl 'https://mail.contoso.com/EWS/Exchange.asmx' `
    -MRSProxyEnabled $true `
    -OAuthAuthentication $true `
    -WindowsAuthentication $true `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-ActiveSyncVirtualDirectory
# ------------------------------------------------------------------------------------------------
Set-ActiveSyncVirtualDirectory -Identity "$TargetServer\Microsoft-Server-ActiveSync (Default Web Site)" `
    -BasicAuthEnabled $true `
    -ClientCertAuth 'Ignore' `
    -ExternalUrl 'https://mail.contoso.com/Microsoft-Server-ActiveSync' `
    -InternalUrl 'https://mail.contoso.com/Microsoft-Server-ActiveSync' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-OabVirtualDirectory
# ------------------------------------------------------------------------------------------------
Set-OabVirtualDirectory -Identity "$TargetServer\OAB (Default Web Site)" `
    -ExternalUrl 'https://mail.contoso.com/OAB' `
    -InternalUrl 'https://mail.contoso.com/OAB' `
    -OfflineAddressBooks @('\Default Offline Address Book') `
    -PollInterval 480 `
    -RequireSSL $true `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-MapiVirtualDirectory
# ------------------------------------------------------------------------------------------------
Set-MapiVirtualDirectory -Identity "$TargetServer\mapi (Default Web Site)" `
    -ExternalUrl 'https://mail.contoso.com/mapi' `
    -IISAuthenticationMethods @('Ntlm', 'OAuth', 'Negotiate') `
    -InternalUrl 'https://mail.contoso.com/mapi' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-PowerShellVirtualDirectory
# ------------------------------------------------------------------------------------------------
Set-PowerShellVirtualDirectory -Identity "$TargetServer\PowerShell (Default Web Site)" `
    -BasicAuthentication $false `
    -CertificateAuthentication $true `
    -InternalUrl "http://$TargetFqdn/powershell" `
    -RequireSSL $false `
    -WindowsAuthentication $false `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-AutodiscoverVirtualDirectory
# ------------------------------------------------------------------------------------------------
Set-AutodiscoverVirtualDirectory -Identity "$TargetServer\Autodiscover (Default Web Site)" `
    -BasicAuthentication $true `
    -OAuthAuthentication $true `
    -WindowsAuthentication $true `
    -WSSecurityAuthentication $true `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-OutlookAnywhere
# ------------------------------------------------------------------------------------------------
Set-OutlookAnywhere -Identity "$TargetServer\Rpc (Default Web Site)" `
    -ExternalClientAuthenticationMethod 'Negotiate' `
    -ExternalClientsRequireSsl $true `
    -ExternalHostname 'mail.contoso.com' `
    -IISAuthenticationMethods @('Basic', 'Ntlm', 'Negotiate') `
    -InternalClientAuthenticationMethod 'Ntlm' `
    -InternalClientsRequireSsl $true `
    -InternalHostname 'mail.contoso.com' `
    -SSLOffloading $false `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-TransportService (paths left out, use -IncludePaths)
# ------------------------------------------------------------------------------------------------
Set-TransportService -Identity $TargetServer `
    -InternalDNSServers @('10.0.0.1') `
    -MaxConcurrentMailboxDeliveries 20 `
    -MaxOutboundConnections 'Unlimited' `
    -MessageTrackingLogMaxAge '30.00:00:00' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-FrontendTransportService (paths left out, use -IncludePaths)
# ------------------------------------------------------------------------------------------------
Set-FrontendTransportService -Identity $TargetServer `
    -ReceiveProtocolLogMaxAge '30.00:00:00' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-MailboxTransportService (paths left out, use -IncludePaths)
# ------------------------------------------------------------------------------------------------
Set-MailboxTransportService -Identity $TargetServer `
    -MailboxDeliveryThrottlingLogMaxAge '7.00:00:00' `
    -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-ImapSettings and service startup type
# ------------------------------------------------------------------------------------------------
Set-ImapSettings -Server $TargetServer `
    -Banner "$TargetServer IMAP4 ready" `
    -ExternalConnectionSettings @('mail.contoso.com:993:SSL') `
    -LoginType 'SecureLogin' `
    -SSLBindings @('[::]:993', '0.0.0.0:993') `
    -UnencryptedOrTLSBindings @('10.0.0.12:143') `
    -X509CertificateName 'mail.contoso.com' `
    -WhatIf:$WhatIf

Set-Service -ComputerName $TargetFqdn -Name 'MSExchangeIMAP4' -StartupType Automatic -WhatIf:$WhatIf
if (-not $WhatIf) { Get-Service -ComputerName $TargetFqdn -Name 'MSExchangeIMAP4' | Start-Service }
Set-Service -ComputerName $TargetFqdn -Name 'MSExchangeIMAP4BE' -StartupType Automatic -WhatIf:$WhatIf
if (-not $WhatIf) { Get-Service -ComputerName $TargetFqdn -Name 'MSExchangeIMAP4BE' | Start-Service }


# ------------------------------------------------------------------------------------------------
# Set-PopSettings and service startup type
# ------------------------------------------------------------------------------------------------
Set-PopSettings -Server $TargetServer `
    -LoginType 'SecureLogin' `
    -SSLBindings @('0.0.0.0:995') `
    -WhatIf:$WhatIf

Set-Service -ComputerName $TargetFqdn -Name 'MSExchangePOP3' -StartupType Disabled -WhatIf:$WhatIf
Set-Service -ComputerName $TargetFqdn -Name 'MSExchangePOP3BE' -StartupType Disabled -WhatIf:$WhatIf


# ------------------------------------------------------------------------------------------------
# Set-EventLogLevel (levels raised above Lowest on the source)
# ------------------------------------------------------------------------------------------------
Set-EventLogLevel -Identity "$TargetServer\MSExchangeTransport\SmtpReceive" -Level Medium -WhatIf:$WhatIf

# ------------------------------------------------------------------------------------------------
# Set-SettingOverride (overrides scoped to the source server get the target added)
# ------------------------------------------------------------------------------------------------
# Override1: Enabled=true
Set-SettingOverride -Identity 'Override1' -Server @('EX01', $TargetServer) -Confirm:$false -WhatIf:$WhatIf

# ------------------------------------------------------------------------------------------------
# Receive connectors
# ------------------------------------------------------------------------------------------------
# Default connectors already exist on the target under the target name, only Set-ReceiveConnector is written
# for them. Custom connectors get New-ReceiveConnector first. Add-ADPermission lines reproduce the explicit
# ACEs of the source; most are granted by the permission groups anyway and adding them twice is harmless.

# --- Default Frontend EX01 (FrontendTransport) ---
Set-ReceiveConnector -Identity "$TargetServer\Default Frontend $TargetServer" `
    -AuthMechanism 'Tls, Integrated, BasicAuth, BasicAuthRequireTLS, ExchangeServer' `
    -Bindings @('0.0.0.0:25', '[::]:25') `
    -Fqdn "$TargetFqdn" `
    -MaxMessageSize '37748736B' `
    -PermissionGroups 'AnonymousUsers, ExchangeServers, ExchangeLegacyServers' `
    -RemoteIPRanges @('0.0.0.0-255.255.255.255') `
    -TlsCertificateName '<I>CN=Contoso CA<S>CN=mail.contoso.com' `
    -WhatIf:$WhatIf


# --- Relay Apps (FrontendTransport) ---
New-ReceiveConnector -Name 'Relay Apps' -Server $TargetServer -TransportRole FrontendTransport -Custom `
    -Bindings @('10.0.0.12:25') `
    -RemoteIPRanges @('10.0.1.0/24', '10.0.2.5') `
    -Confirm:$false -WhatIf:$WhatIf | Out-Null

Set-ReceiveConnector -Identity "$TargetServer\Relay Apps" `
    -AuthMechanism 'Tls, ExternalAuthoritative' `
    -Banner "220 $TargetServer relay" `
    -Bindings @('10.0.0.12:25') `
    -Fqdn 'relay.contoso.com' `
    -MaxMessageSize '104857600B' `
    -PermissionGroups 'AnonymousUsers' `
    -RemoteIPRanges @('10.0.1.0/24', '10.0.2.5') `
    -WhatIf:$WhatIf

Add-ADPermission -Identity "$TargetServer\Relay Apps" -User 'NT AUTHORITY\ANONYMOUS LOGON' -ExtendedRights @('ms-Exch-SMTP-Accept-Any-Recipient') -Confirm:$false -WhatIf:$WhatIf | Out-Null


# ------------------------------------------------------------------------------------------------
# Send connectors: LIVE ROUTING CHANGE, remove the leading # once the target is validated
# ------------------------------------------------------------------------------------------------
# Set-SendConnector -Identity 'Outbound to Internet' -SourceTransportServers @('EX01', $TargetServer) -Confirm:$false -WhatIf:$WhatIf

# ------------------------------------------------------------------------------------------------
# Transport agents on the source (manual: copy the assembly, then run the Install line)
# ------------------------------------------------------------------------------------------------
# Hub | Custom Agent | enabled True | priority 1 | D:\Agents\Contoso.dll
#   Install-TransportAgent -Name 'Custom Agent' -TransportService Hub -TransportAgentFactory 'Contoso.Agent.Factory' -AssemblyPath 'D:\Agents\Contoso.dll'; Enable-TransportAgent -Identity 'Custom Agent' -TransportService Hub; Set-TransportAgent -Identity 'Custom Agent' -TransportService Hub -Priority 1
# Hub | Transport Rule Agent | enabled True | priority 2 | C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\agents\Rule\Microsoft.Exchange.MessagingPolicies.TransportRuleAgent.dll

# ------------------------------------------------------------------------------------------------
# Hybrid configuration (manual: rerun the Hybrid Configuration Wizard and select the target as well)
# ------------------------------------------------------------------------------------------------
# SendingTransportServers currently: EX01. Add $TargetServer via the wizard or Set-HybridConfiguration -SendingTransportServers.
# ReceivingTransportServers currently: EX01. Add $TargetServer via the wizard or Set-HybridConfiguration -ReceivingTransportServers.

# ------------------------------------------------------------------------------------------------
# Config files to compare by hand (not scriptable, a cumulative update overwrites them)
# ------------------------------------------------------------------------------------------------
# Bin\EdgeTransport.exe.config
# Bin\MSExchangeMailboxReplication.exe.config
# TransportRoles\Shared\agents.config
# FrontEnd\HttpProxy\SharedWebConfig.config
# FrontEnd\HttpProxy\owa\web.config
# FrontEnd\HttpProxy\owa\auth\logon.aspx
# FrontEnd\HttpProxy\ews\web.config
# FrontEnd\HttpProxy\rpc\web.config
# FrontEnd\HttpProxy\mapi\web.config
# ClientAccess\SharedWebConfig.config
# ClientAccess\Owa\web.config
# ClientAccess\exchweb\ews\web.config

# ------------------------------------------------------------------------------------------------
# After the run
# ------------------------------------------------------------------------------------------------
# iisreset on the target, Restart-Service MSExchangeTransport and MSExchangeFrontEndTransport, restart the IMAP and
# POP services when they are used. Then Test-OutlookConnectivity, Test-OwaConnectivity, Test-ActiveSyncConnectivity
# before the target joins the load balancer pool.
