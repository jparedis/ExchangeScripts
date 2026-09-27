# Copy-ExchangeServerConfig

Copies the client access and transport configuration of an existing Exchange server to a newly installed one.
Written for the Exchange Server Subscription Edition migration: the existing server is upgraded in place, new
servers join the organization for a new DAG and must behave exactly like the existing one for clients and mail
flow.

## What it does

The script reads the configuration of the source server, compares it with the target server and applies only
what differs. Every decision is written to `changes.csv` in the output folder, together with a JSON snapshot of
both servers taken before any change. A second run reports "in sync" for everything that was applied, so it
doubles as a drift check between the servers of the DAG.

| Area | Cmdlets | Behaviour |
| --- | --- | --- |
| Certificates | Export, Import and Enable-ExchangeCertificate | Copies every certificate that is not self signed, plus the OAuth certificate from Get-AuthConfig, and enables the same services. Non exportable keys are reported. |
| ExchangeServer | Set-ExchangeServer | Everything except the product key and the static domain controller pins. Reports an unlicensed target. |
| ClientAccess | Set-ClientAccessService | Autodiscover internal URI and site scope. Reports an alternate service account on the source. |
| VirtualDirectories | Set-Owa, Ecp, WebServices, ActiveSync, Oab, Mapi, PowerShell and AutodiscoverVirtualDirectory | URLs, authentication and every other settable property, frontend and backend. |
| OutlookAnywhere | Set-OutlookAnywhere | Hostnames, authentication methods and SSL settings. |
| Transport | Set-TransportService, Set-FrontendTransportService, Set-MailboxTransportService | DNS servers, limits, log retention, message tracking and so on. Paths only with `-IncludePaths`. |
| PopImap | Set-ImapSettings, Set-PopSettings, Set-Service | Protocol settings plus the startup type of the four IMAP and POP services. |
| MailboxServer | Set-MailboxServer | Activation policy, maximum active databases, assistant schedules and the rest. |
| Malware | Set-MalwareFilteringServer | Update settings and bypass state. |
| EventLogLevels | Set-EventLogLevel | Diagnostic levels raised above Lowest on the source. |
| SettingOverrides | Set-SettingOverride | Adds the target to overrides that are scoped to the source server. |
| ReceiveConnectors | New-ReceiveConnector, Set-ReceiveConnector, Add-ADPermission | Default connectors are matched by name and converged, custom connectors are created, explicit ACEs such as anonymous relay are copied. |
| SendConnectors | Set-SendConnector | Adds the target as source transport server. Live routing change, always last. |
| TransportAgents | Get-TransportAgent | Report only, prints the Install-TransportAgent command for agents missing on the target. |
| ConfigFiles | file hashes over the admin share | Report only, saves both versions and a line diff of every differing web.config, EdgeTransport.exe.config, SharedWebConfig.config, agents.config and the OWA logon folder and theme folders. |
| Iis | WebAdministration over WinRM | Root redirect of the Default Web Site and its SSL flags are applied, site bindings are reported. |
| Hybrid | Get-HybridConfiguration | Report only, tells you when the source is listed as sending, receiving or edge transport server. |

## How settings are compared

For every Set cmdlet the script reads the parameter list from the cmdlet metadata and intersects it with the
properties of the source object. That keeps the script complete when a cumulative update adds parameters. Values
are normalised before comparison, so the display format of a size ("35 MB (36,700,160 bytes)") becomes the byte
count the cmdlet accepts. When the bulk Set call fails, for example on a deprecated parameter, the script retries
property by property and records exactly which one failed.

Server specific strings are translated: the source FQDN and NetBIOS name inside a value become those of the
target. This covers connector names such as `Default Frontend EX01`, banners, `InternalNLBBypassUrl` and SPN
lists. Disable it with `-NoServerNameSubstitution`.

Bindings on a specific IP address (receive connectors, IMAP, POP) are translated with `-IpAddressMap`. Without a
mapping such a binding is reported and skipped, wildcard bindings need no mapping.

## Usage

Run from the Exchange Management Shell with Organization Management rights. The config file, IIS and service
checks also need access to the admin share and WinRM of both servers.

```powershell
# 1. Report only
.\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -WhatIf

# 2. Apply everything except the send connectors, keep a PFX copy of the certificates
.\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -SkipArea SendConnectors `
    -CertificatePassword (Read-Host -AsSecureString 'PFX password')

# 3. Relay connector bound to a specific address
.\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -Area ReceiveConnectors `
    -IpAddressMap @{ '10.0.0.11' = '10.0.0.12' }

# 4. After testing: add the target to the send connectors
.\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -Area SendConnectors

# 5. Drift check afterwards, should report everything in sync
.\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -WhatIf
```

Repeat step 2 to 4 for the second new server, or run the script from the first new server once it is confirmed
good so the second one is a copy of a copy that was already validated.

## Output folder

`ExchangeServerConfigCopy_<source>_to_<target>_<timestamp>` in the current directory, or `-OutputPath`:

- `changes.csv`: every decision with Area, Identity, Property, SourceValue, TargetValueBefore, Status and Message.
  Status is Applied, WhatIf, Failed, ReviewRequired, Skipped or Info.
- `<source>-config.json` and `<target>-config-before.json`: snapshot of the objects that were read, usable as
  rollback reference.
- `Certificates\<thumbprint>.pfx`: only when `-CertificatePassword` was given.
- `ConfigFiles\...`: both versions plus `.diff.txt` for every differing file.
- `transcript.log`: full console output.

## After the run

The script prints these hints based on what it applied:

- `iisreset` on the target after virtual directory, Outlook Anywhere or certificate changes.
- Restart of MSExchangeTransport and MSExchangeFrontEndTransport after connector changes.
- Restart of the IMAP or POP services after their settings changed.

Then, outside the scope of the script: build the DAG and the database copies, add the target to the load balancer
pool once `Test-OutlookConnectivity`, `Test-OwaConnectivity` and `Test-ActiveSyncConnectivity` pass, and apply
the config file customisations from the diffs by hand and document them, because the next cumulative update
overwrites those files again.

## Alternative: Export-ExchangeServerConfigScript.ps1

Same reading logic, different delivery. `Export-ExchangeServerConfigScript.ps1 -SourceServer EX01 -TargetServer EX02`
only reads the source and writes a plain PowerShell file without functions or logic: one Set command per
object with every parameter filled in, one parameter per line, sections with comments. Open it, trim it, set
`$WhatIf = $false` at the top and run it in the Exchange Management Shell against the target. Strings that
contain the source server name are written with `$TargetServer` or `$TargetFqdn`, so the same file serves the
second DAG member after changing two variables. Send connector lines are written as comments because they change
live routing. Use it when the change must be reviewable line by line or handed to someone else to execute; use
Copy-ExchangeServerConfig.ps1 when you want the comparison, the change report and the config file diffs.

A complete example of the generated output, produced against a fictional lab, is in
`examples/Set-ExchangeServerConfig_from_EX01_example.ps1`.

## Verification: Compare-ExchangeServerConfig.ps1

Read only check to run after either approach, and later as a drift check between the DAG members:

```powershell
.\Compare-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02
.\Compare-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -ShowEqual -OutputPath C:\Docs\EX02-asbuilt
```

It reads the same objects on both servers and compares every parameter the matching Set cmdlet accepts, with
the same normalisation and server name substitution as the copy script, so "Default Frontend EX01" equals
"Default Frontend EX02" and a size shown as "35 MB" equals its byte count. Every property becomes a row with
status Equal, Different, SourceOnly or TargetOnly. Certificates, explicit connector permissions, send connector
membership, setting overrides, service startup types, config file hashes and the IIS root redirect are
included. Output: differences on the console, `compare.csv` with all rows and `compare.html` with a coloured
table per area. `-ShowEqual` turns it into a full side by side as built of both servers. The exit code equals
the number of differences, so it can gate a change window.

## Not in scope

DAG creation, mailbox databases and database copies, organization wide objects (accepted domains, email address
policies, transport rules, OWA and mobile device mailbox policies), the alternate service account credential,
the product key, DNS records and load balancer configuration.
