# ExchangeScripts

PowerShell scripts for Exchange Server assessments and day to day operations.
Everything here is read only unless the script itself says otherwise, and is written for Windows
PowerShell 5.1 inside the Exchange Management Shell.

## Scripts

### Test-ExchangeASAConfiguration.ps1

Validates the Alternate Service Account (ASA) that Exchange uses for Kerberos authentication on
load balanced namespaces, and reports every finding as a pass, warning or failure.

Kerberos for Exchange breaks in a handful of well known ways, and none of them announce
themselves: clients quietly fall back to NTLM, or a single server in the farm starts failing
authentication. The script checks the whole chain:

| Area | Check |
| --- | --- |
| Credential | An ASA credential is deployed on every Client Access service |
| Credential | The deployed credential is not older than the rotation threshold |
| Credential | The deployed account matches `-ExpectedASAAccount` when given |
| Consistency | Every server uses the same ASA account |
| Consistency | Every server carries the same credential generation, so the last password roll reached all of them |
| Active Directory | The ASA account exists and is enabled |
| Active Directory | `pwdLastSet` is not newer than the credential deployed on the servers |
| SPN | Every client namespace has a matching `http/` SPN on the ASA account |
| SPN | None of those SPNs is registered on a second account |
| Authentication | Outlook Anywhere and the MAPI virtual directories offer Negotiate |

Namespaces are collected from Outlook Anywhere, the MAPI, EWS and OAB virtual directories and the
Autodiscover service URI, so the SPN check follows the configuration instead of a hardcoded list.

An organization where no server has an ASA credential at all is not treated as broken. Kerberos is
simply not in use there and clients authenticate with NTLM, which is a valid configuration, so the
report says so in one line instead of failing every check. That only applies when every server
answered conclusively: if the state of one server could not be read, absence was never established
and the findings stay failures.

Active Directory is queried through `System.DirectoryServices`, so RSAT and the ActiveDirectory
module are not required. SPN duplicates are searched forest wide through the global catalog.

The ASA credential itself lives in the registry of every server, so the script reads that subtree
on each of them. That read goes local for the server running the session, then over WinRM, and
only as a last resort through the RemoteRegistry service. `-RegistryAccess Remoting` keeps it off
RemoteRegistry entirely, `RemoteRegistry` forces the old path and `None` skips the read. Every
server reports the same check with the same evidence, including how its registry was read, so two
servers in the same state never come back worded differently.

```powershell
# validate the whole organization
.\Test-ExchangeASAConfiguration.ps1

# validate against a known account and write an HTML report
.\Test-ExchangeASAConfiguration.ps1 -ExpectedASAAccount 'CONTOSO\EXCHANGE$' -ReportPath C:\Temp\asa.html

# two servers, stricter rotation threshold, no directory lookups
.\Test-ExchangeASAConfiguration.ps1 -Server EX01,EX02 -PasswordMaximumAgeDays 30 -SkipSPNCheck

# never touch the RemoteRegistry service, read over WinRM only
.\Test-ExchangeASAConfiguration.ps1 -RegistryAccess Remoting
```

The findings are also returned as objects (`Category`, `Target`, `Check`, `Status`, `Details`), so
the script can be dropped into a larger health check:

```powershell
$findings = .\Test-ExchangeASAConfiguration.ps1
$findings | Where-Object Status -eq 'Fail'
```

Requirements: Exchange Management Shell, view only Exchange permissions, directory read access for
the SPN checks. Tested against Exchange Server 2016 and 2019, with a `Get-ClientAccessServer`
fallback for Exchange 2013.

#### Troubleshooting

The ASA credential lives in the registry of every server, under
`HKLM\SYSTEM\CurrentControlSet\Services\MSExchangeServiceHost\ServiceAccounts`, and not in
Active Directory. Reading it therefore reaches out to each server in turn. When Exchange answers
with

```
Failed to read Alternate Service Account configuration data from the registry subtree
HLKM\SYSTEM\CurrentControlSet\Services\MSExchangeServiceHost\ServiceAccounts.
```

then one specific server could not be read. The script reports that server as a single finding and
keeps validating the rest, and its own registry read tells you which of these it is:

| Situation | What the script reports |
| --- | --- |
| Server offline, RemoteRegistry stopped, firewall blocking, or no rights | Fail, with the probe error |
| Registry reachable but the subtree is missing | Fail, no ASA credential was ever deployed there |
| Subtree present but empty | Fail, no ASA credential is deployed on this server. Exchange throws on an empty subtree instead of reporting an unset credential, so this is the normal state of a server that never received the credential |
| Subtree present and holding subkeys or values | Warning, the data is there so it points at rights on that subtree |

### Get-ExUrlInfo.ps1

Reads the virtual directory configuration of the active Exchange organization and builds an HTML
report. Dot source the file first, then call `Get-ExUrlInfo`.

### Install-ExchangeSEPrerequisites.ps1

Downloads and installs the Windows prerequisites for the Exchange Server Subscription Edition Mailbox
role on Windows Server 2025: the Windows features from the Microsoft prerequisites page, .NET
Framework 4.8.1, the Visual C++ 2012 and 2013 x64 redistributables, UCMA 4.0, IIS URL Rewrite 2.1
and the Remote Registry service on Automatic. This one runs before Exchange exists, so it needs a
plain elevated PowerShell session, not the Exchange Management Shell.

Three modes: `Download` on a machine with internet access, `Install` on the server from the copied
folder, or `DownloadAndInstall` on the same machine (the default). Every item ends as Downloaded,
AlreadyDownloaded, Installed, AlreadyInstalled, Set, Failed or Skipped, everything goes to a log file
in the download folder, and the script ends with a read only check of every requirement and a
verdict: `SERVER READY FOR EXCHANGE SETUP` (exit 0), `SERVER READY FOR EXCHANGE SETUP AFTER REBOOT`
(exit 3010) or `SERVER NOT READY FOR EXCHANGE SETUP` with the missing items listed (exit 1).

```powershell
# download on a machine with internet, copy the folder to the server afterwards
.\Install-ExchangeSEPrerequisites.ps1 -Mode Download -Path D:\ExchangePrereqs

# install from that folder on the server and reboot when done
.\Install-ExchangeSEPrerequisites.ps1 -Mode Install -Path D:\ExchangePrereqs -Restart

# readiness check only, nothing is changed
.\Install-ExchangeSEPrerequisites.ps1 -Mode Install -WhatIf
```

Follows the SE RTM prerequisites. Exchange SE CU1 drops UCMA and moves to the Visual C++ 2022
runtime, so check the Microsoft page again once CU1 is released.

### Copy-ExchangeServerConfig.ps1

Copies the client access and transport configuration of an existing Exchange server to a newly
installed one. Written for the Exchange SE migration where the existing server is upgraded in place
and new servers join the organization for a new DAG. Reads the source, compares with the target and
applies only what differs, so a second run reports "in sync" and doubles as a drift check. Every
decision goes to a CSV in the output folder next to a JSON snapshot of both servers. Run with
`-WhatIf` first. Full documentation in [Copy-ExchangeServerConfig.md](Copy-ExchangeServerConfig.md).

```powershell
.\Copy-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02 -WhatIf
```

### Export-ExchangeServerConfigScript.ps1

The read only alternative to the script above. Reads the same configuration from the source server
and writes a plain PowerShell script of filled in Set commands, one per object, without functions or
logic. Review or trim the generated file and run it in the Exchange Management Shell against the new
server. Every command carries `-WhatIf:$WhatIf`, so one variable at the top turns the whole file into
a dry run. See [examples/Set-ExchangeServerConfig_from_EX01_example.ps1](examples/Set-ExchangeServerConfig_from_EX01_example.ps1)
for what the output looks like.

```powershell
.\Export-ExchangeServerConfigScript.ps1 -SourceServer EX01 -TargetServer EX02 -OutputPath C:\Temp\Set-EX02.ps1
```

### Compare-ExchangeServerConfig.ps1

Read only comparison of two Exchange servers over exactly the parameters the two scripts above can
set: certificates, virtual directories, Outlook Anywhere, transport, POP and IMAP, receive and send
connectors, config files, IIS bindings and more. Use it to prove the new server matches the source
and later as a drift check between DAG members. Writes `compare.csv` and `compare.html` and exits
with the number of differences, so it can gate a change window.

```powershell
.\Compare-ExchangeServerConfig.ps1 -SourceServer EX01 -TargetServer EX02
```

### Install-ExchangeSEServerIsolated.ps1

Quick and dirty runbook for installing an Exchange SE Mailbox server, the first one of a new
organization or an extra one in an existing organization, without it taking part in client access or
mail transport before it is configured. Three steps: `Setup`
runs Exchange Setup unattended with `/DoNotStartTransport`, `Isolate` puts every server component in
Inactive (maintenance mode, survives the reboot), points the Autodiscover SCP at the shared namespace
and stops transport, `Release` sets everything back to Active once the configuration is done.
Variables for the Setup path, the organization name (new organization only) and the Autodiscover
URI at the top of the file.

```powershell
.\Install-ExchangeSEServerIsolated.ps1 -Step Setup
.\Install-ExchangeSEServerIsolated.ps1 -Step Isolate
# reboot, configure with Copy-ExchangeServerConfig.ps1, check with Compare-ExchangeServerConfig.ps1
.\Install-ExchangeSEServerIsolated.ps1 -Step Release
```

## Author

Jente Paredis, jentech consulting BV. jente@jentech.be
