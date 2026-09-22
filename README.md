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

Active Directory is queried through `System.DirectoryServices`, so RSAT and the ActiveDirectory
module are not required. SPN duplicates are searched forest wide through the global catalog.

```powershell
# validate the whole organization
.\Test-ExchangeASAConfiguration.ps1

# validate against a known account and write an HTML report
.\Test-ExchangeASAConfiguration.ps1 -ExpectedASAAccount 'CONTOSO\EXCHANGE$' -ReportPath C:\Temp\asa.html

# two servers, stricter rotation threshold, no directory lookups
.\Test-ExchangeASAConfiguration.ps1 -Server EX01,EX02 -PasswordMaximumAgeDays 30 -SkipSPNCheck
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
keeps validating the rest, and it probes the registry itself to tell you which of these it is:

| Situation | What the script reports |
| --- | --- |
| Server offline, RemoteRegistry stopped, firewall blocking, or no rights | Fail, with the probe error |
| Registry reachable but the subtree is missing | Fail, no ASA credential was ever deployed there |
| Registry and subtree both present | Warning, rights on that subtree rather than a missing credential |

### Get-ExUrlInfo.ps1

Reads the virtual directory configuration of the active Exchange organization and builds an HTML
report. Dot source the file first, then call `Get-ExUrlInfo`.

## Author

Jente Paredis, jentech consulting BV. jente@jentech.be
