Function Get-ExUrlInfo
    {


<#
.SYNOPSIS
   Reads the actual Virtual Directory configuration for the active Exchange Organization and crafts an HTML-based report

.DESCRIPTION
   This function is written to quickly gain an overview of all Virtual Directories within Exchange Server. 
   I Use it all the time when performing assessments or drafting roadmaps for Exchange Server Migrations.
   This script is tested with Exchange Server 2010, Exchange Server 2016 and Exchange Server 2019

.EXAMPLE
   Get-ExUrlInfo



.NOTES
   Author: Jente Paredis - jente@jentech.be

   Credits: heavily inspired on Get-VirDirInfo.ps1 from Michael Van Horenbeeck (https://github.com/enptmps/MessagingDiscovery/blob/master/get-virdirinfo v1.7 (1).ps1), which I used many times before.
   At the time of script creation I did not have the sources available. This script will therefore not be an exact copy with the exact same approach and possibilities, but it must be duely noted that I was not the first to come up with this idea.
   I tested it with all newest versions of Exchange Server, and added some additional information (IP addresses of both Exchange Servers and used namespaces, ...)

   Requirements:
   - this script is a function. It has to be dot-sourced first, and then called. 
   - It requires to be run within an Exchange Management Shell (or have the Exchange Commandlets otherwise available)

#>

# defining variables and arrays
$versioncontent = @()
$mailboxperserver = @()
$srvcontent = @()
$owacontent = @()
$ecpcontent = @()
$ewscontent = @()
$asynccontent = @()
$oabcontent = @()
$allex = "" 

# validating that the session the script is ran in has the proper commandlets loaded. If this is not the case, the scripts stops.
try {$null = get-excommand}
catch [System.Management.Automation.CommandNotFoundException] {Write-Warning "This script must be run in the Exchange Management Shell"; break;}

#nothing's wrong, continuing...
Write-Host "Exchange Powershell Commandlets are detected. Continuing..." -ForegroundColor "Green"


# crafting CSS to format the HTML File that will be the output of thie function.
        $head = @'
<style>
body { background-color:white; font-family:Calibri; font-size:12pt; }
th { border-bottom:1px solid black; background-color:#00004d; color:white; text-align:left; }
td { color: black;  text-align:left;}
table, tr, td, th { padding: 2px; margin: 0px }
table { margin-left:50px; }
h1 {text-align: center; color:#00004d;}
h2 {color:#00004d;}

</style>

'@


# retrieving all exchange servers performing client access activities
$all2010exchangeservers = Get-ExchangeServer |Where-Object {(($_.AdminDisplayVersion).Major -eq "14") -and ($_.ServerRole -eq "ClientAccess")}
$all2013exchangeservers = Get-ExchangeServer |Where-Object {(($_.AdminDisplayVersion).Major -eq "15") -and (($_.AdminDisplayVersion).Minor -eq "0") -and ($_.ServerRole -eq "ClientAccess")}
$all2016exchangeservers = Get-ExchangeServer |Where-Object {(($_.AdminDisplayVersion).Major -eq "15") -and (($_.AdminDisplayVersion).Minor -eq "1") }
$all2019exchangeservers = Get-ExchangeServer |Where-Object {(($_.AdminDisplayVersion).Major -eq "15") -and (($_.AdminDisplayVersion).Minor -eq "2") }

$allex += $all2010exchangeservers
$allex += $all2013exchangeservers
$allex += $all2016exchangeservers
$allex += $all2019exchangeservers

 $count2010 = $all2010exchangeservers.Count
 $count2013 = $all2013exchangeservers.Count
 $count2016 = $all2016exchangeservers.Count
 $count2019 = $all2019exchangeservers.Count

  

    foreach ($item in $allex)
{



$mailboxperserver += @{Name=$item.Name;Count=(Get-Mailbox -Server $item.Name -ResultSize "Unlimited").Count}
$mailboxinfo = $mailboxperserver | ForEach-Object { new-object PSObject -Property $_} |Select Name,Count


$versioncontent = @{EX2010=$count2010;EX2013=$count2013;EX2016=$count2016;EX2019=$count2019}
$versioninfo = $versioncontent |  ForEach-Object { new-object PSObject -Property $_}

    $major = $item.AdminDisplayversion.Major
    $minor = $item.AdminDisplayVersion.Minor
    $serverversion = "$major.$minor"
    $srvcontent+= @{Name=$item.Name; ServerRole=$item.ServerRole; Version=$serverversion}
    $srvinfo = $srvcontent |ForEach-Object { new-object PSObject -Property $_} |select-Object name,Version,ServerRole

    $owadata = Get-OwaVirtualDirectory -Server $item.Name
    $intauth = $owadata.InternalAuthenticationMethods |Out-String
    $extauth = $owadata.InternalAuthenticationMethods |Out-String
    $owacontent+= @{Name=$item.Name; InternalURL=$owadata.InternalURL; ExternalURL=$owadata.ExternalURL; InternalAuth=$intauth; ExternalAuth=$extauth}
    $owainfo = $owacontent |ForEach-Object { new-object PSObject -Property $_} |select-Object name,internalURL,ExternalURL,InternalAuth,ExternalAuth

    $ecpdata = Get-EcpVirtualDirectory -Server $item.Name
    $intauth = $ecpdata.InternalAuthenticationMethods |Out-String
    $extauth = $ecpdata.InternalAuthenticationMethods |Out-String
    $ecpcontent+= @{Name=$item.Name; InternalURL=$ecpdata.InternalURL; ExternalURL=$ecpdata.ExternalURL; InternalAuth=$intauth; ExternalAuth=$extauth}
    $ecpinfo = $ecpcontent |ForEach-Object { new-object PSObject -Property $_} |select-Object name,internalURL,ExternalURL,InternalAuth,ExternalAuth

    $ewsdata = Get-WebServicesVirtualDirectory -Server $item.Name
    $intauth = $ewsdata.InternalAuthenticationMethods |Out-String
    $extauth = $ewsdata.InternalAuthenticationMethods |Out-String
    $ewscontent+= @{Name=$item.Name; InternalURL=$ewsdata.InternalURL; ExternalURL=$ewsdata.ExternalURL; InternalAuth=$intauth; ExternalAuth=$extauth;MRSProxyEnabled=$ewsdata.MRSProxyEnabled}
    $ewsinfo = $ewscontent |ForEach-Object { new-object PSObject -Property $_} | select-Object name,internalURL,ExternalURL,InternalAuth,ExternalAuth,MRSProxyEnabled

    $asyncdata = Get-ActiveSyncVirtualDirectory -Server $item.Name
    $intauth = $asyncdata.InternalAuthenticationMethods |Out-String
    $extauth = $asyncdata.InternalAuthenticationMethods |Out-String
    $asynccontent+= @{Name=$item.Name; InternalURL=$asyncdata.InternalURL; ExternalURL=$asyncdata.ExternalURL; BasicAuthEnabled=$asyncdata.BasicAuthEnabled; WindowsAuthEnabled=$asyncdata.WindowsAuthEnabled;ClientCertAuth=$asyncdata.ClientCertAuth}
    $asyncinfo = $asynccontent |ForEach-Object { new-object PSObject -Property $_} |select-Object name,internalURL,ExternalURL,BasicAuthEnabled,WindowsAuthEnabled,ClientCertAuth

    $oabdata = Get-OabVirtualDirectory -Server $item.Name
    $intauth = $oabdata.InternalAuthenticationMethods |Out-String
    $extauth = $oabdata.InternalAuthenticationMethods |Out-String
    $oabcontent+= @{Name=$item.Name; InternalURL=$oabdata.InternalURL; ExternalURL=$oabdata.ExternalURL; InternalAuth=$intauth; ExternalAuth=$extauth}
    $oabinfo = $oabcontent |ForEach-Object { new-object PSObject -Property $_} |select-Object name,internalURL,ExternalURL,InternalAuth,ExternalAuth
    }

$obj0 = $versioninfo | ConvertTo-HTML -PreContent "<h2>Exchange Version Information</h2>" -Fragment |Out-String
$obj1 = $srvinfo | ConvertTo-HTML -PreContent "<h2>Server Information</h2>" -Fragment |Out-String
$obj2 = $owainfo | ConvertTo-HTML -PreContent "<h2>OWA Virtual Directory Configuration</h2>" -Fragment |Out-String
$obj3 = $ecpinfo | ConvertTo-HTML -PreContent "<h2>ECP Virtual Directory Configuration</h2>" -Fragment |Out-String
$obj4 = $ewsinfo | ConvertTo-HTML -PreContent "<h2>Web Services Virtual Directory Configuration</h2>" -Fragment |Out-String
$obj5 = $asyncinfo | ConvertTo-HTML -PreContent "<h2>ActiveSync Virtual Directory Configuration</h2>" -Fragment |Out-String
$obj6 = $oabinfo | ConvertTo-HTML -PreContent "<h2>Offline Address Book Virtual Directory Configuration</h2>" -Fragment |Out-String 
$obj7 = $mailboxinfo | ConvertTo-HTML -PreContent "<h2>Number of mailboxes per server</h2>" -Fragment |Out-String 
ConvertTo-Html -Head $head -PreContent "<h1>Get-ClientAccessConfig.ps1</h1>" -PostContent $obj7,$obj0,$obj1,$obj2,$obj3,$obj4,$obj5,$obj6 |Out-File C:\temp\test1.html
}
#