<#
.SYNOPSIS
Configure-NewExchangeServer.ps1

.DESCRIPTION 
Configures a new Exchange Server install to a baseline constant.

.PARAMETER Server
The name(s) of the server(s) you want configure

.PARAMETER KeepUMEnabled
Leaves UM Services active on the servers.

.PARAMETER ExchangeURL
Base URL for Exchange. i.e. mail.yourdomain.com

.PARAMETER ExternalSSL
Toggle SSL requirements for External Clients
Boolean value - Defaults to $true

.PARAMETER InternalSSL
Toggle SSL requirements for Internal Clients
Boolean value - Defaults to $true

.PARAMETER DefaultAuthenticationMethod
Update the Default Authentication Method.
String value - Defaults to Ntlm

.PARAMETER DAG
Specifies the name of the DAG to join the server(s) to.

.EXAMPLE
.\Configure-NewExchangeServer.ps1 -Server <ServerName>,<ServerName>

.LINK
https://www.jeremycorbello.com

.NOTES
Written by: Jeremy Corbello

* Website:	https://www.jeremycorbello.com
* Twitter:	https://twitter.com/JeremyCorbello
* LinkedIn:	https://www.linkedin.com/in/jacorbello/
* Github:	https://github.com/jacorbello

Change Log:
V1.00 - 8/31/2017 - Initial version.
V1.01 - 9/2/2017 - Added Receive Connector configuration.
V1.02 - 9/6/2017 - Removed client specific content.
#>




[CmdletBinding()]
param(
	[Parameter( Position=0,Mandatory=$true)]
	[string[]]$Server,

    [Parameter( Mandatory=$false)]
    [switch]$KeepUMEnabled,

    [Parameter( Mandatory=$false)]
    [string]$ExchangeURL = "mail.jeremycorbello.com",

    [Parameter( Mandatory=$false)]
    [bool]$ExternalSSL = $true,

    [Parameter( Mandatory=$false)]
    [bool]$InternalSSL = $true,

    [Parameter( Mandatory=$false)]
    [string]$DefaultAuthenticationMethod = "Ntlm",

    [Parameter( Mandatory=$false)]
    [string]$DAG = "DAG1"
	)

# Declaring Variables
$MimicSendConnectorsFromServer = "*Server Name*"
$ReceiveConnectorTemplateServer = "Server Name"
$OWALogonFormat = "UserName"
$OWADefaultDomain = "jeremycorbello.com"
$IMCertThumb = "A4XX1FXXE6XXF4XX26XXD1XX40XXAFXXDFXX2AXX"
$IMServerName = "skype.jeremycorbello.com"
$IMEnabled = $true
$IMType = "Ocs"
$ProductKey = "7WXV6-XXXXX-F4X67-XXXXX-F6XBY"
$EXCert = "\\Server Name\c$\root\mail-cert.pfx"
$EXCertPW = ConvertTo-SecureString "abcd1234" -AsPlainText -Force
$Expected = @{
    OA = [PSCustomObject]@{
        InternalHostname = "$ExchangeURL"
        ExternalHostname = "$ExchangeURL"
        }
    OWA = [PSCustomObject]@{
        InternalUrl = "https://$ExchangeURL/owa"
        ExternalUrl = "https://$ExchangeURL/owa"
        }
    ECP = [PSCustomObject]@{
        InternalUrl = "https://$ExchangeURL/ecp"
        ExternalUrl = "https://$ExchangeURL/ecp"
        }
    OAB = [PSCustomObject]@{
        InternalUrl = "https://$ExchangeURL/oab"
        ExternalUrl = "https://$ExchangeURL/oab"
        }
    EWS = [PSCustomObject]@{
        InternalUrl = "https://$ExchangeURL/EWS/Exchange.asmx"
        ExternalUrl = "https://$ExchangeURL/EWS/Exchange.asmx"
        }
    MAPI = [PSCustomObject]@{
        InternalUrl = "https://$ExchangeURL/mapi"
        ExternalUrl = "https://$ExchangeURL/mapi"
        }
    EAS = [PSCustomObject]@{
        InternalUrl = "https://$ExchangeURL/Microsoft-Server-ActiveSync"
        ExternalUrl = "https://$ExchangeURL/Microsoft-Server-ActiveSync"
        }
    AutoD = [PSCustomObject]@{
        AutoDiscoverServiceInternalUri = "https://$ExchangeURL/autodiscover/autodiscover.xml"
        }
}

# End Declaring Variables

Function GetURLs {
    [CmdletBinding()]
    param(
        [Parameter( Position=0,Mandatory=$true)]
        [string]$ServerName
        )
    $Results = @{
        OA = @()
        OWA = @()
        ECP = @()
        OAB = @()
        EWS = @()
        MAPI = @()
        EAS = @()
        AutoD = @()
        }
    $Results.OA = Get-OutlookAnyWhere -Server $ServerName -AdPropertiesOnly | Select-Object InternalHostName,ExternalHostName
    $Results.OWA = Get-OWAVirtualDirectory -Server $ServerName -AdPropertiesOnly | Select-Object InternalURL,ExternalURL
    $Results.ECP = Get-ECPVirtualDirectory -Server $ServerName -AdPropertiesOnly | Select-Object InternalURL,ExternalURL
    $Results.OAB = Get-OABVirtualDirectory -Server $ServerName -AdPropertiesOnly | Select-Object InternalURL,ExternalURL
    $Results.EWS = Get-WebServicesVirtualDirectory -Server $ServerName -AdPropertiesOnly | Select-Object InternalURL,ExternalURL
    $Results.MAPI = Get-MAPIVirtualDirectory -Server $ServerName -AdPropertiesOnly | Select-Object InternalURL,ExternalURL
    $Results.EAS = Get-ActiveSyncVirtualDirectory -Server $ServerName -AdPropertiesOnly | Select-Object InternalURL,ExternalURL
    $Results.AutoD = Get-ClientAccessService $ServerName | Select-Object AutoDiscoverServiceInternalUri

    return $Results
}

Function ReplaceURLs {
    [CmdletBinding()]
    param(
        [Parameter( Position=0,Mandatory=$true)]
        [string]$ServerName,

        [Parameter( Mandatory=$true)]
        [Hashtable]$URLs
        )
    # Comparing and replacing URLs
            Get-OutlookAnyWhere -Server $ServerName | Set-OutlookAnyWhere-Object -ExternalHostname $Expected.OA.ExternalHostname -ExternalClientAuthenticationMethod $DefaultAuthenticationMethod -ExternalClientsRequireSSL $ExternalSSL
        if (!($URLs.OA.InternalHostname.HostnameString.Equals($Expected.OA.InternalHostname)) -OR !($URLs.OA.InternalHostname.HostnameString)) {
            Get-OutlookAnyWhere -Server $ServerName | Set-OutlookAnyWhere-Object -InternalHostname $Expected.OA.InternalHostname -InternalClientAuthenticationMethod $DefaultAuthenticationMethod -InternalClientsRequireSSL $InternalSSL
            }
            Get-OwaVirtualDirectory -Server $ServerName | Set-OwaVirtualDirectory -ExternalUrl $Expected.OWA.ExternalUrl -LogonFormat $OWALogonFormat -DefaultDomain $OWADefaultDomain -InstantMessagingCertificateThumbprint $IMCertThumb -InstantMessagingType $IMType -InstantMessagingEnabled $IMEnabled -InstantMessagingServerName $IMServerName
        if (!($URLs.OWA.InternalUrl.AbsoluteUri.Equals($Expected.OWA.InternalUrl)) -OR !($URLs.OWA.InternalUrl.AbsoluteUri)) {
            Get-OwaVirtualDirectory -Server $ServerName | Set-OwaVirtualDirectory -InternalUrl $Expected.OWA.InternalUrl -LogonFormat $OWALogonFormat -DefaultDomain $OWADefaultDomain -InstantMessagingCertificateThumbprint $IMCertThumb -InstantMessagingType $IMType -InstantMessagingEnabled $IMEnabled -InstantMessagingServerName $IMServerName
            }
            Get-EcpVirtualDirectory -Server $ServerName | Set-EcpVirtualDirectory -ExternalUrl $Expected.ECP.ExternalUrl
        if (!($URLs.ECP.InternalUrl.AbsoluteUri.Equals($Expected.ECP.InternalUrl)) -OR !($URLs.ECP.InternalUrl.AbsoluteUri)) {
            Get-EcpVirtualDirectory -Server $ServerName | Set-EcpVirtualDirectory -InternalUrl $Expected.ECP.InternalUrl
            }
            Get-OabVirtualDirectory -Server $ServerName | Set-OabVirtualDirectory -ExternalUrl $Expected.OAB.ExternalUrl
        if (!($URLs.OAB.InternalUrl.AbsoluteUri.Equals($Expected.OAB.InternalUrl)) -OR !($URLs.OAB.InternalUrl.AbsoluteUri)) {
            Get-OabVirtualDirectory -Server $ServerName | Set-OabVirtualDirectory -InternalUrl $Expected.OAB.InternalUrl
            }
            Get-WebServicesVirtualDirectory -Server $ServerName | Set-WebServicesVirtualDirectory -ExternalUrl $Expected.EWS.ExternalUrl
        if (!($URLs.EWS.InternalUrl.AbsoluteUri.Equals($Expected.EWS.InternalUrl)) -OR !($URLs.EWS.InternalUrl.AbsoluteUri)) {
            Get-WebServicesVirtualDirectory -Server $ServerName | Set-WebServicesVirtualDirectory -InternalUrl $Expected.EWS.InternalUrl
            }
            Get-MapiVirtualDirectory -Server $ServerName | Set-MapiVirtualDirectory -ExternalUrl $Expected.MAPI.ExternalUrl
        if (!($URLs.MAPI.InternalUrl.AbsoluteUri.Equals($Expected.MAPI.InternalUrl)) -OR !($URLs.MAPI.InternalUrl.AbsoluteUri)) {
            Get-MapiVirtualDirectory -Server $ServerName | Set-MapiVirtualDirectory -InternalUrl $Expected.MAPI.InternalUrl
            }
            Get-ActiveSyncVirtualDirectory -Server $ServerName | Set-ActiveSyncVirtualDirectory -ExternalUrl $Expected.EAS.ExternalUrl
        if (!($URLs.EAS.InternalUrl.AbsoluteUri.EndsWith($Expected.EAS.InternalUrl)) -OR !($URLs.EAS.InternalUrl.AbsoluteUri)) {
            Get-ActiveSyncVirtualDirectory -Server $ServerName | Set-ActiveSyncVirtualDirectory -InternalUrl $Expected.EAS.InternalUrl
            }
        if (!($URLs.AutoD.AutoDiscoverServiceInternalUri.AbsoluteUri.Equals($Expected.AutoD.AutoDiscoverServiceInternalUri))) {
            Get-ClientAccessService -Identity $ServerName | Set-ClientAccessService -AutoDiscoverServiceInternalUri $Expected.AutoD.AutoDiscoverServiceInternalUri
            }
}

Function DisableUM {
    [CmdletBinding()]
    param(
        [Parameter( Position=0,Mandatory=$true)]
        [string]$ServerName
        )
    $Services = Get-Service -ComputerName $ServerName -Name "MSExchangeUM*"
    if ($Services.Status -ne "Stopped") {
        Get-Service -ComputerName $ServerName -Name "MSExchangeUM*" | Stop-Service
        Get-Service -ComputerName $ServerName -Name "MSExchangeUM*" | Set-Service -StartupType Disabled
        }
}

Function Add-SendConnector {
    [CmdletBinding()]
    param(
        [Parameter( Position=0,Mandatory=$true)]
        [string]$ServerName
        )
    $Senders = Get-SendConnector | Where-Object {$_.SourceTransportServers -like $MimicSendConnectorsFromServer}
    foreach ($Sender in $Senders) {
        $Members = $Sender.SourceTransportServers
        $Members += $ServerName
        Set-SendConnector -Identity $Sender -SourceTransportServers $Members
        }
}

Function Set-Certificate {
    [CmdletBinding()]
    param(
        [Parameter( Position=0,Mandatory=$true)]
        [string]$ServerName
        )
    Import-ExchangeCertificate -Server $ServerName -FileName $EXCert -Password $EXCertPW -FriendlyName "Exchange 2016 SSL Certificate" -PrivateKeyExportable $true -Confirm:$false
    $Thumpprint = (Get-ExchangeCertificate -Server $ServerName | Where-Object {$_.Subject -like "*$ExchangeURL*"}).Thumbprint
    Enable-ExchangeCertificate -Server $ServerName -Thumbprint $Thumpprint -Services IMAP,POP,IIS,SMTP -Confirm:$false
}

Function Add-ReceiveConnector {
    [CmdletBinding()]
    param(
        [Parameter( Position=0,Mandatory=$true)]
        [string]$ServerName
        )
    $connectors = Get-ReceiveConnector -Server $ReceiveConnectorTemplateServer | Where-Object {$_.name -notlike "*$ReceiveConnectorTemplateServer*"}
    foreach ($connector in $connectors) {
        foreach ($server in $ServerName) {
            if (!(Get-ReceiveConnector -Identity "$server\$($connector.name)" -ErrorAction silentlycontinue)) {
                New-ReceiveConnector -Server $server -Name $connector.Name -MaxHopCount $connector.MaxHopCount -MaxLocalHopCount $connector.MaxLocalHopCount -MaxMessageSize $connector.MaxMessageSize -RequireEHLODomain $connector.RequireEHLODomain -SuppressXAnonymousTls $connector.SuppressXAnonymousTls -MessageRateSource $connector.MessageRateSource -MessageRateLimit $connector.MessageRateLimit -BinaryMimeEnabled $connector.BinaryMimeEnabled -EightBitMimeEnabled $connector.EightBitMimeEnabled -TlsDomainCapabilities $connector.TlsDomainCapabilities -Bindings $connector.Bindings -Fqdn $connector.Fqdn -ConnectionInactivityTimeout $connector.ConnectionInactivityTimeout -MaxHeaderSize $connector.MaxHeaderSize -RejectSingleLabelRecipientDomains $connector.RejectSingleLabelRecipientDomains -DeliveryStatusNotificationEnabled $connector.DeliveryStatusNotificationEnabled -MaxProtocolErrors $connector.MaxProtocolErrors -SizeEnabled $connector.SizeEnabled -PermissionGroups $connector.PermissionGroups -Comment $connector.Comment -DomainSecureEnabled $connector.DomainSecureEnabled -MaxInboundConnectionPercentagePerSource $connector.MaxInboundConnectionPercentagePerSource -LongAddressesEnabled $connector.LongAddressesEnabled -AdvertiseClientSettings $connector.AdvertiseClientSettings -OrarEnabled $connector.OrarEnabled -Enabled $connector.Enabled -DefaultDomain $connector.DefaultDomain -MaxRecipientsPerMessage $connector.MaxRecipientsPerMessage -ServiceDiscoveryFqdn $connector.ServiceDiscoveryFqdn -EnhancedStatusCodesEnabled $connector.EnhancedStatusCodesEnabled -ConnectionTimeout $connector.ConnectionTimeout -MaxInboundConnectionPerSource $connector.MaxInboundConnectionPerSource -RejectReservedTopLevelRecipientDomains $connector.RejectReservedSecondLevelRecipientDomains -MaxLogonFailures $connector.MaxLogonFailures -AuthMechanism $connector.AuthMechanism -RequireTLS $connector.RequireTLS -RemoteIPRanges $connector.RemoteIPRanges -EnableAuthGSSAPI $connector.EnableAuthGSSAPI -MaxAcknowledgementDelay $connector.MaxAcknowledgementDelay -MaxInboundConnection $connector.MaxInboundConnection
                }
            }
        }
    }




# Add Exchange snapin if not already loaded in the PowerShell session
if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
{
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}
else
{
    Write-Warning "Exchange Server management tools are not installed on this computer."
    EXIT
}

############## PROCESS ###############

foreach ($i in $Server) {
    Write-Console "Gathering Virtual Directory URLs" -ForegroundColor Green
    $URLs = GetURLs -ServerName $i
    Write-Console "Updating Virtual Directory URLs" -ForegroundColor Green
    ReplaceURLs -ServerName $i -URLs $URLs
    if (!$KeepUMEnabled) {
        Write-Console "Disabling UM Services" -ForegroundColor Green
        DisableUM -ServerName $i
        }
    Write-Console "Applying Exchange Certificate" -ForegroundColor Green
    Set-Certificate -ServerName $i
    Write-Console "Adding $i to DAG $DAG" -ForegroundColor Green
    Add-DatabaseAvailabilityGroupServer -Identity $DAG -MailboxServer $i
    Write-Console "Adding $i to Send Connectors" -ForegroundColor Green
    Add-SendConnector -ServerName $i
    Write-Console "Activating Exchange Server" -ForegroundColor Green
    Set-ExchangeServer -Identity $i -ProductKey $ProductKey
    Write-Console "Configuring Receive Connectors on server $i" -ForegroundColor Green
    Add-ReceiveConnector -ServerName $i
    Write-Console "Restarting the Microsoft Exchange Information Store Service on $i" -ForegroundColor Green    
    Get-Service -ComputerName $i -Name "MSExchangeIS"
    Write-Console "Restarting OWA App Pool on server $i" -ForegroundColor Green
    Invoke-Command -ComputerName $i -ScriptBlock {Restart-WebAppPool -Name MSExchangeOWAAppPool}
    }
     