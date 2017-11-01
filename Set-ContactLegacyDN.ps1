<#
.SYNOPSIS
    Set-ContactLegacyDN.ps1 - Used to correct issue with changes in LegacyExchangeDN during Exchange upgrade

.DESCRIPTION 
    Please see this article for detailed explaination of the issue and solution that this script solves:
    https://www.jeremycorbello.com/index.php/knowledge-base/13-microsoft-exchange

.PARAMETER BounceBackAddress
    Address retrieved from NDR after attempting to send email to mail contact.
    For Example: IMCEAEX-_o=NT5_ou=00000000000000000000000000000000_cn=35B54X608F1BX34BB97580663XF890X9@domain.xyz

.PARAMETER MailContact
    Alias or Email Address of the mail contact that sending was attempted to.

.EXAMPLE
    .\Set-ContactLegacyDN.ps1 -BounceBackAddress "IMCEAEX-_o=NT5_ou=00000000000000000000000000000000_cn=35B54X608F1BX34BB97580663XF890X9@domain.xyz" -MailContact test-contact

.LINK
    https://www.jeremycorbello.com

.NOTES
    Written by: Jeremy Corbello

    * Website:	https://www.jeremycorbello.com
    * Twitter:	https://twitter.com/JeremyCorbello
    * LinkedIn:	https://www.linkedin.com/in/jacorbello/
    * Github:	https://github.com/jacorbello

    Change Log:
    V1.00 - 10/17/2017 - Initial version
#>

[CmdletBinding()]
param (
    [Parameter( Mandatory=$false)]
    [String]$BounceBackAddress,

    [Parameter( Mandatory=$false)]
    [String]$MailContact
)

#Add Exchange snapin if not already loaded in the PowerShell session
if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
{
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
    Write-Host "Established Remote Exchange Session"
}
else
{
    Write-Warning "Exchange Server management tools are not installed on this computer."
    EXIT
}

#Importing the Active Directory Module
Import-Module ActiveDirectory

Function Get-X500Address {
    param (
        [Parameter( Mandatory=$true)]
        [String]$Address
    )
    $Address = $Address.TrimStart("IMCEAEX-")
    $Address = $Address.Replace("_","/")
    $Address = $Address.Replace("+20"," ")
    $Address = $Address.Replace("+28","(")
    $Address = $Address.Replace("+29",")")
    $Address = $Address.Replace("+2E",".")
    $Address = $Address.Substring(0,$Address.LastIndexOf("@"))
    return $Address
}

Function Set-X500Address {
    param (
        [Parameter( Mandatory=$true)]
        [String]$Contact,

        [Parameter( Mandatory=$true)]
        [String]$NewX500
    )

    $new = (Get-ADObject -Identity $Contact -Properties ProxyAddresses).ProxyAddresses
    $new += "X500:"+"$NewX500"
    Set-ADObject -Identity $Contact -Replace @{proxyAddresses=$new}
}

Set-X500Address -Contact (Get-ADObject -Identity (Get-MailContact -Identity $MailContact).DistinguishedName -Properties ProxyAddresses).DistinguishedName -NewX500 (Get-X500Address -Address $BounceBackAddress)