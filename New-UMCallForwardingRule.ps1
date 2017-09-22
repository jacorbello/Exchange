<#
.Synopsis
   New-UMCallForwardingRule.ps1 - Configures UM call forwarding
.DESCRIPTION
   This Cmdlet configures UM call forwarding to another UM Enabled mailbox.
.PARAMETER RuleName
    Name for the new UM Call Answering Rule
.PARAMETER Mailbox
    Alias of the user for the rule to apply to
.PARAMETER ForwardTo
    Alias of the user for the rule to foward the voicemail to
.PARAMETER TransferMessage
    String value that the auto attendant will read to the user when reaching voicemail
.PARAMETER OnlyWhenOOO
    Switch to configure rule to only apply when calendar shows user as Out of Office
.PARAMETER OnlyWhenBusy
    Switch to configure rule to only apply when calendar shows user as Busy
.EXAMPLE
   New-UMCallForwardingRule.ps1 -RuleName "ForwardToReceptionist" -Mailbox "TestBox1" -ForwardTo "RecBox1"
   Forwards calls for TestBox1 to RecBox1
.EXAMPLE 
   New-UMCallForwardingRule.ps1 -RuleName "ForwardToReceptionist" -Mailbox "TestBox1" -ForwardTo "RecBox1" -OnlyWhenBusy
   Forwards calls for TestBox1 to RecBox1 - Only when calendar shows TestBox1 as "Busy"
.NOTES
   Author - Jeremy Corbello
   V1.0 - 7/20/2017
#>

[CmdletBinding()]
param (  
        [Parameter( Mandatory=$true)]
        [string]$RuleName,
        
        [Parameter( Mandatory=$true)]
        [string]$Mailbox,

        [Parameter( Mandatory=$true)]
        [string]$ForwardTo,

        [Parameter( Mandatory=$false)]
        [string]$TransferMessage = "The person you have tried to reach is unavailable",

        [Parameter( Mandatory=$false)]
        [switch]$OnlyWhenOOO,

        [Parameter( Mandatory=$false)]
        [switch]$OnlyWhenBusy

    )

if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1) {
    if (-not (Get-PSSession).ConfigurationName -eq "Microsoft.Exchange") {
	    . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	    Connect-ExchangeServer -auto -AllowClobber
        Write-Host "Established Remote Exchange Session"
        } else {
        Write-Host "Exchange Session Already Established"
        }
    }
else {
    Write-Warning "Exchange Server management tools are not installed on this computer."
    EXIT
    }

$ForwardToDN = (Get-UMMailbox $ForwardTo).LegacyExchangeDN

if ($OnlyWhenOOO) {
    New-UmCallAnsweringRule -Name $RuleName -Mailbox $Mailbox -Priority 2 -KeyMappings "3,1,$($TransferMessage),,0,,0,,$($ForwardToDN)" -ScheduleStatus 0x8
    }
if ($OnlyWhenBusy) { 
    New-UmCallAnsweringRule -Name $RuleName -Mailbox $Mailbox -Priority 2 -KeyMappings "3,1,$($TransferMessage),,0,,0,,$($ForwardToDN)" -ScheduleStatus 0x4
    }
if (!$OnlyWhenOOO -AND !$OnlyWhenBusy) {
    New-UmCallAnsweringRule -Name $RuleName -Mailbox $Mailbox -Priority 2 -KeyMappings "3,1,$($TransferMessage),,0,,0,,$($ForwardToDN)"
    }