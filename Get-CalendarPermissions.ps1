<#
.Synopsis
   Get-CalendarPermissions - Queries all mailboxes to determine who the specified user has access for
.EXAMPLE
   .\Get-CalendarPermissions -User palacia -VerboseMode
.EXAMPLE
   .\Get-CalendarPermissions -User palacia -ReportMode -VerboseMode
.PARAMETER User
   Alias or UserName to query all mailboxes against
.NOTES
    Written by Jeremy Corbello
    V1.0 - 7/10/2017
#>

param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$User, 

        [Parameter(Mandatory=$false)]
        [switch]$ReportMode,

        [Parameter(Mandatory=$false)]
        [switch]$VerboseMode
        
        )

$OutputPath = "C:\temp\CalendarPermissions$(Get-Date -f 'MMddyy').csv"

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

$Report = @()
   
    $Mailboxes = Get-Mailbox -ResultSize 100 -Filter {RecipientTypeDetails -eq "UserMailbox"} 
    $i = 1

   ForEach ($Mailbox in $Mailboxes) 
     {
       $Calendar = $Mailbox.PrimarySmtpAddress.ToString() + ":\Calendar"
       #$Inbox = $Mailbox.PrimarySmtpAddress.ToString() + ":\Inbox"
       $Permissions = Get-MailboxFolderPermission -Identity $Calendar |  where-object {$_.User -like "$User" -and $_.AccessRights –notlike “None”} 
       Write-Progress -Activity "Scanning Mailboxes" -Status "$i out of $($Mailboxes.count)" -percentComplete ($i/$Mailboxes.count*100)
       $i++

      foreach ($Permission in $Permissions) 
         { 
  $permission | Add-Member -MemberType NoteProperty -Name "Calendar" -value $Mailbox.DisplayName
  $Report = $Report + $permission

        }
      }

if ($VerboseMode) {
    $Report | Select-Object Calendar,User,@{label="AccessRights";expression={$_.AccessRights}} | Out-GridView -PassThru
    }
if ($ReportMode) {
    $Report | Select-Object Calendar,User,@{label="AccessRights";expression={$_.AccessRights}} | Export-Csv -Path $OutputPath -NoTypeInformation
    }