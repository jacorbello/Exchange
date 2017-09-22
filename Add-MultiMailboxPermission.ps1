<#
.SYNOPSIS
Add-MultiMailboxPermission.ps1

.DESCRIPTION 
Adds mailbox permissions to multiple Exchange mailboxes.

.PARAMETER Mailbox
The name(s) of the mailbox(es) you want apply permissions to. Can be used in conjunction with $ImportFile parameter.

.PARAMETER ImportFile
Import list of mailbox(es) you want to apply permissions to. Can be used in conjunction with $Mailbox parameter.

.PARAMETER ApplicableUser
Alias of user to add as $Permission to list of mailboxes.

.PARAMETER Permission
Type of permission to add to list of mailboxes. If nothing specified, FullAccess is applied.
Valid values are as follows:
    FullAccess
    ExternalAccount
    DeleteItem
    ReadPermission
    ChangePermission
    ChangeOwner

.EXAMPLE
.\Add-MultiMailboxPermission.ps1 -Mailbox <user1>,<user2> -ApplicableUser jeremy.corbello -Permission FullAccess

.EXAMPLE
.\Add-MultiMailboxPermission.ps1 -ImportFile .\UserList.txt -ApplicableUser jeremy.corbello -Permission DeleteItem,ReadPermission

.LINK
https://www.jeremycorbello.com

.NOTES
Written by: Jeremy Corbello

* Website:	https://www.jeremycorbello.com
* Twitter:	https://twitter.com/JeremyCorbello
* LinkedIn:	https://www.linkedin.com/in/jacorbello/
* Github:	https://github.com/jacorbello

Change Log:
V1.00 - 9/7/2017 - Initial version
#>

[CmdletBinding()]
param(
	[Parameter( Position=0,Mandatory=$false)]
	[string[]]$Mailbox,

    [Parameter( Position=0,Mandatory=$false)]
    [string]$ImportFile,

    [Parameter( Mandatory=$true)]
    [string]$ApplicableUser,

    [Parameter( Mandatory=$false)]
    [string[]]$Permission = "FullAccess"
	)

$allUsers = @()
if ($ImportFile) {
    $allUsers += Get-Content -Path $ImportFile -Force
    }
if ($Mailbox) {
    $allUsers += $Mailbox
    }

$counter = 1
foreach ($user in $allUsers) {
    Write-Progress -Activity "Adding permission to $user" -PercentComplete ($counter/$allUsers.count*100) -Status "$counter out of $($allUsers.count)"
    Add-MailboxPermission -Identity $user -User $ApplicableUser -AccessRights $Permission
    $counter++
    }