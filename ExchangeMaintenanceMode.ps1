<#
    Place Exchange server into or pull out of maintenance mode
        Used for Cumulative Updates
    Author - Jeremy Corbello
    www.JeremyCorbello.com
#>

$EnterExit = Read-Host -Prompt 'Enter or Exit Maintenance Mode? [E] Enter or [X] Exit (default is "E")'

if ($EnterExit -like "x" -OR $EnterExit -like "exit") {
    $exServer = Read-Host -Prompt 'Server name to remove from Maintenance Mode (Not FQDN)'

    Write-Host "Updating state of component 'ServerWideOffline' to 'Active'"
    Set-ServerComponentState $exServer -Component ServerWideOffline -State Active -Requester Maintenance

    Write-Host "Resuming $exServer Cluster Node"
    Resume-ClusterNode -Name $exServer.tostring()

    Write-Host "Configuring Database Copy settings"
    Set-MailboxServer $exServer -DatabaseCopyAutoActivationPolicy Unrestricted
    Set-MailboxServer $exServer -DatabaseCopyActivationDisabledAndMoveNow $false

    Write-Host "Updating state of component 'HubTransport' to 'Active'"
    Set-ServerComponentState $exServer –Component HubTransport –State Active –Requester Maintenance

    Write-Host "Run the following to rebalance the Databases based on Activation Preference:" -ForegroundColor Black -BackgroundColor Yellow
    Write-Host "     cd $exscripts" -ForegroundColor Black -BackgroundColor Yellow
    Write-Host "     .\RedistributeActiveDatabases.ps1 -DagName <DAG NAME> -BalanceDBsByActivationPreference" -ForegroundColor Black -BackgroundColor Yellow

    Write-Host "$exServer is now out of maintenance mode" -ForegroundColor Black -BackgroundColor Green

} else {
    $exServer = Read-Host -Prompt 'Server name to place in Maintenance Mode (Not FQDN)'
    $exServerBackup = Read-Host -Prompt 'Secondary server name to make primary (Not FQDN)'
    $domain = (Get-ADDomain).dnsroot
    $exServer2 = $exServerBackup + $domain
    $dbCopies = Get-MailboxDatabaseCopyStatus -Server $exServer | Where {$_.Status -eq "Mounted"}
    $dbNames = $dbCopies.databasename

    Write-Host "Updating stae of component 'HubTransport' to 'Draining'"
    Set-ServerComponentState $exServer -Component HubTransport -State Draining –Requester Maintenance

    Write-Host "Redirecting messages from $exServer to $exServer2"
    Redirect-Message -Server $exServer -Target $exServer2

    Write-Host "Suspending $exServer Cluster Node"
    Suspend-ClusterNode -Name $exServer.tostring()

    Write-Host "Configuring Database Copy settings"
    Set-MailboxServer $exServer –DatabaseCopyActivationDisabledAndMoveNow $true
    Set-MailboxServer $exServer –DatabaseCopyAutoActivationPolicy Blocked

    foreach ($db in $dbNames) {Write-Host "Moving mailbox database from server $db to server $exServerBackup" -ForegroundColor Green -BackgroundColor Blue; Move-ActiveMailboxDatabase $db -ActivateOnServer $exServerBackup -SkipClientExperienceChecks -MountDialOverride:Lossless}
    Write-Host "Updating state of component 'ServerWideOffline' to 'Inactive'"
    Set-ServerComponentState $exServer –Component ServerWideOffline –State InActive –Requester Maintenance

    Write-Host "$exServer is now in maintenance mode" -ForegroundColor Black -BackgroundColor Green
}
