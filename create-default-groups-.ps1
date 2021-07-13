
#region create default groups
$siteUrl = 'https://tenant.sharepoint.com/sites/site1/'
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
Connect-PnPOnline -Url $siteUrl -WarningAction Ignore -UseWebLogin
$web = Get-PnPWeb
$groupNames = @('Owners', 'Members', 'Visitors')
foreach ($groupName in $groupNames) {
    New-PnPGroup -Title ($web.Title + ' ' + $groupName)
    Set-PnPGroup -Identity ($web.Title + ' ' + $groupName) -Owner ($web.Title + ' Owners') -SetAssociatedGroup $groupName
}

Set-PnPGroupPermissions -Identity ($web.Title + ' Owners') -AddRole 'Full Control'
Set-PnPGroupPermissions -Identity ($web.Title + ' Members') -AddRole 'Contribute'
Set-PnPGroupPermissions -Identity ($web.Title + ' Visitors') -AddRole 'Read'
#endregion

#region set associated group
$siteUrl = 'https://tenant.sharepoint.com/sites/site1/'
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
Connect-PnPOnline -Url $siteUrl -WarningAction Ignore -UseWebLogin

Set-PnPGroup -Identity 'xxx Owners' -SetAssociatedGroup 'Owners'


#endregion

