
#region get owners
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force

$sites = Import-Csv -Path 'C:\temp\site-list.csv'
$sites = $sites | select-object -property url -unique


[System.Collections.ArrayList]$Array = @()

foreach ($site in $sites) {
    $siteUrl = $site.Url.Trim()
    Write-Host $($site.url)
    Connect-PnPOnline -Url $siteUrl -WarningAction Ignore -UseWebLogin
    $OwnersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction SilentlyContinue
    if ($OwnersGroup) {
        Write-Host -ForegroundColor Green "Found!"
        $Owners = $OwnersGroup | Get-PnpGroupMember
        foreach ($Owner in $Owners) {
            if ($Owner.Title -eq 'System Account') {
                continue
            }
            if (!$Owner.Email) {
                continue
            }
            $object = New-Object PSObject -Property @{
                Name = $Owner.Title
                Email = $Owner.Email
                Group = $OwnersGroup.Title
                SiteUrl = $siteUrl
            }
            $Array.Add($object) | Out-Null
        }
    }
    else {
        Write-Host -ForegroundColor Yellow "No default owners group found! :("
        $object = New-Object PSObject -Property @{
            Name = 'None found'
            Email = 'None found'
            Group = 'None found'
            SiteUrl = $siteUrl
        }
        $Array.Add($object) | Out-Null
    }
    Write-Host
}

$Array | Select-Object SiteUrl, Group, Name, Email | Export-Csv -Path 'C:\temp\site-list-owners.csv' -NoTypeInformation
#endregion

#region set default owners group
$siteUrl = 'https://tenant.sharepoint.com/sites/site1/'

Connect-PnPOnline -Url $siteUrl -WarningAction Ignore -UseWebLogin

Get-PnPGroup -AssociatedOwnerGroup -ErrorAction SilentlyContinue
Get-PnPGroup -AssociatedMemberGroup -ErrorAction SilentlyContinue
Get-PnPGroup -AssociatedVisitorGroup -ErrorAction SilentlyContinue


Set-PnPGroup -Identity 'xxx Owners' -SetAssociatedGroup Owners
Set-PnPGroup -Identity 'xxx Members' -SetAssociatedGroup Members
Set-PnPGroup -Identity 'xxx Visitors' -SetAssociatedGroup Visitors


Set-PnPGroup -AssociatedOwnerGroup

$OwnersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction SilentlyContinue
#endregion