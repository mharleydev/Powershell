<#
    .DESCRIPTION
    Use this script to create the defaults.json file. Edit these values then run it.

    .OUTPUTS defaults.json
    This file is used to initialize the UID process.
#>
[System.Collections.ArrayList]$Array = @()
$object = New-Object PSObject -Property @{
    Name = 'ARSServerName'
    Value = 'servername.docmain.local'
    Description = 'The name of the Active Roles server'
}
$Array.Add($object) | Out-Null

$object = New-Object PSObject -Property @{
    Name = 'IDMSiteUrl'
    Value = 'https://Contoso.com/sites/sitename'
    Description = 'Site URL for the Identity Warehouse SharePoint Site'
}
$Array.Add($object) | Out-Null

$object = New-Object PSObject -Property @{
    Name = 'UIDRecordsListName'
    Value = 'uid_records'
    Description = 'Name of the list that contains the teammate records'
}
$Array.Add($object) | Out-Null

$object = New-Object PSObject -Property @{
    Name = 'UIDMasterListName'
    Value = 'uid_master_list'
    Description = 'Name of teh list that contains the master list of uid numbers'
}
$Array.Add($object) | Out-Null

$Array | ConvertTo-Json | Out-File -Path 'defaults.json'
