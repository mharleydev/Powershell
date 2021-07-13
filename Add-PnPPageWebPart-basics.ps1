#region init
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
$siteUrl = 'https://tenant.sharepoint.com/sites/site1/'
Connect-PnPOnline -Url $siteUrl -WarningAction Ignore -UseWebLogin
$siteRelativeUrl = '/sites/site1/'

#endregion

#region import csv
$ImportArray = Import-Csv -Path 'C:\list-fields.csv'

$List = Get-PnPList -Identity 'list-name'

foreach ($row in $ImportArray) {
    if ($row.'I/S' -eq 'V') {
        $FullName = 'Vacant'
    } 
    else {
        $FullName = ($row.'LAST NAME' + ', ' + $row.'FIRST NAME')
    }

    $ListItemValues = @{
        Title               = $row.title
        Status              = $row.status
        MissionArea         = $row.'MISSION AREA'
        MissionAgency       = $row.'MISSION AGENCY'
        Agency              = $row.'Agency Abbreviation'
        LastName            = $row.'LAST NAME'
        FirstName           = $row.'FIRST NAME'
        Full_x0020_Name     = $FullName
        PositionType        = $row.'CR/G'
        Appointment         = $row.'TYPE'
        Series              = $row.SERIES
        Location            = $row.Location
        State               = $row.STATE
    }
    $NewListItem = Add-PnPListItem -List $List -Values $ListItemValues
}

#endregion

#region create agency report pages
$AllocationsList = Get-PnPList -Identity 'list-name'
$Agencies = (Get-PnPListItem -List $AllocationsList -PageSize 500).FieldValues.Agency | Sort-Object -Unique
$items = @()
$itemId = 0
Write-Host "Creating Report Pages..."
foreach ($Agency in $Agencies) {
    Write-Host -ForegroundColor Cyan $($Agency)
    $itemId++
    $NewPageName = 'Agency Report for ' + $Agency
    $NewPageParams = @{
        Name                = $NewPageName
        HeaderLayoutType    = 'NoImage'
        CommentsEnabled     = $false
        Publish             = $true
    }
    $NewPage = Add-PnPPage @NewPageParams
}
#endregion

#region allocations report list views
$AllocationsList = Get-PnPList -Identity 'list-name'
$Agencies = (Get-PnPListItem -List $AllocationsList -PageSize 500).FieldValues.Agency | Sort-Object -Unique
$AllocationsReportList = Get-PnPList -Identity 'Allocation Reports'
$Fields = @(
    'TotalOnBoard',
    'TotalVacancies',
    'TotalAllocations',
    'Career',
    'Non_x002d_Career',
    'LTA'
)
Write-Host "Creating Allocations Report List Views..."
foreach ($Agency in $Agencies) {
    $ViewName = 'Agency: ' + $Agency
    Write-Host -ForegroundColor Cyan $($ViewName)
    $Query = "<OrderBy><FieldRef Name='Title' /></OrderBy><Where><And><Eq><FieldRef Name='Title' /><Value Type='Text'>$Agency</Value></Eq><Eq><FieldRef Name='Status' /><Value Type='Text'>Current</Value></Eq></And></Where>"

    $NewViewParams = @{
        List        = $AllocationsReportList
        Title       = $ViewName
        Fields      = $Fields
        Query       = $Query
        Paged       = $true 
        RowLimit    = 100
    }
    $NewView = Add-PnPView @NewViewParams
    $UpdatedView = Set-PnPView -List $AllocationsReportList -Identity $NewView.Title -Values @{'TabularView' = $false}

    $WebPartProperties = @{
        isDocumentLibrary = $false;
        selectedListId = $AllocationsReportList.ID.Guid;
        selectedListUrl = $AllocationsReportList.RootFolder.ServerRelativeUrl;
        webRelativeListUrl = '/Lists/' + $AllocationsReportList.rootfolder.name;
        selectedViewId = $NewView.Id.Guid;
        webpartHeightKey = 4;
        hideCommandBar = $true
    }
    $NewPageName = 'Agency Report for ' + $Agency
    $NewWebPartPage = Add-PnPPageWebPart -Page $NewPageName -DefaultWebPartType List -WebPartProperties $WebPartProperties
}
#endregion

#region allocations list
$AllocationsList = Get-PnPList -Identity 'Allocations'
$Fields = @(
    'LinkTitle',
    'Status',
    'Full_x0020_Name',
    'MissionArea',
    'MissionAgency',
    'PositionType',
    'Appointment',
    'Series',
    'Location',
    'State',
    'Employee_x002f_JOAStatus_x002f_D'
)
Write-Host "Creating Allocations List Views..."
foreach ($Agency in $Agencies) {
    $ViewName = 'Agency: ' + $Agency
    Write-Host -ForegroundColor Cyan $($ViewName)
    $Query="<OrderBy><FieldRef Name='Title' /></OrderBy><Where><Eq><FieldRef Name='Agency' /><Value Type='Text'>$Agency</Value></Eq></Where>"
    $NewView = Add-PnPView -List $AllocationsList -Title $ViewName -Fields $Fields -Query $Query -Paged -RowLimit 100
    $UpdatedView = Set-PnPView -List $AllocationsList -Identity $NewView.Title -Values @{'TabularView' = $false}

    $WebPartProperties = @{
        isDocumentLibrary = $false;
        selectedListId = $AllocationsList.ID.Guid;
        selectedListUrl = $AllocationsList.RootFolder.ServerRelativeUrl;
        webRelativeListUrl = '/Lists/' + $AllocationsList.rootfolder.name;
        selectedViewId = $NewView.Id.Guid;
        webpartHeightKey = 4;
        hideCommandBar = $true
    }
    $NewPageName = 'Agency Report for ' + $Agency
    $NewWebPart = Add-PnPPageWebPart -Page $NewPageName -DefaultWebPartType List -WebPartProperties $WebPartProperties
}
#endregion

#region create ses reports page
$AllocationsList = Get-PnPList -Identity 'Allocations'
$SitePages = Get-PnPList -Identity 'Site Pages'
$Agencies = (Get-PnPListItem -List $AllocationsList -PageSize 500).FieldValues.Agency | Sort-Object -Unique
$tenantUrl = 'https://usdagcc.sharepoint.com'


$items = @()
$itemId = 0

foreach ($Agency in $Agencies) {
    $itemId++
    $PageName = 'Agency Report for ' + $Agency
    $AgencyReportPage = Get-PnPListItem -List $SitePages | Where-Object {$_.FieldValues.Title -like $PageName}
    $url = $tenantUrl + $AgencyReportPage.FieldValues.FileRef
    $title = "$Agency"
    
    $sourceItem=@{
        url=$url;
        itemType=2;
        fileExtension="";
        progId=""
    }
    $linkDetails = @{
        sourceItem=$sourceItem
        title=$title;
        thumbnailType=3;
        id=$itemId
    }
    $items += $linkDetails
}
$AllocationsReportList = Get-PnPList -Identity 'Allocation Reports'
$SESView = Get-PnPView -List $AllocationsReportList -Identity 'SES'
$WebPartProperties = @{
    isDocumentLibrary = $false;
    selectedListId = $AllocationsReportList.ID.Guid;
    selectedListUrl = $AllocationsReportList.RootFolder.ServerRelativeUrl;
    webRelativeListUrl = '/Lists/' + $AllocationsReportList.rootfolder.name;
    selectedViewId = $SESView.Id.Guid;
    webpartHeightKey = 4;
    hideCommandBar = $true
}
$NewPageName = 'SES-Reporting'
$NewWebPartPage = Add-PnPPageWebPart -Page $NewPageName -DefaultWebPartType List -WebPartProperties $WebPartProperties

$WebPartProperties = @{
    title="Agencies";
    items=$items;
    hideWebPartWhenEmpty=$true
}  
$page = Get-PnPPage -Identity "SES-Reporting.aspx"
$NewWebPartParams = @{
    Page                    = $page
    DefaultWebPartType      = "QuickLinks"
    WebPartProperties       = $WebPartProperties
}
Add-PnPPageWebPart @NewWebPartParams

#endregion

#region create worker page web parts
$AllocationsList = Get-PnPList -Identity 'Allocations'
$SitePages = Get-PnPList -Identity 'Site Pages'
$Agencies = (Get-PnPListItem -List $AllocationsList -PageSize 500).FieldValues.Agency | Sort-Object -Unique

$items = @()
$itemId = 0

foreach ($Agency in $Agencies) {
    $itemId++
    $ListViewName = 'Agency: ' + $Agency
    $ListView = Get-PnPView -List $AllocationsList -Identity $ListViewName
    $url = $listview.ServerRelativeUrl
    $title = "$Agency"

    $sourceItem=@{
        url=$url;
        itemType=2;
        fileExtension="";
        progId=""
    }
    $linkDetails = @{
        sourceItem=$sourceItem
        title=$title;
        thumbnailType=3;
        id=$itemId
    }
    $items += $linkDetails
}
$WebPartProperties = @{
    title="Agencies";
    items=$items;
    hideWebPartWhenEmpty=$true
}  
$page = Get-PnPPage -Identity "Executive-Allocations.aspx"
$NewWebPartParams = @{
    Page                    = $page
    DefaultWebPartType      = "QuickLinks"
    WebPartProperties       = $WebPartProperties
}
Add-PnPPageWebPart @NewWebPartParams

#endregion






#region RESET
$SitePages = Get-PnPList -Identity 'Site Pages'

$PagesToDelete = Get-PnPListItem -List $SitePages | Where-Object {$_.FieldValues.Title -like 'Agency Report*'}
foreach ($page in $PagesToDelete) {
    Remove-PnPListItem -List $SitePages -Identity $page.Id -Force
}

$AllocationsList = Get-PnPList -Identity 'Allocations'
$ViewsToDelete = Get-PnPView -List $AllocationsList | Where-Object {$_.Title -like 'Agency:*'}
foreach ($ViewToDelete in $ViewsToDelete) {
    Remove-PnPView -List $AllocationsList -Identity $ViewToDelete -Force
}

$AllocationsReportList = Get-PnPList -Identity 'Allocation Reports'
$ViewsToDelete = Get-PnPView -List $AllocationsReportList | Where-Object {$_.Title -like 'Agency:*'}
foreach ($ViewToDelete in $ViewsToDelete) {
    Remove-PnPView -List $AllocationsList -Identity $ViewToDelete -Force
}
#endregion



