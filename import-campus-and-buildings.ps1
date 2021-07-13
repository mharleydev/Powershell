#region init
$ErrorActionPreference = "Stop"
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
#endregion

#region connect to the site
Write-Host
Write-Host "Initializing SPO site."
Write-Host
$SiteUrl = 'https://tenant.sharepoint.com/sites/site1'

Write-Host " -> Connecting to " -NoNewLine
Write-Host -Foregroundcolor Cyan $($SiteUrl) -NoNewLine
Write-Host "..." -NoNewLine

try {
	Connect-PnPOnline -Url $SiteUrl -UseWebLogin | Out-Null
	Write-Host -Foregroundcolor Green "Connected."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}
#endregion

#region import campus
$CampusListName = 'Campus List'
$CampusList = Get-PnPList -Identity $CampusListName

$CPAISImportArray = Import-Csv -Path 'C:\temp\cpais-import.csv' | Where-Object {$_.AGENCY -ne 11}
$CPAISImportArray | Add-Member -Name 'TEMP_ID' -Type noteproperty -value $Null
$CPAISImportArray | 
    ForEach-Object {
        $InstallationID = $_.INSTALLATION_ID.Trim()
        $SiteNumber = $_.SITE_NUMBER.Trim()
        $InstallationName =  ($_.INSTALLATION_NAME.Replace(' ','')).Trim()
        $State = ($_.PHYSICAL_LOCATION_STATE_NAME.Replace(' ','')).Trim()
        $City = ($_.PHYSICAL_CITY_NAME.Replace(' ','')).Trim()
        $TempID = $InstallationID + '-' + $SiteNumber + '-' + $InstallationName + '-' + $State + '-' + $City
        $_.TEMP_ID = $TempID
    }
$CampusIDs = $CPAISImportArray | Select-Object -Property TEMP_ID,INSTALLATION_ID,INSTALLATION_NAME,SITE_NUMBER,SITE_NAME | Where-Object {$_.TEMP_ID -ne $Null}
$UniqueCampusID = $CampusIDs | Sort-Object -Property TEMP_ID -Unique | Select-Object -First 50

$Campuses = Get-PnPListItem -List $CampusList -Pagesize 500

Write-Host "-> Adding " -NoNewline
Write-Host -ForegroundColor Cyan $(($UniqueCampusID | Measure-Object).Count) -NoNewline
Write-Host " campuses."

$iCampus = 0

$ListItemValues = @{
    Title               = 'Unassigned'
    campusStatus       = 'Active'
    tempID             = $null
}
$NewListItem = Add-PnPListItem -List $CampusList -Values $ListItemValues

foreach ($ID in $UniqueCampusID) {
    if ($ID.TEMP_ID -like '-*') {
        #do nothing
    
    }
    else {
        Write-Host "--> Processing " -NoNewline
        Write-Host -ForegroundColor Cyan $($ID.INSTALLATION_NAME + ' - ' + $ID.SITE_NAME) -NoNewline
        Write-Host "..." -NoNewline
        $ExistingCampus = $Campuses | Where-Object {$_.FieldValues.tempID -eq $ID.TEMP_ID}
        if (!$ExistingCampus) {
            Write-Host "Not found. Adding..." -NoNewline
            $CampusName = $ID.INSTALLATION_NAME + ' - ' + $ID.SITE_NAME
            $ListItemValues = @{
                Title               = $CampusName
                campusStatus       = 'Active'
                tempID             = $ID.TEMP_ID
                requestStatus       = 'Import'
            }
            $NewListItem = Add-PnPListItem -List $CampusList -Values $ListItemValues
            $iCampus++
            Write-Host -ForegroundColor Green "Complete."
        }
        else {
            Write-Host "$($ID.TEMP_ID) already exists. Doing nothing."
        }
    }
}
#endregion

#region import assets
$AssetListName = 'Asset List'
$AssetList = Get-PnPList -Identity $AssetListName
$CampusListName = 'Campus List'
$CampusList = Get-PnPList -Identity $CampusListName

$AssetColumns = Import-Csv -Path 'C:\temp\columns-assets.csv'

$CPAISImportArray = Import-Csv -Path 'C:\temp\cpais-import.csv' | Where-Object {$_.AGENCY -ne 11}
$CPAISImportArray | Add-Member -Name 'TEMP_ID' -Type noteproperty -value $Null
$CPAISImportArray | 
    ForEach-Object {
        $InstallationID = $_.INSTALLATION_ID.Trim()
        $SiteNumber = $_.SITE_NUMBER.Trim()
        $InstallationName =  ($_.INSTALLATION_NAME.Replace(' ','')).Trim()
        $State = ($_.PHYSICAL_LOCATION_STATE_NAME.Replace(' ','')).Trim()
        $City = ($_.PHYSICAL_CITY_NAME.Replace(' ','')).Trim()
        $TempID = $InstallationID + '-' + $SiteNumber + '-' + $InstallationName + '-' + $State + '-' + $City
        $_.TEMP_ID = $TempID
    }

$Campuses = Get-PnPListItem -List $CampusList -PageSize 500
$Assets = Get-PnPListItem -List $AssetList -PageSize 500

$iAssets = 0

[System.Collections.ArrayList]$ExceptionArray = @()

Write-Host "Importing assets"
foreach ($Campus in $Campuses) {
    Write-Host "-> Processing " -NoNewline
    Write-Host -ForegroundColor Cyan "$($Campus.FieldValues.Title)" -NoNewline
    Write-Host "..."

    $FolderName = $Campus.FieldValues.ID
    $ChildFolderPath = Join-Path -Path $AssetList.RootFolder.Name -ChildPath $FolderName
    $ChildFolder = Get-PnPFolder $ChildFolderPath -ErrorAction SilentlyContinue

    if (!$ChildFolder) {
        Write-Host "--> Adding Campus folder " -NoNewline
        Write-Host -ForegroundColor Cyan $($FolderName) -NoNewline
        Write-Host "..." -NoNewline
        try {
            $NewFolder = Add-PnPFolder -Name $FolderName -Folder $AssetList.RootFolder.Name
            $iAssets++
            Write-Host -ForegroundColor Green "Success."
        }
        catch {
            $err = $_
            Write-Host -Foregroundcolor Red $($err.Exception.Message)
        }
        
    }

    $Buildings = $CPAISImportArray | Where-Object {$_.TEMP_ID -eq $Campus.FieldValues.tempID}
    Write-Host "--> Found " -NoNewline
    Write-Host -ForegroundColor Cyan $(($Buildings | Measure-Object).Count) -NoNewline
    Write-Host " assets."
    
    $PrimaryBuildingName = $Buildings | Sort-Object -Property NO_OF_PERSONNEL | Select-Object -Last 1

    foreach ($Building in $Buildings) {
        Write-Host "---> " -NoNewline
        Write-Host -ForegroundColor Cyan $($Building.ASSET_NAME) -NoNewline
        $ExistingAsset = $Assets | Where-Object {$_.FieldValues.rpuid -eq $Building.RPUID}
        if (!$ExistingAsset) {
            Write-Host " not found. Adding... " -NoNewline

            if ($Building.PLANT_REPLACEMENT_VALUE -eq '#N/A') {
                $PlantReplacementValue = $Null
            }
            else {
                $PlantReplacementValue = $Building.PLANT_REPLACEMENT_VALUE
            }

            if ($ListItemValues.LEASE_EXPIRATION_DATE -eq '#N/A') {
                $LeaseExpirationDate = $Null
            }
            else {
                $LeaseExpirationDate = $ListItemValues.LEASE_EXPIRATION_DATE
            }

            if ($Building.ASSET_NAME -eq $PrimaryBuildingName.ASSET_NAME) {
                $PrimaryBuilding = $True
            } 
            else {
                $PrimaryBuilding = $False
            }

            try {
                $ListItemValues = @{
                    Title = $Building.ASSET_NAME
                    tempID = $Building.TEMP_ID
                    personnelAssigned = $Building.NO_OF_PERSONNEL
                    grossSquareFootage = $Building.GROSS_SQFT
                    usableSquareFootage = $Building.TOTAL_USABLE_SF
                    rentableSquareFootage = $Building.TOTAL_RENTABLE_SF
                    agency = $Null
                    agencyCode = $Building.AGENCY
                    addressLine1 = $Building.PHYSICAL_STREET_ADDR_1
                    addressLine2 = $Building.PHYSICAL_STREET_ADDR_2
                    city = $Building.PHYSICAL_CITY_NAME
                    state = $Building.PHYSICAL_LOCATION_STATE_NAME
                    zip = $Building.PHYSICAL_ZIP_CODE
                    campusID = $Campus.FieldValues.ID
                    campusName = $Campus.FieldValues.Title
                    primaryAsset = $PrimaryBuilding
                    rpuid = $Building.RPUID
                    installationID = $Building.INSTALLATION_ID
                    installationName = $Building.INSTALLATION_NAME
                    siteNumber = $Building.SITE_NUMBER
                    siteName = $Building.SITE_NAME
                    propertyType = $Building.PROPERTY_TYPE
                    assetType = $Building.ASSET_TYPE
                    predominantUseCategory = $Building.PREDOMINANT_USE_CATEGORY
                    predominantUseSubcategory = $Building.PREDOMINANT_USE_SUBCATEGORY
                    missionDependency = $Building.MISSION_DEPENDENCY
                    status = $Building.STATUS
                    assetStatus = $Building.PRIMARY_ASSET
                    childcare = $Building.CHILDCARE_CTR_IN_THE_BUILDING
                    orgCode = $Building.ORG_CODE
                    orgCodeName = $Building.ORG_CODE_NAME
                    gsaRegion = $Building.GSA_REGION
                    gsaRegionName = $Building.GSA_REGION_NAME
                    plantReplacementValue = $PlantReplacementValue 
                    leaseExpirationDate = $LeaseExpirationDate
                }
                $NewPnPListItem = Add-PnPListItem -List $AssetList -Folder $FolderName -Values $ListItemValues
                $iAssets++
                Write-Host -ForegroundColor Green "Success."
            }
            catch {
                $err = $_
                Write-Host -Foregroundcolor Red $($err.Exception.Message)
            }
        }
        else {
            Write-Host " already exists. Doing nothing."
        }
    }
}
#endregion

#region roll up asset numbers to campus
$AssetListName = 'Asset List'
$AssetList = Get-PnPList -Identity $AssetListName
$Assets = Get-PnPListItem -List $AssetList

$CampusListName = 'Campus List'
$CampusList = Get-PnPList -Identity $CampusListName
$Campuses = Get-PnPListItem -List $CampusList

if (!$Agencies) {
    $AgencyListName = 'Agency Code Matrix'
    $AgencyList = Get-PnPList -Identity $AgencyListName
    $Agencies = Get-PnPListItem -List $AgencyList
}

[System.Collections.ArrayList]$AgencyListNames = @()
[System.Collections.ArrayList]$StateNamesList = @()

Write-Host "-> Rolling up " -NoNewline
Write-Host -ForegroundColor Cyan $(($Campuses | Measure-Object).Count) -NoNewline
Write-Host ' campuses.'

foreach ($Campus in $Campuses) {
    Write-Host '--> ' -NoNewline
    Write-Host -ForegroundColor Cyan $Campus.FieldValues.ID $Campus.FieldValues.Title -NoNewline
    Write-Host '...' -NoNewline
    $Buildings = $Assets | Where-Object {$_.FieldValues.campusID -eq $Campus.FieldValues.ID}
    if ($Buildings) {
        $NO_OF_PERSONNEL = ($Buildings.FieldValues.personnelAssigned | Measure-Object -Sum).Sum
        $PrimaryBuilding = $Buildings | Where-Object {$_.FieldValues.primaryAsset -eq $True}
        [System.Collections.ArrayList]$CampusSQFTArray = @()
        foreach ($Building in $Buildings) {
            [System.Collections.ArrayList]$BuildingSQFTArray = @()
            $GROSS_SQFT = ($Building.FieldValues.grossSquareFootage | Measure-Object -Sum).Sum
            $TOTAL_USABLE_SF = ($Building.FieldValues.usableSquareFootage | Measure-Object -Sum).Sum
            $TOTAL_RENTABLE_SF = ($Building.FieldValues.rentableSquareFootage | Measure-Object -Sum).Sum
        
            $BuildingSQFTArray.Add($GROSS_SQFT) | Out-Null
            $BuildingSQFTArray.Add($TOTAL_USABLE_SF) | Out-Null
            $BuildingSQFTArray.Add($TOTAL_RENTABLE_SF) | Out-Null
        
            $BuildingSQFT = $BuildingSQFTArray | Sort-Object | Select-Object -Last 1
        
            $CampusSQFTArray.Add($BuildingSQFT) | Out-Null
        }
        $TOTAL_SQFT = ($CampusSQFTArray | Measure-Object -Sum).Sum

        $AgencyIDs = $Buildings.FieldValues.agencyCode
        $UniqueAgencyIDs = $AgencyIDs | Sort-Object -Unique
        [string]$AgencyList = $Null
        foreach ($Agency in $UniqueAgencyIDs) {
            $AgencyName = ($Agencies | Where-Object {$_.FieldValues.Title -eq $Agency}).FieldValues.Name
            $AgencyList += $AgencyName + ', '
            $object = New-Object PSObject -Property @{
                AgencyName = $AgencyName
                AgencyCode = $Agency
            }
            $AgencyListNames.Add($object) | Out-Null
        }

        $StateNamesList.Add($PrimaryBuilding.FieldValues.state) | Out-Null

        $ListItemValues = @{
            personnelAssigned   = $NO_OF_PERSONNEL
            squareFootage       = $TOTAL_SQFT
            agencyList          = $AgencyList.TrimEnd(', ')
            addressLine1        = $PrimaryBuilding.FieldValues.addressLine1
            addressLine2        = $PrimaryBuilding.FieldValues.addressLine2
            state               = $PrimaryBuilding.FieldValues.state
            city                = $PrimaryBuilding.FieldValues.city
            zip                 = $PrimaryBuilding.FieldValues.zip
        }
        Set-PnPListItem -List $CampusList -Identity $Campus.Id -Values $ListItemValues | Out-Null
        Write-Host -ForegroundColor Green 'Done.'
    }
    else {
        Write-Host "No buildings found. Doing nothing."
    }
}

$Fields = @(
    "LinkTitle",
    "agencyList",
    "addressLine1",
    "addressLine2",
    "city",
    "state",
    "zip",
    "personnelAssigned",
    "squareFootage",
    "fslLevel",
    "upcomingAssessment",
    "outstandingAssessment",
    "lastAssessmentDate",
    "upcomingAssessmentDate"
)
$Query="<OrderBy><FieldRef Name='LinkTitle' Ascending='TRUE'/></OrderBy>"
$NewView = Add-PnPView -List 'Campus List' -Title 'Summary' -Fields $Fields -SetAsDefault -Query $Query -Paged -RowLimit 100

$UniqueAgencyListNames = $AgencyListNames | Sort-Object -Property AgencyCode -Unique
foreach ($AgencyListName in $UniqueAgencyListNames) {
    $QueryValue = $AgencyListName.AgencyName
    $Query = "<Where><Eq><FieldRef Name = 'agencyList' /><Value Type = 'Text'>$QueryValue</Value></Eq></Where><OrderBy><FieldRef Name='LinkTitle'/></OrderBy>"
    $ViewTitle = "Agency: " + $AgencyListName.AgencyCode
    $NewView = Add-PnPView -List $CampusListName -Title $ViewTitle -Fields $Fields -Query $Query -Paged -RowLimit 100
}

$UniqueStates = $StateNamesList | Sort-Object -Unique
foreach ($State in $UniqueStates) {
    $QueryValue = $State
    $Query = "<Where><Eq><FieldRef Name = 'state' /><Value Type = 'Text'>$QueryValue</Value></Eq></Where><OrderBy><FieldRef Name='LinkTitle'/></OrderBy>"
    $ViewTitle = "State: " + (Get-Culture).TextInfo.ToTitleCase($State.ToLower())
    $NewView = Add-PnPView -List $CampusListName -Title $ViewTitle -Fields $Fields -Query $Query -Paged -RowLimit 30

}
#endregion

#region set permissions
$AssetListName = 'Asset List'
$AssetList = Get-PnPList -Identity $AssetListName
$Assets = Get-PnPListItem -List $AssetList

$CampusListName = 'Campus List'
$CampusList = Get-PnPList -Identity $CampusListName
$Campuses = Get-PnPListItem -List $CampusList

$OSSPAdminsGroupName = 'OSSP Admins'

if (!$Agencies) {
    $AgencyListName = 'Agency Code Matrix'
    $AgencyList = Get-PnPList -Identity $AgencyListName
    $Agencies = Get-PnPListItem -List $AgencyList
}

[System.Collections.ArrayList]$SiteGroups = @()
foreach ($Group in (Get-PnPGroup)) {
    $SiteGroups.Add($Group) | Out-Null
}

Write-Host '-> Processing campus permissions...'
foreach ($Campus in $Campuses) {
    Write-Host '--> ' -NoNewline
    Write-Host -ForegroundColor Cyan $($Campus.FieldValues.Title)
    

    #region validate groups

    #region do group
    $DOGroupName = 'Designated Official - ' + $Campus.FieldValues.ID
    $DOGroup = $SiteGroups | Where-Object {$_.Title -eq $DOGroupName}
    if ($DOGroup) {
        # do nothing
    }
    else {
        try {
            $NewPnPGroupParams = @{
                Title = $DOGroupName
                Description = "Group used to grant permissions to Designated Officals for $($Campus.FieldValues.ID) $($Campus.FieldValues.Title)" 
            }
            $DOGroup = New-PnPGroup @NewPnPGroupParams
            $SiteGroups.Add($DOGroup) | Out-Null
        }
        catch {
            $err = $_
            Write-Host -Foregroundcolor Red $($err.Exception.Message)
        }
    }
    #endregion

    #region agency groups
    $Buildings = $Assets | Where-Object {$_.FieldValues.campusID -eq $Campus.FieldValues.ID}
    $AgencyIDs = $Buildings.FieldValues.agencyCode
    $UniqueAgencyIDs = $AgencyIDs | Sort-Object -Unique
    foreach ($Agency in $UniqueAgencyIDs) {
        $AgencyName = ($Agencies | Where-Object {$_.FieldValues.Title -eq $Agency}).FieldValues.Name
        $APSSGroupName = 'APSS Agency Admins - ' + $AgencyName
        Write-Host "----> " -NoNewline
        Write-Host -ForegroundColor Cyan $($APSSGroupName) -NoNewline
        Write-Host "..." -NoNewline
        $APSSGroup = $SiteGroups | Where-Object {$_.Title -eq $APSSGroupName}
        if ($APSSGroup) {
            Write-Host -ForegroundColor Green 'Found.'
        }
        else {
            Write-Host "Not found. Creating..." -NoNewline
            try {
                $NewPnPGroupParams = @{
                    Title = $APSSGroupName
                    Description = "Group used to grant permissions to APSS Agency Admins for all $($AgencyName) campuses." 
                }
                $APSSGroup = New-PnPGroup @NewPnPGroupParams
                $SiteGroups.Add($APSSGroup) | Out-Null
                Write-Host -Foregroundcolor Green "Complete."
            }
            catch {
                $err = $_
                Write-Host -Foregroundcolor Red $($err.Exception.Message)
            }
        }
    }
    #endregion

    #endregion

    #region setting permissions
    $FolderName = $Campus.FieldValues.ID
    $ChildFolderPath = Join-Path -Path $AssetList.RootFolder.Name -ChildPath $FolderName
    $ChildFolder = Get-PnPFolder $ChildFolderPath -ErrorAction SilentlyContinue

    Set-PnPFolderPermission -List $AssetList -Identity $ChildFolder -Group $OSSPAdminsGroupName -ClearExisting -AddRole 'Full Control'
    Set-PnPListItemPermission -List $CampusList -Identity $Campus.Id -Group $OSSPAdminsGroupName -ClearExisting -AddRole 'Full Control'

    Set-PnPFolderPermission -List $AssetList -Identity $ChildFolder -Group $DOGroup -AddRole 'Read'
    Set-PnPListItemPermission -List $CampusList -Identity $Campus.Id -Group $DOGroup -AddRole 'Read'

    $Buildings = $Assets | Where-Object {$_.FieldValues.campusID -eq $Campus.FieldValues.ID}
    $AgencyIDs = $Buildings.FieldValues.agencyCode
    $UniqueAgencyIDs = $AgencyIDs | Sort-Object -Unique
    foreach ($Agency in $UniqueAgencyIDs) {
        $AgencyName = ($Agencies | Where-Object {$_.FieldValues.Title -eq $Agency}).FieldValues.Name
        $APSSGroupName = 'APSS Agency Admins - ' + $AgencyName
        Set-PnPFolderPermission -List $AssetList -Identity $ChildFolder -Group $APSSGroupName -AddRole 'NoDelete'
        Set-PnPListItemPermission -List $CampusList -Identity $Campus.Id -Group $APSSGroupName -AddRole 'NoDelete'

    }
    Set-PnPFolderPermission -List $AssetList -Identity $ChildFolder -User 'user@email.com' -RemoveRole 'Full Control'
    Set-PnPListItemPermission -List $CampusList -Identity $Campus.Id -User 'user@email.com' -RemoveRole 'Full Control'
    #endregion
}
#endregion

