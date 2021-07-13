#Requires -Modules PnP.PowerShell

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

#region create OSSP Admins group
$NewGroupName = 'OSSP Admins'
$ExistingGroup = Get-PnPGroup -Identity $NewGroupName -ErrorAction SilentlyContinue
Write-Host " -> Checking for group, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($NewGroupName) -NoNewLine
Write-Host "..." -NoNewLine
if ($ExistingGroup) {
	Write-Host -ForegroundColor Green 'Found.'
}
else {
	Write-Host "Not found. Creating..." -NoNewline
	try {
		$NewPnPGroupParams = @{
			Title = $NewGroupName
			Description = "Group used to grant permissions to OSSP Admins" 
		}
		$NewGroup = New-PnPGroup @NewPnPGroupParams
		Set-PnPGroup -Identity $NewGroupname -AddRole 'Full Control'
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}

}

#endregion

#region create NoDelete permissions role
Write-Host '-> Verifying ' -NoNewline
Write-Host -ForegroundColor Cyan 'NoDelete' -NoNewline
Write-Host ' permission level is present...' -NoNewline
$ExistingNoDeleteRole = Get-PnPRoleDefinition -Identity 'NoDelete' -ErrorAction SilentlyContinue
if ($ExistingNoDeleteRole) {
	Write-Host -ForegroundColor Green 'Found.'
}
else {
	Write-Host 'Not found. Creating...' -NoNewline
	try {
		$NoDeleteRole = Add-PnPRoleDefinition -RoleName "NoDelete" -Clone "Contribute" -Exclude DeleteListItems -Description "Contribute without Delete."
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}	
}
#endregion

#region create Agency Code Matrix
$ListName = 'Agency Code Matrix'
if (Get-PnPList -Identity $ListName) {
	Write-Host " -> Found " -NoNewLine
	Write-Host -Foregroundcolor Cyan $($ListName) -NoNewLine
	Write-Host ". Doing nothing."
}
else {
	$ListUrl = 'agency-code-matrix'
	Write-Host " -> Creating list, " -NoNewLine
	Write-Host -Foregroundcolor Cyan $($ListName) -NoNewLine
	Write-Host "..." -NoNewLine
	try {
		$NewList = New-PnPList -Title $ListName -Template GenericList -Url $ListUrl
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
	$NewFieldName = 'Name'
	Write-Host " -> Adding " -NoNewline
	Write-Host -ForegroundColor Cyan $($NewFieldName) -NoNewline
	Write-Host " field. " -NoNewLine
	try {
		$NewField = Add-PnPField -List $NewList -Type 'Text' -InternalName $NewFieldName -DisplayName $NewFieldName -AddToDefaultView
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
	$TitleFieldValues = @{
		Title = 'Code'
		Indexed = $true
	}
	Set-PnPField -List $NewList -Identity "Title" -Values $TitleFieldValues
	
	$ImportArray = Import-Csv -Path 'C:\temp\agency-code-matrix.csv'
	foreach ($Row in $ImportArray) {
		$ListItemValues = @{
			Title 	= $Row.Code
			Name 	= $Row.Name
		}
		$NewListItem = Add-PnPListItem -List $NewList -Values $ListItemValues
	}
}


#endregion

#region create campus list
$CampusListName = 'Campus List'
$CampusListUrl = 'campus-list'
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($CampusListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$CampusList = New-PnPList -Title $CampusListName -Template GenericList -Url $CampusListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}

$CampusListColumns = Import-Csv -Path 'C:\temp\columns-campus.csv'

$CampusList = Get-PnPList -Identity $CampusListName

Write-Host " -> Adding columns... "
foreach ($Column in $CampusListColumns) {
	Write-Host "  -> $($Column.Name)... " -NoNewLine
	try {
		if ($Column.AddToDefaultView -eq 0) {
			$AddToDefaultView = $False
		} else {
			$AddToDefaultView = $False
		}
		$AddPnPFieldParams = @{
			List 				= $CampusList
			Type 				= $Column.FieldType
			InternalName 		= (($Column.Name.Substring(0,1)).ToLower() + $Column.Name.Substring(1)).Replace(' ','') 
			DisplayName 		= $Column.Name
			AddToDefaultView 	= $AddToDefaultView
		}
		$NewField = Add-PnPField @AddPnPFieldParams
		if ($Column.Index -eq 'Yes') {
			$NewField.Indexed = $True
			$NewField.Update()
		}
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
	
}
$Choices = @("On-site", "Self-service")
$AddPnPFieldParams = @{
	List 				= $CampusList
	Type 				= "Choice"
	InternalName 		= 'upcomingAssessmentType'
	DisplayName 		= 'Upcoming Assessment Type'
	AddToDefaultView 	= 1
	Choices 			= $Choices
}
$NewField = Add-PnpField @AddPnPFieldParams

$AddPnPFieldParams = @{
	List 				= $CampusList
	Type 				= 'User'
	InternalName 		= 'assessmentTeamLead'
	DisplayName 		= 'Assessment Team Lead'
	AddToDefaultView 	= 1
}
$NewField = Add-PnpField @AddPnPFieldParams

$AddPnPFieldParams = @{
	List 				= $CampusList
	Type 				= 'DateTime'
	InternalName 		= 'upcomingAssessment30Days'
	DisplayName 		= 'Upcoming Assessment 30 Days'
	AddToDefaultView 	= 1
}
$NewField = Add-PnpField @AddPnPFieldParams
[xml]$schemaXml=$NewField.SchemaXml
$schemaXml.Field.SetAttribute('Format','DateOnly')
Set-PnPField -List $CampusList -Identity $NewField.InternalName -Values @{SchemaXml=$schemaXml.OuterXml}


$AddPnPFieldParams = @{
	List 				= $CampusList
	Type 				= 'DateTime'
	InternalName 		= 'upcomingAssessment60Days'
	DisplayName 		= 'Upcoming Assessment 60 Days'
	AddToDefaultView 	= 1
}
$NewField = Add-PnpField @AddPnPFieldParams
[xml]$schemaXml=$NewField.SchemaXml
$schemaXml.Field.SetAttribute('Format','DateOnly')
Set-PnPField -List $CampusList -Identity $NewField.InternalName -Values @{SchemaXml=$schemaXml.OuterXml}

$AddPnPFieldParams = @{
	List 				= $CampusList
	Type 				= 'UserMulti'
	InternalName 		= 'assessmentTeamMembers'
	DisplayName 		= 'Assessment Team Members'
	AddToDefaultView 	= 1
}

$AddPnPFieldParams = @{
	List 				= $CampusList
	Type 				= 'Text'
	InternalName 		= 'requestAction'
	DisplayName 		= 'RequestAction'
	AddToDefaultView 	= 1
}
$NewField = Add-PnpField @AddPnPFieldParams
$NewField.Indexed = $True
$NewField.Update()

$xml = '<Field Type="UserMulti" Name="assessmentTeamMembers" DisplayName="Assessment Team Members" ID="{b7ec2ed1-c723-4d94-a2d4-6815e46a3a5c}" Group="" Required="FALSE" SourceID="{6e854e0b-8d32-4ac6-a4e5-42bc429a1482}" StaticName="assessmentTeamMembers" ColName="int3" RowOrdinal="0" EnforceUniqueValues="FALSE" ShowField="ImnName" UserSelectionMode="PeopleAndGroups" UserSelectionScope="0" Mult="TRUE" Sortable="FALSE" Version="1" />'
$xml = '<Field Type="UserMulti" Name="assessmentTeamMembers" DisplayName="Assessment Team Members" ID="{b7ec2ed1-c723-4d94-a2d4-6815e46a3a5c}" Group="" Required="FALSE" SourceID="{6e854e0b-8d32-4ac6-a4e5-42bc429a1482}" Mult="TRUE" />'
$NewField = Add-PnpFieldFromXml -List $CampusList -FieldXml $xml

$Field = Get-PnPField -List $CampusName -Identity 'assessmentTeamMembers'


$TitleFieldValues = @{
	Title 	= 'Campus'
	Indexed = $true
}
Set-PnPField -List $CampusList -Identity "Title" -Values $TitleFieldValues
$TitleFieldValues = @{
	Title 	= 'FSL Level'
}
Set-PnPField -List $CampusList -Identity 'fslLevel' -Values $TitleFieldValues

$DateFields = @('Last Assessment Date', 'Upcoming Assessment Date', 'Upcoming Assessment 30 Days', 'Upcoming Assessment 60 Days')
foreach ($DateFieldName in $DateFields) {
	$field = Get-PnPField -List $CampusListName -Identity $DateFieldName
	[xml]$schemaXml=$field.SchemaXml
	$schemaXml.Field.SetAttribute('Format','DateOnly')
	Set-PnPField -List $CampusListName -Identity $DateFieldName -Values @{SchemaXml=$schemaXml.OuterXml}
}




#endregion

#region create asset list
$AssetListName = 'Asset List'
$AssetListUrl = 'asset-list'
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($AssetListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$AssetList = New-PnPList -Title $AssetListName -Template GenericList -Url $AssetListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}
Set-PnPList -Identity $AssetListName -EnableFolderCreation $True
$AssetListColumns = Import-Csv -Path 'C:\temp\columns-assets.csv'
Write-Host " -> Adding columns... "
foreach ($Column in $AssetListColumns) {
	Write-Host "  -> $($Column.DisplayName)... " -NoNewLine
	try {
		$AddPnPFieldParams = @{
			List 				= $AssetList
			Type 				= $Column.FieldType
			InternalName 		= $Column.InternalName
			DisplayName 		= $Column.DisplayName
			AddToDefaultView 	= $true
		}
		$NewField = Add-PnPField @AddPnPFieldParams
		if ($Column.Index -eq 'Yes') {
			$NewField.Indexed = $True
			$NewField.Update()
		}
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
	
}
$TitleFieldValues = @{
	Title 	= 'Asset'
	Indexed = $true
}
Set-PnPField -List $AssetList -Identity "Title" -Values $TitleFieldValues
#endregion

#region create config lists



#region questions config
$NewListName = 'Questions Config'
$NewListUrl = ($NewListName.ToLower()).Replace(' ','-')
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($NewListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$NewList = New-PnPList -Title $NewListName -Template GenericList -Url $NewListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}

[System.Collections.ArrayList]$NewListColumns = @()
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Number'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'sectionID'
	"InternalName" 		= 'sectionID'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Number'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'worksheetVersionID'
	"InternalName" 		= 'worksheetVersionID'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Note'
	"Index" 			= 'No'
	"DisplayName" 		= 'config'
	"InternalName" 		= 'config'
	"AddToDefaultView" 	= 0
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Number'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'weight'
	"InternalName" 		= 'weight'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null

Write-Host " -> Adding columns... "
foreach ($Column in $NewListColumns) {
	Write-Host "  -> $($Column.DisplayName)... " -NoNewLine
	try {
		if ($Column.AddToDefaultView -eq 0) {
			$AddToDefaultView = $False
		} else {
			$AddToDefaultView = $True
		}
		$AddPnPFieldParams = @{
			List 				= $NewList
			Type 				= $Column.FieldType
			InternalName 		= $Column.InternalName
			DisplayName 		= $Column.DisplayName
			AddToDefaultView 	= $AddToDefaultView
		}
		$NewField = Add-PnPField @AddPnPFieldParams
		if ($Column.Index -eq 'Yes') {
			$NewField.Indexed = $True
			$NewField.Update()
		}
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
}
#endregion

#region WorksheetVersion
$NewListName = 'WorksheetVersion'
$NewListUrl = ($NewListName.ToLower()).Replace(' ','-')
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($NewListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$NewList = New-PnPList -Title $NewListName -Template GenericList -Url $NewListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}

$NewList = Get-PnPList -Identity $NewListName

[System.Collections.ArrayList]$NewListColumns = @()
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Text'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'version'
	"InternalName" 		= 'version'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'DateTime'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'effectiveStartDate'
	"InternalName" 		= 'effectiveStartDate'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'DateTime'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'effectiveEndDate'
	"InternalName" 		= 'effectiveEndDate'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Boolean'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'isActive'
	"InternalName" 		= 'isActive'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null


Write-Host " -> Adding columns... "
foreach ($Column in $NewListColumns) {
	Write-Host "  -> $($Column.DisplayName)... " -NoNewLine
	try {
		if ($Column.AddToDefaultView -eq 0) {
			$AddToDefaultView = $False
		} else {
			$AddToDefaultView = $True
		}
		$AddPnPFieldParams = @{
			List 				= $NewList
			Type 				= $Column.FieldType
			InternalName 		= $Column.InternalName
			DisplayName 		= $Column.DisplayName
			AddToDefaultView 	= $AddToDefaultView
		}
		$NewField = Add-PnPField @AddPnPFieldParams
		if ($Column.Index -eq 'Yes') {
			$NewField.Indexed = $True
			$NewField.Update()
		}
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
}

Set-PnPField -Identity "Title" -List $NewList -Values @{DefaultValue="defaultTitle"}
#endregion

#region WorksheetSection
$NewListName = 'WorksheetSection'
$NewListUrl = ($NewListName.ToLower()).Replace(' ','-')
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($NewListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$NewList = New-PnPList -Title $NewListName -Template GenericList -Url $NewListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}

$NewList = Get-PnPList -Identity $NewListName

[System.Collections.ArrayList]$NewListColumns = @()
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Text'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'sectionID'
	"InternalName" 		= 'sectionID'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Lookup'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'worksheetVersionID'
	"InternalName" 		= 'worksheetVersionID'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Number'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'weight'
	"InternalName" 		= 'weight'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Text'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'name'
	"InternalName" 		= 'name'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Boolean'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'isValid'
	"InternalName" 		= 'isValid'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Boolean'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'isComplete'
	"InternalName" 		= 'isComplete'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Boolean'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'isReadOnly'
	"InternalName" 		= 'isReadOnly'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null

Write-Host " -> Adding columns... "
foreach ($Column in $NewListColumns) {
	Write-Host "  -> $($Column.DisplayName)... " -NoNewLine
	try {
		if ($Column.AddToDefaultView -eq 0) {
			$AddToDefaultView = $False
		} else {
			$AddToDefaultView = $True
		}
		$AddPnPFieldParams = @{
			List 				= $NewList
			Type 				= $Column.FieldType
			InternalName 		= $Column.InternalName
			DisplayName 		= $Column.DisplayName
			AddToDefaultView 	= $AddToDefaultView
		}
		$NewField = Add-PnPField @AddPnPFieldParams
		if ($Column.Index -eq 'Yes') {
			$NewField.Indexed = $True
			$NewField.Update()
		}
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
}

Set-PnPField -Identity "Title" -List $NewList -Values @{DefaultValue="defaultTitle"}

#set lookup field
$sourcelistID = (Get-PnPList -Identity 'WorksheetVersion').Id.Guid
$sourcefieldname = 'version'
$ctx = Get-PnPContext
$lookupField = Get-PnPField -List $NewList -Identity 'worksheetVersionID'
$lookupField = $lookupField.TypedObject
$lookupField.LookupList = $sourcelistID
$lookupField.LookupField = $sourcefieldname
$lookupField.update()
$ctx.ExecuteQuery()

#endregion

#region WorksheetPart
$NewListName = 'WorksheetPart'
$NewListUrl = ($NewListName.ToLower()).Replace(' ','-')
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($NewListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$NewList = New-PnPList -Title $NewListName -Template GenericList -Url $NewListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}

$NewList = Get-PnPList -Identity $NewListName

[System.Collections.ArrayList]$NewListColumns = @()
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Lookup'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'sectionID'
	"InternalName" 		= 'sectionID'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Lookup'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'worksheetVersionID'
	"InternalName" 		= 'worksheetVersionID'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Note'
	"Index" 			= 'No'
	"DisplayName" 		= 'config'
	"InternalName" 		= 'config'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Note'
	"Index" 			= 'No'
	"DisplayName" 		= 'options'
	"InternalName" 		= 'options'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null


Write-Host " -> Adding columns... "
foreach ($Column in $NewListColumns) {
	Write-Host "  -> $($Column.DisplayName)... " -NoNewLine
	try {
		if ($Column.AddToDefaultView -eq 0) {
			$AddToDefaultView = $False
		} else {
			$AddToDefaultView = $True
		}
		$AddPnPFieldParams = @{
			List 				= $NewList
			Type 				= $Column.FieldType
			InternalName 		= $Column.InternalName
			DisplayName 		= $Column.DisplayName
			AddToDefaultView 	= $AddToDefaultView
		}
		$NewField = Add-PnPField @AddPnPFieldParams
		if ($Column.Index -eq 'Yes') {
			$NewField.Indexed = $True
			$NewField.Update()
		}
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
}

Set-PnPField -Identity "Title" -List $NewList -Values @{DefaultValue="defaultTitle"}


#set lookup field
$sourcelistID = (Get-PnPList -Identity 'WorksheetVersion').Id.Guid
$sourcefieldname = 'version'
$ctx = Get-PnPContext
$lookupField = Get-PnPField -List $NewList -Identity 'worksheetVersionID'
$lookupField = $lookupField.TypedObject
$lookupField.LookupList = $sourcelistID
$lookupField.LookupField = $sourcefieldname
$lookupField.update()
$ctx.ExecuteQuery()


#set lookup field
$sourcelistID = (Get-PnPList -Identity 'WorksheetSection').Id.Guid
$sourcefieldname = 'sectionID'
$ctx = Get-PnPContext
$lookupField = Get-PnPField -List $NewList -Identity 'sectionID'
$lookupField = $lookupField.TypedObject
$lookupField.LookupList = $sourcelistID
$lookupField.LookupField = $sourcefieldname
$lookupField.update()
$ctx.ExecuteQuery()
#endregion

#region WorksheetPartData
$NewListName = 'WorksheetPartData'
$NewListUrl = ($NewListName.ToLower()).Replace(' ','-')
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($NewListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$NewList = New-PnPList -Title $NewListName -Template GenericList -Url $NewListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}

$NewList = Get-PnPList -Identity $NewListName

[System.Collections.ArrayList]$NewListColumns = @()
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Lookup'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'worksheetPartId'
	"InternalName" 		= 'worksheetPartId'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Number'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'assessmentId'
	"InternalName" 		= 'assessmentId'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Note'
	"Index" 			= 'No'
	"DisplayName" 		= 'model'
	"InternalName" 		= 'model'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null


Write-Host " -> Adding columns... "
foreach ($Column in $NewListColumns) {
	Write-Host "  -> $($Column.DisplayName)... " -NoNewLine
	try {
		if ($Column.AddToDefaultView -eq 0) {
			$AddToDefaultView = $False
		} else {
			$AddToDefaultView = $True
		}
		$AddPnPFieldParams = @{
			List 				= $NewList
			Type 				= $Column.FieldType
			InternalName 		= $Column.InternalName
			DisplayName 		= $Column.DisplayName
			AddToDefaultView 	= $AddToDefaultView
		}
		$NewField = Add-PnPField @AddPnPFieldParams
		if ($Column.Index -eq 'Yes') {
			$NewField.Indexed = $True
			$NewField.Update()
		}
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
}

Set-PnPField -Identity "Title" -List $NewList -Values @{DefaultValue="defaultTitle"}


#set lookup field
$sourcelistID = (Get-PnPList -Identity 'WorksheetPart').Id.Guid
$sourcefieldname = 'ID'
$ctx = Get-PnPContext
$lookupField = Get-PnPField -List $NewList -Identity 'worksheetPartId'
$lookupField = $lookupField.TypedObject
$lookupField.LookupList = $sourcelistID
$lookupField.LookupField = $sourcefieldname
$lookupField.update()
$ctx.ExecuteQuery()
#endregion

#region Assessment
$NewListName = 'Assessment'
$NewListUrl = ($NewListName.ToLower()).Replace(' ','-')
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($NewListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$NewList = New-PnPList -Title $NewListName -Template GenericList -Url $NewListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}

$NewList = Get-PnPList -Identity $NewListName

[System.Collections.ArrayList]$NewListColumns = @()
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Text'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'designatedOfficial'
	"InternalName" 		= 'designatedOfficial'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Text'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'assessmentTeamLead'
	"InternalName" 		= 'assessmentTeamLead'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Number'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'campusId'
	"InternalName" 		= 'campusId'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Text'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'campus'
	"InternalName" 		= 'campus'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Lookup'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'worksheetVersionID'
	"InternalName" 		= 'worksheetVersionID'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'DateTime'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'dateOfAssessment'
	"InternalName" 		= 'dateOfAssessment'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Number'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'fslLevel'
	"InternalName" 		= 'fslLevel'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Boolean'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'isPreAssessmentComplete'
	"InternalName" 		= 'isPreAssessmentComplete'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null
$object = New-Object PSObject -Property @{
	"FieldType" 		= 'Boolean'
	"Index" 			= 'Yes'
	"DisplayName" 		= 'isAssessmentComplete'
	"InternalName" 		= 'isAssessmentComplete'
	"AddToDefaultView" 	= 1
}
$NewListColumns.Add($object) | Out-Null



Write-Host " -> Adding columns... "
foreach ($Column in $NewListColumns) {
	Write-Host "  -> $($Column.DisplayName)... " -NoNewLine
	try {
		if ($Column.AddToDefaultView -eq 0) {
			$AddToDefaultView = $False
		} else {
			$AddToDefaultView = $True
		}
		$AddPnPFieldParams = @{
			List 				= $NewList
			Type 				= $Column.FieldType
			InternalName 		= $Column.InternalName
			DisplayName 		= $Column.DisplayName
			AddToDefaultView 	= $AddToDefaultView
		}
		$NewField = Add-PnPField @AddPnPFieldParams
		if ($Column.Index -eq 'Yes') {
			$NewField.Indexed = $True
			$NewField.Update()
		}
		Write-Host -Foregroundcolor Green "Complete."
	}
	catch {
		$err = $_
		Write-Host -Foregroundcolor Red $($err.Exception.Message)
	}
}

Set-PnPField -Identity "Title" -List $NewList -Values @{DefaultValue="defaultTitle"}


#set lookup field
$sourcelistID = (Get-PnPList -Identity 'WorksheetVersion').Id.Guid
$sourcefieldname = 'version'
$ctx = Get-PnPContext
$lookupField = Get-PnPField -List $NewList -Identity 'worksheetVersionID'
$lookupField = $lookupField.TypedObject
$lookupField.LookupList = $sourcelistID
$lookupField.LookupField = $sourcefieldname
$lookupField.update()
$ctx.ExecuteQuery()
#endregion


#endregion

#region create import library
$ListName = 'Location Import Library'
$ListUrl = 'location_import_library'
Write-Host " -> Creating list, " -NoNewLine
Write-Host -Foregroundcolor Cyan $($ListName) -NoNewLine
Write-Host "..." -NoNewLine
try {
	$PSASLocationImportList = New-PnPList -Title $ListName -Template DocumentLibrary -Url $ListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}
Write-Host " -> Adding status field and setting default value..." -NoNewLine
try {
	$FieldName = 'ImportStatus'
	$ImportStatusField = Add-PnPField -List $PSASLocationImportList -Type 'Text' -InternalName $FieldName -DisplayName $FieldName -AddToDefaultView
	Set-PnPDefaultColumnValues -List $PSASLocationImportList -Field $ImportStatusField -Value 'New'
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}

Write-Host " -> Uploading test data..." -NoNewLine
try {
	$NewTestFile = Add-PnPFile -Path "C:\Users\mike.harley\Documents\PSAS\cpais-import.csv" -Folder $ListUrl
	Write-Host -Foregroundcolor Green "Complete."
}
catch {
	$err = $_
	Write-Host -Foregroundcolor Red $($err.Exception.Message)
}
#endregion

#region the big finish
Write-Host
Write-Host "Initialization of SPO site " -NoNewLine
Write-Host -Foregroundcolor Black -Backgroundcolor Green "Complete"
Set-ExecutionPolicy Restricted -Scope CurrentUser -Force
#endregion


