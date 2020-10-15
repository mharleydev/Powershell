Function Initialize-UIDVariables {
    <#
    .SYNOPSIS
    Initializes semantic variables based on a SharePoint list item.

    .EXAMPLE
    Initialize-UIDVariables $ChangeRequest
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$True,ValueFromPipeline=$True)][Microsoft.SharePoint.Client.SecurableObject]$item
    )
    begin {}
    process {
        Set-Variable -Name "SPRecordID" -Value $item.Id -Scope 'Global'
        Set-Variable -Name "CreatedByEmail" -Value $item.FieldValues.Author.Email -Scope 'Global'
        Set-Variable -Name "CreatedByDisplayName" -Value $item.FieldValues.Author.LookupValue -Scope 'Global'
        Set-Variable -Name "RequesterDisplayName" -Value $item.FieldValues.Editor.LookupValue -Scope 'Global'
        Set-Variable -Name "RequesterEmail" -Value $item.FieldValues.Editor.Email -Scope 'Global'

        if ($item.FieldValues.Title) {Set-Variable -Name "TeammateName" -Value $item.FieldValues.Title -Scope 'Global'}
        if ($item.FieldValues.FirstName1) {Set-Variable -Name "FirstName" -Value $item.FieldValues.FirstName1 -Scope 'Global'}
        if ($item.FieldValues.LastName) {Set-Variable -Name "LastName" -Value $item.FieldValues.LastName -Scope 'Global'}
        if ($item.FieldValues.JobTitle1) {Set-Variable -Name "JobTitle" -Value $item.FieldValues.JobTitle1 -Scope 'Global'}
        if ($item.FieldValues._x0055_ID1) {Set-Variable -Name "UID" -Value $item.FieldValues._x0055_ID1 -Scope 'Global'}
        if ($item.FieldValues.PositionType) {Set-Variable -Name "PositionType" -Value $item.FieldValues.PositionType -Scope 'Global'}
        if ($item.FieldValues.StartDate1) {Set-Variable -Name "StartDate" -Value $item.FieldValues.StartDate1.ToString('MM/dd/yyyy') -Scope 'Global'}
        if ($item.FieldValues.Manager) {
            Set-Variable -Name "ManagerDisplayName" -Value $item.FieldValues.Manager.LookupValue -Scope 'Global'
            Set-Variable -Name "ManagerEmail" -Value $item.FieldValues.Manager.Email -Scope 'Global'
        }
        if ($item.FieldValues.Author) {
            Set-Variable -Name "CreatedByEmail" -Value $item.FieldValues.Author.Email -Scope 'Global'
            Set-Variable -Name "CreatedByDisplayName" -Value $item.FieldValues.Author.Email -Scope 'Global'
        }
        if ($item.FieldValues.HRManagerGroup) {Set-Variable -Name "HRManagerGroupName" -Value $item.FieldValues.HRManagerGroup.LookupValue -Scope 'Global'}
        if ($item.FieldValues.bs_Company) {Set-Variable -Name "Company" -Value $item.FieldValues.bs_Company -Scope 'Global'}
        if ($item.FieldValues.PhoneNumber) {Set-Variable -Name "PhoneNumber" -Value $item.FieldValues.PhoneNumber -Scope 'Global'}
        if ($item.FieldValues.ol_Department) {Set-Variable -Name "Department" -Value $item.FieldValues.ol_Department -Scope 'Global'}
        if ($item.FieldValues.Address1) {Set-Variable -Name "Address" -Value $item.FieldValues.Address1 -Scope 'Global'}
        if ($item.FieldValues.WorkCity) {Set-Variable -Name "City" -Value $item.FieldValues.WorkCity -Scope 'Global'}
        if ($item.FieldValues.WorkState) {Set-Variable -Name "State" -Value $item.FieldValues.WorkState -Scope 'Global'}
        if ($item.FieldValues.WorkZip) {Set-Variable -Name "Zip" -Value $item.FieldValues.WorkZip -Scope 'Global'}
        if ($item.FieldValues.Country) {Set-Variable -Name "Country" -Value $item.FieldValues.Country -Scope 'Global'}
        if ($item.FieldValues.ManagerSubmitted) {Set-Variable -Name "ManagerSubmitted" -Value $item.FieldValues.ManagerSubmitted -Scope 'Global'}
        if ($item.FieldValues.HRManagerSubmitted) {Set-Variable -Name "HRManagerSubmitted" -Value $item.FieldValues.HRManagerSubmitted -Scope 'Global'}
        if ($item.FieldValues.ManagerUID) {Set-Variable -Name "ManagerUID" -Value $item.FieldValues.ManagerUID -Scope 'Global'}
        if ($item.FieldValues.TermDate) {Set-Variable -Name "TermDate" -Value $item.FieldValues.TermDate.ToString('MM/dd/yyyy') -Scope 'Global'}
    }
    end {}
}