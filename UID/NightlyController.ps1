#region bootstrap
<#
.DESCRIPTION
  The nightly processing script. Is scheduled to run nightly at 2am eastern. Works in conjunction with 
  UIDController.ps1 to run the Identity Warehouse.
#>

# Initialisations
Set-Location -Path (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)
$Settings = Get-Content 'settings.json' | ConvertFrom-Json

$ProcessName    = 'NightlyController'
$ScriptVersion  = '0.0.1'
$ScriptDir      = $Settings.ScriptDir
$RootScriptDir  = $Settings.RootScriptDir
$ModulesDir     = "$RootScriptDir\Modules"
$LogDir         = "$ScriptDir\Logs"
$TempDir        = "$ScriptDir\Temp"
$CredentialFile = "$ScriptDir\Credentials\myCred_${env:USERNAME}_${env:COMPUTERNAME}.xml"

# Import Modules & Snap-ins
Import-Module "$ModulesDir\SharePointPnPPowerShell2016" -Force -DisableNameChecking
Import-Module ".\Modules\Initialize-UIDVariables"
Import-Module ".\Modules\PSLogging"

$IDMSiteUrl = $Settings.IDMSiteUrl
$UIDRecordsListName = $Settings.UIDRecordsListName

#deactivation requests
$TermRecordFullPath = Join-Path -Path $TempDir -ChildPath ('IDW_TERMS_' + (Get-Date -f yyyyMMddhhmmss) + '.CSV')
$FeedsSiteUrl = "$IDMSiteUrl/feeds/"
$HRTermsFeedLibraryName = 'identwarehousefeed'
#endregion

#region connect to idm site
$LogFileName = "PSLogging_$ProcessName" + "_" + (Get-Date -f yyyyMMddhhmmss) + '.TXT'
$LogFileFullPath = Join-Path -Path $LogDir -ChildPath $LogFileName
Start-Log -LogPath $LogDir -LogName $LogFileName -ScriptVersion $ScriptVersion
Write-LogInfo -LogPath $LogFileFullPath -Message "Running $($ProcessName)"
try {
    $Credential = Import-CliXml -Path $CredentialFile
    Connect-PnPOnline -URL $IDMSiteUrl -Credential $Credential -ErrorAction 'Stop'
}
catch {
    $err = $_
    Write-LogInfo -LogPath $LogFileFullPath -Message "Error connecting to SharePoint site."
    Write-LogInfo -LogPath $LogFileFullPath -Message $($err.Exception.Message)
    stop
}
#endregion

#region expired teammates
$DateParams = @{
    Hour	= 0
    Minute 	= 0
    Second 	= 0
}
$Today = (Get-Date @DateParams)
$ExpiredTeammates = (
    Get-PnPListItem -List $UIDRecordsListName | 
    Where-Object {$_.FieldValues.EmploymentStatus -eq 'Active' -and $_.FieldValues.EndDate1 -lt $Today -and ($_.FieldValues.EndDate1)} -ErrorAction 'Stop'
)
Write-LogInfo -LogPath $LogFileFullPath -Message "Found $(($ExpiredTeammates | Measure-Object).Count) expired teammate(s)."
if ($ExpiredTeammates) {
    $TeammateRecords = $ExpiredTeammates
    foreach ($TeammateRecord in $TeammateRecords) {
        try {
            $Precontent = "<p>Deactivation requested by System. EndDate reached.</p>"
            $Content = @(
                "<b>Teammate</b>: $($TeammateRecord.FieldValues.Title)<br>"
                "<b>UID</b>: $($TeammateRecord.FieldValues._x0055_ID1)<br>"
                "<b>Job Title</b>: $($TeammateRecord.FieldValues.JobTitle1)<br>"
                "<b>Position Type</b>: $($TeammateRecord.FieldValues.PositionType)<br>"
                "<b>Start Date</b>: $($TeammateRecord.FieldValues.StartDate1.ToString("MM/dd/yyyy"))<br>"
                "<b>End Date</b>: $($TeammateRecord.FieldValues.EndDate1.ToString("MM/dd/yyyy"))<br>"
                "<b>Manager</b>: $($TeammateRecord.FieldValues.Manager.LookupValue)<br>"
            )
            $Postcontent = "<p>Identity Access & Management Team</p>"
            [System.Collections.ArrayList]$To = @()
            $To.Add($TeammateRecord.FieldValues.Manager.Email) | Out-Null
            $HRManagerGroup = Get-PnPGroupMembers -Identity $TeammateRecord.FieldValues.HRManagerGroup.LookupValue
            foreach ($HRManager in $HRManagerGroup) {
                $To.Add($HrManager.Email) | Out-Null
            }
            #$To.Add('harleymichael@bfusa.com') | Out-Null
            try {
                $sendEmailParams = @{
                    From		= 'DoNotReply@Bfusa.com'
                    To 			= $To
                    Cc          = 'harleymichael@bfusa.com'
                    Subject 	= "INFO: Deactivation request for $($TeammateRecord.FieldValues.Title)"
                    Body 		= "$Precontent $Content $Postcontent"
                    smtpServer 	= 'akmailedge.bfusa.com'
                    BodyAsHTML  = $true
                }
                Send-MailMessage @sendEmailParams
                Write-LogInfo -LogPath $LogFileFullPath -Message "Success: Sent deactivation notification to HR Managers"
            }
            catch {
                $err = $_
                Write-LogInfo -LogPath $LogFileFullPath -Message "Error: Unable to send deactivation notification to HR Managers: $($err.Exception.Message)"
                continue
            }
            $approvalHistory = $(
                "[$Today] - Deactivation requested by System. EndDate reached.`r`n`r`n" +
                $TeammateRecord.FieldValues.approvalHistory
            )
            try {
                $SetPnPListItemParams = @{
                    'approvalHistory'     = $approvalHistory
                    'RequestStatus'     = 'DeactivationAcknowledged'
                }
                Set-PnPListItem -List $UIDRecordsListName -Identity $TeammateRecord.Id -Values $SetPnPListItemParams | Out-Null
                Write-LogInfo -LogPath $LogFileFullPath -Message "Success: SetPnPListItem for $($TeammateRecord.Id)"
            }
            catch {
                $err = $_
                Write-LogInfo -LogPath $LogFileFullPath -Message "Error: Unable to SetPnPListItem for $($TeammateRecord.Id)"
                Write-LogInfo -LogPath $LogFileFullPath -Message ""
                continue
            }
        }
        catch {
            $err = $_
            Write-LogInfo -LogPath $LogFileFullPath -Message "Error marking $($TeammateRecord.Id) as deactivated."
            Write-LogInfo -LogPath $LogFileFullPath -Message $($err.Exception.Message)
            continue
        }
        
    }
}
#endregion

#region send expiring teammates notification
$OffSetTimeFrames = @(
    1,
    7,
    14,
    21
)
foreach ($OffSetTimeFrame in $OffSetTimeFrames) {
    $DateParams = @{
        Hour	= 0
        Minute 	= 0
        Second 	= 0
    }
    $StartDateFilter = (Get-Date @DateParams).AddDays($OffSetTimeFrame)
    $EndDateFilter = ((Get-Date @DateParams).AddDays(1).AddSeconds(-1)).AddDays($OffSetTimeFrame)
    $AllExpiringTeammates = (
        Get-PnPListItem -List $UIDRecordsListName | 
        Where-Object {$_.FieldValues.EndDate1 -gt $StartDateFilter -and $_.FieldValues.EndDate1 -lt $EndDateFilter}
    )
    Write-LogInfo -LogPath $LogFileFullPath -Message "Found $(($AllExpiringTeammates | Measure-Object).Count) teammates expiring in $($OffSetTimeFrame) days, $($StartDateFilter.ToString("MM/dd/yyyy"))"
    if ($AllExpiringTeammates) {
        $ExpiringTeammatesUrl = 'https://on.contoso.com/sites/ims/_layouts/15/start.aspx#/Lists/uid_records/EditDatasheet.aspx'
        # Iterate through the selections if there are multple locations returned to group email messages by location
        $AllReturnedLocations = $AllExpiringTeammates.FieldValues.bs_Company | Select-Object -Unique
        foreach ($Location in $AllReturnedLocations) {
            # get the hr managers group member emails
            $HRManagersGroupName = (Get-PnPListItem -List 'Company Master' | Where-Object {$_.FieldValues.Title -eq $Location}).FieldValues.AssociatedGroupName
            $HRManagers = Get-PnPGroupMembers -Identity $HRManagersGroupName
            [System.Collections.ArrayList]$HRManagersEmailAddresses = @()
            foreach ($HRManager in $HRManagers) {
                $HRManagersEmailAddresses.Add($HRManager.Email) | Out-Null
            }
            # setup email variables
            $Precontent = (
            "<p>This is a courtesy notification from the Identity Access & Management Team concerning the Identity Management Warehouse. The following teammates have an " +
            "end date of $($StartDateFilter.ToString('MM/dd/yyyy')). " +
            "If this is correct then no action is required by you and these teammate accounts will be disabled on $($StartDateFilter.AddDays(1).ToString('MM/dd/yyyy')).</p>" +
            "<p>If you need to change the End Date for these teammates, please click <a href=$($ExpiringTeammatesUrl)>here</a> to modify the records in question. " +
            "There's no need to use the Update link on the form as this information is not syned to Active Directory.</p>"
            )
            $ExpiringTeammates = $AllExpiringTeammates | Where-Object {$_.FieldValues.bs_Company -eq $Location}
            $Body = $Null
            foreach ($ExpiringTeammate in $ExpiringTeammates) {
                $Content = @(
                    "<p><b>Teammate</b>: $($ExpiringTeammate.FieldValues.Title)<br>"
                    "<b>UID</b>: $($ExpiringTeammate.FieldValues._x0055_ID1)<br>"
                    "<b>Job Title</b>: $($ExpiringTeammate.FieldValues.JobTitle1)<br>"
                    "<b>Position Type</b>: $($ExpiringTeammate.FieldValues.PositionType)<br>"
                    "<b>Start Date</b>: $($ExpiringTeammate.FieldValues.StartDate1.ToString("MM/dd/yyyy"))<br>"
                    "<b>End Date</b>: $($ExpiringTeammate.FieldValues.EndDate1.ToString("MM/dd/yyyy"))</p>"
                )
                $Body += $Content
            }
            $Postcontent = "<p>Regards,<br>Identity Access & Management Team</p>"
            $sendEmailParams = @{
                From		= 'DoNotReply@Bfusa.com'
                To 			= $HRManagersEmailAddresses
                Cc          = 'harleymichael@bfusa.com'
                Subject 	= "INFO: Expiring Identity Warehouse Teammates for $($Location)"
                Body 		= "$Precontent $Body $Postcontent"
                smtpServer 	= 'akmailedge.bfusa.com'
                BodyAsHTML  = $true
            }
            Send-MailMessage @sendEmailParams
        }
    }
}
#endregion

#region deactivation requests
$DeactivationRequests = Get-PnPListItem -List $UIDRecordsListName | Where-Object {$_.FieldValues.RequestStatus -eq 'DeactivationAcknowledged'}
Write-LogInfo -LogPath $LogFileFullPath -Message "Found $(($DeactivationRequests | Measure-Object).Count) deactivation requests."
if ($DeactivationRequests) {
    $TeammateRecords = $DeactivationRequests
    foreach ($TeammateRecord in $TeammateRecords) {
        Initialize-UIDVariables $TeammateRecord
        Write-LogInfo -LogPath $LogFileFullPath -Message "Processing $($UID) $($TeammateName)"
        # update the term record in sharepoint
        try {
            if ([string]::IsNullOrWhiteSpace($StartDate)) {
                $StartDate = $Null
            }
            if ([string]::IsNullOrWhiteSpace($EndDate)) {
                $EndDate = $Null
            }
            $Today = $(Get-Date).ToString('MM/dd/yyyy')
            $approvalHistory = $(
                "[$Today] - Processing deactivation request. Resetting fields.`r`n" +
                "--> Setting Manager to Null, was $($ManagerDisplayName)`r`n" +
                "--> Setting ManagerEmail to Null, was $($ManagerEmail)`r`n" +
                "--> Setting ManagerSubmitted to Null, was $($ManagerSubmitted)`r`n" +
                "--> Setting ManagerUID to Null, was $($ManagerUID)`r`n" +
                "--> Setting StartDate to Null, was $StartDate`r`n" +
                "--> Setting EndDate to Null, was $EndDate`r`n" +
                "--> Setting TermDate to $Today`r`n" +
                "--> Setting Comments to Null, was $($TeammateRecord.FieldValues._Comments)`r`n`r`n" +
                $TeammateRecord.FieldValues.approvalHistory
            )
            $editFormUrl = 'https://on.contoso.com/sites/ims/Lists/uid_records/IT%20UID%20Record/editifs.aspx?'
            $listID = 'List=4dec276e-9b41-4c5f-b003-8086ef0c52d4&'
            $sourceUrl = '&Source=https%3a%2f%2fon.contoso.com%2fsites%2fims%2fSitePages%2fuid-manager-dashboard.aspx&ID='
            $defaultView = '&DefaultView=Reactivate'
            $termActionURL = $editFormUrl + $listID + $sourceUrl + $SPRecordID + $defaultView
            
            $SetPnPListItemParams = @{
                'termAction'        = "$termActionURL , Reactivate"
                'RequestStatus'     = 'Deactivated'
                'EmploymentStatus'  = 'Inactive'
                'TermDate'          = $Today
                'approvalHistory'   = $approvalHistory
                'Manager'           = $Null
                'ManagerName'       = $Null
                'ManagerEmail'      = $Null
                'ManagerSubmitted'  = $Null
                'ManagerUID'        = $Null
                'StartDate1'         = $Null
                'EndDate1'           = $Null
                '_Comments'          = $Null
            }
            Set-PnPListItem -List $UIDRecordsListName -Identity $SPRecordID -Values $SetPnPListItemParams | Out-Null
            Write-LogInfo -LogPath $LogFileFullPath -Message "--> SharePoint list item udpated successfully." 
        }
        catch {
            $err = $_
            Write-LogInfo -LogPath $LogFileFullPath -Message "Error updating list itme." 
            Write-LogInfo -LogPath $LogFileFullPath -Message $($err.Exception.Message) 
            continue
        }

        # reset record permissions in sharepoint
        try {
            Set-PnPListItemPermission -List $UIDRecordsListName -Identity $SPRecordID -Group 'uid_Approvers' -AddRole 'Full Control' -ClearExisting
            Set-PnPListItemPermission -List $UIDRecordsListName -Identity $SPRecordID -Group $HRManagerGroupName -AddRole 'Contribute'
            Write-LogInfo -LogPath $LogFileFullPath -Message "--> SharePoint list item permissions successfully reset." 
        }
        catch {
            $err = $_
            Write-LogInfo -LogPath $LogFileFullPath -Message "Error resetting permissions." 
            Write-LogInfo -LogPath $LogFileFullPath -Message $($err.Exception.Message) 
            continue
        }

        # save all term records to WorkDir
        try {
            $object = New-Object PSObject -Property @{
                'TeammateUID'       = $UID
                'TeammateName'      = $TeammateName
                'TeammateEmail'     = $Email
                'TeammateLocation'  = $Company
                'ManagerName'       = $ManagerDisplayName
                'ManagerEmail'      = $ManagerEmail
                'ManagerUID'        = $ManagerUID
                'TermDate'          = $Today
            }
            Export-CSV -Path $TermRecordFullPath -InputObject $object -NoTypeInformation -Append
            Write-LogInfo -LogPath $LogFileFullPath -Message "--> Record successfully exported"
        }
        catch {
            $err = $_
            Write-LogInfo -LogPath $LogFileFullPath -Message "Error: $($err.Exception.Message)"
            continue
        }
        Write-LogInfo -LogPath $LogFileFullPath -Message "--> Processing for $($UID) $($TeammateName) complete"
    }

    # save the term records csv to the hr feeds offboarding site
    try {
        Connect-PnPOnline -URL $FeedsSiteUrl -Credential $Credential
        Add-PnPFile -Path $TermRecordFullPath -Folder $HRTermsFeedLibraryName | Out-Null
        Write-LogInfo -LogPath $LogFileFullPath -Message "$($TermRecordFullPath) successfully uploaded to $($HRTermsFeedLibraryName)." 
    }
    catch {
        $err = $_
        Write-LogInfo -LogPath $LogFileFullPath -Message "Error uploading the term file to the feeds library." 
        Write-LogInfo -LogPath $LogFileFullPath -Message $($err.Exception.Message) 
        continue
    }
}
#endregion

#region the big finish
Stop-Log -LogPath $LogFileFullPath
$SendLogParams = @{
    SMTPServer      = 'smtp.domain.local'
    LogPath         = $LogFileFullPath
    EmailFrom       = 'sysaccount@domain.local'
    EmailTo         = 'user@domain.local'
    EmailSubject    = "PSLogging: $($ProcessName) $(Get-Date -f yyyyMMddhhmmss)"
}
Send-Log @SendLogParams
#endregion

