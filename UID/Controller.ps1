#region bootstrap
<#
.DESCRIPTION
  The main processing script. Is scheduled to run every 5 minutes. Works in conjunction 
  with NightlyController.ps1 to run the project.
#>
param (
  #Script parameters go here
)

# Initialisations
Set-Location -Path (Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)
$Settings = Get-Content 'settings.json' | ConvertFrom-Json

$ProcessName    = 'Controller'
$ScriptVersion  = '0.0.1'
$ScriptDir      = $Settings.ScriptDir
$RootScriptDir  = $Settings.RootScriptDir
$ModulesDir     = "$RootScriptDir\Modules"
$LogDir         = "$ScriptDir\Logs"
$CredentialFile = "$ScriptDir\Credentials\myCred_${env:USERNAME}_${env:COMPUTERNAME}.xml"
$ARSServerName  = $Settings.ARSServerName

#Import Modules & Snap-ins
Import-Module "$ModulesDir\SharePointPnPPowerShell2016" -Force -DisableNameChecking
Import-Module 'ActiveRolesManagementShell' -Force -DisableNameChecking
Import-Module ".\Modules\Initialize-UIDVariables"
Import-Module ".\Modules\PSLogging"

$IDMSiteUrl = $Settings.IDMSiteUrl
$UIDRecordsListName = $Settings.UIDRecordsListName
$UIDMasterListName = $Settings.UIDMasterListName
$CompanyMasterListName = 'Company Master'
#endregion

#region connect to idm site
try {
    $Credential = Import-CliXml -Path $CredentialFile
    Connect-PnPOnline -URL $IDMSiteUrl -Credential $Credential
}
catch {
    $err = $_
    #Send-CrashReport -ProcessName $ProcessName -LogMessage "Error running Connect-PnPOnline command."
    continue
}
#endregion

#region new company master requests
$NewCompanyMasterRequest = Get-PnPListItem -List $CompanyMasterListName | Where-Object {$_.FieldValues.Status -eq 'New'} | Select-Object -First 1
if ($NewCompanyMasterRequest) {
  if (!$LogFileName) {
    $LogFileName = "PSLogging_$ProcessName" + "_" + (Get-Date -f yyyyMMddhhmmss) + '.TXT'
    $LogFileFullPath = Join-Path -Path $LogDir -ChildPath $LogFileName
    Start-Log -LogPath $LogDir -LogName $LogFileName -ScriptVersion $ScriptVersion
    Write-LogInfo -LogPath $LogFileFullPath -Message "Running $($ProcessName)"
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Processing New Company Master Request, $($NewCompanyMasterRequest.FieldValues.Title)"
  # create the new group
  $NewHRManagerGroupName = "HR Managers $($NewCompanyMasterRequest.FieldValues.Title)"
  try {
    $NewPnPGroupParams = @{
      Title = $NewHRManagerGroupName
      Owner = 'Owners'
    }
    New-PnPGroup @NewPnPGroupParams | Out-Null
    Write-LogInfo -LogPath $LogFileFullPath -Message "--> Created $($NewHRManagerGroupName)"
  }
  catch {
    $err = $_
    Write-LogInfo -LogPath $LogFileFullPath -Message "Error creating $($NewHRManagerGroupName): $($err.Exception.Message)"
    continue
  }

  # if there are any HR managers identified, add them to the new group
  if ($NewCompanyMasterRequest.FieldValues.HRManagers) {
    Write-LogInfo -LogPath $LogFileFullPath -Message "--> Found $(($NewCompanyMasterRequest.FieldValues.HRManagers | Measure-Object).Count) HR Managers"
    foreach ($HRManager in $NewCompanyMasterRequest.FieldValues.HRManagers) {
      try {
        Add-PnPUserToGroup -LoginName $HRManager.Email -Identity $NewHRManagerGroupName
        Write-LogInfo -LogPath $LogFileFullPath -Message "--> Added $($HRManager.LookupValue) to $($NewHRManagerGroupName)"
      }
      catch {
        $err = $_
        Write-LogInfo -LogPath $LogFileFullPath -Message "Error adding $($HRManager.LookupValue) to $($NewHRManagerGroupName): $($err.Exception.Message)"
        continue
      }
    }
  }

  # give the new hr managers group read permissions to the site
  try {
    Set-PnPGroupPermissions -Identity $NewHRManagerGroupName -AddRole 'Read'
    Write-LogInfo -LogPath $LogFileFullPath -Message "--> Added $($NewHRManagerGroupName) with Read permissions"
  }
  catch {
    $err = $_
    Write-LogInfo -LogPath $LogFileFullPath -Message "Error adding $($NewHRManagerGroupName) with Read permissions: $($err.Exception.Message)"
    continue
  }

  # write the changes back to the list item
  try {
    $SetPnPListItemParams = @{
      Status              = 'Active'
      AssociatedGroupName = $NewHRManagerGroupName
    }
    Set-PnPListItem -List $CompanyMasterListName -Identity $NewCompanyMasterRequest.Id -Values $SetPnPListItemParams | Out-Null
  }
  catch {
    $err = $_
    Write-LogInfo -LogPath $LogFileFullPath -Message "Error: $($err.Exception.Message)"
    continue
  }

  # send email to the request
  try {
    $To = $NewCompanyMasterRequest.FieldValues.Author.Email
    $EmailPrecontent = (
      "<p>This is a courtesy notification from the Identity Access & Management Team concerning the Identity Management Warehouse.</p>" +
      "<p>Your new Company Request, $($NewCompanyMasterRequest.FieldValues.Title) has been processed and is ready to use.</p>"
    )
    $EmailSignature = "<p>Regards,<br>Department Name</p>"
    $sendEmailParams = @{
      From		    = 'DoNotReply@domain.local'
      To 			    = $To
      Cc          = 'email@domain.local'
      Subject 	  = "INFO: Processing New Company Master Request, $($NewCompanyMasterRequest.FieldValues.Title) complete"
      Body 		    =  "$EmailPrecontent $EmailSignature"
      smtpServer 	= 'mail.domain.local'
      BodyAsHTML  = $true
  }
  Send-MailMessage @sendEmailParams
  Write-LogInfo -LogPath $LogFileFullPath -Message "--> Email notification sent to $($To)"
  }
  catch {
    $err = $_
    Write-LogInfo -LogPath $LogFileFullPath -Message "Error sending email notification to $($To): $($err.Exception.Message)"
    continue
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Processing New Company Master Request, $($NewCompanyMasterRequest.FieldValues.Title) complete"
}
#endregion

#region new requests
$NewRequest = Get-PnPListItem -List $UIDRecordsListName | Where-Object {$_.FieldValues.RequestStatus -eq 'New'} | Select-Object -First 1
if ($NewRequest) {
  $TeammateRecord = $NewRequest
  Initialize-UIDVariables $TeammateRecord
  if (!$LogFileName) {
    $LogFileName = "PSLogging_$ProcessName" + "_" + (Get-Date -f yyyyMMddhhmmss) + '.TXT'
    $LogFileFullPath = Join-Path -Path $LogDir -ChildPath $LogFileName
    Start-Log -LogPath $LogDir -LogName $LogFileName -ScriptVersion $ScriptVersion
    Write-LogInfo -LogPath $LogFileFullPath -Message "Running $($ProcessName)"
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Found $(($TeammateRecord | Measure-Object).Count) new requests"
  Write-LogInfo -LogPath $LogFileFullPath -Message "Processing $($SPRecordID)"
  $editFormUrl = 'https://contoso.com/sites/ims/Lists/list_url/editifs.aspx?'
  $listID = 'List=4dec276e-9b41-4c5f-b003-8086ef0c52d4&'
  $sourceUrl = '&Source=https%3a%2f%2fon.bsaconnect.com%2fsites%2fims%2fSitePages%2fuid-manager-dashboard.aspx&ID='
  $defaultView = '&DefaultView=Deactivate'
  $defaultUpdateInfoView = '&DefaultView=UpdateInfo'
  $termActionUrl = $editFormUrl + $listID + $sourceUrl + $SPRecordID + $defaultView 
  $updateActionURL = $editFormUrl + $listID + $sourceUrl + $SPRecordID + $defaultUpdateInfoView

  # determine if this is a manager submitted record
  $ManagerSubmitted = 'False'
  if ($ManagerEmail -eq $CreatedByEmail) {
    $ManagerSubmitted = 'True'
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "--> Manager submitted: $($ManagerSubmitted)"

  # determine if this is a hr manager submitted record
  $HRManagerSubmitted = 'False'
  $HRManagerGroupName = "HR Managers $($Company)"
  $HRManagerGroup = Get-PnPGroup -Identity $HRManagerGroupName 
  [string]$HRManagerGroupID = $HRManagerGroup.Id
  $HRManagerGroupMembers = Get-PnPGroupMembers -Identity $HRManagerGroupName
  foreach ($HRManager in $HRManagerGroupMembers) {
    if ($HRManager.email -eq $CreatedByEmail) {
      $HRManagerSubmitted = 'True'
    }
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "--> HR Manager submitted: $($HRManagerSubmitted)"

  # get manager UID from AD
  $manager = Get-ADUser -filter {(enabled -eq $true) -and (Employeenumber -like "*") -and (mail -eq $ManagerEmail)} -Properties ('mail', 'EmployeeNumber')

  # format approval history
  $Today = (Get-Date).ToString('MM/dd/yyyy')
  $approvalHistory = $(
    "[$Today] - Request submitted by $($CreatedByDisplayName)"
  )
  $SetPnPListItemParams = @{
    Title                 = $LastName + ', ' + $FirstName
    _x0055_ID1            = '00000000'
    ManagerSubmitted      = $ManagerSubmitted
    HRManagerSubmitted    = $HRManagerSubmitted
    ManagerName           = $ManagerDisplayName
    ManagerEmail          = $ManagerEmail
    ManagerUID            = $manager.EmployeeNumber
    RequesterDisplayName  = $CreatedByDisplayName
    RequestStatus         = 'Pending IT Review'
    EmploymentStatus      = 'Pending'
    HRManagerGroup        = $HRManagerGroupID
    approvalHistory       = $approvalHistory
    termAction            = "$termActionURL , Deactivate"
    updateAction          = "$updateActionURL , Update"
  } 
  try {
    Set-PnPListItem -List $UIDRecordsListName -Identity $SPRecordID -Values $SetPnPListItemParams | Out-Null
    Write-LogInfo -LogPath $LogFileFullPath -Message "--> Successfully updated $($SPRecordID) on SharePoint"
  }
  catch {
    $err = $_
    Write-LogInfo -LogPath $sLogFile -Message "--> Error updating $($SPRecordID) SharePoint record: $($err.Exception.Message)"
    continue
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "New requests processing complete"
}
#endregion

#region hr manager approved
$managerApprovedUIDRequests = (
    Get-PnPListItem -List $UIDRecordsListName | 
    Where-Object {$_.FieldValues.RequestStatus -eq 'HR Manager Approved'} | 
    Select-Object -First 1
)
if ($managerApprovedUIDRequests) {
  $TeammateRecords = $managerApprovedUIDRequests
  if (!$LogFileName) {
    $LogFileName = "PSLogging_$ProcessName" + "_" + (Get-Date -f yyyyMMddhhmmss) + '.TXT'
    $LogFileFullPath = Join-Path -Path $LogDir -ChildPath $LogFileName
    Start-Log -LogPath $LogDir -LogName $LogFileName -ScriptVersion $ScriptVersion
    Write-LogInfo -LogPath $LogFileFullPath -Message "Running $($ProcessName)"
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Found $(($TeammateRecords | Measure-Object).Count) HR Manager Approved requests"
	foreach ($TeammateRecord in $TeammateRecords) {
    Initialize-UIDVariables $TeammateRecord

    Write-LogInfo -LogPath $LogFileFullPath -Message "Processing $($TeammateName)"
    #create new UID item in master list
    try {
      $NewUIDItemTitle = (
        $SPRecordID.ToString() +
        $LastName +
        $FirstName
      )
      $NewUIDItem = @{
        Title = $NewUIDItemTitle
      }
      $NewUIDItem = Add-PnPListItem -List $UIDMasterListName -Values $NewUIDItem
      $NewUID = 'IT' + $("{0:000000}" -f $NewUIDItem.Id).ToString()
      Set-PnPListItem -List $UIDMasterListName -Identity $NewUIDItem.Id -Values @{UID = $NewUID} | Out-Null
      Write-LogInfo -LogPath $LogFileFullPath -Message "--> UID created: $($NewUID)"
    }
    catch {
      $err = $_
      Write-LogInfo -LogPath $LogFileFullPath -Message "Error: creating new UID: $($err.Exception.Message)"
      continue
    }
    #assign new UID to request
    try {
      Set-PnPListItem -List $UIDRecordsListName -Identity $SPRecordID -Values @{"RequestStatus"='UID Assigned'; "_x0055_ID1"=$NewUID} | Out-Null
      Write-LogInfo -LogPath $LogFileFullPath -Message "--> UID assigned to $($TeammateName)"
    }
    catch {
      $err = $_
      Write-LogInfo -LogPath $LogFileFullPath -Message "--> Error assigning UID to $($TeammateName): $($err.Exception.Message)"
      continue
    }
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Processing complete for $($TeammateName)"
}
#endregion

#region acknowledge deactivation requests
$DeactivationRequests = Get-PnPListItem -List $UIDRecordsListName | Where-Object {$_.FieldValues.RequestStatus -eq 'DeactivationRequested'}
if ($DeactivationRequests) {
  $StepName = "Deactivation Requests"
  $TeammateRecords = $DeactivationRequests
  if (!$LogFileName) {
    $LogFileName = "PSLogging_$ProcessName" + "_" + (Get-Date -f yyyyMMddhhmmss) + '.TXT'
    $LogFileFullPath = Join-Path -Path $LogDir -ChildPath $LogFileName
    Start-Log -LogPath $LogDir -LogName $LogFileName -ScriptVersion $ScriptVersion
    Write-LogInfo -LogPath $LogFileFullPath -Message "Running $($ProcessName)"
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Found $(($TeammateRecords | Measure-Object).Count) deactivation request(s)"
  foreach ($TeammateRecord in $TeammateRecords) {
    Initialize-UIDVariables $TeammateRecord

    Write-LogInfo -LogPath $LogFileFullPath -Message "Processing $($TeammateName) - $($UID)"
    try {
      $Precontent = "<p>A new deactivation request has been submitted by $($RequesterDisplayName).</p>"
      $Content = @(
          "<b>Teammate</b>: $($TeammateName)<br>"
          "<b>UID</b>: $($UID)<br>"
          "<b>Job Title</b>: $($JobTitle )<br>"
          "<b>Position Type</b>: $($PositionType)<br>"
          "<b>Start Date</b>: $($StartDate)<br>"
          "<b>Manager</b>: $($ManagerDisplayName)<br>"
      )
      $Postcontent = "<p>Identity Access & Management Team</p>"
      [System.Collections.ArrayList]$To = @()
      $To.Add($TeammateRecord.FieldValues.Manager.Email) | Out-Null
      $HRManagerGroup = Get-PnPGroupMembers -Identity $HRManagerGroupName
      foreach ($HRManager in $HRManagerGroup) {
          $To.Add($HrManager.Email) | Out-Null
      }
      $sendEmailParams = @{
          From		    = 'DoNotReply@Bfusa.com'
          To 			    = $To
          Cc          = 'harleymichael@bfusa.com'
          Subject 	  = "INFO: Deactivation request for $($TeammateName)"
          Body 		    = "$Precontent $Content $Postcontent"
          smtpServer 	= 'akmailedge.bfusa.com'
          BodyAsHTML  = $true
      }
      Send-MailMessage @sendEmailParams
      Write-LogInfo -LogPath $LogFileFullPath -Message "--> Notification email sent"

      $Today = $(Get-Date).ToString('MM/dd/yyyy')
      $approvalHistory = $(
          "[$Today] - Deactivation requested by $($RequesterDisplayName)`r`n`r`n" +
          $TeammateRecord.FieldValues.approvalHistory
      )
      $SetPnPListItemParams = @{
          'approvalHistory'     = $approvalHistory
          'RequestStatus'     = 'DeactivationAcknowledged'
      }
      Set-PnPListItem -List $UIDRecordsListName -Identity $SPRecordID -Values $SetPnPListItemParams | Out-Null
      Write-LogInfo -LogPath $LogFileFullPath -Message "--> SharePoint list item updated"
    }
    catch {
        $err = $_
        Write-LogInfo -LogPath $sLogFile -Message "--Error deactiving teammate: $($err.Exception.Message)"
        continue
    }
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Deactivation processing for $($TeammateName) - $($UID) is complete"
}
#endregion

#region reactivations
$ReactivationRequests = (
    Get-PnPListItem -List $UIDRecordsListName | 
    Where-Object {$_.FieldValues.RequestStatus -eq 'ReactivationRequested'} | 
    Select-Object -First 1
)
if ($ReactivationRequests) {
  $StepName = "Reactivation Requests"
  $TeammateRecords = $ReactivationRequests
  if (!$LogFileName) {
    $LogFileName = "PSLogging_$ProcessName" + "_" + (Get-Date -f yyyyMMddhhmmss) + '.TXT'
    $LogFileFullPath = Join-Path -Path $LogDir -ChildPath $LogFileName
    Start-Log -LogPath $LogDir -LogName $LogFileName -ScriptVersion $ScriptVersion
    Write-LogInfo -LogPath $LogFileFullPath -Message "Running $($ProcessName)"
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Found $(($TeammateRecords | Measure-Object).Count) reactivation request(s)"
  if (!$ARSConnection) {
    $ARSConnection = Connect-QADService -Service $ARSServerName -Credential $Credential -Proxy
    Write-LogInfo -LogPath $LogFileFullPath -Message "Active Directory connection type is $($ARSConnection.Type)"
  }
  foreach ($TeammateRecord in $TeammateRecords) {

    try {
        Initialize-UIDVariables $TeammateRecord
    }
    catch {
        $err = $_
        Write-LogInfo -LogPath $LogFileFullPath -Message "Error: $($err.Exception.Message)"
        break
    }
    
    Write-LogInfo -LogPath $LogFileFullPath -Message "Reactivating $($UID ) $($TeammateName)"
    # build term action and update action URLs
    $editFormUrl = 'https://on.bsaconnect.com/sites/ims/Lists/uid_records/IT%20UID%20Record/editifs.aspx?'
    $listID = 'List=4dec276e-9b41-4c5f-b003-8086ef0c52d4&'
    $sourceUrl = '&Source=https%3a%2f%2fon.bsaconnect.com%2fsites%2fims%2fSitePages%2fuid-manager-dashboard.aspx&ID='
    $defaultView = '&DefaultView=Deactivate'
    $defaultUpdateInfoView = '&DefaultView=UpdateInfo'
    $termActionUrl = $editFormUrl + $listID + $sourceUrl + $SPRecordID + $defaultView 
    $updateActionURL = $editFormUrl + $listID + $sourceUrl + $SPRecordID + $defaultUpdateInfoView

    # get manager UID from AD
    Write-LogInfo -LogPath $LogFileFullPath -Message "--> Getting manager Active Directory record for $($ManagerEmail)"
    [System.Collections.ArrayList]$CurrentChangesArray = @()
    try {
      $ManagerADRecord = Get-QADUser -Email $ManagerEmail -Properties *
      $Message = "--> Found $($ManagerADRecord)"
      Write-LogInfo -LogPath $LogFileFullPath -Message $Message
    }
    catch {
      $err = $_
      Write-LogInfo -LogPath $LogFileFullPath -Message "Error getting manager AD reord for $($ManagerEmail): $($err.Exception.Message)"
      continue
    }

    # assign manager permissions to list item in SharePoint
    try {
      Set-PnPListItemPermission -List $UIDRecordsListName -Identity $SPRecordID -User $ManagerDisplayName -AddRole 'Contribute'
      Write-LogInfo -LogPath $LogFileFullPath -Message "--> Successfully assigned manager $($ManagerDisplayName), Contribute permissions to list item."
    }
    catch {
      $err = $_
      Write-LogInfo -LogPath $LogFileFullPath -Message "Error assigning manager permissions: $($err.Exception.Message)"
      continue
    }

    # update sharepoint list item
    try {
        $CurrentChangesArray.Add("--> Setting Manager to $($ManagerDisplayName)") | Out-Null
        $CurrentChangesArray.Add("--> Setting ManagerEmail to $($ManagerEmail)") | Out-Null
        $CurrentChangesArray.Add("--> Setting ManagerUID to $($ManagerADRecord.employeeid)") | Out-Null
        $CurrentChangesArray.Add("--> Setting TermDate to Null. Was $($TermDate)") | Out-Null
        $CurrentChangesArray.Add("--> Setting EmploymentStatus to Active") | Out-Null

        $CurrentChangesPreContent = "[$(Get-Date -f MM/dd/yyyy)] - Reactivation requested by $($RequesterDisplayName).`r`n"
        [string]$CurrentChangesBody = @()
        foreach ($row in $CurrentChangesArray) {
            $CurrentChangesBody += "$($row)`r`n"
        }
        $CurrentChanges = $($CurrentChangesPrecontent + $CurrentChangesBody) + "`r`n"

        $SetPnPListItemValues = @{
        ManagerName       = $ManagerDisplayName
        ManagerEmail      = $ManagerEmail
        ManagerUID        = $ManagerADRecord.employeeid
        RequestStatus     = 'Complete'
        EmploymentStatus  = 'Active'
        TermDate          = $Null
        approvalHistory   = $CurrentChanges + $TeammateRecord.FieldValues.approvalHistory
        termAction        = "$termActionURL , Deactivate"
        updateAction      = "$updateActionURL , Update"
        }
        Set-PnPListItem -List $UIDRecordsListName -Identity $SPRecordID -Values $SetPnPListItemValues | Out-Null
        Write-LogInfo -LogPath $LogFileFullPath -Message "--> Successfully updated SharePoint list item for $($UID) $($TeammateName)"
    }
    catch {
      $err = $_
      Write-LogInfo -LogPath $LogFileFullPath -Message "Error updating SharePoint list item for $($UID) $($TeammateName): $($err.Exception.Message)"
      continue
    }

    # send notification to requester
    try {
        $EmailStepName = ($StepName.ToLower()).TrimEnd('s')
        $EmailPrecontent = (
            "<p>This is a courtesy notification from the Identity Access & Management Team concerning the Identity Management Warehouse.</p>" +
            "<p>Your $($EmailStepName) for $($UID) $($TeammateName) has been processed.</p>"
        )
        [string]$EmailBody = @()
        foreach ($row in $CurrentChangesArray) {
            $EmailBody += "$($row)<br>"
        }
        $EmailPostcontent = @(
            "<p><i>Note</i>: This step only reactives the record on the Identity Management Warehouse Dashboard. To complete the reactivation, you will " +
            "need to submit an onboarding request for this teammate.</p>"
        )
        $EmailSignature = "<p>Regards,<br>Identity Access & Management Team</p>"
        $sendEmailParams = @{
            From		= 'DoNotReply@Bfusa.com'
            To 			= $RequesterEmail
            Cc          = 'harleymichael@bfusa.com'
            Subject 	= "INFO: $($StepName.TrimEnd('s')) for $($UID) $($TeammateName) is complete"
            Body 		= "$EmailPrecontent $EmailBody $EmailPostcontent $EmailSignature"
            smtpServer 	= 'akmailedge.bfusa.com'
            BodyAsHTML  = $true
        }
        Send-MailMessage @sendEmailParams
        Write-LogInfo -LogPath $LogFileFullPath -Message "--> Email notification sent to $($RequesterEmail)"
    }
    catch {
        $err = $_
        Write-LogInfo -LogPath $LogFileFullPath -Message "Error sending email notification to $($RequesterEmail): $($err.Exception.Message)"
        continue
    }
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Reactivation step complete."
}
#endregion

#region change requests
$ChangeRequests = Get-PnPListItem -List $UIDRecordsListName | Where-Object {$_.FieldValues.RequestStatus -eq 'SyncRequested'}
if ($ChangeRequests) {
  $StepName = "Change Requests"
  $TeammateRecords = $ChangeRequests
  if (!$LogFileName) {
    $LogFileName = "PSLogging_$ProcessName" + "_" + (Get-Date -f yyyyMMddhhmmss) + '.TXT'
    $LogFileFullPath = Join-Path -Path $LogDir -ChildPath $LogFileName
    Start-Log -LogPath $LogDir -LogName $LogFileName -ScriptVersion $ScriptVersion
    Write-LogInfo -LogPath $LogFileFullPath -Message "Running $($ProcessName)"
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Found $(($TeammateRecords | Measure-Object).Count) change request(s)"
  if (!$ARSConnection) {
    $ARSConnection = Connect-QADService -Service $ARSServerName -Credential $Credential -Proxy
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Active Directory connection type is $($ARSConnection.Type)"
  foreach ($TeammateRecord in $TeammateRecords) {
    Initialize-UIDVariables $TeammateRecord
    Write-LogInfo -LogPath $LogFileFullPath -Message "Processing $($UID) $($TeammateName)"
    # Pull all accounts associated with this uid
    # Could include test and admin accounts
    $ADUsers = Get-QADUser -ObjectAttributes @{EmployeeNumber=$UID} -Properties *
    $ADUser = $ADUsers | Where-Object {$_.employeeType -eq 'Teammate' -or $_.employeeType -eq 'Contractor' -or $_.employeeType -eq 'International'}
    if ($ADUser) {
        Write-LogInfo -LogPath $LogFileFullPath -Message "--> Found AD record for $($UID) $($TeammateName)"
        # setup arrays and hash to compare sharepoint values to active directory values

        # values in SharePoint uid list
        $ChangeRequestObject = New-Object PSObject -Property @{
          'givenname'			  = $FirstName
          'LastName' 			  = $LastName
          'Title' 			    = $JobTitle
          'employeeType' 		= $PositionType
          'Manager' 			  = $ManagerEmail
          'telephonenumber' = $PhoneNumber
          'Department' 		  = $Department
          'StreetAddress' 	= $Address
          'City' 				    = $City
          'StateOrProvince' = $State
          'PostalCode'      = $Zip
          'co'              = $Country
          'Company'         = $Company
        }
            
        # values in AD
        if ($ADUser.Manager) {
          $ManagerADRecord = Get-QADUser -Identity $ADUser.Manager
        }
        else {
          $ManagerADRecord = $Null
        }

        $ADUserObject = New-Object PSObject -Property @{
          'givenname'			  = $adUser.givenname
          'LastName' 			  = $adUser.LastName
          'Title' 			    = $adUser.Title
          'employeeType' 		= $adUser.employeeType
          'Manager' 			  = $ManagerADRecord.mail
          'telephonenumber' = $adUser.telephonenumber
          'Department' 		  = $adUser.Department
          'StreetAddress' 	= $adUser.StreetAddress
          'City' 				    = $adUser.City
          'StateOrProvince' = $adUser.StateOrProvince
          'PostalCode'      = $adUser.PostalCode
          'co' 			        = $adUser.co
          'Company' 		    = $adUser.Company
        }

      # compare sharepoint list object to active directory object
      $UpdateParams = @{}
      [System.Collections.ArrayList]$CurrentChangesArray = @()
      $Columns = $($ChangeRequestObject.PsObject.Properties).Name
      foreach ($Column in $Columns) {
        if (!$ADUserObject.$Column) {
          if ($ChangeRequestObject.$Column) {
              $message = "--> $($Column) is Null in Active Directory. Updating to $($ChangeRequestObject.$Column)"
              Write-LogInfo -LogPath $LogFileFullPath -Message $message
              $CurrentChangesArray.Add($message) | Out-Null
              $UpdateParams.Add($Column, $($ChangeRequestObject.$Column))
              continue
          }
          else {
              #Write-LogInfo -LogPath $LogFileFullPath -Message "--> $($Column) is Null in Active Directory and SharePoint. Doing nothing"
              continue
          }
        }
        if ($ChangeRequestObject.$Column) {
          $Comparison = Compare-Object -ReferenceObject $ChangeRequestObject.$Column -DifferenceObject $ADUserObject.$Column
          if ($Comparison) {
              $message = "--> Updating $($Column) from $($ADUserObject.$Column) to $($ChangeRequestObject.$Column)"
              Write-LogInfo -LogPath $LogFileFullPath -Message $message
              $CurrentChangesArray.Add($message) | Out-Null
              $UpdateParams.Add($Column, $($ChangeRequestObject.$Column))
          }
          else {
            #Write-LogInfo -LogPath $LogFileFullPath -Message "--> $($Column) match: SharePoint: $($ChangeRequestObject.$Column). Active Directory: $($ADUserObject.$Column). Doing Nothing"
          }
        }
        else {
          $message = "--> $($Column) is Null in SharePoint. Updating from $($ADUserObject.$Column) to Null"
          Write-LogInfo -LogPath $LogFileFullPath -Message $message
          $CurrentChangesArray.Add($message) | Out-Null
          $UpdateParams.Add($Column, $Null)
        }
      }

      # update the active directory record
      if ($UpdateParams.count -gt 0) {
        try {
          if ($UpdateParams.Manager) {
            $NewManager = Get-QADUser -LdapFilter "(mail=$ManagerEmail)"
            $UpdateParams.Manager = $NewManager.DN
          }
          if ($UpdateParams.City) {
            $UpdateParams.Add('l', $($UpdateParams.City))
            $UpdateParams.Remove('City')
          }
          $ADUser | Set-QADUser -ObjectAttributes $UpdateParams | Out-Null
          Write-LogInfo -LogPath $LogFileFullPath -Message "--> Successfully updated Active Diretory record for $($UID) $($ADUser.Name)"
        }
        catch {
          $err = $_
          Write-LogInfo -LogPath $LogFileFullPath -Message "Error updating Active Diretory record in $($UID) $($ADUser.Name) AD: $($err.Exception.Message)"
          continue
        }
      }
      else {
        $Message = "--> No differences detected for $($ADUser.Name). No changes made in Active Directory."
        Write-LogInfo -LogPath $LogFileFullPath -Message $Message
        $CurrentChangesArray.Add($message) | Out-Null
      }
    }
    else {
      [System.Collections.ArrayList]$CurrentChangesArray = @()
      $message = "--> No Active Directory account found for $($UID), $($TeammateName)"
      Write-LogInfo -LogPath $LogFileFullPath -Message $message
      $CurrentChangesArray.Add($message) | Out-Null
    }

  # update sharepoint record
  try {
    $CurrentChangesPreContent = "[$(Get-Date -f MM/dd/yyyy)] - Sync requested by $($RequesterDisplayName).`r`n"
    [string]$CurrentChangesBody = @()
    foreach ($row in $CurrentChangesArray) {
      $CurrentChangesBody += "$($row)`r`n"
    }
    $CurrentChanges = $($CurrentChangesPrecontent + $CurrentChangesBody) + "`r`n"
    $SetPnPListItemParams = @{
      RequestStatus   = 'Complete'
      approvalHistory = $CurrentChanges + $TeammateRecord.FieldValues.approvalHistory
    }
    if ($UpdateParams.Manager) {
      $NewManager = Get-QADUser -LdapFilter "(mail=$ManagerEmail)" -Properties *
      $SetPnPListItemParams.Add('ManagerName', $ManagerDisplayName)
      $SetPnPListItemParams.Add('ManagerEmail', $ManagerEmail)
      $SetPnPListItemParams.Add('ManagerUID', $NewManager.employeenumber)
    }
    Set-PnPListItem -List $UIDRecordsListName -Identity $SPRecordID -Values $SetPnPListItemParams | Out-Null
    Write-LogInfo -LogPath $LogFileFullPath -Message "--> Successfully updated SharePoint list item."
  }
  catch {
    $err = $_
    Write-LogInfo -LogPath $LogFileFullPath -Message "Error: $($err.Exception.Message)"
    continue
  }

  # send notification to requester
  try {
    $EmailStepName = ($StepName.ToLower()).TrimEnd('s')
    $EmailPrecontent = (
        "<p>This is a courtesy notification from the Identity Access & Management Team concerning the Identity Management Warehouse.</p>" +
        "<p>Your $($EmailStepName) for $($UID) $($TeammateName) has been processed.</p>"
    )
    [string]$EmailBody = @()
    foreach ($row in $CurrentChangesArray) {
      $EmailBody += "$($row)<br>"
    }
    $EmailPostcontent = "<p>Regards,<br>Identity Access & Management Team</p>"
    $sendEmailParams = @{
        From		= 'DoNotReply@Bfusa.com'
        To 			= $RequesterEmail
        Cc          = 'harleymichael@bfusa.com'
        Subject 	= "INFO: $($StepName.TrimEnd('s')) for $($UID) $($TeammateName) is complete"
        Body 		= "$EmailPrecontent $EmailBody $EmailPostcontent"
        smtpServer 	= 'akmailedge.bfusa.com'
        BodyAsHTML  = $true
    }
    Send-MailMessage @sendEmailParams
    Write-LogInfo -LogPath $LogFileFullPath -Message "--> Email notification sent to $($RequesterEmail)"
  }
  catch {
    $err = $_
    Write-LogInfo -LogPath $LogFileFullPath -Message "Error sending email notification to $($RequesterEmail): $($err.Exception.Message)"
    continue
  }
  Write-LogInfo -LogPath $LogFileFullPath -Message "Processing $($UID) $($TeammateName) complete."
  }
}
#endregion

#region the big finish
if ($LogFileName) {
  Stop-Log -LogPath $LogFileFullPath
  $SendLogParams = @{
      SMTPServer      = 'smtp.domain.local'
      LogPath         = $LogFileFullPath
      EmailFrom       = 'sysaccount@domain.local'
      EmailTo         = 'account@domain.local'
      EmailSubject    = "PSLogging: $($ProcessName) $((Get-Date -f yyyyMMddhhmmss))"
  }
  Send-Log @SendLogParams
}
#endregion