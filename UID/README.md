# Introduction 
This set of scripts was origionally published as part of an internal project. These scripts have been santized of any internally identifiable information and is being published here as part of my portfolio. 

# Getting Started
Follow these steps to get everything up and running on a new server. 

## Prerequisties

* You'll need to have [Git for Windows](https://git-scm.com/download/win) installed on the server.
* You'll need to have the ActiveRoles Management Shell and the ADSI helper application installed on the server.
* You'll need access to the [Project](https://company@dev.azure.com/business-unit/project/) on the [Contoso Azure DevOps](https://dev.azure.com/Contoso/) instance.
* At the time of development, this process was running on SharePoint 2016 farm. As such, Initialize-Controllers.ps1 looks for and downloads the SharePoint 2016 [PnP Module](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps) modules. If or when the site is migrated to SharePoint Online, the script will need to be changed to use the SharePoint Online module instead.

## Clone the PowerShell scripots to the local server.
The PowerShell scripts need a place to live. Copy and paste the commands below into a new PowerShell command prompt. 

_Note_: To run PowerShell with elevated privlages, right click your PowerShell shortcut and select _Run as administrator_.

```powershell
$LocalDir = 'D:\PSScripts\UID' 
git clone https://company@dev.azure.com/business-unit/project/_git/repo $LocalDir
```

## Initialize the Identity Warehouse Processes
Copy and paste this PowerShell snippet into the existing Powershell window:

```powershell
Set-Location $LocalDir
.\Initialize-Controllers.ps1
```
The following items are completed during this step.

- Creates the following directories if they do not exist.
  - ..\Modules
  - .\Credentials
  - .\Logs
  - .\Temp
- Downloads the SharePoint 2016 [PnP Module](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps) into ..\Modules if it is not already downloaded. If it's already present, then the script checks to see if the module needs to be updated.
- Prompt for the service account credentials and securely save the encrypted credentials for use by the Task Scheduler jobs.
- Setup the UIDController and UIDNightlyController Task Scheduler jobs.

# File descriptions
 - __defaults.json__ - These are the default settings that the process starts with.
 - __Initialize-Controllers.ps1__ - Run this script to initalize the UID processes.
 - __Initialize-CredentialFile.ps1__ - This script is used along with Initialize-Controllers.ps1 to setup the Identity Warehouse processes.
 - __Initialize-Defaults.ps1__ - This script is used to create the defaults.json file. Edit the script then run it to generate an updated defaults.json.
 - __UIDController.ps1__ - The main processing script. Is scheduled to run every 5 minutes. Works in conjunction with UIDNightlyController.ps1__ to run the Identity Warehouse.
 - __UIDNightlyController.ps1__ - The nightly processing script. Is scheduled to run nightly at 2am eastern. Works in conjunction with UIDController.ps1 to run the Identity Warehouse.