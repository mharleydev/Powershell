#region bootstrap
<#
    .DESCRIPTION
    Run this script to initalize the  processes.
#>
$ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$RootScriptDir = Split-Path -Parent $ScriptDir
$StartDate = Get-Date -Format 'MM/dd/yyyy hh:mm:ss tt'
Write-Host "------------------------------------------"
Write-Host "| Initializing controllers               |"
Write-Host "| Started: $startDate        |"
Write-Host "------------------------------------------"
Write-Host
#endregion

#region create settings.json
Write-Host "Intializing the settings file"
Write-Host
Set-Location -Path $ScriptDir
$DefaultSettings = Get-Content 'defaults.json' | ConvertFrom-Json
Write-Host " - Loaded " -NoNewline
Write-Host -ForegroundColor Cyan ($DefaultSettings | Measure-Object).Count -NoNewline
Write-Host " default settings."
foreach ($DefaultSetting in $DefaultSettings) {
    Write-Host " - Setting " -NoNewLine
    Write-Host -ForegroundColor Cyan $DefaultSetting.Name -NoNewLine
    Write-Host " is " -NoNewline
    Write-Host -ForegroundColor Cyan $DefaultSetting.Value
}
Write-Host
$TakeDefaults = Read-Host " - Do you want to accept defaults? (Y/N)"
Write-Host
$settings = @{}
if ($TakeDefaults -eq 'Y')  {
    foreach ($DefaultSetting in $DefaultSettings) {
        $settings.Add($DefaultSetting.Name, $DefaultSetting.Value)
    }
} else {
    foreach ($DefaultSetting in $DefaultSettings) {
        Set-Variable -Name $DefaultSetting.Name -Value $DefaultSetting.Value -Scope 'Global'
        Write-Host " - Enter $($DefaultSetting.Description.ToLower())."
        $prompt = Read-Host "Press enter to accept the default [$($DefaultSetting.Value)]"
        if ([string]::IsNullOrWhiteSpace($prompt)) {
            $settings.Add($DefaultSetting.Name, $DefaultSetting.Value) | Out-Null
        } 
        else {
            $settings.Add($DefaultSetting.Name, $prompt) | Out-Null
        }
        Write-Host
    }
}
Write-Host " - Adding " -NoNewLine
Write-Host -Foregroundcolor Cyan 'ScriptDir' -NoNewline
Write-Host " setting as " -NoNewline
Write-Host -ForegroundColor Cyan $($Scriptdir) -NoNewline
Write-Host "..." -NoNewline
$settings.Add('ScriptDir', $Scriptdir)
Write-Host -ForegroundColor Green "Success"

Write-Host " - Adding " -NoNewLine
Write-Host -Foregroundcolor Cyan 'RootScriptDir' -NoNewline
Write-Host " setting as " -NoNewline
Write-Host -ForegroundColor Cyan $($RootScriptdir) -NoNewline
Write-Host "..." -NoNewline
$settings.Add('RootScriptDir', $RootScriptdir)
Write-Host -ForegroundColor Green "Success"

Write-Host " - Writing " -NoNewLine
Write-Host -Foregroundcolor Cyan ($settings).Count -NoNewline
Write-Host " settings to " -NoNewline
Write-Host -ForegroundColor Cyan "$ScriptDir\settings.json" -NoNewline
Write-Host "..." -NoNewline
try {
    $settings | ConvertTo-Json | Out-File 'settings.json'
    Write-Host -ForegroundColor Green "Success"
}
catch {
    $err = $_
    Write-Host -ForegroundColor Red "Error: $($err.Exception.Message)"
    continue
}
Write-Host
Write-Host "------------------------------------------"
Write-Host
#endregion

#region setup system dirs
Write-Host "Setting up the system directories"
Write-Host
$ModuleDir = "$RootScriptDir\Modules"
$ModuleName = 'SharePointPnPPowerShell2016'
$SetupDirs = @(
    $ModuleDir
    "$ScriptDir\Temp"
    "$ScriptDir\Credentials"
    "$ScriptDir\Logs"
)
foreach ($SetupDir in $SetupDirs) {
    if (!(Test-Path -Path $SetupDir)) {
        Write-Host " - Creating " -NoNewLine
        Write-Host -ForegroundColor Cyan $SetupDir -NoNewLine
        Write-Host "..." -NoNewLine
        try {
            New-Item $SetupDir -ItemType Directory | Out-Null
            Write-Host -ForegroundColor Green "Success."
        }
        catch {
            $err = $_
            Write-Host -ForegroundColor Red "Error: $($err.Exception.Message)"
        }
    }
    else {
        Write-Host " - " -NoNewLine
        Write-Host -ForegroundColor Cyan $SetupDir -NoNewLine
        Write-Host " already present. " -NoNewLine
        Write-Host -ForegroundColor Green "Doing nothing."
    }
}

# download sharepoint 2016 pnp module and update if already present
if (!(Test-Path -Path "$ModuleDir\$ModuleName")) {
    Write-Host " - Downloading " -NoNewline
    Write-Host -ForegroundColor Cyan $ModuleName -NoNewline
    Write-Host " module..." -NoNewline
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Save-Module -Name $ModuleName -Path $ModuleDir
        Write-Host -ForegroundColor Green "Success."
    }
    catch {
        $err = $_
        Write-Host -ForegroundColor Red "Error: $($err.Exception.Message)"
        continue
    }
    
}
else {
    Import-Module "$ModuleDir\$ModuleName" -Force -DisableNameChecking
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $PublishedVersion = Find-Module $ModuleName
        $LocalModuleVersion = Get-Module $ModuleName
        if ($LocalModuleVersion.Version -ne $PublishedVersion.Version) {
            try {
                Write-Host " - Updating " -NoNewline
                Write-Host -ForegroundColor Cyan $($ModuleName) -NoNewline
                Write-Host '...' -NoNewline
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Save-Module -Name $ModuleName -Path $ModuleDir
                Write-Host -ForegroundColor Green 'Success.'
            }
            catch {
                $err = $_
                Write-Host -ForegroundColor Red $err.Exception.Message
                continue
            }
        } 
        else {
            Write-Host " - " -NoNewline
            Write-Host -ForegroundColor Cyan $($ModuleName) -NoNewline
            Write-Host ' is up to date. ' -NoNewline
            Write-Host -ForegroundColor Green 'Doing nothing.'
        }
    }
    catch {
        $err = $_
        Write-Host "Error: $($err.Exception.Message)"
        continue
    }
}
Write-Host
Write-Host "------------------------------------------"
Write-Host
#endregion

#region setup credentials
Write-Host "Setting up the credential file"
Write-Host
Write-Host ' - Press any key to enter the service account credentials that will run the Task Schedule job on this server.'
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
$ServiceAccountCredential = Get-Credential -Message 'Enter credentials for the service account that will run scheduled tasks'
Write-Host ' - Press any key to enter the password for ' -NoNewline
Write-Host -ForegroundColor cyan 'domain.local/service-account' -NoNewline
Write-Host '.'
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
Start-Process -FilePath "powershell.exe" -ArgumentList "$ScriptDir\Initialize-CredentialFile.ps1" -Credential $ServiceAccountCredential -Wait
Write-Host
Write-Host "------------------------------------------"
Write-Host
#endregion

#region setup task schedule jobs
Write-Host "Setting up the automated job in Task Scheduler"
Write-Host
# controller
$JobName = 'Controller'
$ScriptName = $JobName + '.ps1'
$ScriptPath = Join-Path -Path $ScriptDir -ChildPath $ScriptName
$Script = '"' + $ScriptPath + '"'
$Repeat = (New-TimeSpan -Minutes 5)
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-file $Script" -WorkingDirectory $ScriptDir
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval $Repeat
$UserName = $ServiceAccountCredential.UserName
$Password = $ServiceAccountCredential.GetNetworkCredential().Password
$Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([timespan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable -DontStopOnIdleEnd
try {
    Register-ScheduledTask -TaskName $JobName -Action $Action -Trigger $Trigger -RunLevel Highest -User $UserName -Password $Password -Settings $Settings | Out-Null
    Write-Host " - Scheduled task " -NoNewLine
    Write-Host -ForegroundColor Cyan $JobName -NonewLine
    Write-Host " successfully setup."
}
catch {
    $err = $_
    Write-Host -ForegroundColor Red "Error setting up $($JobName) scheduled task: $($err.Exception.Message)"
    continue
}

# nightlycontroller
$JobName = 'NightlyController'
$ScriptName = $JobName + '.ps1'
$ScriptPath = Join-Path -Path $ScriptDir -ChildPath $ScriptName
$Script = '"' + $ScriptPath + '"'
$Repeat = (New-TimeSpan -Days 1)
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-file $Script" -WorkingDirectory $ScriptDir
$Trigger = New-ScheduledTaskTrigger -Once -At 2am -RepetitionInterval $Repeat
$UserName = $ServiceAccountCredential.UserName
$Password = $ServiceAccountCredential.GetNetworkCredential().Password
$Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([timespan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable -DontStopOnIdleEnd
try {
    Register-ScheduledTask -TaskName $JobName -Action $Action -Trigger $Trigger -RunLevel Highest -User $UserName -Password $Password -Settings $Settings | Out-Null
    Write-Host " - Scheduled task " -NoNewLine
    Write-Host -ForegroundColor Cyan $JobName -NonewLine
    Write-Host " successfully setup."
}
catch {
    $err = $_
    Write-Host -ForegroundColor Red "Error setting up $($JobName) scheduled task: $($err.Exception.Message)"
    continue
}
Write-Host
#endregion

#region the end
Write-Host "------------------------------------------"
Write-Host "| Initialize Controllers Complete        |"
Write-Host "------------------------------------------"
#endregion