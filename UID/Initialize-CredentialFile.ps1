<#
    .DESCRIPTION
    This scirpt is used along with Initialize-Controllers.ps1 to setup the processes
#>
$ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$CredentialFileName = "myCred_${env:USERNAME}_${env:COMPUTERNAME}.xml"
if (!(Test-Path -Path "$ScriptDir\Credentials\$CredentialFileName")) {
    $Credential = Get-Credential -Credential domain.local\service_account
    $Credential | Export-Clixml -Path "$ScriptDir\Credentials\$CredentialFileName"
    Write-Host "--> Created " -NoNewLine
    Write-Host -ForegroundColor Cyan "$ScriptDir\Credentials\$CredentialFileName" -NoNewLine
    Write-Host "."

}
else {
    Write-Host "--> " -NoNewLine
    Write-Host -ForegroundColor Cyan "$ScriptDir\Credentials\$CredentialFileName" -NoNewLine
    Write-Host " is already present. Delete the file if you'd like to recreate it."
}
Write-Host "Script complete. Press any key to close this window..." -NoNewLine
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
