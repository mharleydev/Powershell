
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
$localDir = 'C:\Users\mike.harley\projects\HTML\'

$files = Get-ChildItem -path $localDir

foreach ($file in $files) {
    Write-Host 'Processing ' -NoNewline
    Write-Host -ForegroundColor Cyan $($file.Name) -NoNewline
    Write-Host '...' -NoNewline
    if (($file.Name).ToCharArray() -contains '+') {
        $newName = $file.Name.Replace('+', '_')
        Rename-Item -Path $file.Fullname -newname ($newName) 
        Write-Host 'Changed to ' -NoNewline
        Write-Host -ForegroundColor Green $($newName)
    }
    else {
        Write-Host No changes needed
    }
}
