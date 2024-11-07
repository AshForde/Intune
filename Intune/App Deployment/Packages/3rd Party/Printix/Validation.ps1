$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "2025.1.0.117"
$validationFile = "$validation\Printix Client.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}