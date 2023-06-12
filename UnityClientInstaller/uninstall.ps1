$PackageName = "HylandUnityClient22.1" #

$Path_local = "$Env:Programfiles\_MEM" # "$ENV:LOCALAPPDATA\_MEM" for user context installations
Start-Transcript -Path "$Path_local\uninstall\$PackageName-install.txt" -Force

try{
    Write-Output "Uninstalling $PackageName and all packages like it"
    (Get-WmiObject Win32_Product -filter "vendor like 'Hyland%'"| % { $_.Uninstall() }) | Out-Null

    if ($null -eq (Get-WmiObject Win32_Product -filter "vendor like 'Hyland%' AND name like 'Hyland%Unity%'"))
        {Write-Output "No installations found."}
        Else {Write-Output "Installations remain. Re-run this script."}
}catch{
    Write-Host "_____________________________________________________________________"
    Write-Host "ERROR while uninstalling $PackageName"
    Write-Host "$_"
}

Stop-Transcript

