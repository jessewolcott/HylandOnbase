# Application Friendly Name
$ApplicationFriendlyName = "HylandUnityClient22.1" # your install script (Install.ps1, appname.ps1, whatever)
$SourceFolder            = "C:\InTune Packages\Hyland\UnityClient\22.1" # Location of install script

# Where is your Win32 Prep Tool?
$WorkingFolder = "C:\Git\Windows\Microsoft-Win32-Content-Prep-Tool-master"

# Where should we output the .intunewin file?
$OutPutFolder            =  "C:\DEV"

# Setup
$SourceSetupFile         = "$SourceFolder\$ApplicationFriendlyName.ps1"
$IntuneWinAppUtil        = "IntuneWinAppUtil.exe"
$AppPath                 = "$WorkingFolder\$IntuneWinAppUtil"

Start-Process $AppPath -ArgumentList @(("-c `"$SourceFolder`""),("-s `"$SourceSetupFile`""),("-o `"$OutPutFolder`"")) -wait -PassThru