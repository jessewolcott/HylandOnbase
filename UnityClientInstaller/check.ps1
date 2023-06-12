######################################################################################################################
# Program EXE with target Version or higher
######################################################################################################################
$ProgramPath = "C:\Program Files (x86)\Hyland\Unity Client\obunity.exe"


if (Test-Path -Path "C:\Program Files (x86)\Hyland\Unity Client\obunity.exe") {
    $ProgramVersion_target = [System.Version]'22.1.9.1000' 
    $ProgramVersion_current = [System.Version](Get-Item $ProgramPath).VersionInfo.FileVersion
        if($ProgramVersion_current -ge $ProgramVersion_target){
            Write-Host "Found it!"
    }
}

