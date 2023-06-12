# Hyland OnBase

Document management for the masses. 

## OnBase Unity Client and Office Tools 2019 Installer

This script was made for Intune but should work in your RMM of choice, so long as it handles Powershell reasonably. Also, this is tested with Unity Client 22.1.9.1000 and Office Tools 2019.

This is a free script, test with your build, don't yell at me.

### **HylandUnityClient22.1.ps1**

This installer should be placed in a folder with ```Hyland Unity Client.msi``` and ```Hyland Office 2019 Integration x86.msi```. You can configure ```IntuneAppCreator.ps1``` with your values to spit out an .intunewin package. Note, this installs the Office 2019 for x86 office. If you use 64 bit office, you'll have to change that out. 

I don't think I missed any settings in the top of the script for install. Its important to call out that Hyland installer doesn't have switches for half of the configuration options that are required for a meaningful install, so this script generates the configuration files for both the Unity Client and the Office tools in order to actually install. If there are more options or other things, please feel free to fork and PR.

### **check.ps1**
This is the version / presence check. Simply returns "Found it!" if ```"C:\Program Files (x86)\Hyland\Unity Client\obunity.exe"``` is present with a version greater than or equal to 22.1.9.1000

### **uninstall.ps1**
Removes everything on the machine with a vendor named "hyland". Careful with this one! If you have restricted mode (or core-only) powershell, it references the Uninstall() .net class and might not work