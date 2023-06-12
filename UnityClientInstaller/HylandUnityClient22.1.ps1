$PackageName          = "HylandUnityClient22.1-DEV"
$Path_local           = "$Env:Programfiles\_MEM" # "$ENV:LOCALAPPDATA\_MEM" for user context installations
Start-Transcript -Path "$Path_local\$PackageName-install.txt" -Force
$ErrorActionPreference = 'Stop'


# Onbase Specific Settings - These are used for both Unity Client and Office Integration Tools
# ServicePath =           The URL to the Service.asmx page of the Application Server
# DataSource =            The data source name (configured at the Application Server) to connect to
# FriendlyName =          The "friendly name" of the service location
# DefaultDomain =         The name of the domain to use by default
# UseInstitution =        A "true" or "false" value indicating if Institutional database is in use
# IdpUrl =                The URL to the IDP server to be used for authentication
# IdpClientId =           The client id configured on the IDP Server that is associated with this client
# AuthenticationType =    The authentication type to use =  Standard, NTAuthentication, ADFS, Idp

$ServicePath                             = '""' # Ex: '"https://contoso.com/appserver/service.asmx"'
$DataSource                              = '""' # Ex: '"OnBase"'
$FriendlyName                            = '""' # EX: '"OnBase Production 22.1 IDP"'
$DefaultDomain                           = '""'
$UseInstitution                          = '"false"'
$IdpUrl                                  = '""' # Ex: '"https://onbase.contoso.com/identityprovider"'
$IdpClientId                             = '""' # Ex: '"big-long-2134234-app-ID-2134"'
$AuthenticationType                      = '""' # Ex: '"Idp"'

# Unity Client Advanced Settings - These are the defaults
$IsPersistentLogonEnabled                = '"false"'
$enableWorkflowDebugTrace                = '"false"'
$MailClient                              = '"outlook"'
$EmailLinkAs                             = '"upop"'
$docpopurl                               = '"http://[server]/[virtual directory]/docpop/docpop.aspx?docid={0}&amp;clienttype=html"'
$reportingWebViewerUrl                   = '"http://[server]/[virtual directory]/Viewer.aspx"'
$MaxConcurrentAutoVueControls            = '"3"'
$EnableAutoVueControlPool                = '"false"'
$enableGECentricityIndexingMode          = '"false"'
$scanlog                                 = '"false"'
$enableModernScanning                    = '"true"'
$DisableUnityFormChromiumCache           = '"false"'
$suppressScriptErrors                    = '"false"'
$webBrowserRenderMode                    = '"edge"'
$useSystemBrowserForIdpLogin             = '"false"'
$enableEdgePrintForPDFs                  = '"false"'
# Language
$DisplayLanguage                         = 'en-US'
$Culture                                 = 'en-US'
# Application Enabler
$DefaultConfigFile                       = '""'
$UsePersistentWindowForRetrieval         = '"false"'
$UsePersistentWindowForCustomQuery       = '"false"'
$UsePersistentWindowForEformCreation     = '"false"' 
$UsePersistentWindowForFolderRetrieval   = '"false"'
$UsePersistentWindowForWorkflowRetrieval = '"false"'
$FolderPop                               = '"http://localhost/appnet/FolderPop/FolderPop.aspx"'
$openInNewWindow                         = '"true"'
$sendSessionID                           = '"true"'
$useURL                                  = '"false"'
$DefPopurl                               = '"http://[server]/virtualRoot"'
$DefPopExternal                          = '"no"'
$sendSessionIDdefpop                     = '"false"'
$OracleURL                               = '"http://localhost/appnet/OracleLogin.aspx"'
$PatientWindowURL                        = '"http://localhost/Medical.Pop/login.aspx"'
$sendSessionIDPatientWindow              = '"false"'
$ShowBanner                              = '"true"'

## end onbase config

try{

    Write-Output "Attempting to remove OnBase Unity 18"
    start-process 'msiexec.exe' -ArgumentList '/X{93BD81D2-3E88-4140-A290-9DD18C9BEB6F} /qn' -wait -Verbose -ErrorAction SilentlyContinue

    Write-Output "Attempting to remove OnBase Office Integration Tools 2016"
    start-process 'msiexec.exe' -ArgumentList '/X{9B4B1682-D536-4796-BD71-16CA11B7135A} /qn' -wait -Verbose -ErrorAction SilentlyContinue
   
    Write-Output "VCRedist Check"
    $x64VCRedist2012Version = [System.Version]((Get-ItemProperty "HKLM:\SOFTWARE\Classes\Installer\Dependencies\{ca67548a-5ebe-413a-b50c-4b9ceb6d66c6}" -ErrorAction SilentlyContinue).Version)
    If ($null -ne $x64VCRedist2012Version){
        $ProgramVersion_target = [System.Version]'11.0.61030.0'
        if($x64VCRedist2012Version -ge $ProgramVersion_target){
            Write-Output "VCRedist Check - Found VCRedist 2012 x64"
        }
            else {
                Write-Output "VCRedist Check - VCRedist 2012 x64 NOT FOUND. Exiting."
                Stop-Transcript
                BREAK
            }
    }
    
    $x86VCRedist2012Version = [System.Version]((Get-ItemProperty "HKLM:\SOFTWARE\Classes\Installer\Dependencies\{33d1fd90-4274-48a1-9bc1-97e33d9c2d6f}" -ErrorAction SilentlyContinue).Version)

    If ($null -ne $x86VCRedist2012Version){
        $ProgramVersion_target = [System.Version]'11.0.61030.0'
        if($x86VCRedist2012Version -ge $ProgramVersion_target){
            Write-Output "VCRedist Check - Found VCRedist 2012 x86"
        }
        else {
            Write-Output "VCRedist Check - VCRedist 2012 x86 NOT FOUND. Exiting."
            Stop-Transcript
            BREAK
        }
    }

    $2015_2022_Version = (Get-ItemProperty "HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" -ErrorAction SilentlyContinue).bld

    If ($null -ne $2015_2022_Version){
        $ProgramVersion_target = '31938'
        if($2015_2022_Version -ge $ProgramVersion_target){
            Write-Output "VCRedist Check - Found VCRedist 2015-2022 x64"
        }
        else {
            Write-Output "VCRedist Check - VCRedist 2015-2022 x64 NOT FOUND. Exiting."
            Stop-Transcript
            BREAK
        }
    }


            Write-Output "Installing $Packagename"
            Start-Process 'msiexec.exe' -ArgumentList '/i "Hyland Unity Client.msi" /q ADDLOCAL="Unity_Client,UpopAutomation" CREATE_DESKTOP_SHORTCUTS="1"' -wait
            

$UnityClientConfig = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <configSections>
    <section name="Hyland.Canvas.Client" type="Hyland.Canvas.CanvasClientConfiguration, Hyland.Canvas"/>
    <section name="Hyland.Services.Client" type="Hyland.Services.Client.ServiceClientConfiguration, Hyland.Services.Client"/>
    <section name="Hyland.Canvas" type="Hyland.Canvas.CanvasConfiguration, Hyland.Canvas"/>
    <section name="Hyland.Canvas.Automation.Services" type="Hyland.Canvas.Automation.Services.AutomationConfiguration, Hyland.Canvas.Automation.Services"/>
    <!-- section name="Hyland.Authentication" type="Hyland.Authentication.Configuration.AuthenticationConfigurationSection, Hyland.Authentication" / -->
  </configSections>
  <runtime>
    <generatePublisherEvidence enabled="false"/>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="IdentityModel" publicKeyToken="e7877f4675df049f" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-3.10.9.0" newVersion="3.10.9.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Extensions.DependencyInjection.Abstractions" publicKeyToken="adb9793829ddae60" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-2.1.1.0" newVersion="2.1.1.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Extensions.Logging" publicKeyToken="adb9793829ddae60" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-2.1.1.0" newVersion="2.1.1.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Extensions.Logging.Abstractions" publicKeyToken="adb9793829ddae60" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-2.1.1.0" newVersion="2.1.1.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.IdentityModel.Tokens" publicKeyToken="31bf3856ad364e35" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-5.4.0.0" newVersion="5.4.0.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Owin" publicKeyToken="31bf3856ad364e35" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-4.0.1.0" newVersion="4.0.1.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Newtonsoft.Json" publicKeyToken="30ad4fe6b2a6aeed" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-11.0.0.0" newVersion="11.0.0.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.Buffers" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-4.0.3.0" newVersion="4.0.3.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.Configuration.ConfigurationManager" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-4.0.3.0" newVersion="4.0.3.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.IdentityModel.Tokens.Jwt" publicKeyToken="31bf3856ad364e35" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-5.4.0.0" newVersion="5.4.0.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.Memory" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-4.0.1.1" newVersion="4.0.1.1"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.Runtime.CompilerServices.Unsafe" publicKeyToken="b03f5f7f11d50a3a" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-4.0.6.0" newVersion="4.0.6.0"/>
      </dependentAssembly>
    </assemblyBinding>
  </runtime>
  <system.diagnostics>
    <switches>
      <add name="LoggingLevel" value="2"/>
    </switches>
  </system.diagnostics>
  <Hyland.Canvas.Client>
    <ServiceMode allowExit="true" enabled="true" autoLaunch="false">
      <Feature name="Unity Client" enabled="true"/>
      <Feature name="Application Enabler" enabled="false"/>
      <Feature name="Unity Client Automation API" enabled="false"/>
      <Feature name="Virtual Print Driver" enabled="false"/>
      <Feature name="Unity Pop" enabled="true"/>
    </ServiceMode>
  </Hyland.Canvas.Client>
  <Hyland.Services.Client>
    <!--
    SERVICE LOCATIONS
    Service Locations appear in the Unity Client login dialog, and are 
    pre-configured items that contain web service connection information.
    They can be configured in the registry by use of the OnBase Desktop Control Panel,
    or they can be added here.
    
    Service Locations have the following attributes:
    
      ServicePath:          The URL to the Service.asmx page of the Application Server
      DataSource:           The data source name (configured at the Application Server) to connect to
      FriendlyName:         The "friendly name" of the service location
      DefaultDomain:        The name of the domain to use by default
      UseInstitution:       A "true" or "false" value indicating if Institutional database is in use
      IdpUrl:               The URL to the IDP server to be used for authentication
      IdpClientId:          The client id configured on the IDP Server that is associated with this client
      AuthenticationType:   The authentication type to use: Standard, NTAuthentication, ADFS, Idp
    -->
    <ServiceLocations>
	  <add DefaultDomain=$DefaultDomain UseInstitution=$UseInstitution IdpUrl=$IdpUrl IdpClientId=$IdpClientId ServicePath=$ServicePath DataSource=$DataSource FriendlyName=$FriendlyName AuthenticationType=$AuthenticationType/>
    </ServiceLocations>
    <AllowAllSSLCertificates Enabled="false"/>
    <OptimizedServicePipeline Enabled="true"/>
  </Hyland.Services.Client>
  <appSettings>
    <!--Toggles the visibility of the workflow trace ribbon group in the workflow layout-->
    <add key="enableWorkflowDebugTrace" value=$enableWorkflowDebugTrace/>
    <add key="mailclient" value=$mailclient/>
    <!--Uncomment the next line to enable sending links from the client. Valid values include "disabled", "weblink", "upop" and "upop-file"-->
    <add key="emailLinkAs" value=$emailLinkAs/>
    <!--
    When using docpop links make sure that you update the server and the virtual directory in the line below. There
    is no additional configuration when using upop links.
    -->
    <add key="docpopurl" value=$docpopurl/>
    <!--
     When using web links for dashboards or reports make sure that you update the server and the virtual directory in the line below. 
     -->
    <add key="reportingWebViewerUrl" value=$reportingWebViewerUrl/>
    <!-- The maximum number of concurrent instances of the AutoVue control. -->
    <add key="MaxConcurrentAutoVueControls" value=$MaxConcurrentAutoVueControls/>
    <!-- Enable the pooling of AutoVue controls (to speed loading of AutoVue viewer windows). -->
    <add key="EnableAutoVueControlPool" value=$EnableAutoVueControlPool/>
    <!-- Set this option to true to enable the Indexing Integration for GE Centricity. -->
    <add key="enableGECentricityIndexingMode" value=$enableGECentricityIndexingMode/>
    <add key="scanlog" value=$scanlog/>
    <!-- Enables the Modern version of Batch Scanning, featuring subprocess isolation. -->
    <add key="enableModernScanning" value=$enableModernScanning/>
    <!-- Forces Unity/Image Forms to skip the cache when retrieving resources (scripts, etc.) -->
    <add key="DisableUnityFormChromiumCache" value=$DisableUnityFormChromiumCache/>
    <!-- Stop script errors from displaying to the user -->
    <add key="suppressScriptErrors" value=$suppressScriptErrors/>
    <!-- Set rendering browser used by default. Valid values include "edge" and "ie" -->
    <add key="webBrowserRenderMode" value=$webBrowserRenderMode/>
    <!-- Set this option to true to use the default system web browser for Idp authentication-->
    <add key="useSystemBrowserForIdpLogin" value=$useSystemBrowserForIdpLogin/>
    <!-- Set this option to true to use Edge to print PDFs -->
    <add key="enableEdgePrintForPDFs" value=$enableEdgePrintForPDFs/>
  </appSettings>
  <Hyland.Canvas>
    <!--
    Unity Client will be displayed in the language configured as the 
    "Display language" or the "Default input language" in Windows.
    To display Canvas in a different language than Windows, set the DisplayLanguage to a culture setting like 
    en-US for US English, de-DE for German, fr-FR for French etc.
    -->
    <!--<DisplayLanguage>$DisplayLanguage</DisplayLanguage>-->
    <!--
    Unity Client will display date, time, currency, and numeric values using the default settings configured in 
    Windows under "Regional and Language Options", the "Formats" or "Standards and formats" section.
    To display the values in Unity Client using different cultural settings than the ones configured in Windows,
    set the Culture to a culture setting like en-US for US English, de-DE for German, fr-FR for French etc.
    -->
    <!--<Culture>$Culture</Culture>-->
    <!--
    To disable handling of Unity Client contexts from Application Enabler, create an attribute of EnableHandlingUnityClientEvents="false".
    
    When a scrape event with persistent window enabled is performed, they will be displayed in a 
    persistent window that is shared among all contexts. If it is false, the context will open in 
    a new window every time.
    -->
    <ApplicationEnabler DefaultConfigFile=$DefaultConfigFile UsePersistentWindowForRetrieval=$UsePersistentWindowForRetrieval UsePersistentWindowForCustomQuery=$UsePersistentWindowForCustomQuery UsePersistentWindowForEformCreation=$UsePersistentWindowForEformCreation UsePersistentWindowForFolderRetrieval=$UsePersistentWindowForFolderRetrieval UsePersistentWindowForWorkflowRetrieval=$UsePersistentWindowForWorkflowRetrieval>
      <!--
        FolderPop
        useURL: by setting this attribute to true the scraped data will be added to the URL instead of added as post data.  If the length of the scrape data
                is larger than the max length of the URL the data may be truncated.
        -->
      <FolderPop url=$FolderPop openInNewWindow=$openInNewWindow sendSessionID=$sendSessionID useURL=$useURL/>
      <!--
        DefPop
      -->
      <DefPop url=$DefPopurl external=$DefPopExternal sendSessionID=$sendSessionIDdefpop/>
      <!--
        The "Oracle" node is temporary and may be removed in future releases 
       -->
      <Oracle url=$OracleURL />
      <PatientWindow url=$PatientWindowURL sendSessionID=$sendSessionIDPatientWindow showBanner=$showBanner/>
    </ApplicationEnabler>
    <!--
    IsPersistentLogonEnabled - When set to "true", this option will allow
    users to store their credentials securely so that the next time they run
    Unity Client, they will be automatically logged in. Their stored
    credentials will be held until their login fails to work, this option
    is set back to false, or the user manually logs out.
    -->
    <Application IsPersistentLogonEnabled=$IsPersistentLogonEnabled/>
    <!-- 
    The default home layout for the Unity Client exposes an RSS Feed ticker and a web browser.
    The list of websites here will be made available in the navigation list for browsing, and 
    the list of RSS feeds will be loaded and displayed in the ticker.
    For a website, the URL is required, but the name is optional (the URL will be used when name is absent).
    For a feed, the URL is required.
    -->
    <DefaultHomeLayout>
      <Websites>
        <Website Name="Unity" Url="https://unity.hyland.com"/>
        <Website Name="Community" Url="https://community.hyland.com/"/>
        <Website Name="OnBase" Url="https://www.hyland.com/onbase"/>
      </Websites>
      <Feeds>
        <Feed Url="https://blog.Hyland.com/feed/"/>
      </Feeds>
    </DefaultHomeLayout>
    <!--
    The UserGroupTimeoutMode setting controls how timeout is detected.  This is only used when UserGroup Timeout is configured.
      Request (Default): Requests to the application server will serve as means to identify when a user is active within the application.
      Windows: User interaction with any application within a Windows session will serve as means to identify when a user is active.
    -->
    <UserGroupTimeoutMode>Request</UserGroupTimeoutMode>
  </Hyland.Canvas>
  <Hyland.Canvas.Automation.Services>
    <WCF enabled="true"/>
    <VPDMailslot enabled="true" deleteAfterUpload="true"/>
    <HttpAutomation enabled="false" port="15412"/>
    <HttpsAutomation enabled="false" port="15425" certificateLocation=""/>
  </Hyland.Canvas.Automation.Services>
  <startup>
    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.7.2"/>
  </startup>
  <!--
  <system.web>
  <webServices>
    <soapExtensionTypes>
      <add type="Hyland.Authentication.ADFS.CustomCanvasADFSAuthSoapExtension, Hyland.Authentication" />
    </soapExtensionTypes>
  </webServices>
</system.web>

  <Hyland.Authentication>
    <adfs enabled="true" logClientEventsToEventLog="true">
        <wsTrust forceNTLM="false">
          <adfsEndpointAddress>https://<ADFS_SERVER>/adfs/services/trust/2005/windowstransport</adfsEndpointAddress>
          <securityMode>Transport</securityMode>
          <trustVersion>WSTrustFeb2005</trustVersion>
          <appliesTo>http://mydomain.com/AppNet/</appliesTo>
        </wsTrust>
    </adfs>
  </Hyland.Authentication>
  -->
</configuration>
"@

Write-Output "Writing configuration file for Unity Client"
$UnityClientConfigFilePath = "C:\Program Files (x86)\Hyland\Unity Client\obunity.exe.config"
if ($UnityClientConfigFilePath){
    Write-Output "Overwriting configuration file"
    New-Item $UnityClientConfigFilePath -force
    $UnityClientConfig | Add-Content $UnityClientConfigFilePath
}
    else { 
        Write-Output "No configuration file found. Creating new."
        New-Item $OfficeConfigFilePath -force
        $UnityClientConfig | Add-Content $OfficeConfigFilePath
    }

        Write-Output "Installing Office Integration"

    if (Test-Path -Path "C:\Program Files (x86)\Hyland\Office Integration\Office 2019\x64\RenderProcess.exe") {
        $ProgramVersion_target = '2.6.6.0' 
        $ProgramVersion_current = (Get-Item "C:\Program Files (x86)\Hyland\Office Integration\Office 2019\x64\RenderProcess.exe").VersionInfo.FileVersion
            if($ProgramVersion_current -ge [System.Version]$ProgramVersion_target){
                Write-Output "Found Office Integration Installed."
        }
    }
        else{
            Write-Output "Installing Office Integration 2019 (x86)"

            # https://support.hyland.com/r/OnBase/Document-Composition/English/Foundation-22.1/Document-Composition/Hyland-Office-Products-Installers/Installation-Using-the-Standard-MSI-Installer/Controlling-the-Installer-from-the-Command-Line/Feature-and-Property-Names/Feature-Names/Outlook-Integration-Associated-Properties

            $CompleteCommandArgs = '/q /CompleteCommandArgs "ADDLOCAL=Outlook_Integration_Module SERVICE_LOCATION_DATA_SOURCE=\"OnBase\" SERVICE_LOCATION_DISPLAY_NAME=\"OnBase EP5 - Prod\" SERVICE_LOCATION_IDP_ENABLED_2019=\"true\" OUTLOOKINTEGRATION_WORKFLOW_BUTTON_OPTION=\"true\" OUTLOOKINTEGRATION_WORKFLOW_UPDATE_NOTIFICATION=\"true\" OUTLOOKINTEGRATION_AUTO_SELECT_ATTACHMENTS=\"true\" OUTLOOKINTEGRATION_DOCUMENT_DATE_SENT_DATE=\"true\" OUTLOOKINTEGRATION_ENABLE_IMPORT_FROM_FILE=\"true\" OUTLOOKINTEGRATION_HELP_BUTTON_OPTION=\"true\" OUTLOOKINTEGRATION_IMPORT_BUTTON_OPTION=\"true\" OUTLOOKINTEGRATION_KEEP_MESSAGES=\"true\" OUTLOOKINTEGRATION_RETRIEVAL_BUTTON_OPTION=\"true\""'

        Start-Process ".\setup.exe" -ArgumentList $CompleteCommandArgs -Wait -Verbose

$OfficeConfigFile = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <configSections>
    <section name="Hyland.Services.Client" type="Hyland.Services.Client.ServiceClientConfiguration, Hyland.Services.Client"/>
    <section name="Hyland.Logging" type="Hyland.Logging.Configuration.ClientConfigurationHandler,Hyland.Logging"/>
    <section name="Hyland.Canvas" type="Hyland.Canvas.CanvasConfiguration, Hyland.Canvas"/>
    <!-- section name="Hyland.Authentication" type="Hyland.Authentication.Configuration.AuthenticationConfigurationSection, Hyland.Authentication" / -->
    <sectionGroup name="userSettings" type="System.Configuration.UserSettingsGroup, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089">
      <section name="Hyland.Office2019.Outlook.Addin.Properties.Settings" type="System.Configuration.ClientSettingsSection, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" allowExeDefinition="MachineToLocalUser" requirePermission="false"/>
    </sectionGroup>
    <sectionGroup name="applicationSettings" type="System.Configuration.ApplicationSettingsGroup, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089">
      <section name="Hyland.Office2019.Outlook.Addin.Properties.Settings" type="System.Configuration.ClientSettingsSection, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" requirePermission="false"/>
    </sectionGroup>
  </configSections>
  <runtime>
    <generatePublisherEvidence enabled="false"/>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Extensions.Logging.Abstractions" publicKeyToken="adb9793829ddae60" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-2.1.1.0" newVersion="2.1.1.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Extensions.DependencyInjection.Abstractions" publicKeyToken="adb9793829ddae60" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-2.1.1.0" newVersion="2.1.1.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Newtonsoft.Json" publicKeyToken="30ad4fe6b2a6aeed" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-11.0.0.0" newVersion="11.0.0.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.IdentityModel.Tokens" publicKeyToken="31bf3856ad364e35" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-5.4.0.0" newVersion="5.4.0.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="IdentityModel" publicKeyToken="e7877f4675df049f" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-3.10.9.0" newVersion="3.10.9.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.Configuration.ConfigurationManager" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-4.0.3.0" newVersion="4.0.3.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Extensions.Logging" publicKeyToken="adb9793829ddae60" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-2.1.1.0" newVersion="2.1.1.0"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.Runtime.CompilerServices.Unsafe" publicKeyToken="b03f5f7f11d50a3a" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-4.0.4.1" newVersion="4.0.4.1"/>
      </dependentAssembly>
    </assemblyBinding>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.Memory" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral"/>
        <bindingRedirect oldVersion="0.0.0.0-4.0.1.1" newVersion="4.0.1.1"/>
      </dependentAssembly>
    </assemblyBinding>
  </runtime>
  <system.diagnostics>
    <switches>
      <!-- Logging Level 2 is Warning -->
      <add name="LoggingLevel" value="2"/>
    </switches>
  </system.diagnostics>
  <Hyland.Services.Client>
    <!--
    SERVICE LOCATIONS
    Service Locations appear in the Outlook Integration login dialog, and are 
    pre-configured items that contain web service connection information.
    They can be configured in the registry by use of the OnBase Desktop Control Panel,
    or they can be added here.
    
    Service Locations have the following attributes:
    
      ServicePath:          The URL to the Service.asmx page of the Application Server
      DataSource:           The data source name (configured at the Application Server) to connect to
      FriendlyName:         The "friendly name" of the service location
      DefaultDomain:        The name of the domain to use by default
      UseInstitution:       A "true" or "false" value indicating if Institutional database is in use
      IdpUrl:               The URL to the IDP server to be used for authentication
      IdpClientId:          The client id configured on the IDP Server that is associated with this client
      AuthenticationType:   The authentication type to use: Standard, NTAuthentication, ADFS, Idp
    -->
    <ServiceLocations>
      <add DefaultDomain=$DefaultDomain UseInstitution=$UseInstitution IdpClientId=$IdpClientId ServicePath=$ServicePath DataSource=$DataSource FriendlyName=$FriendlyName AuthenticationType=$AuthenticationType IdpUrl=$IdpUrl/>
    </ServiceLocations>
    <OptimizedServicePipeline Enabled="true"/>
    <AllowAllSSLCertificates Enabled="false"/>
  </Hyland.Services.Client>
  <appSettings>
    <add key="autoRegisterUpop" value="true"/>
    <add key="enableUpopAutomation" value="false"/>
    <!-- 
	A bug in WPF can cause high CPU usage and freezing of the UI on the client. 
	This is triggered by other applications running on the machine that use the 
	Microsoft UI Automation APIs - like screen readers, touch features, 
	or AppEnabler type of applications that scrape values from the screen. 
	If experiencing this issue, you can attempt to mitigate the issue by setting 
	enableUIAutomationWorkaround to true.
	-->
    <add key="enableUIAutomationWorkaround" value="false"/>
    <!--Toggles the visibility of the workflow trace ribbon group in the workflow layout-->
    <add key="enableWorkflowDebugTrace" value="false"/>
    <add key="mailclient" value="outlook"/>
    <!--Uncomment the next line to enable sending links from the client. Valid valies include "disabled", "docpop", and "upop"-->
    <add key="emailLinkAs" value="disabled"/>
    <!--
	When using docpop links make sure that you update the server and the virtual directory in the line below. There
	is no additional configuration when using upop links.
	-->
    <add key="docpopurl" value="http://[server]/[virtual directory]/docpop/docpop.aspx?docid={0}&amp;clienttype=html"/>
    <!-- Set this option to true to use the default system web browser for Idp authentication-->
    <add key="useSystemBrowserForIdpLogin" value="false"/>
  </appSettings>
  <Hyland.Canvas>
    <!--
	Unity Client will be displayed in the language configured as the 
	"Display language" or the "Default input language" in Windows.
	To display Canvas in a different language than Windows, set the DisplayLanguage to a culture setting like 
	en-US for US English, de-DE for German, fr-FR for French etc.
	-->
    <!--<DisplayLanguage>en-US</DisplayLanguage>-->
    <!--
	Unity Client will display date, time, currency, and numeric values using the default settings configured in 
	Windows under "Regional and Language Options", the "Formats" or "Standards and formats" section.
	To display the values in Unity Client using different cultural settings than the ones configured in Windows,
	set the Culture to a culture setting like en-US for US English, de-DE for German, fr-FR for French etc.
	-->
    <!--<Culture>en-US</Culture>-->
    <!--
		To allow this add-in to recieve Application Enabler events, set EnableHandlingUnityClientEvents = true.
		-->
    <ApplicationEnabler EnableHandlingUnityClientEvents="True"/>
  </Hyland.Canvas>
  <Hyland.Logging>
    <WindowsEventLogging sourcename="Hyland Application Server"/>
    <Routes>
      <Route name="DiagnosticsConsole">
        <add key="Http" value="http://localhost:8989"/>
        <!-- <add key="minimum-level" value="Trace" /> -->
      </Route>
      <Route name="ErrorEventLog">
        <!-- Write errors to the Event Log using the source specified above -->
        <add key="HylandLog"/>
        <add key="minimum-level" value="Error"/>
      </Route>
      <!--<Route name="ELK">
            <add key="Http" value="http://localhost:5586"/>
            <add key="CompactHttpFormat"/>
        </Route>-->
    </Routes>
    <AuditRoutes>
      <!--<Route name="DiagnosticsConsole">
            <add key="Http" value="http://localhost:8989" /> 
        </Route>-->
    </AuditRoutes>
  </Hyland.Logging>
  <!--
  <system.web>
  <webServices>
    <soapExtensionTypes>
      <add type="Hyland.Authentication.ADFS.CustomCanvasADFSAuthSoapExtension, Hyland.Authentication" />
    </soapExtensionTypes>
  </webServices>
</system.web>

  <Hyland.Authentication>
    <adfs enabled="true" logClientEventsToEventLog="true">
        <wsTrust forceNTLM="false">
          <adfsEndpointAddress>https://<ADFS_SERVER>/adfs/services/trust/2005/windowstransport</adfsEndpointAddress>
          <securityMode>Transport</securityMode>
          <trustVersion>WSTrustFeb2005</trustVersion>
          <appliesTo>http://mydomain.com/AppNet/</appliesTo>
        </wsTrust>
    </adfs>
  </Hyland.Authentication>
  -->
  <startup>
    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.7.2"/>
  </startup>
  <userSettings>
    <Hyland.Office2019.Outlook.Addin.Properties.Settings>
      <setting name="FullTextSearchTimeout" serializeAs="String">
        <value>30</value>
      </setting>
    </Hyland.Office2019.Outlook.Addin.Properties.Settings>
  </userSettings>
  <applicationSettings>
    <Hyland.Office2019.Outlook.Addin.Properties.Settings>
      <setting name="AdminConfiguredFolders" serializeAs="Xml">
        <value>
          <FolderNodeList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
            <FolderDictionary/>
          </FolderNodeList>
        </value>
      </setting>
      <setting name="ShowUploadTaskPaneOnLeft" serializeAs="String">
        <value>False</value>
      </setting>
      <setting name="AdminExplorerNode" serializeAs="Xml">
        <value>
          <ExplorerNode xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
            <OutlookIntegrationDeployed>true</OutlookIntegrationDeployed>
            <WorkViewIntegrationDeployed>false</WorkViewIntegrationDeployed>
            <ImportButton>true</ImportButton>
            <ExecutableButton>false</ExecutableButton>
            <RetrievalButton>true</RetrievalButton>
            <WorkflowButton>true</WorkflowButton>
            <WorkviewButton>false</WorkviewButton>
            <HelpButton>true</HelpButton>
            <CreateFormButton>false</CreateFormButton>
            <UploadFromFileButton>true</UploadFromFileButton>
            <ExecutablePath></ExecutablePath>
            <ExecutableSwitches></ExecutableSwitches>
            <DeleteMessage>false</DeleteMessage>
            <MoveMessage>false</MoveMessage>
            <MoveFolder></MoveFolder>
            <KeepMessageInPlace>true</KeepMessageInPlace>
            <SelectAllAttachments>true</SelectAllAttachments>
            <DocDateAsSent>true</DocDateAsSent>
            <MSGFormat>false</MSGFormat>
            <SeparateAttachments>false</SeparateAttachments>
            <ShowImaging>false</ShowImaging>
            <MyReadingGroupsButton>false</MyReadingGroupsButton>
            <ReportingDashboardsButton>false</ReportingDashboardsButton>
            <AdhocTaskComplete>true</AdhocTaskComplete>
            <AdminConfig>false</AdminConfig>
            <AllowUserConfig>true</AllowUserConfig>
            <AllowUserFolderConfig>true</AllowUserFolderConfig>
            <Default>false</Default>
            <UpdateClient>false</UpdateClient>
            <TimeStamp>05/23/2023 16:44:01</TimeStamp>
          </ExplorerNode>
        </value>
      </setting>
    </Hyland.Office2019.Outlook.Addin.Properties.Settings>
  </applicationSettings>
</configuration>
"@

            Write-Output "Writing configuration file for Office Integration. IDP Client set to $IDPClientID"
            $OfficeConfigFilePath = "C:\Program Files (x86)\Hyland\Office Integration\Office 2019\Hyland.Office2019.Outlook.Addin.dll.config"
            if ($OfficeConfigFilePath){
                Write-Output "Overwriting configuration file"
                New-Item $OfficeConfigFilePath -force
                $OfficeConfigFile | Add-Content $OfficeConfigFilePath
            }
                else { 
                    Write-Output "No configuration file found. Creating new."
                    New-Item $OfficeConfigFilePath -force
                    $OfficeConfigFile | Add-Content $OfficeConfigFilePath
                }

        }
        Exit 3010
}catch{
    Write-Host "_____________________________________________________________________"
    Write-Host "ERROR while installing $PackageName"
    Write-Host "$_"
}

Stop-Transcript

