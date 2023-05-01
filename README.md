![](https://www.codetriage.com/mallockey/install-office365suite/badges/users.svg)
# Install-Office365Suite

## Updates

10/16/22

Important: Moved `Get-XMLFile` and  `Get-ODTURL` to an external module in this repo. So if you plan to deploy this script it will be easiest to just download it from the [PowerShell Gallery](https://www.powershellgallery.com/packages/Install-Office365Suite/). Otherwise you'll need to include the `InstallOffice.psm1` with your deployment or manually move the functions inside of the `Install-Office365Suite.ps1` for a single script deployment.

10/4/22
* Added `-LanguageIDs` parameter
* Added `-IncludeProject` parameter
* Added `-IncludeVisio` parameter
* Removed `-LoggingPath` as its not longer an option

7/26/22
* Fixed ConfigurationXMLFile Bug

## Description
A PowerShell script that installs Office 365 on a workstation with parameters that talor the install to your specific needs.
## Installing the script

`Install-Script -Name Install-Office365Suite`

## Features
The script will download the Office Deployment Tool from Microsoft's website first. If you have an XML file that you'd like to use you can supply it to the **-ConfigurationXMLFile** parameter like below:

`.\Install-Office365Suite.ps1 -ConfigurationXMLFile "C:\Kits\OfficeConfig.xml"`

If you don't, you can run it without any parameters and it will install with the default settings below:
```xml
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="MatchOS" />
    </Product>
  </Add>
  <Property Name="PinIconsToTaskbar" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="0" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
</Configuration>
```

Alternatively, you can set many settings from the command line that you'd like to include, below is a list of the settings and their values:

 Parameter | Possible Values 
--- | --- |
-AcceptEULA | TRUE,FALSE
-Channel | SemiAnnualPreview, SemiAnnual, MonthlyEnterprise, CurrentPreview, Current
-DisplayInstall | [Switch]
-EnableUpdates | TRUE, FALSE
-LanguageIDs | [Array] en-us, ar-sa 
-IncludeProject | [Switch]
-IncludeVisio | [Switch]
-ExcludeApps | Groove, Outlook, OneNote, Access, OneDrive, Publisher, Word, Excel, PowerPoint, Teams, Lync
-OfficeArch | 64, 32
-OfficeEdition | O365ProPlusRetail, O365BusinessRetail
-OfficeInstallerDownloadPath   | [String] *Specify path*
-SharedComputerLicensing | 0,1
-SourcePath | [String] *Specify path*
-PinItemsToTaskbar  | TRUE, FALSE (Windows 7 / 8 only!)
-KeepMSI | [Switch]
-SetFileFormat | [Switch]
-CleanUpInstallFiles | [Switch]

## Additional Info
By default, the script will create and download the ODT tool to "C:\Scripts\Office365Install" folder. You can change this with the **-OfficeInstallDownloadPath** parameter
