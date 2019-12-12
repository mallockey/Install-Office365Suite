# Install-Office365
## Description
A PowerShell script that installs Office 365 on a workstation with parameters that talor the install to your specific needs.
## Installing the script
For the time being please copy and paste the code directly from GitHub. I will publish to the PowerShell Gallery when I've done more testing.
## Features
The script will download the Office Deployment Tool from Microsoft's website first. If you have an XML file that you'd like to use you can supply it to the **-ConfigurationXMLFile** parameter like below:

`.\Install-Office365.ps1 -ConfigurationXMLFile "C:\Kits\OfficeConfig.xml"`

If you don't you can run it without any parameters and it will install with the default settings below:

`<Configuration>
  <Add OfficeClientEdition="64" Channel="Broad">
    <Product ID="O365ProPlusRetail">
      <Language ID="MatchOS" />
    </Product>
  </Add>
  <Property Name="PinIconsToTaskbar" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="0" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
</Configuration>`
