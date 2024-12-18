[CmdletBinding(DefaultParameterSetName = 'XMLFile')]
param(
  [Parameter(ParameterSetName = 'XMLFile')][String]$ConfigurationXMLFile,
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$AcceptEULA = 'TRUE',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('SemiAnnualPreview', 'SemiAnnual', 'MonthlyEnterprise', 'CurrentPreview', 'Current')]$Channel = 'Current',
  [Parameter(ParameterSetName = 'NoXML')][Switch]$DisplayInstall,
  [Parameter(ParameterSetName = 'NoXML')][Switch]$IncludeProject,
  [Parameter(ParameterSetName = 'NoXML')][Switch]$IncludeVisio,
  [Parameter(ParameterSetName = 'NoXML')][Array]$LanguageIDs,
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('Groove', 'Outlook', 'OneNote', 'Access', 'OneDrive', 'Publisher', 'Word', 'Excel', 'PowerPoint', 'Teams', 'Lync')][Array]$ExcludeApps,
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('64', '32')]$OfficeArch = '64',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('O365ProPlusRetail', 'O365BusinessRetail')]$OfficeEdition = 'O365ProPlusRetail',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet(0, 1)]$SharedComputerLicensing = '0',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$EnableUpdates = 'TRUE',
  [Parameter(ParameterSetName = 'NoXML')][String]$SourcePath,
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$PinItemsToTaskbar = 'TRUE',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$ForceOpenAppShutdown = 'FALSE',
  [Parameter(ParameterSetName = 'NoXML')][Switch]$KeepMSI,
  [Parameter(ParameterSetName = 'NoXML')][Switch]$RemoveAllProducts,
  [Parameter(ParameterSetName = 'NoXML')][Switch]$SetFileFormat,
  [Parameter(ParameterSetName = 'NoXML')][Switch]$ChangeArch,
  [String]$OfficeInstallDownloadPath = 'C:\Scripts\Office365Install',
  [Switch]$CleanUpInstallFiles = $False
)

$VerbosePreference = 'Continue'
$ErrorActionPreference = 'Stop'

$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
  Write-Warning 'Script is not running as Administrator'
  Write-Warning 'Please rerun this script as Administrator.'
  exit
}

Import-Module '.\InstallOffice.psm1'

if (-Not(Test-Path $OfficeInstallDownloadPath )) {
  New-Item -Path $OfficeInstallDownloadPath -ItemType Directory | Out-Null
}

if (!($ConfigurationXMLFile)) {

  if ($ExcludeApps) {
    $OfficeXML = Get-XMLFile -AcceptEULA $AcceptEULA `
      -Channel $Channel `
      -DisplayInstall:$DisplayInstall `
      -IncludeProject:$IncludeProject `
      -IncludeVisio:$IncludeVisio `
      -LanguageIDs $LanguageIDs `
      -ExcludeApps $ExcludeApps `
      -OfficeArch $OfficeArch `
      -OfficeEdition $OfficeEdition `
      -SharedComputerLicensing $SharedComputerLicensing `
      -EnableUpdate $EnableUpdates `
      -PinItemsToTaskBar $PinItemsToTaskbar `
      -ForceOpenAppShutdown $ForceOpenAppShutdown `
      -KeepMSI:$KeepMSI `
      -RemoveAllProducts:$RemoveAllProducts `
      -SetFileFormat:$SetFileFormat `
      -ChangeArch:$ChangeArch `
  
  }
  else {
    $OfficeXML = Get-XMLFile -AcceptEULA $AcceptEULA `
      -Channel $Channel `
      -DisplayInstall:$DisplayInstall `
      -IncludeProject:$IncludeProject `
      -IncludeVisio:$IncludeVisio `
      -LanguageIDs $LanguageIDs `
      -OfficeArch $OfficeArch `
      -OfficeEdition $OfficeEdition `
      -SharedComputerLicensing $SharedComputerLicensing `
      -EnableUpdate $EnableUpdates `
      -PinItemsToTaskBar $PinItemsToTaskbar `
      -ForceOpenAppShutdown $ForceOpenAppShutdown `
      -KeepMSI:$KeepMSI `
      -RemoveAllProducts:$RemoveAllProducts `
      -SetFileFormat:$SetFileFormat `
      -ChangeArch:$ChangeArch `
  
  }

  $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")

  $ConfigurationXMLFile = "$OfficeInstallDownloadPath\OfficeInstall.xml"
}
else {

  if (!(Test-Path $ConfigurationXMLFile)) {
    Write-Warning 'The configuration XML file is not a valid file'
    Write-Warning 'Please check the path and try again'
    exit
  }

}

$ODTInstallLink = Get-ODTURL

if ($Null -eq $ODTInstallLink) {
  Write-Error "Could not find ODT install link, exiting"
  exit
}

#Download the Office Deployment Tool
Write-Verbose 'Downloading the Office Deployment Tool...'
try {
  Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"
}
catch {
  Write-Warning 'There was an error downloading the Office Deployment Tool.'
  Write-Warning 'Please verify the below link is valid:'
  Write-Warning $ODTInstallLink
  exit
}

#Run the Office Deployment Tool setup
try {
  Write-Verbose 'Running the Office Deployment Tool...'
  Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:`"$OfficeInstallDownloadPath`"" -Wait
}
catch {
  Write-Warning 'Error running the Office Deployment Tool. The error is below:'
  Write-Warning $_
}

#Run the O365 install
try {
  Write-Verbose 'Downloading and installing Microsoft 365'
  $Silent = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure `"$ConfigurationXMLFile`"" -Wait -PassThru
}
catch {
  Write-Warning 'Error running the Office install. The error is below:'
  Write-Warning $_
}

#Check if Office 365 suite was installed correctly.
$RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

$OfficeInstalled = $False
foreach ($Key in (Get-ChildItem $RegLocations) ) {
  if ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
    $OfficeVersionInstalled = $Key.GetValue('DisplayName')
    $OfficeInstalled = $True
  }
}

if ($OfficeInstalled) {
  Write-Verbose "$($OfficeVersionInstalled) installed successfully!"
}
else {
  Write-Warning 'Microsoft 365 was not detected after the install ran'
}

if ($CleanUpInstallFiles) {
  Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse
}
