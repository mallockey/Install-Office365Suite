[CmdletBinding(DefaultParameterSetName = 'XMLFile')]
Param(
  [Parameter(ParameterSetName = "XMLFile")][ValidateNotNullOrEmpty()][String]$ConfiguratonXMLFile,
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("TRUE","FALSE")]$AcceptEULA = "TRUE",
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("Broad","Targeted","Monthly")]$Channel = "Broad",
  [Parameter(ParameterSetName = "NoXML")][Switch]$DisplayInstall = $False,
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("Groove","Outlook","OneNote","Access","OneDrive","Publisher","Word","Excel","PowerPoint","Teams","Lync")][Array]$ExcludeApps,
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("64","32")]$OfficeArch = "64",
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("O365ProPlusRetail","O365BusinessRetail")]$OfficeEdition = "O365ProPlusRetail",
  [Parameter(ParameterSetName = "NoXML")][ValidateSet(0,1)]$SharedComputerLicensing = "0",
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("TRUE","FALSE")]$EnableUpdates = "TRUE",
  [Parameter(ParameterSetName = "NoXML")][String]$LoggingPath,
  [Parameter(ParameterSetName = "NoXML")][String]$SourcePath,
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("TRUE","FALSE")]$PinItemsToTaskbar = "TRUE",
  [Parameter(ParameterSetName = "NoXML")][Switch]$KeepMSI = $False,
  [String]$OfficeInstallDownloadPath = "C:\Scripts\Office365Install"
)

$VerbosePreference = "Continue"
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If(!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))){
    Write-Warning "Script is not running as Administrator"
    Write-Warning "Please rerun this script as Administrator."
    Exit
}

If(-Not(Test-Path $OfficeInstallDownloadPath )){
  New-Item -Path $OfficeInstallDownloadPath  -ItemType Directory -ErrorAction Stop | Out-Null
}

Function Generate-XMLFile{

  If($ExcludeApps){
    $ExcludeApps | ForEach-Object{
      $ExcludeAppsString += "<ExcludeApp ID =`"$_`" />"
    }
  }

  If($OfficeArch){
    $OfficeArchString = "`"$OfficeArch`""
  }

  If($KeepMSI){
    $RemoveMSIString = $Null
  }Else{
    $RemoveMSIString =  "<RemoveMSI />"
  }

  If($Channel){
    $ChannelString = "Channel=`"$Channel`""
  }Else{
    $ChannelString = $Null
  }

  If($SourcePath){
    $SourcePathString = "SourcePath=`"$SourcePath`"" 
  }Else{
    $SourcePathString = $Null
  }

  If($DisplayInstall){
    $SilentInstallString = "Full"
  }Else{
    $SilentInstallString = "None"
  }

  If($LoggingPath){
    $LoggingString = "<Logging Level=`"Standard`" Path=`"$LoggingPath`" />"
  }Else{
    $LoggingString = $Null
  }
  #XML data that will be used for the download/install
  $OfficeXML = [XML]@"
  <Configuration>
    <Add OfficeClientEdition=$OfficeArchString $ChannelString $SourcePathString  >
      <Product ID="$OfficeEdition">
        <Language ID="MatchOS" />
        $ExcludeAppsString
      </Product>
    </Add>  
    <Property Name="PinIconsToTaskbar" Value="$PinItemsToTaskbar" />
    <Property Name="SharedComputerLicensing" Value="$SharedComputerlicensing" />
    <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
    <Updates Enabled="$EnableUpdates" />
    $RemoveMSIString
    $LoggingString
  </Configuration>
"@
  #Save the XML file
  $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")
  Return "$OfficeInstallDownloadPath\OfficeInstall.xml"
}
If(!($ConfiguratonXMLFile)){

  $ConfiguratonXMLFile = Generate-XMLFile

}Else{
  If(!(Test-Path $ConfiguratonXMLFile)){
    Write-Warning "The configuration XML file is not a valid file"
    Write-Warning "Please check the path and try again"
    Exit
  }
}

$ErrorActionPreference = "Stop"
$ODTInstallLink = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_12130-20272.exe"

#Download the Office Deployment Tool
Write-Verbose "Downloading the Office Deployment Tool..."
Try{
  Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"
}Catch{
  Write-Warning "There was an error downloading the Office Deployment Tool."
  Write-Warning "Please verify the below link is valid:"
  Write-Warning $ODTInstallLink
  Exit
}

#Run the Office Deployment Tool
Try{
  Write-Verbose "Running the Office Deployment Tool..."
  Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait
}Catch{
  Write-Warning "Error running the Office Deployment Tool. The error is below:"
  Write-Warning $_
}

#Run the install
Try{
  Write-Verbose "Downloading and installing Office 365"
  $OfficeInstall = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure $ConfiguratonXMLFile" -Wait -PassThru
}Catch{
  Write-Warning "Error running the Office install. The error is below:"
  Write-Warning $_
}

$RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                  )
$OfficeInstalled = $False
Foreach ($Key in (Get-ChildItem $RegLocations) ) {
  If($Key.GetValue("DisplayName") -like "*Office 365*") {
    $OfficeVersionInstalled = $Key.GetValue("DisplayName")
    $OfficeInstalled = $True
  }
}
If($OfficeInstalled){
  Write-Verbose "$($OfficeVersionInstalled) installed successfully!"
}Else{
  Write-Warning "Office 365 was not detected after the install ran"
}


