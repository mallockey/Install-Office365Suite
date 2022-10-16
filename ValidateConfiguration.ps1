param(
  [Parameter(ParameterSetName = 'XMLFile')][String]$ConfigurationXMLFilePath
)

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

Import-Module './InstallOffice.psm1'

try {
  if (!$ConfigurationXMLFilePath) {
    $OfficeXML = Get-XMLFile
  }
  else {
    $OfficeXML = Get-Content -Path $ConfigurationXMLFilePath
  }
}
catch {
  Write-Verbose 'There was an error generating the XML config file'
}

try {
  Write-Verbose 'Uploading XML file to clients.config.office.net...'

  Invoke-RestMethod -Uri 'https://clients.config.office.net/intents/v1.0/DeploymentSettings/ImportConfiguration' `
    -Method Post  `
    -Body $OfficeXML `
    -ContentType 'text/xml'

  Write-Verbose 'XML File imported successfully'
}
catch {
  Write-Error 'The XML is not formatted correctly. Check '
} 