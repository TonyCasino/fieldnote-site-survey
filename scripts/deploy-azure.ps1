param(
  [string]$ResourceGroup = "fieldnote-site-survey-rg",
  [string]$AppName = "fieldnote-site-survey-3mdtech"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$az = "C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd"
if (-not (Test-Path $az)) { $az = "az" }
$env:AZURE_CONFIG_DIR = Join-Path $env:TEMP "fieldnote-azure-cli"

Push-Location $projectRoot
try {
  $env:AZURE_STATIC_EXPORT = "1"
  & node node_modules/next/dist/bin/next build
  if ($LASTEXITCODE -ne 0) { throw "Azure build failed." }
  $token = & $az staticwebapp secrets list --resource-group $ResourceGroup --name $AppName --query properties.apiKey -o tsv
  if ($LASTEXITCODE -ne 0 -or -not $token) { throw "Azure deployment token could not be loaded." }
  & (Join-Path $projectRoot "node_modules/.bin/swa.cmd") deploy out --api-location api --api-language node --api-version 20 --deployment-token $token --env production
  if ($LASTEXITCODE -ne 0) { throw "Azure deployment failed." }
} finally {
  Pop-Location
}
