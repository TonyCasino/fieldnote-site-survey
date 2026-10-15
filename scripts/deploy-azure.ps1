param(
  [string]$ResourceGroup = "fieldnote-site-survey-rg",
  [string]$AppName = "fieldnote-site-survey-3mdtech"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$versionPath = Join-Path $projectRoot "public/version.json"
$outputPath = Join-Path $projectRoot ".next-azure"
$buildId = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds().ToString()
$az = "C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd"
if (-not (Test-Path $az)) { $az = "az" }
$env:AZURE_CONFIG_DIR = Join-Path $env:TEMP "fieldnote-azure-cli"

Push-Location $projectRoot
try {
  $env:AZURE_STATIC_EXPORT = "1"
  $env:NEXT_PUBLIC_FIELDNOTE_BUILD_ID = $buildId
  [System.IO.File]::WriteAllText($versionPath, "{`"buildId`":`"$buildId`"}", [System.Text.UTF8Encoding]::new($false))
  & node node_modules/next/dist/bin/next build
  if ($LASTEXITCODE -ne 0) { throw "Azure build failed." }
  [System.IO.File]::WriteAllText((Join-Path $outputPath "version.json"), "{`"buildId`":`"$buildId`"}", [System.Text.UTF8Encoding]::new($false))
  Copy-Item -LiteralPath (Join-Path $projectRoot "public/staticwebapp.config.json") -Destination (Join-Path $outputPath "staticwebapp.config.json") -Force
  $token = & $az staticwebapp secrets list --resource-group $ResourceGroup --name $AppName --query properties.apiKey -o tsv
  if ($LASTEXITCODE -ne 0 -or -not $token) { throw "Azure deployment token could not be loaded." }
  & (Join-Path $projectRoot "node_modules/.bin/swa.cmd") deploy $outputPath --api-location api --api-language node --api-version 20 --deployment-token $token --env production
  if ($LASTEXITCODE -ne 0) { throw "Azure deployment failed." }
} finally {
  if (Test-Path -LiteralPath $versionPath) { Remove-Item -LiteralPath $versionPath -Force }
  Pop-Location
}
