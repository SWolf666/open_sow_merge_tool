param(
  [string]$DeployPath = 'C:\sow_main\excel\excel_merge_tool',
  [string]$Version = '2026-09-04.update92',
  [switch]$SkipNative,
  [switch]$SkipDeploy
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
& (Join-Path $repo 'tools\verify.ps1') -Profile Full
if (-not $?) { throw 'Full quality gate failed.' }
if (-not $SkipNative) {
  & (Join-Path $repo 'tools\test.ps1') -Profile Native
  if (-not $?) { throw 'Native GUI gate failed.' }
}
& (Join-Path $repo 'tools\package.ps1') -Version $Version
if (-not $?) { throw 'Package failed.' }
if (-not $SkipDeploy) {
  & (Join-Path $repo 'tools\deploy.ps1') -DeployPath $DeployPath
  if (-not $?) { throw 'Deployment failed.' }
}
Write-Host 'Release workflow passed.' -ForegroundColor Green
