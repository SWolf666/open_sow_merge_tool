param(
  [string]$BuildDir = '',
  [string]$Version = '2026-09-04.update92'
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
$buildRoot = if ($BuildDir) { [IO.Path]::GetFullPath($BuildDir) } else { Join-Path $repo 'artifacts\build' }
$exe = Join-Path $buildRoot 'dist\sow_merge_tool.exe'
if (-not (Test-Path -LiteralPath $exe)) { throw "Build output missing: $exe" }

$slug = $Version -replace '[^A-Za-z0-9._-]', '_'
$release = Join-Path $repo "artifacts\release\$slug"
if (Test-Path -LiteralPath $release) { Remove-Item -LiteralPath $release -Recurse -Force }
New-Item -ItemType Directory -Force -Path $release | Out-Null
Copy-Item -LiteralPath $exe -Destination (Join-Path $release 'sow_merge_tool.exe')
$releaseFiles = @()
# Resolve owned documentation/scripts by extension and size so this remains
# reliable under legacy Windows PowerShell code pages.
$releaseFiles += @(Get-ChildItem -LiteralPath $repo -File -Filter '*.bat')
$markdownFiles = @(Get-ChildItem -LiteralPath $repo -File -Filter '*.md' | Sort-Object Length)
if ($markdownFiles.Count -ge 2) { $releaseFiles += $markdownFiles[1] }
foreach ($file in $releaseFiles) {
  Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $release $file.Name)
}
$hash = (Get-FileHash -LiteralPath (Join-Path $release 'sow_merge_tool.exe') -Algorithm SHA256).Hash
Set-Content -LiteralPath (Join-Path $release 'SHA256SUMS.txt') -Value "$hash  sow_merge_tool.exe" -Encoding ASCII
$manifest = [ordered]@{
  version = $Version
  package = "sow_merge_tool_$slug.zip"
  directory = $slug
  files = @('sow_merge_tool.exe') + @($releaseFiles.Name) + @('SHA256SUMS.txt')
  sha256 = $hash
  generatedAt = (Get-Date).ToUniversalTime().ToString('o')
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $repo 'artifacts\release\latest.json') -Encoding UTF8
$zip = Join-Path $repo "artifacts\release\sow_merge_tool_$slug.zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $release '*') -DestinationPath $zip -Force
Write-Host "Package: $zip" -ForegroundColor Green
Write-Host "SHA256: $hash" -ForegroundColor Green
