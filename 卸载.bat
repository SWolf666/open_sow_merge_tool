@echo off
setlocal
set "SOW_BAT_FILE=%~s0"
set "SOW_BAT_ARGS=%*"
set "SOW_PS_FILE=%TEMP%\SowMergeUninstall_%RANDOM%.ps1"

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$raw=[IO.File]::ReadAllText($env:SOW_BAT_FILE);$start=$raw.LastIndexOf('# POWERSHELL-BEGIN');$end=$raw.LastIndexOf('# POWERSHELL-END');$raw.Substring($start+18,$end-$start-18) | Set-Content -LiteralPath $env:SOW_PS_FILE -Encoding UTF8"
if errorlevel 1 exit /b %ERRORLEVEL%
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SOW_PS_FILE%"
set "SOW_EXIT=%ERRORLEVEL%"
del /q "%SOW_PS_FILE%" >nul 2>&1
if not "%SOW_BAT_ARGS:/quiet=%"=="%SOW_BAT_ARGS%" exit /b %SOW_EXIT%
pause
exit /b %SOW_EXIT%

# POWERSHELL-BEGIN
$ErrorActionPreference = 'Stop'
$argsText = ($env:SOW_BAT_ARGS -split '\s+') | Where-Object { $_ }
$contextOnly = $argsText -contains '/context-only'
$statePath = Join-Path (Join-Path $env:LOCALAPPDATA 'SowMergeTool') 'install-state.json'
$contextPaths = @(
  'HKCU:\Software\Classes\SystemFileAssociations\.xlsx\shell\SowMultiBranchSVNSubmit',
  'HKCU:\Software\Classes\SystemFileAssociations\.xlsm\shell\SowMultiBranchSVNSubmit',
  'HKCU:\Software\Classes\Directory\shell\SowMultiBranchSVNSubmit',
  'HKCU:\Software\Classes\Directory\Background\shell\SowMultiBranchSVNSubmit'
)
$compareContextPaths = @(
  'HKCU:\Software\Classes\SystemFileAssociations\.xlsx\shell\SowExcelCompareMerge',
  'HKCU:\Software\Classes\SystemFileAssociations\.xlsm\shell\SowExcelCompareMerge',
  'HKCU:\Software\Classes\Directory\shell\SowExcelCompareMerge',
  'HKCU:\Software\Classes\Directory\Background\shell\SowExcelCompareMerge'
)
foreach ($path in $contextPaths) { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force } }
foreach ($path in $compareContextPaths) { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force } }
if ($contextOnly) { Write-Host '右键菜单卸载完成。'; exit 0 }
if (-not (Test-Path -LiteralPath $statePath)) { Write-Host '未找到安装状态，只完成了右键菜单卸载。'; exit 0 }

$state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
$skipped = @()
function Get-RegistryValue($entry) {
  $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($entry.path.Substring(6), $false)
  if ($null -eq $key -or $key.GetValueNames() -notcontains $entry.name) {
    if ($key) { $key.Dispose() }
    return $null
  }
  $value = [string]$key.GetValue($entry.name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
  $key.Dispose()
  return $value
}
function Set-RegistryString($entry, [string]$value) {
  $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($entry.path.Substring(6))
  try { $key.SetValue($entry.name, $value, [Microsoft.Win32.RegistryValueKind]::String) }
  finally { $key.Dispose() }
}
function Remove-RegistryValue($entry) {
  $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($entry.path.Substring(6), $true)
  if ($null -ne $key) { try { $key.DeleteValue($entry.name, $false) } finally { $key.Dispose() } }
}
foreach ($entry in @($state.entries)) {
  $current = Get-RegistryValue $entry
  $hasValue = $null -ne $current
  if ($current -ne [string]$entry.installed) { $skipped += ($entry.path + ' [' + $entry.name + ']'); continue }
  if ($entry.existed) {
    Set-RegistryString $entry ([string]$entry.original)
  } elseif ($hasValue) {
    Remove-RegistryValue $entry
  }
}
if ($skipped.Count -gt 0) {
  Write-Host '检测到外部修改，已跳过这些配置，未覆盖其他工具或用户的改动：'
  $skipped | ForEach-Object { Write-Host ('  ' + $_) }
  exit 2
}
Remove-Item -LiteralPath $statePath -Force
Write-Host '卸载完成，已恢复安装前的 TortoiseSVN 配置。'
# POWERSHELL-END
