@echo off
setlocal
set "SOW_BAT_FILE=%~s0"
set "SOW_BAT_ARGS=%*"
set "SOW_TOOL_PATH=%~dp0sow_merge_tool.exe"
if not exist "%SOW_TOOL_PATH%" set "SOW_TOOL_PATH=%~dp0artifacts\build\dist\sow_merge_tool.exe"
set "SOW_PS_FILE=%TEMP%\SowMergeInstall_%RANDOM%.ps1"

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
$tool = [IO.Path]::GetFullPath($env:SOW_TOOL_PATH)
if (-not (Test-Path -LiteralPath $tool)) { throw "未找到 sow_merge_tool.exe：$tool" }

$stateRoot = Join-Path $env:LOCALAPPDATA 'SowMergeTool'
$statePath = Join-Path $stateRoot 'install-state.json'
New-Item -ItemType Directory -Force -Path $stateRoot | Out-Null
$quote = [char]34
$label = [string]::Concat([char[]](0x591A,0x5206,0x652F,0x20,0x53,0x56,0x4E,0x20,0x63D0,0x4EA4))
$diff = $quote + $tool + $quote + ' --base ' + $quote + '%base' + $quote + ' --mine ' + $quote + '%mine' + $quote + ' --title ' + $quote + '%bname' + $quote
$merge = $quote + $tool + $quote + ' --base ' + $quote + '%base' + $quote + ' --mine ' + $quote + '%mine' + $quote + ' --theirs ' + $quote + '%theirs' + $quote + ' --merged ' + $quote + '%merged' + $quote + ' --title ' + $quote + '%bname' + $quote
$settings = @(
  @{ Path='HKCU:\Software\TortoiseSVN\DiffTools'; Name='.xlsx'; Value=$diff },
  @{ Path='HKCU:\Software\TortoiseSVN\DiffTools'; Name='.xlsm'; Value=$diff },
  @{ Path='HKCU:\Software\TortoiseSVN\MergeTools'; Name='.xlsx'; Value=$merge },
  @{ Path='HKCU:\Software\TortoiseSVN\MergeTools'; Name='.xlsm'; Value=$merge },
  @{ Path='HKCU:\Software\TortoiseSVN\DiffTools\XLSX'; Name='command'; Value=$tool },
  @{ Path='HKCU:\Software\TortoiseSVN\DiffTools\XLSX'; Name='args'; Value='--base %base --mine %mine --title %bname' },
  @{ Path='HKCU:\Software\TortoiseSVN\DiffTools\XLSM'; Name='command'; Value=$tool },
  @{ Path='HKCU:\Software\TortoiseSVN\DiffTools\XLSM'; Name='args'; Value='--base %base --mine %mine --title %bname' },
  @{ Path='HKCU:\Software\TortoiseSVN\MergeTools\XLSX'; Name='command'; Value=$tool },
  @{ Path='HKCU:\Software\TortoiseSVN\MergeTools\XLSX'; Name='args'; Value='--base %base --mine %mine --theirs %theirs --merged %merged --title %bname' },
  @{ Path='HKCU:\Software\TortoiseSVN\MergeTools\XLSM'; Name='command'; Value=$tool },
  @{ Path='HKCU:\Software\TortoiseSVN\MergeTools\XLSM'; Name='args'; Value='--base %base --mine %mine --theirs %theirs --merged %merged --title %bname' }
)
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

function Get-RegistryEntry($setting) {
  $subKeyName = $setting.Path.Substring(6)
  $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($subKeyName, $false)
  if ($null -eq $key -or $key.GetValueNames() -notcontains $setting.Name) {
    if ($key) { $key.Dispose() }
    return [ordered]@{ path=$setting.Path; name=$setting.Name; existed=$false; original=$null; installed=$setting.Value }
  }
  $original = [string]$key.GetValue($setting.Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
  $key.Dispose()
  return [ordered]@{ path=$setting.Path; name=$setting.Name; existed=$true; original=$original; installed=$setting.Value }
}

function Set-RegistryString($setting) {
  $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($setting.Path.Substring(6))
  try { $key.SetValue($setting.Name, [string]$setting.Value, [Microsoft.Win32.RegistryValueKind]::String) }
  finally { $key.Dispose() }
}

if (-not $contextOnly) {
  $state = $null
  if (Test-Path -LiteralPath $statePath) {
    try { $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $state = $null }
  }
  if ($null -eq $state -or $state.tool -ne $tool) {
    $state = [ordered]@{ version=1; tool=$tool; installedAt=(Get-Date).ToUniversalTime().ToString('o'); entries=@() }
    foreach ($setting in $settings) { $state.entries += [pscustomobject](Get-RegistryEntry $setting) }
  } else {
    $state.tool = $tool
  }
  foreach ($setting in $settings) {
    Set-RegistryString $setting
  }
  $state | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $statePath -Encoding UTF8
}

foreach ($path in $contextPaths) {
  New-Item -ItemType Directory -Force -Path $path | Out-Null
  Set-Item -LiteralPath $path -Value $label
  New-ItemProperty -LiteralPath $path -Name Icon -PropertyType String -Value ($quote + $tool + $quote + ',0') -Force | Out-Null
  New-ItemProperty -LiteralPath $path -Name Position -PropertyType String -Value Top -Force | Out-Null
  if ($path -like '*SystemFileAssociations*') { New-ItemProperty -LiteralPath $path -Name MultiSelectModel -PropertyType String -Value Player -Force | Out-Null }
  $token = if ($path -like '*Background*') { '%V' } else { '%1' }
  $commandPath = Join-Path $path 'command'
  New-Item -ItemType Directory -Force -Path $commandPath | Out-Null
  Set-Item -LiteralPath $commandPath -Value ($quote + $tool + $quote + ' --branch-submit ' + $quote + $token + $quote)
}
$compareLabel = 'Excel 文件比较/合并'
foreach ($path in $compareContextPaths) {
  New-Item -ItemType Directory -Force -Path $path | Out-Null
  Set-Item -LiteralPath $path -Value $compareLabel
  New-ItemProperty -LiteralPath $path -Name Icon -PropertyType String -Value ($quote + $tool + $quote + ',0') -Force | Out-Null
  New-ItemProperty -LiteralPath $path -Name Position -PropertyType String -Value Top -Force | Out-Null
  if ($path -like '*SystemFileAssociations*') { New-ItemProperty -LiteralPath $path -Name MultiSelectModel -PropertyType String -Value Player -Force | Out-Null }
  $token = if ($path -like '*Background*') { '%V' } else { '%1' }
  $commandPath = Join-Path $path 'command'
  New-Item -ItemType Directory -Force -Path $commandPath | Out-Null
  Set-Item -LiteralPath $commandPath -Value ($quote + $tool + $quote + ' --compare ' + $quote + $token + $quote)
}
if ($contextOnly) { Write-Host '右键菜单安装完成。' } else { Write-Host '安装完成，已保存原有 TortoiseSVN 配置。' }
# POWERSHELL-END
