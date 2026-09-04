param(
  [string]$Python311Path = '',
  [string]$Python312Path = ''
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path

function Resolve-Runtime([string]$requested, [string]$versionPrefix, [string[]]$fallbacks) {
  $candidates = @()
  if ($requested) { $candidates += $requested }
  $candidates += $fallbacks
  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) {
      $version = (& $candidate --version 2>$null | Select-Object -First 1)
      if ([string]$version -match "Python $([regex]::Escape($versionPrefix))") {
        return (Resolve-Path -LiteralPath $candidate).Path
      }
    }
  }
  throw "Python $versionPrefix runtime not found. Pass -Python$($versionPrefix.Replace('.', ''))Path explicitly."
}

$userRoot = $env:USERPROFILE
$python311 = Resolve-Runtime $Python311Path '3.11' @(
  (Join-Path $userRoot 'AppData\Roaming\uv\python\cpython-3.11.15-windows-x86_64-none\python.exe'),
  (Join-Path $userRoot 'AppData\Roaming\uv\python\cpython-3.11-windows-x86_64-none\python.exe')
)
$python312 = Resolve-Runtime $Python312Path '3.12' @(
  (Join-Path $userRoot 'AppData\Roaming\uv\python\cpython-3.12.13-windows-x86_64-none\python.exe'),
  (Join-Path $userRoot 'AppData\Roaming\uv\python\cpython-3.12-windows-x86_64-none\python.exe')
)

$oldPythonPath = $env:PYTHONPATH
$src = (Resolve-Path (Join-Path $repo 'src')).Path
$deps = Join-Path $repo '.venv\Lib\site-packages'
$env:PYTHONPATH = if ($oldPythonPath) { "$src;$deps;$oldPythonPath" } else { "$src;$deps" }

try {
  foreach ($runtime in @(@{ Label = '3.11'; Path = $python311 }, @{ Label = '3.12'; Path = $python312 })) {
    $python = $runtime.Path
    & $python -m compileall -q (Join-Path $repo 'src')
    if ($LASTEXITCODE -ne 0) { throw "Python $($runtime.Label) compileall failed." }
    & $python -c "import sow_merge_tool; assert sow_merge_tool.APP_VERSION.endswith('update93')"
    if ($LASTEXITCODE -ne 0) { throw "Python $($runtime.Label) import gate failed." }
    $help = & $python -m sow_merge_tool --help 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Python $($runtime.Label) --help gate failed: $help" }
    Write-Host "PASS Python $($runtime.Label): compile/import/--help" -ForegroundColor Green
  }
  Write-Host 'Python compatibility gate passed for 3.11 and 3.12.' -ForegroundColor Green
}
finally {
  $env:PYTHONPATH = $oldPythonPath
}
