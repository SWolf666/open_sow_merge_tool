param(
  [ValidateSet('Fast', 'Full', 'Adversarial')]
  [string]$Profile = 'Fast'
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
$python = Join-Path $repo '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) { throw 'Run tools\bootstrap.ps1 first.' }

$status = @(git status --porcelain)
if ($status.Count -gt 0) {
  throw "Quality gate requires a clean worktree:`n$($status -join "`n")"
}

& (Join-Path $repo 'tools\test.ps1') -Profile $Profile
if ($LASTEXITCODE -ne 0) { throw "Test profile $Profile failed." }

& $python -m ruff check src\sow_merge_tool\cli.py src\sow_merge_tool\__main__.py tests\unit
if ($LASTEXITCODE -ne 0) { throw 'Ruff gate failed.' }
$ruffBase = (git rev-parse 'HEAD^').Trim()
& $python (Join-Path $repo 'tools\ruff_changed.py') --base $ruffBase
if ($LASTEXITCODE -ne 0) { throw 'Changed-lines Ruff gate failed.' }

& (Join-Path $repo 'tools\build.ps1') -Clean
if ($LASTEXITCODE -ne 0) { throw 'Build gate failed.' }

$exe = Join-Path $repo 'artifacts\build\dist\sow_merge_tool.exe'
$helpProbe = @"
import subprocess
exe = r'''$exe'''
result = subprocess.run([exe, '--help'], timeout=30)
raise SystemExit(result.returncode)
"@
$helpProbe | & $python -
if ($LASTEXITCODE -ne 0) { throw 'Frozen EXE --help probe failed.' }

$tree = (git rev-parse 'HEAD^{tree}').Trim()
$gateDir = Join-Path $repo '.local\quality-gates'
New-Item -ItemType Directory -Force -Path $gateDir | Out-Null
$payload = [ordered]@{
  tree = $tree
  head = (git rev-parse HEAD).Trim()
  profile = $Profile
  exe_sha256 = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
  created_at = (Get-Date).ToUniversalTime().ToString('o')
}
$payload | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $gateDir "$tree.json") -Encoding UTF8
Write-Host "Quality gate passed for tree $tree" -ForegroundColor Green
