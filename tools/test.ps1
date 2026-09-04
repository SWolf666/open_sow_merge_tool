param(
  [ValidateSet('Fast', 'Full', 'Integration', 'Native', 'Adversarial')]
  [string]$Profile = 'Fast',
  [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'
$managedEnvironmentNames = @(
  'LOCALAPPDATA', 'SOW_TEST_TMPDIR', 'PYTHONUTF8', 'SOW_SKIP_REAL_WC_TESTS', 'SOW_SVN_BIN'
)
$originalEnvironment = @{}
$originalEnvironmentPresent = @{}
foreach ($name in $managedEnvironmentNames) {
  $item = Get-Item -LiteralPath "Env:$name" -ErrorAction SilentlyContinue
  $originalEnvironmentPresent[$name] = $null -ne $item
  $originalEnvironment[$name] = if ($item) { $item.Value } else { $null }
}

try {
Set-Location (Split-Path -Parent $PSScriptRoot)

$repo = (Get-Location).Path
$python = Join-Path $repo '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) {
  throw "Virtual environment not found: $python. Run tools\bootstrap.ps1 first."
}

# Keep synthetic fixtures outside the system temp root. The application uses
# the system temp root as one signal for SVN-generated sidecars; putting test
# fixtures there changes the behavior the tests are meant to observe.
$testRoot = Join-Path $repo 'tmp\test_tmp'
New-Item -ItemType Directory -Force -Path $testRoot | Out-Null
$testAppData = Join-Path $testRoot 'appdata'
New-Item -ItemType Directory -Force -Path $testAppData | Out-Null
$env:LOCALAPPDATA = $testAppData
$env:SOW_TEST_TMPDIR = $testRoot
$env:PYTHONUTF8 = '1'
$env:SOW_SKIP_REAL_WC_TESTS = if ($Profile -eq 'Native') { '0' } else { '1' }

function Invoke-PythonFile {
  param([string]$Path)
  $psi = [Diagnostics.ProcessStartInfo]::new()
  $psi.FileName = $python
  $psi.Arguments = '"' + $Path + '"'
  $psi.WorkingDirectory = $repo
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.Environment['SOW_TEST_TMPDIR'] = $testRoot
  $psi.Environment['LOCALAPPDATA'] = $testAppData
  $psi.Environment['PYTHONUTF8'] = '1'
  $psi.Environment['SOW_SKIP_REAL_WC_TESTS'] = $env:SOW_SKIP_REAL_WC_TESTS
  if ($env:SOW_SVN_BIN) { $psi.Environment['SOW_SVN_BIN'] = $env:SOW_SVN_BIN }
  $process = [Diagnostics.Process]::new()
  $process.StartInfo = $psi
  [void]$process.Start()
  $watch = [Diagnostics.Stopwatch]::StartNew()
  $finished = $process.WaitForExit($TimeoutSeconds * 1000)
  if (-not $finished) {
    $process.Kill($true)
    $process.WaitForExit()
    throw "Timeout after $TimeoutSeconds seconds: $Path"
  }
  $watch.Stop()
  $stdout = $process.StandardOutput.ReadToEnd()
  $stderr = $process.StandardError.ReadToEnd()
  if ($process.ExitCode -ne 0) {
    throw "Failed ($($process.ExitCode)) $Path`n$stdout`n$stderr"
  }
  Write-Host ("PASS {0} ({1:N2}s)" -f (Split-Path -Leaf $Path), $watch.Elapsed.TotalSeconds) -ForegroundColor Green
}

$testScriptRoot = Join-Path $repo 'tests\regression'
$allSmoke = @(Get-ChildItem -LiteralPath $testScriptRoot -File -Filter '_smoke_test*.py' | Sort-Object Name)
$fastSmokeNames = @(
  '_smoke_test_branch_submit.py',
  '_smoke_test_fast_branch_analysis.py',
  '_smoke_test_automatic_merge_semantics.py',
  '_smoke_test_save_and_diff_fidelity.py',
  '_smoke_test_formula_cache_undo.py',
  '_smoke_test_svn_conflict_detection.py',
  '_smoke_test_svn_merge_role_semantics.py'
)
$selectedSmoke = if ($Profile -eq 'Full') {
  $allSmoke
} elseif ($Profile -eq 'Fast') {
  @($allSmoke | Where-Object { $_.Name -in $fastSmokeNames })
} elseif ($Profile -eq 'Adversarial') {
  @($allSmoke | Where-Object { $_.Name -in @(
    '_smoke_test_branch_submit.py',
    '_smoke_test_svn_conflict_detection.py',
    '_smoke_test_svn_merge_role_semantics.py'
  ) })
} else {
  @()
}
if ($Profile -eq 'Fast') {
  $missingFastSmoke = @($fastSmokeNames | Where-Object { $_ -notin $allSmoke.Name })
  if ($missingFastSmoke.Count -gt 0) {
    throw "Fast smoke manifest contains missing files: $($missingFastSmoke -join ', ')"
  }
}
foreach ($test in $selectedSmoke) { Invoke-PythonFile $test.FullName }

if ($Profile -in @('Fast', 'Full', 'Adversarial')) {
  $pytest = Join-Path $repo '.venv\Scripts\pytest.exe'
  if (-not (Test-Path -LiteralPath $pytest)) { throw "pytest not found: $pytest" }
  & $pytest -q
  if ($LASTEXITCODE -ne 0) { throw "pytest failed with exit code $LASTEXITCODE" }
}

if ($Profile -in @('Full', 'Integration', 'Adversarial')) {
  $svnBin = (& (Join-Path $repo 'tools\setup_svn_test_runtime.ps1') | Select-Object -Last 1)
  if (-not $svnBin) { throw 'SVN test runtime setup returned no bin path.' }
  $env:SOW_SVN_BIN = [string]$svnBin
  if ($Profile -eq 'Integration') {
    $pytest = Join-Path $repo '.venv\Scripts\pytest.exe'
    if (-not (Test-Path -LiteralPath $pytest)) { throw "pytest not found: $pytest" }
    & $pytest -q (Join-Path $repo 'tests\unit\test_svn_status_policy.py')
    if ($LASTEXITCODE -ne 0) { throw "SVN policy matrix failed with exit code $LASTEXITCODE" }
  }
  Invoke-PythonFile (Join-Path $repo 'tests\integration\_integration_test_svn_headless_end_to_end.py')
}

if ($Profile -eq 'Native') {
  & $python (Join-Path $testScriptRoot '_gui_self_test_branch_submit_workbench.py')
  if ($LASTEXITCODE -ne 0) { throw "Native GUI test failed with exit code $LASTEXITCODE" }
  & $python (Join-Path $testScriptRoot '_gui_self_test_merge_file_paths.py')
  if ($LASTEXITCODE -ne 0) { throw "Native merge-path GUI test failed with exit code $LASTEXITCODE" }
  & $python (Join-Path $testScriptRoot '_gui_self_test_start_center.py')
  if ($LASTEXITCODE -ne 0) { throw "Native start-centre GUI test failed with exit code $LASTEXITCODE" }
  & $python (Join-Path $testScriptRoot '_gui_self_test_comparison_sessions.py')
  if ($LASTEXITCODE -ne 0) { throw "Native comparison-session GUI test failed with exit code $LASTEXITCODE" }
}

Write-Host "Test profile $Profile passed." -ForegroundColor Green
} finally {
  foreach ($name in $managedEnvironmentNames) {
    if ($originalEnvironmentPresent[$name]) {
      Set-Item -LiteralPath "Env:$name" -Value $originalEnvironment[$name]
    } else {
      Remove-Item -LiteralPath "Env:$name" -ErrorAction SilentlyContinue
    }
  }
}
