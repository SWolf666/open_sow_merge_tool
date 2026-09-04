param(
  [ValidateSet('Fast', 'Full', 'Integration', 'Native', 'Visual', 'Adversarial')]
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
$initialMergePids = @(
  Get-Process -Name 'sow_merge_tool' -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Id
)
if ($Profile -in @('Native', 'Visual') -and $initialMergePids.Count -gt 0) {
  throw "Native/Visual 测试前发现已有 sow_merge_tool.exe 实例，请先关闭后再运行；工具不会强杀业务实例。"
}

function Assert-NoUnexpectedMergeProcess {
  if ($Profile -notin @('Fast', 'Full', 'Integration', 'Adversarial')) { return }
  $current = @(
    Get-Process -Name 'sow_merge_tool' -ErrorAction SilentlyContinue |
      Select-Object -ExpandProperty Id
  )
  $unexpected = @($current | Where-Object { $_ -notin $initialMergePids })
  if ($unexpected.Count -gt 0) {
    throw "无界面 profile 意外启动了 sow_merge_tool.exe 实例：$($unexpected -join ', ')"
  }
}

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
$visualSmokeNames = @(
  '_smoke_test_2way_formula_cache_save.py',
  '_smoke_test_2way_row_replay.py',
  '_smoke_test_3way_alignment.py',
  '_smoke_test_3way_only_diff_base_insert.py',
  '_smoke_test_3way_pristine_base_tail_block.py',
  '_smoke_test_3way_tail_append_split.py',
  '_smoke_test_blank_shared_formula_b_save.py',
  '_smoke_test_cursor_block.py',
  '_smoke_test_difference_browser_real_workbooks.py',
  '_smoke_test_excel_com_blank_cell.py',
  '_smoke_test_formula_cache_undo.py',
  '_smoke_test_large_3way_merge_open.py',
  '_smoke_test_large_3way_only_diff.py',
  '_smoke_test_large_only_diff_row_insert.py',
  '_smoke_test_manual_merge_row_insert.py',
  '_smoke_test_only_diff_minimap.py',
  '_smoke_test_recovery_progress_close.py',
  '_smoke_test_sheet_level_ops.py',
  '_smoke_test_xlsm_support.py',
  '_smoke_test.py'
)
$selectedSmoke = if ($Profile -eq 'Full') {
  @($allSmoke | Where-Object { $_.Name -notin $visualSmokeNames })
} elseif ($Profile -eq 'Fast') {
  @($allSmoke | Where-Object { $_.Name -in $fastSmokeNames -and $_.Name -notin $visualSmokeNames })
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
foreach ($test in $selectedSmoke) {
  Invoke-PythonFile $test.FullName
  Assert-NoUnexpectedMergeProcess
}

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

if ($Profile -eq 'Visual') {
  Write-Warning 'Visual profile 会打开真实 Tk/Win32 窗口；仅使用临时验收数据，不要与业务实例同时运行。'
  $visualFiles = @(
    '_gui_self_test_branch_submit_workbench.py',
    '_gui_self_test_merge_file_paths.py',
    '_gui_self_test_start_center.py',
    '_gui_self_test_comparison_sessions.py',
    '_gui_self_test_logical_column_actions.py'
  )
  foreach ($name in $visualFiles) {
    $path = Join-Path $testScriptRoot $name
    if (-not (Test-Path -LiteralPath $path)) { throw "Visual manifest contains missing file: $name" }
    Invoke-PythonFile $path
  }
}

Assert-NoUnexpectedMergeProcess
if ($Profile -in @('Native', 'Visual')) {
  $remaining = @(Get-Process -Name 'sow_merge_tool' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  if ($remaining.Count -gt 0) {
    Write-Warning "Native/Visual 测试结束仍有 sow_merge_tool.exe 实例：$($remaining -join ', ')；未强杀，请人工确认。"
  }
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
