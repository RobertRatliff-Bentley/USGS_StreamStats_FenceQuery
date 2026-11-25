<#
  USGS StreamStats & VDOT Hydrology - Fence Query : installer
  Created by Robert Ratliff

  1. Finds MicroStation's Python (PowerPlatformPython).
  2. Installs the Python libraries from the bundled wheels (offline); falls back to the internet if needed.
  3. Copies the Python tool and the VBA project into a MicroStation Macros folder.
  4. Verifies the libraries import.

  Usage:  Install.bat                         (interactive)
          Install.ps1 -MacrosDir "<folder>"   (no folder prompt)
          Install.ps1 -SkipLibraries          (copy files only)
#>
[CmdletBinding()]
param(
    [string]$MacrosDir,
    [string]$PythonExe,
    [switch]$SkipLibraries,
    [switch]$Quiet
)
$ErrorActionPreference = "Continue"   # pip writes warnings to stderr; failures are checked with $LASTEXITCODE
$here   = Split-Path -Parent $MyInvocation.MyCommand.Path
$wheels = Join-Path $here "wheels"
$files  = Join-Path $here "files"
$req    = Join-Path $here "requirements.txt"
$log    = Join-Path $env:TEMP "USGS_StreamStats_Install.log"
$toolFiles = @("USGS_StreamStats_FenceQuery.py", "USGS_StreamStats_VDOT.mvba")
"Install started $(Get-Date)" | Set-Content $log

function Say($msg, $color = "Gray") { Write-Host $msg -ForegroundColor $color; Add-Content $log $msg }
function Fail($msg) { Say "ERROR: $msg" Red; if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }; exit 1 }

Say ""
Say "USGS StreamStats & VDOT Hydrology - Fence Query installer" Cyan
Say "Created by Robert Ratliff" Cyan
Say ""

# ---- 1. MicroStation Python -------------------------------------------------------------
if (-not $PythonExe) {
    $candidates = @(
        "$env:ProgramData\Bentley\PowerPlatformPython\python\python.exe",
        "$env:ProgramFiles\Bentley\PowerPlatformPython\python\python.exe"
    )
    $PythonExe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $PythonExe -or -not (Test-Path $PythonExe)) {
    Fail "MicroStation's Python was not found. Install MicroStation 2026 Python support (PowerPlatformPython), or run Install.ps1 -PythonExe <path to python.exe>."
}
$pyVer = & $PythonExe -c "import sys;print('%d.%d' % sys.version_info[:2])"
Say "[1/4] MicroStation Python: $PythonExe (Python $pyVer)" Green
if ($pyVer -ne "3.12") {
    Say "      Warning: the bundled libraries are built for Python 3.12; the installer will download matching ones from the internet instead." Yellow
}

# ---- 2. Python libraries ----------------------------------------------------------------
if ($SkipLibraries) {
    Say "[2/4] Skipping Python libraries (-SkipLibraries)." Yellow
} else {
    if (Get-Process -Name ustation -ErrorAction SilentlyContinue) {
        Say "      MicroStation is running. Close it first so library files are not locked." Yellow
        if (-not $Quiet) { Read-Host "      Close MicroStation, then press Enter to continue" | Out-Null }
    }
    & $PythonExe -c "import pip" 2>$null
    if ($LASTEXITCODE -ne 0) {
        Say "      pip is missing; installing it from the bundled wheel..."
        $pipWhl = Get-ChildItem $wheels -Filter "pip-*.whl" | Select-Object -First 1
        & $PythonExe "$($pipWhl.FullName)\pip" install --no-index "$($pipWhl.FullName)" 2>&1 | ForEach-Object { "$_" } | Add-Content $log
    }
    # Install for all users when site-packages is writable, otherwise for the current user.
    $site = & $PythonExe -c "import sysconfig;print(sysconfig.get_paths()['purelib'])"
    $scope = @()
    try { $t = Join-Path $site "_ss_write_test"; [IO.File]::WriteAllText($t, "x"); Remove-Item $t } catch { $scope = @("--user") }
    Say "[2/4] Installing Python libraries ($(if ($scope) {'current user'} else {'all users'}))..." Green
    $pipArgs = @("-m", "pip", "install", "--disable-pip-version-check", "--no-warn-script-location") + $scope + @("-r", $req)
    & $PythonExe @pipArgs --no-index --find-links $wheels 2>&1 | ForEach-Object { "$_" } | Add-Content $log
    if ($LASTEXITCODE -ne 0) {
        Say "      Offline install did not finish; trying the internet (pypi.org)..." Yellow
        & $PythonExe @pipArgs 2>&1 | ForEach-Object { "$_" } | Add-Content $log
        if ($LASTEXITCODE -ne 0) { Fail "Python libraries could not be installed. See $log" }
    }
}

# ---- 3. Copy tool files -----------------------------------------------------------------
if (-not $MacrosDir) {
    $root = "$env:ProgramData\Bentley\MicroStation 2026\Configuration"
    $found = @()
    if (Test-Path $root) {
        $found += Get-ChildItem "$root\WorkSpaces\*\WorkSets\*\Standards\Macros" -Directory -ErrorAction SilentlyContinue
        $found += Get-ChildItem "$root\WorkSpaces\*\Standards\Macros" -Directory -ErrorAction SilentlyContinue
        $found += Get-ChildItem "$root\Organization\Macros" -Directory -ErrorAction SilentlyContinue
    }
    $found = @($found | Sort-Object FullName -Unique)
    if ($found.Count -eq 1 -or ($Quiet -and $found.Count -ge 1)) {
        $MacrosDir = $found[0].FullName
    } elseif ($found.Count -gt 1) {
        Say "      Choose the MicroStation Macros folder to install into:"
        for ($i = 0; $i -lt $found.Count; $i++) { Say ("        {0}. {1}" -f ($i + 1), $found[$i].FullName) }
        $pick = Read-Host "      Number (or type a folder path)"
        if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $found.Count) { $MacrosDir = $found[[int]$pick - 1].FullName }
        else { $MacrosDir = $pick }
    } else {
        $MacrosDir = Read-Host "      No Macros folder was found. Type the folder to install into"
    }
}
if (-not $MacrosDir) { Fail "No Macros folder was chosen." }
try { New-Item -ItemType Directory -Path $MacrosDir -Force -ErrorAction Stop | Out-Null } catch { Fail "Cannot create $MacrosDir. $_" }
Say "[3/4] Copying tool files to $MacrosDir" Green
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
foreach ($f in $toolFiles) {
    $dest = Join-Path $MacrosDir $f
    if (Test-Path $dest) {
        Copy-Item $dest "$dest.bak-$stamp" -Force -ErrorAction SilentlyContinue
        Say "      Backed up existing $f -> $f.bak-$stamp"
    }
    try { Copy-Item (Join-Path $files $f) $dest -Force -ErrorAction Stop } catch { Fail "Could not copy $f (is it open in MicroStation?). $_" }
    Say "      $f"
}

# ---- 4. Verify --------------------------------------------------------------------------
Say "[4/4] Checking libraries..." Green
$check = & $PythonExe -W ignore -c "import importlib`nfor m in ('pyproj','shapely','openpyxl','numpy','certifi'):`n  try:`n    x=importlib.import_module(m); print('  OK      %-9s %s' % (m, x.__version__))`n  except Exception as e: print('  MISSING %-9s %s' % (m, e))" 2>&1
$check | ForEach-Object { Say "    $_" }

Say ""
Say "Done. Log: $log" Cyan
Say "  Python tool:  key in   python load `"$MacrosDir\USGS_StreamStats_FenceQuery.py`"" Cyan
Say "  VBA tool:     key in   vba load USGS_StreamStats_VDOT   then   vba run StreamStatsShow" Cyan
if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }


