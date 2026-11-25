<#
  Removes the USGS StreamStats tool files from a MicroStation Macros folder.
  Python libraries are left installed because other tools may use them.
#>
[CmdletBinding()]
param([string]$MacrosDir, [switch]$Quiet)
$toolFiles = @("USGS_StreamStats_FenceQuery.py", "USGS_StreamStats_VDOT.mvba")
if (-not $MacrosDir) {
    $root = "$env:ProgramData\Bentley\MicroStation 2026\Configuration"
    $hits = Get-ChildItem $root -Recurse -Filter "USGS_StreamStats_FenceQuery.py" -ErrorAction SilentlyContinue |
            ForEach-Object { $_.DirectoryName } | Sort-Object -Unique
    if (-not $hits) { Write-Host "The tool was not found under $root." -ForegroundColor Yellow; if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }; exit 0 }
    $hits = @($hits)
    if ($hits.Count -eq 1) { $MacrosDir = $hits[0] }
    else {
        for ($i = 0; $i -lt $hits.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i + 1), $hits[$i]) }
        $MacrosDir = $hits[[int](Read-Host "Number of the folder to remove it from") - 1]
    }
}
foreach ($f in $toolFiles) {
    $p = Join-Path $MacrosDir $f
    if (Test-Path $p) { Remove-Item $p -Force; Write-Host "Removed $p" -ForegroundColor Green }
}
Write-Host "Done. Backups (*.bak-*) and Python libraries were left in place."
if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }
