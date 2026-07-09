# RoadReviewer / Site Inspector Review Tool build script.
# Drives Excel via COM: for each requested product it creates a new workbook,
# imports every .bas from ../src, bakes the product id in (SetProduct), runs
# BuildWorkbook, and saves the .xlsm at the repo root:
#   Standard  -> RoadReviewer.xlsm
#   Inspector -> Site Inspector Review Tool.xlsm
# Requires HKCU AccessVBOM=1 (verified before running).

param(
  [string]$SrcDir = (Join-Path $PSScriptRoot '..\src'),
  [ValidateSet('Standard', 'Inspector', 'Both')]
  [string]$Product = 'Both',
  # Default output dir is the repo root. SaveAs can't land directly in the
  # build\ folder (or overwrite in the repo root) because those dirs carry an
  # `Everyone Deny DeleteSubdirectoriesAndFiles` ACE which breaks Excel's
  # temp-file -> rename pattern; each build stages into %TEMP% and
  # Copy-Item -Force overwrites into place instead.
  [string]$OutDir = (Split-Path $PSScriptRoot -Parent),
  [switch]$Visible
)

$ErrorActionPreference = 'Stop'
$XlOpenXMLWorkbookMacroEnabled = 52

function Resolve-Strict([string]$p) {
  $r = Resolve-Path -LiteralPath $p -ErrorAction SilentlyContinue
  if (-not $r) { throw "Path not found: $p" }
  return $r.ProviderPath
}

$SrcDir = Resolve-Strict $SrcDir
$OutDir = Resolve-Strict $OutDir

$importOrder = @(
  'modConstants.bas',
  'modUtil.bas',
  'modHttp.bas',
  'modBuild.bas',
  'modSources.bas',
  'modClassify.bas',
  'modGeocode.bas',
  'modImagery.bas',
  'modMaps.bas',
  'modMapImage.bas',
  'modExport.bas',
  'modExportMenu.bas'
)

# Verify all files exist before launching Excel
foreach ($f in $importOrder) {
  $p = Join-Path $SrcDir $f
  if (-not (Test-Path -LiteralPath $p)) { throw "Missing source: $p" }
}

$outNames = [ordered]@{
  'Standard'  = 'RoadReviewer.xlsm'
  'Inspector' = 'Site Inspector Review Tool.xlsm'
}
$productsToBuild = if ($Product -eq 'Both') { @('Standard', 'Inspector') } else { @($Product) }

Write-Host "Starting Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = [bool]$Visible
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable - we trust our own code; this just suppresses prompts on open

try {
  foreach ($p in $productsToBuild) {
    $outPath = Join-Path $OutDir $outNames[$p]
    Write-Host ""
    Write-Host ("=== Building product: {0} -> {1} ===" -f $p, $outPath) -ForegroundColor Cyan

    $wb = $excel.Workbooks.Add()
    $proj = $wb.VBProject
    Write-Host "Importing modules..."
    foreach ($f in $importOrder) {
      $src = Join-Path $SrcDir $f
      $comp = $proj.VBComponents.Import($src)
      Write-Host ("  imported {0,-20} (lines: {1})" -f $comp.Name, $comp.CodeModule.CountOfLines)
    }

    # Inject an error-capturing wrapper so we get a usable diagnostic
    # instead of Excel sitting silently in VBE break mode.
    $helperPath = Join-Path $PSScriptRoot 'BuildHelper.bas'
    if (Test-Path -LiteralPath $helperPath) {
      $proj.VBComponents.Import($helperPath) | Out-Null
      Write-Host "  imported BuildHelper        (build-time only)"
    }

    $errFile = Join-Path $env:TEMP 'RoadReviewer_build_error.txt'
    if (Test-Path -LiteralPath $errFile) { Remove-Item -LiteralPath $errFile -Force }

    Write-Host ("Baking product id ({0}) + running BuildWorkbook..." -f $p)
    try {
      $excel.Run('SetProduct', $p) | Out-Null
      $excel.Run('BuildWorkbookSafe') | Out-Null
    }
    catch {
      if (Test-Path -LiteralPath $errFile) {
        Write-Host "---- VBA error capture ----" -ForegroundColor Yellow
        Get-Content -LiteralPath $errFile | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        Write-Host "---------------------------" -ForegroundColor Yellow
      }
      throw
    }

    # Remove the build-time helper before saving so it isn't in the shipped workbook.
    try {
      $helper = $proj.VBComponents.Item('BuildHelper')
      $proj.VBComponents.Remove($helper)
      Write-Host "  removed BuildHelper from saved workbook"
    } catch { }

    # Excel's SaveAs writes to a temp file in the destination dir, then
    # deletes the existing destination and renames the temp file. The repo
    # root (and build\) carry an `Everyone Deny DeleteSubdirectoriesAndFiles`
    # ACE that breaks both delete steps. So stage the SaveAs into %TEMP%
    # (always writable) and Copy-Item -Force into place, which overwrites
    # without needing delete permission on the parent dir.
    $stagingPath = Join-Path $env:TEMP ('rr-build-' + [guid]::NewGuid().ToString('N') + '.xlsm')
    Write-Host "Saving to staging: $stagingPath"
    $wb.SaveAs($stagingPath, $XlOpenXMLWorkbookMacroEnabled)
    $wb.Close($false)
    Write-Host "Copying to $outPath"
    Copy-Item -LiteralPath $stagingPath -Destination $outPath -Force
    Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue
    Write-Host ("Build complete: " + $outNames[$p]) -ForegroundColor Green
  }
}
catch {
  Write-Host ("BUILD FAILED: " + $_.Exception.Message) -ForegroundColor Red
  if ($_.Exception.InnerException) {
    Write-Host ("  inner: " + $_.Exception.InnerException.Message) -ForegroundColor Red
  }
  throw
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# Post-build compile check. VBA compiles each module lazily, so a syntax error
# in a module the build never executes (modHttp / modClassify) would otherwise
# only surface at the user's first Check Roads. compile-check.ps1 force-compiles
# those modules now (with a timeout so a compile-error modal can't hang us).
# Only reached on a successful build (a thrown build error skips post-finally).
foreach ($p in $productsToBuild) {
  # Clear any Excel instance left over from the previous check's background
  # job before starting the next - two compile-check Excels racing each other
  # can make Workbooks.Open transiently block and trip the timeout falsely.
  Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
  Start-Sleep -Seconds 2
  $checkPath = Join-Path $OutDir $outNames[$p]
  Write-Host ("Compile-checking " + $outNames[$p] + " ...")
  & (Join-Path $PSScriptRoot 'compile-check.ps1') -XlsmPath $checkPath
  if ($LASTEXITCODE -ne 0) { throw ("Compile check FAILED for " + $outNames[$p] + " - the workbook has a VBA compile error.") }
}
