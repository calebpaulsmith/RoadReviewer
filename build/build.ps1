# RoadReviewer build script.
# Drives Excel via COM: creates a new workbook, imports every .bas from ../src,
# runs BuildWorkbook, and saves as RoadReviewer.xlsm in this folder.
# Requires HKCU AccessVBOM=1 (verified before running).

param(
  [string]$SrcDir = (Join-Path $PSScriptRoot '..\src'),
  # Default to %TEMP% because the repo's build/ folder has a sandbox-imposed
  # `Everyone Deny DeleteSubdirectoriesAndFiles` ACE that breaks Excel's
  # SaveAs (it writes to a temp file then deletes/renames). %TEMP% is
  # always writable. The path can be overridden via -OutPath.
  [string]$OutPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'),
  [switch]$Visible
)

$ErrorActionPreference = 'Stop'
$XlOpenXMLWorkbookMacroEnabled = 52
$VbeModuleStandard = 1

function Resolve-Strict([string]$p) {
  $r = Resolve-Path -LiteralPath $p -ErrorAction SilentlyContinue
  if (-not $r) { throw "Path not found: $p" }
  return $r.ProviderPath
}

$SrcDir = Resolve-Strict $SrcDir
$OutPath = [System.IO.Path]::GetFullPath($OutPath)

$importOrder = @(
  'modConstants.bas',
  'modUtil.bas',
  'modHttp.bas',
  'modBuild.bas',
  'modClassify.bas',
  'modGeocode.bas',
  'modImagery.bas',
  'modMaps.bas',
  'modExport.bas'
)

# Verify all files exist before launching Excel
foreach ($f in $importOrder) {
  $p = Join-Path $SrcDir $f
  if (-not (Test-Path -LiteralPath $p)) { throw "Missing source: $p" }
}

if (Test-Path -LiteralPath $OutPath) {
  Write-Host "Removing existing $OutPath"
  Remove-Item -LiteralPath $OutPath -Force
}

Write-Host "Starting Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = [bool]$Visible
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable - we trust our own code; this just suppresses prompts on open

try {
  $wb = $excel.Workbooks.Add()
  $proj = $wb.VBProject
  Write-Host "Importing modules..."
  foreach ($f in $importOrder) {
    $p = Join-Path $SrcDir $f
    $comp = $proj.VBComponents.Import($p)
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

  Write-Host "Running BuildWorkbook..."
  try {
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

  Write-Host "Saving as $OutPath"
  $wb.SaveAs($OutPath, $XlOpenXMLWorkbookMacroEnabled)
  $wb.Close($false)
  Write-Host "Build complete." -ForegroundColor Green
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
