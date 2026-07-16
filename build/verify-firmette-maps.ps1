# Verification for Workflow 3 ported from the prototypes:
#   DownloadFirmettes (FEMA GP submitJob -> poll -> OutputFile -> PDF)
#   PrepareMapPages + ExportCombinedMapPdf
#
# Uses a SINGLE Kalamazoo test row to keep the network round-trip short.
# A FEMA GP job typically takes 5-30s; allow up to 3 min before failing.
#
# Inspector product only - the standard RoadReviewer.xlsm has no FIRMette /
# MapPages buttons or WO/DI named ranges.

param([string]$XlsmPath = (Join-Path $env:TEMP 'Site Inspector Review Tool.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

$outFolder = Join-Path $env:TEMP 'RoadReviewer_Workflow3'
if (Test-Path -LiteralPath $outFolder) {
  Get-ChildItem -LiteralPath $outFolder -Filter '*.pdf' | Remove-Item -Force
  Get-ChildItem -LiteralPath $outFolder -Filter '*.kml' | Remove-Item -Force
} else {
  New-Item -ItemType Directory -Path $outFolder -Force | Out-Null
}

# NEVER open the committed workbook read-write: OneDrive AutoSave persists
# macro side effects (§7d) - this script's old sheet-delete "cleanup" once
# corrupted the committed file's job named ranges to #REF! that way. Work on
# a %TEMP% copy instead.
$tempXlsm = Join-Path $env:TEMP 'rr-verify-w3-copy.xlsm'
Copy-Item -LiteralPath $XlsmPath -Destination $tempXlsm -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($tempXlsm)
  $sites = $wb.Worksheets('Sites')

  # Setup values
  $wb.Names('JobWO').RefersToRange.Value2 = 'TEST'
  $wb.Names('JobDI').RefersToRange.Value2 = '0001'
  $wb.Names('JobDisaster').RefersToRange.Value2 = 'DR-TEST'
  $wb.Names('JobApplicant').RefersToRange.Value2 = 'Test Applicant'
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ($outFolder + '\')

  # Clear and write one Sites row (row 1 header, data from row 2; tests use row 3;
  # Lat=5, Lon=6, Description=7, Category=9)
  $sites.Range($sites.Cells(3,1), $sites.Cells(11,30)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $sites.Cells(3, 3).Value2 = 1                     # Site #
  $sites.Cells(3, 4).Value2 = 'Kalamazoo test'      # Site Name
  $sites.Cells(3, 5).Value2 = [double]42.28536
  $sites.Cells(3, 6).Value2 = [double]-85.57025
  $sites.Cells(3, 7).Value2 = 'pothole'              # Description
  $sites.Cells(3, 9).Value2 = 'C'                    # Category

  $excel.Run('SetHeadless', $true) | Out-Null
  $excel.Run('SetTrace', (Join-Path $env:TEMP 'RoadReviewer_w3_trace.txt')) | Out-Null

  # ---- Per-row WO/DI override in the FIRMette namer (F14 / §9.5, no network) ----
  # The row's own WO/DI (Sites cols A/B) must override the Setup WO/DI in the
  # file name, the same way the stamp already does. FirmettePreview() names the
  # first valid row (row 3 here), so it exercises the exact FirmetteFileName path.
  # Use NON-numeric Setup WO/DI here so Excel doesn't coerce them (e.g. '0001'
  # -> 1); restore the real ones afterwards for the download leg below.
  Write-Host "=== Per-row WO/DI override (FirmettePreview, no network) ===" -ForegroundColor Cyan
  $wb.Names('JobWO').RefersToRange.Value2 = 'AAA'
  $wb.Names('JobDI').RefersToRange.Value2 = 'BBB'
  $sites.Cells(3, 1).Value2 = '999'   # COL_WO row override
  $sites.Cells(3, 2).Value2 = '888'   # COL_DI row override
  $ovr = [string]$excel.Run('FirmettePreview')
  Write-Host ("  row override  : " + $ovr)
  if ($ovr -notlike '*WO999 DI888*') { throw ("Row WO/DI did not override Setup in the FIRMette name: '" + $ovr + "'") }
  $sites.Cells(3, 2).ClearContents()  # blank the row DI only -> Setup DI, row WO kept
  $mix = [string]$excel.Run('FirmettePreview')
  Write-Host ("  mixed (WO row, DI setup): " + $mix)
  if ($mix -notlike '*WO999 DIBBB*') { throw ("Per-field WO/DI fallback wrong: '" + $mix + "'") }
  $sites.Cells(3, 1).ClearContents()  # blank both -> full Setup fallback
  $fb = [string]$excel.Run('FirmettePreview')
  Write-Host ("  setup fallback: " + $fb)
  if ($fb -notlike '*WOAAA DIBBB*') { throw ("Blank row WO/DI did not fall back to Setup: '" + $fb + "'") }
  $wb.Names('JobWO').RefersToRange.Value2 = 'TEST'   # restore for the download leg
  $wb.Names('JobDI').RefersToRange.Value2 = '0001'
  Write-Host "  per-row override OK" -ForegroundColor Green

  # ---- DownloadFirmettes ----
  Write-Host "=== DownloadFirmettes (hits FEMA Print FIRMette GP service) ===" -ForegroundColor Cyan
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $excel.Run('DownloadFirmettes') | Out-Null
  $sw.Stop()
  Write-Host ("  finished in " + $sw.Elapsed.TotalSeconds + "s")
  $firmStatus = [string]$sites.Cells(3, 29).Value2   # FIRMette Status (col 29 in the new layout)
  Write-Host ("  FIRMette Status: " + $firmStatus)

  $pdfFiles = Get-ChildItem -LiteralPath $outFolder -Filter '*.pdf'
  Write-Host ("  PDFs in folder: " + ($pdfFiles.Count))
  $pdfFiles | ForEach-Object { Write-Host ("    " + $_.Name + " (" + $_.Length + " bytes)") }
  if ($firmStatus -notlike "Downloaded:*") { throw ("Status didn't say Downloaded: '" + $firmStatus + "'") }
  if ($pdfFiles.Count -lt 1) { throw "No PDF was written" }
  if ($pdfFiles[0].Length -lt 5000) { throw ("Output PDF suspiciously small: " + $pdfFiles[0].Length + " bytes") }
  Write-Host "  FIRMette download OK" -ForegroundColor Green

  # ---- PrepareMapPages + ExportCombinedMapPdf ----
  Write-Host ""
  Write-Host "=== PrepareMapPages ===" -ForegroundColor Cyan
  $excel.Run('PrepareMapPages') | Out-Null
  if (-not ($wb.Worksheets | ForEach-Object { $_.Name } | Where-Object { $_ -eq 'Map Pages' })) {
    throw "Map Pages sheet was not created"
  }
  $wsMap = $wb.Worksheets('Map Pages')
  $shapeCount = $wsMap.Shapes.Count
  $mergedCheck = $wsMap.Range('A1').MergeArea.Address
  Write-Host ("  MapPages created with " + $shapeCount + " shape(s); A1 merge range=" + $mergedCheck)
  if ($shapeCount -lt 1) { throw "Expected at least 1 textbox shape on MapPages" }

  Write-Host ""
  Write-Host "=== ExportCombinedMapPdf ===" -ForegroundColor Cyan
  $excel.Run('ExportCombinedMapPdf') | Out-Null
  # DisasterTag strips separators (PR #37): typed 'DR-TEST' -> 'DRTEST'.
  $pdfFiles = Get-ChildItem -LiteralPath $outFolder -Filter 'WO* DI* - DRTEST - Location Map.pdf'
  Write-Host ("  Map PDFs found: " + ($pdfFiles.Count))
  $pdfFiles | ForEach-Object { Write-Host ("    " + $_.Name + " (" + $_.Length + " bytes)") }
  if ($pdfFiles.Count -lt 1) { throw "Combined map PDF was not exported" }
  if ($pdfFiles[0].Length -lt 1000) { throw ("Map PDF suspiciously small: " + $pdfFiles[0].Length + " bytes") }
  Write-Host "  Map PDF export OK" -ForegroundColor Green

  # No workbook "cleanup" needed: everything ran on a throwaway %TEMP% copy.
  # (The old cleanup deleted the Map Pages sheet - which is PERMANENT since
  # §7d and carries the job named ranges - and AutoSave synced that
  # corruption into the committed file once. Never again.)
  $wb.Close($false)

  Write-Host ""
  Write-Host "VERIFICATION PASSED (Workflow 3 ported)" -ForegroundColor Green
}
catch {
  Write-Host ("VERIFICATION FAILED: " + $_.Exception.Message) -ForegroundColor Red
  if (Test-Path (Join-Path $env:TEMP 'RoadReviewer_w3_trace.txt')) {
    Write-Host "Last 25 trace lines:" -ForegroundColor Yellow
    Get-Content (Join-Path $env:TEMP 'RoadReviewer_w3_trace.txt') | Select-Object -Last 25 | ForEach-Object { Write-Host ("  " + $_) }
  }
  throw
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
