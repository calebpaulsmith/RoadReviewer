# Verification for Workflow 3 ported from the prototypes:
#   DownloadFirmettes (FEMA GP submitJob -> poll -> OutputFile -> PDF)
#   PrepareMapPages + ExportCombinedMapPdf
#
# Uses a SINGLE Kalamazoo test row to keep the network round-trip short.
# A FEMA GP job typically takes 5-30s; allow up to 3 min before failing.

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

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

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath)
  $sites = $wb.Worksheets('Sites')

  # Setup values
  $wb.Names('JobWO').RefersToRange.Value2 = 'TEST'
  $wb.Names('JobDI').RefersToRange.Value2 = '0001'
  $wb.Names('JobDisaster').RefersToRange.Value2 = 'DR-TEST'
  $wb.Names('JobApplicant').RefersToRange.Value2 = 'Test Applicant'
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ($outFolder + '\')

  # Clear and write one Sites row
  $sites.Range($sites.Cells(2,1), $sites.Cells(10,23)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $sites.Cells(2, 3).Value2 = 1                     # Site #
  $sites.Cells(2, 4).Value2 = 'Kalamazoo test'      # Site Name
  $sites.Cells(2, 6).Value2 = [double]42.28536
  $sites.Cells(2, 7).Value2 = [double]-85.57025
  $sites.Cells(2, 8).Value2 = 'C'                    # Category
  $sites.Cells(2, 9).Value2 = 'pothole'              # Description

  $excel.Run('SetHeadless', $true) | Out-Null
  $excel.Run('SetTrace', (Join-Path $env:TEMP 'RoadReviewer_w3_trace.txt')) | Out-Null

  # ---- DownloadFirmettes ----
  Write-Host "=== DownloadFirmettes (hits FEMA Print FIRMette GP service) ===" -ForegroundColor Cyan
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $excel.Run('DownloadFirmettes') | Out-Null
  $sw.Stop()
  Write-Host ("  finished in " + $sw.Elapsed.TotalSeconds + "s")
  $firmStatus = [string]$sites.Cells(2, 22).Value2
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
  if (-not ($wb.Worksheets | ForEach-Object { $_.Name } | Where-Object { $_ -eq 'MapPages' })) {
    throw "MapPages sheet was not created"
  }
  $wsMap = $wb.Worksheets('MapPages')
  $shapeCount = $wsMap.Shapes.Count
  $mergedCheck = $wsMap.Range('A1').MergeArea.Address
  Write-Host ("  MapPages created with " + $shapeCount + " shape(s); A1 merge range=" + $mergedCheck)
  if ($shapeCount -lt 1) { throw "Expected at least 1 textbox shape on MapPages" }

  Write-Host ""
  Write-Host "=== ExportCombinedMapPdf ===" -ForegroundColor Cyan
  $excel.Run('ExportCombinedMapPdf') | Out-Null
  $pdfFiles = Get-ChildItem -LiteralPath $outFolder -Filter 'WO* DI* - DR-TEST - Location Map.pdf'
  Write-Host ("  Map PDFs found: " + ($pdfFiles.Count))
  $pdfFiles | ForEach-Object { Write-Host ("    " + $_.Name + " (" + $_.Length + " bytes)") }
  if ($pdfFiles.Count -lt 1) { throw "Combined map PDF was not exported" }
  if ($pdfFiles[0].Length -lt 1000) { throw ("Map PDF suspiciously small: " + $pdfFiles[0].Length + " bytes") }
  Write-Host "  Map PDF export OK" -ForegroundColor Green

  # ---- Cleanup workbook state ----
  $excel.Run('SetHeadless', $false) | Out-Null
  $excel.Run('SetTrace', '') | Out-Null
  $sites.Range($sites.Cells(2,1), $sites.Cells(10,23)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  # Reset Setup so the user's saved workbook doesn't carry test values
  $wb.Names('JobWO').RefersToRange.Value2 = ''
  $wb.Names('JobDI').RefersToRange.Value2 = ''
  $wb.Names('JobDisaster').RefersToRange.Value2 = ''
  $wb.Names('JobApplicant').RefersToRange.Value2 = ''
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ''
  $excel.DisplayAlerts = $false
  $wb.Worksheets('MapPages').Delete()
  $wb.Save()
  $wb.Close($true)

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
