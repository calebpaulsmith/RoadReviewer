# Verify that DownloadFirmettes, ExportCombinedMapPdf, the MapPages
# textbox, and the default output folder all behave cleanly when the
# inspector leaves WO and/or DI blank. No dangling "WO ", trailing
# " DI", "\WO-DI\" folder, or empty "WO # DI #" line in the stamp.

# Inspector product only - the standard RoadReviewer.xlsm has no FIRMette /
# MapPages buttons or WO/DI named ranges.
param([string]$XlsmPath = (Join-Path $env:TEMP 'Site Inspector Review Tool.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

$outFolder = Join-Path $env:TEMP 'RoadReviewer_BlankWoDi'
if (Test-Path -LiteralPath $outFolder) {
  Get-ChildItem -LiteralPath $outFolder | Remove-Item -Force
} else {
  New-Item -ItemType Directory -Path $outFolder -Force | Out-Null
}

function Split-StampLines([string]$s) {
  return ($s -replace "`r","" -split "`n")
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath)
  $sites = $wb.Worksheets('Sites')

  $wb.Names('JobWO').RefersToRange.Value2 = ''
  $wb.Names('JobDI').RefersToRange.Value2 = ''
  $wb.Names('JobDisaster').RefersToRange.Value2 = 'DR-TEST'
  $wb.Names('JobApplicant').RefersToRange.Value2 = 'Test Applicant'
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ($outFolder + '\')

  # Row 1 header, data from row 2; tests use row 3; Lat=5, Lon=6, Category=9.
  $sites.Range($sites.Cells(3,1), $sites.Cells(11,30)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $sites.Cells(3, 3).Value2 = 1
  $sites.Cells(3, 4).Value2 = 'BlankIDs site'
  $sites.Cells(3, 5).Value2 = [double]42.28536
  $sites.Cells(3, 6).Value2 = [double]-85.57025
  $sites.Cells(3, 9).Value2 = 'C'

  $excel.Run('SetHeadless', $true) | Out-Null

  Write-Host "=== Case A: WO and DI BOTH blank ===" -ForegroundColor Cyan
  $excel.Run('DownloadFirmettes') | Out-Null
  $statusA = [string]$sites.Cells(3, 29).Value2   # FIRMette Status (col 29 in the new layout)
  Write-Host ("  FIRMette Status: " + $statusA)
  $pdfs = @(Get-ChildItem -LiteralPath $outFolder -Filter '*.pdf')
  Write-Host ("  PDFs: " + $pdfs.Count)
  $pdfs | ForEach-Object { Write-Host ("    " + $_.Name) }
  if ($pdfs.Count -ne 1) { throw "Expected exactly 1 PDF" }
  $nameA = $pdfs[0].Name
  if ($nameA.Contains('WO ')) { throw ("Filename has dangling 'WO ': " + $nameA) }
  if ($nameA.Contains(' DI')) { throw ("Filename has dangling ' DI': " + $nameA) }
  if ($nameA.StartsWith('WO') -or $nameA.StartsWith('DI')) { throw ("Filename starts with WO/DI when both blank: " + $nameA) }
  if (-not $nameA.StartsWith('DR-TEST')) { throw ("Filename should start with DR-TEST when no WO/DI: " + $nameA) }
  Write-Host ("  filename clean: " + $nameA) -ForegroundColor Green

  $excel.Run('PrepareMapPages') | Out-Null
  $wsMap = $wb.Worksheets('Map Pages')
  $textbox = $wsMap.Shapes('Textbox_Page_1')
  $stampA = [string]$textbox.TextFrame2.TextRange.Text
  Write-Host "  Map textbox stamp:"
  $lines = Split-StampLines $stampA
  foreach ($ln in $lines) { Write-Host ("    | " + $ln) }
  if ($stampA.Contains('WO #') -or $stampA.Contains('DI #')) {
    throw "Textbox stamp should not contain 'WO #' or 'DI #' when both blank"
  }
  if (-not $stampA.Contains('Test Applicant')) { throw "Textbox missing applicant" }
  Write-Host "  textbox clean (no empty WO #/DI # line)" -ForegroundColor Green

  $excel.Run('ExportCombinedMapPdf') | Out-Null
  $mapPdfs = @(Get-ChildItem -LiteralPath $outFolder -Filter '*Location Map.pdf')
  if ($mapPdfs.Count -ne 1) { throw "Expected exactly 1 Map PDF" }
  $mapNameA = $mapPdfs[0].Name
  Write-Host ("  Map PDF: " + $mapNameA)
  if (-not $mapNameA.StartsWith('DR-TEST')) { throw ("Map PDF should start with DR-TEST: " + $mapNameA) }
  Write-Host "  Map PDF clean" -ForegroundColor Green

  # Reset Case B
  Write-Host ""
  Write-Host "=== Case B: WO=123, DI blank ===" -ForegroundColor Cyan
  Get-ChildItem -LiteralPath $outFolder | Remove-Item -Force
  $wb.Worksheets('Map Pages').Delete()
  $wb.Names('JobWO').RefersToRange.Value2 = '123'
  $sites.Cells(3, 29).Value2 = ''   # FIRMette Status
  $sites.Cells(3, 30).Value2 = ''   # Map Status

  $excel.Run('DownloadFirmettes') | Out-Null
  $pdfs = @(Get-ChildItem -LiteralPath $outFolder -Filter '*.pdf')
  if ($pdfs.Count -ne 1) { throw "Expected exactly 1 PDF in Case B" }
  $nameB = $pdfs[0].Name
  Write-Host ("  filename: " + $nameB)
  if (-not $nameB.StartsWith('WO123 -')) { throw ("Expected 'WO123 - ...', got: " + $nameB) }
  if ($nameB.Contains('DI')) { throw ("Expected no DI in filename, got: " + $nameB) }
  Write-Host "  Case B filename clean (WO only)" -ForegroundColor Green

  $excel.Run('PrepareMapPages') | Out-Null
  $stampB = [string]$wb.Worksheets('Map Pages').Shapes('Textbox_Page_1').TextFrame2.TextRange.Text
  $linesB = Split-StampLines $stampB
  Write-Host ("  textbox first line: " + $linesB[0])
  if ($linesB[0] -ne 'WO #123') { throw ("Expected stamp first line 'WO #123', got: '" + $linesB[0] + "'") }
  if ($stampB.Contains('DI #')) { throw ("Expected no DI # in stamp, got: " + $stampB) }
  Write-Host "  Case B textbox clean (WO # only)" -ForegroundColor Green

  # Cleanup
  $wb.Worksheets('Map Pages').Delete()
  $sites.Range($sites.Cells(3,1), $sites.Cells(11,30)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $wb.Names('JobWO').RefersToRange.Value2 = ''
  $wb.Names('JobDI').RefersToRange.Value2 = ''
  $wb.Names('JobDisaster').RefersToRange.Value2 = ''
  $wb.Names('JobApplicant').RefersToRange.Value2 = ''
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ''
  $excel.Run('SetHeadless', $false) | Out-Null
  $wb.Close($false)

  Write-Host ""
  Write-Host "VERIFICATION PASSED -- blank WO/DI handled cleanly" -ForegroundColor Green
}
catch {
  Write-Host ("VERIFICATION FAILED: " + $_.Exception.Message) -ForegroundColor Red
  throw
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
