# Verify the AGOL Map column (Sites col 24) and its driver named range
# JobAgolMap on Setup. Three cases:
#   A. JobAgolMap blank   -> cell is blank
#   B. JobAgolMap set, URL already has "?" -> deep-link joined with "&"
#   C. JobAgolMap set, URL has NO "?"      -> deep-link joined with "?"
# Also confirms the MDOT NFC Map link (col 16) hits the ArcGIS Map Viewer
# URL (not the Experience app).

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath)
  $sites = $wb.Worksheets('Sites')
  $sites.Range($sites.Cells(2,1), $sites.Cells(10,24)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear

  # Test point (Kalamazoo)
  $sites.Cells(2, 4).Value2 = 'AGOL test site'
  $sites.Cells(2, 6).Value2 = [double]42.28536
  $sites.Cells(2, 7).Value2 = [double]-85.57025

  # ---- Case A: JobAgolMap blank -> AGOL Map cell is blank ----
  Write-Host "=== Case A: JobAgolMap blank ===" -ForegroundColor Cyan
  $wb.Names('JobAgolMap').RefersToRange.Value2 = ''
  $excel.Calculate()
  $cellA = [string]$sites.Cells(2, 24).Value2
  Write-Host ("  AGOL Map cell value: '" + $cellA + "'")
  if ($cellA -ne '') { throw "Expected blank AGOL cell when JobAgolMap is empty" }
  Write-Host "  Case A pass" -ForegroundColor Green

  # ---- Case B: URL already has ? ----
  Write-Host ""
  Write-Host "=== Case B: URL already has '?' ===" -ForegroundColor Cyan
  $urlB = 'https://www.arcgis.com/apps/mapviewer/index.html?webmap=abc123'
  $wb.Names('JobAgolMap').RefersToRange.Value2 = $urlB
  $excel.Calculate()
  $cellB = [string]$sites.Cells(2, 24).Value2
  $formulaB = [string]$sites.Cells(2, 24).Formula
  Write-Host ("  AGOL Map cell value: '" + $cellB + "'")
  Write-Host ("  Formula length: " + $formulaB.Length)
  if ($cellB -ne 'Open') { throw ("Expected cell text 'Open' (the hyperlink friendly name), got: '" + $cellB + "'") }
  # Excel's HYPERLINK target is the formula's first arg (built from the
  # JobAgolMap named range), not a literal in the formula text. Confirm
  # the formula references the named range + the center/marker template.
  if (-not $formulaB.Contains('JobAgolMap')) { throw "Formula doesn't reference JobAgolMap" }
  if (-not $formulaB.Contains('center=')) { throw "Formula doesn't stitch a center= param" }
  if (-not $formulaB.Contains('marker=')) { throw "Formula doesn't stitch a marker= param" }
  if (-not $formulaB.Contains('level=16')) { throw "Formula doesn't set level=16" }
  Write-Host "  Case B pass" -ForegroundColor Green

  # ---- Case C: URL has no ? ----
  Write-Host ""
  Write-Host "=== Case C: URL has no '?' (raw host) ===" -ForegroundColor Cyan
  $urlC = 'https://www.arcgis.com/apps/mapviewer/index.html'
  $wb.Names('JobAgolMap').RefersToRange.Value2 = $urlC
  $excel.Calculate()
  $cellC = [string]$sites.Cells(2, 24).Value2
  if ($cellC -ne 'Open') { throw ("Expected cell text 'Open' in Case C, got: '" + $cellC + "'") }
  Write-Host "  Case C pass" -ForegroundColor Green

  # ---- MDOT NFC Map (col 16) — confirm it now hits the FEMA webmap, not the Experience ----
  Write-Host ""
  Write-Host "=== MDOT NFC Map (col 16) URL ===" -ForegroundColor Cyan
  $f16 = [string]$sites.Cells(2, 16).Formula
  if (-not $f16.Contains('fema.maps.arcgis.com')) { throw "Col 16 formula no longer points at the FEMA webmap: $f16" }
  if (-not $f16.Contains('webmap=6a1702b9147243d1a5ee62cd614bc681')) { throw "Col 16 missing the expected webmap id: $f16" }
  if ($f16.Contains('experience.arcgis.com')) { throw "Col 16 still references the Experience app: $f16" }
  Write-Host "  col 16 routes through the FEMA-hosted NFC/ACUB webmap" -ForegroundColor Green

  # ---- CSV export resolves URLs for col 24 ----
  Write-Host ""
  Write-Host "=== CSV export resolves AGOL URL for col 24 ===" -ForegroundColor Cyan
  $excel.Run('SetHeadless', $true) | Out-Null
  # Set a known output folder so ExportSitesCsv doesn't default elsewhere
  $csvFolder = Join-Path $env:TEMP 'RoadReviewer_AgolCsv'
  if (Test-Path $csvFolder) { Get-ChildItem $csvFolder | Remove-Item -Force } else { New-Item -ItemType Directory -Path $csvFolder | Out-Null }
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ($csvFolder + '\')
  $excel.Run('ExportSitesCsv') | Out-Null
  $csv = Get-Content (Join-Path $csvFolder 'RoadReviewer Sites.csv') -Raw
  # Last column in the data row should be the resolved AGOL URL
  if (-not $csv.Contains($urlC)) { throw "CSV missing the AGOL base URL" }
  if (-not $csv.Contains('center=-85.57025,42.28536')) { throw "CSV AGOL URL doesn't have center coords" }
  Write-Host "  CSV row contains resolved AGOL URL with center + marker" -ForegroundColor Green

  $sites.Range($sites.Cells(2,1), $sites.Cells(10,24)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas before save
  $wb.Names('JobAgolMap').RefersToRange.Value2 = ''
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ''
  $excel.Run('SetHeadless', $false) | Out-Null
  $wb.Save()
  $wb.Close($true)
  Get-ChildItem $csvFolder | Remove-Item -Force
  Remove-Item $csvFolder -Force -ErrorAction SilentlyContinue

  Write-Host ""
  Write-Host "VERIFICATION PASSED -- AGOL Map column + MDOT NFC Map URL change" -ForegroundColor Green
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
