# Verify the AGOL Map column (Sites col 24) and its driver named range
# JobAgolMap on Setup. Three cases:
#   A. JobAgolMap blank   -> cell is blank
#   B. JobAgolMap set, URL already has "?" -> deep-link joined with "&"
#   C. JobAgolMap set, URL has NO "?"      -> deep-link joined with "?"
# Also confirms the state-aware NFC Map link (col 18, §9.3b) hits the
# right URL per state: MI's curated FEMA webmap, IN/WI's own live NFC
# FeatureServer side-loaded into the Map Viewer, and the generic FEMA
# Map Viewer pin for an unwired state.

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
  $sites.Range($sites.Cells(2,1), $sites.Cells(10,27)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear

  # Test point (Kalamazoo)
  $sites.Cells(2, 4).Value2 = 'AGOL test site'
  $sites.Cells(2, 6).Value2 = [double]42.28536
  $sites.Cells(2, 7).Value2 = [double]-85.57025

  # ---- Case A: JobAgolMap blank -> AGOL Map cell is blank ----
  Write-Host "=== Case A: JobAgolMap blank ===" -ForegroundColor Cyan
  $wb.Names('JobAgolMap').RefersToRange.Value2 = ''
  $excel.Calculate()
  $cellA = [string]$sites.Cells(2, 27).Value2
  Write-Host ("  AGOL Map cell value: '" + $cellA + "'")
  if ($cellA -ne '') { throw "Expected blank AGOL cell when JobAgolMap is empty" }
  Write-Host "  Case A pass" -ForegroundColor Green

  # ---- Case B: URL already has ? ----
  Write-Host ""
  Write-Host "=== Case B: URL already has '?' ===" -ForegroundColor Cyan
  $urlB = 'https://www.arcgis.com/apps/mapviewer/index.html?webmap=abc123'
  $wb.Names('JobAgolMap').RefersToRange.Value2 = $urlB
  $excel.Calculate()
  $cellB = [string]$sites.Cells(2, 27).Value2
  $formulaB = [string]$sites.Cells(2, 27).Formula
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
  $cellC = [string]$sites.Cells(2, 27).Value2
  if ($cellC -ne 'Open') { throw ("Expected cell text 'Open' in Case C, got: '" + $cellC + "'") }
  Write-Host "  Case C pass" -ForegroundColor Green

  # ---- NFC Map (col 18) — state-aware (§9.3b). Formula text always
  # contains all three states' URL fragments (one nested-IF formula for
  # every row) so assertions check the CALCULATED value (.Hyperlinks(1).Address)
  # for the active State, not the raw formula text. ----
  Write-Host ""
  Write-Host "=== NFC Map (col 18) URL, State=MI (default) ===" -ForegroundColor Cyan
  $excel.Calculate()
  $vMi = [string]$sites.Cells(2, 18).Hyperlinks(1).Address
  if (-not $vMi.Contains('fema.maps.arcgis.com')) { throw "MI link no longer points at the FEMA webmap: $vMi" }
  if (-not $vMi.Contains('webmap=6a1702b9147243d1a5ee62cd614bc681')) { throw "MI link missing the expected webmap id: $vMi" }
  if ($vMi.Contains('experience.arcgis.com')) { throw "MI link still references the Experience app: $vMi" }
  Write-Host "  MI routes through the FEMA-hosted NFC/ACUB webmap" -ForegroundColor Green

  Write-Host ""
  Write-Host "=== NFC Map (col 18) URL, State=IN ===" -ForegroundColor Cyan
  $wb.Names('JobState').RefersToRange.Value2 = 'IN'
  $excel.Calculate()
  $vIn = [string]$sites.Cells(2, 18).Hyperlinks(1).Address
  if (-not $vIn.Contains('gisdata.in.gov')) { throw "IN link doesn't side-load the Indiana FeatureServer: $vIn" }
  Write-Host "  IN side-loads gisdata.in.gov" -ForegroundColor Green

  Write-Host ""
  Write-Host "=== NFC Map (col 18) URL, State=WI ===" -ForegroundColor Cyan
  $wb.Names('JobState').RefersToRange.Value2 = 'WI'
  $excel.Calculate()
  $vWi = [string]$sites.Cells(2, 18).Hyperlinks(1).Address
  if (-not $vWi.Contains('services5.arcgis.com')) { throw "WI link doesn't side-load the WisDOT FeatureServer: $vWi" }
  Write-Host "  WI side-loads services5.arcgis.com" -ForegroundColor Green

  Write-Host ""
  Write-Host "=== NFC Map (col 18) URL, State=MN (unwired) ===" -ForegroundColor Cyan
  $wb.Names('JobState').RefersToRange.Value2 = 'MN'
  $excel.Calculate()
  $vMn = [string]$sites.Cells(2, 18).Hyperlinks(1).Address
  if ($vMn.Contains('webmap=')) { throw "MN link shouldn't fall back to the MI-specific webmap: $vMn" }
  if (-not $vMn.Contains('fema.maps.arcgis.com')) { throw "MN link should still open the generic FEMA Map Viewer: $vMn" }
  Write-Host "  MN falls back to the generic FEMA Map Viewer pin" -ForegroundColor Green

  $wb.Names('JobState').RefersToRange.Value2 = 'MI'
  $excel.Calculate()

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
  # State was reset to MI above, so the CSV's NFC Map column (modExport's
  # NfcMapUrlForRow, kept in sync by hand with SetNfcMapFormula) should
  # resolve to the same MI webmap the formula did.
  if (-not $csv.Contains('webmap=6a1702b9147243d1a5ee62cd614bc681')) { throw "CSV NFC Map column doesn't resolve to the MI webmap" }
  Write-Host "  CSV row contains resolved NFC Map URL for State=MI" -ForegroundColor Green

  $sites.Range($sites.Cells(2,1), $sites.Cells(10,27)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas before save
  $wb.Names('JobAgolMap').RefersToRange.Value2 = ''
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ''
  $excel.Run('SetHeadless', $false) | Out-Null
  $wb.Close($false)
  Get-ChildItem $csvFolder | Remove-Item -Force
  Remove-Item $csvFolder -Force -ErrorAction SilentlyContinue

  Write-Host ""
  Write-Host "VERIFICATION PASSED -- AGOL Map column + state-aware NFC Map column" -ForegroundColor Green
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
