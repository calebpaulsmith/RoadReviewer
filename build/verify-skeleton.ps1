# Verification §5.2 + §5.3: skeleton present, every button wired, Sites
# hyperlinks/validation working, product-specific surface correct. Pure
# inspection - does NOT call out to any remote service.
#
# Product-aware: reads the hidden RR_Product defined name baked in by
# build.ps1 (Standard -> RoadReviewer.xlsm, Inspector -> Site Inspector
# Review Tool.xlsm) and asserts that product's expected inputs, buttons
# and hidden columns. Run it once per built workbook.

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

$expectedSheets = @('Start Here', 'Sites', 'Sources')

# New canonical Sites layout: row 1 toolbar, row 2 header, data from row 3.
$HeaderRow = 2
$FirstDataRow = 3
$expectedHeaders = @('WO #','DI #','Site #','Site Name','Latitude','Longitude','Description (optional)','Address (optional)','Category (optional)','Costs (optional)','Work Completion (optional)','Geocode Status','NFC Map','FHWA Class','Urban/Rural','ACUB Name','Road Name','Street Name','Federal Aid Status','Review Reason','Google Maps','Street View','Bing','Google Earth','FEMA Viewer','FIRMette Portal','FIRMette Status','Map Status','AGOL Map','AGOL NFC Layer')

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath, $false, $true)   # ReadOnly=true so we don't dirty the file
  $proj = $wb.VBProject

  # ---- Which product is this file? ----
  $product = 'Inspector'   # missing name defaults to Inspector, same as the VBA
  try {
    $refersTo = [string]$wb.Names('RR_Product').RefersTo   # looks like ="Standard"
    if ($refersTo -match 'Standard') { $product = 'Standard' }
  } catch { }
  $isInspector = ($product -eq 'Inspector')
  Write-Host ("=== Product: {0} ===" -f $product) -ForegroundColor Cyan

  Write-Host "=== Sheets ===" -ForegroundColor Cyan
  $present = @{}
  foreach ($ws in $wb.Worksheets) { $present[$ws.Name] = $true; Write-Host ("  " + $ws.Name) }
  $missing = $expectedSheets | Where-Object { -not $present.ContainsKey($_) }
  if ($missing) { throw ("Missing sheets: " + ($missing -join ', ')) }
  Write-Host "  all expected sheets present" -ForegroundColor Green

  # Collect every Public Sub name across modules so we can verify button OnActions.
  $publicSubs = @{}
  foreach ($comp in $proj.VBComponents) {
    if ($comp.Type -ne 1) { continue }   # 1 = standard module
    $cm = $comp.CodeModule
    for ($i = 1; $i -le $cm.CountOfLines; $i++) {
      $line = $cm.Lines($i, 1)
      if ($line -match '^\s*Public\s+Sub\s+(\w+)') {
        $publicSubs[$Matches[1]] = $comp.Name
      } elseif ($line -match '^\s*Sub\s+(\w+)') {
        # Bare `Sub Foo()` is also publicly callable from Excel UI
        $publicSubs[$Matches[1]] = $comp.Name
      }
    }
  }
  Write-Host ("  found " + $publicSubs.Count + " callable subs across modules")

  Write-Host "=== Buttons ===" -ForegroundColor Cyan
  $btnCount = 0; $orphans = @(); $onActions = @{}
  foreach ($ws in $wb.Worksheets) {
    foreach ($sh in $ws.Shapes) {
      $oa = ""
      try { $oa = [string]$sh.OnAction } catch {}
      if ($oa) {
        $btnCount++
        $onActions[$oa] = $true
        $resolved = $publicSubs[$oa]
        $caption = $sh.TextFrame2.TextRange.Text
        $status = if ($resolved) { "OK   ($resolved)" } else { "ORPHAN"; $orphans += "$($ws.Name)::$caption -> $oa" }
        Write-Host ("  [{0,-12}] {1,-44} -> {2,-28} {3}" -f $ws.Name, $caption, $oa, $status)
      }
    }
  }
  if ($orphans) {
    Write-Host "ORPHAN BUTTONS:" -ForegroundColor Red
    $orphans | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    throw "Some buttons reference subs that don't exist"
  } else {
    Write-Host ("  $btnCount buttons, every OnAction resolves") -ForegroundColor Green
  }

  # Product button surface: shared actions everywhere; workflow-3 actions
  # inspector-only.
  $sharedActions = @('CheckRoads','ReRunFailedRows','OpenImageryForSelection','ExportSitesCsv','ExportSitesToKML','SendSitesToAgolMap','OpenSitesOnNfcLayer','BuildWorkbook')
  $inspectorActions = @('DownloadFirmettes','ReRunFailedFirmettes','PrepareMapPages','AddMapPage','ExportCombinedMapPdf')
  foreach ($a in $sharedActions) {
    if (-not $onActions.ContainsKey($a)) { throw "Missing expected button for: $a" }
  }
  foreach ($a in $inspectorActions) {
    if ($isInspector -and -not $onActions.ContainsKey($a)) { throw "Inspector build missing button for: $a" }
    if (-not $isInspector -and $onActions.ContainsKey($a)) { throw "Standard build must NOT have a button for: $a" }
  }
  Write-Host "  product button surface correct" -ForegroundColor Green

  Write-Host "=== Named ranges ===" -ForegroundColor Cyan
  $sharedNames = @('JobState','JobOutputFolder','JobAgolMap','JobBufferFeet')
  $inspectorNames = @('JobWO','JobDI','JobDisaster','JobApplicant')
  foreach ($n in $sharedNames) {
    try { $r = $wb.Names($n); Write-Host ("  " + $n + " -> " + $r.RefersTo) } catch { throw "Missing named range: $n" }
  }
  foreach ($n in $inspectorNames) {
    $found = $true
    try { $r = $wb.Names($n) } catch { $found = $false }
    if ($isInspector) {
      if (-not $found) { throw "Inspector build missing named range: $n" }
      Write-Host ("  " + $n + " -> " + $r.RefersTo)
    } else {
      if ($found) { throw "Standard build must NOT have named range: $n" }
    }
  }
  Write-Host "  product named-range surface correct" -ForegroundColor Green

  Write-Host "=== Sites headers (row $HeaderRow) ===" -ForegroundColor Cyan
  $sites = $wb.Worksheets('Sites')
  for ($c = 1; $c -le $expectedHeaders.Count; $c++) {
    $got = [string]$sites.Cells($HeaderRow, $c).Value2
    $want = $expectedHeaders[$c-1]
    if ($got -ne $want) { throw ("Header mismatch at col " + $c + ": got '" + $got + "', want '" + $want + "'") }
  }
  Write-Host ("  all " + $expectedHeaders.Count + " headers match constants") -ForegroundColor Green

  Write-Host "=== Product column hiding ===" -ForegroundColor Cyan
  # Inspector-only columns: WO(1), DI(2), FIRMette Status(27), Map Status(28)
  foreach ($c in @(1, 2, 27, 28)) {
    $hidden = [bool]$sites.Columns($c).Hidden
    if ($isInspector -and $hidden) { throw "Inspector build should show column $c" }
    if (-not $isInspector -and -not $hidden) { throw "Standard build should hide column $c" }
  }
  Write-Host ("  inspector-only columns " + $(if ($isInspector) { 'visible' } else { 'hidden' }) + " as expected") -ForegroundColor Green

  Write-Host "=== Sites toolbar (row 1) ===" -ForegroundColor Cyan
  $toolbarNames = @()
  foreach ($sh in $sites.Shapes) { $toolbarNames += $sh.Name }
  foreach ($n in @('RR_CheckRoads', 'RR_PhotoLinks')) {
    if ($toolbarNames -notcontains $n) { throw "Sites toolbar missing shape: $n" }
  }
  Write-Host "  Check Roads + Photo Links buttons present on the Sites sheet" -ForegroundColor Green

  Write-Host "=== Sites hyperlink resolution (test coord) ===" -ForegroundColor Cyan
  # Use the Kalamazoo test coord. We have to open in r/w to write cells; reopen.
  $wb.Close($false)
  $wb = $excel.Workbooks.Open($XlsmPath, $false, $false)
  $sites = $wb.Worksheets('Sites')
  $sites.Cells($FirstDataRow, 4).Value2 = 'Test - Kalamazoo'   # Site Name
  $sites.Cells($FirstDataRow, 5).Value2 = [double]42.28536      # Lat
  $sites.Cells($FirstDataRow, 6).Value2 = [double]-85.57025     # Lon
  $excel.Calculate()
  # Verify each hyperlink formula resolves to a non-empty string
  $linkCols = @{ 13='NFC Map'; 21='Google Maps'; 22='Street View'; 23='Bing'; 24='Google Earth'; 25='FEMA Viewer'; 26='FIRMette Portal'; 30='AGOL NFC Layer' }
  foreach ($k in $linkCols.Keys | Sort-Object) {
    $cell = $sites.Cells($FirstDataRow, $k)
    $f = [string]$cell.Formula
    $v = [string]$cell.Value2
    if (-not $f.StartsWith('=')) { throw "Col $k formula empty: $f" }
    if (-not $v) { throw ("Col $k ({0}) shows empty with a valid coord" -f $linkCols[$k]) }
    Write-Host ("  col {0,2} ({1,-16}) shows: '{2}'  -- formula intact: {3} chars" -f $k, $linkCols[$k], $v, $f.Length)
  }

  Write-Host "=== Sites validation ===" -ForegroundColor Cyan
  $latVal = $sites.Cells($FirstDataRow, 5).Validation
  $lonVal = $sites.Cells($FirstDataRow, 6).Validation
  # xlValidateDecimal = 2
  Write-Host ("  Latitude  validation type=" + $latVal.Type + " (2=Decimal)  formula1=" + $latVal.Formula1 + "  formula2=" + $latVal.Formula2)
  Write-Host ("  Longitude validation type=" + $lonVal.Type + " (2=Decimal)  formula1=" + $lonVal.Formula1 + "  formula2=" + $lonVal.Formula2)
  if ($latVal.Type -ne 2 -or $lonVal.Type -ne 2) { throw "Coordinate columns missing decimal validation" }
  if ($latVal.Formula1 -ne '-90' -or $latVal.Formula2 -ne '90') { throw "Latitude validation range wrong" }
  if ($lonVal.Formula1 -ne '-180' -or $lonVal.Formula2 -ne '180') { throw "Longitude validation range wrong" }

  Write-Host "=== Conditional formatting on Federal Aid Status columns ===" -ForegroundColor Cyan
  $r = $sites.Range($sites.Cells($FirstDataRow, 14), $sites.Cells($FirstDataRow, 20))
  $fcCount = $r.FormatConditions.Count
  Write-Host ("  format conditions count on class..eligibility row " + $FirstDataRow + ": " + $fcCount)
  if ($fcCount -lt 3) { throw ("Expected 3 conditional-format rules (federal aid / non-federal aid / review), got " + $fcCount) }

  Write-Host "=== Start Here disclaimer + version ===" -ForegroundColor Cyan
  $start = $wb.Worksheets('Start Here')
  $startBlob = ''
  foreach ($cell in $start.UsedRange.Cells) { $startBlob += ([string]$cell.Value2) + "`n" }
  if ($startBlob -notmatch 'NOT AN AUTHORITATIVE') { throw "Start Here is missing the disclaimer block" }
  if ($startBlob -notmatch 'eligibility determination') { throw "Start Here disclaimer missing the eligibility clause" }
  if ($startBlob -notmatch 'PR #') { throw "Start Here missing the PR/version label" }
  Write-Host "  disclaimer block + version label present" -ForegroundColor Green

  Write-Host "=== Sources sheet content ===" -ForegroundColor Cyan
  $sources = $wb.Worksheets('Sources')
  $sourcesBlob = ''
  foreach ($cell in $sources.UsedRange.Cells) { $sourcesBlob += ([string]$cell.Value2) + "`n" }
  $sourcesCellCount = $excel.WorksheetFunction.CountA($sources.UsedRange)
  Write-Host ("  Sources sheet non-empty cells: " + $sourcesCellCount)
  if ($sourcesCellCount -lt 20) { throw "Sources sheet looks empty" }
  if ($sourcesBlob -notmatch 'BOUNDARY ROADS') { throw "Sources sheet missing the boundary-roads caveat" }
  if ($sourcesBlob -notmatch 'does NOT authoritatively') { throw "Sources sheet missing the not-authoritative disclaimer" }
  Write-Host "  boundary caveat + disclaimer present on Sources" -ForegroundColor Green

  # Clean up the test row so the saved file stays empty
  $sites.Range($sites.Cells($FirstDataRow, 4), $sites.Cells($FirstDataRow, 6)).ClearContents()
  $wb.Save()

  Write-Host "VERIFICATION PASSED ($product)" -ForegroundColor Green
  $wb.Close($true)
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
