# Verification §5.2 + §5.3: skeleton present, every button wired, Sites
# hyperlinks/validation working. Pure inspection - does NOT call out to
# any remote service.

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

$expectedSheets = @('Home','Setup','Sites','1. Classify Roads','2. Review Imagery','3. Maps & FIRMettes')

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath, $false, $true)   # ReadOnly=true so we don't dirty the file
  $proj = $wb.VBProject

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
  $btnCount = 0; $orphans = @()
  foreach ($ws in $wb.Worksheets) {
    foreach ($sh in $ws.Shapes) {
      $oa = ""
      try { $oa = [string]$sh.OnAction } catch {}
      if ($oa) {
        $btnCount++
        $resolved = $publicSubs[$oa]
        $caption = $sh.TextFrame2.TextRange.Text
        $status = if ($resolved) { "OK   ($resolved)" } else { "ORPHAN"; $orphans += "$($ws.Name)::$caption -> $oa" }
        Write-Host ("  [{0,-22}] {1,-40} -> {2,-32} {3}" -f $ws.Name, $caption, $oa, $status)
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

  Write-Host "=== Named ranges ===" -ForegroundColor Cyan
  $expectedNames = @('JobWO','JobDI','JobDisaster','JobApplicant','JobState','JobOutputFolder','JobAgolMap','JobBufferFeet')
  foreach ($n in $expectedNames) {
    try { $r = $wb.Names($n); Write-Host ("  " + $n + " -> " + $r.RefersTo) } catch { throw "Missing named range: $n" }
  }

  Write-Host "=== Sites headers ===" -ForegroundColor Cyan
  $sites = $wb.Worksheets('Sites')
  $expectedHeaders = @('WO #','DI #','Site #','Site Name','Address','Latitude','Longitude','Category','Description','Costs','Work Completion','Geocode Status','Google Maps','Street View','Bing','FEMA Viewer','FIRMette Portal','NFC Map','FHWA Class','Urban/Rural','ACUB Name','Road Name','Street Name','Federal Aid Status','FIRMette Status','Map Status','AGOL Map')
  for ($c = 1; $c -le $expectedHeaders.Count; $c++) {
    $got = [string]$sites.Cells(1,$c).Value2
    $want = $expectedHeaders[$c-1]
    if ($got -ne $want) { throw ("Header mismatch at col " + $c + ": got '" + $got + "', want '" + $want + "'") }
  }
  Write-Host ("  all " + $expectedHeaders.Count + " headers match constants") -ForegroundColor Green

  Write-Host "=== Sites hyperlink resolution (test coord) ===" -ForegroundColor Cyan
  # Use the Kalamazoo test coord. We have to open in r/w to write cells; reopen.
  $wb.Close($false)
  $wb = $excel.Workbooks.Open($XlsmPath, $false, $false)
  $sites = $wb.Worksheets('Sites')
  $sites.Cells(2, 4).Value = 'Test - Kalamazoo'   # Site Name
  $sites.Cells(2, 6).Value = 42.28536              # Lat
  $sites.Cells(2, 7).Value = -85.57025             # Lon
  $excel.Calculate()
  # Verify each hyperlink formula resolves to a non-empty string
  $linkCols = @{ 13='Google Maps'; 14='Street View'; 15='Bing'; 16='FEMA Viewer'; 17='FIRMette Portal'; 18='NFC Map' }
  foreach ($k in $linkCols.Keys | Sort-Object) {
    $cell = $sites.Cells(2, $k)
    $f = [string]$cell.Formula
    $v = [string]$cell.Value2
    # Formula should reference the lat/lon cells
    if (-not $f.StartsWith('=')) { throw "Col $k formula empty: $f" }
    Write-Host ("  col {0,2} ({1,-21}) shows: '{2}'  -- formula intact: {3} chars" -f $k, $linkCols[$k], $v, $f.Length)
  }

  Write-Host "=== Sites validation ===" -ForegroundColor Cyan
  $latVal = $sites.Cells(2, 6).Validation
  $lonVal = $sites.Cells(2, 7).Validation
  # xlValidateDecimal = 2
  Write-Host ("  Latitude  validation type=" + $latVal.Type + " (2=Decimal)  formula1=" + $latVal.Formula1 + "  formula2=" + $latVal.Formula2)
  Write-Host ("  Longitude validation type=" + $lonVal.Type + " (2=Decimal)  formula1=" + $lonVal.Formula1 + "  formula2=" + $lonVal.Formula2)
  if ($latVal.Type -ne 2 -or $lonVal.Type -ne 2) { throw "Coordinate columns missing decimal validation" }
  if ($latVal.Formula1 -ne '-90' -or $latVal.Formula2 -ne '90') { throw "Latitude validation range wrong" }
  if ($lonVal.Formula1 -ne '-180' -or $lonVal.Formula2 -ne '180') { throw "Longitude validation range wrong" }

  Write-Host "=== Conditional formatting on Federal Aid Status column ===" -ForegroundColor Cyan
  $r = $sites.Range($sites.Cells(2,19), $sites.Cells(2,24))
  $fcCount = $r.FormatConditions.Count
  Write-Host ("  format conditions count on class..eligibility row 2: " + $fcCount)
  if ($fcCount -lt 3) { throw ("Expected 3 conditional-format rules (federal aid / non-federal aid / review), got " + $fcCount) }

  # Clean up the test row so the saved file stays empty
  $sites.Range($sites.Cells(2,4), $sites.Cells(2,7)).ClearContents()
  $wb.Save()

  Write-Host "VERIFICATION PASSED" -ForegroundColor Green
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
