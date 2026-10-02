# HPMS double-check (2026-09-30) + the non-inventory twin drop in Excel's HPMS
# lookups. The state DOT layer always decides the verdict; with the Y/N box
# (JobHpmsCheck) set to Y every row is also looked up in FHWA's HPMS layer and
# a row where the two disagree on federal aid becomes "Review - Sources
# disagree".
#
# Legs (all live - WisDOT, HPMS, NTAD, Census):
#   1. STH 52 (45.169879,-89.102452), box N: Rural Minor Collector, no note.
#   2. Same row, box Y: HPMS carries the road twice (class 6 inventory record,
#      class 3 non-inventory record). The twin is dropped, HPMS agrees, the
#      verdict and the empty Review Reason are unchanged.
#   3. State layer down (dead URL in Svc_WI_LOCAL_ROADS) -> HPMS fallback must
#      ALSO read Minor Collector (the twin drop), not the class-3 twin and not
#      a "Conflicting classes" tie.
#   4. Box Y with HPMS down: the verdict stands, Review Reason says the check
#      was unavailable, the row does not fail.
#   5. A real state/HPMS mismatch (-MismatchLat/-MismatchLon/-MismatchState,
#      found with build\audit-hpms-vs-state.mjs): N gives a confident verdict,
#      Y turns it into "Review - Sources disagree" naming the HPMS class.
#      Skipped when no point is passed.
# Also prints the Check Roads time for N vs Y on the same rows (the cost of Y).
# Works against either product. Runs on a %TEMP% copy.
param(
  [string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'),
  [double]$MismatchLat = 0, [double]$MismatchLon = 0, [string]$MismatchState = ''
)

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }
$tempXlsm = Join-Path $env:TEMP 'rr-verify-hpmscheck-copy.xlsm'
Copy-Item -LiteralPath $XlsmPath -Destination $tempXlsm -Force

$deadUrl = 'https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($tempXlsm)
  $sites = $wb.Worksheets('Sites')
  $sites.Range($sites.Cells(3, 1), $sites.Cells(20, 30)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null
  $excel.Run('SetHeadless', $true) | Out-Null
  $excel.Run('SetTrace', (Join-Path $env:TEMP 'RoadReviewer_hpmscheck_trace.txt')) | Out-Null
  $check = $wb.Names('JobHpmsCheck').RefersToRange
  function Row3 { @{ elig = [string]$sites.Cells(3, 16).Value2; reason = [string]$sites.Cells(3, 17).Value2; cls = [string]$sites.Cells(3, 18).Value2 } }
  function SetPoint([double]$lat, [double]$lon, [string]$name) {
    $sites.Cells(3, 4).Value2 = $name; $sites.Cells(3, 5).Value2 = $lat; $sites.Cells(3, 6).Value2 = $lon
  }

  $wb.Names('JobState').RefersToRange.Value2 = 'WI'
  SetPoint 45.169879 -89.102452 'STH 52'

  Write-Host "=== 1. STH 52, double-check N ===" -ForegroundColor Cyan
  $check.Value2 = 'N'
  $excel.Run('CheckRoads') | Out-Null
  $r = Row3; Write-Host ("  elig: '" + $r.elig + "'  reason: '" + $r.reason + "'")
  if ($r.elig -ne 'Non-federal aid - Rural Minor Collector') { throw "N: expected Non-federal aid - Rural Minor Collector, got '$($r.elig)'" }
  if ($r.reason -ne '') { throw "N: Review Reason should be blank, got '$($r.reason)'" }
  Write-Host "  PASSED" -ForegroundColor Green

  Write-Host "=== 2. STH 52, double-check Y (HPMS twin dropped -> HPMS agrees) ===" -ForegroundColor Cyan
  $check.Value2 = 'Y'
  $excel.Run('CheckRoads') | Out-Null
  $r = Row3; Write-Host ("  elig: '" + $r.elig + "'  reason: '" + $r.reason + "'")
  if ($r.elig -ne 'Non-federal aid - Rural Minor Collector') { throw "Y: the class-3 non-inventory twin must not turn this row into a review, got '$($r.elig)'" }
  if ($r.reason -ne '') { throw "Y: HPMS agrees once the twin is dropped, Review Reason should be blank, got '$($r.reason)'" }
  Write-Host "  PASSED" -ForegroundColor Green

  Write-Host "=== 3. State layer down -> HPMS fallback reads the inventory record ===" -ForegroundColor Cyan
  $check.Value2 = 'N'
  $wb.Names('Svc_WI_LOCAL_ROADS').RefersToRange.Value2 = $deadUrl
  $excel.Run('CheckRoads') | Out-Null
  $r = Row3; Write-Host ("  elig: '" + $r.elig + "'  reason: '" + $r.reason + "'  class: '" + $r.cls + "'")
  if ($r.elig -ne 'Non-federal aid - Rural Minor Collector') { throw "HPMS fallback should read Rural Minor Collector at STH 52, got '$($r.elig)'" }
  if ($r.reason -notlike 'HPMS fallback*') { throw "Fallback row should be tagged HPMS fallback, got '$($r.reason)'" }
  if ($r.cls -like '*Principal Arterial*') { throw "The class-3 non-inventory twin should have been dropped, class cell is '$($r.cls)'" }
  $wb.Names('Svc_WI_LOCAL_ROADS').RefersToRange.Value2 = ''
  Write-Host "  PASSED" -ForegroundColor Green

  Write-Host "=== 4. Double-check Y with HPMS down -> verdict stands, check reported unavailable ===" -ForegroundColor Cyan
  $check.Value2 = 'Y'
  $wb.Names('Svc_HPMS').RefersToRange.Value2 = $deadUrl
  $excel.Run('CheckRoads') | Out-Null
  $r = Row3; Write-Host ("  elig: '" + $r.elig + "'  reason: '" + $r.reason + "'")
  if ($r.elig -ne 'Non-federal aid - Rural Minor Collector') { throw "A failed HPMS check must not change the verdict, got '$($r.elig)'" }
  if ($r.reason -ne 'HPMS check unavailable') { throw "Review Reason should read 'HPMS check unavailable', got '$($r.reason)'" }
  $wb.Names('Svc_HPMS').RefersToRange.Value2 = ''
  Write-Host "  PASSED" -ForegroundColor Green

  if ($MismatchState) {
    Write-Host ("=== 5. Real mismatch at $MismatchLat, $MismatchLon ($MismatchState) ===") -ForegroundColor Cyan
    $wb.Names('JobState').RefersToRange.Value2 = $MismatchState
    SetPoint $MismatchLat $MismatchLon 'Mismatch'
    $check.Value2 = 'N'
    $excel.Run('CheckRoads') | Out-Null
    $n = Row3; Write-Host ("  N  elig: '" + $n.elig + "'  reason: '" + $n.reason + "'")
    if ($n.elig -notlike 'Federal aid*' -and $n.elig -notlike 'Non-federal aid*') { throw "N: expected a confident verdict from the state layer, got '$($n.elig)'" }
    $check.Value2 = 'Y'
    $excel.Run('CheckRoads') | Out-Null
    $y = Row3; Write-Host ("  Y  elig: '" + $y.elig + "'  reason: '" + $y.reason + "'")
    if ($y.elig -ne 'Review - Sources disagree') { throw "Y: expected Review - Sources disagree, got '$($y.elig)'" }
    if ($y.reason -notlike 'Sources disagree | HPMS: *') { throw "Y: Review Reason should name the HPMS class, got '$($y.reason)'" }
    $yellow = $sites.Cells(3, 16).DisplayFormat.Interior.Color
    Write-Host ("  status cell colour: $yellow")
    Write-Host "  PASSED" -ForegroundColor Green
  } else {
    Write-Host "=== 5. (skipped - no -MismatchState/-MismatchLat/-MismatchLon given) ===" -ForegroundColor DarkGray
  }

  Write-Host "=== cost of Y: Check Roads on 10 rows, N vs Y ===" -ForegroundColor Cyan
  $wb.Names('JobState').RefersToRange.Value2 = 'WI'
  $pts = @(@(45.169879,-89.102452), @(43.0389,-87.9065), @(45.4711,-89.7345), @(44.764850,-91.406533), @(43.765678,-91.174513),
           @(44.595298,-89.796621), @(43.516423,-88.225545), @(43.02159,-88.265124), @(44.293536,-91.462206), @(44.590891,-89.767131))
  for ($i = 0; $i -lt $pts.Count; $i++) {
    $sites.Cells(3 + $i, 4).Value2 = "Timing $($i + 1)"; $sites.Cells(3 + $i, 5).Value2 = [double]$pts[$i][0]; $sites.Cells(3 + $i, 6).Value2 = [double]$pts[$i][1]
  }
  $times = @{}
  foreach ($v in @('N', 'Y', 'N', 'Y')) {
    $check.Value2 = $v
    $sw = [Diagnostics.Stopwatch]::StartNew(); $excel.Run('CheckRoads') | Out-Null; $sw.Stop()
    if (-not $times.ContainsKey($v)) { $times[$v] = @() }
    $times[$v] += $sw.Elapsed.TotalSeconds
  }
  $nAvg = ($times['N'] | Measure-Object -Average).Average; $yAvg = ($times['Y'] | Measure-Object -Average).Average
  Write-Host ("  N: {0:N1} s   Y: {1:N1} s   for {2} rows  ->  {3:N2} s per row extra with Y" -f $nAvg, $yAvg, $pts.Count, (($yAvg - $nAvg) / $pts.Count))
  for ($i = 0; $i -lt $pts.Count; $i++) { Write-Host ("    row {0}: {1}  [{2}]" -f (3 + $i), [string]$sites.Cells(3 + $i, 16).Value2, [string]$sites.Cells(3 + $i, 17).Value2) }
  Write-Host "VERIFY PASSED" -ForegroundColor Green
}
finally {
  if ($wb) { $wb.Close($false) }
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  Remove-Item -LiteralPath $tempXlsm -Force -Confirm:$false -ErrorAction SilentlyContinue
}
