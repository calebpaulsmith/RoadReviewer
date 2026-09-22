# HPMS fallback (2026-09-21): when a state DOT class layer is unavailable the row
# is classified from FHWA's national HPMS layer, tagged "HPMS fallback" in Review
# Reason, tinted light blue (Review Reason + FHWA Class) and its state-site link
# bolded. Also covers the service-URL swap: the state layer is broken and then
# repaired purely by typing into the Sources sheet's Svc_MI_NFC cell - no rebuild.
#
# The "unavailable" layer is REAL: MDOT stopped Widget/NextGenPrFinderPub on
# 2026-09-21 and it answers HTTP 200 with {"error":{"code":500,...not started}}
# - exactly the shape that used to read as "no road found". Works against either
# product. Runs on a %TEMP% copy (never open the committed workbook read-write).
param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }
$tempXlsm = Join-Path $env:TEMP 'rr-verify-hpms-copy.xlsm'
Copy-Item -LiteralPath $XlsmPath -Destination $tempXlsm -Force

$deadUrl = 'https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353'
$BLUE = 16247773   # CLR_HPMS_FALLBACK, RGB(221,235,247)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($tempXlsm)
  $sites = $wb.Worksheets('Sites')
  $sites.Range($sites.Cells(3, 1), $sites.Cells(11, 30)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null
  $excel.Run('SetHeadless', $true) | Out-Null
  $excel.Run('SetTrace', (Join-Path $env:TEMP 'RoadReviewer_hpms_trace.txt')) | Out-Null
  $wb.Names('JobState').RefersToRange.Value2 = 'MI'
  $sites.Cells(3, 4).Value2 = 'Kalamazoo culvert'
  $sites.Cells(3, 5).Value2 = [double]42.28536
  $sites.Cells(3, 6).Value2 = [double]-85.57025

  Write-Host "=== state layer DOWN (dead URL pasted into Svc_MI_NFC) -> HPMS fallback ===" -ForegroundColor Cyan
  $wb.Names('Svc_MI_NFC').RefersToRange.Value2 = $deadUrl
  $excel.Run('CheckRoads') | Out-Null
  $elig = [string]$sites.Cells(3, 16).Value2; $reason = [string]$sites.Cells(3, 17).Value2; $cls = [string]$sites.Cells(3, 18).Value2
  Write-Host ("  elig: '$elig'`n  reason: '$reason'`n  class: '$cls'")
  if ($elig -ne 'Federal aid - Urban Minor Collector') { throw "HPMS fallback should still classify the Kalamazoo point as Federal aid - Urban Minor Collector, got '$elig'" }
  if ($reason -notlike '*HPMS fallback*verify on the state site*') { throw "Review Reason should carry the HPMS fallback tag, got '$reason'" }
  $cReason = $sites.Cells(3, 17).DisplayFormat.Interior.Color
  $cClass  = $sites.Cells(3, 18).DisplayFormat.Interior.Color
  $cElig   = $sites.Cells(3, 16).DisplayFormat.Interior.Color
  $bold    = $sites.Cells(3, 14).DisplayFormat.Font.Bold
  Write-Host ("  colours: reason=$cReason class=$cClass status=$cElig   state-link bold=$bold")
  if ($cReason -ne $BLUE -or $cClass -ne $BLUE) { throw "Review Reason + FHWA Class should be light blue ($BLUE) on a fallback row" }
  if ($cElig -eq $BLUE) { throw "Federal Aid Status must keep its verdict colour, not the blue" }
  if (-not $bold) { throw "The state-site link (col 14) should be bold on a fallback row" }
  Write-Host "  PASSED" -ForegroundColor Green

  Write-Host "=== URL swap, no rebuild: clear the override -> state layer answers again ===" -ForegroundColor Cyan
  $wb.Names('Svc_MI_NFC').RefersToRange.Value2 = ''
  $excel.Run('CheckRoads') | Out-Null
  $elig = [string]$sites.Cells(3, 16).Value2; $reason = [string]$sites.Cells(3, 17).Value2
  Write-Host ("  elig: '$elig'  reason: '$reason'")
  if ($elig -ne 'Federal aid - Urban Minor Collector') { throw "State layer should classify the point, got '$elig'" }
  if ($reason -like '*HPMS fallback*') { throw "A row answered by the state layer must not be tagged HPMS fallback" }
  if ($sites.Cells(3, 17).DisplayFormat.Interior.Color -eq $BLUE) { throw "Blue tint should be gone once the state layer answers" }
  if ($sites.Cells(3, 14).DisplayFormat.Font.Bold) { throw "State-site link should not be bold once the state layer answers" }
  Write-Host "  PASSED" -ForegroundColor Green

  Write-Host "=== both DOWN -> the row FAILS (re-runnable), never a silent 'no road found' ===" -ForegroundColor Cyan
  $wb.Names('Svc_MI_NFC').RefersToRange.Value2 = $deadUrl
  $wb.Names('Svc_HPMS').RefersToRange.Value2 = $deadUrl
  $excel.Run('CheckRoads') | Out-Null
  $elig = [string]$sites.Cells(3, 16).Value2
  Write-Host ("  elig: '$elig'")
  if ($elig -notlike 'Failed - NFC query*HPMS fallback*') { throw "With both sources down the row should be 'Failed - NFC query (...; HPMS fallback: ...)', got '$elig'" }
  Write-Host "  PASSED" -ForegroundColor Green
  Write-Host "VERIFY PASSED" -ForegroundColor Green
}
finally {
  if ($wb) { $wb.Close($false) }
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  Remove-Item -LiteralPath $tempXlsm -Force -Confirm:$false -ErrorAction SilentlyContinue
}
