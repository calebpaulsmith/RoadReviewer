# Verification §5.7 (Re-run Failed Rows) + §5.8 (State selector).
#
# §5.8: set State to an unwired value (MN — MI/IN/WI are all wired now,
#       see CLAUDE.md §4.2/§4.2a/§4.2b, so this must be one of the
#       still-unwired states). Run Check Roads on a coord that reports ACUB
#       regardless of state. Expect the ACUB lookup to still run, but
#       NFC/class/road name to be blank with a "class lookup not wired
#       for this state" verdict message.
#
# §5.7: with State back to MI, plant a fake "Failed - ..." marker in the
#       Federal Aid Status column of one row, run ReRunFailedRows,
#       confirm ONLY that row got reprocessed (others keep their values).
#
# Works against either product's build (the classify path is shared);
# default is the standard RoadReviewer.xlsm.

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

# Sites layout: row 1 toolbar, row 2 header, data from row 3.
# Columns: SiteName=4, Lat=5, Lon=6, Class=20, Urban/Rural=21, ACUB=22, Elig=25.

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath)
  $sites = $wb.Worksheets('Sites')

  Write-Host "Clearing rows..."
  $sites.Range($sites.Cells(3, 1), $sites.Cells(11, 29)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $excel.Run('SetHeadless', $true) | Out-Null
  $excel.Run('SetTrace', (Join-Path $env:TEMP 'RoadReviewer_state_trace.txt')) | Out-Null

  # ---- §5.8: MN state (unwired), classify a Detroit-area coord ----
  Write-Host ""
  Write-Host "=== §5.8: State=MN on Detroit coord ===" -ForegroundColor Cyan
  # Use the annotated dropdown value to also exercise modUtil.StateCode,
  # which must strip "(not wired)" back to the bare "MN" code.
  $wb.Names('JobState').RefersToRange.Value2 = 'MN (not wired)'
  $sites.Cells(3, 4).Value2 = 'MN-state Detroit'
  $sites.Cells(3, 5).Value2 = [double]42.331
  $sites.Cells(3, 6).Value2 = [double]-83.045
  $excel.Run('CheckRoads') | Out-Null

  $row_class = [string]$sites.Cells(3, 14).Value2
  $row_urban = [string]$sites.Cells(3, 15).Value2
  $row_acub  = [string]$sites.Cells(3, 16).Value2
  $row_elig  = [string]$sites.Cells(3, 19).Value2
  Write-Host ("  class: '" + $row_class + "'  urban/rural: '" + $row_urban + "'  ACUB: '" + $row_acub + "'  elig: '" + $row_elig + "'")
  if ($row_class -ne '') { throw "Row 3 should have blank class when state=MN (NFC not wired), got '$row_class'" }
  if ($row_acub -notlike "*Detroit*") { throw ("Row 3 ACUB should still resolve to Detroit (ACUB is nationwide), got '" + $row_acub + "'") }
  if ($row_elig -notlike "*not wired*") { throw ("Row 3 verdict should mention NFC not wired, got '" + $row_elig + "'") }
  Write-Host "  §5.8 PASSED" -ForegroundColor Green

  # ---- §5.7: switch back to MI, set up failure + clean rows, re-run failed ----
  Write-Host ""
  Write-Host "=== §5.7: Re-run only Failed rows ===" -ForegroundColor Cyan
  $wb.Names('JobState').RefersToRange.Value2 = 'MI'
  # Clear + plant: row 3 has a stale "Failed - simulated" mark (should retry);
  # row 4 has a previous OK classify (should NOT change).
  $sites.Range($sites.Cells(3, 1), $sites.Cells(11, 29)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $sites.Cells(3, 4).Value2 = 'Failed row (will retry)'
  $sites.Cells(3, 5).Value2 = [double]42.28536
  $sites.Cells(3, 6).Value2 = [double]-85.57025
  $sites.Cells(3, 19).Value2 = 'Failed - simulated test failure'   # Federal Aid Status

  $sites.Cells(4, 4).Value2 = 'Already-classified row (should not retry)'
  $sites.Cells(4, 5).Value2 = [double]42.6911
  $sites.Cells(4, 6).Value2 = [double]-84.5360
  $sites.Cells(4, 14).Value2 = 'PREVIOUS-CLASS-MARKER'   # FHWA Class
  $sites.Cells(4, 19).Value2 = 'Non-federal aid - Urban Local (sticky)'   # Federal Aid Status

  Write-Host "  Before re-run:"
  Write-Host ("    row 3 elig = '" + [string]$sites.Cells(3, 19).Value2 + "'")
  Write-Host ("    row 4 class = '" + [string]$sites.Cells(4, 14).Value2 + "'  elig = '" + [string]$sites.Cells(4, 19).Value2 + "'")

  $excel.Run('ReRunFailedRows') | Out-Null

  Write-Host "  After re-run:"
  $r3 = [string]$sites.Cells(3, 19).Value2
  $r4_class = [string]$sites.Cells(4, 14).Value2
  $r4_elig = [string]$sites.Cells(4, 19).Value2
  Write-Host ("    row 3 elig = '" + $r3 + "'")
  Write-Host ("    row 4 class = '" + $r4_class + "'  elig = '" + $r4_elig + "'")

  if ($r3 -notlike "*Federal aid*Urban Minor Collector*") { throw ("Row 3 should have been re-classified to 'Federal aid - Urban Minor Collector', got '" + $r3 + "'") }
  if ($r4_class -ne 'PREVIOUS-CLASS-MARKER') { throw ("Row 4 was reclassified when it shouldn't have been; class is now '" + $r4_class + "'") }
  if ($r4_elig -ne 'Non-federal aid - Urban Local (sticky)') { throw ("Row 4 Federal Aid Status changed when it shouldn't have; got '" + $r4_elig + "'") }
  Write-Host "  §5.7 PASSED" -ForegroundColor Green

  # Cleanup
  $sites.Range($sites.Cells(3, 1), $sites.Cells(11, 29)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $wb.Names('JobState').RefersToRange.Value2 = 'MI'
  $excel.Run('SetHeadless', $false) | Out-Null
  $excel.Run('SetTrace', '') | Out-Null
  $wb.Save()
  $wb.Close($true)

  Write-Host ""
  Write-Host "VERIFICATION PASSED (§5.7 + §5.8)" -ForegroundColor Green
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
