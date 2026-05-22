# Verification §5.7 (Re-run Failed Rows) + §5.8 (State selector).
#
# §5.8: set State to a non-MI value (WI). Run Classify on a coord that
#       both Michigan AND Wisconsin would happily report ACUB for.
#       Expect the ACUB lookup to still run, but NFC/class/road name
#       to be blank with a "class lookup not wired for this state"
#       eligibility message.
#
# §5.7: with State back to MI, plant a fake "Failed - ..." marker in the
#       Eligibility column of one row, run ReRunFailedClassifications,
#       confirm ONLY that row got reprocessed (others keep their values).

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
  $setup = $wb.Worksheets('Setup')

  Write-Host "Clearing rows..."
  $sites.Range($sites.Cells(2,1), $sites.Cells(10, 23)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $excel.Run('SetHeadless', $true) | Out-Null
  $excel.Run('SetTrace', (Join-Path $env:TEMP 'RoadReviewer_state_trace.txt')) | Out-Null

  # ---- §5.8: WI state, classify a Detroit-area coord ----
  Write-Host ""
  Write-Host "=== §5.8: State=WI on Detroit coord ===" -ForegroundColor Cyan
  $wb.Names('JobState').RefersToRange.Value2 = 'WI'
  $sites.Cells(2, 4).Value2 = 'WI-state Detroit'
  $sites.Cells(2, 6).Value2 = [double]42.331
  $sites.Cells(2, 7).Value2 = [double]-83.045
  $excel.Run('ClassifyAllRows') | Out-Null

  $row2_class = [string]$sites.Cells(2,17).Value2
  $row2_urban = [string]$sites.Cells(2,18).Value2
  $row2_acub  = [string]$sites.Cells(2,19).Value2
  $row2_elig  = [string]$sites.Cells(2,21).Value2
  Write-Host ("  class: '" + $row2_class + "'  urban/rural: '" + $row2_urban + "'  ACUB: '" + $row2_acub + "'  elig: '" + $row2_elig + "'")
  if ($row2_class -ne '') { throw "Row 2 should have blank class when state=WI (NFC not wired), got '$row2_class'" }
  if ($row2_acub -notlike "*Detroit*") { throw ("Row 2 ACUB should still resolve to Detroit (ACUB is nationwide), got '" + $row2_acub + "'") }
  if ($row2_elig -notlike "*not wired*") { throw ("Row 2 eligibility should mention NFC not wired, got '" + $row2_elig + "'") }
  Write-Host "  §5.8 PASSED" -ForegroundColor Green

  # ---- §5.7: switch back to MI, set up failure + clean rows, re-run failed ----
  Write-Host ""
  Write-Host "=== §5.7: Re-run only Failed rows ===" -ForegroundColor Cyan
  $wb.Names('JobState').RefersToRange.Value2 = 'MI'
  # Clear + plant: row 2 has a stale "Failed - simulated" mark (should retry);
  # row 3 has a previous OK Classify (should NOT change).
  $sites.Range($sites.Cells(2,1), $sites.Cells(10, 23)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $sites.Cells(2, 4).Value2 = 'Failed row (will retry)'
  $sites.Cells(2, 6).Value2 = [double]42.28536
  $sites.Cells(2, 7).Value2 = [double]-85.57025
  $sites.Cells(2, 21).Value2 = 'Failed - simulated test failure'   # Eligibility (col 21 = U)

  $sites.Cells(3, 4).Value2 = 'Already-classified row (should not retry)'
  $sites.Cells(3, 6).Value2 = [double]42.6911
  $sites.Cells(3, 7).Value2 = [double]-84.5360
  $sites.Cells(3, 17).Value2 = 'PREVIOUS-CLASS-MARKER'   # FHWA Class (col 17 = Q)
  $sites.Cells(3, 21).Value2 = 'Non-federal aid - Urban Local (sticky)'   # Federal Aid Status

  Write-Host "  Before re-run:"
  Write-Host ("    row 2 elig = '" + [string]$sites.Cells(2,21).Value2 + "'")
  Write-Host ("    row 3 class = '" + [string]$sites.Cells(3,17).Value2 + "'  elig = '" + [string]$sites.Cells(3,21).Value2 + "'")

  $excel.Run('ReRunFailedClassifications') | Out-Null

  Write-Host "  After re-run:"
  $r2 = [string]$sites.Cells(2,21).Value2
  $r3_class = [string]$sites.Cells(3,17).Value2
  $r3_elig = [string]$sites.Cells(3,21).Value2
  Write-Host ("    row 2 elig = '" + $r2 + "'")
  Write-Host ("    row 3 class = '" + $r3_class + "'  elig = '" + $r3_elig + "'")

  if ($r2 -notlike "*Federal aid*Urban Minor Collector*") { throw ("Row 2 should have been re-classified to 'Federal aid - Urban Minor Collector', got '" + $r2 + "'") }
  if ($r3_class -ne 'PREVIOUS-CLASS-MARKER') { throw ("Row 3 was reclassified when it shouldn't have been; class is now '" + $r3_class + "'") }
  if ($r3_elig -ne 'Non-federal aid - Urban Local (sticky)') { throw ("Row 3 Federal Aid Status changed when it shouldn't have; got '" + $r3_elig + "'") }
  Write-Host "  §5.7 PASSED" -ForegroundColor Green

  # Cleanup
  $sites.Range($sites.Cells(2,1), $sites.Cells(10, 23)).ClearContents()
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
