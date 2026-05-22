# Verification §5.4 (+ touches §5.7, §5.8): Classify Roads workflow.
#
# Writes the three §4.2 test coordinates + one out-of-state (TN) row to
# the Sites table, runs ClassifyAllRows, then reads back FHWA Class,
# Urban/Rural, ACUB Name, Eligibility and checks them against the
# expected outcomes in CLAUDE.md §4.2.
#
# Network-dependent: this hits mdotgis.state.mi.us (NFC + Route) and
# services.arcgis.com (NTAD ACUB). Both must be reachable from this
# workstation; cloud sandboxes can't run this.

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

# (row, lat, lon, expected eligibility substring, expected class substring, expected ACUB name)
$tests = @(
  @{ Row=2; Name='Kalamazoo Urban Minor Collector'; Lat=42.28536; Lon=-85.57025; ExpectElig='INELIGIBLE';  ExpectClass='Minor Collector'; ExpectAcub='Kalamazoo' }
  @{ Row=3; Name='Lansing Urban Local';             Lat=42.6911;  Lon=-84.5360;  ExpectElig='ELIGIBLE';    ExpectClass='Local';           ExpectAcub='Lansing'   }
  @{ Row=4; Name='Tawas Rural Local';               Lat=44.2700;  Lon=-83.5200;  ExpectElig='ELIGIBLE';    ExpectClass='Local';           ExpectAcub=''          }
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath)
  $sites = $wb.Worksheets('Sites')

  Write-Host "Clearing any prior test rows (2..10)..."
  $sites.Range($sites.Cells(2,1), $sites.Cells(10, 23)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear

  Write-Host "Writing test rows..."
  foreach ($t in $tests) {
    # Cast to typed scalars + use Value2 so PowerShell's COM late-binding
    # doesn't pick the String overload of the parameterized Value property.
    $sites.Cells([int]$t.Row, 4).Value2 = [string]$t.Name        # Site Name (col D)
    $sites.Cells([int]$t.Row, 6).Value2 = [double]$t.Lat         # Lat (col F)
    $sites.Cells([int]$t.Row, 7).Value2 = [double]$t.Lon         # Lon (col G)
  }

  $tracePath = Join-Path $env:TEMP 'RoadReviewer_classify_trace.txt'
  Write-Host ("Setting gHeadless=True and tracing to " + $tracePath)
  $excel.Run('SetHeadless', $true) | Out-Null
  $excel.Run('SetTrace', $tracePath) | Out-Null

  Write-Host "Running ClassifyAllRows (this hits MDOT + NTAD)..."
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $excel.Run('ClassifyAllRows') | Out-Null
  $sw.Stop()
  Write-Host ("  finished in " + $sw.Elapsed.TotalSeconds + "s") -ForegroundColor Cyan

  $excel.Run('SetHeadless', $false) | Out-Null
  $excel.Run('SetTrace', '') | Out-Null

  if (Test-Path -LiteralPath $tracePath) {
    Write-Host "=== HTTP/classify trace ===" -ForegroundColor Cyan
    Get-Content -LiteralPath $tracePath | ForEach-Object { Write-Host ("  " + $_) }
  }

  Write-Host "=== Results ===" -ForegroundColor Cyan
  $failures = @()
  foreach ($t in $tests) {
    $r = $t.Row
    $cls = [string]$sites.Cells($r, 17).Value2  # FHWA Class
    $ur  = [string]$sites.Cells($r, 18).Value2  # Urban/Rural
    $ac  = [string]$sites.Cells($r, 19).Value2  # ACUB Name
    $rn  = [string]$sites.Cells($r, 20).Value2  # Road Name
    $el  = [string]$sites.Cells($r, 21).Value2  # Eligibility
    Write-Host ""
    Write-Host ("  row {0}: {1}" -f $r, $t.Name)
    Write-Host ("    class       : {0}" -f $cls)
    Write-Host ("    urban/rural : {0}" -f $ur)
    Write-Host ("    ACUB        : {0}" -f $ac)
    Write-Host ("    road name   : {0}" -f $rn)
    Write-Host ("    eligibility : {0}" -f $el)
    if ($t.ExpectElig -and ($el -notlike ("*" + $t.ExpectElig + "*"))) {
      $failures += ("row " + $r + " eligibility: expected '" + $t.ExpectElig + "', got '" + $el + "'")
    }
    if ($t.ExpectClass -and ($cls -notlike ("*" + $t.ExpectClass + "*"))) {
      $failures += ("row " + $r + " class: expected to contain '" + $t.ExpectClass + "', got '" + $cls + "'")
    }
    if ($t.ExpectAcub -ne '' -and ($ac -notlike ("*" + $t.ExpectAcub + "*"))) {
      $failures += ("row " + $r + " ACUB: expected to contain '" + $t.ExpectAcub + "', got '" + $ac + "'")
    }
    if ($t.ExpectAcub -eq '' -and $ac -ne '') {
      # Tawas should be rural (no ACUB hit)
      if ($t.Row -eq 4) {
        $failures += ("row " + $r + " ACUB: expected blank (rural), got '" + $ac + "'")
      }
    }
  }

  # Cleanup the test data so the saved file stays empty for the user.
  $sites.Range($sites.Cells(2,1), $sites.Cells(10, 23)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $wb.Save()
  $wb.Close($true)

  if ($failures) {
    Write-Host ""
    Write-Host "FAILURES:" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
    throw "Classify verification failed"
  } else {
    Write-Host ""
    Write-Host "VERIFICATION PASSED (§5.4)" -ForegroundColor Green
  }
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
