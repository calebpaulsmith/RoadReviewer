# Verification §5.4 (+ touches §5.7, §5.8): Classify Roads workflow.
#
# Writes the wired-state test coordinates from CLAUDE.md §4.2 (MI),
# §4.2a (IN) and §4.2b (WI) to the Sites table, runs ClassifyAllRows on
# each state in turn, then reads back FHWA Class, Urban/Rural, ACUB Name,
# Eligibility and checks them against the expected outcomes.
#
# Network-dependent: this hits mdotgis.state.mi.us (MI NFC + Route),
# gisdata.in.gov (IN NFC + road name), services5.arcgis.com (WI state
# trunk + local roads) and services.arcgis.com (NTAD ACUB, nationwide).
# All must be reachable from this workstation; cloud sandboxes can reach
# the IN/WI/ACUB hosts but not mdotgis.state.mi.us.

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

# (row, state, lat, lon, expected Federal Aid Status substring, expected class substring, expected ACUB name)
$tests = @(
  @{ Row=2; State='MI'; Name='Kalamazoo Urban Minor Collector'; Lat=42.28536; Lon=-85.57025;  ExpectElig='Federal aid';     ExpectClass='Minor Collector';  ExpectAcub='Kalamazoo'  }
  @{ Row=3; State='MI'; Name='Lansing Urban Local';             Lat=42.6911;  Lon=-84.5360;   ExpectElig='Non-federal aid'; ExpectClass='Local';            ExpectAcub='Lansing'    }
  @{ Row=4; State='MI'; Name='Tawas Rural Local';               Lat=44.2700;  Lon=-83.5200;   ExpectElig='Non-federal aid'; ExpectClass='Local';            ExpectAcub=''           }
  @{ Row=5; State='IN'; Name='Indianapolis Minor Collector';    Lat=39.7684;  Lon=-86.1581;   ExpectElig='Federal aid';     ExpectClass='Minor Collector';  ExpectAcub='Indianapolis' }
  @{ Row=6; State='IN'; Name='Hancock County Rural Local';      Lat=39.9876;  Lon=-86.0128;   ExpectElig='Non-federal aid'; ExpectClass='Local';            ExpectAcub=''           }
  @{ Row=7; State='WI'; Name='Milwaukee Minor Arterial';        Lat=43.0389;  Lon=-87.9065;   ExpectElig='Federal aid';     ExpectClass='Minor Arterial';   ExpectAcub='Milwaukee'  }
  @{ Row=8; State='WI'; Name='Tomahawk Rural Major Collector';  Lat=45.4711;  Lon=-89.7345;   ExpectElig='Federal aid';     ExpectClass='Major Collector';  ExpectAcub=''           }
  # Regression tests for the Increment 4 FederalAidVerdict/PrefixedClass fixes (CLAUDE.md §7a):
  # a rural class-6 segment must read "...Rural Minor Collector", never "...Rural Local",
  # and class 1-3 (here Interstate) must get the Urban/Rural prefix same as 4-6.
  @{ Row=9; State='WI'; Name='STH 52 Rural Minor Collector';    Lat=45.169879; Lon=-89.102452; ExpectElig='Non-federal aid - Rural Minor Collector'; ExpectClass='Minor Collector'; ExpectAcub='' }
  @{ Row=10; State='WI'; Name='I-94 Eau Claire Urban Interstate'; Lat=44.764850; Lon=-91.406533; ExpectElig='Federal aid - Urban Interstate'; ExpectClass='Interstate'; ExpectAcub='Eau Claire' }
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath)
  $sites = $wb.Worksheets('Sites')
  $tracePath = Join-Path $env:TEMP 'RoadReviewer_classify_trace.txt'
  $excel.Run('SetHeadless', $true) | Out-Null

  Write-Host "Clearing any prior test rows (2..10)..."
  $sites.Range($sites.Cells(2,1), $sites.Cells(10, 27)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear

  $failures = @()

  # JobState applies to the whole ClassifyAllRows run, not per-row, so each
  # wired state needs its own write-then-classify pass.
  foreach ($state in @('MI', 'IN', 'WI')) {
    $stateTests = $tests | Where-Object { $_.State -eq $state }
    if (-not $stateTests) { continue }

    Write-Host ""
    Write-Host ("=== State=" + $state + " ===") -ForegroundColor Cyan
    $wb.Names('JobState').RefersToRange.Value2 = $state

    foreach ($t in $stateTests) {
      # Cast to typed scalars + use Value2 so PowerShell's COM late-binding
      # doesn't pick the String overload of the parameterized Value property.
      $sites.Cells([int]$t.Row, 4).Value2 = [string]$t.Name        # Site Name (col D)
      $sites.Cells([int]$t.Row, 6).Value2 = [double]$t.Lat         # Lat (col F)
      $sites.Cells([int]$t.Row, 7).Value2 = [double]$t.Lon         # Lon (col G)
    }

    $excel.Run('SetTrace', $tracePath) | Out-Null
    Write-Host ("Running ClassifyAllRows for " + $state + "...")
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $excel.Run('ClassifyAllRows') | Out-Null
    $sw.Stop()
    Write-Host ("  finished in " + $sw.Elapsed.TotalSeconds + "s") -ForegroundColor Cyan
    $excel.Run('SetTrace', '') | Out-Null

    if (Test-Path -LiteralPath $tracePath) {
      Write-Host "  --- HTTP/classify trace ---"
      Get-Content -LiteralPath $tracePath | ForEach-Object { Write-Host ("    " + $_) }
    }

    foreach ($t in $stateTests) {
      $r = $t.Row
      $cls = [string]$sites.Cells($r, 19).Value2  # FHWA Class
      $ur  = [string]$sites.Cells($r, 20).Value2  # Urban/Rural
      $ac  = [string]$sites.Cells($r, 21).Value2  # ACUB Name
      $rn  = [string]$sites.Cells($r, 22).Value2  # Road Name (state-specific layer)
      $st  = [string]$sites.Cells($r, 23).Value2  # Street Name (Census TIGER)
      $el  = [string]$sites.Cells($r, 24).Value2  # Federal Aid Status
      Write-Host ""
      Write-Host ("  row {0}: {1}" -f $r, $t.Name)
      Write-Host ("    class       : {0}" -f $cls)
      Write-Host ("    urban/rural : {0}" -f $ur)
      Write-Host ("    ACUB        : {0}" -f $ac)
      Write-Host ("    road name   : {0}" -f $rn)
      Write-Host ("    street name : {0}" -f $st)
      Write-Host ("    eligibility : {0}" -f $el)
      if ($t.ExpectElig -and ($el -notlike ("*" + $t.ExpectElig + "*"))) {
        $failures += ($state + " row " + $r + " eligibility: expected '" + $t.ExpectElig + "', got '" + $el + "'")
      }
      if ($t.ExpectClass -and ($cls -notlike ("*" + $t.ExpectClass + "*"))) {
        $failures += ($state + " row " + $r + " class: expected to contain '" + $t.ExpectClass + "', got '" + $cls + "'")
      }
      if ($t.ExpectAcub -ne '') {
        if ($ac -notlike ("*" + $t.ExpectAcub + "*")) {
          $failures += ($state + " row " + $r + " ACUB: expected to contain '" + $t.ExpectAcub + "', got '" + $ac + "'")
        }
      } elseif ($ac -ne '') {
        $failures += ($state + " row " + $r + " ACUB: expected blank (rural), got '" + $ac + "'")
      }
    }
  }

  # Cleanup the test data so the saved file stays empty for the user.
  $sites.Range($sites.Cells(2,1), $sites.Cells(10, 27)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear
  $wb.Names('JobState').RefersToRange.Value2 = 'MI'
  $excel.Run('SetHeadless', $false) | Out-Null
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
