# Verification §5.4 (+ touches §5.7, §5.8): Check Roads workflow.
#
# Writes the wired-state test coordinates from CLAUDE.md §4.2 (MI),
# §4.2a (IN) and §4.2b (WI) to the Sites table, runs CheckRoads on
# each state in turn, then reads back FHWA Class, Urban/Rural, ACUB Name,
# Federal Aid Status and checks them against the expected outcomes.
# The MI pass also includes an address-only row to exercise the
# auto-geocode step folded into Check Roads.
#
# Network-dependent: this hits mdotgis.state.mi.us (MI NFC + Route),
# gisdata.in.gov (IN NFC + road name), services5.arcgis.com (WI state
# trunk + local roads), services.arcgis.com (NTAD ACUB, nationwide) and
# geocoding.geo.census.gov (the address row). All must be reachable from
# this workstation.
#
# Works against either product's build (the classify path is shared);
# default is the standard RoadReviewer.xlsm.

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

# Sites layout: row 1 header, data from row 2 (row-1 toolbar retired). Tests
# use rows 3-12, all valid data rows; row 2 is left empty and skipped.
# Columns: SiteName=4, Lat=5, Lon=6, Address=8, Geocode=12,
#          Class=16, Urban/Rural=17, ACUB=18, RoadName=19, Street=20,
#          Elig=21, ReviewReason=22.
$FirstRow = 3
$LastTestRow = 12

# (row, state, lat, lon, expected Federal Aid Status substring, expected class substring, expected ACUB name)
$tests = @(
  @{ Row=3;  State='MI'; Name='Kalamazoo Urban Minor Collector'; Lat=42.28536; Lon=-85.57025;  ExpectElig='Federal aid';     ExpectClass='Minor Collector';  ExpectAcub='Kalamazoo'  }
  @{ Row=4;  State='MI'; Name='Lansing Urban Local';             Lat=42.6911;  Lon=-84.5360;   ExpectElig='Non-federal aid'; ExpectClass='Local';            ExpectAcub='Lansing'    }
  @{ Row=5;  State='MI'; Name='Tawas Rural Local';               Lat=44.2700;  Lon=-83.5200;   ExpectElig='Non-federal aid'; ExpectClass='Local';            ExpectAcub=''           }
  # Address-only row: Check Roads must geocode it (Census), then classify.
  # Michigan State Capitol - downtown Lansing, always urban; loose checks
  # (Geocoded status + a real verdict), the exact class isn't pinned.
  @{ Row=6;  State='MI'; Name='Geocode test - MI Capitol'; Address='100 N Capitol Ave, Lansing, MI 48933'; ExpectGeocode=$true; ExpectAcub='Lansing' }
  @{ Row=7;  State='IN'; Name='Indianapolis Minor Collector';    Lat=39.7684;  Lon=-86.1581;   ExpectElig='Federal aid';     ExpectClass='Minor Collector';  ExpectAcub='Indianapolis' }
  # CLAUDE.md §4.2a's table labeled this point "rural Hancock County", but a
  # live run shows it sits inside the Indianapolis, IN ACUB polygon (it's at
  # the Marion/Hamilton county line, not Hancock). Class 7 (Local) was the
  # live-verified part; the verdict is Urban Local accordingly.
  @{ Row=8;  State='IN'; Name='Indianapolis-area Urban Local';   Lat=39.9876;  Lon=-86.0128;   ExpectElig='Non-federal aid'; ExpectClass='Local';            ExpectAcub='Indianapolis' }
  @{ Row=9;  State='WI'; Name='Milwaukee Minor Arterial';        Lat=43.0389;  Lon=-87.9065;   ExpectElig='Federal aid';     ExpectClass='Minor Arterial';   ExpectAcub='Milwaukee'  }
  @{ Row=10; State='WI'; Name='Tomahawk Rural Major Collector';  Lat=45.4711;  Lon=-89.7345;   ExpectElig='Federal aid';     ExpectClass='Major Collector';  ExpectAcub=''           }
  # Regression tests for the Increment 4 FederalAidVerdict/PrefixedClass fixes (CLAUDE.md §7a):
  # a rural class-6 segment must read "...Rural Minor Collector", never "...Rural Local",
  # and class 1-3 (here Interstate) must get the Urban/Rural prefix same as 4-6.
  @{ Row=11; State='WI'; Name='STH 52 Rural Minor Collector';    Lat=45.169879; Lon=-89.102452; ExpectElig='Non-federal aid - Rural Minor Collector'; ExpectClass='Minor Collector'; ExpectAcub='' }
  @{ Row=12; State='WI'; Name='I-94 Eau Claire Urban Interstate'; Lat=44.764850; Lon=-91.406533; ExpectElig='Federal aid - Urban Interstate'; ExpectClass='Interstate'; ExpectAcub='Eau Claire' }
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath)
  $sites = $wb.Worksheets('Sites')
  $tracePath = Join-Path $env:TEMP 'RoadReviewer_classify_trace.txt'
  $excel.Run('SetHeadless', $true) | Out-Null

  Write-Host "Clearing any prior test rows ($FirstRow..$LastTestRow)..."
  $sites.Range($sites.Cells($FirstRow, 1), $sites.Cells($LastTestRow, 30)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null   # restore link-col formulas after the wide clear

  $failures = @()

  # JobState applies to the whole CheckRoads run, not per-row, so each
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
      $sites.Cells([int]$t.Row, 4).Value2 = [string]$t.Name        # Site Name
      if ($t.ContainsKey('Address')) {
        $sites.Cells([int]$t.Row, 8).Value2 = [string]$t.Address   # Address (no coords)
      } else {
        $sites.Cells([int]$t.Row, 5).Value2 = [double]$t.Lat       # Lat
        $sites.Cells([int]$t.Row, 6).Value2 = [double]$t.Lon       # Lon
      }
    }

    $excel.Run('SetTrace', $tracePath) | Out-Null
    Write-Host ("Running CheckRoads for " + $state + "...")
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $excel.Run('CheckRoads') | Out-Null
    $sw.Stop()
    Write-Host ("  finished in " + $sw.Elapsed.TotalSeconds + "s") -ForegroundColor Cyan
    $excel.Run('SetTrace', '') | Out-Null

    if (Test-Path -LiteralPath $tracePath) {
      Write-Host "  --- HTTP/classify trace ---"
      Get-Content -LiteralPath $tracePath | ForEach-Object { Write-Host ("    " + $_) }
    }

    foreach ($t in $stateTests) {
      $r = $t.Row
      $cls = [string]$sites.Cells($r, 16).Value2  # FHWA Class
      $ur  = [string]$sites.Cells($r, 17).Value2  # Urban/Rural
      $ac  = [string]$sites.Cells($r, 18).Value2  # ACUB Name
      $rn  = [string]$sites.Cells($r, 19).Value2  # Road Name (merged, with distances)
      $st  = [string]$sites.Cells($r, 20).Value2  # Street Name (Census TIGER)
      $el  = [string]$sites.Cells($r, 21).Value2  # Federal Aid Status
      $rr  = [string]$sites.Cells($r, 22).Value2  # Review Reason (yellow note)
      $gc  = [string]$sites.Cells($r, 12).Value2  # Geocode Status
      Write-Host ""
      Write-Host ("  row {0}: {1}" -f $r, $t.Name)
      Write-Host ("    class       : {0}" -f $cls)
      Write-Host ("    urban/rural : {0}" -f $ur)
      Write-Host ("    ACUB        : {0}" -f $ac)
      Write-Host ("    road name   : {0}" -f $rn)
      Write-Host ("    street name : {0}" -f $st)
      Write-Host ("    eligibility : {0}" -f $el)
      Write-Host ("    review note : {0}" -f $rr)
      if ($t.ContainsKey('ExpectGeocode')) {
        $glat = [string]$sites.Cells($r, 5).Value2
        $glon = [string]$sites.Cells($r, 6).Value2
        Write-Host ("    geocode     : {0}  (lat='{1}', lon='{2}')" -f $gc, $glat, $glon)
        if ($gc -ne 'Geocoded') { $failures += ($state + " row " + $r + " geocode status: expected 'Geocoded', got '" + $gc + "'") }
        if (-not $glat -or -not $glon) { $failures += ($state + " row " + $r + " geocode: lat/lon not filled") }
        if (-not $el -or $el.StartsWith('Failed')) { $failures += ($state + " row " + $r + " geocoded row got no verdict: '" + $el + "'") }
      }
      if ($t.ContainsKey('ExpectElig') -and $t.ExpectElig -and ($el -notlike ("*" + $t.ExpectElig + "*"))) {
        $failures += ($state + " row " + $r + " eligibility: expected '" + $t.ExpectElig + "', got '" + $el + "'")
      }
      if ($t.ContainsKey('ExpectClass') -and $t.ExpectClass -and ($cls -notlike ("*" + $t.ExpectClass + "*"))) {
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
  $sites.Range($sites.Cells($FirstRow, 1), $sites.Cells($LastTestRow, 30)).ClearContents()
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
