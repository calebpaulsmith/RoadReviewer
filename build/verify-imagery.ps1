# Verification for Fetch Imagery (PR #35): the auto aerial-download path on
# Map Pages.
#
#   PrepareMapPages (3 sites) -> sabotage row 2's latitude -> FetchMapImagery
#   (expect 2 placed / 1 Failed, batch keeps going) -> restore + sentinel ->
#   ReRunFailedImagery (ONLY the failed row is retried) -> ExportCombinedMapPdf
#   -> PyMuPDF asserts: page count == site count, attribution + stamp text on
#   every page, real imagery pixels, red site pin at page center.
#   (Per §9.8: trust page count + rendered pixels, never get_image_rects.)
#
# Network: services.arcgisonline.com World_Imagery export (one ~1 MB PNG per
# site). Inspector product recommended (Map Pages is its landing), but the
# subs are shared so either build works.
#
# The workbook is COPIED to %TEMP% and the copy is opened - never the committed
# file (OneDrive AutoSave persists macro effects into any read-write open, §7d).

param([string]$XlsmPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Site Inspector Review Tool.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

$workCopy = Join-Path $env:TEMP 'rr-verify-imagery.xlsm'
Copy-Item -LiteralPath $XlsmPath -Destination $workCopy -Force

$outFolder = Join-Path $env:TEMP 'RoadReviewer_Imagery'
if (Test-Path -LiteralPath $outFolder) { Remove-Item -LiteralPath $outFolder -Recurse -Force }
New-Item -ItemType Directory -Path $outFolder -Force | Out-Null

$trace = Join-Path $env:TEMP 'RoadReviewer_imagery_trace.txt'

function Test-Shape($ws, [string]$name) {
  try { $null = $ws.Shapes($name); return $true } catch { return $false }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($workCopy)
  $sites = $wb.Worksheets('Sites')

  $excel.Run('SetHeadless', $true) | Out-Null
  $excel.Run('SetTrace', $trace) | Out-Null

  # Job values (inspector named ranges; Map Status = col 30)
  $wb.Names('JobWO').RefersToRange.Value2 = 'TEST'
  $wb.Names('JobDI').RefersToRange.Value2 = '0001'
  $wb.Names('JobDisaster').RefersToRange.Value2 = 'DR-TEST'
  $wb.Names('JobApplicant').RefersToRange.Value2 = 'Test Applicant'
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = ($outFolder + '\')

  # Three site rows (header row 1, data from row 2).
  $sites.Range($sites.Cells(2,1), $sites.Cells(12,30)).ClearContents()
  $excel.Run('RefreshSitesFormulas') | Out-Null
  $coords = @(
    @(1, 'Kalamazoo test', 42.28536, -85.57025),
    @(2, 'Lansing test',   42.6911,  -84.5360),
    @(3, 'Tawas test',     44.2700,  -83.5200)
  )
  for ($i = 0; $i -lt 3; $i++) {
    $r = 2 + $i
    $sites.Cells($r, 3).Value2 = [double]$coords[$i][0]   # Site #
    $sites.Cells($r, 4).Value2 = [string]$coords[$i][1]   # Site Name
    $sites.Cells($r, 5).Value2 = [double]$coords[$i][2]   # Lat
    $sites.Cells($r, 6).Value2 = [double]$coords[$i][3]   # Lon
  }

  Write-Host "=== PrepareMapPages (3 sites) ===" -ForegroundColor Cyan
  $excel.Run('PrepareMapPages') | Out-Null
  $wsMap = $wb.Worksheets('Map Pages')
  foreach ($n in 1..3) {
    if (-not (Test-Shape $wsMap "Textbox_Page_$n")) { throw "Missing page stamp Textbox_Page_$n" }
  }
  Write-Host "  3 map pages created" -ForegroundColor Green

  # ---- FetchMapImagery with one bad row (failure path: batch must continue) ----
  Write-Host "=== FetchMapImagery (row 3 sabotaged to non-numeric lat) ===" -ForegroundColor Cyan
  $sites.Cells(3, 5).Value2 = 'not-a-lat'
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $excel.Run('FetchMapImagery') | Out-Null
  $sw.Stop()
  Write-Host ("  finished in " + [math]::Round($sw.Elapsed.TotalSeconds,1) + "s")

  $s1 = [string]$sites.Cells(2, 30).Value2
  $s2 = [string]$sites.Cells(3, 30).Value2
  $s3 = [string]$sites.Cells(4, 30).Value2
  Write-Host "  Map Status: [$s1] [$s2] [$s3]"
  if ($s1 -notlike 'Imagery placed:*') { throw "Row 2 (Kalamazoo) not placed: '$s1'" }
  if ($s2 -notlike 'Failed - imagery:*') { throw "Row 3 (sabotaged) should be Failed: '$s2'" }
  if ($s3 -notlike 'Imagery placed:*') { throw "Row 4 (Tawas) not placed - batch did not continue past the failure: '$s3'" }

  foreach ($n in @(1,3)) {
    foreach ($p in @("MapImage_Page_$n", "MapPin_Page_$n", "MapAttr_Page_$n")) {
      if (-not (Test-Shape $wsMap $p)) { throw "Missing shape $p" }
    }
  }
  if (Test-Shape $wsMap 'MapImage_Page_2') { throw "Failed page 2 should have NO image" }
  Write-Host "  pages 1+3 have image+pin+attribution; failed page 2 has none" -ForegroundColor Green

  foreach ($f in @('Site_1.png','Site_3.png')) {
    $p = Join-Path (Join-Path $outFolder 'maps') $f
    if (-not (Test-Path -LiteralPath $p)) { throw "PNG copy missing: $p" }
    if ((Get-Item -LiteralPath $p).Length -lt 50000) { throw "PNG copy suspiciously small: $p" }
  }
  Write-Host "  maps\Site_N.png copies written (manual-flow interop)" -ForegroundColor Green

  # ---- ReRunFailedImagery: only the failed row is retried ----
  Write-Host "=== ReRunFailedImagery (row fixed; sentinel proves others skipped) ===" -ForegroundColor Cyan
  $sites.Cells(3, 5).Value2 = [double]42.6911
  $sites.Cells(2, 30).Value2 = 'SENTINEL - must not be touched'
  $excel.Run('ReRunFailedImagery') | Out-Null
  $s1 = [string]$sites.Cells(2, 30).Value2
  $s2 = [string]$sites.Cells(3, 30).Value2
  if ($s2 -notlike 'Imagery placed:*') { throw "Re-run did not fix row 3: '$s2'" }
  if ($s1 -ne 'SENTINEL - must not be touched') { throw "Re-run touched a non-failed row: '$s1'" }
  if (-not (Test-Shape $wsMap 'MapImage_Page_2')) { throw "Page 2 image missing after re-run" }
  if (-not (Test-Shape $wsMap 'MapPin_Page_2')) { throw "Page 2 pin missing after re-run" }
  Write-Host "  only the failed row was retried; page 2 now has image+pin" -ForegroundColor Green

  # ---- Imagery-source override (Map Pages "Imagery URL" cell) ----
  # Re-point page 2's fetch at World_Street_Map, pasted WITH a query string and
  # a trailing /export (proves ImageryServiceBase normalizes a browser-copied
  # URL), and confirm the attribution line switches to the custom host.
  Write-Host "=== Imagery URL override (World_Street_Map) ===" -ForegroundColor Cyan
  $wb.Names('JobImagerySvc').RefersToRange.Value2 = 'https://services.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/export?f=json'
  $sites.Cells(3, 30).Value2 = 'Failed - imagery: retest with override'
  $excel.Run('ReRunFailedImagery') | Out-Null
  $s2 = [string]$sites.Cells(3, 30).Value2
  if ($s2 -notlike 'Imagery placed:*') { throw "Override fetch failed: '$s2'" }
  $attr = [string]$wsMap.Shapes('MapAttr_Page_2').TextFrame2.TextRange.Text
  if ($attr -ne 'Imagery: services.arcgisonline.com') { throw "Attribution didn't switch to the custom host: '$attr'" }
  $wb.Names('JobImagerySvc').RefersToRange.Value2 = ''
  Write-Host "  override fetch OK; attribution = '$attr'" -ForegroundColor Green

  # ---- ExportCombinedMapPdf + PyMuPDF assertions ----
  Write-Host "=== ExportCombinedMapPdf ===" -ForegroundColor Cyan
  $excel.Run('ExportCombinedMapPdf') | Out-Null
  $pdf = Get-ChildItem -LiteralPath $outFolder -Filter '*Location Map.pdf' | Select-Object -First 1
  if (-not $pdf) { throw "Combined map PDF was not exported" }
  Write-Host ("  " + $pdf.Name + " (" + $pdf.Length + " bytes)")

  $py = 'C:\Users\caleb\OneDrive\Desktop\Scripts\.venv\Scripts\python.exe'
  if (-not (Test-Path -LiteralPath $py)) { $py = 'python' }
  $pyScript = Join-Path $env:TEMP 'rr_verify_imagery_pdf.py'
  @'
import sys
import fitz

doc = fitz.open(sys.argv[1])
assert doc.page_count == 3, f"page count {doc.page_count} != 3"

for i, page in enumerate(doc):
    text = page.get_text()
    # Page 2 was re-fetched from a custom source, so its attribution reads
    # "Imagery: <host>" instead of the Esri credit line.
    assert "Imagery" in text, f"page {i+1}: attribution line missing from text"
    assert "Site" in text, f"page {i+1}: stamp text missing"

assert "Esri" in doc[0].get_text(), "page 1: Esri attribution missing"

page = doc[0]
pix = page.get_pixmap(dpi=100)
w, h = pix.width, pix.height

# Imagery coverage: sampled fraction of non-near-white pixels must dominate.
nonwhite = total = 0
for yy in range(0, h, 7):
    for xx in range(0, w, 7):
        r, g, b = pix.pixel(xx, yy)[:3]
        total += 1
        if not (r > 235 and g > 235 and b > 235):
            nonwhite += 1
frac = nonwhite / total
assert frac > 0.5, f"page 1 looks blank: only {frac:.0%} non-white"

# Red site pin at the geometric center of the page.
cx, cy = w // 2, h // 2
found = False
for yy in range(cy - h // 12, cy + h // 12):
    for xx in range(cx - w // 12, cx + w // 12):
        r, g, b = pix.pixel(xx, yy)[:3]
        if r > 150 and g < 90 and b < 90:
            found = True
            break
    if found:
        break
assert found, "no red pin pixels near page center"

print(f"PYMUPDF OK: 3 pages, attribution+stamp text, {frac:.0%} imagery coverage, center pin found")
'@ | Set-Content -Path $pyScript -Encoding utf8
  $pyOut = & $py $pyScript $pdf.FullName
  Write-Host ("  " + $pyOut)
  if ($LASTEXITCODE -ne 0 -or $pyOut -notmatch 'PYMUPDF OK') { throw "PyMuPDF assertions failed" }
  Write-Host "  PDF structure + rendered pixels OK" -ForegroundColor Green

  $wb.Close($false)
  Write-Host ""
  Write-Host "VERIFICATION PASSED (Fetch Imagery)" -ForegroundColor Green
}
catch {
  Write-Host ("VERIFICATION FAILED: " + $_.Exception.Message) -ForegroundColor Red
  if (Test-Path $trace) {
    Write-Host "Last 25 trace lines:" -ForegroundColor Yellow
    Get-Content $trace | Select-Object -Last 25 | ForEach-Object { Write-Host ("  " + $_) }
  }
  throw
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Remove-Item -LiteralPath $workCopy -Force -ErrorAction SilentlyContinue
}
