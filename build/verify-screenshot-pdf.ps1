# Verify the MANUAL screenshot flow end-to-end + the self-healing export
# geometry (2026-07-15):
#   PrepareMapPages (6 sites) -> 6 synthetic "Google Earth screenshots"
#   (2340x1562, aspect ~1.50 like real GE captures) -> InsertMapImages ->
#   ExportCombinedMapPdf -> PyMuPDF asserts 6 pages with the imagery cropped
#   to EXACTLY the 760x568 block (thin even frame, nothing outside the
#   print area) and the right image on the right page.
#
#   Then the DRIFT test: sabotage the sheet the way stale-layout workbooks
#   were observed to break (page rows re-heighted 142 -> 152, a picture
#   shoved a full row down) and export again. NormalizeMapLayoutForPrint +
#   SnapShapesToPages must produce a byte-different but geometry-IDENTICAL
#   PDF - this is the regression test for the "screenshots outside the
#   print area" bug (images printed 760x606 instead of 760x568).
#
# Needs the INSPECTOR build (WO/DI named ranges). Copies the workbook to
# %TEMP% itself - never opens the committed file read-write (§7d).

param([string]$XlsmPath = "C:\Users\caleb\OneDrive\Desktop\Scripts\RoadReviewer\Site Inspector Review Tool.xlsm")

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

$py = 'C:\Users\caleb\OneDrive\Desktop\Scripts\.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $py)) { $py = 'python' }

$work = Join-Path $env:TEMP 'rr-screenshot-verify'
if (Test-Path $work) { Remove-Item -Recurse -Force $work }
New-Item -ItemType Directory -Path $work | Out-Null
$tempXlsm = Join-Path $work 'inspector-copy.xlsm'
Copy-Item -LiteralPath $XlsmPath -Destination $tempXlsm
$outFolder = Join-Path $work 'out'
New-Item -ItemType Directory -Path $outFolder | Out-Null
$mapsFolder = Join-Path $outFolder 'maps'
New-Item -ItemType Directory -Path $mapsFolder | Out-Null

# ---- 6 synthetic screenshots, distinct fill colors, GE-like aspect ----
$genScript = Join-Path $work 'gen.py'
@'
import sys
from PIL import Image, ImageDraw
colors = [(200,60,60),(60,160,60),(60,60,200),(190,160,40),(150,60,170),(40,160,170)]
for i, c in enumerate(colors, start=1):
    im = Image.new("RGB", (2340, 1562), c)
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, 2339, 1561], outline=(0, 0, 0), width=24)
    im.save(rf"{sys.argv[1]}\Page_{i}.png")
print("generated 6")
'@ | Set-Content -Path $genScript -Encoding utf8
$genOut = & $py $genScript $mapsFolder
if ($genOut -notmatch 'generated 6') { throw "screenshot generation failed: $genOut" }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($tempXlsm)
  $excel.Run('SetHeadless', $true) | Out-Null

  $sites = $wb.Worksheets('Sites')
  $coords = @(
    @(42.28536, -85.57025), @(42.69110, -84.53600), @(44.27000, -83.52000),
    @(42.33100, -83.04500), @(43.03890, -87.90650), @(39.76840, -86.15810))
  for ($i = 0; $i -lt 6; $i++) {
    $r = 2 + $i
    $sites.Cells($r, 3).Value2 = [string]($i + 1)          # Site #
    $sites.Cells($r, 4).Value2 = "Verify Site $($i + 1)"   # Site Name
    $sites.Cells($r, 5).Value2 = [double]$coords[$i][0]    # Lat
    $sites.Cells($r, 6).Value2 = [double]$coords[$i][1]    # Lon
  }
  $wb.Names('JobWO').RefersToRange.Value2 = 'WOTEST'
  $wb.Names('JobDI').RefersToRange.Value2 = '1'
  $wb.Names('JobDisaster').RefersToRange.Value2 = 'TEST'
  $wb.Names('JobOutputFolder').RefersToRange.Value2 = $outFolder + '\'

  Write-Host "=== PrepareMapPages (6 sites) ===" -ForegroundColor Cyan
  $excel.Run('PrepareMapPages') | Out-Null
  $wsMap = $wb.Worksheets('Map Pages')

  Write-Host "=== InsertMapImages (6 Page_N.png screenshots) ===" -ForegroundColor Cyan
  $excel.Run('InsertMapImages') | Out-Null
  $placed = 0
  foreach ($sh in $wsMap.Shapes) { if ($sh.Name -like 'MapImage_Page_*') { $placed++ } }
  if ($placed -ne 6) { throw "expected 6 placed images, found $placed" }
  Write-Host "  6 images placed" -ForegroundColor Green

  Write-Host "=== Export 1: clean geometry ===" -ForegroundColor Cyan
  $excel.Run('ExportCombinedMapPdf') | Out-Null
  $pdf1 = Get-ChildItem -LiteralPath $outFolder -Filter '*Location Map.pdf' | Select-Object -First 1
  if (-not $pdf1) { throw "PDF 1 was not exported" }
  $pdf1Clean = Join-Path $work 'export-clean.pdf'
  Copy-Item -LiteralPath $pdf1.FullName -Destination $pdf1Clean

  Write-Host "=== Sabotage: stale-layout drift (rows 142->152, image shoved a row down) ===" -ForegroundColor Cyan
  # This mimics the observed corrupted demo workbook: page rows at the older
  # 152pt height and a picture anchored one row below its block.
  $firstPageRow = 23   # MAP_FIRST_PAGE_ROW (MAP_HEADER_ROWS=22 + 1)
  for ($r = $firstPageRow; $r -lt $firstPageRow + 24; $r++) { $wsMap.Rows($r).RowHeight = 152 }
  $img1 = $wsMap.Shapes('MapImage_Page_1')
  $img1.Top = $img1.Top + 142
  $img1.Height = $img1.Height + 38     # the observed 566 -> ~606 stretch
  Write-Host "  rows re-heighted + image 1 displaced/stretched"

  Write-Host "=== Export 2: must self-heal ===" -ForegroundColor Cyan
  $excel.Run('ExportCombinedMapPdf') | Out-Null
  $pdf2 = Get-ChildItem -LiteralPath $outFolder -Filter '*Location Map.pdf' | Select-Object -First 1
  $pdf2Healed = Join-Path $work 'export-healed.pdf'
  Copy-Item -LiteralPath $pdf2.FullName -Destination $pdf2Healed

  $wb.Close($false)

  Write-Host "=== PyMuPDF geometry assertions (both PDFs) ===" -ForegroundColor Cyan
  $checkScript = Join-Path $work 'check.py'
  @'
import sys
import fitz

# Expected geometry: 792x612 landscape Letter, 760x568 block centered ->
# imagery frame x in [14,19]/[773,778], y in [19,25]/[587,593].
EXPECT = dict(x0=(13.0, 19.0), x1=(772.0, 779.0), y0=(18.0, 26.0), y1=(586.0, 594.0))
COLORS = [(200,60,60),(60,160,60),(60,60,200),(190,160,40),(150,60,170),(40,160,170)]

def check(path, label):
    doc = fitz.open(path)
    assert doc.page_count == 6, f"{label}: page count {doc.page_count} != 6"
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=100)
        w, h = pix.width, pix.height
        s = 100 / 72.0
        # content bbox of non-white pixels
        minx = miny = 10**9
        maxx = maxy = -1
        for yy in range(0, h, 2):
            for xx in range(0, w, 2):
                r, g, b = pix.pixel(xx, yy)[:3]
                if not (r > 240 and g > 240 and b > 240):
                    if xx < minx: minx = xx
                    if xx > maxx: maxx = xx
                    if yy < miny: miny = yy
                    if yy > maxy: maxy = yy
        x0, y0, x1, y1 = minx/s, miny/s, (maxx+2)/s, (maxy+2)/s
        assert EXPECT["x0"][0] <= x0 <= EXPECT["x0"][1], f"{label} p{i+1}: left {x0:.1f}pt outside {EXPECT['x0']}"
        assert EXPECT["x1"][0] <= x1 <= EXPECT["x1"][1], f"{label} p{i+1}: right {x1:.1f}pt outside {EXPECT['x1']}"
        assert EXPECT["y0"][0] <= y0 <= EXPECT["y0"][1], f"{label} p{i+1}: top {y0:.1f}pt outside {EXPECT['y0']}"
        assert EXPECT["y1"][0] <= y1 <= EXPECT["y1"][1], f"{label} p{i+1}: bottom {y1:.1f}pt outside {EXPECT['y1']}"
        # right image on the right page: sample right-of-center (clear of the stamp)
        er, eg, eb = COLORS[i]
        pr, pg, pb = pix.pixel(int(w*0.75), int(h*0.5))[:3]
        assert abs(pr-er) < 40 and abs(pg-eg) < 40 and abs(pb-eb) < 40, \
            f"{label} p{i+1}: expected color {COLORS[i]}, sampled ({pr},{pg},{pb})"
        # stamp text present
        assert "WO #" in page.get_text(), f"{label} p{i+1}: stamp text missing"
        print(f"  {label} p{i+1}: imagery ({x0:.1f},{y0:.1f})-({x1:.1f},{y1:.1f}) = {x1-x0:.1f}x{y1-y0:.1f}pt OK")

check(sys.argv[1], "clean")
check(sys.argv[2], "healed")
print("GEOMETRY OK: both exports render 6 pages with the 760x568 block exactly filled")
'@ | Set-Content -Path $checkScript -Encoding utf8
  $pyOut = & $py $checkScript $pdf1Clean $pdf2Healed
  $pyOut | ForEach-Object { Write-Host $_ }
  if ($LASTEXITCODE -ne 0 -or ($pyOut -join ' ') -notmatch 'GEOMETRY OK') { throw "PyMuPDF geometry assertions failed" }

  Write-Host ""
  Write-Host "VERIFICATION PASSED (screenshot flow + self-healing export)" -ForegroundColor Green
}
catch {
  Write-Host ("VERIFICATION FAILED: " + $_.Exception.Message) -ForegroundColor Red
  throw
}
finally {
  try { $excel.Quit() } catch {}
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
