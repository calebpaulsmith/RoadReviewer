# Verification §5.2 + §5.3: skeleton present, every button wired, Sites
# hyperlinks/validation working, product-specific surface correct. Pure
# inspection - does NOT call out to any remote service.
#
# Product-aware: reads the hidden RR_Product defined name baked in by
# build.ps1 (Standard -> RoadReviewer.xlsm, Inspector -> Site Inspector
# Review Tool.xlsm) and asserts that product's expected inputs, buttons
# and hidden columns. Run it once per built workbook.

param([string]$XlsmPath = (Join-Path $env:TEMP 'RoadReviewer.xlsm'))

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

# The hub sheet is named per product: "Start Here" (standard, visible) or
# "Tools and Exports" (inspector, hidden). The map workspace tab is "Map Pages".
# $startName / $mapName are set once the product is known (below); the expected
# sheet set is built from them.

# Canonical Sites layout: row 1 IS the header, data from row 2. The row-1
# toolbar was retired; all actions live on Start Here.
$HeaderRow = 1
$FirstDataRow = 2
# PR #37 layout: the two state map-link columns sit together (13, 14) with the
# AGOL column after (15); the verdict columns LEAD the reviewer block (16, 17).
$expectedHeaders = @('WO #','DI #','Site #','Site Name','Latitude','Longitude','Description (optional)','Address (optional)','Category (optional)','Costs (optional)','Work Completion (optional)','Geocode Status','NFC Layer (Map Viewer)','State NFC App','User-Defined AGOL Layer','Federal Aid Status','Review Reason','FHWA Class','Urban/Rural','ACUB Name','Road Name','Street Name','Google Maps','Street View','Bing','Google Earth','FEMA Viewer','FIRMette Portal','FIRMette Status','Map Status')

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($XlsmPath, $false, $true)   # ReadOnly=true so we don't dirty the file
  $proj = $wb.VBProject

  # ---- Which product is this file? ----
  $product = 'Inspector'   # missing name defaults to Inspector, same as the VBA
  try {
    $refersTo = [string]$wb.Names('RR_Product').RefersTo   # looks like ="Standard"
    if ($refersTo -match 'Standard') { $product = 'Standard' }
  } catch { }
  $isInspector = ($product -eq 'Inspector')
  Write-Host ("=== Product: {0} ===" -f $product) -ForegroundColor Cyan

  # Product-specific sheet names (see the header comment).
  $startName = if ($isInspector) { 'Tools and Exports' } else { 'Start Here' }
  $mapName = 'Map Pages'
  $expectedSheets = @($startName, 'Sites', $mapName, 'Sources')

  Write-Host "=== Sheets ===" -ForegroundColor Cyan
  $present = @{}
  foreach ($ws in $wb.Worksheets) { $present[$ws.Name] = $true; Write-Host ("  " + $ws.Name) }
  $missing = $expectedSheets | Where-Object { -not $present.ContainsKey($_) }
  if ($missing) { throw ("Missing sheets: " + ($missing -join ', ')) }
  Write-Host "  all expected sheets present" -ForegroundColor Green

  # Sources tab is hidden on the inspector product (still in the file), visible
  # on the standard product. xlSheetVisible = -1, xlSheetHidden = 0.
  $srcVisible = [int]$wb.Worksheets('Sources').Visible
  if ($isInspector -and $srcVisible -ne 0) { throw "Inspector build should hide the Sources tab (Visible=$srcVisible)" }
  if (-not $isInspector -and $srcVisible -ne -1) { throw "Standard build should show the Sources tab (Visible=$srcVisible)" }
  Write-Host ("  Sources tab " + $(if ($isInspector) { 'hidden' } else { 'visible' }) + " as expected") -ForegroundColor Green

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
  # $onActions = every button in the file; $startActions = Start Here only. The
  # product-surface assertions below key off Start Here, because the MapPages
  # tools panel (Prepare / Insert / Update Stamps / Export) now exists on BOTH
  # products - it's the map workflow's home, not an inspector-only feature.
  $btnCount = 0; $orphans = @(); $onActions = @{}; $startActions = @{}; $mapActions = @{}
  foreach ($ws in $wb.Worksheets) {
    foreach ($sh in $ws.Shapes) {
      $oa = ""
      try { $oa = [string]$sh.OnAction } catch {}
      if ($oa) {
        $btnCount++
        $onActions[$oa] = $true
        if ($ws.Name -eq $startName) { $startActions[$oa] = $true }
        if ($ws.Name -eq $mapName)   { $mapActions[$oa] = $true }
        $resolved = $publicSubs[$oa]
        $caption = $sh.TextFrame2.TextRange.Text
        $status = if ($resolved) { "OK   ($resolved)" } else { "ORPHAN"; $orphans += "$($ws.Name)::$caption -> $oa" }
        Write-Host ("  [{0,-12}] {1,-44} -> {2,-28} {3}" -f $ws.Name, $caption, $oa, $status)
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

  # Product button surface (Start Here).
  #   Standard Start Here = the FHWA + imagery hub: Check Roads, Re-run, the
  #     export dropdown Go, its own Output-Folder Browse, Build/Reset.
  #   Inspector Start Here = a slim landing page: an "Open Map Pages tab" jump
  #     (GoToMapPages), a roads dropdown Go (demoted optional FHWA), an export
  #     dropdown Go, Build/Reset. The map/FIRMette buttons are on MapPages now.
  $sharedStart = @('RunSelectedExport','BuildWorkbook','ResetWorkbookFull')
  if ($isInspector) {
    $inspectorStart = @('GoToMapPages','RunSelectedRoadsAction')
    $forbiddenStart = @('PrepareMapPages','ExportSitesToKML','DownloadFirmettes','ReRunFailedFirmettes','ExportCombinedMapPdf','InsertMapImages')
    foreach ($a in ($sharedStart + $inspectorStart)) {
      if (-not $startActions.ContainsKey($a)) { throw "Inspector Start Here missing button for: $a" }
    }
    foreach ($a in $forbiddenStart) {
      if ($startActions.ContainsKey($a)) { throw "Inspector Start Here should NOT have $a (it moved to MapPages)" }
    }
  } else {
    $standardStart = @('CheckRoads','ReRunFailedRows','SelectOutputFolder')
    foreach ($a in ($sharedStart + $standardStart)) {
      if (-not $startActions.ContainsKey($a)) { throw "Standard Start Here missing button for: $a" }
    }
  }
  # The map workflow lives on MapPages for BOTH products (hidden until opted-in
  # on standard). Output-Folder Browse is the one that differs: on the inspector
  # it's on MapPages (Output Folder is canonical there); on the standard product
  # it's on Start Here, so its MapPages has no Browse.
  # One-click hero (CreateMapPagesPdf, 2026-07-15) + FIRMettes, with the
  # individual steps (Prepare / Fetch / Export), the re-run/re-stamp refreshers
  # and the manual Google Earth alternative behind the Advanced toggle.
  foreach ($a in @('CreateMapPagesPdf','ToggleMapAdvanced','PrepareMapPages','FetchMapImagery','ReRunFailedImagery','ExportCombinedMapPdf','ExportSitesToKML','InsertMapImages','UpdateMapStamps','DownloadFirmettes','ReRunFailedFirmettes')) {
    if (-not $mapActions.ContainsKey($a)) { throw "MapPages tools panel missing button for: $a" }
  }
  # Advanced shapes ship hidden (collapsed) - the uncluttered default view.
  $mapSheet = $wb.Worksheets($mapName)
  foreach ($sh in $mapSheet.Shapes) {
    if ($sh.Name -like 'MapCtrl_Adv*' -and $sh.Name -ne 'MapCtrl_AdvToggle') {
      if ([bool]$sh.Visible) { throw "Advanced control should ship hidden: $($sh.Name)" }
    }
    if ($sh.Name -eq 'MapCtrl_AdvToggle' -and -not [bool]$sh.Visible) { throw "Advanced toggle itself must be visible" }
  }
  if ($isInspector) {
    if (-not $mapActions.ContainsKey('SelectOutputFolder')) { throw "Inspector MapPages missing the Output-Folder Browse" }
    # MapPages is the landing on the inspector; it carries the door to the hidden
    # Tools & Exports (SH_START) sheet.
    if (-not $mapActions.ContainsKey('GoToOtherTools')) { throw "Inspector MapPages missing the 'Exports & other tools' button" }
  }
  Write-Host "  product button surface correct (Start Here + MapPages tools)" -ForegroundColor Green

  Write-Host "=== Sheet roles / visibility ===" -ForegroundColor Cyan
  # Inspector: Map Pages is the landing (visible), Start Here demotes to a HIDDEN
  # Tools & Exports sheet. Standard: Start Here is the hub (visible), MapPages
  # hidden until opted in.
  $mpVisible = [int]$wb.Worksheets($mapName).Visible
  $shVisible = [int]$wb.Worksheets($startName).Visible
  if ($isInspector) {
    if ($mpVisible -ne -1) { throw "Inspector should show MapPages (Visible=$mpVisible)" }
    if ($shVisible -eq -1) { throw "Inspector should hide Start Here / Tools (Visible=$shVisible)" }
  } else {
    if ($mpVisible -eq -1) { throw "Standard should ship MapPages hidden (Visible=$mpVisible)" }
    if ($shVisible -ne -1) { throw "Standard should show Start Here (Visible=$shVisible)" }
  }
  Write-Host ("  " + $(if ($isInspector) { 'Map Pages landing (visible); Start Here hidden (Tools)' } else { 'Start Here landing (visible); Map Pages hidden (opt-in)' })) -ForegroundColor Green

  Write-Host "=== Named ranges ===" -ForegroundColor Cyan
  # All eight job named ranges now exist on BOTH products (the job block lives on
  # MapPages, which the standard product also has - just hidden). Their homes
  # differ: State/AGOL/Buffer on Start Here; WO/DI/Disaster/Applicant on MapPages;
  # Output Folder on Start Here (standard) or MapPages (inspector).
  foreach ($n in @('JobState','JobOutputFolder','JobAgolMap','JobBufferFeet','JobWO','JobDI','JobDisaster','JobApplicant')) {
    try { $r = $wb.Names($n); Write-Host ("  " + $n + " -> " + $r.RefersTo) } catch { throw "Missing named range: $n" }
  }
  Write-Host "  all job named ranges present" -ForegroundColor Green

  Write-Host "=== Sites headers (row $HeaderRow) ===" -ForegroundColor Cyan
  $sites = $wb.Worksheets('Sites')
  for ($c = 1; $c -le $expectedHeaders.Count; $c++) {
    $got = [string]$sites.Cells($HeaderRow, $c).Value2
    $want = $expectedHeaders[$c-1]
    if ($got -ne $want) { throw ("Header mismatch at col " + $c + ": got '" + $got + "', want '" + $want + "'") }
  }
  Write-Host ("  all " + $expectedHeaders.Count + " headers match constants") -ForegroundColor Green

  Write-Host "=== Product column hiding ===" -ForegroundColor Cyan
  # Inspector-only columns: WO(1), DI(2), FIRMette Status(29), Map Status(30)
  foreach ($c in @(1, 2, 29, 30)) {
    $hidden = [bool]$sites.Columns($c).Hidden
    if ($isInspector -and $hidden) { throw "Inspector build should show column $c" }
    if (-not $isInspector -and -not $hidden) { throw "Standard build should hide column $c" }
  }
  Write-Host ("  inspector-only columns " + $(if ($isInspector) { 'visible' } else { 'hidden' }) + " as expected") -ForegroundColor Green

  Write-Host "=== Auto-reviewer columns hidden on a fresh (unclassified) build ===" -ForegroundColor Cyan
  # Federal Aid Status(16) .. Street Name(22) ship hidden until CheckRoads
  # reveals them (BEFORE its loop since PR #37).
  foreach ($c in @(16, 22)) {
    if (-not [bool]$sites.Columns($c).Hidden) { throw "Reviewer column $c should be hidden on a fresh build (classifier hasn't run)" }
  }
  Write-Host "  reviewer columns 16..22 hidden until first Check Roads" -ForegroundColor Green

  Write-Host "=== Photo/NFC column hiding (PR #37, Earth unhidden 2026-07-15) ===" -ForegroundColor Cyan
  # BOTH products hide Google Maps(23), Bing(25) and FEMA Viewer(27) - the FEMA
  # pin lives in the AGOL column's default link now. Street View(24) and Google
  # Earth(26) are the two visible photo links in BOTH products (per user,
  # 2026-07-15: Earth used to be inspector-hidden).
  foreach ($c in @(23, 25, 27)) {
    if (-not [bool]$sites.Columns($c).Hidden) { throw "Column $c should be hidden by default on both products" }
  }
  if ([bool]$sites.Columns(24).Hidden) { throw "Street View (24) should stay visible" }
  if ([bool]$sites.Columns(26).Hidden) { throw "Google Earth (26) should be visible in both products" }
  # The two state map-link columns (13, 14) ride with the reviewer block on the
  # inspector (hidden until Check Roads); standard shows them always.
  foreach ($c in @(13, 14)) {
    $hidden = [bool]$sites.Columns($c).Hidden
    if ($isInspector -and -not $hidden) { throw "Inspector build should hide column $c on a fresh build" }
    if (-not $isInspector -and $hidden) { throw "Standard build should show column $c" }
  }
  Write-Host ("  photo links hidden per PR #37; NFC link columns " + $(if ($isInspector) { 'hidden' } else { 'visible' }) + " as expected") -ForegroundColor Green

  Write-Host "=== Sites toolbar retired + export dropdown present ===" -ForegroundColor Cyan
  # The row-1 toolbar was removed; assert no leftover RR_* buttons on Sites.
  foreach ($sh in $sites.Shapes) {
    if ($sh.Name -like 'RR_*') { throw "Unexpected leftover toolbar shape on Sites: $($sh.Name)" }
  }
  # The export picker lives on the hub sheet and drives RunSelectedExport.
  $startSheet = $wb.Worksheets($startName)
  $hasPicker = $false
  foreach ($sh in $startSheet.Shapes) { if ($sh.Name -eq 'RR_ExportPicker') { $hasPicker = $true } }
  if (-not $hasPicker) { throw "Start Here is missing the RR_ExportPicker export dropdown" }
  # Inspector also collapses Check Roads / Re-run / Photo Links into a second
  # dropdown (RR_RoadsPicker driving RunSelectedRoadsAction).
  if ($isInspector) {
    $hasRoads = $false
    foreach ($sh in $startSheet.Shapes) { if ($sh.Name -eq 'RR_RoadsPicker') { $hasRoads = $true } }
    if (-not $hasRoads) { throw "Inspector Start Here is missing the RR_RoadsPicker dropdown" }
  }
  Write-Host "  no leftover Sites toolbar; action dropdown(s) present on Start Here" -ForegroundColor Green

  Write-Host "=== Sites hyperlink resolution (test coord) ===" -ForegroundColor Cyan
  # Use the Kalamazoo test coord. Reopen READONLY - cell/name edits below happen
  # in memory (readonly only blocks SAVE), and Close($false) discards them, so
  # this can never persist test values (e.g. JobState='MI') into the committed
  # workbook the way an r/w reopen could under OneDrive.
  $wb.Close($false)
  $wb = $excel.Workbooks.Open($XlsmPath, $false, $true)
  $sites = $wb.Worksheets('Sites')
  # State ships BLANK now, which would make the NFC columns show the "Set State"
  # prompt; set it so the state-dependent links resolve for this resolution check
  # (the blank-State placeholder is exercised separately below).
  $wb.Names('JobState').RefersToRange.Value2 = 'MI'
  $sites.Cells($FirstDataRow, 4).Value2 = 'Test - Kalamazoo'   # Site Name
  $sites.Cells($FirstDataRow, 5).Value2 = [double]42.28536      # Lat
  $sites.Cells($FirstDataRow, 6).Value2 = [double]-85.57025     # Lon
  $excel.Calculate()
  # Verify each hyperlink formula resolves to a non-empty string. The AGOL
  # column (15) now defaults to the FEMA pin ("FEMA AGOL Map Viewer") when
  # JobAgolMap is blank, so it's included (PR #37).
  $linkCols = @{ 13='NFC Layer (Map Viewer)'; 14='State NFC App'; 15='User-Defined AGOL'; 23='Google Maps'; 24='Street View'; 25='Bing'; 26='Google Earth'; 27='FEMA Viewer'; 28='FIRMette Portal' }
  foreach ($k in $linkCols.Keys | Sort-Object) {
    $cell = $sites.Cells($FirstDataRow, $k)
    $f = [string]$cell.Formula
    $v = [string]$cell.Value2
    if (-not $f.StartsWith('=')) { throw "Col $k formula empty: $f" }
    if (-not $v) { throw ("Col $k ({0}) shows empty with a valid coord" -f $linkCols[$k]) }
    Write-Host ("  col {0,2} ({1,-16}) shows: '{2}'  -- formula intact: {3} chars" -f $k, $linkCols[$k], $v, $f.Length)
  }
  # PR #37 link labels: descriptive per-state text instead of "Open", the AGOL
  # column's FEMA default, and Wisconsin's local/trunk relabel + layer split.
  if ([string]$sites.Cells($FirstDataRow, 13).Value2 -ne 'Review NFC AGOL Layer') { throw ("Col 13 label (MI) should be 'Review NFC AGOL Layer', got '" + $sites.Cells($FirstDataRow, 13).Value2 + "'") }
  if ([string]$sites.Cells($FirstDataRow, 14).Value2 -ne 'Review State NFC Layer') { throw ("Col 14 label (MI) should be 'Review State NFC Layer', got '" + $sites.Cells($FirstDataRow, 14).Value2 + "'") }
  if ([string]$sites.Cells($FirstDataRow, 15).Value2 -ne 'FEMA AGOL Map Viewer') { throw ("Col 15 default should be 'FEMA AGOL Map Viewer', got '" + $sites.Cells($FirstDataRow, 15).Value2 + "'") }
  $wb.Names('JobState').RefersToRange.Value2 = 'WI'
  $excel.Calculate()
  if ([string]$sites.Cells($FirstDataRow, 13).Value2 -ne 'Review Local Roads Layer') { throw ("Col 13 label (WI) should be 'Review Local Roads Layer', got '" + $sites.Cells($FirstDataRow, 13).Value2 + "'") }
  if ([string]$sites.Cells($FirstDataRow, 14).Value2 -ne 'Review State Trunk Hwy Layer') { throw ("Col 14 label (WI) should be 'Review State Trunk Hwy Layer', got '" + $sites.Cells($FirstDataRow, 14).Value2 + "'") }
  if ([string]$sites.Cells($FirstDataRow, 13).Formula -notmatch 'Functional_Class_Local_Non_Prod') { throw "Col 13 WI branch should side-load the local-roads layer" }
  if ([string]$sites.Cells($FirstDataRow, 14).Formula -notmatch 'FFCL_gdb') { throw "Col 14 WI branch should side-load the state-trunk layer" }
  Write-Host "  PR #37 link labels + WI local/trunk split OK" -ForegroundColor Green
  $wb.Names('JobState').RefersToRange.Value2 = 'MI'
  $excel.Calculate()
  # The two state-dependent NFC columns (13, 14) carry a blank-State placeholder
  # that directs the user to set State (instead of silently linking to Michigan).
  foreach ($k in @(13, 14)) {
    if ([string]$sites.Cells($FirstDataRow, $k).Formula -notmatch 'Set State') {
      throw "Col $k should carry the blank-State 'Set State' placeholder branch"
    }
  }
  # Standard's placeholder is a real HYPERLINK to the JobState cell; inspector's
  # is plain text (its State cell is on a hidden sheet).
  if (-not $isInspector -and [string]$sites.Cells($FirstDataRow, 13).Formula -notmatch 'HYPERLINK\("#JobState"') {
    throw "Standard col 13 placeholder should hyperlink to #JobState"
  }
  # Fires only on the FIRST data row: give rows 2 AND 3 coords, blank JobState,
  # recalc; row 2 shows the directive, row 3 stays blank.
  $sites.Cells($FirstDataRow, 5).Value2 = [double]42.28536
  $sites.Cells($FirstDataRow+1, 5).Value2 = [double]42.29
  $sites.Cells($FirstDataRow+1, 6).Value2 = [double]-85.57
  $jsOld = [string]$wb.Names('JobState').RefersToRange.Value2
  $wb.Names('JobState').RefersToRange.Value2 = ''
  $excel.Calculate()
  $ph2 = [string]$sites.Cells($FirstDataRow, 13).Value2
  $ph3 = [string]$sites.Cells($FirstDataRow+1, 13).Value2
  if ($ph2 -notmatch 'Set State') { throw "Blank State should show the directive in col 13 row 2, got '$ph2'" }
  if ($ph3 -ne '') { throw "Col 13 row 3 should be BLANK when State is blank (first-row-only), got '$ph3'" }
  Write-Host ("  blank-State placeholder (row 2 only): '{0}'  row 3: '{1}'" -f $ph2, $ph3)
  $wb.Names('JobState').RefersToRange.Value2 = $jsOld
  $sites.Cells($FirstDataRow+1, 5).Value2 = ''
  $sites.Cells($FirstDataRow+1, 6).Value2 = ''
  $excel.Calculate()

  Write-Host "=== Sites validation ===" -ForegroundColor Cyan
  $latVal = $sites.Cells($FirstDataRow, 5).Validation
  $lonVal = $sites.Cells($FirstDataRow, 6).Validation
  # xlValidateDecimal = 2
  Write-Host ("  Latitude  validation type=" + $latVal.Type + " (2=Decimal)  formula1=" + $latVal.Formula1 + "  formula2=" + $latVal.Formula2)
  Write-Host ("  Longitude validation type=" + $lonVal.Type + " (2=Decimal)  formula1=" + $lonVal.Formula1 + "  formula2=" + $lonVal.Formula2)
  if ($latVal.Type -ne 2 -or $lonVal.Type -ne 2) { throw "Coordinate columns missing decimal validation" }
  if ($latVal.Formula1 -ne '-90' -or $latVal.Formula2 -ne '90') { throw "Latitude validation range wrong" }
  if ($lonVal.Formula1 -ne '-180' -or $lonVal.Formula2 -ne '180') { throw "Longitude validation range wrong" }

  Write-Host "=== Conditional formatting on Federal Aid Status columns ===" -ForegroundColor Cyan
  $r = $sites.Range($sites.Cells($FirstDataRow, 16), $sites.Cells($FirstDataRow, 22))
  $fcCount = $r.FormatConditions.Count
  Write-Host ("  format conditions count on class..eligibility row " + $FirstDataRow + ": " + $fcCount)
  if ($fcCount -lt 3) { throw ("Expected 3 conditional-format rules (federal aid / non-federal aid / review), got " + $fcCount) }

  Write-Host "=== Hub-sheet disclaimer + version ===" -ForegroundColor Cyan
  $start = $wb.Worksheets($startName)
  $startBlob = ''
  foreach ($cell in $start.UsedRange.Cells) { $startBlob += ([string]$cell.Value2) + "`n" }
  # Standard keeps the on-sheet disclaimer box. Inspector moved it to a dialog
  # shown on Check Roads, so its Start Here should NOT carry the box text.
  if ($isInspector) {
    if ($startBlob -match 'NOT AN AUTHORITATIVE') { throw "Inspector Start Here should NOT carry the disclaimer box (it moved to the Check Roads dialog)" }
  } else {
    if ($startBlob -notmatch 'NOT AN AUTHORITATIVE') { throw "Start Here is missing the disclaimer block" }
    # The body was trimmed to end at "informational only" (the eligibility /
    # boundary sentences were dropped per user request); the Sources sheet still
    # carries the long-form version.
    if ($startBlob -notmatch 'informational\s+only') { throw "Start Here disclaimer missing the 'informational only' clause" }
    # Match the dropped BODY sentence, not 'eligibility determination' on its own:
    # the header line ("...FHWA OR ELIGIBILITY DETERMINATION") would match that
    # under PowerShell's case-insensitive -match and give a false failure.
    if ($startBlob -match 'do NOT constitute') { throw "Start Here disclaimer should end at 'informational only'" }
    if ($startBlob -notmatch 'Federal-aid road checker and review tool') { throw "Start Here missing the new subtitle" }
  }
  # The PR/version stamp was removed from Start Here (it lives in the Sources
  # footer now), so assert it is NOT here rather than that it is.
  if ($startBlob -match 'PR #') { throw "Start Here should no longer carry the PR/version label" }
  Write-Host "  disclaimer surface correct for product; no version stamp on Start Here" -ForegroundColor Green

  Write-Host "=== Sources sheet content ===" -ForegroundColor Cyan
  $sources = $wb.Worksheets('Sources')
  $sourcesBlob = ''
  foreach ($cell in $sources.UsedRange.Cells) { $sourcesBlob += ([string]$cell.Value2) + "`n" }
  $sourcesCellCount = $excel.WorksheetFunction.CountA($sources.UsedRange)
  Write-Host ("  Sources sheet non-empty cells: " + $sourcesCellCount)
  if ($sourcesCellCount -lt 20) { throw "Sources sheet looks empty" }
  if ($sourcesBlob -notmatch 'BOUNDARY ROADS') { throw "Sources sheet missing the boundary-roads caveat" }
  if ($sourcesBlob -notmatch 'does NOT authoritatively') { throw "Sources sheet missing the not-authoritative disclaimer" }
  Write-Host "  boundary caveat + disclaimer present on Sources" -ForegroundColor Green

  # Service-URL override table (PR "WI layer swap"): the section + the Svc_ named
  # ranges the ServiceUrl() resolver reads must exist so a user can swap a layer.
  if ($sourcesBlob -notmatch 'SERVICE URLs') { throw "Sources sheet missing the Service URLs override section" }
  foreach ($svcName in @('Svc_WI_LOCAL_ROADS', 'Svc_WI_STATE_TRUNK', 'Svc_ACUB')) {
    $found = $false
    foreach ($nm in $wb.Names) { if ($nm.Name -eq $svcName -or $nm.Name -like ("*!" + $svcName)) { $found = $true; break } }
    if (-not $found) { throw "Sources sheet missing named range $svcName for the service-override table" }
  }
  Write-Host "  service-URL override table + Svc_ named ranges present on Sources" -ForegroundColor Green

  # Clean up the test row so the saved file stays empty
  $sites.Range($sites.Cells($FirstDataRow, 4), $sites.Cells($FirstDataRow, 6)).ClearContents()

  Write-Host "VERIFICATION PASSED ($product)" -ForegroundColor Green
  $wb.Close($false)
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
