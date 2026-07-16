# verify-output-folder.ps1 - regression test for the OneDrive-for-Business /
# SharePoint URL -> local synced folder mapping (modMaps.OneDriveLocalFolder).
#
# The bug this guards (SharePoint-library user, 2026-07-16): a workbook opened
# from a "OneDrive - FEMA" library reports ThisWorkbook.Path as an https://
# SharePoint URL. The mapper must resolve that back to the local sync folder;
# it used to hand back a mangled "<base>\https:\...sharepoint.com\..." path
# (Dir$ raised on the malformed leading tail and, under On Error Resume Next,
# execution fell INTO the Then branch). This builds a fake sync tree, points an
# OneDrive env var at it, and asserts the mapper returns the clean local folder
# - never a path containing "https"/"sharepoint".
#
# Opens the workbook READ-ONLY (only Application.Run of a pure function; no
# workbook mutation) per the §7d "never open the committed workbook read-write".

param(
  [string]$XlsmPath = "$PSScriptRoot\..\Site Inspector Review Tool.xlsm"
)
$ErrorActionPreference = 'Stop'

# --- build a fake OneDrive sync tree + dummy workbook file -------------------
$root = Join-Path $env:TEMP 'rr_od_test'
$fakeSync = Join-Path $root 'OneDrive - FEMA'
$localToolkit = Join-Path $fakeSync 'Desktop\Working\03. Data and Mapping\Mapping Toolkit'
$wbName = 'RR OneDrive Test.xlsm'
if (Test-Path $root) { Remove-Item $root -Recurse -Force }
New-Item -ItemType Directory -Path $localToolkit -Force | Out-Null
Set-Content -Path (Join-Path $localToolkit $wbName) -Value 'x' -Encoding ascii

# Point OneDriveCommercial at the fake sync root BEFORE Excel launches so the
# COM child inherits it (this is exactly what OneDrive sets for a business
# library). Note the SharePoint URL maps ".../Documents/" onto the sync ROOT,
# so "Documents" is NOT a local folder - the mapper must skip it.
$env:OneDriveCommercial = $fakeSync

$sharepointUrl = 'https://usfema-my.sharepoint.com/personal/04927734585_fema_dhs_gov1/Documents/Desktop/Working/03.%20Data%20and%20Mapping/Mapping%20Toolkit'
$noMatchUrl    = 'https://usfema-my.sharepoint.com/personal/04927734585_fema_dhs_gov1/Documents/Nowhere/Missing'

$failures = @()
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $full = (Resolve-Path $XlsmPath).Path
  $wb = $excel.Workbooks.Open($full, $false, $true)   # readonly
  $excel.Run('SetHeadless', $true) | Out-Null

  # 1. Real SharePoint URL -> the fake local Toolkit folder.
  $mapped = [string]$excel.Run('OneDriveLocalFolder', $sharepointUrl, $wbName)
  Write-Host ("mapped: '" + $mapped + "'")
  $expected = $localToolkit
  if ($mapped.TrimEnd('\') -ne $expected.TrimEnd('\')) {
    $failures += "SharePoint URL did not map to the local Toolkit folder. Expected '$expected', got '$mapped'"
  }
  if ($mapped -match 'https' -or $mapped -match 'sharepoint') {
    $failures += "Mapped path still contains a URL fragment (the mangled-concatenation bug): '$mapped'"
  }

  # 2. A URL whose tail doesn't exist locally -> "" (caller then falls back).
  $none = [string]$excel.Run('OneDriveLocalFolder', $noMatchUrl, $wbName)
  Write-Host ("no-match: '" + $none + "'")
  if ($none -ne '') { $failures += "Expected '' for an unmappable URL, got '$none'" }

  $excel.Run('SetHeadless', $false) | Out-Null
  $wb.Close($false)
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failures) {
  Write-Host ""
  Write-Host "FAILURES:" -ForegroundColor Red
  $failures | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
  throw "Output-folder mapping verification failed"
} else {
  Write-Host ""
  Write-Host "VERIFICATION PASSED (OneDrive/SharePoint URL -> local folder)" -ForegroundColor Green
}
