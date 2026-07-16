# Force VBA to compile the modules the build itself never executes.
#
# The build runs BuildWorkbook (modBuild), so a syntax error there fails the
# build - but modHttp / modClassify / modMaps are compiled lazily by VBA only
# when something calls into them, i.e. at the user's first "Check Roads". This
# check calls one harmless function from each of those modules, which forces
# VBA to compile the WHOLE module and surfaces any syntax error NOW.
#
# A compile error pops a modal even in hidden Excel, which would hang forever,
# so the work runs in a background job with a hard timeout: clean compile
# returns "COMPILE OK" in a second or two; a compile error times out and we
# report failure (exit 1) instead of hanging.

param([string]$XlsmPath)

$ErrorActionPreference = 'Stop'
$XlsmPath = [System.IO.Path]::GetFullPath($XlsmPath)
if (-not (Test-Path -LiteralPath $XlsmPath)) { throw "Workbook not found: $XlsmPath" }

$job = Start-Job -ScriptBlock {
  param($p)
  $x = New-Object -ComObject Excel.Application
  $x.Visible = $false; $x.DisplayAlerts = $false; $x.AutomationSecurity = 1
  try {
    # READONLY: the target is the committed workbook, and OneDrive AutoSave
    # persists any macro side effect of a read-write open (§7d) - one no-op
    # hook that revealed a hidden sheet got silently saved that way.
    $wb = $x.Workbooks.Open($p, 0, $true)
    # Each Run compiles that function's ENTIRE module (all procedures).
    $x.Run('SetHeadless', $true) | Out-Null           # modUtil
    [void]$x.Run('BareStateCode', 'MI')               # modUtil
    [void]$x.Run('FunctionalSystemLabel', 1)          # modConstants
    [void]$x.Run('UrlEncode', 'x')                    # modHttp (geometry/distance helpers)
    [void]$x.Run('BufferFeet')                        # modClassify (verdict/query logic)
    [void]$x.Run('ResolveOutputFolder')               # modMaps
    [void]$x.Run('ExportItemCount')                   # modExportMenu (pure hook)
    $x.Run('RemoveMapImages') | Out-Null              # modMapImage (no-op w/o MapPages)
    $x.Run('ReRunFailedImagery') | Out-Null           # modMapFetch (no-op: 0 pages + headless)
    [void]$x.Run('PdfSelfTest')                       # modPdf (direct PDF writer)
    $wb.Close($false)
    'COMPILE OK'
  } catch {
    'COMPILE ERROR: ' + $_.Exception.Message
  } finally {
    try { $x.Quit() } catch {}
  }
} -ArgumentList $XlsmPath

if (Wait-Job $job -Timeout 60) {
  $res = Receive-Job $job
  Remove-Job $job
  Write-Host $res
  if ($res -match 'COMPILE OK') { exit 0 } else { exit 1 }
} else {
  Stop-Job $job; Remove-Job $job -Force
  Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
  Write-Host "COMPILE CHECK TIMED OUT - the project has a compile error (VBE modal blocked the run)."
  exit 1
}
