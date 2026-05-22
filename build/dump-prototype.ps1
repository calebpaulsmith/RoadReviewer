# Extract VBA modules from the two prototype workbooks to text files we
# can read.

$ErrorActionPreference = 'Stop'
$outDir = Join-Path $PSScriptRoot 'prototype-vba'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  foreach ($wbName in @('Site Inspector Tool 1.xlsm', 'GPS Checker - TN updated 3.5.2026.xlsm')) {
    $path = Join-Path (Split-Path $PSScriptRoot -Parent) $wbName
    if (-not (Test-Path -LiteralPath $path)) { Write-Host "skip: $path"; continue }
    Write-Host ("Opening " + $wbName)
    $wb = $excel.Workbooks.Open($path, $false, $true)
    foreach ($comp in $wb.VBProject.VBComponents) {
      $cm = $comp.CodeModule
      $lines = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { '' }
      $stem = ($wbName -replace '\.xlsm$','') -replace '[^A-Za-z0-9]+','_'
      $file = Join-Path $outDir ($stem + '__' + $comp.Name + '.bas')
      Set-Content -LiteralPath $file -Value $lines -Encoding UTF8
      Write-Host ("  wrote " + $file + " (" + $cm.CountOfLines + " lines)")
    }
    $wb.Close($false)
  }
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
