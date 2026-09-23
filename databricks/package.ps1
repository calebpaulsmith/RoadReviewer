# Package the FHWA Road Checker as a Databricks App upload.
#
#   .\databricks\package.ps1                # -> dist\RoadReviewer-databricks-app.zip (tiles included, ~350 MB)
#   .\databricks\package.ps1 -NoTiles       # tiles left out: serve them from a UC Volume via RR_TILES_DIR
#   .\databricks\package.ps1 -StageOnly     # just the staged folder (for `databricks sync`), no zip
#
# The zip unpacks to ONE folder holding app.yaml, app.py, requirements.txt,
# setup\build_reviewer_tables.py, README.md and web\ (the site). Upload that
# folder to your workspace (Workspace > Import, or `databricks sync`) and
# point the app's source code path at it - see README.md.
param(
    [switch]$NoTiles,
    [switch]$StageOnly,
    [string]$OutDir = (Join-Path (Split-Path -Parent $PSScriptRoot) "dist")
)
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
$stage = Join-Path $OutDir "RoadReviewer-databricks-app"
if (Test-Path $stage) { Remove-Item -Recurse -Force $stage }
New-Item -ItemType Directory -Force (Join-Path $stage "setup") | Out-Null

foreach ($f in "app.yaml", "app.py", "requirements.txt", "README.md") {
    Copy-Item (Join-Path $PSScriptRoot $f) (Join-Path $stage $f)
}
Copy-Item (Join-Path $PSScriptRoot "setup\build_reviewer_tables.py") (Join-Path $stage "setup\build_reviewer_tables.py")

# the site, minus the tile archives when asked (they are the bulk of the size)
$webSrc = Join-Path $repo "web"
$webDst = Join-Path $stage "web"
robocopy $webSrc $webDst /E /NFL /NDL /NJH /NJS /NP $(if ($NoTiles) { "/XD"; (Join-Path $webSrc "tiles") }) | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy failed ($LASTEXITCODE)" }

$size = [math]::Round(((Get-ChildItem $stage -Recurse -File | Measure-Object Length -Sum).Sum / 1MB), 1)
Write-Host "staged $stage ($size MB)"
if ($StageOnly) { exit 0 }

$zip = Join-Path $OutDir "RoadReviewer-databricks-app.zip"
if (Test-Path $zip) { Remove-Item -Force $zip }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($stage, $zip, [System.IO.Compression.CompressionLevel]::Optimal, $true)
Write-Host "wrote $zip ($([math]::Round((Get-Item $zip).Length / 1MB, 1)) MB)"

exit 0
