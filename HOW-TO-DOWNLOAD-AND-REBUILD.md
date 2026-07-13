# Opening the workbooks — "file format or file extension is not valid"

If Excel refuses to open `RoadReviewer.xlsm` or
`Site Inspector Review Tool.xlsm` with:

> Excel cannot open the file '…xlsm' because the file format or file
> extension is not valid. Verify that the file has not been corrupted and
> that the file extension matches the format of the file.

the workbook committed in this repository is **not** the problem — every
committed `.xlsm` is a valid, macro-enabled Excel file (verified: ZIP
container, OOXML parts, workbook relationships, and the `vbaProject.bin`
VBA stream all check out; `.gitattributes` marks `*.xlsm binary` so git
never rewrites line endings on a Windows checkout). The corruption is
almost always introduced **between GitHub and your PC**. Fix it with one
of the two paths below.

## 1. Download the file the right way (most common cause)

Excel shows the "not valid" error when the file it opened is actually an
**HTML page saved under an `.xlsm` name** — which is what you get if you
open the file's GitHub *view* page (a `github.com/.../blob/...` URL) and
use the browser's **Save As** / **Save link as**.

Download the real binary instead:

- On the file's GitHub page, click the **⤓ Download raw file** button
  (top-right of the file view), **or**
- Use the green **Code → Download ZIP** button and extract, **or**
- Open the raw URL directly:
  `https://raw.githubusercontent.com/calebpaulsmith/RoadReviewer/main/RoadReviewer.xlsm`

Then, on the downloaded file: **right-click → Properties →** tick
**Unblock** (bottom of the General tab) **→ OK**. Windows adds a
"Mark of the Web" to anything downloaded from the internet, which makes
Office distrust the macros (and occasionally refuse the file outright).

A file downloaded this way opens directly — no rebuild needed.

## 2. Rebuild the workbooks locally (guarantees a fresh, good copy)

Use this if your local working copy is stale/damaged, or after any change
under `src/`. **Excel must be installed** (the cloud build environment has
none — this only works on your Windows machine). Close Excel first; the
`.xlsm` files are locked while open.

```powershell
# from the repo root
git pull origin main

# build BOTH products into the repo root
#   -> RoadReviewer.xlsm  and  Site Inspector Review Tool.xlsm
& ".\build\build.ps1"

# commit and push the refreshed workbooks
git add RoadReviewer.xlsm "Site Inspector Review Tool.xlsm"
git commit -m "Rebuild workbooks locally"
git push origin <your-branch>
```

`build\build.ps1` imports every `.bas` from `src\`, bakes in each product
id, runs `BuildWorkbook`, and force-compiles the VBA (so a syntax error
can't ship silently). One-time prerequisite: Excel must allow programmatic
access to the VBA project — **File → Options → Trust Center → Trust Center
Settings → Macro Settings → Trust access to the VBA project object model**.

## Quick self-check that a file is a real workbook

An `.xlsm` is a ZIP archive. Rename a copy to `.zip` and try to open it —
a genuine workbook opens and shows `xl\`, `docProps\`, `[Content_Types].xml`.
If it won't open as a ZIP (or opening it in a text editor shows `<!DOCTYPE
html>`), you downloaded the HTML page, not the file — go back to path 1.
