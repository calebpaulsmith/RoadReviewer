# Map images: web renders, Excel places

## Why this split

The per-site figure is a **vector composite**: Esri street basemap tiles under
the ACUB urban-boundary polygon under the state's functional-class polylines,
each drawn with that source's **own published `drawingInfo.renderer`**.

Only MDOT (MI) exposes a `MapServer/export` operation. Indiana, Wisconsin and
the nationwide ACUB layer are AGOL-hosted **Query-only** feature services —
no `export`, no `/legend` (confirmed live 2026-07-03; see the header comment
above `renderCombinedCanvas` in `web/index.html`).

VBA has no canvas and no image compositor. A pure-VBA port of `reportFrame`
could only ever produce a **Michigan-only raster** and would silently drop
IN/WI/ACUB. So the split is:

- `web/index.html` renders the figure it *already draws* for the PDF report.
- `src/modMapImage.bas` places the resulting PNGs into the MapPages layout.

One renderer, one symbology, no drift between the web tool and the workbook.

## 1. Add the PNG export to `web/index.html`

`renderCombinedCanvas(...)` already returns a `<canvas>`, and
`buildPointReport(pt, radiusMeters)` already assembles everything it needs.
Add a sibling to the existing `Download PDF Report` button.

Markup — next to `#pdfBtn` (line ~205):

```html
<button id="pngBtn" title="One PNG per site, named Site_&lt;SiteNo&gt;.png — drop the folder into RoadReviewer.xlsm via 'Insert Map Images'">Download Map PNGs</button>
```

Script — next to the CSV/GeoJSON download helpers (lines ~923 / ~961), reusing
the same anchor-click pattern they use:

```js
// Per-site PNG export. Same figure as the PDF report, one file per site,
// named so modMapImage.FindImageForPage picks it up: Site_<SiteNo>.png,
// falling back to Page_<n>.png when a site has no number.
async function downloadMapPngs() {
  const radius = reportRadiusMeters();
  let n = 0;
  for (const pt of points) {              // same array the PDF path walks
    const canvas = await buildPointReport(pt, radius);
    if (!canvas) continue;
    n++;
    const stem = (pt.name && /^\d+$/.test(String(pt.name).trim()))
      ? "Site_" + String(pt.name).trim()
      : "Page_" + n;
    const blob = await new Promise(r => canvas.toBlob(r, "image/png"));
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = stem + ".png";
    a.click();
    URL.revokeObjectURL(a.href);
  }
}
$("pngBtn").addEventListener("click", downloadMapPngs);
```

> Check what `buildPointReport` actually returns before wiring this — if it
> returns a report object rather than the bare canvas, pull `.canvas` off it
> (or call `renderCombinedCanvas` directly with the same args). This is the
> one line worth stepping through in devtools rather than trusting.

Browsers will ask once about "download multiple files." Allow it, and the
files land together in your Downloads folder.

## 2. Import the VBA module

Add `modMapImage.bas` to the import list in `docs/build-and-import.md`,
**after** `modMaps.bas` (it depends on `SH_MAPPAGES`, `MAP_ROWS_PER_PAGE`,
`MAP_COLS_WIDE`, `MAP_PAGE_HEIGHT_PTS`, `ResolveOutputFolder`, `SitesSheet`,
`HasValidCoords`, `SheetExists`, `SetStatus`/`ClearStatus`):

```
- modMaps.bas
- modMapImage.bas      <-- new
- modExport.bas
```

Then re-run **BuildWorkbook** (Alt+F8) and re-save as `.xlsm`.

## 3. Wire the buttons

In `modBuild`'s Maps-sheet button block, add two buttons alongside
`Prepare Map Pages`:

| Caption | Macro |
| --- | --- |
| Insert Map Images | `InsertMapImages` |
| Remove Map Images | `RemoveMapImages` |

## 4. Use it

1. Classify your sites in `RoadReviewer.xlsm` as usual.
2. Paste the same coordinates into `web/index.html`, set **PDF map width**
   (that's the frame zoom), click **Download Map PNGs**.
3. Move the PNGs into `<output folder>\maps\` — that path is auto-detected,
   so no picker appears.
4. In the workbook: **Prepare Map Pages**, then **Insert Map Images**.
5. **Export Combined Map PDF** as before.

`Insert Map Images` is re-run safe — it deletes any `MapImage_Page_*` shapes
it previously created before placing new ones. Pages with no matching file
keep their "Paste screenshot here" placeholder, so a partial export degrades
gracefully instead of erroring.

## Behavior notes

- Images are embedded (`LinkToFile:=msoFalse`), so the workbook still renders
  after you delete the PNG folder.
- Aspect ratio is preserved and the image is centered in the page area; the
  WO/DI textbox is kept on top via `ZOrder msoSendToBack`.
- Page N maps to the Nth Sites row with valid coordinates — the same order
  `PrepareMapPages` walks. Prefer `Site_<SiteNo>.png` names, which survive
  row reordering; `Page_<n>.png` does not.
- No network calls in `modMapImage`. Nothing new for Zscaler to categorize.
