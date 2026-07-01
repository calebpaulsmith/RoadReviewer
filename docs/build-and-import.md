# Building RoadReviewer.xlsm from the VBA source

The V1 macros live as text in [`src/`](../src) so they can be reviewed and
version-controlled. Excel can't be driven from the cloud build environment,
so you assemble the workbook once on your Windows laptop. It takes about two
minutes and needs nothing beyond Excel with macros enabled (no add-ins, no
"trust access to the VBA project object model" setting).

## One-time assembly

1. Open Excel and create a **new blank workbook**.
2. Press **Alt + F11** to open the VBA editor (VBE).
3. In the VBE, select **File → Import File…** (or just drag each file onto the
   *Project* pane) and import every `.bas` file from `src/`:
   - `modConstants.bas`
   - `modUtil.bas`
   - `modHttp.bas`
   - `modBuild.bas`
   - `modClassify.bas`
   - `modGeocode.bas`
   - `modImagery.bas`
   - `modMaps.bas`
   - `modExport.bas`
4. Back in Excel, press **Alt + F8**, choose **BuildWorkbook**, and click
   **Run**. This creates the Home, Setup, Sites and three workflow sheets,
   wires every button, and lays out the Sites table.
5. **File → Save As**, choose **Excel Macro-Enabled Workbook (\*.xlsm)**, and
   name it `RoadReviewer.xlsm`.

That's it. From then on you just open `RoadReviewer.xlsm`. If you ever change
the layout, the **Build / Reset Workbook** button on the Home sheet rebuilds
the sheets while preserving the data you typed into the Sites table.

## What works in this build

- **Skeleton** — Home / Setup / Sites + three numbered workflow sheets, all
  buttons wired (verification step §5.2).
- **Sites table** — hyperlink columns (Google, Street View, Bing, FEMA Viewer,
  FIRMette portal, Open-in-map) computed from lat/lon, decimal validation on
  the coordinate columns, INELIGIBLE rows auto-highlight red (§5.3).
- **Geocode Addresses** — fills lat/lon from the Address column via the Census
  geocoder; never overwrites coordinates already present (F4).
- **1. Classify Roads** — MI/IN/WI NFC class + NTAD ACUB urban/rural + road
  name, eligibility verdict, "Re-run Failed Rows" (§5.4, F7, F12).
- **2. Review Imagery** — opens the curated imagery set for the selected
  Sites row(s) (§5.5).
- **3. Maps & FIRMettes** — Download FIRMettes / Re-run Failed FIRMettes
  (FEMA Print FIRMette GP service), Prepare Map Pages, Export Combined
  Map PDF / Export Individual Map PDFs.
- **KML export** and **CSV export** of the full Sites table (F10).

See `CLAUDE.md` §7a for which of these have an automated verifier behind
them versus "built but needs a manual smoke test."

## Trusting the macros after a move (verification step §5.9)

When you copy the `.xlsm` from a download folder or a `OneDrive - FEMA` path,
Office may open it with macros disabled (Mark-of-the-Web). To clear it:
right-click the file → **Properties** → check **Unblock** at the bottom →
**OK**. Then reopen and click **Enable Content** on the yellow bar.
