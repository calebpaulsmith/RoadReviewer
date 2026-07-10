# Publishing the RoadReviewer classifier as an ArcGIS Online Web Tool, and calling it from Experience Builder

This guide takes you from the notebook file in `notebooks/roadreviewer_classify.ipynb` to a
working app where a user **pastes coordinates and sees classified, verdict-colored points on a
map** — no spreadsheet, no file upload. It is written for someone who has **never** published a
web tool or used Experience Builder. Every click is spelled out and every term is defined the
first time it appears.

Read **Section 0 first** — on a locked-down / managed ArcGIS Online (AGOL) organization you may
not have the rights this needs, and it is much better to find that out before you start.

**Term primer (you'll meet these below):**

| Term | Plain meaning |
|---|---|
| **AGOL** | ArcGIS Online — Esri's cloud GIS at `arcgis.com`, where your organization ("org") hosts maps and apps. |
| **Notebook** | A page of Python cells you run in AGOL's cloud (like Google Colab, but Esri's). Ours is `roadreviewer_classify.ipynb`. |
| **Web tool** | A published notebook that other apps can **call like a function** — you give it inputs, it runs in the cloud, and hands back outputs. Also called a "geoprocessing (GP) service". |
| **Parameter** | A named input or output of the web tool. Ours: two String inputs (`coords_text`, `radius_ft`) and one **Feature set** output. |
| **Feature set** | A bundle of map features (here, points) with their attributes and locations, passed around **in memory** — nothing is permanently saved. This is what the tool returns. |
| **Async job** | "Asynchronous": the tool doesn't answer instantly. You **submit a job**, it runs for a few seconds to a few minutes, and you **poll** (ask "done yet?") until the result is ready. |
| **Experience Builder (EB)** | AGOL's drag-and-drop **app builder** — you place widgets (map, buttons, panels) on a page and wire them together. |
| **Analysis widget** | The EB widget that **runs a web tool** from inside your app and puts the result on the map. |
| **Runtime** | The Python environment your notebook runs in. **Standard** (cheaper, no `arcpy`) vs **Advanced** (has `arcpy`). Ours needs only **Standard**. |

---

## 0. Prerequisites & gotchas (read before you touch anything)

**0.1 — You need specific privileges.** Ask your AGOL org administrator to confirm your account
has, at minimum:

- **A user type that allows notebooks.** Notebooks require a **Creator** (or higher, e.g. GIS
  Professional) user type **plus the ArcGIS Notebooks capability** enabled for your org. Viewer
  user types cannot create notebooks.
- **"Create, update, and delete content"** privilege (to make the notebook and the app).
- **"Publish web tools"** privilege (administrative privilege — to turn the notebook into a web
  tool). This is separate from publishing hosted layers.
- **"Run web tools / standard feature analysis"** privilege (to call the tool from the app).
- **"Publish web-hosted apps"** (to share the Experience).

If you don't know, **ask the admin first** — on a managed/government org these are frequently
withheld, and there is no self-service way to grant them. If "Publish web tools" is not
available to you, this whole path is blocked; the fallbacks are the drag-drop GeoJSON / Feature
Info approach in `docs/agol-review-app.md` §3B, or the coded page in that doc's §6, neither of
which needs publish rights.

**0.2 — Notebooks and web-tool runs consume CREDITS.** AGOL bills **credits** for compute:
running a notebook consumes credits per minute the runtime is active, and every web-tool run
(including each Experience Builder call and each self-test run) consumes credits too. On a
metered org this is real money. Practical consequences:

- Don't leave the notebook runtime idle-open; shut it down when you're done editing.
- **Turn off the self-test before publishing** (Section 3.5) so it doesn't re-run — and re-hit
  11 live services — on *every* production call.
- Ask your admin roughly how many credits you have and whether there's a budget.

**0.3 — Everything the tool queries is a public service.** The notebook queries the same public
state-DOT (MDOT/INDOT/WisDOT), USDOT NTAD (ACUB), and Census TIGER REST services the Excel and
web tools use. No credentials are needed for those; the credit cost is purely AGOL compute.

**0.4 — This runs on the Standard runtime.** The notebook deliberately uses only the `arcgis`
Python API, `requests`, and the standard library — **no `arcpy`** — so it runs on the cheaper
**Standard** runtime. Don't switch it to Advanced; you'll pay more for nothing.

---

## 1. Create the notebook in AGOL and load the file

1. Sign in to your organization at `https://www.arcgis.com` (or your org's URL, e.g.
   `https://yourorg.maps.arcgis.com`).
2. Top ribbon → **Content**.
3. **New item** (or **+ New** → **Notebook**). Choose **Notebook**.
4. **Runtime**: pick **ArcGIS Notebook Python 3 Standard**. (If a version picker appears, choose
   the newest; you'll confirm it's ≥ 8.0 in the next step. Do **not** pick "Advanced".)
5. The notebook opens. Now get our code in. Two ways:

   **A — Upload (preferred).** If your **Content** page's **New item** flow offers *"Your device"*,
   upload `notebooks/roadreviewer_classify.ipynb` directly as a new Notebook item, then open it.
   Some orgs also expose a **Files** panel inside the notebook editor (a folder icon in the left
   sidebar) where you can **Import**/**Upload** the `.ipynb` and then open it.

   **B — Paste (if upload is disabled).** Create a blank notebook (steps 1–4), open
   `notebooks/roadreviewer_classify.py` from this repo, and **copy each cell**. The `.py` mirror
   uses `# %%` markers to show where each notebook cell begins and `# %% [markdown]` for text
   cells. In the blank notebook, add a code cell for each `# %%` block and a Markdown cell (the
   **+ Markdown** button, or change a cell's type via the cell-type dropdown) for each
   `# %% [markdown]` block, pasting the code (drop the leading `# ` from markdown lines). The two
   files are kept identical, so either one reproduces the notebook exactly.

6. **Confirm the runtime is ≥ 8.0**, which AGOL requires for web-tool publishing. In the notebook,
   run this in a scratch cell (delete it afterward):
   ```python
   import arcgis, sys
   print("arcgis", arcgis.__version__, "| python", sys.version)
   ```
   The **runtime version** is shown on the notebook's item page under **Settings → Notebook →
   Runtime** (e.g. "ArcGIS Notebook Python 3 Standard 9.0"). If it reads below 8.0, open
   **Settings → Notebook → Runtime**, pick a newer Standard runtime, and save. (The `arcgis`
   package version is related but not the same number — what the publisher checks is the
   **runtime** version.)

---

## 2. Run it once, top to bottom, to confirm it works

Before publishing anything, prove the notebook runs.

1. The **first code cell** holds the sample inputs:
   ```python
   coords_text = """42.28536, -85.57025
   Culvert on Q Ave, 42.6911, -84.5360
   Site 12, 44.2700, -83.5200"""
   radius_ft = "250"
   ```
   Leave these as-is for the first run (they're three known Michigan points).
2. Run every cell in order: **Cell → Run All**, or click each cell and press **Shift+Enter**.
3. Watch two checkpoints:
   - The **self-test cell** (Section "10" in the notebook) prints `ok` lines and ends with
     **`11 passed, 0 failed`**. This asserts the Python verdicts match the JS `rr-core` for the
     confirmed CLAUDE.md coordinates. If it fails, a state service was down (see 7.5) — re-run.
   - The **"Package the results as a Feature set"** cell prints
     `Built FeatureSet with 3 feature(s).` On the AGOL runtime `arcgis` is preinstalled, so this
     succeeds. (It only prints the "arcgis not installed" skip message when run off-AGOL.)
4. The **last cell** shows `output_features` — the FeatureSet object. That's what the web tool
   will return.

If all three checkpoints pass, the notebook works. **Save** (the disk icon, or **Ctrl+S**).

---

## 3. Configure the web-tool parameters and publish

This is the part that turns a notebook into a callable tool. You'll tell AGOL which variables are
inputs, which is the output, and what types they are.

### 3.1 — Open the Parameters pane

In the notebook editor's **right sidebar**, click the **Parameters** tab (a sliders / `{ }`
icon). This pane is where web-tool inputs and outputs are declared.

### 3.2 — Declare the two String inputs

1. Click **Add parameter** (or **+**).
2. **Name**: type exactly `coords_text`. **Direction**: *Input*. **Type**: **String**.
   - *Default value* (optional): paste the same sample multi-line string. A default lets the tool
     run even if a caller forgets the parameter, and gives the publisher an example.
3. Add a second parameter: **Name** `radius_ft`, **Input**, **String**, default `250`.

> **Why String, not a number, for `radius_ft`?** The notebook parses and clamps it itself
> (`buffer_feet()` → 1–1000, default 250), exactly like the web tool. Keeping it a String avoids
> type-coercion surprises across the web-tool boundary.

### 3.3 — "Insert as variables" (wire inputs into the code)

With `coords_text` (and `radius_ft`) selected, click **Insert as variables** (or the *insert*
icon on the parameter). AGOL rewrites the **first code cell** so those variable assignments are
the ones a caller overrides at run time. This is why the notebook keeps both variables isolated
in that first cell — after "Insert as variables," that cell is the parameter-injection point.
Confirm the cell still simply reads:
```python
coords_text = "...caller value arrives here..."
radius_ft = "250"
```

### 3.4 — Declare the Feature set output and insert the write snippet

1. In **Parameters**, **Add parameter**. **Name**: `output_features`. **Direction**: *Output*.
   **Type**: **Feature set**.
2. Select it and click **Insert output variable** (or **Insert as variable**). AGOL appends a
   snippet to the **bottom** of the notebook that references the output. Our notebook already ends
   with a cell whose final expression is `output_features` (the FeatureSet built two cells up) —
   make sure AGOL's inserted snippet maps the **`output_features`** variable to this output
   parameter. If AGOL inserts a differently named placeholder, either rename it to
   `output_features` or set the last line to your FeatureSet variable. The rule: **the FeatureSet
   must be the last thing referenced at the very bottom.**
3. **Save** the notebook.

### 3.5 — Turn OFF the self-test before publishing

Find the cell:
```python
RUN_SELF_TEST = True
```
Change it to `RUN_SELF_TEST = False` (or delete the self-test markdown+code cells entirely).
Otherwise every production call re-runs 11 live-service classifications — slow and credit-wasting.
**Save.**

### 3.6 — Publish

1. From the notebook editor: **Notebook menu (⋯ or "Save as web tool" / "Create web tool")** →
   **Publish as web tool** (wording varies slightly by AGOL version — look for **web tool**).
2. Give it a **Title** (e.g. `RoadReviewer FHWA Classifier`), **tags** (`roadreviewer`), and a
   **summary**.
3. Confirm the parameter list it shows matches: inputs `coords_text` (String), `radius_ft`
   (String); output `output_features` (Feature set).
4. **Publish.** AGOL creates a **Web Tool item** (a geoprocessing service) in your Content.

### 3.7 — Set sharing

Open the new **Web Tool** item → **Share**. Choose:
- **Organization** — anyone in your org (including the Experience) can run it. Simplest.
- **Group** — only members of a specific group. Use this to limit who can spend credits.
- (Avoid **Everyone/Public** unless you intend an anonymous public tool — public runs still bill
  *your* org's credits.)

> **Async, by design.** Notebook-backed web tools run **asynchronously**: a caller *submits a job*,
> the tool runs in the cloud, and the caller *polls* until the FeatureSet is ready. Experience
> Builder's Analysis widget handles this submit-and-wait for you; if you call the REST endpoint
> directly (Section 5) you'll do the submit/poll by hand.

---

## 4. Test the web tool by itself (before building any app)

Prove the *published tool* works, independent of any app.

**Option A — from the item page.** Open the Web Tool item. Some AGOL versions add an **Analysis**
or **Open in Map Viewer** action, or show a **"Try it" / task form**. If present, fill
`coords_text` with a couple of lines (`42.28536, -85.57025`), leave `radius_ft` at `250`, and
**Run**. When the async job finishes, the returned points draw on the map — hover/click one to
see the `FedAidStatus` attribute.

**Option B — the REST submitJob form (always available).** Every web tool exposes a REST page:

1. Open the Web Tool item → **URL** (right column) → click it to open the **service directory**
   (`.../GPServer`), then click the **task** name (matches the tool title).
2. You'll see the parameters and a **Submit Job (POST)** or a link with a form. Enter
   `coords_text` and `radius_ft`, submit.
3. The response gives a **jobId**. Refresh the **job status** URL until it reads
   `esriJobSucceeded`, then open the **output_features** result link to see the returned
   FeatureSet JSON. Confirm each feature has `FedAidStatus`, `Verdict`, `VerdictColor`, etc.

If a point you expect to be classified comes back as
`ACUB only - class lookup not wired for this state`, its auto-detected state isn't one of the
wired three (MI/IN/WI) — that's expected for MN/IL/OH points.

---

## 5. Build the Experience (paste → classify → map)

Now assemble the app: a map with the authoritative reference layers, the Analysis widget pointed
at your tool, the returned points drawn on the map, and a stepper that shows each site's verdict
big and colored.

### 5.1 — (One-time) build the reference Web Map

The Analysis result is more useful over the authoritative road-class + urban-boundary layers.
Build a Web Map once, following **`docs/agol-review-app.md` §2** (full click list there). In brief:

1. **Content → Map** (opens **Map Viewer**).
2. **Add → Add layer from URL**, and paste the exact URLs from **`agol-review-app.md` §1**. Start
   with the states you cover plus the nationwide ACUB layer:
   - **MI (MDOT NFC):** `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353`
   - **IN (INDOT class):** `https://gisdata.in.gov/server/rest/services/Hosted/LRSE_Functional_Class/FeatureServer/22`
   - **WI (state trunk):** `https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/FFCL_gdb/FeatureServer/3`
   - **WI (local roads):** `https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/Functional_Class_Local_Non_Prod/FeatureServer/1`
   - **ACUB (all states):** `https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0`
   - (MN/IL/OH reference layers are also listed in `agol-review-app.md` §1 if you cover them.)
3. Style the **ACUB** polygon at ~10% fill with a distinct outline, placed **below** the road
   lines; leave each NFC line layer with its **own published symbology** (INDOT is single-symbol —
   optionally restyle by `functional_class` with the FHWA palette; see `agol-review-app.md` §2
   step 5). **Basemap → Streets** for road names.
4. **Save** the Web Map (title `RoadReviewer — FHWA Review (MI/IN/WI)`, tag `roadreviewer`), share
   to your org/group.

### 5.2 — Create the Experience and add a Map widget

1. From the Web Map item page → **Create Web App → Experience Builder** (or **Content → New item →
   Experience**). Pick a template with a map and a side panel (e.g. **"Pocket"** or **"Foldable"**).
2. In the builder, drag a **Map** widget onto the page (if the template doesn't already have one).
   In its settings, **select the Web Map from 5.1** as its data source. The reference NFC + ACUB
   layers now show on the map.

### 5.3 — Add the Analysis widget and point it at your web tool

1. From the widget panel (left), find **Analysis** and drag it onto the page (or into the side
   panel).
2. In the Analysis widget settings, choose the mode that **runs a tool** (labeled **"Analysis
   tool"** / **"Tools"** — as opposed to the built-in standard analysis list).
3. **Select tool → browse your organization → pick your published `RoadReviewer FHWA Classifier`
   web tool.** (If it doesn't appear, see 7.1.)
4. The widget reads the tool's parameters and renders input controls:
   - `coords_text` — a **text box**. This is where the user pastes coordinates. Make it
     multi-line if the widget offers it.
   - `radius_ft` — a text box; set its default to `250`.
   - Output `output_features` — set the widget to **add the result to the Map widget** from 5.2
     (there's an "output" / "result layer" setting; point it at your Map).
5. Save.

### 5.4 — Draw the returned points, colored by verdict

The tool returns a `VerdictColor` hex field and a `Verdict` bucket field on every feature, so
styling is **direct** — no Arcade needed (unlike the KML path in `agol-review-app.md` §5, where
the verdict had to be parsed out of a description string).

- In the Analysis widget's **result symbology** (or, if the result becomes a temporary map layer,
  in that layer's **Styles**), choose **Types (unique symbols)** on the **`Verdict`** field and set:
  - `Federal aid` → red (`#d73027`), `Non-federal aid` → green (`#1a9850`), `Review` → amber
    (`#b58900`). These are the exact `rr-core` bucket colors the tool also writes into
    `VerdictColor`, so if the widget lets you drive color directly from the `VerdictColor` field
    instead, use that and the colors come through automatically.
- *(Optional polish)* Add a **"Data/analysis complete" → zoom to result** action so the map frames
  the classified points the moment the run finishes.

### 5.5 — Add a Feature Info stepper (◀ / ▶, "1 of N") with a big colored verdict

1. Drag a **Feature Info** widget into the side panel.
2. Set its data mode to **"Interact with a Map widget"** and point it at the Map from 5.2 (this
   lets it read the runtime-added result layer — see `agol-review-app.md` §3B for why this mode
   matters).
3. Enable **Feature navigation** and **Show index** → you get **◀ / ▶** buttons and a **"1 of N"**
   counter to step through every classified site.
4. Configure the displayed fields, verdict first and prominent: show **`FedAidStatus`** (the full
   verdict text) large, then `FHWAClass`, `UrbanRural`, `ACUBName`, `RoadName`, `ReviewNote`,
   `Latitude`/`Longitude`.
5. **Coloring the verdict:** because `Verdict`/`VerdictColor` are **real fields here**, you can
   style the verdict text element's color directly from the field, or drop in the badge from
   `agol-review-app.md` §5 (there it reads a parsed field; here the field already exists, so point
   the same expression at **`$feature.FedAidStatus`** / **`$feature.VerdictColor`**). Example
   Arcade for a colored badge element, adapted to the real field:
   ```arcade
   var v = $feature.FedAidStatus;
   var c = $feature.VerdictColor;        // already the right hex, no parsing needed
   return `<div style="font-size:15px;font-weight:700;color:#fff;background:${c};
     padding:6px 10px;border-radius:6px;display:inline-block">${v}</div>`;
   ```

### 5.6 — Header, disclaimer, publish

Add a **Text** widget header with the disclaimer that mirrors the rest of the tool: *"Classifies
the road, not the project. A candidate for human review, not a federal-aid determination — verify
every point on the source map."* Then **Publish** (top-right) and **set sharing** (org/group) the
same way as the web tool.

**The end-user flow:** open the Experience → paste coordinates into the Analysis text box →
**Run** → points appear on the map colored red/green/amber → step through them with Feature Info's
◀ / ▶, reading each site's verdict. No spreadsheet, no file upload.

---

## 6. (Recap) how the pieces connect

```
   user pastes coords
          │
          ▼
  Analysis widget  ──submitJob(coords_text, radius_ft)──►  Web Tool  ──► notebook runs
  (in the EB app)                                          (async GP)      classify_point()
          ▲                                                                    │
          │◄──────────── output_features (Feature set) ◄────────── returns ────┘
          ▼
  Map widget draws points by Verdict/VerdictColor
          │
          ▼
  Feature Info ◀/▶ steps through each site, verdict shown big + colored
```

---

## 7. Verify + troubleshoot

### 7.1 — 5-minute check: does the Analysis widget actually list & run *your* tool?

Not every AGOL / Experience Builder version surfaces custom notebook web tools in the Analysis
widget the same way. **Before building the whole app**, do this quick check:

1. New scratch Experience → add a **Map** + an **Analysis** widget.
2. In Analysis, switch to the **run-a-tool** mode and **browse your org's tools**. Confirm
   `RoadReviewer FHWA Classifier` **appears** and can be **selected**.
3. Run it once with two pasted coordinates and confirm points come back on the map.

If it does **not** appear or won't run:
- Re-check the web tool's **sharing** (Section 3.7) — it must be shared to you / your org / the
  same group.
- Confirm your account has **"Run web tools"** (Section 0.1).
- Some versions only list web tools under a specific Analysis sub-mode ("Tools" vs "Standard
  analysis") — look for both.
- **Fallback if your EB version can't consume the tool:** use the no-tool path in
  `docs/agol-review-app.md` §3B — export the classified points as GeoJSON, load them with the
  **Add Data** widget, and step through with Feature Info. You lose the "paste-in-app" step but
  keep the map + stepper with zero publishing rights.

### 7.2 — "You don't have privileges to publish web tools" / tool won't publish

You're missing the **Publish web tools** administrative privilege (Section 0.1). Only an admin can
grant it. Until then, use the GeoJSON drag-drop fallback above.

### 7.3 — "Insufficient credits" / the run stops partway

You (or the org) are **out of credits** (Section 0.2). Ask the admin to top up or raise your
budget. To reduce burn: keep the self-test **off** (Section 3.5), classify smaller batches, and
shut the notebook runtime down when not editing.

### 7.4 — The async job times out / spins forever

- Large paste + a slow state service can exceed the Analysis widget's wait. Try a **smaller
  batch** (e.g. ≤ 25 points) first.
- A stuck job usually means an upstream service is hanging. Re-submit; the notebook's built-in
  retry (`http_get_json`, 3 tries with backoff) rides over brief blips, but a fully-down service
  will still fail the row.

### 7.5 — Empty FeatureSet, or every point says "Review - no road within N ft"

- **Empty output:** every pasted line failed to parse. Check the format — each line needs a
  CONUS-looking `lat lon` pair (lat 17–72, lon −180…−64). A leading site number alone isn't a
  coordinate.
- **"no road within N ft" everywhere for wired states:** widen `radius_ft` (e.g. `500`), or the
  points are genuinely off-network. For MN/IL/OH points, `ACUB only - class lookup not wired for
  this state` is **expected** (those states aren't wired for road class).

### 7.6 — MDOT 503 / transient state-service errors (mirror the repo's retry note)

Michigan's MDOT service throws **occasional transient 503s** (noted throughout CLAUDE.md and
`web/index.html`). The notebook already retries each request up to 3 times with linear backoff, so
most blips self-heal. If a whole run fails on MDOT, **just run it again** — the same coordinates
will usually succeed on the next attempt. (A Michigan point that intermittently returns
`Failed - NFC query (...)` is this, not a logic bug.)

---

## What matches the repo (parity guarantees)

- **Verdict logic** is a line-for-line port of `web/index.html`'s `rr-core` (`computeVerdict`,
  `classIsFederal`, `determineAcub`, `classifyPoint`) — same **250 ft** default buffer, same
  **nearest-segment-decides** rule, same **yellow-only-downgrades-green** reasons ("Second road
  close" / "Nearby FHWA road" / "Urban boundary edge"), and **red never downgrades**.
- **Buckets/colors/fields** match the repo's GeoJSON export (`rr-core` `BUCKET_GEO_LABEL` /
  `BUCKET_COLOR` and `modMaps.WriteSitesGeoJson`): `Name, FedAidStatus, Verdict, VerdictColor,
  FHWAClass, UrbanRural, ACUBName, RoadName, StreetName, ReviewNote, Category, Description,
  Latitude, Longitude`.
- **Data sources / filters** match `web/sources.html`: `RHRetireDate IS NULL` (MI),
  `record_status=5` (IN), WisDOT state-trunk-then-local order with the local category unpacked,
  browser User-Agent for MDOT.
- The notebook's **self-test asserts the exact CLAUDE.md §4.2–§4.2b coordinates** (the same set as
  `build/verify-web-core.mjs`, including the two Wisconsin regression points) and was confirmed
  passing **11/11 against the live services**.
