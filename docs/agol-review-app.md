# AGOL site-review app — build guide & Excel hand-off

**Goal.** Inspector clicks one button in the Excel tool → an ArcGIS Online app
opens with the state's functional-class (NFC) + Adjusted Census Urban Boundary
(ACUB) layers already loaded → they step through each pasted site one at a time
and confirm (eyes-on) whether it's an FHWA/federal-aid road. The same pattern
extends to Floodplain review (swap in the NFHL flood layer) and Photo review
(swap in an imagery basemap). The auto-classifier becomes an *initial pass* — a
colored hypothesis per site — not the final word.

This is **Option A** from the design conversation: no authenticated AGOL push.
The button exports a colored KML and opens the app; the sites arrive by
drag-drop. Read [§0](#0-the-honest-constraint-read-first) before building —
it explains what AGOL can and can't do here without auth, and why a small
coded companion page is recommended alongside it.

---

## 0. The honest constraint (read first)

Two facts about AGOL bound how this can work with **no authentication**:

1. **URL fragments can't feed an AGOL app.** There's no supported way to encode
   the sites into the app URL and have Experience Builder / an Instant App parse
   them into a layer. (That "zero-click, pre-loaded" trick only exists in a page
   we write ourselves — see [§6](#6-optional-coded-companion-page-true-zero-click).)
2. **Widgets differ on runtime data.** The **List** widget binds to a data
   source chosen at *design* time, so it will **not** auto-fill from a file the
   user adds at runtime. BUT the **Feature Info** widget, set to *interact with a
   Map widget*, exposes **all map layers** — including a layer the user adds at
   runtime via the **Add Data** widget — and it has a built-in **Feature
   navigation** stepper (Next/Previous, "1 of 25" index) plus an attribute panel.
   So "user adds the file → clicks Next through every site → sees the verdict" **is
   achievable natively, no auth** (see [§3B](#3b-experience-builder--add-data--feature-info-recommended-native-step-through)).
   The *only* thing that still needs auth is the *fully automatic* "click Excel
   button → zero actions at all → sites already loaded," which would require an
   authenticated overwrite of a hosted layer each run (V2, `CLAUDE.md` §3.3).

**What this means in practice:**

| You want… | No-auth AGOL gives you… | How |
|---|---|---|
| Authoritative NFC/ACUB layers, your org's symbology, team-shareable | ✅ A saved Web Map + Instant App/Experience | this guide, §2 |
| Add today's sites, colored by verdict, click pins/table to review | ✅ Drag-drop into Map Viewer, or Add Data widget in an Experience | §4 |
| **User adds the file → Prev/Next through each site → verdict shown inline** | ✅ Experience: Add Data widget + Feature Info (map-interactive) navigation | §3B |
| **Click button → *zero* further actions → sites already loaded & stepping** | ⚠️ Needs auth (hosted-layer overwrite) — or use the coded page | §4C / §6 |

**Export GeoJSON or CSV, not just KML, for the AGOL Add Data path.** KML added to
AGOL becomes a "KML layer" that is often not fully queryable/selectable by widgets;
GeoJSON and CSV become proper client-side feature layers the Feature Info
navigation consumes cleanly. The Add Data widget accepts shapefile, CSV, KML,
GeoJSON, GPX, and FGDB.

The recommended architecture uses **both**: AGOL as the authoritative reference
map, the coded page as the fast zero-click review driver. They share the same
KML/verdict data and the same public layer URLs.

---

## 1. Layers you'll add (all 6 states, public, no auth)

Every one of these is a public REST service already used by the Excel tool, so
"is it addable to my map / usable in an Experience?" is **yes** for all of them —
you add them as *reference layers by URL*; the data does not have to be copied
into your org, and reference layers are fully usable in Web Maps, Experiences and
Instant Apps.

| State | Layer (what it shows) | Add-by-URL |
|---|---|---|
| **MI** | MDOT NFC — `FunctionalSystem` 0–7 | `https://mdotgis.state.mi.us/arcgis/rest/services/DataAccess/NfcNhsPub/MapServer/353` |
| **IN** | INDOT `LRSE_Functional_Class` — `functional_class` 1–7 | `https://gisdata.in.gov/server/rest/services/Hosted/LRSE_Functional_Class/FeatureServer/22` |
| **WI** | WisDOT State Trunk — `FED_FC_CD` 1–7 | `https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/FFCL_gdb/FeatureServer/3` |
| **WI** | WisDOT Local Roads — `FNCT_CLS_CTGY_TYCD` (urban/rural in code) | `https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/Functional_Class_Local_Non_Prod/FeatureServer/1` |
| **MN** | MnDOT common layers — `FUNCTIONAL_CLASS` | `https://dotapp9.dot.state.mn.us/egis12/rest/services/BASEMAP/mndot_commonlayers2/MapServer/11` |
| **IL** | IDOT Functional Class — `FC` | `https://gis1.dot.illinois.gov/arcgis/rest/services/AdministrativeData/FunctionalClass/MapServer/0` |
| **OH** | ODOT Functional Class — `FUNCTION_CLASS_CD` | `https://tims.dot.state.oh.us/ags/rest/services/Roadway_Information/Functional_Class/MapServer/0` |
| **All** | ACUB (Adjusted Census Urban Boundary), nationwide | `https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0` |

Notes:

- **MI needs no browser-UA workaround here.** The HTTP 403 the Excel tool works
  around (`BROWSER_UA`) is specific to VBA's `MSXML2.ServerXMLHTTP` default
  User-Agent. A real browser (which is what Map Viewer / an Experience runs in)
  sends a normal UA, so MDOT answers fine. It's also CORS-enabled.
- **MN/IL/OH are reference-only in the Excel tool** and haven't been CORS-verified
  from a browser. Add each with "Add layer from URL" and confirm it draws before
  relying on it; if one refuses (CORS), you can still add it in Map Viewer via the
  item's *hosted* copy or a WMS alternative. All three expose a bare FHWA 1–7 class
  field, so their pop-ups configure the same way as IN.
- **Floodplain variant:** add FEMA's National Flood Hazard Layer (NFHL) —
  `https://hazards.fema.gov/arcgis/rest/services/public/NFHL/MapServer` — instead
  of (or alongside) the NFC layers.
- **Photo variant:** switch the basemap to Esri **World Imagery**, no operational
  layer needed.

---

## 2. Build the Web Map (one-time, ~15 min)

1. Sign in to ArcGIS Online → **Map** (opens Map Viewer).
2. **Add** → **Add layer from URL** → paste each URL from §1 you want. Start with
   the states you actually cover (MI/IN/WI) + the ACUB layer; add MN/IL/OH later.
3. **Basemap** → choose **Streets** (roads + labels for orientation). The Photo
   variant uses **Imagery**.
4. Style the **ACUB** polygon layer: fill ~10 % opacity, a distinct outline
   (e.g. purple), so urban areas read as a wash under the roads. Put it *below*
   the NFC line layers in the layer list.
5. Leave the NFC line layers with **their own published symbology** (each state's
   REST service ships a renderer that colors by class — that's the authoritative
   coloring; don't reinvent it). Exception: INDOT publishes a single-symbol
   renderer (every class one color), so in Map Viewer restyle IN's layer with a
   *Types (unique symbols)* renderer on `functional_class` using the standard FHWA
   palette if you want per-class colors (optional).
6. Set **visible range** on the NFC layers so they only draw when zoomed in
   (e.g. Streets zoom ~12+) — keeps the statewide view uncluttered.
7. **Save** → title `RoadReviewer — FHWA Review (MI/IN/WI)`, tag `roadreviewer`,
   share to your org (or a group your reviewers are in). Copy the item's URL.

You now have a reusable authoritative reference map. This alone is valuable —
it's the "actual eyes on the prize" layer stack leadership wants.

---

## 3. Turn the Web Map into a step-through app

Pick the template that matches how polished you want it. The reviewer loads today's
sites at runtime (§4) — via drag-drop into Map Viewer or the **Add Data** widget in
an Experience. The **Experience Builder Add Data + Feature Info** path (§3B) is the
recommended one: it gives a real Next/Previous stepper over the added sites with no
auth. Only the *fully automatic, zero-action* variant needs auth (§4C).

### 3A. Instant App — **Media Map** or **Attachment Viewer** (fastest to stand up)

- From the Web Map item page → **Create Web App** → **Instant Apps**.
- **Media Map**: gives a clean map + a pop-up-driven experience; enable the
  **"Featured content"/bookmark** stepping if you pre-save bookmarks. Best when
  the site layer is known ahead of time.
- **Attachment Viewer**: steps through the features of one layer **one at a time**
  with a large side panel showing that feature's attributes (and photos, if the
  layer has attachments) — this is the closest Instant App to "tab through each
  site and see what it is." It requires the site layer to be a **hosted layer in
  the map at design time**, which (per §0) an ad-hoc dragged KML isn't. Use this
  only if you decide to publish the KML as a hosted layer each run (§4C).
- Configure the layer's **pop-up** (§5) so the verdict shows big and colored.
- Publish; copy the app URL.

### 3B. Experience Builder — **Add Data + Feature Info (recommended native step-through)**

This is the path that delivers "user adds the file → clicks Next through each
site → sees the verdict," with **no auth** and **no per-run rebuild**, because
Feature Info reads the runtime-added layer instead of a design-time binding.

1. Web Map item → **Create Web App** → **Experience Builder** → start from the
   **"Pocket"** or **"Foldable"** template (map + side panel).
2. Add a **Map** widget bound to your §2 Web Map (the NFC + ACUB reference stack).
3. Add an **Add Data** widget. Configure it to accept file upload; this is how the
   reviewer loads today's sites (GeoJSON/CSV preferred over KML — see §0).
4. Add a **Feature Info** widget and set its data mode to **"Interact with a Map
   widget"** (pointed at the Map from step 2). This makes it list **all map
   layers** — including the layer added at runtime. Enable **Feature navigation**
   and **Show index** so you get **◀ / ▶** and a **"1 of 25"** counter. Configure
   its display to show the verdict (big/colored — §5), FHWA class, urban/rural,
   ACUB name, road name, lat/lon.
5. *(Optional polish)* Add a **Data added** trigger on the Add Data widget → target
   the Map widget → *zoom to the added layer*, so the map frames the sites the
   moment they're loaded.
6. Style: header with the "candidate for human review, not a determination"
   disclaimer, your org's colors, a legend. Publish; copy the app URL.

Reviewer flow: open app → **Add Data** → pick the file the Excel button exported →
map frames the sites, colored by verdict → **Feature Info ▶** steps through each
one showing the tool's finding. That's the whole eyes-on loop.

**One thing to verify hands-on (5 min):** confirm the Feature Info widget in
map-interactive mode actually lists the *runtime-added* layer (Esri's docs confirm
it lists "all map layers" and that Add Data layers become map layers, but don't
spell out the runtime case in a single sentence). If a KML layer isn't pickable,
switch the export to GeoJSON/CSV.

### 3B-alt. Experience Builder — **List + Map** (only with a hosted sites layer)

If you go the hosted-layer route (§4C), a **List** widget bound to that layer gives
the fanciest cards (bold name + conditional-formatted verdict chip + click-to-pan).
List binds at design time, so this needs the sites as a real hosted layer, not an
ad-hoc add. Use it for a curated/standing review, not per-run ad-hoc batches.

### 3C. Plain **Map Viewer** (no app at all — the true no-auth path)

Honestly the most robust no-auth option: **skip the app**, and have the button
open the *saved Web Map* directly in Map Viewer. The reviewer drags the KML in
(§4A), then uses Map Viewer's built-in **feature list / table** to click each
site; pop-ups (§5) show the verdict. No "Next" button, but zero binding problems
and nothing to rebuild per run. This is what the Excel button targets by default
(§4A) until/unless you commit to the hosted-layer path (§4C).

---

## 4. Getting today's sites into the app (the hand-off)

### 4A. Drag-drop KML (default, no auth) — what the Excel button does

The Excel tool already writes a **verdict-colored KML** (red = federal aid,
green = non-federal, blue = review) grouped by Category, and can open a target
URL + pop Explorer with the KML highlighted for a one-drag drop. See
`modMaps.OpenSitesOnNfcLayer` / `SendSitesToAgolMap`. The new button (§7) opens
**your saved Web Map / app URL** instead of the raw layer.

**GeoJSON is now written too.** As of the GeoJSON export change,
`OpenSitesOnNfcLayer` drops a `… Sites.geojson` next to the KML (same output
folder), and there is a standalone `modMaps.ExportSitesToGeoJson` sub. Each
feature carries flat properties — `Name`, `FedAidStatus`, `Verdict` (bucket),
`VerdictColor` (hex), `FHWAClass`, `UrbanRural`, `ACUBName`, `RoadName`,
`StreetName`, `ReviewNote`, `Category`, `Description`, `Latitude`, `Longitude`.
**Use the `.geojson` (not the `.kml`) in the Experience's Add Data widget** so
Feature Info's step-through and verdict-field symbology work (§0). The `.kml`
stays for Google Earth and plain Map Viewer drag-drop.

Flow the inspector sees:
1. Paste sites in Excel, click **Check Roads** (auto-classify, initial pass).
2. Click **Open Sites in AGOL Review App**.
3. Browser opens your Web Map/app; Explorer opens with the KML highlighted.
4. Drag the KML onto the Map Viewer window → sites appear, colored by verdict.
5. Click each pin / table row to review; the pop-up shows the tool's finding.

AGOL Map Viewer ingests KML as a temporary layer with working pop-ups. This is
the whole Option-A loop, no accounts-per-site, no upload of damage data (the KML
carries only name + coords + verdict).

### 4B. Fragment fallback — **coded page only**

If you also stand up the coded companion page (§6), the same button can pass the
sites in the URL `#fragment` (zero drag). AGOL apps can't read that; the coded
page can. That's the only place "zero-click" is real.

### 4C. Hosted-layer path (needed for 3A Attachment Viewer / 3B List binding)

If you want the polished bound step-through (Attachment Viewer / EB List), the
sites must be a hosted feature layer:
- **Manual, no auth each run:** Content → **New item** → *Your device* → the KML
  → **Publish** as a hosted feature layer. The app must be re-pointed at the new
  item (its ID changes each publish) — clunky, only worth it for a one-off review.
- **Automated (V2, needs auth):** an ArcGIS API token stored in Setup + VBA calls
  to the layer's `applyEdits`/`truncate`+`addFeatures` to overwrite a *fixed*
  hosted layer each run. Then the app never needs re-pointing and the whole thing
  becomes truly one-click. This is the real "productionize it" upgrade; it's out
  of V1 no-auth scope but is the natural next step if the workflow proves out.

---

## 5. Show "what the tool found" — pop-up / card config (the extra you asked for)

Whatever template, the payoff is a pop-up that surfaces the classifier's verdict
big and colored. The KML placemarks carry the verdict in their name/description;
if you go the hosted-layer route (§4C) you'll have real fields. An **Arcade**
pop-up expression gives you a colored badge from the verdict text:

```arcade
// Pop-up "Arcade" expression: colored verdict badge
var v = $feature.FederalAidStatus;      // or parse from the KML description
var color =
  Left(v, 10) == "Federal ai" ? "#c0392b" :   // red
  Left(v, 6)  == "Review"     ? "#f39c12" :    // amber
                                "#27ae60";      // green
return `<div style="font-size:15px;font-weight:700;color:#fff;
  background:${color};padding:6px 10px;border-radius:6px;display:inline-block">
  ${v}</div>`;
```

Then in the pop-up body list the FHWA class, urban/rural, ACUB name, road name,
and the site's lat/lon. In an EB List card, drive the card's background or a chip
element off the same verdict field with conditional formatting. Add the front-page
disclaimer ("candidate for human review, not a determination") to the app header
so the eyes-on framing is explicit.

---

## 6. Optional coded companion page (true zero-click)

Because §0 blocks the zero-click experience in pure AGOL, a small static page —
a sibling of `web/index.html`, deployed on the **GitHub Pages** the repo already
publishes — closes the gap and needs no AGOL account at all:

- **Input:** reads the sites from the URL `#fragment` the Excel button writes
  (name, lat/lon, verdict, class, road) — nothing hits a server, so it stays
  private — with a drag-drop KML zone as fallback for oversized batches.
- **Map:** the same public NFC + ACUB layers from §1, side-loaded into the Leaflet
  map already built in `web/index.html` (`fetchClassLayers` / `fetchAcubLayer`).
- **Review:** reuses the existing Prev/Next stepping, verdict-colored pins, and
  per-site panel (`selectSite`, the review legend) — but instead of re-classifying,
  it *displays the verdict Excel already computed*.
- **Result:** click Excel button → browser opens the page fully populated → hit
  **Next** to walk every site, verdict shown inline. Exactly the stated goal.

This is a few hours of work reusing existing code and is the recommended way to
get the polished zero-click loop. The AGOL Web Map (§2) still stands as the
authoritative, team-shareable reference the coded page links out to per site.

---

## 7. Excel wiring (once your app URL exists)

Small, self-consistent addition mirroring the existing `JobAgolMap` plumbing:

1. **`modConstants.bas`** — add the app URL as a constant *or* (better) reuse an
   input cell so it's editable without a rebuild:
   - Add a named range `NR_REVIEWAPP` (e.g. `JobReviewApp`) and a labeled input
     cell on **Start Here** (inspector) / **Sites** toolbar (standard), exactly
     like `JobAgolMap` (`modBuild` writes the cell; `SetupValue(NR_REVIEWAPP)`
     reads it).
2. **`modMaps.bas`** — add `Public Sub OpenSitesInReviewApp()`, a near-clone of
   `OpenSitesOnNfcLayer` (§ modMaps.bas:616): call `WriteSitesKml`, then open
   `SetupValue(NR_REVIEWAPP)` (fall back to the §2 Web Map URL if blank) instead
   of `NfcLayerUrlForFirstSite()`, then `Shell "explorer.exe /select,"` the KML.
   If you also build §6's page, append the fragment: read each row's coords +
   verdict, build `#s=<base64 json>`, and append to the URL before `FollowHyperlink`.
3. **`modBuild.bas`** — add the button to the toolbar (`WriteSitesToolbar` for
   standard, Start Here for inspector), label **"Open Sites in AGOL Review App"**,
   `OnAction = "OpenSitesInReviewApp"`.
4. **Rebuild + verify** (the cloud env has no Excel — do this on the Windows
   laptop, `CLAUDE.md` §9.1/§9.2):
   ```powershell
   & "…\RoadReviewer\build\build.ps1"
   & "…\RoadReviewer\build\verify-skeleton.ps1" -XlsmPath "…\Site Inspector Review Tool.xlsm"
   ```
   Add the new button's `OnAction` + named range to `verify-skeleton.ps1`'s
   asserted surface so a missing wiring fails the build.
5. Bump `BUILD_REFERENCE` in `modConstants.bas`.

---

## 8. Extending to Floodplain & Photo review

Same Web Map + app + button pattern, different reference layer:

- **Floodplain:** a second Web Map with the FEMA **NFHL** MapServer (§1) as the
  operational layer; the verdict chip becomes flood-zone (A/AE/X…) once you add a
  zone lookup (a point-in-polygon against NFHL layer 28, "Flood Hazard Zones").
  The KML/site hand-off is identical.
- **Photo:** a Web Map on the **World Imagery** basemap; the app steps through
  sites over aerials for pre-disaster condition. Pair with the existing "Photo
  Links" button (opens Google/Bing/Street View/Earth) for source diversity.

One `JobReviewApp`-style cell per review type (or a small dropdown) lets one button
target the right app.

---

## 9. Quick recommendation

1. **Build the §2 Web Map** (MI/IN/WI + ACUB) — high value on its own, ~15 min.
2. **Build the §3B Experience** (Add Data + Feature Info navigation) — this is the
   native AGOL app where the reviewer adds today's file and clicks Next through
   each site with the verdict shown. No auth, no per-run rebuild. Do the 5-minute
   runtime-layer check in §3B before committing.
3. **Add GeoJSON/CSV to the Excel export** (§0) so the Add Data layer is fully
   queryable, and **wire the button** (§7) to open the app + export the file.
4. **Add §6's coded page** only if you want the *fully automatic zero-action* loop
   (button → sites already loaded, no Add-Data step) without paying for auth.
5. **Consider §4C automated hosted-layer push (V2/auth)** to make the AGOL app
   itself fully one-click, once the workflow is proven and leadership wants it
   productionized inside AGOL.
