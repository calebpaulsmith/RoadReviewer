# Pre-built HPMS class tiles (web/tiles/*.pmtiles)

The web tool's road-class map layer is served from **pre-built vector
tiles** generated from FHWA's HPMS full-extent data (every public road,
FHWA class 1-7), one PMTiles file per state in `web/tiles/`. Tiles are
plain static files — GitHub Pages serves them as-is (PMTiles is read via
HTTP range requests, which Pages' CDN supports) — so the class map draws
instantly at any zoom with **zero live queries** and no record caps.

Division of labor (decided 2026-09-14, per user direction):

- **Display** (the map layer) = baked HPMS tiles. Progressive by class,
  baked into the tileset itself: interstates from z6, other principal
  arterials z8, minor arterials z9, major collectors z10, minor
  collectors z11, locals z12 (see `MINZOOM_BY_CLASS` in
  `fetch-hpms-state.mjs`). Rationale: high-class assignments are stable
  year to year, and even for lower classes the *display* being one HPMS
  submission old is acceptable because…
- **Verdicts** (federal-aid classification of pasted points) stay **live**
  against each state DOT's authoritative layer, exactly as before —
  that's where currency matters (collector/local reclassifications move
  the federal-aid line) and it's only a few point queries per site.
- States with no tileset yet (or if the tile fetch fails) fall back to
  the live progressive class mirror automatically.

## Source

`HPMS_National_Current` on the USDOT/BTS AGOL org (the org that hosts
the NTAD ACUB layer):
`https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/HPMS_National_Current/FeatureServer/0`
— confirmed live 2026-09-14: full extent incl. locals, `STATE_ID` =
state FIPS, `F_SYSTEM` = FHWA 1-7, pagination + `f=geojson`. Per-state
segment counts: WI 902,611 · IN 530,981 · MI 354,291 · MN 509,570 ·
IL 445,757 · OH 487,411 (≈3.23M total). The layer name says which HPMS
submission year it carries — record it in `web/sources.html` when
refreshing.

## Build (any machine with node + tippecanoe)

```sh
# per state: FIPS codes WI 55 · IN 18 · MI 26 · MN 27 · IL 17 · OH 39
bash build/tiles/build-state-tiles.sh 26 mi     # -> web/tiles/mi.pmtiles
```

The script pages the state out of the HPMS service (~2-4 min per 100k
segments) and runs tippecanoe (`-l roads -Z6 -z13`, per-feature minzoom
honored from the `tippecanoe` key, `--drop-densest-as-needed` so dense
cores thin instead of overflowing). `tippecanoe` ≥ 2.x: `apt-get install
tippecanoe` or https://github.com/felt/tippecanoe.

## Refresh cadence

HPMS publishes annually. When BTS updates `HPMS_National_Current`,
re-run the six builds and commit the new `.pmtiles` (and update the year
noted in `web/sources.html` + the web legend's `HPMS_TILE_YEAR`).
Verdicts are unaffected — they never read these tiles.

## Size constraints

GitHub's hard limit is 100 MB per file. Per-state tiles measured at
build time are recorded here; if a state ever crosses ~95 MB, split it
(e.g. classes 1-5 / 6-7 as two PMTiles) or lower `-z`.

Measured (2026-09-14, HPMS_National_Current):

| state | features harvested | pmtiles (with state-linkage attrs) |
|---|---|---|
| MI | 349,233 | 41 MB |
| IN | 529,343 | 50 MB |
| OH | 487,331 | 57 MB |
| IL | 445,757 | 46 MB |
| MN | 509,570 | 53 MB |
| WI | 902,580 | 70 MB |

The state-linkage attributes (round 2, below) roughly double-to-triple
each tileset — every file still under the 100 MB limit; ≈ 317 MB total
across the six states (+ 50 MB basemap + 3 MB ACUB).

## State-linkage attributes in the class tiles (2026-09-14, round 2)

Per user direction ("connect the segments to state segments"), the
state tilesets are rebuilt carrying each segment's **state LRS keys** —
HPMS is built from the states' own submissions, so these point straight
back at the state inventory:

| tile key | HPMS field | meaning |
|---|---|---|
| `F` | F_SYSTEM | FHWA class 1-7 (as before) |
| `R` | ROUTE_ID | the STATE's LRS route id (MI: the MDOT PR number — live-verified `0006904` at the Kalamazoo test point; IN: INDOT LRS id; IL: IDOT key-route; OH: ODOT NLFID-style; MN/WI: their LRS ids) |
| `B`/`E` | BEGIN/END_POINT | state mileposts (3 decimals) |
| `N` | RouteName | street/route name — offline names, incl. states whose own layers publish none (MN, IL) |
| `RN` | RouteNumber | signed route number where present |

Null/empty attrs are omitted per feature. The cached classifier reads
them into segments (routeId/mpFrom/mpTo/name), the row detail shows a
"State route <id> · MP <a–b>" chip on the closest segment, and the
CSV / GeoJSON / KMZ exports carry State Route ID + Milepost Range.

**Service gotcha (2026-09-14 evening):** the HPMS service stopped
accepting `resultRecordCount` (every query with it → 400 "Invalid query
parameters"; a republish — the layer now names itself
HPMS_National_2024_FullJoin). The harvester dropped the parameter; the
layer's own maxRecordCount (2000) still caps pages and
`exceededTransferLimit` still signals the quadtree split.

## Cached-tile CLASSIFICATION (2026-09-14, per user direction)

Verdicts now run against the hosted tilesets by default — "couldn't we
just check every single point with the data we have cached?" — through
the SAME computeVerdict logic as the live path:

- **Road segments + FHWA class** from the per-state HPMS tilesets
  (z13 tiles read in-browser via `protomapsL.PmtilesSource`, geometry
  converted back to lon/lat, true point-to-segment distances).
- **Urban/rural + urban-area name** from **`web/tiles/acub.pmtiles`**
  (3 MB, 549 Region V polygons from the NTAD 2020 Adjusted Urban Areas
  service, built by `build-acub.sh` → z12): point-in-polygon, plus a
  ring-distance check for the "Urban boundary edge" yellow rule.
  Tile-clip edges can't fake a boundary hit — tippecanoe's tile buffer
  keeps them farther from any in-tile point than the 76 m rule floor.
- Census TIGER **street names** stay a live, non-fatal backfill
  (verdict-irrelevant); offline they're simply blank.
- Rows carry a "cached data" chip; **"Live verdicts"** under Data
  service URLs switches back to per-point state-DOT + NTAD queries
  (slower, but reflects reclassifications newer than the tilesets).
- Falls back to the live path automatically on file://, a missing
  tileset, a read error, or an out-of-region point.

NTAD harvest gotchas (fetch-acub.mjs): resultOffset pagination hits the
same ~55 s server give-up as the HPMS table, and full-precision
polygons 504 — hence ids-then-objectId-batches with
`maxAllowableOffset≈3 m`.

## Offline road basemap (web/tiles/basemap.pmtiles)

The map's ROAD BASEMAP is also served from this repo — a Protomaps/OSM
vector extract of the six states (z0-11, 50 MB), built by
`build-basemap.sh` + `filter-basemap-layers.py`: `pmtiles extract` with a
six-rectangle region polygon, then a wire-level layer filter keeping only
earth / water / roads / boundaries / places (buildings, POIs and landuse
are most of a full basemap's bulk — the unfiltered z12 extract measured
504 MB). Rendered by the already-vendored protomaps-leaflet `light`
theme; tiles beyond z11 overzoom (the HPMS class tiles carry every road
from z12, so the basemap's job above that is context and names).
Satellite imagery deliberately stays a LIVE Esri layer — fetched only
when the user switches to it. If `basemap.pmtiles` is missing (or on a
`file://` open) the page falls back to the live Esri street layer.
Refresh occasionally from a newer Protomaps daily build (OSM edits).

- **`pmtiles extract` does not checksum tiles** — one 504 MB pull through
  the agent proxy arrived with ~2,700 truncated/zeroed tiles (silently!).
  `filter-basemap-layers.py` gunzips every tile and aborts on the first
  bad one; if it aborts, re-extract.
- **tippecanoe 2.49's `tile-join` cannot filter these tiles** — its MVT
  reader errors ("PBF decoding error") on some Protomaps tiles even from
  a clean archive, hence the wire-level python filter (which never
  decodes features at all).
- **Always pass `maxDataZoom` (= the archive's maxzoom) to
  `protomapsL.leafletLayer`** — the library defaults it to 15 and never
  reads the PMTiles header, so past the archive's real maxzoom it
  fetches nonexistent tiles and renders BLANK instead of overzooming
  (the HPMS class lines silently vanished at z14+; Leaflet's stale
  lower-zoom canvases masked it until the canvases were cleared).
  Class tilesets pass 13, the basemap 11, the live OpenFreeMap detail
  layer 14.
- **Serve `.css` with `text/css` in any local test server** — Chromium
  silently rejects a stylesheet served as octet-stream; with leaflet.css
  rejected every Leaflet pane loses `position:absolute` and the map
  renders as misplaced patches with invisible pins (cost a long
  debugging round in the Playwright harness; the PAGE was never broken).

## Build gotchas (cost a debugging round each)

- **Never use per-feature `"tippecanoe":{"minzoom":…}` for the class
  bands.** tippecanoe (2.49) silently RATE-DROPS line features carrying
  that key even at maxzoom — Michigan collapsed from 349k features to
  exactly one line per tile (`dropped_by_rate` in the tile `strategies`
  metadata; `-r1` does not prevent it; repro: 5 clean parallel lines →
  1 survivor). The `-j` `$zoom` feature-filter in `build-state-tiles.sh`
  implements the same banding correctly.
- **Whole-state envelopes and any attribute-filtered offset page 400
  after ~55 s** on the ~30M-row national table (server give-up, blank
  message). The harvester treats that error as a split signal and seeds
  a 0.5° grid — ~1° dense-metro cells burn the give-up before splitting;
  ~0.2-0.5° cells answer in seconds.
- `tippecanoe-decode`'s per-tile `lines` count is MVT features, and
  features split at tile borders — z13 counts exceed the input count;
  judge completeness by decoding a known urban tile (Kalamazoo
  13/2148/3032 held 762 features), not by totals alone.

## Basemap renderer / schema (2026-09-14)

`web/tiles/basemap.pmtiles` is extracted from the daily Protomaps build,
which is the **v4 tile schema** (`kind`, `kind_detail`, water lines in the
`water` layer). The vendored renderer must read that schema:
`web/vendor/protomaps-leaflet/protomaps-leaflet.js` is **5.1.0**
(`flavor:"light", lang:"en"`). A 3.x renderer (v3 `pmap:kind` themes)
draws only earth + water — blank land and river lines filled as polygon
slivers — which is exactly what shipped first. If the renderer is ever
downgraded or the build ever changes schema, re-check one tile's property
keys (`build/tiles`'s sibling probe: decode a z11 tile and list layer
keys) before committing.

The class tilesets start at z6 (`-Z6`); the page renders them with
`levelDiff: 0` (display z from data z) so the region view at map minZoom
6 has data — the renderer default (levelDiff 1) asked for z5 tiles and
drew nothing at the region zoom.

