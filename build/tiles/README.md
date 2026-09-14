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

| state | features harvested | pmtiles |
|---|---|---|
| MI | 349,233 | 19 MB |
| IN | 529,343 | 6.4 MB |
| OH | 487,331 | 26 MB |
| IL | 445,757 | 22 MB |
| MN | 509,570 | 22 MB |
| WI | 902,580 | 36 MB |

All six fit comfortably under the 100 MB/file limit; total ≈ 132 MB
across files. (IN is an outlier at 6.4 MB despite its feature count —
INDOT submits shorter, simpler segment geometry.)

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
