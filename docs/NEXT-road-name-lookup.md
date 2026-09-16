# NEXT: bare road names, county identifiers, and the place search

**Status: planned, not built. Nothing in this file has shipped.**
Written 2026-09-16 to hand off to a desktop session. Everything below was
measured live on that date; the commands to re-measure are in §7.

The shipped behaviour this builds on is PR #44 (merged, deployed):
addresses and road names in the coordinates box, geocoded on one explicit
button, snapped to the named street. See CLAUDE.md §7b "Addresses and road
names in the same box".

---

## 1. The problem we are solving

`web/index.html` accepts a **road line** — a street plus a place — and
classifies the whole road:

```
Portage Rd, Portage MI          ✅ works today
CR 550 N, Hamilton County IN    ✅ works today
Q Ave                           ❌ "Could not read a lat/lon, address or road"
```

A road name with **no place** is refused. That is deliberate (a bare name
would mean a six-state search) but it is a dead end: people will type bare
street names, and the message does not say what to add.

Two things need deciding, and they are related:

1. **How does a bare road name get a place?** (§4)
2. **Where does the place come from — can the road segments tell us?** (§3)

---

## 2. Why "just search everywhere and let them pick" does not work

Measured against Census TIGERweb, exact-name search, clustering the returned
segments into distinct places:

| scope | `Main St` | `Q Ave` | `Portage Rd` | query time |
|---|---|---|---|---|
| Six-state envelope | **904 places** | 29 | 20 | 4.4 s |
| Kalamazoo County, MI | **2** | 4 | 1 | 0.3–0.6 s |
| Wayne County, MI (Detroit) | 11 | 0 | 0 | 0.5 s |
| Cook County, IL | 35 | 0 | 0 | 0.8 s |

Three conclusions:

- A region-wide picker is useless — 904 entries for the name people most
  commonly type.
- **County scope is ~10× faster** (0.3–0.8 s vs 4.4 s), cheap enough to run
  automatically rather than behind a button.
- At county scope the list is short and made of **real named places**
  ("Main St in Brady township vs Kalamazoo charter township") — a question
  an inspector can answer. `Q Ave` and `Portage Rd` return **0** in
  Wayne/Cook, so the picker disappears when the name is not there.

Also: the six-state *rectangle* covers chunks of Iowa, Missouri and Kentucky
(the `Q Ave` results included Iowa counties). The region is a mask, not a
box — one more reason not to run unscoped searches.

---

## 3. What the road data actually contains

### 3.1 HPMS carries county, urban area and ownership

`HPMS_National_Current/FeatureServer/0` — the layer `build/tiles/fetch-hpms-state.mjs`
already harvests. We currently pull only
`OBJECTID,F_SYSTEM,ROUTE_ID,BEGIN_POINT,END_POINT,RouteName,RouteNumber`.

| field | what it gives | coverage / verification |
|---|---|---|
| `COUNTY_ID` | 3-digit county FIPS (pair with `STATE_ID`) | **99.6%** populated in MI (1,477 null of 354,291). Verified `77` at the §4.2 Kalamazoo test point = FIPS 26077 ✓; Lac du Flambeau bbox returned 51/125 = Iron/Vilas ✓. **No name domain** — needs our own FIPS→name table (§3.3) |
| `URBAN_ID` | urban area, **with a 514-entry coded-value name domain** | `43723 → "Kalamazoo; MI"` — the exact UACE our ACUB layer returns; `47719 → "Lansing; MI"`; **`99999 → "Rural"`**. ~98.6% populated in MI |
| `OWNERSHIP` | **named domain**: `1` State Hwy Agency, `2` County Hwy Agency, `3` Town or Township, `4` City or Municipal, `50` Indian Tribe Nation, `62` BIA, plus USFS / NPS / Corps / BLM etc. | dense — Lac du Flambeau bbox: 268 township, 66 county, 26 state |
| `RouteName` | street name | already baked as tile property `N` |

`OWNERSHIP` is the sleeper find: it is effectively **who owns the road**,
which in PA terms is usually **who the applicant is**. Nothing surfaces it today.

### 3.2 City and township are NOT on the segments

| state | county | city / township |
|---|---|---|
| **IN** | `county_l`/`county_r`, `county_fips` | ✅ `incmuni_l` / `incmuni_r` — real municipality names |
| **OH** | `COUNTY_CD` (e.g. `'FRA'`) | ❌ `JURISDICTION_CD` is `'M'`/`'S'` — a maintenance code, not a place |
| **WI** | `DOT_CNTY_CD` | ❌ nothing, in all 97 fields |
| **MN** | ❌ (7 fields total) | ❌ |
| **IL** | ❌ (21 fields total) | ❌ |
| **MI** | ❌ (MDOT 353 is LRS keys only) | ❌ |

Only Indiana has usable municipality names, so **city/township cannot come
from the segments**. That search stays with Census TIGERweb
(`Places_CouSub_ConCity_SubMCD` layer 1 = county subdivisions, layer 4 =
incorporated places) — one boundary query per search, not per segment, so
the cost is fine.

### 3.3 County names are a 13 KB static table

TIGERweb `State_County/MapServer/1` filtered to the six states returns
**524 counties**, which serialises to **13,013 bytes** of `{GEOID: NAME}`.
Spot-checked: `26077 → Kalamazoo County`, `55125 → Vilas County`,
`26051 → Gladwin County`. Bake it at tile-build time; it never changes.

### 3.4 Tribal: the ownership codes do not find reservations

`OWNERSHIP` has `50 = Indian Tribe Nation` and `62 = Bureau of Indian
Affairs`, but inside the Lac du Flambeau reservation bbox **zero of 360
segments** carried either — the roads there are town, county and state
owned. So ownership marks tribally-*owned* roads, not tribal *areas*.

For a tribal **search**, the source is TIGERweb's `AIANNHA` service:

```
0  Alaska Native Regional Corporations      6  Alaska Native Village Statistical Areas
1  Tribal Subdivisions                      7  Oklahoma Tribal Statistical Areas
2  Federal American Indian Reservations     8  State Designated Tribal Statistical Areas
3  Off-Reservation Trust Lands              9  Tribal Designated Statistical Areas
4  State American Indian Reservations      10  American Indian Joint-Use Areas
5  Hawaiian Home Lands
```

Layers 2, 3 and 4 are the relevant ones for Region V.

---

## 4. The design

### 4.1 Default Area box

A **"Default area"** input above the coordinates box, applied to every road
line that does not carry its own place. `Q Ave` then resolves exactly as
`Q Ave, Kalamazoo County MI` does today — one bounded query, fast.

- **Pre-fill it** from what is already on screen: if the batch has
  coordinates, use the county they fall in (they nearly always share one);
  otherwise offer the current map area. Editable always.
- Show it on the row (`whole road · Kalamazoo County, MI`) so the scope is
  never invisible.
- A line that names its own place keeps overriding it.

This matches how the work is scoped — a disaster + applicant is one county
or a handful — and mirrors the Excel product's per-job State/Disaster/
Applicant fields.

### 4.2 Disambiguation picker, gated on county-or-smaller

Per user direction, a picker **only when the selected area is a county or
smaller** — that is what keeps the list short and the query fast (§2):

- **Default area unset, or set to a whole state** → no search. Row reads:
  *"Road name without a place — set a Default area (county or smaller)."*
  Instant and honest; a state-wide bare `Main St` is hundreds of places.
- **County or smaller** → resolve automatically:
  - **1 place** → use it, no picker (`Portage Rd` in Kalamazoo County).
  - **2–12 places** → picker in the row: `Main St — 2 places: Brady
    township · Kalamazoo charter township`. Click resolves that row and the
    pick is remembered.
  - **more than 12** → show the first 12 plus *"35 places have a Main St in
    Cook County — pick one, or narrow the area to a city/township."*
    Never a wall, never a dead end.
- **Order candidates by distance from the batch's other sites** (or the map
  centre). In a real disaster the other points sit near the road in
  question, so the right township lands first and the picker becomes a
  one-click confirm.
- A pick is a **decision the user made**, so it shows on the row and in
  exports the same way the address snap does — not as something the tool
  knew.

### 4.3 Bake the three HPMS fields into the tiles

Add `COUNTY_ID`, `URBAN_ID`, `OWNERSHIP` to the harvester's `outFields` and
to the tile properties (three small ints per segment), plus the two static
name tables (13 KB counties, ~10 KB urban areas). That buys, all offline:

- a **County** column on rows and in every export — useful on its own,
  since PA work is organised by county and applicant;
- **free picker labels** — no reverse county lookup per cluster, which is
  what made the disambiguation step slow;
- **attribute-scoped county search** rather than envelope-scoped, so
  "roads named X in Kalamazoo County" is exact instead of a bounding box
  that leaks into Iowa;
- an **owner chip** ("County Highway Agency") pointing at the likely
  applicant;
- an offline urban-area name that independently agrees with the ACUB
  lookup — as a **cross-check only**. ACUB stays the single source of truth
  for urban/rural per CLAUDE.md §4.2; do not relitigate that.

---

## 5. Plan, in order

1. **Michigan-only trial first.** Add the three fields to
   `build/tiles/fetch-hpms-state.mjs` (`outFields` + the `props` block
   around line 150) and rebuild **only MI**. Measure the resulting
   `mi.pmtiles` against today's 40 MB. This de-risks §6's size question
   before spending the full six-state harvest.
2. Bake the county FIPS→name and urban-area name tables (static JSON in
   `web/tiles/` or inlined — decide once the sizes are known).
3. Wire the tile properties through `classifyPointTiles` into the row,
   the chips and the exports (County, Owner columns).
4. Default Area box + auto-fill (§4.1).
5. County-or-smaller gate + picker with distance ordering and the 12 cap
   (§4.2).
6. Better message for the unset/state-level case — **this one is worth
   doing first and separately**, since it removes today's dead end in one
   line and does not depend on any of the above.
7. Re-harvest the remaining five states; re-commit tiles.
8. Verifier legs: bare name with area set (1 place → silent), several →
   picker, over-cap message, state-level refusal, per-line place still
   overriding the default, and county/owner reaching the exports.

---

## 6. Risks and open questions

- **Tile size vs GitHub's 100 MB per-file limit.** `wi.pmtiles` is already
  **67 MB**; `oh` 55 MB, `mn` 51 MB, `in` 49 MB, `il` 44 MB, `mi` 40 MB.
  Three extra integer attributes per segment will grow them. If WI crosses
  100 MB the options are dropping a field, coarser `maxAllowableOffset`, or
  splitting the state. **Measure on the MI trial before committing.**
- **The harvest is the expensive part.** Adding fields is one line; re-running
  the six-state HPMS harvest is hours, and it is the harvest with the
  quadtree splits and the "Unable to perform query" 400s (CLAUDE.md §7b).
  Do it once and take all three fields rather than coming back for ownership
  later.
- **HPMS refuses attribute-only filters.** `STATE_ID=26 AND OWNERSHIP=50`
  returns HTTP 400 ("Unable to perform query") on the 30M-row table. Every
  probe must carry a spatial filter. This is why §3.4's tribal check used a
  reservation bbox rather than a statewide count.
- **`COUNTY_ID` has no coded-value domain** — unlike `URBAN_ID` and
  `OWNERSHIP`, which ship their own name tables. Hence §3.3.
- **1.4% of MI segments have a null `URBAN_ID`**, distinct from the explicit
  `99999 = Rural`. Treat null as unknown, not rural.
- **Unverified:** whether a *city/township*-scoped search (smaller than a
  county) stays as fast as the county-scoped one. Expected to be faster —
  smaller envelope — but not measured.

---

## 7. Re-measuring any of this

```bash
H="https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/HPMS_National_Current/FeatureServer/0"

# HPMS attributes at the Kalamazoo federal-aid test point (CLAUDE.md §4.2)
curl -s "$H/query?geometry=-85.57025,42.28536&geometryType=esriGeometryPoint&inSR=4326\
&distance=150&units=esriSRUnit_Foot&outFields=STATE_ID,COUNTY_ID,URBAN_ID,OWNERSHIP,RouteName\
&returnGeometry=false&f=json"
#   -> STATE_ID 26, COUNTY_ID 77, URBAN_ID 43723, OWNERSHIP 4, RouteName "S Pitcher St"

# The coded-value name tables (URBAN_ID: 514 entries; OWNERSHIP: 27)
curl -s "$H?f=pjson" | python3 -c "
import sys,json; d=json.load(sys.stdin)
for f in d['fields']:
    if f['name'] in ('URBAN_ID','OWNERSHIP') and f.get('domain'):
        print(f['name'], len(f['domain']['codedValues']))"

# County FIPS -> name table for the six states (524 rows, ~13 KB)
curl -s "https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/State_County/MapServer/1/query\
?where=STATE%20IN%20('17','18','26','27','39','55')&outFields=GEOID,NAME,STATE&returnGeometry=false&f=json"
```

The ambiguity measurements in §2 were produced by two throwaway scripts
(region-wide clustering, and county-scoped clustering with township naming)
against TIGERweb `Transportation/MapServer/8`. They are not committed —
§2's table is the result, and re-deriving it is only worth it if the
approach changes.
