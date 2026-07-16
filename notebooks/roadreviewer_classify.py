# %% [markdown]
# # RoadReviewer — FHWA Federal-Aid Road Classifier (ArcGIS Online Notebook / Web Tool)
#
# This notebook is a **faithful Python port of the `rr-core` classifier** shipped in
# `web/index.html` (which is itself the JS port of the Excel tool's `src/modClassify.bas`
# + `src/modConstants.bas`). It takes pasted coordinates, classifies each point against the
# authoritative state functional-class (NFC) layers + the nationwide 2020 Adjusted Census
# Urban Boundary (ACUB) layer, and returns verdict-colored points as a **Feature set** so an
# Experience Builder app can draw them on a map with no spreadsheet and no file upload.
#
# **It classifies the ROAD, never the project.** "Federal aid" means the road is on the
# federal-aid system (Urban Minor Collector or greater per FHWA functional classification
# inside a 2020 ACUB). A human still verifies every point.
#
# **How the color is decided (identical to the web tool):**
# 1. Every road within the search radius is found and the true distance to each is measured.
# 2. The **closest** road decides red vs green — it is the road the point is on.
# 3. **Yellow only downgrades green** (red never downgrades) when an ambiguity could flip a
#    green result to federal-aid — reasons: *Second road close* / *Nearby FHWA road* /
#    *Urban boundary edge*.
#
# **Publishing note:** parameters are isolated in the first code cell; the Feature set output
# is assigned at the very bottom for the "Insert output" snippet. Use the **Standard** runtime
# (version **≥ 8.0**) — this notebook uses `arcgis` + `requests` + the standard library only
# (**no `arcpy`**), so it runs on the cheaper Standard runtime. Full walkthrough:
# `docs/notebook-web-tool-implementation.md`.

# %% [markdown]
# ## 1. Web-tool parameters (isolated at the top)
#
# When you publish this notebook as a web tool, the **Parameters** pane turns these two
# variables into **String** inputs. Click *Insert as variables* and it (re)writes THIS cell.
# Leave the variable names exactly as they are — the rest of the notebook reads them.
#
# * `coords_text` — the pasted coordinates, **one point per line**. Each line is
#   `lat lon` or `lat, lon`, with an optional free-text **name** before or after
#   (e.g. `Culvert on Q Ave, 42.6911, -84.5360`). Commas, tabs or spaces all work.
# * `radius_ft` — search radius in feet when no road passes exactly under the point
#   (string; default `"250"`, clamped to 1–1000).

# %%
# ─── WEB-TOOL PARAMETERS ─────────────────────────────────────────────────────
# (The Parameters pane overwrites this cell with the caller's values at run time.)
coords_text = """42.28536, -85.57025
Culvert on Q Ave, 42.6911, -84.5360
Site 12, 44.2700, -83.5200"""

radius_ft = "250"
# ─────────────────────────────────────────────────────────────────────────────

# %% [markdown]
# ## 2. Imports & constants
#
# Only `requests` (HTTP) and the standard library are needed for classification; `arcgis`
# is imported later, solely to package the output Feature set. Every constant below is copied
# verbatim from `rr-core` in `web/index.html` / `src/modConstants.bas`.

# %%
import re
import time
import math
import json

import requests

# --- Authoritative REST endpoints -------------------------------------------
# Key names match the Excel Sources-sheet Svc_ table and the web tool's REST /
# RR_SERVICE_OVERRIDES, so a URL swap is the same mental model in every product.
REST = {
    # Michigan MDOT — NFC class (layer 353) + trunkline route name (layer 543)
    "MI_NFC":         "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353",
    "MI_ROUTE":       "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/543",
    # Nationwide 2020 Adjusted Census Urban Boundary (USDOT NTAD)
    "ACUB":           "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0",
    # Indiana INDOT — authoritative Roads_and_Highways functional-class layer
    # (the one behind INDOT's official Functional Class Map) + separate
    # road-name centerline layer
    "IN_NFC":         "https://gis.indot.in.gov/ro/rest/services/RAH_GIO_Collaboration/LRSE_Functional_Class/FeatureServer/22",
    "IN_ROADNAME":    "https://gisdata.in.gov/server/rest/services/Hosted/Road_Centerlines_of_Indiana_2021/FeatureServer/15",
    # Minnesota / Illinois / Ohio (wired PR #36) — bare FHWA 1-7 class, no filter needed
    "MN_NFC":         "https://dotapp9.dot.state.mn.us/egis12/rest/services/BASEMAP/mndot_commonlayers2/MapServer/11",
    "IL_NFC":         "https://gis1.dot.illinois.gov/arcgis/rest/services/AdministrativeData/FunctionalClass/MapServer/0",
    "OH_NFC":         "https://tims.dot.state.oh.us/ags/rest/services/Roadway_Information/Functional_Class/MapServer/0",
    # Wisconsin WisDOT — local-roads layer (queried FIRST) + state-trunk layer (fallback).
    # WisDOT moved/locked the old WI_Local_Roads_Flood_Damage_Assessment_Snapshot
    # (now token-gated); this is the live public local layer.
    "WI_LOCAL_ROADS": "https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/Functional_Class_Local_Non_Prod/FeatureServer/1",
    "WI_STATE_TRUNK": "https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/FFCL_gdb/FeatureServer/3",
    # Census TIGER road centerlines — street names for every state (best-effort)
    "TIGER_ROADS":    "https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/Transportation/MapServer/8",
}

# --- Service-URL overrides ("lego swapout") ---------------------------------
# Paste a replacement layer URL here to re-point any service without editing the
# query code — e.g. if WisDOT moves its local-roads layer again:
#   SERVICE_OVERRIDES["WI_LOCAL_ROADS"] = "https://services5.arcgis.com/.../FeatureServer/1"
# Keys match REST above (and the Excel/web override tables). Empty = use default.
SERVICE_OVERRIDES: dict = {}

def svc(key):
    """The overriding URL for `key` if one is set, else the built-in default."""
    return SERVICE_OVERRIDES.get(key) or REST[key]

# Search radius when the exact point-on-road intersect misses (UI-adjustable; the
# Excel tool's "Search buffer" cell, same 250 ft default).
DEFAULT_BUFFER_FEET = 250
# The urban-boundary (ACUB) check never narrows below this, even when the road buffer
# is tightened — a point a few feet on the rural side of a road that itself forms the
# urban boundary must still resolve Urban (modClassify.ACUB_MIN_BUFFER_FEET).
ACUB_MIN_BUFFER_FEET = 250
# Two roads whose distances differ by less than this are "close together": GPS alone
# can't tell which one the point is on, so a federal-aid second road there earns a
# yellow flag (modClassify.CLOSE_ROAD_FEET).
CLOSE_ROAD_FEET = 30

# Browser-like User-Agent. MDOT (`mdotgis.state.mi.us`) returns HTTP 403 to the default
# non-browser UA (CLAUDE.md §4.2 / §9.3); a normal UA fixes it. Sent on every request —
# harmless on the other services, exactly like modHttp.HttpGetText.
BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# Verdict bucket -> display label + hex color. These are the SAME buckets/labels/colors
# the web tool's GeoJSON export writes (rr-core `BUCKET_GEO_LABEL` / `BUCKET_COLOR`), so a
# point classified here matches the same point classified in the browser.
BUCKET_GEO_LABEL = {"fed": "Federal aid", "nonfed": "Non-federal aid",
                    "review": "Review", "failed": "Review"}
BUCKET_COLOR = {"fed": "#d73027", "nonfed": "#1a9850",
                "review": "#b58900", "failed": "#777777"}

# %% [markdown]
# ## 3. HTTP helper
#
# Direct REST access with `requests` (mirrors `rr-core`'s `httpGetJson`). Adds the browser UA
# and a small retry so a transient MDOT 503 doesn't sink a whole batch (the web tool caches
# per-row and offers a manual retry link; a web tool has no UI, so we retry in place). Retries
# do **not** change any verdict — only robustness.

# %%
_SESSION = requests.Session()
_SESSION.headers.update({"User-Agent": BROWSER_UA})

def http_get_json(url, params=None, retries=3, timeout=45):
    """GET `url` and parse JSON. Retries transient failures with linear backoff.
    Raises on a hard failure so the caller marks the row 'Failed - ...' (F12)."""
    last_err = None
    for attempt in range(retries):
        try:
            r = _SESSION.get(url, params=params, timeout=timeout)
            r.raise_for_status()
            return r.json()
        except Exception as e:            # network / 5xx / JSON decode
            last_err = e
            if attempt < retries - 1:
                time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(str(last_err))

# %% [markdown]
# ## 4. Class-label + domain helpers
#
# The FHWA functional-system code table and the WisDOT local-category unpacker, both copied
# verbatim from `rr-core` (`functionalSystemLabel` / `wisconsinLocalCategoryToFhwa`).

# %%
def functional_system_label(code):
    """Bare FHWA functional-class code (0-7) -> its label (rr-core functionalSystemLabel)."""
    return {
        0: "Non-Certified Roadway",
        1: "Interstate",
        2: "Other Freeway",
        3: "Other Principal Arterial",
        4: "Minor Arterial",
        5: "Major Collector",
        6: "Minor Collector",
        7: "Local",
    }.get(code, "Unknown ({})".format(code))

def wisconsin_local_category_to_fhwa(cat):
    """WisDOT local-roads FNCT_CLS_CTGY_TYCD -> bare FHWA 1-7 (rr-core
    wisconsinLocalCategoryToFhwa; 96 -> Major Collector is safe, see CLAUDE.md §4.2b)."""
    return {10: 3, 60: 3, 20: 4, 86: 4, 30: 5, 96: 5, 40: 6, 45: 7, 97: 7}.get(cat, 0)

def clean_str(v):
    """rr-core cleanStr: None -> '', else trimmed string."""
    return "" if v is None else str(v).strip()

# %% [markdown]
# ## 5. Per-road distance math
#
# True point-to-polyline distance in feet, using a local equirectangular projection and
# point-to-segment distance (not just nearest vertex). Direct port of
# `modHttp.MinDistanceFt` / `PointSegDistM` / `rr-core minDistanceFt`.

# %%
def _point_seg_dist_m(ax, ay, bx, by):
    """Distance in metres from the local origin (0,0) to segment A-B (planar)."""
    dx, dy = bx - ax, by - ay
    l2 = dx * dx + dy * dy
    if l2 == 0:
        return math.hypot(ax, ay)
    t = -(ax * dx + ay * dy) / l2      # projection of origin onto AB, clamped to [0,1]
    if t < 0:
        t = 0
    if t > 1:
        t = 1
    return math.hypot(ax + t * dx, ay + t * dy)

def min_distance_ft(geometry, lat, lon):
    """Minimum distance in FEET from (lat, lon) to one feature's geometry (polyline paths /
    polygon rings / a point). float('inf') when the feature carries no geometry."""
    if not geometry:
        return math.inf
    if geometry.get("paths"):
        parts = geometry["paths"]
    elif geometry.get("rings"):
        parts = geometry["rings"]
    elif geometry.get("x") is not None:
        parts = [[[geometry["x"], geometry["y"]]]]
    else:
        return math.inf
    if not parts:
        return math.inf
    m_per_lat = 111320.0
    m_per_lon = 111320.0 * math.cos(lat * math.pi / 180.0)
    best = math.inf
    for part in parts:
        pv = [((x - lon) * m_per_lon, (y - lat) * m_per_lat) for x, y in part]
        if len(pv) == 1:
            best = min(best, math.hypot(pv[0][0], pv[0][1]))
        for i in range(len(pv) - 1):
            best = min(best, _point_seg_dist_m(pv[i][0], pv[i][1], pv[i + 1][0], pv[i + 1][1]))
    return best * 3.28084   # metres -> feet

def feature_entries(js, lat, lon):
    """One {'attrs':..., 'distFt':...} per feature of a with-geometry response
    (rr-core featureEntries)."""
    out = []
    for f in (js.get("features") if js else None) or []:
        out.append({"attrs": f.get("attributes") or {},
                    "distFt": min_distance_ft(f.get("geometry"), lat, lon)})
    return out

# %% [markdown]
# ## 6. Query builders (exact-point → buffer fallback)
#
# Every road query tries an exact point-on-road intersect first, then retries with the search
# buffer if nothing hit (rr-core `buildQueryUrl` / `runQuery` / `queryWithFallback`).

# %%
def _query_params(lat, lon, out_fields, where_clause, distance_ft, with_geom):
    """ArcGIS `query` params for a point intersect (rr-core buildQueryUrl)."""
    p = {
        "where": where_clause,
        "geometry": "{},{}".format(lon, lat),
        "geometryType": "esriGeometryPoint",
        "inSR": "4326",
        "spatialRel": "esriSpatialRelIntersects",
        "outFields": out_fields,
        "f": "json",
    }
    if with_geom:
        p["returnGeometry"] = "true"
        p["outSR"] = "4326"
    else:
        p["returnGeometry"] = "false"
    if distance_ft > 0:
        p["distance"] = str(distance_ft)
        p["units"] = "esriSRUnit_Foot"
    return p

def feature_count(js):
    return len(js["features"]) if js and js.get("features") else 0

def first_string(js, field):
    """First non-blank value of `field` across returned features (rr-core firstString)."""
    for f in (js.get("features") if js else None) or []:
        v = (f.get("attributes") or {}).get(field)
        if v is not None and str(v).strip() != "":
            return str(v)
    return ""

def run_query(base_url, lat, lon, out_fields, where_clause, distance_ft, with_geom):
    js = http_get_json(base_url + "/query",
                       _query_params(lat, lon, out_fields, where_clause, distance_ft, with_geom))
    if js and js.get("error"):
        err = js["error"]
        raise RuntimeError("service error {} {}".format(err.get("code", ""), err.get("message", "")))
    return js

def query_with_fallback(base_url, lat, lon, out_fields, where_clause, fallback_ft, with_geom):
    """Exact point intersect first; if no hit, retry with the fallback buffer (CLAUDE.md §4.2)."""
    js = run_query(base_url, lat, lon, out_fields, where_clause, 0, with_geom)
    if feature_count(js) == 0:
        js = run_query(base_url, lat, lon, out_fields, where_clause, fallback_ft, with_geom)
    return js

# %% [markdown]
# ## 7. State NFC queries (MI / IN / WI)
#
# Each returns `(segments, roads)` where `segments` is `[{code, name, distFt}]` (bare FHWA class
# per detected road segment, nearest measured) and `roads` is a separate `[{name, distFt}]` list
# for the merged Road Name display. Retired-segment filters (`RHRetireDate IS NULL` for MI)
# come straight from `web/sources.html`; Indiana's authoritative RO layer needs none (`1=1`).

# %%
def query_michigan_nfc(lat, lon, buffer_ft):
    """MDOT layer 353 class (RHRetireDate IS NULL) + layer 543 trunkline route name."""
    nfc = query_with_fallback(svc("MI_NFC"), lat, lon, "FunctionalSystem,PR",
                              "RHRetireDate IS NULL", buffer_ft, True)
    segments = []
    for e in feature_entries(nfc, lat, lon):
        raw = e["attrs"].get("FunctionalSystem")
        if clean_str(raw) == "":
            continue
        try:
            code = int(float(raw))
        except (TypeError, ValueError):
            continue
        segments.append({"code": code, "name": "", "distFt": e["distFt"]})
    roads = []
    try:   # route name is best-effort, matching modClassify
        route = query_with_fallback(svc("MI_ROUTE"), lat, lon, "RouteDesignation,RouteNumber",
                                    "RHRetireDate IS NULL", buffer_ft, True)
        for e in feature_entries(route, lat, lon):
            nm = " ".join([s for s in (clean_str(e["attrs"].get("RouteDesignation")),
                                       clean_str(e["attrs"].get("RouteNumber"))) if s])
            if nm:
                roads.append({"name": nm, "distFt": e["distFt"]})
    except Exception:
        pass
    return segments, roads

def query_indiana_nfc(lat, lon, buffer_ft):
    """Authoritative INDOT Roads_and_Highways LRSE_Functional_Class (UPPERCASE
    FUNCTIONAL_CLASS, where=1=1 — no record-status filter; the layer holds only
    statuses {1,4,5,null}, none retired, and null-status segments carry real
    classes) + separate centerline road name."""
    nfc = query_with_fallback(svc("IN_NFC"), lat, lon, "FUNCTIONAL_CLASS",
                              "1=1", buffer_ft, True)
    segments = []
    for e in feature_entries(nfc, lat, lon):
        raw = e["attrs"].get("FUNCTIONAL_CLASS")
        if clean_str(raw) == "":
            continue
        try:
            code = int(float(raw))
        except (TypeError, ValueError):
            continue
        segments.append({"code": code, "name": "", "distFt": e["distFt"]})
    roads = []
    try:   # best-effort
        nm = query_with_fallback(svc("IN_ROADNAME"), lat, lon, "st_full", "1=1", buffer_ft, True)
        for e in feature_entries(nm, lat, lon):
            s = clean_str(e["attrs"].get("st_full"))
            if s:
                roads.append({"name": s, "distFt": e["distFt"]})
    except Exception:
        pass
    return segments, roads

def query_wisconsin_nfc(lat, lon, buffer_ft):
    """Local Road Network FIRST (local roads + most collectors — most points), then the
    State Trunk Network only if the local layer has no usable class OR shows a state-highway
    "stub" (null/0 class). Closest-road logic downstream picks the verdict (PR "WI layer swap")."""
    segments, roads = [], []
    saw_stub = False
    loc = query_with_fallback(svc("WI_LOCAL_ROADS"), lat, lon,
                              "FNCT_CLS_CTGY_TYCD,ST_LABL_NM", "1=1", buffer_ft, True)
    for e in feature_entries(loc, lat, lon):
        a = e["attrs"]
        if clean_str(a.get("FNCT_CLS_CTGY_TYCD")) == "":
            saw_stub = True
            continue
        try:
            cat = int(float(a.get("FNCT_CLS_CTGY_TYCD")))
        except (TypeError, ValueError):
            saw_stub = True
            continue
        code = wisconsin_local_category_to_fhwa(cat)
        if code < 1:                       # code 0 / unrecognized = state-highway stub
            saw_stub = True
            continue
        name = clean_str(a.get("ST_LABL_NM"))
        segments.append({"code": code, "name": name, "distFt": e["distFt"]})
        if name:
            roads.append({"name": name, "distFt": e["distFt"]})
    if not segments or saw_stub:
        trunk = query_with_fallback(svc("WI_STATE_TRUNK"), lat, lon,
                                    "FED_FC_CD,HWYTYPE,HWYNUM,HWYDIR", "1=1", buffer_ft, True)
        for e in feature_entries(trunk, lat, lon):
            a = e["attrs"]
            if clean_str(a.get("FED_FC_CD")) == "":
                continue
            try:
                code = int(float(a.get("FED_FC_CD")))
            except (TypeError, ValueError):
                continue
            name = " ".join([s for s in (clean_str(a.get("HWYTYPE")), clean_str(a.get("HWYNUM")),
                                         clean_str(a.get("HWYDIR"))) if s])
            segments.append({"code": code, "name": name, "distFt": e["distFt"]})
            if name:
                roads.append({"name": name, "distFt": e["distFt"]})
    return segments, roads

def query_minnesota_nfc(lat, lon, buffer_ft):
    """MnDOT Functional Class layer 11 - bare FHWA 1-7 integer, no active/retired filter.
    No street-name field (ROUTE_ID is an LRS key); Census TIGER backfills names."""
    nfc = query_with_fallback(svc("MN_NFC"), lat, lon, "FUNCTIONAL_CLASS", "1=1", buffer_ft, True)
    segments = []
    for e in feature_entries(nfc, lat, lon):
        raw = e["attrs"].get("FUNCTIONAL_CLASS")
        if clean_str(raw) == "":
            continue
        try:
            code = int(float(raw))
        except (TypeError, ValueError):
            continue
        segments.append({"code": code, "name": "", "distFt": e["distFt"]})
    return segments, []

def query_illinois_nfc(lat, lon, buffer_ft):
    """IDOT Functional Class layer 0 - FC is a STRING "1".."7"; no filter needed.
    Route-system labels ("FAU 1422") aren't street names; TIGER covers names."""
    nfc = query_with_fallback(svc("IL_NFC"), lat, lon, "FC", "1=1", buffer_ft, True)
    segments = []
    for e in feature_entries(nfc, lat, lon):
        raw = e["attrs"].get("FC")
        if clean_str(raw) == "":
            continue
        try:
            code = int(float(raw))
        except (TypeError, ValueError):
            continue
        segments.append({"code": code, "name": "", "distFt": e["distFt"]})
    return segments, []

def query_ohio_nfc(lat, lon, buffer_ft):
    """ODOT Functional Class layer 0 - FUNCTION_CLASS_CD bare FHWA 1-7; ROUTE_TYPE+
    ROUTE_NBR give route names for IR/US/SR systems ("US 23"); municipal "MR" codes
    are skipped (cryptic) and Census TIGER fills local street names."""
    nfc = query_with_fallback(svc("OH_NFC"), lat, lon,
                              "FUNCTION_CLASS_CD,ROUTE_TYPE,ROUTE_NBR", "1=1", buffer_ft, True)
    segments, roads = [], []
    for e in feature_entries(nfc, lat, lon):
        a = e["attrs"]
        raw = a.get("FUNCTION_CLASS_CD")
        if clean_str(raw) == "":
            continue
        try:
            code = int(float(raw))
        except (TypeError, ValueError):
            continue
        rt = clean_str(a.get("ROUTE_TYPE")).upper()
        rn = clean_str(a.get("ROUTE_NBR"))
        name = "{} {}".format(rt, int(rn)) if rt in ("IR", "US", "SR") and rn.isdigit() else ""
        segments.append({"code": code, "name": name, "distFt": e["distFt"]})
        if name:
            roads.append({"name": name, "distFt": e["distFt"]})
    return segments, roads

# States whose NFC layer is wired (rr-core NFC_WIRED) - all six Region V states (PR #36).
NFC_WIRED = {"MI": query_michigan_nfc, "IN": query_indiana_nfc, "WI": query_wisconsin_nfc,
             "MN": query_minnesota_nfc, "IL": query_illinois_nfc, "OH": query_ohio_nfc}

# %% [markdown]
# ## 8. Verdict logic + ACUB + full per-point pipeline
#
# The verdict model is a direct port of `rr-core computeVerdict`: the **closest** road decides
# red/green; **yellow only downgrades green**; **red never downgrades**. `determine_acub` is the
# three-state urban check (exact-inside / boundary-edge / rural, with the 250 ft floor).

# %%
def class_labels(segments, buffer_ft):
    """Distinct FHWA class labels across segments, nearest-first (rr-core classLabels)."""
    if not segments:
        return "No road segment within {} ft".format(buffer_ft)
    out = []
    for s in segments:
        label = functional_system_label(s["code"])
        if label not in out:
            out.append(label)
    return " | ".join(out)

def _dist(x):
    d = x.get("distFt")
    return math.inf if d is None else d

_COMPASS_TOKENS = {"N", "S", "E", "W", "NE", "NW", "SE", "SW",
                   "NORTH", "SOUTH", "EAST", "WEST"}
_STREET_TYPE_CANON = {
    "ST": "ST", "STREET": "ST", "RD": "RD", "ROAD": "RD", "AVE": "AVE",
    "AV": "AVE", "AVENUE": "AVE", "PKWY": "PKWY", "PKY": "PKWY",
    "PARKWAY": "PKWY", "PKWAY": "PKWY", "DR": "DR", "DRIVE": "DR", "LN": "LN",
    "LANE": "LN", "CT": "CT", "COURT": "CT", "BLVD": "BLVD",
    "BOULEVARD": "BLVD", "TRL": "TRL", "TRAIL": "TRL", "HWY": "HWY",
    "HIGHWAY": "HWY", "CIR": "CIR", "CIRCLE": "CIR", "PL": "PL", "PLACE": "PL",
    "TER": "TER", "TERR": "TER", "TERRACE": "TER", "PT": "PT", "POINT": "PT",
    "SQ": "SQ", "SQUARE": "SQ", "CV": "CV", "COVE": "CV", "WAY": "WAY",
}

def normalize_road_key(name):
    """Canonical dedup key so the same road written two ways ('N Meridian St'
    vs 'MERIDIAN ST', 'Harrison Pkwy' vs 'HARRISON PKY') collapses to one entry
    (modClassify.NormalizeRoadKey / rr-core normalizeRoadKey)."""
    toks = [t for t in re.sub(r"[^A-Z0-9]+", " ", str(name).upper()).split() if t]
    if not toks:
        return ""
    if len(toks) > 1 and toks[0] in _COMPASS_TOKENS:
        toks = toks[1:]
    if len(toks) > 1 and toks[-1] in _COMPASS_TOKENS:
        toks = toks[:-1]
    if len(toks) > 1:
        toks[-1] = _STREET_TYPE_CANON.get(toks[-1], toks[-1])
    return " ".join(toks)

def merge_road_list(roads):
    """Merge detected roads into one nearest-first list, deduped by NORMALIZED
    name (keeping the nearest distance, preferring a mixed-case display over
    ALL-CAPS) — rr-core mergeRoadList / modClassify.FormatRoadList."""
    best = {}
    for r in roads:
        nm = clean_str(r.get("name"))
        if not nm:
            continue
        k = normalize_road_key(nm) or nm.lower()
        d = _dist(r)
        mixed = (nm != nm.upper())
        if k not in best:
            best[k] = {"name": nm, "distFt": d, "mixed": mixed, "nameDist": d}
            continue
        cur = best[k]
        if d < cur["distFt"]:
            cur["distFt"] = d
        if mixed and not cur["mixed"]:
            cur["name"], cur["mixed"], cur["nameDist"] = nm, True, d
        elif mixed == cur["mixed"] and d < cur["nameDist"]:
            cur["name"], cur["nameDist"] = nm, d
    return sorted(best.values(), key=lambda r: r["distFt"])

def format_road_list(road_list):
    return " | ".join(
        r["name"] + (" ({} ft)".format(round(r["distFt"])) if math.isfinite(r["distFt"]) else "")
        for r in road_list)

def class_is_federal(code, is_urban):
    """1-5 always federal-aid; 6 (Minor Collector) only when urban; 7 (Local) never;
    0 (non-certified) handled as 'review' upstream (rr-core classIsFederal)."""
    if 1 <= code <= 5:
        return True
    if code == 6:
        return is_urban
    return False

def compute_verdict(segments, exact_urban, boundary_ambiguous, buffer_ft,
                    named_road_nearby=False):
    """The core verdict — direct port of rr-core computeVerdict. Returns (verdict, reason)."""
    if not segments:
        # A named road is present but the state layer assigns it no functional
        # class (e.g. Mohawk Trail) — flag for review rather than "no road".
        if named_road_nearby:
            return ("Review - road not classified", "Unclassified road")
        return ("Review - no road within {} ft".format(buffer_ft), "No road found")
    ordered = sorted(segments, key=_dist)
    p = ordered[0]

    if p["code"] == 0:
        return ("Review - non-certified class, check manually", "Non-certified")
    if class_is_federal(p["code"], exact_urban):
        return ("Federal aid - " + ("Urban " if exact_urban else "Rural ")
                + functional_system_label(p["code"]), "")

    # Primary road is non-federal -> green base text.
    if p["code"] == 6:
        base_text = "Non-federal aid - Rural Minor Collector"           # a non-fed 6 is by definition rural
    elif p["code"] == 7:
        base_text = "Non-federal aid - " + ("Urban" if exact_urban else "Rural") + " Local"
    else:
        base_text = "Non-federal aid - " + ("Urban " if exact_urban else "Rural ") \
            + functional_system_label(p["code"])

    # Ambiguities that turn green -> yellow, most-specific reason first.
    reason = ""
    if (len(ordered) > 1 and (_dist(ordered[1]) - _dist(p)) < CLOSE_ROAD_FEET
            and class_is_federal(ordered[1]["code"], exact_urban)):
        reason = "Second road close"
    if not reason and any(class_is_federal(s["code"], exact_urban) for s in ordered[1:]):
        reason = "Nearby FHWA road"
    if not reason and p["code"] == 6 and boundary_ambiguous:
        reason = "Urban boundary edge"

    return ("Review - " + reason, reason) if reason else (base_text, "")

def determine_acub(lat, lon, buffer_ft):
    """Point-in-ACUB-polygon (exact) first; if outside every polygon, buffer to the boundary
    floor to detect 'right on the urban edge' (rr-core determineAcub). Returns a dict with
    exact_urban / boundary / name."""
    j = run_query(svc("ACUB"), lat, lon, "NAME,UACE,state_1", "1=1", 0, False)
    if feature_count(j) > 0:
        return {"exact_urban": True, "boundary": False, "name": first_string(j, "NAME")}
    j = run_query(svc("ACUB"), lat, lon, "NAME,UACE,state_1", "1=1",
                  max(buffer_ft, ACUB_MIN_BUFFER_FEET), False)
    if feature_count(j) > 0:
        return {"exact_urban": False, "boundary": True, "name": first_string(j, "NAME")}
    return {"exact_urban": False, "boundary": False, "name": ""}

def classify_point(lat, lon, state_code, buffer_ft=DEFAULT_BUFFER_FEET):
    """Full per-point pipeline, mirroring rr-core classifyPoint / modClassify.ClassifyOneRow:
    ACUB three-state -> state NFC segments + named roads (all with distances) -> TIGER street
    names (non-fatal) -> class label + ambiguity-aware verdict."""
    out = {"lat": lat, "lon": lon, "state": state_code, "urbanRural": "", "acubName": "",
           "acubBoundary": False, "classLabel": "", "roadName": "", "streets": "",
           "streetList": [], "roadList": [], "verdict": "", "reviewReason": "",
           "failed": False, "segments": [], "bufferFt": buffer_ft}

    try:
        acub = determine_acub(lat, lon, buffer_ft)
    except Exception as e:
        out["verdict"] = "Failed - ACUB query ({})".format(e)
        out["failed"] = True
        return out
    out["urbanRural"] = "Urban" if acub["exact_urban"] else "Rural"
    out["acubName"] = acub["name"]
    out["acubBoundary"] = acub["boundary"]

    nfc_query = NFC_WIRED.get(state_code)
    if nfc_query is None:
        out["verdict"] = "ACUB only - class lookup not wired for this state"
        return out

    try:
        segments, roads = nfc_query(lat, lon, buffer_ft)
    except Exception as e:
        out["verdict"] = "Failed - NFC query ({})".format(e)
        out["failed"] = True
        return out
    out["segments"] = sorted(segments, key=_dist)

    # Census TIGER street names (non-fatal) — covers local streets the state layers name poorly.
    try:
        tiger = run_query(svc("TIGER_ROADS"), lat, lon, "NAME", "1=1", buffer_ft, True)
        seen = set()
        for e in feature_entries(tiger, lat, lon):
            nm = clean_str(e["attrs"].get("NAME"))
            if not nm:
                continue
            roads.append({"name": nm, "distFt": e["distFt"]})
            k = nm.lower()
            if k not in seen:
                seen.add(k)
                out["streetList"].append(nm)
        out["streets"] = " | ".join(out["streetList"])
    except Exception:
        pass

    out["roadList"] = merge_road_list(roads)
    out["roadName"] = format_road_list(out["roadList"])
    out["classLabel"] = class_labels(out["segments"], buffer_ft)
    verdict, reason = compute_verdict(out["segments"], acub["exact_urban"], acub["boundary"], buffer_ft,
                                      len(out["roadList"]) > 0)
    out["verdict"] = verdict
    out["reviewReason"] = reason
    return out

# %% [markdown]
# ## 9. Coordinate parsing + state auto-detection
#
# `parse_coordinates` matches the JS parser semantics exactly: one point per line, the last
# adjacent CONUS-looking number pair (lon,lat swap accepted), optional free-text name before or
# after. `detect_state` uses the same rough bounding boxes; a wrong guess fails soft to the
# ACUB-only path.

# %%
def detect_state(lat, lon):
    """Rough CONUS-region auto-detection for the three wired states (rr-core detectState).
    Returns 'MI'/'IN'/'WI' or None (None -> ACUB-only path)."""
    if 37.75 <= lat <= 41.77 and -88.15 <= lon <= -84.75:
        return "IN"
    in_mi = 41.69 <= lat <= 48.35 and -90.45 <= lon <= -82.10
    in_wi = 42.45 <= lat <= 47.35 and -92.95 <= lon <= -86.20
    if in_mi and in_wi:
        return "MI" if lon >= (-88.7 if lat >= 45.4 else -86.7) else "WI"
    if in_mi:
        return "MI"
    if in_wi:
        return "WI"
    return None

_NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")

def parse_coordinates(text):
    """Parse pasted text: one point per line, 'lat lon' or 'lat, lon' with an optional
    free-text name before/after. Uses the last adjacent number pair that looks like a CONUS
    lat/lon (also accepts lon,lat swapped). Direct port of rr-core parseCoordinates."""
    points = []
    for i, raw_line in enumerate(text.split("\n")):
        line = raw_line.replace("\r", "").strip()
        if not line:
            continue
        nums = list(_NUM_RE.finditer(line))
        pair = None
        for j in range(len(nums) - 2, -1, -1):
            a = float(nums[j].group()); b = float(nums[j + 1].group())
            i0, i1 = nums[j].start(), nums[j + 1].end()
            if 17 <= a <= 72 and -180 <= b <= -64:
                pair = {"lat": a, "lon": b, "i0": i0, "i1": i1}; break
            if 17 <= b <= 72 and -180 <= a <= -64:
                pair = {"lat": b, "lon": a, "i0": i0, "i1": i1}; break
        if pair is None:
            points.append({"line": i + 1, "raw": line, "invalid": True})
            continue
        name = (line[:pair["i0"]] + " " + line[pair["i1"]:])
        name = re.sub(r"[,;\t]+", " ", name)
        name = re.sub(r"\s+", " ", name).strip()
        if not name:
            # Mirrors the JS default-name expression exactly, including its string-concatenation
            # quirk: "Point " + <valid-count-so-far> + "1"  (e.g. first unnamed point -> "Point 01").
            valid_so_far = len([p for p in points if not p.get("invalid")])
            name = "Point " + str(valid_so_far) + "1"
        points.append({"line": i + 1, "name": name, "lat": pair["lat"], "lon": pair["lon"]})
    return points

def buffer_feet(radius_str):
    """Parse `radius_ft`, clamp to 1..1000, else default 250 (rr-core bufferFeet)."""
    try:
        v = int(str(radius_str).strip())
    except (TypeError, ValueError):
        return DEFAULT_BUFFER_FEET
    return v if 1 <= v <= 1000 else DEFAULT_BUFFER_FEET

def verdict_bucket(v):
    """Verdict string -> color bucket (rr-core verdictBucket)."""
    if not v:
        return "pending"
    if v.startswith("Federal aid"):
        return "fed"
    if v.startswith("Non-federal aid"):
        return "nonfed"
    if v.startswith("Failed"):
        return "failed"
    return "review"

# %% [markdown]
# ## 10. Self-test — assert the CLAUDE.md test coordinates match the JS
#
# Runs the exact confirmed coordinates from `build/verify-web-core.mjs` / CLAUDE.md §4.2–§4.2b
# (including the two Wisconsin regression points) and asserts the Python verdicts match the JS.
# This hits the live services.
#
# **Set `RUN_SELF_TEST = False` (or delete this cell) before publishing as a web tool** so it
# does not re-run on every web-tool invocation and burn time/credits.

# %%
RUN_SELF_TEST = True

def _run_self_test():
    # [state, lat, lon, expectations] — copied from build/verify-web-core.mjs CASES.
    cases = [
        ["MI", 42.28536, -85.57025, {"verdict": "Federal aid - Urban Minor Collector",
                                     "urbanRural": "Urban", "acubName": "Kalamazoo, MI",
                                     "classIncludes": "Minor Collector"}],
        ["MI", 42.6911, -84.5360, {"verdict": "Non-federal aid - Urban Local",
                                   "urbanRural": "Urban", "acubName": "Lansing, MI"}],
        ["MI", 44.2700, -83.5200, {"verdict": "Non-federal aid - Rural Local", "urbanRural": "Rural"}],
        ["IN", 39.7684, -86.1581, {"verdictStarts": "Federal aid -", "urbanRural": "Urban",
                                   "classIncludes": "Minor Collector"}],
        ["IN", 39.4234, -86.7628, {"verdictStarts": "Federal aid -",
                                   "classIncludes": "Other Principal Arterial"}],
        ["IN", 39.9876, -86.0128, {"verdictStarts": "Non-federal aid -", "classIncludes": "Local"}],
        ["WI", 43.0389, -87.9065, {"verdict": "Federal aid - Urban Minor Arterial",
                                   "urbanRural": "Urban", "acubName": "Milwaukee, WI"}],
        # Local-first WI order (PR "WI layer swap"): point is on W Wisconsin Ave
        # (Rural Minor Collector, ~26 ft), much closer than STH 86 (Major Collector,
        # ~199 ft) which the stub-triggered trunk fallback still surfaces -> yellow.
        ["WI", 45.4711, -89.7345, {"verdict": "Review - Nearby FHWA road",
                                   "urbanRural": "Rural", "classIncludes": "Major Collector"}],
        ["WI", 45.169879, -89.102452, {"verdict": "Non-federal aid - Rural Minor Collector",
                                       "urbanRural": "Rural"}],                       # §7a fix #2 regression
        ["WI", 44.764850, -91.406533, {"verdict": "Federal aid - Urban Interstate",
                                       "acubName": "Eau Claire, WI"}],                # §7a fix #3 regression
        ["MN", 44.9778, -93.2650, {"verdict": "ACUB only - class lookup not wired for this state",
                                   "urbanRural": "Urban"}],                           # state gate
    ]
    npass = nfail = 0
    for state, lat, lon, want in cases:
        try:
            r = classify_point(lat, lon, state)
        except Exception as e:
            print("  FAIL {} {},{} — threw {}".format(state, lat, lon, e)); nfail += 1; continue
        problems = []
        if "verdict" in want and r["verdict"] != want["verdict"]:
            problems.append('verdict "{}" != "{}"'.format(r["verdict"], want["verdict"]))
        if "verdictStarts" in want and not r["verdict"].startswith(want["verdictStarts"]):
            problems.append('verdict "{}" !^ "{}"'.format(r["verdict"], want["verdictStarts"]))
        if "urbanRural" in want and r["urbanRural"] != want["urbanRural"]:
            problems.append('urbanRural "{}" != "{}"'.format(r["urbanRural"], want["urbanRural"]))
        if "acubName" in want and r["acubName"] != want["acubName"]:
            problems.append('acubName "{}" != "{}"'.format(r["acubName"], want["acubName"]))
        if "classIncludes" in want and want["classIncludes"] not in r["classLabel"]:
            problems.append('classLabel "{}" !~ "{}"'.format(r["classLabel"], want["classIncludes"]))
        # PR #24 parity: wired states must carry finite per-segment distances + a sorted road list.
        if state in ("MI", "IN", "WI"):
            if not r["segments"] or not all(math.isfinite(s["distFt"]) for s in r["segments"]):
                problems.append("segments missing finite distFt")
            if r["roadName"] and not re.search(r"\(\d+ ft\)", r["roadName"]):
                problems.append('roadName "{}" missing (N ft) distances'.format(r["roadName"]))
            dists = [x["distFt"] for x in r["roadList"]]
            if any(i > 0 and d < dists[i - 1] for i, d in enumerate(dists)):
                problems.append("roadList not sorted nearest-first")
        tag = "{} {},{} -> {}".format(state, lat, lon, r["verdict"])
        if r["reviewReason"]:
            tag += " ({})".format(r["reviewReason"])
        if r["roadName"]:
            tag += " [{}]".format(r["roadName"])
        if problems:
            print("  FAIL " + tag + " — " + "; ".join(problems)); nfail += 1
        else:
            print("  ok   " + tag); npass += 1
    print("\n{} passed, {} failed".format(npass, nfail))
    assert nfail == 0, "self-test failed: {} case(s) do not match the JS rr-core".format(nfail)
    return npass, nfail

if RUN_SELF_TEST:
    _run_self_test()

# %% [markdown]
# ## 11. Run the classifier over the pasted coordinates
#
# Parse `coords_text`, auto-detect each point's state, classify it, and collect one flat record
# per valid point. Invalid lines are skipped (a web tool has no UI to surface them). Records use
# the same field names as the repo's GeoJSON export shape.

# %%
def classify_coords_text(text, radius_str):
    """Parse + classify every point, returning a list of flat attribute dicts (GeoJSON export
    shape). Category/Description are always '' here (the web-tool input carries only name +
    coordinates), but the fields are kept so the output schema matches the Excel/web export."""
    buf = buffer_feet(radius_str)
    records = []
    for p in parse_coordinates(text):
        if p.get("invalid"):
            continue
        state = detect_state(p["lat"], p["lon"]) or "XX"   # XX -> ACUB-only path
        r = classify_point(p["lat"], p["lon"], state, buf)
        bucket = verdict_bucket(r["verdict"])
        records.append({
            "Name":         p["name"],
            "FedAidStatus": r["verdict"],
            "Verdict":      BUCKET_GEO_LABEL.get(bucket, "Review"),
            "VerdictColor": BUCKET_COLOR.get(bucket, BUCKET_COLOR["review"]),
            "FHWAClass":    r["classLabel"],
            "UrbanRural":   r["urbanRural"],
            "ACUBName":     r["acubName"],
            "RoadName":     r["roadName"],
            "StreetName":   r["streets"],
            "ReviewNote":   r["reviewReason"],
            "Category":     "",
            "Description":  "",
            "Latitude":     p["lat"],
            "Longitude":    p["lon"],
        })
    return records

classified_records = classify_coords_text(coords_text, radius_ft)
print("Classified {} point(s):".format(len(classified_records)))
for rec in classified_records:
    print("  {:<24} {}".format(rec["Name"][:24], rec["FedAidStatus"]))

# %% [markdown]
# ## 12. Package the results as a Feature set (WGS84)
#
# Builds an `arcgis` **FeatureSet** of Point features in WGS84, with the flat attributes above.
# This is the object the web tool returns. `arcgis` ships preinstalled on the AGOL notebook
# runtime; no install needed. (This cell is skipped automatically if `arcgis` is unavailable —
# e.g. running the notebook locally for the self-test — so the rest still executes.)

# %%
def build_feature_set(records):
    """List of flat attribute dicts -> arcgis.features.FeatureSet (Point, WGS84)."""
    from arcgis.features import Feature, FeatureSet
    feats = []
    for rec in records:
        geom = {"x": rec["Longitude"], "y": rec["Latitude"], "spatialReference": {"wkid": 4326}}
        # Attributes minus the lat/lon we carry in geometry-adjacent fields (kept in attrs too,
        # per the GeoJSON export shape, so the attribute table shows them).
        feats.append(Feature(geometry=geom, attributes=dict(rec)))
    return FeatureSet(feats, geometry_type="esriGeometryPoint",
                      spatial_reference={"wkid": 4326})

try:
    output_features = build_feature_set(classified_records)
    print("Built FeatureSet with {} feature(s).".format(len(output_features.features)))
except ImportError:
    output_features = None
    print("arcgis not installed here — FeatureSet build skipped (fine for local self-test; "
          "on the AGOL runtime arcgis is preinstalled).")

# %% [markdown]
# ## 13. Web-tool output (very bottom — the "Insert output" snippet)
#
# When you add a **Feature set** *output* parameter in the Parameters pane and click
# **Insert output variable**, AGOL appends a snippet here that references the FeatureSet
# variable. Point that output parameter at **`output_features`** (built in cell 12). Keeping
# the FeatureSet as the last expression is what the web-tool publisher captures as the result.

# %%
# ─── WEB-TOOL OUTPUT ─────────────────────────────────────────────────────────
# The Parameters pane's "Insert output variable" maps `output_features` (an arcgis FeatureSet)
# to the Feature set output parameter. Leaving it as the final expression returns it to callers
# (Map Viewer, the Experience Builder Analysis widget, or a REST submitJob request).
output_features
