# Databricks notebook source
# MAGIC %md
# MAGIC # Build the `reviewer_*` Delta tables for the FHWA Road Checker app
# MAGIC
# MAGIC Import this file as a notebook (Workspace > Import > File) and **Run all**
# MAGIC on any cluster with internet access. It harvests the same public sources the
# MAGIC web tool's tile pipeline uses (`build/tiles/` in the repo) and lands them in
# MAGIC **two** Unity Catalog Delta tables - deliberately as few as possible:
# MAGIC
# MAGIC | table | rows | what |
# MAGIC |---|---|---|
# MAGIC | `reviewer_roads` | ~3.2 M | every HPMS road segment in the six Region V states: FHWA class 1-7 (`f_system`), the state's own LRS keys (`route_id`, mileposts, route name/number), bbox columns and the GeoJSON line |
# MAGIC | `reviewer_boundaries` | ~1.1 k | `kind` = `urban_area` (NTAD 2020 Adjusted Urban Areas = ACUB), `county` (Census TIGER counties) and `state` (the six state outlines), each with name, code, bbox and the GeoJSON polygon |
# MAGIC
# MAGIC Sources (all public, no auth):
# MAGIC * HPMS: `https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/HPMS_National_Current/FeatureServer/0` (USDOT/BTS)
# MAGIC * ACUB: `.../NTAD_Adjusted_Urban_Areas/FeatureServer/0` (USDOT/BTS NTAD)
# MAGIC * Counties + states: Census TIGERweb `State_County/MapServer` layers 1 and 0
# MAGIC
# MAGIC Re-running is safe: each state's rows are replaced, not duplicated. The whole
# MAGIC harvest takes roughly 30-60 minutes (the HPMS layer answers envelope queries in
# MAGIC seconds but refuses offset pagination - see the quadtree note below).

# COMMAND ----------

dbutils.widgets.text("catalog", "main")
dbutils.widgets.text("schema", "roadreviewer")
dbutils.widgets.text("states", "MI,IN,WI,MN,IL,OH")   # subset to rebuild, e.g. "MI"
dbutils.widgets.dropdown("boundaries", "yes", ["yes", "no"])

CATALOG = dbutils.widgets.get("catalog").strip()
SCHEMA = dbutils.widgets.get("schema").strip()
STATES = [s.strip().upper() for s in dbutils.widgets.get("states").split(",") if s.strip()]
DO_BOUNDARIES = dbutils.widgets.get("boundaries") == "yes"

T_ROADS = f"`{CATALOG}`.`{SCHEMA}`.`reviewer_roads`"
T_BOUND = f"`{CATALOG}`.`{SCHEMA}`.`reviewer_boundaries`"

spark.sql(f"CREATE SCHEMA IF NOT EXISTS `{CATALOG}`.`{SCHEMA}`")
spark.sql(f"""
CREATE TABLE IF NOT EXISTS {T_ROADS} (
  state        STRING  COMMENT '2-letter state code',
  state_fips   INT,
  objectid     BIGINT  COMMENT 'HPMS OBJECTID (dedupe key)',
  f_system     INT     COMMENT 'FHWA functional class 1-7 (1 Interstate ... 7 Local)',
  route_id     STRING  COMMENT 'state LRS route id as submitted to HPMS',
  begin_point  DOUBLE  COMMENT 'begin milepost',
  end_point    DOUBLE  COMMENT 'end milepost',
  route_name   STRING,
  route_number STRING,
  min_lon DOUBLE, min_lat DOUBLE, max_lon DOUBLE, max_lat DOUBLE,
  geometry     STRING  COMMENT 'GeoJSON LineString/MultiLineString, WGS84',
  source       STRING,
  loaded_at    TIMESTAMP
) USING DELTA PARTITIONED BY (state)
COMMENT 'FHWA HPMS road segments for the FHWA Road Checker (reviewer app): class + state LRS keys + geometry'
""")
spark.sql(f"""
CREATE TABLE IF NOT EXISTS {T_BOUND} (
  kind      STRING COMMENT 'urban_area | county | state',
  name      STRING,
  code      STRING COMMENT 'UACE for urban areas, GEOID for counties, FIPS for states',
  state     STRING COMMENT '2-letter state code (urban areas may span states: first listed)',
  min_lon DOUBLE, min_lat DOUBLE, max_lon DOUBLE, max_lat DOUBLE,
  geometry  STRING COMMENT 'GeoJSON Polygon/MultiPolygon, WGS84',
  source    STRING,
  loaded_at TIMESTAMP
) USING DELTA
COMMENT 'Boundaries for the FHWA Road Checker (reviewer app): 2020 Adjusted Urban Areas (ACUB), counties, states'
""")
print("tables ready:", T_ROADS, T_BOUND)

# COMMAND ----------

import json, time, math
from collections import deque
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
from datetime import datetime, timezone
import requests
from pyspark.sql import types as T

HPMS = "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/HPMS_National_Current/FeatureServer/0"
ACUB = "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0"
TIGER_COUNTIES = "https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/State_County/MapServer/1"
TIGER_STATES = "https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/State_County/MapServer/0"

# Generous per-state boxes; the STATE_ID where-clause clips exactly.
STATE_INFO = {
    "MI": (26, [-90.5, 41.6, -82.0, 48.4]),
    "WI": (55, [-93.0, 42.4, -86.1, 47.4]),
    "MN": (27, [-97.3, 43.4, -89.4, 49.5]),
    "IL": (17, [-91.6, 36.9, -87.0, 42.6]),
    "IN": (18, [-88.2, 37.7, -84.7, 41.8]),
    "OH": (39, [-84.9, 38.3, -80.4, 42.0]),
}
FIPS_TO_ST = {v[0]: k for k, v in STATE_INFO.items()}

PAGE_CAP = 2000        # the HPMS layer's maxRecordCount
CONCURRENCY = 4
MIN_CELL_DEG = 0.005   # never split below ~500 m
SEED_STEP = 0.5        # ~0.5 degree seed cells answer in seconds; 1 degree metro cells don't answer at all
FLUSH_ROWS = 25_000

session = requests.Session()
session.headers["User-Agent"] = "RoadReviewer-databricks/1.0 (+https://github.com/calebpaulsmith/RoadReviewer)"


class TooBig(Exception):
    """The server gave up on a cell that covers too many rows ('Unable to perform
    query' after ~55 s) - not transient: split the cell, don't retry."""


def get_json(url, params, tries=5, timeout=170):
    for i in range(tries):
        try:
            r = session.get(url, params=params, timeout=timeout)
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code}")
            j = r.json()
            if "error" in j:
                if "Unable to perform query" in json.dumps(j["error"]):
                    raise TooBig()
                raise RuntimeError("service error " + json.dumps(j["error"])[:200])
            return j
        except TooBig:
            raise
        except Exception as e:  # noqa: BLE001
            if i == tries - 1:
                raise
            time.sleep(3 * (i + 1))


def bbox_of(coords, acc=None):
    acc = acc or [180.0, 90.0, -180.0, -90.0]
    if coords and isinstance(coords[0], (int, float)):
        x, y = coords[0], coords[1]
        acc[0] = min(acc[0], x); acc[1] = min(acc[1], y); acc[2] = max(acc[2], x); acc[3] = max(acc[3], y)
    else:
        for c in coords:
            bbox_of(c, acc)
    return acc


ROADS_SCHEMA = T.StructType([
    T.StructField("state", T.StringType()), T.StructField("state_fips", T.IntegerType()),
    T.StructField("objectid", T.LongType()), T.StructField("f_system", T.IntegerType()),
    T.StructField("route_id", T.StringType()), T.StructField("begin_point", T.DoubleType()),
    T.StructField("end_point", T.DoubleType()), T.StructField("route_name", T.StringType()),
    T.StructField("route_number", T.StringType()),
    T.StructField("min_lon", T.DoubleType()), T.StructField("min_lat", T.DoubleType()),
    T.StructField("max_lon", T.DoubleType()), T.StructField("max_lat", T.DoubleType()),
    T.StructField("geometry", T.StringType()), T.StructField("source", T.StringType()),
    T.StructField("loaded_at", T.TimestampType()),
])
BOUND_SCHEMA = T.StructType([
    T.StructField("kind", T.StringType()), T.StructField("name", T.StringType()),
    T.StructField("code", T.StringType()), T.StructField("state", T.StringType()),
    T.StructField("min_lon", T.DoubleType()), T.StructField("min_lat", T.DoubleType()),
    T.StructField("max_lon", T.DoubleType()), T.StructField("max_lat", T.DoubleType()),
    T.StructField("geometry", T.StringType()), T.StructField("source", T.StringType()),
    T.StructField("loaded_at", T.TimestampType()),
])

# COMMAND ----------

# MAGIC %md
# MAGIC ## HPMS roads - quadtree envelope harvest
# MAGIC
# MAGIC The national table is ~30 M rows and any attribute-filtered *offset* page times
# MAGIC out server-side (~55 s then HTTP 400), while an *envelope* query answers in
# MAGIC seconds because it rides the spatial index. So: seed the state box with 0.5-degree
# MAGIC cells; any cell that comes back at the 2,000-feature cap (or that the server gives
# MAGIC up on) splits into four; leaf cells stream to the table. Segments crossing cell
# MAGIC borders arrive twice and are deduped by OBJECTID.

# COMMAND ----------

def cell_params(fips, cell):
    x0, y0, x1, y1 = cell
    return {
        "where": f"STATE_ID={fips}",
        "geometry": f"{x0:.5f},{y0:.5f},{x1:.5f},{y1:.5f}",
        "geometryType": "esriGeometryEnvelope", "inSR": 4326,
        "spatialRel": "esriSpatialRelIntersects",
        "outFields": "OBJECTID,F_SYSTEM,ROUTE_ID,BEGIN_POINT,END_POINT,RouteName,RouteNumber",
        "returnGeometry": "true", "outSR": 4326, "geometryPrecision": 6, "f": "geojson",
    }


def fetch_cell(fips, cell):
    return get_json(f"{HPMS}/query", cell_params(fips, cell))


def harvest_state(st):
    fips, box = STATE_INFO[st]
    loaded_at = datetime.now(timezone.utc)
    spark.sql(f"DELETE FROM {T_ROADS} WHERE state = '{st}'")
    queue = deque()
    x0, y0, x1, y1 = box
    x = x0
    while x < x1:
        y = y0
        while y < y1:
            queue.append((x, y, min(x + SEED_STEP, x1), min(y + SEED_STEP, y1)))
            y += SEED_STEP
        x += SEED_STEP

    seen, buf = set(), []
    written = cells = splits = 0

    def flush():
        nonlocal buf, written
        if not buf:
            return
        spark.createDataFrame(buf, ROADS_SCHEMA).write.mode("append").saveAsTable(T_ROADS.replace("`", ""))
        written += len(buf)
        buf = []

    with ThreadPoolExecutor(CONCURRENCY) as ex:
        futures = {}
        while queue or futures:
            while queue and len(futures) < CONCURRENCY:
                cell = queue.popleft()
                futures[ex.submit(fetch_cell, fips, cell)] = cell
            done, _ = wait(list(futures), return_when=FIRST_COMPLETED)
            for fut in done:
                cell = futures.pop(fut)
                cx0, cy0, cx1, cy1 = cell
                wide = (cx1 - cx0) > MIN_CELL_DEG

                def split():
                    nonlocal splits
                    mx, my = (cx0 + cx1) / 2, (cy0 + cy1) / 2
                    queue.extend([(cx0, cy0, mx, my), (mx, cy0, cx1, my), (cx0, my, mx, cy1), (mx, my, cx1, cy1)])
                    splits += 1

                try:
                    j = fut.result()
                except TooBig:
                    if wide:
                        print(f"  give-up cell {cx1 - cx0:.2f}deg @ {cx0:.2f},{cy0:.2f} -> split")
                        split()
                        continue
                    raise
                feats = j.get("features") or []
                exceeded = bool((j.get("properties") or {}).get("exceededTransferLimit") or j.get("exceededTransferLimit"))
                if (exceeded or len(feats) >= PAGE_CAP) and wide:
                    split()
                    continue
                for f in feats:
                    p = f.get("properties") or {}
                    oid = p.get("OBJECTID")
                    g = f.get("geometry")
                    try:
                        cls = int(float(p.get("F_SYSTEM")))
                    except (TypeError, ValueError):
                        continue
                    if oid is None or oid in seen or not (1 <= cls <= 7) or not g or not g.get("coordinates"):
                        continue
                    seen.add(oid)
                    bb = bbox_of(g["coordinates"])
                    rid = p.get("ROUTE_ID")
                    rn = p.get("RouteNumber")
                    buf.append((
                        st, fips, int(oid), cls,
                        str(rid).strip() if rid not in (None, "") else None,
                        float(p["BEGIN_POINT"]) if p.get("BEGIN_POINT") is not None else None,
                        float(p["END_POINT"]) if p.get("END_POINT") is not None else None,
                        str(p["RouteName"]).strip() if p.get("RouteName") not in (None, "") else None,
                        str(rn).strip() if rn not in (None, "", 0) else None,
                        bb[0], bb[1], bb[2], bb[3],
                        json.dumps(g, separators=(",", ":")), "HPMS_National_Current", loaded_at,
                    ))
                cells += 1
                if len(buf) >= FLUSH_ROWS:
                    flush()
                if cells % 25 == 0:
                    print(f"  {st}: {cells} cells, {splits} splits, {written + len(buf)} segments, queue {len(queue)}")
    flush()
    print(f"{st} DONE: {written} segments from {cells} leaf cells ({splits} splits)")
    return written


for st in STATES:
    print(f"=== {st} (FIPS {STATE_INFO[st][0]}) ===")
    harvest_state(st)

# COMMAND ----------

# MAGIC %md
# MAGIC ## Boundaries - urban areas (ACUB 2020), counties, states

# COMMAND ----------

def ids_in_box(layer, box, where="1=1"):
    x0, y0, x1, y1 = box
    j = get_json(f"{layer}/query", {
        "where": where, "geometry": f"{x0},{y0},{x1},{y1}", "geometryType": "esriGeometryEnvelope",
        "inSR": 4326, "spatialRel": "esriSpatialRelIntersects", "returnIdsOnly": "true", "f": "json"})
    return j.get("objectIds") or []


def fetch_by_ids(layer, ids, out_fields, batch=25, offset=0.00003):
    """Geometry in small objectIds batches with ~3 m server-side generalization -
    full precision over a whole state 504s, and the verdict's boundary rule has
    a 76 m floor so <=3 m of shift is noise."""
    for i in range(0, len(ids), batch):
        j = get_json(f"{layer}/query", {
            "objectIds": ",".join(str(x) for x in ids[i:i + batch]), "outFields": out_fields,
            "returnGeometry": "true", "outSR": 4326, "geometryPrecision": 5,
            "maxAllowableOffset": offset, "f": "geojson"})
        for f in j.get("features") or []:
            if f.get("geometry"):
                yield f


def boundary_row(kind, name, code, st, geom, source, loaded_at):
    bb = bbox_of(geom["coordinates"])
    return (kind, name, code, st, bb[0], bb[1], bb[2], bb[3], json.dumps(geom, separators=(",", ":")), source, loaded_at)


if DO_BOUNDARIES:
    loaded_at = datetime.now(timezone.utc)
    rows = []

    # urban areas: union of ids over the six state boxes, then geometry
    ids = set()
    for st, (fips, box) in STATE_INFO.items():
        got = ids_in_box(ACUB, box)
        ids.update(got)
        print(f"ACUB {st}: {len(got)} ids (union {len(ids)})")
    for f in fetch_by_ids(ACUB, sorted(ids), "NAME,UACE,state_1"):
        p = f["properties"]
        rows.append(boundary_row("urban_area", p.get("NAME"), str(p.get("UACE") or ""), (p.get("state_1") or "")[:2],
                                 f["geometry"], "NTAD_Adjusted_Urban_Areas (2020)", loaded_at))
    print(f"urban areas: {len(rows)}")

    # counties + states of the six states (TIGERweb, by STATE FIPS)
    fips_list = ",".join(f"'{v[0]:02d}'" for v in STATE_INFO.values())
    n0 = len(rows)
    cids = get_json(f"{TIGER_COUNTIES}/query", {"where": f"STATE IN ({fips_list})", "returnIdsOnly": "true", "f": "json"}).get("objectIds") or []
    for f in fetch_by_ids(TIGER_COUNTIES, sorted(cids), "NAME,GEOID,STATE,BASENAME", batch=20, offset=0.0001):
        p = f["properties"]
        st = FIPS_TO_ST.get(int(p.get("STATE") or 0), "")
        rows.append(boundary_row("county", p.get("NAME"), str(p.get("GEOID") or ""), st, f["geometry"], "Census TIGERweb State_County/1", loaded_at))
    print(f"counties: {len(rows) - n0}")

    n0 = len(rows)
    sids = get_json(f"{TIGER_STATES}/query", {"where": f"GEOID IN ({fips_list})", "returnIdsOnly": "true", "f": "json"}).get("objectIds") or []
    for f in fetch_by_ids(TIGER_STATES, sorted(sids), "NAME,GEOID,STUSAB", batch=3, offset=0.0005):
        p = f["properties"]
        rows.append(boundary_row("state", p.get("NAME"), str(p.get("GEOID") or ""), (p.get("STUSAB") or "").upper(), f["geometry"], "Census TIGERweb State_County/0", loaded_at))
    print(f"states: {len(rows) - n0}")

    spark.sql(f"DELETE FROM {T_BOUND} WHERE kind IN ('urban_area','county','state')")
    spark.createDataFrame(rows, BOUND_SCHEMA).write.mode("append").saveAsTable(T_BOUND.replace("`", ""))
    print(f"boundaries written: {len(rows)}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## Compact + cluster for point lookups, then sanity-check

# COMMAND ----------

spark.sql(f"OPTIMIZE {T_ROADS} ZORDER BY (min_lon, min_lat)")
spark.sql(f"OPTIMIZE {T_BOUND} ZORDER BY (min_lon, min_lat)")
display(spark.sql(f"SELECT state, f_system, COUNT(*) AS segments FROM {T_ROADS} GROUP BY state, f_system ORDER BY state, f_system"))
display(spark.sql(f"SELECT kind, COUNT(*) AS n FROM {T_BOUND} GROUP BY kind"))

# Known verdict point (CLAUDE.md section 4.2 test #1): Kalamazoo, MI - expect a class-6 segment within ~250 ft
# and the 'Kalamazoo, MI' urban area.
lat, lon, ft = 42.28536, -85.57025, 250
d = ft / 364000.0
display(spark.sql(f"""
SELECT f_system, route_id, route_name FROM {T_ROADS}
WHERE state = 'MI' AND max_lon >= {lon - d * 1.4} AND min_lon <= {lon + d * 1.4} AND max_lat >= {lat - d} AND min_lat <= {lat + d}
"""))
display(spark.sql(f"""
SELECT kind, name FROM {T_BOUND}
WHERE max_lon >= {lon} AND min_lon <= {lon} AND max_lat >= {lat} AND min_lat <= {lat}
"""))
