"""RoadReviewer / FHWA Road Checker as a Databricks App.

Serves the static web tool (web/) exactly as GitHub Pages does - including
the PMTiles archives, which need HTTP Range requests - and adds a JSON API
that reads two Unity Catalog Delta tables (built by
setup/build_reviewer_tables.py):

    reviewer_roads        HPMS road segments (FHWA class 1-7, state LRS keys)
    reviewer_boundaries   ACUB 2020 urban areas + counties + state outlines

Endpoints
    GET /api/health                       {ok, delta, catalog, schema, tables}
    GET /api/features?kind=roads&state=MI&lat=..&lon=..&ft=250
    GET /api/features?kind=acub&lat=..&lon=..&ft=250
    GET /api/features?kind=county&lat=..&lon=..&ft=250
                                          features near a point in the page's
                                          tile-feature shape {props, geomType, parts}
    GET /api/area?lat=..&lon=..           county + urban area + state at a point

The page (web/index.html) probes /api/health at load; when `delta` is true
its cached-verdict path reads /api/features instead of the tiles, through
the same verdict code. Without a warehouse (local dev, or a missing
binding) the API says delta:false and the page uses the tiles - nothing
breaks, the data source just changes.

Local run (for a look, no Delta), from databricks/:
    RR_WEB_DIR=../web uvicorn app:app --reload
"""
from __future__ import annotations

import json
import math
import os
import re
import threading
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import FileResponse, JSONResponse, Response
from starlette.staticfiles import StaticFiles

CATALOG = os.environ.get("RR_CATALOG", "main")
SCHEMA = os.environ.get("RR_SCHEMA", "roadreviewer")
WAREHOUSE_ID = os.environ.get("DATABRICKS_WAREHOUSE_ID", "")
WEB_DIR = Path(os.environ.get("RR_WEB_DIR", "web")).resolve()
TILES_DIR = Path(os.environ.get("RR_TILES_DIR", str(WEB_DIR / "tiles"))).resolve()

T_ROADS = f"`{CATALOG}`.`{SCHEMA}`.`reviewer_roads`"
T_BOUND = f"`{CATALOG}`.`{SCHEMA}`.`reviewer_boundaries`"
TABLES = ["reviewer_roads", "reviewer_boundaries"]

MAX_RADIUS_FT = 6000
FT_PER_DEG_LAT = 364_000.0  # 111,320 m / 0.3048

app = FastAPI(title="FHWA Road Checker (Databricks)", docs_url=None, redoc_url=None)

# ---------------------------------------------------------------- SQL access
_conn_lock = threading.Lock()
_conn: Any = None
_delta_error: str | None = None


def _connect():
    """One shared connection to the bound SQL warehouse, authenticated as the
    app's service principal (DATABRICKS_HOST / CLIENT_ID / CLIENT_SECRET are
    injected by the Apps runtime)."""
    global _conn, _delta_error
    if _conn is not None:
        return _conn
    with _conn_lock:
        if _conn is not None:
            return _conn
        if not WAREHOUSE_ID:
            _delta_error = "no SQL warehouse bound (DATABRICKS_WAREHOUSE_ID is empty)"
            raise RuntimeError(_delta_error)
        try:
            from databricks import sql as dbsql
            from databricks.sdk.core import Config, oauth_service_principal

            cfg = Config()
            host = (cfg.host or "").replace("https://", "").rstrip("/")
            _conn = dbsql.connect(
                server_hostname=host,
                http_path=f"/sql/1.0/warehouses/{WAREHOUSE_ID}",
                credentials_provider=lambda: oauth_service_principal(cfg),
            )
            _delta_error = None
            return _conn
        except Exception as e:  # noqa: BLE001
            _delta_error = f"{type(e).__name__}: {e}"
            raise


def _query(sql: str, params: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    global _conn
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params or {})
            cols = [d[0] for d in cur.description]
            return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception:
        with _conn_lock:  # drop a broken connection so the next call reconnects
            try:
                conn.close()
            except Exception:  # noqa: BLE001
                pass
            _conn = None
        raise


# ------------------------------------------------------------- geometry help
def _bbox(lat: float, lon: float, radius_ft: float) -> tuple[float, float, float, float]:
    d_lat = radius_ft / FT_PER_DEG_LAT
    d_lon = d_lat / max(0.2, math.cos(math.radians(lat)))
    return lon - d_lon, lat - d_lat, lon + d_lon, lat + d_lat


def _parts(geom: dict[str, Any]) -> tuple[int, list[list[list[float]]]]:
    """GeoJSON -> (geomType, parts) in the page's tile-feature shape:
    geomType 2 = lines (each part a path), 3 = polygons (each part a ring)."""
    t = geom.get("type")
    c = geom.get("coordinates") or []
    if t == "LineString":
        return 2, [c]
    if t == "MultiLineString":
        return 2, c
    if t == "Polygon":
        return 3, c
    if t == "MultiPolygon":
        return 3, [ring for poly in c for ring in poly]
    raise ValueError(f"unsupported geometry {t}")


def _point_in_rings(parts: list[list[list[float]]], lat: float, lon: float) -> bool:
    inside = False
    for ring in parts:
        j = len(ring) - 1
        for i in range(len(ring)):
            xi, yi = ring[i][0], ring[i][1]
            xj, yj = ring[j][0], ring[j][1]
            if (yi > lat) != (yj > lat) and lon < (xj - xi) * (lat - yi) / (yj - yi) + xi:
                inside = not inside
            j = i
    return inside


# --------------------------------------------------------------------- API
@app.get("/api/health")
def health() -> JSONResponse:
    info: dict[str, Any] = {"ok": True, "delta": False, "catalog": CATALOG, "schema": SCHEMA, "tables": TABLES}
    try:
        rows = _query(
            f"SELECT (SELECT COUNT(*) FROM {T_ROADS}) AS roads, (SELECT COUNT(*) FROM {T_BOUND}) AS boundaries"
        )
        info.update({"delta": True, "roads": rows[0]["roads"], "boundaries": rows[0]["boundaries"]})
    except Exception as e:  # noqa: BLE001
        info["reason"] = _delta_error or f"{type(e).__name__}: {e}"
    return JSONResponse(info, headers={"Cache-Control": "no-store"})


@app.get("/api/features")
def features(
    kind: str = Query(..., pattern="^(roads|acub|county)$"),
    lat: float = Query(..., ge=-90, le=90),
    lon: float = Query(..., ge=-180, le=180),
    ft: float = Query(250, gt=0, le=MAX_RADIUS_FT),
    state: str = Query("", max_length=2),
) -> JSONResponse:
    """Features within a bbox of `ft` feet around the point, in the shape the
    page's tileFeaturesNear() returns so classifyPointTiles() needs no changes:
    roads -> props {F, R, B, E, N, RN}; acub -> props {NAME, UACE, state_1};
    county -> props {NAME, GEOID, STATE}."""
    x0, y0, x1, y1 = _bbox(lat, lon, ft)
    params: dict[str, Any] = {"x0": x0, "y0": y0, "x1": x1, "y1": y1}
    if kind == "roads":
        st = state.upper()
        if not re.fullmatch(r"[A-Z]{2}", st):
            raise HTTPException(400, "state (2-letter code) is required for kind=roads")
        params["st"] = st
        sql = (
            f"SELECT objectid, f_system, route_id, begin_point, end_point, route_name, route_number, geometry "
            f"FROM {T_ROADS} WHERE state = %(st)s "
            f"AND max_lon >= %(x0)s AND min_lon <= %(x1)s AND max_lat >= %(y0)s AND min_lat <= %(y1)s"
        )
    else:
        params["kind"] = "urban_area" if kind == "acub" else "county"
        sql = (
            f"SELECT name, code, state, geometry FROM {T_BOUND} WHERE kind = %(kind)s "
            f"AND max_lon >= %(x0)s AND min_lon <= %(x1)s AND max_lat >= %(y0)s AND min_lat <= %(y1)s"
        )
    try:
        rows = _query(sql, params)
    except Exception as e:  # noqa: BLE001
        return JSONResponse({"error": f"Delta query failed: {_delta_error or e}"}, status_code=503)

    out = []
    for r in rows:
        try:
            gt, parts = _parts(json.loads(r["geometry"]))
        except Exception:  # noqa: BLE001
            continue
        if kind == "roads":
            props: dict[str, Any] = {"F": int(r["f_system"])}
            if r.get("route_id"):
                props["R"] = str(r["route_id"])
            if r.get("begin_point") is not None:
                props["B"] = float(r["begin_point"])
            if r.get("end_point") is not None:
                props["E"] = float(r["end_point"])
            if r.get("route_name"):
                props["N"] = r["route_name"]
            if r.get("route_number"):
                props["RN"] = r["route_number"]
        elif kind == "acub":
            props = {"NAME": r["name"], "UACE": r["code"], "state_1": r["state"]}
        else:
            props = {"NAME": r["name"], "GEOID": r["code"], "STATE": r["state"]}
        out.append({"props": props, "geomType": gt, "parts": parts})
    return JSONResponse({"features": out, "count": len(out), "source": "delta"})


@app.get("/api/area")
def area(lat: float = Query(..., ge=-90, le=90), lon: float = Query(..., ge=-180, le=180)) -> JSONResponse:
    """County, urban area and state containing the point (exact point-in-polygon)."""
    x0, y0, x1, y1 = _bbox(lat, lon, 10)
    try:
        rows = _query(
            f"SELECT kind, name, code, state, geometry FROM {T_BOUND} "
            f"WHERE max_lon >= %(x0)s AND min_lon <= %(x1)s AND max_lat >= %(y0)s AND min_lat <= %(y1)s",
            {"x0": x0, "y0": y0, "x1": x1, "y1": y1},
        )
    except Exception as e:  # noqa: BLE001
        return JSONResponse({"error": f"Delta query failed: {_delta_error or e}"}, status_code=503)
    hit: dict[str, Any] = {"lat": lat, "lon": lon, "county": None, "urban_area": None, "state": None}
    for r in rows:
        try:
            gt, parts = _parts(json.loads(r["geometry"]))
        except Exception:  # noqa: BLE001
            continue
        if gt == 3 and _point_in_rings(parts, lat, lon):
            hit[r["kind"]] = {"name": r["name"], "code": r["code"], "state": r["state"]}
    return JSONResponse(hit)


# ------------------------------------------------------ static site + tiles
_RANGE = re.compile(r"bytes=(\d*)-(\d*)$")


@app.api_route("/tiles/{name}", methods=["GET", "HEAD"])
def tiles(name: str, request: Request) -> Response:
    """PMTiles with HTTP Range support (protomaps reads the archive by byte
    ranges - a server that ignores Range would ship the whole 50-70 MB file
    per tile request). Served from RR_TILES_DIR, which may point at a Unity
    Catalog Volume mount instead of the app source tree."""
    if not re.fullmatch(r"[a-z0-9_-]+\.pmtiles", name):
        raise HTTPException(404)
    path = TILES_DIR / name
    if not path.is_file():
        raise HTTPException(404)
    size = path.stat().st_size
    headers = {"Accept-Ranges": "bytes", "Cache-Control": "public, max-age=86400"}
    if request.method == "HEAD":  # the page probes each tileset with HEAD before using it
        return Response(status_code=200, headers={**headers, "Content-Length": str(size)})
    rng = request.headers.get("range")
    m = _RANGE.match(rng.strip()) if rng else None
    if not m:
        return FileResponse(path, media_type="application/octet-stream", headers=headers)
    a, b = m.group(1), m.group(2)
    if a == "" and b == "":
        raise HTTPException(416)
    if a == "":  # suffix range
        length = min(int(b), size)
        start, end = size - length, size - 1
    else:
        start = int(a)
        end = min(int(b), size - 1) if b else size - 1
    if start >= size or start > end:
        return Response(status_code=416, headers={**headers, "Content-Range": f"bytes */{size}"})
    with path.open("rb") as f:
        f.seek(start)
        data = f.read(end - start + 1)
    return Response(
        data, status_code=206, media_type="application/octet-stream",
        headers={**headers, "Content-Range": f"bytes {start}-{end}/{size}", "Content-Length": str(len(data))},
    )


@app.get("/")
def index() -> FileResponse:
    return FileResponse(WEB_DIR / "index.html", media_type="text/html")


app.mount("/", StaticFiles(directory=str(WEB_DIR), html=True), name="web")
