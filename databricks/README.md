# FHWA Road Checker as a Databricks App

The web tool in `web/` (the same page GitHub Pages serves) packaged to run
inside Databricks, with its road-class and boundary data in **two Unity
Catalog Delta tables**:

| table | contents |
|---|---|
| `reviewer_roads` | every FHWA HPMS road segment in MI / IN / WI / MN / IL / OH: `f_system` (FHWA class 1-7), the state's LRS keys (`route_id`, `begin_point`, `end_point`, `route_name`, `route_number`), bbox columns, GeoJSON geometry. Partitioned by `state`, Z-ordered on `min_lon, min_lat`. |
| `reviewer_boundaries` | `kind` = `urban_area` (2020 Adjusted Urban Areas / ACUB), `county` (Census TIGER counties) and `state` (the six outlines), with `name`, `code`, bbox, GeoJSON geometry. |

That is the whole data model. The page's cached-verdict path (the same
`computeVerdict` / `classLabels` code the public site runs against the
PMTiles) reads these tables through `app.py`'s `/api/features` endpoint,
so verdicts, review reasons, road names, the LRS chips and every export are
identical to the public site - only the data source changes, and each row's
source note says "Databricks Delta tables". Everything else in the page
still works as on GitHub Pages: the Live Review button (live state
DOT queries), the Census TIGERweb search, address geocoding, FIRMettes,
PDF / CSV / KMZ / GeoJSON exports, the map (live Esri basemap + live class
layers when the tile archives are not packaged; the app also serves PMTiles
with HTTP Range support when they are).

## Files

```
databricks/
  app.yaml                      Databricks App manifest (command + env)
  app.py                        FastAPI: static site + /tiles range server + /api/*
  requirements.txt
  setup/build_reviewer_tables.py   notebook: harvests HPMS + NTAD ACUB + TIGER counties/states
                                   into the two reviewer_* tables (run once, ~30-60 min)
  package.ps1                   stages web/ + the files above into dist/RoadReviewer-databricks-app.zip
```

## Deploy (first time)

1. **Build the tables.** Workspace > Import > the file
   `setup/build_reviewer_tables.py` as a notebook. Set the `catalog` /
   `schema` widgets (defaults `main.roadreviewer`) and Run all on a cluster
   with internet access. Re-runs replace, never duplicate.
2. **Package the app.** On the laptop:
   ```powershell
   .\databricks\package.ps1            # dist\RoadReviewer-databricks-app.zip (~2 MB, no tiles)
   ```
   The package deliberately leaves out `web/tiles/` (350 MB of PMTiles that
   workspace files can't hold). Nothing is lost: verdicts come from the Delta
   tables, and the page falls back automatically to live map layers for
   display - Esri street tiles for the basemap, the state DOT class layers
   and the NTAD urban boundaries for the overlays (the same fallbacks the
   public site uses when a tileset is missing). `-WithTiles` includes them
   if a workspace ever can take the size. Upload the unzipped folder
   (Workspace > Import, or `databricks sync dist\RoadReviewer-databricks-app
   /Workspace/Users/<you>/road-reviewer`).
3. **Create the app** (Compute > Apps > Create, or
   `databricks apps create road-reviewer`), add a **SQL warehouse resource
   named `sql-warehouse`** (that name is what `app.yaml` reads), and set
   `RR_CATALOG` / `RR_SCHEMA` in `app.yaml` if you used other names.
4. **Grant the app's service principal** read access:
   ```sql
   GRANT USE CATALOG ON CATALOG main TO `<app service principal>`;
   GRANT USE SCHEMA  ON SCHEMA  main.roadreviewer TO `<app service principal>`;
   GRANT SELECT ON TABLE main.roadreviewer.reviewer_roads      TO `<app service principal>`;
   GRANT SELECT ON TABLE main.roadreviewer.reviewer_boundaries TO `<app service principal>`;
   ```
   plus CAN USE on the warehouse.
5. **Deploy** from the uploaded folder (`databricks apps deploy road-reviewer
   --source-code-path /Workspace/Users/<you>/road-reviewer`). Open the app
   URL, then `<app url>/api/health` should read `"delta": true` with the row
   counts; if it says `false`, the `reason` field names the binding or grant
   that is missing and the page keeps working from the tiles meanwhile.

Updating the site later = re-run `package.ps1`, re-sync, re-deploy.
Refreshing the data = re-run the notebook (HPMS publishes annually).

## API

| route | returns |
|---|---|
| `GET /api/health` | `{ok, delta, catalog, schema, tables, roads, boundaries}` |
| `GET /api/features?kind=roads&state=MI&lat=&lon=&ft=250` | HPMS segments whose bbox is within `ft` feet, as `{props:{F,R,B,E,N,RN}, geomType:2, parts:[[[lon,lat],…]]}` |
| `GET /api/features?kind=acub&lat=&lon=&ft=250` | urban-area polygons near the point (`props:{NAME,UACE,state_1}`) |
| `GET /api/features?kind=county&lat=&lon=&ft=250` | county polygons near the point (`props:{NAME,GEOID,STATE}`) |
| `GET /api/area?lat=&lon=` | the county, urban area and state containing the point |

## Local check (no Databricks)

```powershell
cd databricks
$env:RR_WEB_DIR = "..\web"; uvicorn app:app --port 8000
```
`http://localhost:8000/` serves the site from the tiles; `/api/health`
reports `delta:false` with the reason (no warehouse bound).
