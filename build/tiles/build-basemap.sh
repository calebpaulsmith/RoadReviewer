#!/usr/bin/env bash
# Build the offline Region V road BASEMAP: web/tiles/basemap.pmtiles
#   bash build/tiles/build-basemap.sh [YYYYMMDD]
# (build date defaults to the latest daily Protomaps build; list them at
#  https://build-metadata.protomaps.dev/builds.json)
#
# Pipeline: pmtiles extract (six-state region polygon, maxzoom 11) ->
# python pmtiles-convert to mbtiles -> filter-basemap-layers.py (drops
# buildings/POIs/landuse at the protobuf wire level; also gunzip-verifies
# every tile) -> pmtiles convert back. Result measured 2026-09-14: 50 MB
# for all six states, z0-11 (protomaps-leaflet overzooms beyond 11; the
# HPMS class tiles carry every road from z12 anyway).
#
# Needs: the go-pmtiles CLI (`pmtiles`) on PATH, `pip install pmtiles`
# (for pmtiles-convert), python3. maxzoom 12 was rejected: ~135 MB even
# filtered (GitHub's limit is 100 MB/file) for detail the HPMS overlay
# already provides.
set -euo pipefail
BUILD="${1:-$(curl -sS https://build-metadata.protomaps.dev/builds.json | python3 -c 'import json,sys;print(json.load(sys.stdin)[-1]["key"])')}"
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
WORK="${TILES_WORKDIR:-$(mktemp -d)}"

python3 - "$WORK/r5-region.geojson" <<'EOF'
import json, sys
boxes = {   # one generous rectangle per Region V state (fetch-hpms-state.mjs STATE_BOX)
 "MI": [-90.5, 41.6, -82.0, 48.4], "WI": [-93.0, 42.4, -86.1, 47.4],
 "MN": [-97.3, 43.4, -89.4, 49.5], "IL": [-91.6, 36.9, -87.0, 42.6],
 "IN": [-88.2, 37.7, -84.7, 41.8], "OH": [-84.9, 38.3, -80.4, 42.0],
}
polys = [[[[x0,y0],[x1,y0],[x1,y1],[x0,y1],[x0,y0]]] for x0,y0,x1,y1 in boxes.values()]
json.dump({"type":"Feature","properties":{},
           "geometry":{"type":"MultiPolygon","coordinates":polys}}, open(sys.argv[1],"w"))
EOF

pmtiles extract "https://build.protomaps.com/$BUILD" "$WORK/r5-full.pmtiles" \
  --region="$WORK/r5-region.geojson" --maxzoom=11
pmtiles-convert "$WORK/r5-full.pmtiles" "$WORK/r5-full.mbtiles"
python3 "$HERE/filter-basemap-layers.py" "$WORK/r5-full.mbtiles" "$WORK/r5-roadmap.mbtiles"
mkdir -p "$REPO/web/tiles"
pmtiles convert "$WORK/r5-roadmap.mbtiles" "$REPO/web/tiles/basemap.pmtiles"
ls -la "$REPO/web/tiles/basemap.pmtiles"
