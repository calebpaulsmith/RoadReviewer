#!/usr/bin/env bash
# Build one state's HPMS class tileset: fetch + tippecanoe -> web/tiles/<abbr>.pmtiles
#   bash build/tiles/build-state-tiles.sh <STATE_FIPS> <abbr>
# e.g. bash build/tiles/build-state-tiles.sh 26 mi
#
# Progressive class-by-zoom display is applied with a -j $zoom filter:
# interstates/freeways always (layer minzoom 6), other principal arterials
# from z8, minor arterials z9, major collectors z10, minor collectors z11,
# everything (locals) z12. Do NOT switch this to per-feature
# "tippecanoe":{"minzoom"} — that key silently rate-drops line features
# even at maxzoom (tippecanoe 2.49, "dropped_by_rate"; measured: Michigan
# collapsed from 349k features to one line per tile).
set -euo pipefail
FIPS="${1:?state FIPS, e.g. 26}"
ABBR="${2:?state abbr lowercase, e.g. mi}"
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
WORK="${TILES_WORKDIR:-$(mktemp -d)}"
NDJSON="$WORK/$ABBR.ndjson"

CLASS_ZOOM_FILTER='{ "roads": [ "any",
  [ "<=", "F", 2 ],
  [ "all", [ ">=", "$zoom", 8 ],  [ "<=", "F", 3 ] ],
  [ "all", [ ">=", "$zoom", 9 ],  [ "<=", "F", 4 ] ],
  [ "all", [ ">=", "$zoom", 10 ], [ "<=", "F", 5 ] ],
  [ "all", [ ">=", "$zoom", 11 ], [ "<=", "F", 6 ] ],
  [ ">=", "$zoom", 12 ] ] }'

node "$HERE/fetch-hpms-state.mjs" "$FIPS" "$NDJSON"
mkdir -p "$REPO/web/tiles"
tippecanoe -o "$REPO/web/tiles/$ABBR.pmtiles" -l roads -Z6 -z13 -P \
  --drop-densest-as-needed -j "$CLASS_ZOOM_FILTER" --force "$NDJSON"
ls -la "$REPO/web/tiles/$ABBR.pmtiles"
