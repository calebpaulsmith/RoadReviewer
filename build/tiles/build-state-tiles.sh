#!/usr/bin/env bash
# Build one state's HPMS class tileset: fetch + tippecanoe -> web/tiles/<abbr>.pmtiles
#   bash build/tiles/build-state-tiles.sh <STATE_FIPS> <abbr>
# e.g. bash build/tiles/build-state-tiles.sh 26 mi
set -euo pipefail
FIPS="${1:?state FIPS, e.g. 26}"
ABBR="${2:?state abbr lowercase, e.g. mi}"
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
WORK="${TILES_WORKDIR:-$(mktemp -d)}"
NDJSON="$WORK/$ABBR.ndjson"

node "$HERE/fetch-hpms-state.mjs" "$FIPS" "$NDJSON"
mkdir -p "$REPO/web/tiles"
tippecanoe -o "$REPO/web/tiles/$ABBR.pmtiles" -l roads -Z6 -z13 -P \
  --drop-densest-as-needed --extend-zooms-if-still-dropping --force "$NDJSON"
ls -la "$REPO/web/tiles/$ABBR.pmtiles"
