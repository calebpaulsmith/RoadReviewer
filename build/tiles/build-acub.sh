#!/usr/bin/env bash
# Build the cached urban-boundary tileset: web/tiles/acub.pmtiles
#   bash build/tiles/build-acub.sh
# Harvests the Region V polygons from the NTAD 2020 Adjusted Urban Areas
# service (two-phase: ids, then small objectId batches with ~3 m
# generalization — see fetch-acub.mjs for why) and bakes them to z12.
# The cached-verdict path reads this for urban/rural + the ACUB name;
# refresh it together with the HPMS state tilesets.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
WORK="${TILES_WORKDIR:-$(mktemp -d)}"

node "$HERE/fetch-acub.mjs" "$WORK/acub.ndjson"
mkdir -p "$REPO/web/tiles"
tippecanoe -o "$REPO/web/tiles/acub.pmtiles" -l acub -Z0 -z12 -P \
  --no-tile-size-limit --force "$WORK/acub.ndjson"
ls -la "$REPO/web/tiles/acub.pmtiles"
