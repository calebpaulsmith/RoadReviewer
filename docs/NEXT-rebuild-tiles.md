# NEXT: rebuild the five state tilesets (cloud session)

**Status: NOT DONE. Written 2026-10-01 for a later cloud session.**
The Windows laptop cannot do this — it has no tippecanoe, no WSL and no
Docker. A Claude Code cloud session (Linux) can.

## Why

The web tool answers from the HPMS tiles in `web/tiles/<st>.pmtiles`. HPMS
records some roads twice on the same centerline (an inventory record and a
"non-inventory direction" record) and the two can carry different classes.
The page drops the non-inventory twin, but only when the tile carries the
`T` attribute (HPMS `FACILITY_TYPE`). **Only `wi.pmtiles` has `T`** (rebuilt
2026-09-30). `mi`, `in`, `mn`, `il` and `oh` were built before `T` existed.

Without `T` nothing is wrong, just less sharp: a twin tie in those states
shows as `Review - Conflicting classes` instead of resolving to the inventory
record's class. Known twins that can flip the answer (full list:
`docs/hpms-twin-conflicts-2026-09-30.csv`): MN 3, OH 2, MI / IN / IL 0. So the
states that actually gain from the rebuild are **MN and OH**; the other three
are for consistency.

WI should be rebuilt too if this is done after a new HPMS release (see
"Refresh cadence" in `build/tiles/README.md`), otherwise leave it.

## Steps

```sh
# 0. tools (cloud sandbox)
apt-get update && apt-get install -y tippecanoe     # needs >= 2.x
node --version                                       # >= 18 (global fetch)

# 1. one state at a time: FIPS + lowercase abbreviation
bash build/tiles/build-state-tiles.sh 27 mn
bash build/tiles/build-state-tiles.sh 39 oh
bash build/tiles/build-state-tiles.sh 26 mi
bash build/tiles/build-state-tiles.sh 18 in
bash build/tiles/build-state-tiles.sh 17 il
```

Each run harvests the state from the HPMS service (10–40 minutes; WI, the
largest, took about an hour) and then runs tippecanoe. Set
`TILES_WORKDIR=/some/dir` to keep the intermediate `.ndjson` and its
`.conflicts.csv` (the twin report) instead of a throwaway temp dir.

Expected feature counts (2026-09-30 harvests) — a run far below these lost
cells and must be repeated:

| state | FIPS | features |
|---|---|---|
| MI | 26 | 349,233 |
| IN | 18 | 529,343 |
| MN | 27 | 509,570 |
| IL | 17 | 445,757 |
| OH | 39 | 487,331 |
| WI | 55 | 902,580 |

## Things that go wrong

- **The harvest can stop silently** with no `DONE:` line and no error (seen
  once on OH, 2026-09-30). Check the last line of the output; if there is no
  `DONE: … features`, run that state again. `build-state-tiles.sh` uses
  `set -e`, but a silent exit code 0 would still hand tippecanoe a partial
  file — compare the feature count with the table above before committing.
- **`WARNING: skipped min-size cell`** means the server gave up on a ~500 m
  cell three times. One or two skipped cells in a dense downtown cost a few
  hundred segments; re-run later if it matters.
- **GitHub's 100 MB per-file limit.** `wi.pmtiles` is 70 MB. If a state comes
  out near 100 MB, do not raise `-z`; see "Size constraints" in
  `build/tiles/README.md`.
- **Do not switch the class-by-zoom banding to per-feature
  `tippecanoe.minzoom`** — it silently drops most line features
  (`build-state-tiles.sh` header).
- MDOT and the HPMS service both return occasional HTTP 500s; the harvester
  retries on its own.

## Verify before committing

```sh
cd build/web-tests && npm install        # once
node verify-hpms-tiles.mjs               # cached verdict path, on whichever tileset sorts first
node verify-map-render.mjs               # class + boundary pixels, zoom budget
node ../verify-web-core.mjs
```

Then check that `T` is really in the new tiles and the twin drop acts on it.
Minnesota's twin at `44.028521, -92.648287` (route `0400006595000005`, class 6
on the inventory record, 5 on the other) is the test: paste it into the page
served from the rebuilt tiles. It should read **Non-federal aid - Rural Minor
Collector**, not `Review - Conflicting classes`. MnDOT's own layer says Minor
Collector there too.

A quick look at one tile's attributes, without the page:

```sh
tippecanoe-decode web/tiles/mn.pmtiles 13 1987 2977 | grep -o '"T": *[0-9]' | sort | uniq -c
```

(Any z13 tile in the state works; a tile with no ramps or twins prints
nothing, which is fine — `T` is only written for facility types 3–7.)

## Commit

`web/tiles/*.pmtiles` are committed binaries served by GitHub Pages. One
commit per state keeps a bad rebuild easy to revert:

```sh
git add web/tiles/mn.pmtiles && git commit -m "Rebuild mn.pmtiles with T = FACILITY_TYPE (<n> HPMS segments, <size> MB)"
```

After the last one, update the per-state size table in
`build/tiles/README.md`, change "WI was rebuilt with `T`; the other five
states need the same rebuild" in that file and in CLAUDE.md §4.2b / §9.7b,
and delete this document.
