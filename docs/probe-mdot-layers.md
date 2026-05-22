# Probing MDOT NFC + NTAD ACUB layers (verification step §5.1)

## Why this doc exists

The build environment for this repo (Claude Code on the web) cannot reach
`mdotgis.state.mi.us` or `services.arcgis.com` — both return HTTP 403
from the sandbox's egress allowlist. Before the VBA macro can be wired
to query MDOT's National Functional Classification layers, the **exact
field names** that come back from those layers must be confirmed against
the live service. This doc walks you through doing that from your
workstation, which has normal internet access.

It is a one-time step. Run it once, paste the answers into
`CLAUDE.md` §4.2, and §8 open question #1 closes.

## What we need to learn

For each of these three MDOT FeatureServer layers, get the list of
fields, identify which one carries the **NFC class code** (e.g. "Urban
Local"), and identify which one carries the **road / route name**:

- `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353` — "Functional System"
- `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/364` — "Classification"
- `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/543` — "Route System"

And from the nationwide ACUB layer, get the urban-area-name field
(probably `NAME` or `NAME20`):

- `https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0`

## Note on "MCP"

You do **not** need an MCP server for this. The GitHub MCP the cloud
session uses is for pushing code to GitHub; it has no concept of
arbitrary HTTP requests. The right tool here is Claude Code's built-in
`WebFetch` (Option A) or plain Python's `urllib` (Option B). Both work
from any machine with internet access — no install required.

## Option A — drive a local Claude Code session (recommended)

1. Clone (or `git pull` your existing clone of) this repo on your
   workstation.
2. Open a terminal in the repo root.
3. Run `claude` to start a local Claude Code session.
4. Paste the prompt below verbatim and hit enter.

The session will fetch the URLs, propose an edit to `CLAUDE.md` §4.2
with the confirmed schemas, and (if you approve) commit + push it.

```text
I need you to run verification step §5.1 of the RoadReviewer design doc.
Read CLAUDE.md §4.2 first for full context. Then:

1. Use WebFetch on each of these four URLs and report back the schemas
   verbatim — for every field give name, type, and alias:

   - https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353?f=pjson
   - https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/364?f=pjson
   - https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/543?f=pjson
   - https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0?f=pjson

2. For each MDOT layer, run a point-intersect query against downtown
   Detroit (lon=-83.045, lat=42.331) to see which fields actually carry
   data. The query template (substitute the layer ID for {ID}):

   https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/{ID}/query?geometry=-83.045,42.331&geometryType=esriGeometryPoint&inSR=4326&spatialRel=esriSpatialRelIntersects&outFields=*&returnGeometry=false&f=json

3. For the NTAD ACUB layer, run the same point query against Detroit and
   confirm the urban-area record is returned. Report which field holds
   the urban-area name.

4. Identify three "known coordinate" test cases for §5.1 of CLAUDE.md:
   - one Urban Minor Collector inside an ACUB (expected INELIGIBLE)
   - one Rural Local inside an ACUB (expected eligible, urban)
   - one Rural Local well outside any ACUB (expected eligible, rural)
   Look them up using the live service, not from memory.

5. Update CLAUDE.md §4.2 in place: fill in the confirmed field names,
   remove the "TBD" annotations, and add the three test coordinates.
   Commit with the message
     "Confirm MDOT NFC + NTAD ACUB schemas (verification §5.1)"
   and push to the current branch.

Keep narration tight. If any layer is unreachable, report what you saw
and stop — do not paper over a failure.
```

## Option B — run the Python probe directly

If you'd rather skip Claude in the loop and just look at the JSON
yourself, run the included script:

```bash
python3 docs/probe.py
```

It hits the four FeatureServers, prints the schemas, runs the Detroit
point-intersect query against each, and prints the matching attributes.
Paste the relevant pieces into `CLAUDE.md` §4.2 by hand.

## After you have the schemas

The exact spots in `CLAUDE.md` §4.2 that need to be filled in:

- **NFC class-code field name** (under the "Functional System" / 353
  bullet — currently TBD).
- **Road-name field name** (under "Route System" / 543, or wherever the
  Detroit query shows the road name attribute — currently TBD).
- **NTAD ACUB urban-area-name field** — replace "expected `NAME` or
  `NAME20`" with whatever the live `?f=pjson` actually returns.
- **Three known test coordinates** for verification step §5.1.

Once those are filled in, §8 open question #1 closes and the VBA
implementation can start.
