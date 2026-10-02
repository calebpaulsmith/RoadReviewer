# AGOL notebook version — ARCHIVED 2026-10-01, do not update

A Python port of the road classifier meant to run as an ArcGIS Online
notebook and be published as a Web Tool for Experience Builder.

**Frozen per Caleb (2026-10-01): the org has no web tools, so this version
gets no further fixes.** It is kept only as a reference.

What it has and has not got, as of the freeze:

- Has the PR #24 closest-road verdict model and the PR #47 tie rule
  ("Review - Conflicting classes").
- Does NOT have the double-check, the route-id comparison, the HPMS
  non-inventory twin drop, or the HPMS outage fallback (CLAUDE.md §9.7a/§9.7b).
- Its service URLs are whatever was current on 2026-09-30; nobody will move
  them when a state moves a layer.

Files:

- `roadreviewer_classify.ipynb` / `roadreviewer_classify.py` — the notebook
  and its jupytext mirror (they lived in `notebooks/`).
- `notebook-web-tool-implementation.md` — the publishing guide (it lived in
  `docs/`; its paths still say `notebooks/`).

The live products are the two Excel workbooks and the web tool under `web/`.
