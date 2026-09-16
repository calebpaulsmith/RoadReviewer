# FHWA Road Checker — manual test pass

Hand tests for the web tool (`web/index.html`), covering what the headless
verifiers in `build/web-tests/` cannot judge: the clipboard, Excel, Google
Earth, PDF layout on a real machine, and the editable export table.

Interactive version (tick + notes, results readable by Claude):
<https://claude.ai/artifact/KMuEX8xpojeYZ36aBrvSnG>

Test ids are stable — quote them ("5.4 failed") when reporting back.

## Test data

Paste these three lines into Coordinate Input. **Line 2 is tab-separated on
purpose** — several tests check that editing a coordinate leaves the user's
own formatting alone.

```
Kalamazoo culvert,42.28536,-85.57025
Site B	42.6911	-84.5360
Tawas rural, 44.2700, -83.5200
```

Expected: Kalamazoo = Federal aid (Urban Minor Collector) · Site B =
Non-federal aid (Urban Local, Lansing) · Tawas = Non-federal aid (Rural Local).

## 1. Auto-Detect panel (~5 min)

- [ ] **1.1** Paste the three test lines. → Three rows classify: Kalamazoo red,
  Site B green (Urban Local), Tawas green (Rural Local); pins match the rows.
- [ ] **1.2** Read the heading and the line under it. → Heading reads
  "Auto-Detect"; the disclaimer is legible and names the responsible governing
  agency.
- [ ] **1.3** Hover the ⓘ beside Detection Buffer. → Tooltip explains: closest
  road decides, a federal-aid road inside the buffer can only downgrade to
  yellow, ACUB never narrows below 250 ft.
- [ ] **1.4** Set the buffer to 50 ft, then 1,000 ft. → Everything re-checks each
  time; fewer roads at 50 ft, more at 1,000 ft and possibly a yellow "Nearby
  FHWA road".

## 2. Results list ↔ map filtering (~4 min)

- [ ] **2.1** Filter rows on `Kalamazoo`, then clear. → Only that row **and only
  its pin**; clearing restores all three of each.
- [ ] **2.2** "filter by map view" on a fresh load. → Unchecked.
- [ ] **2.3** Unchecked, zoom onto Kalamazoo. → All three rows stay listed.
- [ ] **2.4** Tick it, then pan. → Only sites in view; "showing 1 of 3"; the list
  follows the map.
- [ ] **2.5** Still ticked, click a row. → Map zooms to the site, the list does
  **not** collapse to it.
- [ ] **2.6** Still ticked, step with Prev/Next. → Same.
- [ ] **2.7** Untick. → All rows return.

## 3. Export pill and menu (~4 min)

- [ ] **3.1** Locate Export; switch tabs. → Blue split pill with the coordinate
  input (not under Auto-Detect), reachable from both tabs.
- [ ] **3.2** Click the caret, then click away. → Grouped menu opens, closes on
  the outside click.
- [ ] **3.3** "Copy site + coordinates" → paste into Excel. → Toast; three
  columns, one row per site, each value in its own cell.
- [ ] **3.4** "Copy Auto-Detect results" → paste into Excel. → Every column in
  its own cell with the header row.
- [ ] **3.5** Each format entry. → Dialogue opens on that format's tab.

## 4. Export dialogue and the files it writes (~8 min)

- [ ] **4.1** Click through the four tabs. → Each shows its own actions and a
  preview of real content.
- [ ] **4.2** CSV in Excel. → Columns align; road names with commas stay in one
  cell; Note column present.
- [ ] **4.3** KMZ in Google Earth. → Red/green/yellow pins by verdict; pin
  description carries status, class, urban area, roads, State Route ID + MP, note.
- [ ] **4.4** GeoJSON dragged onto the AGOL map. → Three points, same attributes.
- [ ] **4.5** PDF Report at 500 ft. → Cover + one page per site; nothing clipped
  or bleeding off the page.
- [ ] **4.6** PDF Report at 3 mi. → Same layout quality, wider ground. *Flag
  anything off — this control is the one up for redesign.*
- [ ] **4.7** FIRMettes (ZIP). → One FEMA FIRMette PDF per site.

## 5. Editing in the dialogue (~10 min) — newest work

- [ ] **5.1** Note on row 1 → CSV and KMZ. → Present in both.
- [ ] **5.2** Row 1 Longitude → `-84.5360`, click outside the row. → First line of
  the coordinates box updates with **only the longitude changed**; pin moves;
  verdict re-runs.
- [ ] **5.3** Edit the Latitude of the tab-separated Site B line. → The line keeps
  its tabs.
- [ ] **5.4** Latitude → **Tab** → Longitude → click outside the row. → Nothing on
  the Tab; leaving the row applies both and re-checks **once**.
- [ ] **5.5** `abc` or `99` in Latitude. → Cell reverts, toast gives the range,
  nothing exports with the bad value.
- [ ] **5.6** Change a Site Name. → Line rewritten with the new name; row and map
  label follow.
- [ ] **5.7** Watch during a coordinate re-check. → Export buttons grey out with
  "re-checking the edited site…", then return.
- [ ] **5.8** Start an edit, press Escape. → Dialogue closes and the edit applies.
- [ ] **5.9** Collected point: add with name + note, edit coords, reload. → Comes
  back edited.
- [ ] **5.10** Collected point on the same coordinates as a pasted one, different
  notes. → Notes stay separate.

## 6. Regression sanity (~6 min)

- [ ] **6.1** Fresh load, then zoom in. → Atlas map with no class lines at the
  region view; class lines ~z9-10; street detail from z13.
- [ ] **6.2** Zoom region → street and back several times. → No "Page
  Unresponsive" (**known open item**, never reproduced headless).
- [ ] **6.3** Tick "Live verdicts", re-run the three points. → Same verdicts;
  rows lose the "cached data" chip.
- [ ] **6.4** One point per state from `docs/Region V Test Coordinates.xlsx`. →
  Each matches the sheet.
- [ ] **6.5** "⧉ Expand table". → Every site with full detail incl. the State
  route chip.

## 7. Addresses and road names (~8 min) — newest work

Paste, on their own lines under the three coordinates:

```
5201 Portage Rd, Portage MI 49002
Portage Rd, Portage MI
1 Nowhere Ln, Nowhere MI
```

- [ ] **7.1** Watch the box as you paste. → The two Portage lines highlight
  (yellow = address, purple → green = road once it resolves); nothing is sent
  until you click (network log shows no `geocoding.geo.census.gov`).
- [ ] **7.2** The count line. → "3 point(s) parsed, 2 addresses, 1 road".
- [ ] **7.3** `Portage Rd, Portage MI` without any click. → A row appears with a
  **whole road** chip and a class-by-length chip; the road draws on the map.
- [ ] **7.4** Click **Geocode 2 addresses**. → Button counts down, then a toast
  "1 of 2 addresses located"; `5201 Portage Rd` becomes a row with **from
  address · snapped to Portage Rd**; `1 Nowhere Ln` shows "No match" and the
  button now reads **Geocode 1 address**.
- [ ] **7.5** Zoom to the 5201 pin. → It sits **on** Portage Rd's centerline,
  not on the cross street; the verdict is Portage Rd's class.
- [ ] **7.6** The coordinates box after geocoding. → Every line exactly as you
  typed it — no coordinates were written into it.
- [ ] **7.7** Edit the ZIP on the 5201 line. → The line turns yellow again and
  the button offers it again (edits make a line unresolved).
- [ ] **7.8** Export → Excel preview. → The 5201 row's Latitude/Longitude are the
  snapped coordinates and its Site Name is the geocoder's matched address.
- [ ] **7.9** Try a road with no place, e.g. `Q Ave`. → Marked unreadable, with
  no lookup.
- [ ] **7.10** Reload the page and re-paste. → Addresses are unresolved again
  (results are not stored) until you click.


## 8. Pointer coordinate readout (~3 min) — newest work

- [ ] **8.1** Fresh load, don't touch the mouse. → A box in the map's
  bottom-right corner, above the attribution line, showing the centre
  coordinate and nothing else (no label).
- [ ] **8.2** Move over the map. → The value follows the pointer continuously;
  motion stays smooth at street zoom with imagery on.
- [ ] **8.3** Park the pointer on a known feature and compare against the pin
  of a site you already classified. → Same coordinate to ~5 decimals.
- [ ] **8.4** Hold the pointer still and drag the map with the keyboard arrows
  (or scroll-zoom). → The value updates as the ground moves under the cursor.
- [ ] **8.5** Move off the map (over the left panel). → Falls back to the map
  centre.
- [ ] **8.6** Move onto the coordinate, then click it. → The value **freezes**
  as you approach it, and the click copies exactly that value ("Coordinate
  copied to clipboard"). Paste into the coordinates box → a point at that spot.
- [ ] **8.7** Click the small **DD** box beside it. → The coordinate becomes
  `42°17'07.3"N 85°34'12.9"W` and the box now reads **DMS**. Click the
  coordinate again → the DMS text copies, and it pastes back to the same point.
- [ ] **8.8** Reload the page. → It comes back in whichever format you left it.
- [ ] **8.9** With a site legend showing, check the corner. → Legend sits above
  the readout; neither covers the attribution.
- [ ] **8.10** On a phone. → The boxes show the map centre and update as you
  pan; tapping DD/DMS still switches format, and they don't crowd the
  attribution.
