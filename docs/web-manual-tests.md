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
- [ ] **2.8** Paste the three coordinates fresh. → A three-bar chart (Federal
  aid / Needs review / Non-federal aid) appears at once under Auto-Detect and
  its counts tick up as verdicts land. Click a bar → only those rows and pins
  stay, the bar is highlighted, Prev/Next step through just those; click it
  again → everything returns. Two bars can be on at once.

## 3. Quick Export menu and View and Export (~6 min)

- [ ] **3.1** Locate "Quick Export ▾". → One blue dropdown button with the
  coordinate input (not under Auto-Detect). The pane is ONE column now: search
  box on top, coordinates box, Quick Export, Auto-Detect — no tabs.
- [ ] **3.2** Click it, then click away. → Grouped menu opens, closes on the
  outside click.
- [ ] **3.3** "Copy site + coordinates" → paste into Excel. → Toast; three
  columns, one row per site, each value in its own cell.
- [ ] **3.4** "Copy Auto-Detect results" → paste into Excel. → Every column in
  its own cell with the header row.
- [ ] **3.5** Each download entry. → The file downloads straight away, no
  dialogue.
- [ ] **3.6** Blue "View and Export →" pinned at the bottom of the pane (and
  "⧉ View and Export" in the results header). → The pane stretches across the
  screen (a short slide); the map keeps a strip on the right and zooms to fit
  every site — or stays on the selected site if one was selected.
- [ ] **3.7** Click rows in the big table. → The row gets a blue outline, the
  same row highlights in the small list, the map zooms right in on that site;
  Prev/Next on the map strip still steps. "← Back to map" restores the layout.

## 4. View and Export tabs and the files they write (~8 min)

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

## 5. Editing in View and Export (~10 min)

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
- [ ] **5.8** Start an edit, press Escape. → The pane shrinks back and the edit
  applies.
- [ ] **5.9** Pin button (top right of the map): one click → button turns blue,
  cursor becomes a pin; click the map. → A pin drops with a bounce, its label is
  a name box with "Point N" selected, the line "Point N, lat, lon" is at the end
  of the coordinates box, and the site classifies. Type a name without clicking
  anything → the line and the label follow; Enter ends it. The map must NOT
  jump away while you type. Pin mode is off again after the one pin.
- [ ] **5.10** Click the pin button twice (orange, ∞ badge). → Every map click
  drops another pin; scroll-zoom and drag still work; clicking an existing pin
  selects it rather than dropping a new one; Escape or the button turns it off.
- [ ] **5.11** Right-click a pin → "Move pin", drag it. → Its line in the box
  gets the new coordinates (name kept), the verdict re-runs. Right-click →
  "Delete pin". → The line and the pin are gone.

## 5b. Search, default area and pickers (~8 min) — newest work

- [ ] **5b.1** Type `rockford` in the search box (no Find button). → Results
  grouped under State / County / City / Township / Road headings; "Rockford
  city" (IL, MI, MN) is listed under City, the townships under Township.
- [ ] **5b.2** Click "Kalamazoo County". → Map zooms to the dashed outline
  and a blue chip "Area: Kalamazoo County, MI ✕" appears under the search box;
  reload → still there.
- [ ] **5b.3** With the area set, type `main` in the search. → Roads found
  inside the county even when zoomed out ("in Kalamazoo County, MI").
- [ ] **5b.4** Paste `Q Ave` alone. → Resolves inside the county as a whole
  road. Paste `Main St`. → Row says "2 separate stretches in Kalamazoo County,
  MI — pick one" with township/city-named buttons; a click resolves it and the
  row's chip says which stretch you picked.
- [ ] **5b.5** Click ✕ on the area, paste `Q Ave`. → Row says "Road name
  without a place — pick a county, city or township…"; nothing is searched.
- [ ] **5b.6** Area set, paste `5201 Portage Rd` (no city), click Geocode. →
  Located inside the county. Paste an address that exists in several places
  in the county → "N possible matches inside Kalamazoo County, MI — pick one"
  with address buttons.

## 6. Regression sanity (~6 min)

- [ ] **6.1** Fresh load, then zoom in. → Atlas map with no class lines at the
  region view; class lines ~z9-10; street detail from z13.
- [ ] **6.2** Zoom region → street and back several times. → No "Page
  Unresponsive" (**known open item**, never reproduced headless).
- [ ] **6.3** Tick "Live verdicts", re-run the three points. → Same verdicts;
  rows' Source chip changes from "FHWA HPMS 2024 tiles" to "<DOT> live layer"
  and clicking it opens the state's official public map, not a REST page.
- [ ] **6.4** One point per state from `docs/Region V Test Coordinates.xlsx`. →
  Each matches the sheet.
- [ ] **6.5** Outage: with "Live verdicts" ticked, block a state server (e.g.
  DevTools → Network → block `mdotgis.state.mi.us`) and paste a point there. →
  Verdict still appears; chip reads "Source: FHWA HPMS 2024 tiles · MDOT live
  layer down"; an amber strip under Auto-Detect names the outage and rechecks
  every 60 s ("check now" forces it). Unblock → strip turns green "back up —
  re-run 1 row(s)"; clicking it re-classifies live. With 20+ points pasted the
  Network log shows at most 3 requests in flight per host.
- [ ] **6.6** "⧉ View and Export" from the results header. → Same expanded
  view as the bottom button, every site with the State Route ID column.

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
