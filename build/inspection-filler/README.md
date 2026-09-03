# Cat C Road / Low Water Crossing — Site Inspection filler

A self-contained, single-file web app that fills the FEMA Public
Assistance **"Cat C - Road-Low Water Crossing - Fillable.pdf"** Site
Inspection Report entirely in the browser. No server, no install, no
network: the pdf-lib library and the blank form (base64) are inlined into
one HTML file. The output is the *genuine* FEMA PDF with its AcroForm
fields still live — not a lookalike.

## The deliverable

`web/inspection/index.html` (~1 MB) is both:

- **hosted** — deployed by `.github/workflows/pages.yml` with the rest of
  `web/`, so it's reachable at the repo's GitHub Pages URL under
  `/inspection/`. Served over https, the "Use current location" GPS
  button works and phones get the page fresh on every visit; and
- **portable** — the same file can be downloaded and emailed / dropped in
  Teams / carried on a USB stick, and opened from local storage with no
  connectivity at all. (Opened as a local file, browsers block the
  geolocation API; coordinates then come from geotagged photos or typing.)

## What it does

- Every fillable field of the 7-page form, grouped into phone-friendly
  sections (applicant, GPS, facility, description + cross-section
  annotations, four component-damage site blocks, photos, mitigation,
  insurance, EHP questions 1–8, additional notes).
- **Signature & initials pads** — finger/stylus drawn; the signature is
  stamped on the page-1 Applicant Representative line, the initials on
  pages 2–7. Recipient lines stay blank for later signing.
- **Sketch** — draw it or attach an image; lands in the page-2 sketch box.
- **Photos** — first photo fills the page-4 grid; extras are appended as
  additional pages (two per page, captioned) with continued page numbers.
  Photos are downscaled/re-encoded to keep the PDF small.
- **GPS** — geolocation button (https only), plus automatic fallback: the
  first geotagged photo's EXIF coordinates fill the start point.
- **Drafts** — text/checkbox state autosaves to localStorage (best-effort)
  and can be exported/imported as a JSON file to move between devices.
  Photos and drawings are not part of drafts.
- **Page numbers** — the "Page _ of _" blanks fill automatically,
  accounting for appended photo pages.
- Output goes through the native share sheet where available (iOS/Android
  → Files/Mail/Teams), else a browser download. Filename derives from the
  Damage #.

## Build

```
python3 build/inspection-filler/build.py
```

inlines `vendor/pdf-lib.min.js` (pdf-lib 1.17.1 UMD, from cdnjs) and
`template/catc-road-lwc-fillable-clean.pdf` into `form_src.html`'s
placeholders and writes `web/inspection/index.html`. Rebuild and commit
the output whenever `form_src.html` or the template changes.

`template/catc-road-lwc-fillable-clean.pdf` is the original FEMA form
with two invisible fixes pdf-lib needs (rich-text flags cleared; 10pt
/DA on text fields so big notes fields don't auto-size comically).
Regenerate it from an original with
`python3 make_template.py <original.pdf>` (needs pypdf; pikepdf optional
for compression).

## Verify

```
cd build/web-tests && npm install && node verify-inspection-filler.mjs
```

Drives the built page in Chromium from `file://`: imports a draft
covering every section (including the form's typo'd field names like
`Location Addres11`), draws on all three pads (ink asserted on canvas),
attaches a geotagged photo fixture twice, generates, and then re-opens
the output PDF asserting field values, checkbox states, page count,
image XObjects on the stamped pages, page-number auto-fill, and the
EXIF GPS auto-fill.

## Field mapping notes

Field names in `form_src.html`'s spec are verbatim from the PDF,
including its typos (`Location Addres11`, `Quantity 11`,
`%Complete9` vs `% Complete 10`). Checkbox purposes were mapped by
widget coordinates against rendered pages. The `Work Order` / `Damage`
fields in the "For FEMA Use Only" page headers are only filled when the
"Repeat … in the page headers" toggle is on. The 14 `/Sig` digital
signature fields are left untouched (the pads stamp images over them);
true digital signing isn't possible in-browser.

## Survey123 companion

`survey123/gen_xlsform.py` generates two files from one definition (so
they can't drift): `catc_road_lwc_xlsform.xlsx`, an XLSForm to publish
via Survey123 Connect so crews can collect in Esri's field app (offline,
GPS geopoint, photos, signature), and `survey123_map.json`, the
Survey123-name → PDF-field-name mapping that build.py inlines into the
web app.

The round trip needs no webhooks, credits, or extra licensing:

1. Publish the XLSForm; inspectors submit from the Survey123 app.
2. In the Survey123 website's **Data** tab, export responses as **CSV**.
3. In the web app, **Import Survey123 CSV** — single-record exports
   import directly; multi-record exports show a picker. Dates are
   reformatted, select-ones map to the right checkbox, select-multiples
   fan out, and the geopoint's x/y columns fill the GPS start fields.
4. Re-attach photos and sign (attachments don't travel in CSV), generate.

Survey123 has no built-in "email me each response" — that requires a
webhook into an automation platform (e.g. Power Automate's Survey123
connector). The CSV export path above is the zero-infrastructure
equivalent.

## Extending to the other report types

The other Site Inspection Report PDFs (Cat A/D/E/F/G, bridges, culverts,
component sheets) follow the same recipe: run `make_template.py` on the
blank form, map its field names into a new spec block, reuse everything
else. The app shell (pads, photos, EXIF, drafts, share) is
form-agnostic.
