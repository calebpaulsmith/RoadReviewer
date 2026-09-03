#!/usr/bin/env python3
"""Assemble the self-contained Cat C inspection filler.

Inlines vendor/pdf-lib.min.js and the base64-encoded blank form template
into form_src.html's /*PDFLIB*/ and /*PDFB64*/ placeholders, writing the
single-file app to web/inspection/index.html. That one output file is both
the GitHub Pages deployment and the email-around artifact — they are the
same bytes on purpose.

Run from anywhere: python3 build/inspection-filler/build.py
"""
import base64
import pathlib

HERE = pathlib.Path(__file__).resolve().parent
REPO = HERE.parent.parent
OUT = REPO / "web" / "inspection" / "index.html"

src = (HERE / "form_src.html").read_text()
lib = (HERE / "vendor" / "pdf-lib.min.js").read_text()
pdf_b64 = base64.b64encode(
    (HERE / "template" / "catc-road-lwc-fillable-clean.pdf").read_bytes()
).decode()
s123_map = (HERE / "survey123" / "survey123_map.json").read_text()

for marker in ("/*PDFLIB*/", "/*PDFB64*/", "/*S123MAP*/null"):
    if src.count(marker) != 1:
        raise SystemExit(f"expected exactly one {marker} in form_src.html")

out = (src.replace("/*PDFLIB*/", lib)
          .replace("/*PDFB64*/", pdf_b64)
          .replace("/*S123MAP*/null", s123_map))
OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(out)
print(f"wrote {OUT} ({len(out):,} bytes)")
