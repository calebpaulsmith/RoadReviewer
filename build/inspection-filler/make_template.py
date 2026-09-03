#!/usr/bin/env python3
"""Prepare the blank FEMA form for in-browser filling.

Input: the original "Cat C - Road-Low Water Crossing - Fillable.pdf" from
FEMA's Site Inspection Report fillable set. Output:
template/catc-road-lwc-fillable-clean.pdf, which build.py embeds into the
single-file app. Two fixes are applied — both are invisible in Acrobat but
required by pdf-lib (the in-browser fill library):

1. Clear the rich-text field flag (bit 26 of /Ff). Two fields in the
   original carry it and pdf-lib refuses to touch rich-text fields.
2. Give every text field a /DA of "/Helv 10 Tf 0 g" (Helvetica 10pt).
   The original fields have no /DA at all, which pdf-lib treats as
   auto-size — comically large text in the big Notes areas — and its
   setFontSize() errors without a /DA to rewrite.

Usage: python3 make_template.py <original.pdf>
Requires: pypdf (and optionally pikepdf, for a smaller recompressed output).
"""
import pathlib
import sys

from pypdf import PdfWriter
from pypdf.generic import NameObject, NumberObject, TextStringObject

HERE = pathlib.Path(__file__).resolve().parent
OUT = HERE / "template" / "catc-road-lwc-fillable-clean.pdf"
RICH_TEXT_FLAG = 1 << 25
DA = "/Helv 10 Tf 0 g"


def walk(fields, stats):
    for ref in fields:
        obj = ref.get_object()
        flags = int(obj.get("/Ff", 0))
        if flags & RICH_TEXT_FLAG:
            obj[NameObject("/Ff")] = NumberObject(flags & ~RICH_TEXT_FLAG)
            if "/RV" in obj:
                del obj[NameObject("/RV")]
            stats["rich"] += 1
        if obj.get("/FT") == "/Tx":
            obj[NameObject("/DA")] = TextStringObject(DA)
            stats["da"] += 1
        if "/Kids" in obj:
            walk(obj["/Kids"], stats)


def main(src):
    writer = PdfWriter(clone_from=src)
    acro = writer._root_object["/AcroForm"]
    acro[NameObject("/DA")] = TextStringObject(DA)
    stats = {"rich": 0, "da": 0}
    walk(acro["/Fields"], stats)
    tmp = OUT.with_suffix(".tmp.pdf")
    with open(tmp, "wb") as fh:
        writer.write(fh)
    try:  # recompress if pikepdf is available (pypdf clones decompressed)
        import pikepdf

        with pikepdf.open(tmp) as pdf:
            pdf.save(OUT, compress_streams=True,
                     object_stream_mode=pikepdf.ObjectStreamMode.generate)
        tmp.unlink()
    except ImportError:
        tmp.rename(OUT)
    print(f"cleared {stats['rich']} rich-text flags, set /DA on {stats['da']} "
          f"text fields -> {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    main(sys.argv[1])
