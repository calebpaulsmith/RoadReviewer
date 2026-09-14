#!/usr/bin/env python3
"""Strip a Protomaps basemap mbtiles down to road-map layers.

    python3 filter-basemap-layers.py <in.mbtiles> <out.mbtiles>

Keeps only the layers a road basemap needs (earth, water, roads,
boundaries, places) — buildings/POIs/landuse are the bulk of a full
basemap's size. Layers are dropped at the protobuf WIRE level (each MVT
layer is one length-delimited submessage; its name is field 1), so no
feature decoding happens — which also sidesteps tippecanoe 2.49's
tile-join, whose MVT reader chokes on some Protomaps tiles.

Every tile is gunzipped here, so this doubles as a corruption check:
a tile that fails to decompress aborts the run — re-extract (the agent
proxy has been seen truncating range responses, and `pmtiles extract`
does not checksum tiles).
"""
import gzip
import shutil
import sqlite3
import sys

KEEP = {b"earth", b"water", b"roads", b"boundaries", b"places"}


def read_varint(b, i):
    r = s = 0
    while True:
        x = b[i]
        i += 1
        r |= (x & 0x7F) << s
        if not x & 0x80:
            return r, i
        s += 7


def layer_name(b):
    i = 0
    while i < len(b):
        tag, i = read_varint(b, i)
        f, wt = tag >> 3, tag & 7
        if wt == 2:
            ln, i = read_varint(b, i)
            if f == 1:
                return b[i:i + ln]
            i += ln
        elif wt == 0:
            _, i = read_varint(b, i)
        elif wt == 5:
            i += 4
        elif wt == 1:
            i += 8
        else:
            return None
    return None


def filter_tile(raw):
    b = gzip.decompress(raw) if raw[:2] == b"\x1f\x8b" else raw
    out = bytearray()
    i = 0
    while i < len(b):
        start = i
        tag, i = read_varint(b, i)
        f, wt = tag >> 3, tag & 7
        if wt == 2:
            ln, i = read_varint(b, i)
            end = i + ln
            if f == 3:  # Tile.layers
                if layer_name(b[i:end]) in KEEP:
                    out += b[start:end]
            else:
                out += b[start:end]
            i = end
        elif wt == 0:
            _, i = read_varint(b, i)
            out += b[start:i]
        elif wt == 5:
            i += 4
            out += b[start:i]
        elif wt == 1:
            i += 8
            out += b[start:i]
        else:
            raise ValueError("bad wiretype %d at %d" % (wt, start))
    if i != len(b):
        raise ValueError("overrun")
    return gzip.compress(bytes(out), 9)


def main():
    src, dst = sys.argv[1], sys.argv[2]
    shutil.copyfile(src, dst)
    db = sqlite3.connect(dst)
    rows = db.execute(
        "SELECT zoom_level, tile_column, tile_row, tile_data FROM tiles").fetchall()
    print("tiles:", len(rows))
    for z, x, y, data in rows:
        nd = filter_tile(data)   # raises on a corrupt tile: re-extract, don't ship it
        db.execute(
            "UPDATE tiles SET tile_data=? WHERE zoom_level=? AND tile_column=? AND tile_row=?",
            (nd, z, x, y))
    db.commit()
    db.execute("VACUUM")
    db.close()
    print("done:", dst)


if __name__ == "__main__":
    main()
