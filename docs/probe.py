#!/usr/bin/env python3
"""Probe MDOT NFC + NTAD ACUB FeatureServer schemas (verification step §5.1).

Hits the four FeatureServer endpoints, prints each layer's schema, then
runs a point-intersect query at downtown Detroit so you can see which
fields actually carry useful data.

Run from any machine with internet access (the cloud build sandbox is
firewalled off `mdotgis.state.mi.us` and `services.arcgis.com`):

    python3 docs/probe.py
"""
import json
import sys
import urllib.error
import urllib.request

LAYERS = [
    ("MDOT 353 (Functional System)",
     "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353"),
    ("MDOT 364 (Classification)",
     "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/364"),
    ("MDOT 543 (Route System)",
     "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/543"),
    ("NTAD ACUB (nationwide)",
     "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0"),
]

# Downtown Detroit — should hit every Michigan road layer and be inside an ACUB
TEST_LON, TEST_LAT = -83.045, 42.331


def fetch_json(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": "RoadReviewer/probe"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.load(r)


def main():
    for label, base in LAYERS:
        print(f"\n=== {label} ===")
        print(f"  base: {base}")
        try:
            meta = fetch_json(base + "?f=pjson")
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
            print(f"  metadata fetch failed: {e}")
            continue

        print(f"  name:         {meta.get('name')}")
        print(f"  geometryType: {meta.get('geometryType')}")
        sr = (meta.get("sourceSpatialReference") or meta.get("extent", {}).get("spatialReference") or {})
        print(f"  spatialRef:   {sr.get('wkid') or sr.get('latestWkid')}")
        print(f"  fields:")
        for f in meta.get("fields", []):
            print(f"    - {f.get('name', ''):28s} {f.get('type', ''):22s} {f.get('alias', '')}")

        query_url = (
            f"{base}/query"
            f"?geometry={TEST_LON},{TEST_LAT}"
            f"&geometryType=esriGeometryPoint"
            f"&inSR=4326"
            f"&spatialRel=esriSpatialRelIntersects"
            f"&outFields=*"
            f"&returnGeometry=false"
            f"&f=json"
        )
        try:
            qresult = fetch_json(query_url)
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
            print(f"  Detroit point query failed: {e}")
            continue

        feats = qresult.get("features", [])
        print(f"  Detroit point-intersect: {len(feats)} feature(s)")
        if feats:
            for k, v in feats[0].get("attributes", {}).items():
                print(f"    {k:28s} = {v}")


if __name__ == "__main__":
    sys.exit(main())
