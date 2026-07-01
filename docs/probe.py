#!/usr/bin/env python3
"""Probe MI/IN/WI NFC + NTAD ACUB FeatureServer schemas (verification §5.1).

Hits every wired-state FeatureServer endpoint, prints each layer's schema,
then runs a point-intersect query at a known good coordinate for that state
so you can see which fields actually carry useful data.

Run from any machine with internet access. The cloud build sandbox is
firewalled off `mdotgis.state.mi.us`, so the MI layers must be probed from a
local workstation (see docs/probe-mdot-layers.md); `gisdata.in.gov` and
`services5.arcgis.com` (IN/WI) are reachable from the cloud sandbox too.

    python3 docs/probe.py
"""
import json
import sys
import urllib.error
import urllib.request

# Each entry: (label, base FeatureServer URL, test lon, test lat)
LAYERS = [
    ("MDOT 353 (Functional System)",
     "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353",
     -83.045, 42.331),   # downtown Detroit
    ("MDOT 364 (Classification)",
     "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/364",
     -83.045, 42.331),
    ("MDOT 543 (Route System)",
     "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/543",
     -83.045, 42.331),
    ("NTAD ACUB (nationwide)",
     "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0",
     -83.045, 42.331),
    ("INDOT LRSE_Functional_Class",
     "https://gisdata.in.gov/server/rest/services/Hosted/LRSE_Functional_Class/FeatureServer/22",
     -86.1581, 39.7684),   # downtown Indianapolis
    ("INDOT Road_Centerlines_of_Indiana_2021",
     "https://gisdata.in.gov/server/rest/services/Hosted/Road_Centerlines_of_Indiana_2021/FeatureServer/15",
     -86.1581, 39.7684),
    ("WisDOT State Trunk Network (FFCL_gdb/3)",
     "https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/FFCL_gdb/FeatureServer/3",
     -87.9065, 43.0389),   # downtown Milwaukee
    ("WisDOT Local Road Network snapshot",
     "https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/WI_Local_Roads_Flood_Damage_Assessment_Snapshot/FeatureServer/1",
     -87.9065, 43.0389),
]


def fetch_json(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": "RoadReviewer/probe"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.load(r)


def main():
    for label, base, test_lon, test_lat in LAYERS:
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
            f"?geometry={test_lon},{test_lat}"
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
            print(f"  point query failed: {e}")
            continue

        feats = qresult.get("features", [])
        print(f"  point-intersect ({test_lon},{test_lat}): {len(feats)} feature(s)")
        if feats:
            for k, v in feats[0].get("attributes", {}).items():
                print(f"    {k:28s} = {v}")


if __name__ == "__main__":
    sys.exit(main())
