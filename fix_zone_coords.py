#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fix zone coordinates utility
Run this script if zone boundaries on the Gaode map look ~500m off.
This converts zone vertices from WGS-84 to GCJ-02 (Mars coordinates).
Usage: python3 fix_zone_coords.py
"""
import json, os, math

def delta(lng, lat):
    def transform_lat(x, y):
        ret = -100.0 + 2.0*x + 3.0*y + 0.2*y*y + 0.1*x*y + 0.2*abs(x)**0.5
        ret += (20.0*math.sin(6.0*x*math.pi) + 20.0*math.sin(2.0*x*math.pi)) * 2.0/3.0
        ret += (20.0*math.sin(y*math.pi) + 40.0*math.sin(y/3.0*math.pi)) * 2.0/3.0
        ret += (160.0*math.sin(y/12.0*math.pi) + 320.0*math.sin(y*math.pi/30.0)) * 2.0/3.0
        return ret
    def transform_lng(x, y):
        ret = 300.0 + x + 2.0*y + 0.1*x*x + 0.1*x*y + 0.1*abs(x)**0.5
        ret += (20.0*math.sin(6.0*x*math.pi) + 20.0*math.sin(2.0*x*math.pi)) * 2.0/3.0
        ret += (20.0*math.sin(x*math.pi) + 40.0*math.sin(x/3.0*math.pi)) * 2.0/3.0
        ret += (150.0*math.sin(x/12.0*math.pi) + 300.0*math.sin(x/30.0*math.pi)) * 2.0/3.0
        return ret
    dlat = transform_lat(lng - 105.0, lat - 35.0)
    dlng = transform_lng(lng - 105.0, lat - 35.0)
    radlat = lat / 180.0 * math.pi
    magic = math.sin(radlat)
    magic = 1 - 0.00669342162296594323 * magic * magic
    sqrtmagic = math.sqrt(magic)
    dlat = (dlat * 180.0) / ((6378245.0 * (1 - 0.00669342162296594323)) / (magic * sqrtmagic) * math.pi)
    dlng = (dlng * 180.0) / (6378245.0 / sqrtmagic * math.cos(radlat) * math.pi)
    return dlng, dlat

def wgs84_to_gcj02(lng, lat):
    dlng, dlat = delta(lng, lat)
    return lng + dlng, lat + dlat

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ZONES_FILE = os.path.join(SCRIPT_DIR, "zones.json")

def load_zones():
    if os.path.exists(ZONES_FILE):
        with open(ZONES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_zones(zones):
    with open(ZONES_FILE, "w", encoding="utf-8") as f:
        json.dump(zones, f, ensure_ascii=False, indent=2)

def main():
    zones = load_zones()
    if not zones:
        print("No zones found in zones.json")
        return
    print(f"Found {len(zones)} zones")
    for z in zones:
        name = z.get("name", "Unknown")
        poly = z.get("polygon", [])
        print(f"  Converting {name} ({len(poly)} vertices)")
        new_poly = [[round(wgs84_to_gcj02(p[0], p[1])[0], 6),
                      round(wgs84_to_gcj02(p[0], p[1])[1], 6)] for p in poly]
        z["polygon"] = new_poly
    save_zones(zones)
    print("Done! Zones converted from WGS-84 to GCJ-02")
    print("Please verify the zone positions on the map")

if __name__ == "__main__":
    main()
