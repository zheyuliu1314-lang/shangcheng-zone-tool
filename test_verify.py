import requests, json

# Test: compare old (addr-only) vs new (company+addr)
test_cases = [
    ("杭州近江股份经济合作社", "杭州市上城区近江家园四园18幢"),
    ("杭州近江物业管理有限公司", "上城区近江南路1号"),
    ("杭州之江路888号商业运营管理有限公司", "浙江省杭州市上城区之江路888号"),
    ("杭州富阳利丰交通设施有限公司", "浙江省杭州市上城区婺江路217号2号楼607室86号"),
]

print("=== New approach (company name + address) ===")
for name, addr in test_cases:
    item = [{"addr": addr, "name": name}]
    r = requests.post("http://127.0.0.1:5050/api/geocode",
        json={"items": item, "addr_col": "addr", "name_col": "name"}, timeout=15).json()
    res = r["results"][0]
    ok = "OK" if res.get("lng") else "FAIL"
    print(f"  {ok} ({res.get('strategy','-'):18s}) Lng={res.get('lng'):.6f} Lat={res.get('lat'):.6f} level={res.get('level','-'):10s}  {name[:40]}")

print()
print("=== Old approach (address only, no company) ===")
for name, addr in test_cases:
    item = [{"addr": addr}]
    r = requests.post("http://127.0.0.1:5050/api/geocode",
        json={"items": item, "addr_col": "addr", "name_col": ""}, timeout=15).json()
    res = r["results"][0]
    ok = "OK" if res.get("lng") else "FAIL"
    print(f"  {ok} ({res.get('strategy','-'):18s}) Lng={res.get('lng'):.6f} Lat={res.get('lat'):.6f} level={res.get('level','-'):10s}  {addr[:40]}")

# Check classification
print()
print("=== Classification test ===")
with open(r"C:\Users\wy98k\Documents\片区分类开发工具。\zones.json", "r", encoding="utf-8") as f:
    zones = json.load(f)
print(f"Zones loaded: {len(zones)} zones")
for z in zones:
    print(f"  {z['name']}: {len(z['polygon'])} vertices")

# Test full pipeline
print()
print("=== Full pipeline test ===")
items = [{"addr": addr, "name": name} for name, addr in test_cases]
r = requests.post("http://127.0.0.1:5050/api/geocode",
    json={"items": items, "addr_col": "addr", "name_col": "name"}, timeout=30).json()
geocoded = [res for res in r["results"] if res.get("lng")]

# Classify
r2 = requests.post("http://127.0.0.1:5050/api/classify",
    json={"points": geocoded, "zones": zones}, timeout=15).json()
for res in r2["results"]:
    print(f"  {res.get('zone','?'):8s} Lng={res.get('lng'):.6f} Lat={res.get('lat'):.6f}  {res.get('addr','')[:40]}")
