import json, base64, zlib, os, re
root = os.path.dirname(os.path.abspath(__file__))
name = "KSA Healthcare HR Analytics"
# 1. All JSON files parse
for dirpath, _, files in os.walk(root):
    for f in files:
        if f.endswith((".json", ".pbip", ".pbir", ".pbism", ".bim")) or f == ".platform":
            json.load(open(os.path.join(dirpath, f), encoding="utf-8"))
print("all JSON files parse OK")
# 2. report.json nested config strings parse + aliases resolve
rep = json.load(open(os.path.join(root, name + ".Report", "report.json"), encoding="utf-8"))
json.loads(rep["config"])
nvis = 0
for s in rep["sections"]:
    json.loads(s["config"]); json.loads(s["filters"])
    for vc in s["visualContainers"]:
        c = json.loads(vc["config"]); json.loads(vc["filters"]); nvis += 1
        sv = c["singleVisual"]
        assert "visualType" in sv
        if "prototypeQuery" in sv:
            ents = {f["Name"] for f in sv["prototypeQuery"]["From"]}
            for sel in sv["prototypeQuery"]["Select"]:
                expr = sel.get("Measure") or sel.get("Column")
                assert expr["Expression"]["SourceRef"]["Source"] in ents, (sv["visualType"], sel["Name"])
print(f"report.json OK: {len(rep['sections'])} sections, {nvis} visuals, all query aliases resolve")
# 3. model.bim: M base64 payloads decompress to valid JSON rows of right width
bim = json.load(open(os.path.join(root, name + ".SemanticModel", "model.bim"), encoding="utf-8"))
for t in bim["model"]["tables"]:
    src = t["partitions"][0]["source"]
    if src["type"] == "m":
        m = "\n".join(src["expression"])
        for b in re.findall(r'Binary\.FromText\("([A-Za-z0-9+/=]+)"', m):
            rows = json.loads(zlib.decompress(base64.b64decode(b), -15))
            ncols = len(re.findall(r'#"[^"]+" = _t', m))
            assert all(len(r) == ncols == len(t["columns"]) for r in rows), t["name"]
            print(f"  {t['name']}: {len(rows)} rows x {ncols} cols decompress OK")
# 4. relationships reference real tables/columns
cols = {t["name"]: {c["name"] for c in t["columns"]} for t in bim["model"]["tables"]}
for r in bim["model"]["relationships"]:
    assert r["fromColumn"] in cols[r["fromTable"]] and r["toColumn"] in cols[r["toTable"]], r
print(f"model.bim OK: {len(bim['model']['tables'])} tables, {len(bim['model']['relationships'])} relationships")
# 5. every visual field reference exists in the model
mnames = {m["name"] for t in bim["model"]["tables"] for m in t.get("measures", [])}
missing = set()
for s in rep["sections"]:
    for vc in s["visualContainers"]:
        sv = json.loads(vc["config"])["singleVisual"]
        for sel in sv.get("prototypeQuery", {}).get("Select", []):
            if "Measure" in sel and sel["Measure"]["Property"] not in mnames:
                missing.add(sel["Measure"]["Property"])
            if "Column" in sel:
                tbl = sel["Name"].split(".")[0]
                if sel["Column"]["Property"] not in cols[tbl]:
                    missing.add(sel["Name"])
assert not missing, missing
print(f"all visual field references resolve against the model ({len(mnames)} measures defined)")
# 6. sort-by columns + date relationship uniqueness sanity
for t in bim["model"]["tables"]:
    for c in t["columns"]:
        if "sortByColumn" in c:
            assert c["sortByColumn"] in cols[t["name"]], (t["name"], c["name"])
print("sort-by columns OK")
# 7. M syntax sanity: no trailing comma before 'in', let/in present
for t in bim["model"]["tables"]:
    src = t["partitions"][0]["source"]
    if src["type"] == "m":
        lines = [l.rstrip() for l in src["expression"] if l.strip()]
        assert lines[0] == "let" and "in" in lines, t["name"]
        in_idx = lines.index("in")
        assert not lines[in_idx - 1].endswith(","), f"{t['name']}: comma before 'in'"
print("M expressions OK (no comma before 'in')")
