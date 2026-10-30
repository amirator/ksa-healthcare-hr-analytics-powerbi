"""Row-level security (RLS) roles for the KSA Healthcare HR Analytics semantic model.

The roles are written into model.bim as standard TMSL `roles` (modelPermission "read",
tablePermissions with DAX filterExpression), which is the format Power BI Desktop
uses for "Manage roles". They have NOT been opened or tested in Power BI Desktop or
the Power BI Service here; the tests in tests/test_rls.py instead
  1) validate the TMSL structure against the model's own tables/columns/values, and
  2) emulate each filter in Python on the embedded data, propagating it along the
     model's single-direction relationships (dimension -> fact, Employees -> Training/Leave),
     and check that every role sees exactly its own rows and nothing else.

Run `python rls.py --apply` to (re)write the roles into the existing model.bim.
"""
import base64, json, os, re, sys, zlib

ROOT = os.path.dirname(os.path.abspath(__file__))
BIM = os.path.join(ROOT, "KSA Healthcare HR Analytics.SemanticModel", "model.bim")

# (role name, description, {table: DAX filter}); an empty dict means full access.
ROLES = [
    ("HR Director", "Group HR: sees all facilities and departments.", {}),
    ("Region Manager - Western", "Western region facilities only (Jeddah, Makkah).", {"DimFacility": '[Region] = "Western"'}),
    ("Region Manager - Central", "Central region facilities only (Riyadh).", {"DimFacility": '[Region] = "Central"'}),
    ("Region Manager - Eastern", "Eastern region facilities only (Dammam).", {"DimFacility": '[Region] = "Eastern"'}),
    ("Clinical Division Manager", "Clinical departments only, all facilities.", {"DimDepartment": '[Division] = "Clinical"'}),
]

def roles_tmsl():
    out = []
    for name, desc, perms in ROLES:
        r = {"name": name, "description": desc, "modelPermission": "read"}
        if perms:
            r["tablePermissions"] = [{"name": t, "filterExpression": f} for t, f in perms.items()]
        out.append(r)
    return out

# --- emulation helpers (support the subset of DAX used above: [Col] = "literal") ---
_EQ = re.compile(r'^\[([^\]]+)\]\s*=\s*"([^"]*)"$')

def parse_filter(expr):
    m = _EQ.match(expr.strip())
    if not m:
        raise ValueError(f"unsupported filter expression: {expr}")
    return m.group(1), m.group(2)

def load_tables(bim):
    """Decode the embedded rows of every M-sourced table into lists of dicts."""
    data = {}
    for t in bim["model"]["tables"]:
        src = t["partitions"][0]["source"]
        if src["type"] != "m":
            continue
        m = "\n".join(src["expression"])
        mb = re.search(r'Binary\.FromText\("([A-Za-z0-9+/=]+)"', m)
        if not mb:  # e.g. the empty 'HR Measures' table
            continue
        b = mb.group(1)
        rows = json.loads(zlib.decompress(base64.b64decode(b), -15))
        cols = [c["name"] for c in t["columns"]]
        data[t["name"]] = [dict(zip(cols, r)) for r in rows]
    return data

def visible_rows(bim, data, role_name):
    """Apply a role's filters and propagate them along the model relationships
    (filter flows from the 'to' (one) side to the 'from' (many) side)."""
    role = next(r for r in bim["model"].get("roles", []) if r["name"] == role_name)
    vis = {t: list(rows) for t, rows in data.items()}
    for tp in role.get("tablePermissions", []):
        col, val = parse_filter(tp["filterExpression"])
        vis[tp["name"]] = [r for r in vis[tp["name"]] if r[col] == val]
    rels = [r for r in bim["model"]["relationships"] if r["fromTable"] in vis and r["toTable"] in vis]
    for _ in range(len(rels) + 1):  # propagate until stable
        for r in rels:
            keys = {row[r["toColumn"]] for row in vis[r["toTable"]]}
            vis[r["fromTable"]] = [row for row in vis[r["fromTable"]] if row[r["fromColumn"]] in keys]
    return vis

def apply(path=BIM):
    raw = open(path, encoding="utf-8-sig").read()
    bim = json.loads(raw)
    bim["model"]["roles"] = roles_tmsl()
    with open(path, "w", encoding="utf-8") as f:
        f.write(json.dumps(bim, indent=2))
    print(f"wrote {len(ROLES)} roles to {os.path.relpath(path, ROOT)}")

if __name__ == "__main__" and "--apply" in sys.argv:
    apply()
