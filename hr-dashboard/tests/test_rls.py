import json, os, sys, unittest
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.dirname(HERE))
import rls

BIM = json.load(open(rls.BIM, encoding="utf-8-sig"))
DATA = rls.load_tables(BIM)
COLS = {t["name"]: {c["name"] for c in t["columns"]} for t in BIM["model"]["tables"]}
REGIONS = sorted({f["Region"] for f in DATA["DimFacility"]})
FAC_REGION = {f["Facility"]: f["Region"] for f in DATA["DimFacility"]}
DEP_DIV = {d["Department"]: d["Division"] for d in DATA["DimDepartment"]}
EMP = {e["EmployeeID"]: e for e in DATA["Employees"]}

class TestRLSDefinition(unittest.TestCase):
    def test_roles_in_model_match_source(self):
        self.assertEqual(BIM["model"].get("roles"), rls.roles_tmsl())

    def test_tmsl_structure_and_references(self):
        for r in BIM["model"]["roles"]:
            self.assertEqual(r["modelPermission"], "read")
            for tp in r.get("tablePermissions", []):
                self.assertIn(tp["name"], COLS)
                col, val = rls.parse_filter(tp["filterExpression"])
                self.assertIn(col, COLS[tp["name"]])
                self.assertIn(val, {row[col] for row in DATA[tp["name"]]}, "filter value must exist in data")

    def test_every_region_has_a_role(self):
        names = {r["name"] for r in BIM["model"]["roles"]}
        for reg in REGIONS:
            self.assertIn(f"Region Manager - {reg}", names)

class TestRLSBehaviour(unittest.TestCase):
    def test_hr_director_sees_everything(self):
        vis = rls.visible_rows(BIM, DATA, "HR Director")
        for t in DATA:
            self.assertEqual(len(vis[t]), len(DATA[t]), t)

    def test_region_roles_isolate_and_partition(self):
        totals = {t: 0 for t in ("Employees", "Recruitment", "Training", "Leave")}
        for reg in REGIONS:
            vis = rls.visible_rows(BIM, DATA, f"Region Manager - {reg}")
            self.assertTrue(vis["Employees"], reg)
            self.assertTrue(all(FAC_REGION[e["Facility"]] == reg for e in vis["Employees"]))
            self.assertTrue(all(FAC_REGION[r["Facility"]] == reg for r in vis["Recruitment"]))
            # child tables follow the employee filter
            for t in ("Training", "Leave"):
                self.assertTrue(all(FAC_REGION[EMP[x["EmployeeID"]]["Facility"]] == reg for x in vis[t]), t)
            for t in totals:
                totals[t] += len(vis[t])
            # dimensions/date not related to the filter are untouched
            self.assertEqual(len(vis["DimDepartment"]), len(DATA["DimDepartment"]))
        # regions are disjoint and together cover every row: no leakage, nothing lost
        for t, n in totals.items():
            self.assertEqual(n, len(DATA[t]), t)

    def test_clinical_division_role(self):
        vis = rls.visible_rows(BIM, DATA, "Clinical Division Manager")
        exp = [e for e in DATA["Employees"] if DEP_DIV[e["Department"]] == "Clinical"]
        self.assertEqual(len(vis["Employees"]), len(exp))
        self.assertLess(len(vis["Employees"]), len(DATA["Employees"]))
        self.assertTrue(all(DEP_DIV[r["Department"]] == "Clinical" for r in vis["Recruitment"]))
        self.assertEqual(len(vis["DimFacility"]), len(DATA["DimFacility"]))

    def test_unsupported_expression_rejected(self):
        with self.assertRaises(ValueError):
            rls.parse_filter("USERPRINCIPALNAME()")

if __name__ == "__main__":
    unittest.main(verbosity=2)
