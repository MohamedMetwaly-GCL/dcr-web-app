"""Offline tests for drive_routing.py using a fake Drive service (no network). Run: python test_drive_routing.py"""
import drive_routing as R

FOLDER = R.FOLDER_MIME


class FakeDrive:
    """Minimal in-memory Drive: files().list/get/create with the call shapes used by drive_routing."""
    def __init__(self, tree):
        self.items = {}       # id -> dict(id,name,parent,trashed)
        self.calls = {"list": 0, "get": 0, "create": 0}
        self._n = 0
        self._load(tree, None)

    def _load(self, tree, parent):
        for name, sub in tree.items():
            fid = "id_" + name.replace(" ", "_")
            self.items[fid] = {"id": fid, "name": name, "parent": parent, "trashed": False}
            self._load(sub, fid)

    def id_of(self, name):
        return "id_" + name.replace(" ", "_")

    def files(self):
        outer = self

        class _Files:
            def list(self, q, fields=None, pageSize=None, pageToken=None, **kw):
                assert kw.get("supportsAllDrives") is True, "supportsAllDrives must be True"
                outer.calls["list"] += 1
                pid = q.split("'")[1]
                assert "mimeType = '%s'" % FOLDER in q and "trashed = false" in q
                res = [{"id": i["id"], "name": i["name"]} for i in outer.items.values()
                       if i["parent"] == pid and not i["trashed"]]
                return _Exec({"files": res})

            def get(self, fileId, fields=None, **kw):
                assert kw.get("supportsAllDrives") is True
                outer.calls["get"] += 1
                it = outer.items.get(fileId)
                if not it:
                    return _Exec(exc=Exception("404"))
                return _Exec({"id": fileId, "mimeType": FOLDER, "trashed": it["trashed"]})

            def create(self, body, fields=None, **kw):
                assert kw.get("supportsAllDrives") is True
                assert body["mimeType"] == FOLDER and len(body["parents"]) == 1
                outer.calls["create"] += 1
                outer._n += 1
                fid = "new_%d" % outer._n
                outer.items[fid] = {"id": fid, "name": body["name"], "parent": body["parents"][0], "trashed": False}
                return _Exec({"id": fid, "name": body["name"]})
        return _Files()


class _Exec:
    def __init__(self, val=None, exc=None):
        self.val, self.exc = val, exc

    def execute(self):
        if self.exc:
            raise self.exc
        return self.val


class FakeStore:
    def __init__(self, project, dts):
        self.project, self.dts, self.patches = project, dts, []

    def get_project(self, pid):
        return self.project

    def get_doc_type(self, pid, dt_id):
        return self.dts.get(dt_id)

    def set_doc_type_drive_cache(self, pid, dt_id, patch):
        self.patches.append(patch)
        self.dts[dt_id].update(patch)


# ---- name matching
assert R.score_folder_name("SD", "SD", "Shop Drawings") == R.SCORE_EXACT
assert R.score_folder_name("01 - SD", "SD", "Shop Drawings") == R.SCORE_TOKEN
assert R.score_folder_name("SD - Shop Drawings", "SD", "Shop Drawings") == R.SCORE_PREFIX
assert R.score_folder_name("Shop Drawings", "SD", "Shop Drawings") == R.SCORE_EXACT
assert R.score_folder_name("MIR", "IR", "Inspection Request") == 0       # IR must not match MIR
assert R.score_folder_name("SIRIUS", "IR", "Inspection Request") == 0
assert R.score_folder_name("03 Inspection Request (IR)", "IR", "Inspection Request") > 0
assert R.score_folder_name("", "SD", "x") == 0
m = R.match_doc_type_folder([{"id": "a", "name": "01 - SD"}, {"id": "b", "name": "MS"}], "SD", "Shop Drawings")
assert m["id"] == "a"
assert R.match_doc_type_folder([{"id": "b", "name": "MS"}], "SD", "Shop Drawings") is None
try:
    R.match_doc_type_folder([{"id": "a", "name": "SD 2020"}, {"id": "b", "name": "SD 2021"}], "SD", "Shop Drawings")
    raise SystemExit("expected ambiguous")
except R.RoutingError as e:
    assert e.reason == "ambiguous" and len(e.candidates) == 2

# ---- slot matching
fl = [{"id": "s", "name": "Submit"}, {"id": "r", "name": "Replied"}]
assert R.match_slot_folder(fl, "submission")["id"] == "s"
assert R.match_slot_folder(fl, "approval")["id"] == "r"
assert R.match_slot_folder([{"id": "x", "name": "Other"}], "approval") is None
assert R.match_slot_folder([{"id": "r", "name": "Replied"}], "submission") is None

tree = {"PROJ": {"01 - SD": {"Submit": {}, "Replied": {}}, "02 - MS": {}, "MIR": {}}}

def mk(dts=None, drive=None):
    d = drive or FakeDrive(tree)
    st = FakeStore({"drive_folder_id": d.id_of("PROJ")},
                   dts or {"SD": {"code": "SD", "name": "Shop Drawings"}, "IR": {"code": "IR", "name": "Inspection Request"},
                           "MS": {"code": "MS", "name": "Material Submittal"}})
    return d, st

# ---- discovery + cache
d, st = mk()
r = R.resolve_upload_destination("p", "SD", "submission", d, st)
assert r["folder_id"] == d.id_of("Submit") and r["doc_type_folder_id"] == d.id_of("01 - SD") and r["source"] == "discovered"
assert st.dts["SD"]["drive_folder_id"] == d.id_of("01 - SD")
assert st.dts["SD"]["drive_slot_folders"] == {"submission": d.id_of("Submit")}
r2 = R.resolve_upload_destination("p", "SD", "approval", d, st)
assert r2["folder_id"] == d.id_of("Replied")
assert st.dts["SD"]["drive_slot_folders"] == {"submission": d.id_of("Submit"), "approval": d.id_of("Replied")}
# 3rd call: fully cached -> no list calls, no writes
lc, np = d.calls["list"], len(st.patches)
r3 = R.resolve_upload_destination("p", "SD", "submission", d, st)
assert r3["source"] == "cache" and d.calls["list"] == lc and len(st.patches) == np
assert d.calls["create"] == 0
print("discovery + cache OK")

# ---- never creates by default; missing folders raise
d, st = mk()
for dt, reason in (("IR", "doc_type_folder_not_found"),):
    try:
        R.resolve_upload_destination("p", dt, "submission", d, st); raise SystemExit("expected error")
    except R.RoutingError as e:
        assert e.reason == reason
try:   # doc type folder exists but has no Submit/Replied child
    R.resolve_upload_destination("p", "MS", "submission", d, st); raise SystemExit("expected error")
except R.RoutingError as e:
    assert e.reason == "slot_folder_not_found"
assert d.calls["create"] == 0
print("no implicit creation OK")

# ---- create_if_missing creates only what is missing, under the right parent
d, st = mk()
r = R.resolve_upload_destination("p", "MS", "approval", d, st, create_if_missing=True)
created = d.items[r["folder_id"]]
assert created["name"] == "Replied" and created["parent"] == d.id_of("02 - MS") and r["source"] == "created"
assert d.calls["create"] == 1
r = R.resolve_upload_destination("p", "IR", "submission", d, st, create_if_missing=True)
assert d.items[r["doc_type_folder_id"]]["name"] == "IR - Inspection Request"
assert d.items[r["doc_type_folder_id"]]["parent"] == d.id_of("PROJ")
print("create_if_missing OK")

# ---- stale cache (folder deleted on Drive) triggers rediscovery
d, st = mk()
R.resolve_upload_destination("p", "SD", "submission", d, st)
d.items[d.id_of("Submit")]["trashed"] = True
d.items["new_sub"] = {"id": "new_sub", "name": "Submit", "parent": d.id_of("01 - SD"), "trashed": False}
r = R.resolve_upload_destination("p", "SD", "submission", d, st)
assert r["folder_id"] == "new_sub"
d.items[d.id_of("01 - SD")]["trashed"] = True
try:
    R.resolve_upload_destination("p", "SD", "submission", d, st); raise SystemExit("expected error")
except R.RoutingError as e:
    assert e.reason == "doc_type_folder_not_found"
print("stale cache OK")

# ---- guards
for args, reason in ((("p", "SD", "bogus"), "invalid_slot"),):
    try:
        R.resolve_upload_destination(*args, FakeDrive(tree), mk()[1]); raise SystemExit("x")
    except R.RoutingError as e:
        assert e.reason == reason
try:
    R.resolve_upload_destination("p", "SD", "submission", FakeDrive(tree), FakeStore({}, {"SD": {"code": "SD", "name": "n"}}))
    raise SystemExit("x")
except R.RoutingError as e:
    assert e.reason == "project_has_no_drive_folder"
try:
    R.resolve_upload_destination("p", "ZZ", "submission", FakeDrive(tree), mk()[1]); raise SystemExit("x")
except R.RoutingError as e:
    assert e.reason == "doc_type_not_found"
# ---- directional routing & separate directional caching
ltr_tree = {
    "PROJ_LTR": {
        "01 - INCOMING LETTERS": {},
        "02 - OUTGOING LETTERS": {}
    }
}
d_ltr = FakeDrive(ltr_tree)
st_ltr = FakeStore({"drive_folder_id": d_ltr.id_of("PROJ_LTR")},
                   {"LTR": {"code": "LTR", "name": "Letters"}})

# Upload incoming
r_inc = R.resolve_upload_destination("p", "LTR", "submission", d_ltr, st_ltr, flat=True, direction="Received")
assert r_inc["folder_id"] == d_ltr.id_of("01 - INCOMING LETTERS")
assert st_ltr.dts["LTR"].get("drive_direction_folders", {}).get("incoming") == d_ltr.id_of("01 - INCOMING LETTERS")

# Upload outgoing -> must route to OUTGOING and not use incoming cache
r_out = R.resolve_upload_destination("p", "LTR", "submission", d_ltr, st_ltr, flat=True, direction="Issued")
assert r_out["folder_id"] == d_ltr.id_of("02 - OUTGOING LETTERS")
assert st_ltr.dts["LTR"].get("drive_direction_folders", {}).get("outgoing") == d_ltr.id_of("02 - OUTGOING LETTERS")

# 2nd incoming upload -> cache hit for incoming
r_inc2 = R.resolve_upload_destination("p", "LTR", "submission", d_ltr, st_ltr, flat=True, direction="Received")
assert r_inc2["folder_id"] == d_ltr.id_of("01 - INCOMING LETTERS")
assert r_inc2["source"] == "cache"
print("directional routing & caching OK")

print("ALL DRIVE ROUTING TESTS PASSED")
