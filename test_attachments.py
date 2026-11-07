"""Offline tests for attachments.py (no DB, no network). Run: python test_attachments.py"""
import attachments as A

# 1. Legacy record -> submission via fileLocation, approval empty
legacy = {"docNo": "X-001", "fileLocation": "https://drive.google.com/file/d/ABC123/view"}
n = A.normalize_attachments(legacy)
assert n["submission"] == {"url": legacy["fileLocation"], "legacy": True}
assert n["approval"] is None
assert A.attachment_urls(n["submission"]) == (
    legacy["fileLocation"], "https://drive.google.com/uc?export=download&id=ABC123")

# 2. Invalid legacy values are ignored
for bad in ("", "#", "null", "None", None):
    assert A.normalize_attachments({"fileLocation": bad}) == {"submission": None, "approval": None}

# 3. Folder link has no download URL
f = A.normalize_attachments({"fileLocation": "https://drive.google.com/drive/folders/F1"})
assert A.attachment_urls(f["submission"])[1] == ""

# 4. apply_attachment is non-mutating and keeps fileLocation
meta = A.build_attachment_meta("FID1", "a.pdf", 10, "application/pdf", "dc1")
new = A.apply_attachment(legacy, "approval", meta)
assert "attachments" not in legacy
assert new["fileLocation"] == legacy["fileLocation"]
assert new["attachments"]["approval"]["file_id"] == "FID1"
norm = A.normalize_attachments(new)
assert norm["submission"]["legacy"] is True and norm["approval"]["file_id"] == "FID1"

# 5. New submission overrides legacy fallback
new2 = A.apply_attachment(new, "submission", A.build_attachment_meta("FID2", "s.pdf"))
assert A.normalize_attachments(new2)["submission"]["file_id"] == "FID2"
assert A.attachment_urls(A.normalize_attachments(new2)["submission"])[0].endswith("FID2/view")

# 6. remove_attachment
rem = A.remove_attachment(new2, "submission")
assert A.normalize_attachments(rem)["submission"]["legacy"] is True
assert "attachments" in rem
rem2 = A.remove_attachment(A.remove_attachment(rem, "approval"), "approval")
assert "attachments" not in rem2

# 7. Invalid slot / bad input
for fn in (lambda: A.apply_attachment({}, "x", {}), lambda: A.remove_attachment({}, "x")):
    try:
        fn(); raise SystemExit("expected ValueError")
    except ValueError:
        pass
assert A.normalize_attachments(None) == {"submission": None, "approval": None}
print("ALL ATTACHMENT TESTS PASSED")

# 8. Sync merge (Step 3)
assert A.classify_path_kind("Root / Reply / x.pdf") == "approval"
assert A.classify_path_kind("Root / Submit / x.pdf") == "submission"
assert A.classify_path_kind("Root / Submit / Reply / x.pdf") == "approval"
assert A.classify_path_kind("Root / Other") is None
assert A.classify_path_kind("Root / 3. Material Submittal (MS) / Replied / MS-001.pdf") == "approval"
assert A.classify_path_kind("Root / 3. Material Submittal (MS) / Submit / MS-001.pdf") == "submission"
assert A.classify_path_kind("Root / 3. Material Submittal (MS) / MS-001.pdf") is None
assert A.classify_path_kind("Root / 5. Material Inspection Request (MIR) / MIR-001.pdf") is None
base = {"docNo": "X", "fileLocation": "old"}
m = A.merge_sync_links(base, {"submission": "S1", "approval": "R1"})
assert m["attachments"]["submission"] == {"url": "S1", "source": "drive_sync"}
assert m["attachments"]["approval"]["url"] == "R1" and m["fileLocation"] == "old"
assert "attachments" not in base
# DCR-uploaded slot is never overwritten; sync-created slot is refreshed
up = A.apply_attachment(base, "submission", A.build_attachment_meta("FID9", "u.pdf"))
m2 = A.merge_sync_links(up, {"submission": "S2", "approval": "R2"})
assert m2["attachments"]["submission"]["file_id"] == "FID9"
assert m2["attachments"]["approval"]["url"] == "R2"
m3 = A.merge_sync_links(m, {"submission": "S3"})
assert m3["attachments"]["submission"]["url"] == "S3" and m3["attachments"]["approval"]["url"] == "R1"
assert A.merge_sync_links(base, {}) == base and A.merge_sync_links(base, None) == base
print("SYNC MERGE TESTS PASSED")

# 9. NOC 4-Part Attachments (Part A, Part B, Part C, Part D)
assert A.classify_path_kind("Root / 11. Notice of Change (NOC) / Part A / NOC-001.pdf") == "part_a"
assert A.classify_path_kind("Root / 11. Notice of Change (NOC) / Part B / NOC-001.pdf") == "part_b"
assert A.classify_path_kind("Root / 11. Notice of Change (NOC) / Part C / NOC-001.pdf") == "part_c"
assert A.classify_path_kind("Root / 11. Notice of Change (NOC) / Part D / NOC-001.pdf") == "part_d"

noc_rec = {"docNo": "NOC-001", "fileLocation": "https://drive.google.com/file/d/OLD_NOC/view"}
noc_rec = A.apply_attachment(noc_rec, "part_a", A.build_attachment_meta("FA1", "NOC-001-PartA.pdf"))
noc_rec = A.apply_attachment(noc_rec, "part_b", A.build_attachment_meta("FB1", "NOC-001-PartB.pdf", consultant_status="Accepted & to proceed with Part C"))
noc_norm = A.normalize_attachments(noc_rec)
assert noc_norm["part_a"]["file_id"] == "FA1"
assert noc_norm["part_b"]["file_id"] == "FB1"
assert noc_norm["submission"] is None
print("NOC 4-PART ATTACHMENT TESTS PASSED")
