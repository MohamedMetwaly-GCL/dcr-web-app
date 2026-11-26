"""Regression tests for DRIVE-FIX-01: Submit vs. Replied Link Isolation & Drive Sync Hardening.

Covers all 4 acceptance criteria offline (zero network / live DB calls):
1. Record with Reply Code ('A - Approved'), file in Submit/, empty Reply/Replied displays ONLY Submit
   in register table, edit modal (mountDriveAttachmentsBox), and drive_viewer (404 on approval).
2. Record with files in both Submit/ and Replied/ (including legacy Submit) displays BOTH links.
3. File named '*_For_Approval.pdf' inside Submit/ is classified as 'submission' by attachments & drive_service.
4. db.update_record_link scopes updates strictly by project_id, and db.set_doc_type_drive_cache retains
   'drive_direction_folders'.
"""

import json
import os
from pathlib import Path
import re
import subprocess
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-drive-fix-01")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-drive-fix-01")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-drive-fix-01")
os.environ.setdefault("ALLOW_INSECURE_COOKIES_FOR_LOCAL_DEV", "1")

import app as app_mod
import attachments
import blueprints.drive_viewer as dv_mod
import db
import drive_service
import html_render


class _FakeDriveListExec:
    def __init__(self, payload):
        self._payload = payload

    def execute(self):
        return self._payload


class _FakeDriveFilesResource:
    def __init__(self, items):
        self._items = items

    def list(self, q="", fields=None, pageSize=1000, pageToken=None, **kwargs):
        parent_ids = re.findall(r"'([^']+)'\s+in\s+parents", q or "")
        matched = [
            it for it in self._items
            if any(p in (it.get("parents") or []) for p in parent_ids) and not it.get("trashed", False)
        ]
        return _FakeDriveListExec({"files": matched})


class _FakeDriveSyncService:
    def __init__(self, items):
        self._items = items

    def files(self):
        return _FakeDriveFilesResource(self._items)


class TestDriveSubmitReplyIsolationBackend(unittest.TestCase):
    def test_01_directory_only_classification_excludes_filename(self):
        cases_submission = (
            "Root / 2. Shop Drawings (SD) / Submit / DS-001_For_Approval.pdf",
            "Root / 2. Shop Drawings (SD) / Submit / DS-001 For Approval.pdf",
            "Root / 2. Shop Drawings (SD) / Submit / DS-001_Approved_By_Client.pdf",
            "Root / 2. Shop Drawings (SD) / Submit / DS-001_Reply_Sheet.pdf",
            "Root / Submit / DS-001 For Approval REV00 / DS-001_For_Approval.pdf",
            r"Root\01 - DS\Submit\DS-001_For_Approval.pdf",
        )
        for path in cases_submission:
            with self.subTest(path=path):
                self.assertEqual(attachments.classify_path_kind(path), "submission")
                self.assertEqual(drive_service.get_item_priority(path, "application/pdf"), 4)

        # Folder item representing a document subfolder inside Submit/
        doc_folder_in_submit = "Root / 2. Shop Drawings (SD) / Submit / DS-001_For_Approval REV00"
        self.assertEqual(attachments.classify_path_kind(doc_folder_in_submit), "submission")
        self.assertEqual(
            drive_service.get_item_priority(doc_folder_in_submit, "application/vnd.google-apps.folder"),
            3,
        )

        # File inside Replied/ whose filename contains 'Submit' must still be 'approval'
        replied_file = "Root / 2. Shop Drawings (SD) / Replied / DS-001_Submit_Stamped.pdf"
        self.assertEqual(attachments.classify_path_kind(replied_file), "approval")
        self.assertEqual(drive_service.get_item_priority(replied_file, "application/pdf"), 2)

        # File directly under category folder (no Submit/Replied subfolder) must return None / priority 9
        flat_file = "Root / 3. Material Submittal (MS) / MS-001_For_Approval.pdf"
        self.assertIsNone(attachments.classify_path_kind(flat_file))
        self.assertEqual(drive_service.get_item_priority(flat_file, "application/pdf"), 9)

    def test_02_drive_viewer_strictly_isolates_empty_reply_and_noc_slots(self):
        rec_submit_only = {
            "_id": "REC-DS-001",
            "_project_id": "PEM-058",
            "_dt_id": "DS",
            "docNo": "DS-001",
            "status": "A - Approved",
            "fileLocation": "https://drive.google.com/file/d/SUB_FILE_111/view",
            "driveLink": "https://drive.google.com/file/d/SUB_FILE_111/view",
            "attachments": {
                "submission": {
                    "url": "https://drive.google.com/file/d/SUB_FILE_111/view",
                    "source": "drive_sync",
                }
            },
        }
        viewer_user = {"username": "dc_user", "role": "editor"}

        with patch("blueprints.drive_viewer.current_user", return_value=viewer_user), \
             patch("db.get_record_by_id", return_value=rec_submit_only), \
             patch("blueprints.drive_viewer.can_view_project", return_value=True), \
             patch("blueprints.drive_viewer.can_access_doc_type", return_value=True), \
             patch("blueprints.drive_viewer.can_edit", return_value=True), \
             patch("blueprints.drive_viewer._is_workspace_admin", return_value=False):

            with app_mod.app.test_request_context("/api/drive/view-info/REC-DS-001/submission"):
                ctx_sub, err_sub = dv_mod._context("REC-DS-001", "submission")
                self.assertIsNone(err_sub)
                self.assertEqual(ctx_sub["target_id"], "SUB_FILE_111")

            # Empty approval and NOC slots MUST return 404, never falling back to submission or fileLocation
            for empty_slot in ("approval", "part_a", "part_b", "part_c", "part_d"):
                with self.subTest(slot=empty_slot):
                    with app_mod.app.test_request_context(f"/api/drive/view-info/REC-DS-001/{empty_slot}"):
                        ctx_emp, err_emp = dv_mod._context("REC-DS-001", empty_slot)
                        self.assertIsNone(ctx_emp)
                        resp, status_code = err_emp
                        self.assertEqual(status_code, 404)
                        self.assertIn("Nothing is attached", resp.get_json().get("error", ""))

        # Also test via full Flask test_client on /api/drive/view-info and /api/drive/stream
        client = app_mod.app.test_client()
        client.set_cookie("dcr_token", "tok_editor")
        with patch("db.get_session", return_value=viewer_user), \
             patch("db.get_record_by_id", return_value=rec_submit_only), \
             patch("blueprints.drive_viewer.can_view_project", return_value=True), \
             patch("blueprints.drive_viewer.can_access_doc_type", return_value=True):
            r_vi = client.get("/api/drive/view-info/REC-DS-001/approval")
            self.assertEqual(r_vi.status_code, 404)
            r_st = client.get("/api/drive/stream/REC-DS-001/approval/SUB_FILE_111")
            self.assertEqual(r_st.status_code, 404)
            r_zip = client.get("/api/drive/folder-zip/REC-DS-001/approval")
            self.assertEqual(r_zip.status_code, 404)

        # Also verify legacy submission dict ({'url': ..., 'legacy': True}) never leaks into approval
        rec_legacy_sub = {
            "_id": "REC-DS-002",
            "_project_id": "PEM-058",
            "_dt_id": "DS",
            "docNo": "DS-002",
            "status": "B - Approved As Noted",
            "fileLocation": "https://drive.google.com/file/d/LEG_SUB_222/view",
            "attachments": {
                "submission": {"url": "https://drive.google.com/file/d/LEG_SUB_222/view", "legacy": True}
            },
        }
        with patch("blueprints.drive_viewer.current_user", return_value=viewer_user), \
             patch("db.get_record_by_id", return_value=rec_legacy_sub), \
             patch("blueprints.drive_viewer.can_view_project", return_value=True), \
             patch("blueprints.drive_viewer.can_access_doc_type", return_value=True):
            with app_mod.app.test_request_context("/api/drive/view-info/REC-DS-002/approval"):
                ctx_emp2, err_emp2 = dv_mod._context("REC-DS-002", "approval")
                self.assertIsNone(ctx_emp2)
                self.assertEqual(err_emp2[1], 404)

    def test_03_project_scoped_sync_and_update_record_link_isolation(self):
        # In-memory simulation of projects + records across two distinct projects sharing docNo 'DS-001'
        projects_rows = [
            {"id": "PROJ-A", "data": {"drive_folder_id": "https://drive.google.com/drive/folders/ROOT_FOLDER_A?usp=sharing"}},
            {"id": "PROJ-B", "data": {"drive_folder_id": "ROOT_FOLDER_B"}},
        ]
        records_store = {
            "REC-A-1": {
                "id": "REC-A-1",
                "project_id": "PROJ-A",
                "data": {"docNo": "DS-001", "status": "A - Approved"},
            },
            "REC-B-1": {
                "id": "REC-B-1",
                "project_id": "PROJ-B",
                "data": {
                    "docNo": "DS-001",
                    "status": "Under Review",
                    "fileLocation": "https://drive.google.com/file/d/ORIG_B/view",
                },
            },
        }
        executed_updates = []

        def fake_q(sql, params=(), one=False):
            s = " ".join(sql.split())
            if "FROM projects" in s:
                return projects_rows
            if "FROM records" in s:
                self.assertIn("project_id = %s", s, "update_record_link SELECT must filter by project_id")
                pid_arg, doc_arg = params[0], params[1]
                matched = [
                    {"id": r["id"], "data": r["data"]}
                    for r in records_store.values()
                    if r["project_id"] == pid_arg and r["data"].get("docNo") == doc_arg
                ]
                return (matched[0] if matched else None) if one else matched
            return None if one else []

        def fake_exe(sql, params=()):
            s = " ".join(sql.split())
            self.assertIn("AND project_id=%s", s, "update_record_link UPDATE must enforce project_id")
            new_data_json, rec_id, pid_arg = params
            executed_updates.append((rec_id, pid_arg, json.loads(new_data_json)))
            if rec_id in records_store and records_store[rec_id]["project_id"] == pid_arg:
                records_store[rec_id]["data"] = json.loads(new_data_json)
            return 1

        # Simulate Drive tree under ROOT_FOLDER_A:
        # Submit/ contains 'DS-001_For_Approval.pdf'; Replied/ is empty!
        drive_items = [
            {
                "id": "F_DS",
                "name": "2. Document Submittal (DS)",
                "mimeType": "application/vnd.google-apps.folder",
                "parents": ["ROOT_FOLDER_A"],
                "webViewLink": "https://drive.google.com/drive/folders/F_DS",
            },
            {
                "id": "F_SUB",
                "name": "Submit",
                "mimeType": "application/vnd.google-apps.folder",
                "parents": ["F_DS"],
                "webViewLink": "https://drive.google.com/drive/folders/F_SUB",
            },
            {
                "id": "F_REP",
                "name": "Replied",
                "mimeType": "application/vnd.google-apps.folder",
                "parents": ["F_DS"],
                "webViewLink": "https://drive.google.com/drive/folders/F_REP",
            },
            {
                "id": "FILE_A_SUB",
                "name": "DS-001_For_Approval.pdf",
                "mimeType": "application/pdf",
                "parents": ["F_SUB"],
                "webViewLink": "https://drive.google.com/file/d/FILE_A_SUB/view",
            },
        ]

        with patch("db.q", side_effect=fake_q), \
             patch("db.exe", side_effect=fake_exe), \
             patch.object(drive_service, "get_drive_service", return_value=_FakeDriveSyncService(drive_items)):
            drive_service.process_drive_folder("ROOT_FOLDER_A")

        # 1) PROJ-A's DS-001 updated with submission ONLY; approval is NOT populated
        rec_a_after = records_store["REC-A-1"]["data"]
        self.assertEqual(rec_a_after["fileLocation"], "https://drive.google.com/file/d/FILE_A_SUB/view")
        self.assertEqual(
            rec_a_after["attachments"],
            {"submission": {"url": "https://drive.google.com/file/d/FILE_A_SUB/view", "source": "drive_sync"}},
        )
        self.assertNotIn("approval", rec_a_after["attachments"])

        # 2) PROJ-B's DS-001 was NEVER touched
        rec_b_after = records_store["REC-B-1"]["data"]
        self.assertEqual(rec_b_after["fileLocation"], "https://drive.google.com/file/d/ORIG_B/view")
        self.assertNotIn("attachments", rec_b_after)
        self.assertEqual([u[1] for u in executed_updates], ["PROJ-A"])

        # 3) Unscoped call without project_id refuses to overwrite records
        with patch("db.q", side_effect=fake_q), patch("db.exe", side_effect=fake_exe):
            self.assertEqual(db.update_record_link("DS-001", "https://drive.google.com/file/d/X/view"), 0)

    def test_04_set_doc_type_drive_cache_persists_drive_direction_folders(self):
        captured = []

        def fake_q(sql, params=(), one=False):
            captured.append((sql, params))
            return [{"id": "LTR"}]

        patch_payload = {
            "drive_direction_folders": {"incoming": "FLD_INC_1", "outgoing": "FLD_OUT_1"},
            "drive_slot_folders": {"submission": "FLD_SUB_1"},
            "folder_synced_at": "2026-11-25T10:00:00Z",
            "untrusted_extra_key": "must_be_stripped",
        }
        with patch("db.q", side_effect=fake_q):
            ok = db.set_doc_type_drive_cache("PEM-058", "LTR", patch_payload)

        self.assertTrue(ok)
        self.assertEqual(len(captured), 1)
        persisted = json.loads(captured[0][1][0])
        self.assertEqual(
            persisted.get("drive_direction_folders"),
            {"incoming": "FLD_INC_1", "outgoing": "FLD_OUT_1"},
        )
        self.assertEqual(persisted.get("drive_slot_folders"), {"submission": "FLD_SUB_1"})
        self.assertEqual(persisted.get("folder_synced_at"), "2026-11-25T10:00:00Z")
        self.assertNotIn("untrusted_extra_key", persisted)


class TestDriveSubmitReplyIsolationFrontendJS(unittest.TestCase):
    def _render_eng_register(self):
        user = {"username": "eng_editor", "role": "editor"}
        project = {
            "id": "PEM-058",
            "code": "PEM-058",
            "name": "CFC Phase 2",
            "workspace_id": "ENG",
        }
        dt_list = [
            {"id": "DS", "code": "DS", "name": "Document Submittal", "project_id": "PEM-058", "sort_order": 0},
            {"id": "PR", "code": "PR", "name": "Purchase Requisition", "project_id": "PEM-058", "sort_order": 1},
        ]
        with patch("html_render.can_edit", return_value=True), \
             patch("db.can_user_download", return_value=True), \
             patch("db.get_projects_for_user", return_value=[project]), \
             patch("db.get_doc_types", return_value=dt_list), \
             patch("db.get_logo", return_value=None), \
             patch("db.get_user_workspaces", return_value=[{"id": "ENG", "name": "Engineering"}]), \
             patch("db.is_user_workspace_admin", return_value=False):
            return html_render.render_register(user, project, lang="en")

    def test_05_frontend_table_and_modal_slot_isolation_and_dual_preservation(self):
        html = self._render_eng_register()
        scripts = re.findall(r"<script(?:\s+[^>]*)?>([\s\S]*?)</script>", html, re.IGNORECASE)
        main_scripts = [s for s in scripts if "_normalizeRowAttachments" in s and "mountDriveAttachmentsBox" in s]
        self.assertEqual(len(main_scripts), 1)
        main_script = main_scripts[0]

        node_test = r"""
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const mainJs = fs.readFileSync(process.argv[2], 'utf8');

function mkEl(tag) {
  const listeners = {};
  const attrs = {};
  const children = [];
  const classSet = new Set();
  const el = {
    tagName: String(tag || 'DIV').toUpperCase(),
    id: '',
    value: '',
    placeholder: '',
    textContent: '',
    title: '',
    href: '',
    style: { cssText: '' },
    dataset: {},
    children,
    classList: {
      add: (...cs) => cs.forEach(c => classSet.add(c)),
      remove: (...cs) => cs.forEach(c => classSet.delete(c)),
      toggle: (c, force) => {
        if (force === undefined) { classSet.has(c) ? classSet.delete(c) : classSet.add(c); }
        else if (force) classSet.add(c);
        else classSet.delete(c);
      },
      contains: (c) => classSet.has(c),
    },
    set className(v) { classSet.clear(); String(v || '').split(/\s+/).filter(Boolean).forEach(c => classSet.add(c)); },
    get className() { return [...classSet].join(' '); },
    setAttribute: (k, v) => { attrs[k] = String(v); },
    getAttribute: (k) => attrs[k] || null,
    appendChild: (c) => { children.push(c); c.parentElement = el; return c; },
    insertBefore: (c) => { children.push(c); c.parentElement = el; return c; },
    remove: () => {
      if (el.parentElement && Array.isArray(el.parentElement.children)) {
        const i = el.parentElement.children.indexOf(el);
        if (i >= 0) el.parentElement.children.splice(i, 1);
      }
    },
    addEventListener: (ev, fn) => { (listeners[ev] = listeners[ev] || []).push(fn); },
    dispatchEvent: (ev) => { (listeners[ev.type] || []).forEach(fn => fn(ev)); },
    querySelector: (sel) => {
      const all = walkEls(el);
      if (String(sel).startsWith('#')) {
        const wantId = String(sel).slice(1);
        return all.find(x => x.id === wantId) || null;
      }
      return all.find(x => x.tagName === String(sel).toUpperCase()) || null;
    },
    querySelectorAll: () => [],
    closest: () => null,
    get isConnected() { return true; },
  };
  Object.defineProperty(el, 'innerHTML', {
    get() { return this._inner || ''; },
    set(v) { this._inner = String(v || ''); if (!v) children.length = 0; },
  });
  return el;
}

function walkEls(root, out = []) {
  if (!root || !root.children) return out;
  for (const ch of root.children) {
    out.push(ch);
    walkEls(ch, out);
  }
  return out;
}

const byId = {};
[
  'tabsbar', 'ptw-subtabs', 'toolbar', 'btn-add-subtab', 'regtbl', 'thead', 'tbody', 'srchbox',
  'rec-form', 'rec-title', 'rec-modal', 'rec-save-btn', 'rec-cancel-btn',
  's-clock', 's-total', 's-show', 's-ov', 'empty'
].forEach(id => {
  const e = mkEl('div'); e.id = id; byId[id] = e;
});

const bodyEl = mkEl('body');
const viewerCalls = [];

const documentStub = {
  documentElement: { dir: 'ltr', lang: 'en' },
  body: bodyEl,
  createElement: (tag) => mkEl(tag),
  getElementById: (id) => {
    if (byId[id]) return byId[id];
    const found = walkEls(byId['rec-form']).find(e => e.id === id);
    if (found) return found;
    const autoEl = mkEl('div');
    autoEl.id = id;
    byId[id] = autoEl;
    return autoEl;
  },
  querySelector: () => null,
  querySelectorAll: () => [],
  addEventListener: () => {},
};

const ctx = {
  console, Math, Date, RegExp, Map, Set, Array, Object, String, Number, Boolean, JSON, Promise, Error,
  encodeURIComponent, decodeURIComponent, parseInt, parseFloat, isNaN,
  setTimeout: () => 0,
  setInterval: () => 0,
  requestAnimationFrame: () => 0,
  URLSearchParams: function() { return { get: () => null }; },
  location: { search: '' },
  document: documentStub,
  CustomEvent: function(type, init) { this.type = type; this.detail = init && init.detail; },
  Event: function(type) { this.type = type; },
  CSS: { escape: s => s },
  window: {
    t: (k) => {
      const map = {
        'form.slot_subm_title': 'Contractor Submission',
        'form.slot_subm_empty': 'No submission file attached',
        'form.slot_reply_title': 'Consultant Reply',
        'form.slot_reply_empty': 'No reply file attached',
        'form.slot_doc_title': 'Document',
        'form.slot_doc_empty': 'No document attached',
        'form.upload_file': 'Upload File',
        'form.replace_file': 'Replace File',
        'form.upload_folder': 'Upload Folder',
        'form.view_btn': 'View',
        'form.dl_btn': 'Download',
        'form.confirm_unlink': 'Unlink',
        'form.drive_link_legacy': 'Legacy Drive Link',
      };
      return map[k] || k;
    },
    tCol: (_k, l) => l,
    tDt: (_c, n) => n,
    tStatus: (s) => s,
    matchMedia: () => ({ matches: false }),
    addEventListener: () => {},
  },
  escHtml: (s) => String(s || ''),
  confirm: () => true,
  apiFetch: async (u) => (String(u).includes('/api/doc_types/') || String(u).includes('/api/columns/')) ? [] : ({}),
  fetch: async () => ({ ok: true, json: async () => ({}) }),
};
ctx.window.document = documentStub;
ctx.globalThis = ctx;

const api = vm.runInNewContext(
  'function bindDirectionalInput(){} function bindSmartTextarea(){} function openM(){} function closeM(){} function toast(){} function initRz(){}\n' +
  mainJs +
  '\nopenDocViewer = function(rid, slot, opts){ viewerCalls.push({ rid, slot, opts }); };\n' +
  ';({ state, _normalizeRowAttachments, mountDriveAttachmentsBox, renderRows });',
  Object.assign(ctx, { viewerCalls })
);

api.state.tab = 'DS';
api.state.dtList = [
  { id: 'DS', code: 'DS', name: 'Document Submittal' },
  { id: 'PR', code: 'PR', name: 'Purchase Requisition' },
];
api.state.cols = [
  { col_key: 'docNo', label: 'Doc No.', col_type: 'docno', visible: true },
  { col_key: 'status', label: 'Status', col_type: 'dropdown', list_name: 'status', visible: true },
  { col_key: 'fileLocation', label: 'Drive Link', col_type: 'url', visible: true },
];
api.state.allTabCols = api.state.cols;

// ── Criterion 1a: Drive-synced file in Submit/, empty Reply, status = 'A - Approved' ──
const rowSyncedSubmitOnly = {
  _id: 'R-SYNC-1',
  docNo: 'DS-001',
  status: 'A - Approved',
  fileLocation: 'https://drive.google.com/file/d/SUB_SYNC_1/view',
  attachments: {
    submission: { url: 'https://drive.google.com/file/d/SUB_SYNC_1/view', source: 'drive_sync' },
  },
};
const norm1 = api._normalizeRowAttachments(rowSyncedSubmitOnly, 'DS', rowSyncedSubmitOnly.fileLocation);
assert.ok(norm1.submission, 'Submit slot must exist');
assert.strictEqual(norm1.submission.url, 'https://drive.google.com/file/d/SUB_SYNC_1/view');
assert.strictEqual(norm1.approval, null, 'Reply slot must remain null when Reply folder is empty');

// ── Criterion 1b: slot_verified=true or legacy submission object with status 'B - Approved As Noted' / 'C - Revise & Resubmit' ──
for (const stCode of ['A - Approved', 'B - Approved As Noted', 'C - Revise & Resubmit', 'D - Rejected']) {
  const rowVer = {
    _id: 'R-VER',
    docNo: 'DS-002',
    status: stCode,
    attachments: {
      submission: { url: 'https://drive.google.com/file/d/SUB_VER_2/view', slot_verified: true },
    },
  };
  const nv = api._normalizeRowAttachments(rowVer, 'DS');
  assert.ok(nv.submission && nv.submission.url.includes('SUB_VER_2'));
  assert.strictEqual(nv.approval, null, 'slot_verified submission must never move to approval for ' + stCode);

  const rowUnverSub = {
    _id: 'R-LEG-SUB',
    docNo: 'DS-003',
    status: stCode,
    fileLocation: 'https://drive.google.com/file/d/SUB_LEG_3/view',
    attachments: {
      submission: { url: 'https://drive.google.com/file/d/SUB_LEG_3/view', legacy: true },
    },
  };
  const nl = api._normalizeRowAttachments(rowUnverSub, 'DS', rowUnverSub.fileLocation);
  assert.ok(nl.submission && nl.submission.url.includes('SUB_LEG_3'), 'submission must stay in submission for ' + stCode);
  assert.strictEqual(nl.approval, null, 'existing submission must never reclassify into approval for ' + stCode);
}

// ── Criterion 1c: Register table row renders ONLY 'Subm' chip and NO 'Reply' chip ──
api.state.recs = [rowSyncedSubmitOnly];
api.state.visibleRows = [rowSyncedSubmitOnly];
api.renderRows();
const tr1 = byId['tbody'].children[0];
const chips1 = walkEls(tr1).filter(el => el.textContent === 'Subm' || el.textContent === 'Reply');
assert.deepStrictEqual(chips1.map(c => c.textContent), ['Subm'], 'Table must render ONLY Subm chip when Reply is empty');

// ── Criterion 1d: Edit modal (mountDriveAttachmentsBox) shows file in Submit & empty Reply ──
const formRoot1 = mkEl('div');
api.state.editId = 'R-SYNC-1';
api.mountDriveAttachmentsBox(formRoot1, {}, rowSyncedSubmitOnly);
const submInfo1 = walkEls(formRoot1).find(e => e.id === 'drv-slot-info-submission');
const apprInfo1 = walkEls(formRoot1).find(e => e.id === 'drv-slot-info-approval');
assert.ok(submInfo1, 'Submission slot card must exist in modal');
assert.ok(apprInfo1, 'Approval slot card must exist in modal');
const submTexts1 = walkEls(submInfo1).map(e => e.textContent).join(' | ');
const apprTexts1 = walkEls(apprInfo1).map(e => e.textContent).join(' | ');
assert.ok(!submTexts1.includes('No submission file attached'), 'Submission card must show attached file');
assert.ok(apprTexts1.includes('No reply file attached'), 'Approval card must show empty message: got ' + apprTexts1);

// ── Criterion 2a: Both Submit and Reply synced/uploaded -> BOTH preserved and rendered ──
const rowBothSynced = {
  _id: 'R-BOTH-1',
  docNo: 'DS-010',
  status: 'A - Approved',
  fileLocation: 'https://drive.google.com/file/d/REP_SYNC_10/view',
  attachments: {
    submission: { url: 'https://drive.google.com/file/d/SUB_SYNC_10/view', source: 'drive_sync' },
    approval: { url: 'https://drive.google.com/file/d/REP_SYNC_10/view', source: 'drive_sync' },
  },
};
const normBoth1 = api._normalizeRowAttachments(rowBothSynced, 'DS', rowBothSynced.fileLocation);
assert.ok(normBoth1.submission && normBoth1.submission.url.includes('SUB_SYNC_10'));
assert.ok(normBoth1.approval && normBoth1.approval.url.includes('REP_SYNC_10'));

// ── Criterion 2b: Legacy Submit + Reply -> Submit is NOT cleared to null ──
const rowLegacySubPlusReply = {
  _id: 'R-BOTH-LEG',
  docNo: 'DS-011',
  status: 'B - Approved As Noted',
  fileLocation: 'https://drive.google.com/file/d/SUB_LEG_11/view',
  attachments: {
    submission: { url: 'https://drive.google.com/file/d/SUB_LEG_11/view', legacy: true },
    approval: { file_id: 'REP_UP_11', file_name: 'DS-011_Reply.pdf', url: 'https://drive.google.com/file/d/REP_UP_11/view', uploaded_at: '2026-11-25T09:00:00Z' },
  },
};
const normBoth2 = api._normalizeRowAttachments(rowLegacySubPlusReply, 'DS', rowLegacySubPlusReply.fileLocation);
assert.ok(normBoth2.submission && normBoth2.submission.url.includes('SUB_LEG_11'), 'Legacy submission must be preserved when approval exists');
assert.ok(normBoth2.approval && normBoth2.approval.file_id === 'REP_UP_11');

// Also when legacy record had only fileLocation (no attachments.submission) and DCR approval was uploaded:
const rowFileLocPlusReply = {
  _id: 'R-BOTH-FL',
  docNo: 'DS-012',
  status: 'A - Approved',
  fileLocation: 'https://drive.google.com/file/d/SUB_FL_12/view',
  attachments: {
    approval: { file_id: 'REP_UP_12', file_name: 'DS-012_Reply.pdf', url: 'https://drive.google.com/file/d/REP_UP_12/view', uploaded_at: '2026-11-25T09:00:00Z' },
  },
};
const normBoth3 = api._normalizeRowAttachments(rowFileLocPlusReply, 'DS', rowFileLocPlusReply.fileLocation);
assert.ok(normBoth3.submission && normBoth3.submission.url.includes('SUB_FL_12'), 'Legacy fileLocation submission must be preserved alongside new approval');
assert.ok(normBoth3.approval && normBoth3.approval.file_id === 'REP_UP_12');

// Table renders BOTH Subm and Reply chips
api.state.recs = [rowLegacySubPlusReply];
api.state.visibleRows = [rowLegacySubPlusReply];
api.renderRows();
const tr2 = byId['tbody'].children[0];
const chips2 = walkEls(tr2).filter(el => el.textContent === 'Subm' || el.textContent === 'Reply');
assert.deepStrictEqual(chips2.map(c => c.textContent), ['Subm', 'Reply'], 'Table must render BOTH Subm and Reply chips');

// Modal renders files in BOTH Submission and Approval cards
const formRoot2 = mkEl('div');
api.state.editId = 'R-BOTH-LEG';
api.mountDriveAttachmentsBox(formRoot2, {}, rowLegacySubPlusReply);
const submInfo2 = walkEls(formRoot2).find(e => e.id === 'drv-slot-info-submission');
const apprInfo2 = walkEls(formRoot2).find(e => e.id === 'drv-slot-info-approval');
assert.ok(!walkEls(submInfo2).map(e => e.textContent).join(' | ').includes('No submission file attached'));
assert.ok(!walkEls(apprInfo2).map(e => e.textContent).join(' | ').includes('No reply file attached'));
"""

        with tempfile.NamedTemporaryFile("w", suffix=".js", encoding="utf-8", delete=False) as jf, \
             tempfile.NamedTemporaryFile("w", suffix=".js", encoding="utf-8", delete=False) as tf:
            jf.write(main_script)
            main_js_path = jf.name
            tf.write(node_test)
            test_js_path = tf.name

        try:
            proc = subprocess.run(
                ["node", test_js_path, main_js_path],
                capture_output=True,
                text=True,
                timeout=30,
            )
            self.assertEqual(
                proc.returncode,
                0,
                f"Node frontend test failed:\nSTDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}",
            )
        finally:
            Path(main_js_path).unlink(missing_ok=True)
            Path(test_js_path).unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
