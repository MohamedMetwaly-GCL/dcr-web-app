"""Offline unit and Node VM tests for DRIVE-UX-02:
1. drive_upload.DEFAULT_CHUNK == 2 MiB (exact multiple of 256 KiB CHUNK_UNIT).
2. XMLHttpRequest-based _xhrSendBlob and _sendDriveChunkWithRetry:
   - continuous byte progress (99% direct cap, 95% proxy cap)
   - proxy 'forwarding' stage on xhr.upload.onload
   - AbortSignal cancellation calling xhr.abort()
   - automatic direct-to-proxy fallback on attempt 0 network/CORS failure
3. Bilingual stage text (_formatDriveStageText) for preparing/uploading/forwarding/linking.
4. Byte-weighted folder & batch progress math (updateBatchOverallProgress, folderInp.onchange, startBatchDriveUpload).
5. Error persistence in mountDriveAttachmentsBox (progWrap stays visible on upload error).
"""
import json
import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("DATABASE_URL", "")
os.environ.setdefault("SECRET_KEY", "synthetic-test-secret-key-drive-ux-02")
os.environ.setdefault("MAGIC_SECRET", "synthetic-test-magic-secret-drive-ux-02")
os.environ.setdefault("GEMINI_API_KEY", "synthetic-test-gemini-key-drive-ux-02")

import drive_upload
import html_render


def _render_register_html(lang="en"):
    user = {"id": "u1", "username": "testuser", "email": "test@example.com", "name": "Test User", "role": "admin", "is_superadmin": True}
    proj = {"id": "p1", "code": "PRJ-1", "name": "Test Project", "workspace_id": "ws1"}
    dt_list = [{"id": "DS", "code": "DS", "name": "Document Submittal", "project_id": "p1", "sort_order": 0}]
    with patch("html_render.can_edit", return_value=True), \
         patch("db.can_user_download", return_value=True), \
         patch("db.get_projects_for_user", return_value=[proj]), \
         patch("db.get_doc_types", return_value=dt_list), \
         patch("db.get_logo", return_value=None), \
         patch("db.get_user_workspaces", return_value=[{"id": "ws1", "name": "Engineering"}]), \
         patch("db.is_user_workspace_admin", return_value=True):
        return html_render.render_register(user, proj, active_ws_id="ws1", lang=lang)


def _extract_js_function(html_text, start_marker, end_marker):
    s = html_text.index(start_marker)
    e = html_text.index(end_marker, s)
    return html_text[s:e]


class TestDriveUploadProgress(unittest.TestCase):
    def test_default_chunk_is_2mb_and_multiple_of_256kib(self):
        self.assertEqual(drive_upload.DEFAULT_CHUNK, 2 * 1024 * 1024)
        self.assertEqual(drive_upload.CHUNK_UNIT, 256 * 1024)
        self.assertEqual(drive_upload.DEFAULT_CHUNK % drive_upload.CHUNK_UNIT, 0)

    def test_rendered_js_progress_stages_and_error_persistence(self):
        html_en = _render_register_html(lang="en")
        self.assertIn("2*1024*1024", html_en)
        self.assertIn("new XMLHttpRequest()", html_en)
        self.assertIn("xhr.upload.onprogress", html_en)
        self.assertIn("xhr.upload.onload", html_en)
        self.assertIn("if(progWrap.dataset.hasError!=='1')progWrap.style.display='none'", html_en)
        self.assertIn("if(progWrap.dataset.hasError !== '1') progWrap.style.display = 'none'", html_en)
        self.assertIn("totalFolderBytes", html_en)
        self.assertIn("uploadedFolderBytes", html_en)

    def test_node_vm_xhr_progress_fallback_abort_stages_and_batch_math(self):
        html_en = _render_register_html(lang="en")
        fmt_bytes_js = _extract_js_function(html_en, "function _fmtBytes(", "\nfunction ")
        stage_and_xhr_js = _extract_js_function(
            html_en,
            "function _isArLang()",
            "\nfunction promptDriveFolderChoice(",
        )
        batch_overall_js = _extract_js_function(
            html_en,
            "function updateBatchOverallProgress()",
            "\nfunction openBatchDriveUpload(",
        )

        node_script = f"""
const assert = require('assert');

// Minimal browser globals for Node VM
global.window = {{ DCR_LANG: 'en', t: (k, v) => k + ':' + JSON.stringify(v || {{}}) }};
global.document = {{
  documentElement: {{ getAttribute: () => (global.window.DCR_LANG === 'ar' ? 'ar' : 'en') }},
  _els: {{}},
  getElementById: (id) => global.document._els[id] || null,
}};

{fmt_bytes_js}
{stage_and_xhr_js}
{batch_overall_js}

// Mock XMLHttpRequest factory
let xhrInstances = [];
let nextXhrBehavior = [];

class MockXMLHttpRequest {{
  constructor() {{
    this.headers = {{}};
    this.upload = {{ onprogress: null, onload: null }};
    this.status = 0;
    this.responseText = '';
    this.aborted = false;
    this._respHeaders = {{}};
    xhrInstances.push(this);
  }}
  open(method, url) {{
    this.method = method;
    this.url = url;
  }}
  setRequestHeader(k, v) {{
    this.headers[k] = v;
  }}
  getResponseHeader(k) {{
    return this._respHeaders[k] || null;
  }}
  abort() {{
    this.aborted = true;
    if (this.onabort) this.onabort();
  }}
  send(body) {{
    this.body = body;
    const fn = nextXhrBehavior.shift();
    if (fn) fn(this);
  }}
}}
global.XMLHttpRequest = MockXMLHttpRequest;

function makeFakeFile(size) {{
  return {{
    size,
    name: 'sample.pdf',
    slice(start, end) {{
      return {{ size: Math.max(0, end - start), start, end }};
    }}
  }};
}}

(async () => {{
  // 1. Bilingual stage text
  global.window.DCR_LANG = 'en';
  assert.strictEqual(_formatDriveStageText('preparing', 0, 100, ''), 'Preparing Drive folder...');
  assert.strictEqual(_formatDriveStageText('forwarding', 50, 100, ''), 'Transferring to Google Drive...');
  assert.strictEqual(_formatDriveStageText('linking', 100, 100, ''), 'Linking file to record...');
  assert(_formatDriveStageText('uploading', 50, 100, '').includes('Uploading...'));

  global.window.DCR_LANG = 'ar';
  assert.strictEqual(_formatDriveStageText('preparing', 0, 100, ''), 'جاري تجهيز مجلد Drive...');
  assert.strictEqual(_formatDriveStageText('forwarding', 50, 100, ''), 'جاري النقل إلى Google Drive...');
  assert.strictEqual(_formatDriveStageText('linking', 100, 100, ''), 'جاري ربط الملف بالسجل...');
  assert(_formatDriveStageText('uploading', 50, 100, '').includes('جاري الرفع...'));
  global.window.DCR_LANG = 'en';

  // 2. Direct XHR byte progress capped at 99% until Drive 200 resolves
  xhrInstances = [];
  nextXhrBehavior = [
    (xhr) => {{
      setTimeout(() => {{
        xhr.upload.onprogress({{ lengthComputable: true, loaded: 500, total: 1000 }});
        xhr.upload.onprogress({{ lengthComputable: true, loaded: 1000, total: 1000 }});
        xhr.status = 200;
        xhr.responseText = JSON.stringify({{ id: 'DRV_DIRECT_1' }});
        xhr.onload();
      }}, 5);
    }}
  ];
  const directEvents = [];
  const directRes = await _sendDriveChunkWithRetry(
    'https://upload.drive/sess',
    't1',
    makeFakeFile(1000),
    0,
    1000,
    (done, total, lbl, stage) => directEvents.push({{ done, total, stage }}),
    {{ proxy: false }},
    null
  );
  assert.strictEqual(directRes.done, true);
  assert.strictEqual(directRes.fileId, 'DRV_DIRECT_1');
  assert.strictEqual(directEvents[0].done, 500);
  assert.strictEqual(directEvents[1].done, 990); // 99% cap
  assert.strictEqual(directEvents[2].done, 1000); // confirmed 100%

  // 3. Direct failure on attempt 0 -> automatic proxy fallback, 95% cap, 'forwarding' stage
  xhrInstances = [];
  nextXhrBehavior = [
    (xhr) => {{
      setTimeout(() => {{ xhr.onerror(); }}, 5);
    }},
    (xhr) => {{
      setTimeout(() => {{
        xhr.upload.onprogress({{ lengthComputable: true, loaded: 1000, total: 1000 }});
        xhr.upload.onload();
        xhr.status = 200;
        xhr.responseText = JSON.stringify({{ done: true, file: {{ id: 'DRV_PROXY_1' }} }});
        xhr.onload();
      }}, 5);
    }}
  ];
  const proxyEvents = [];
  const proxyRef = {{ proxy: false }};
  const proxyRes = await _sendDriveChunkWithRetry(
    'https://upload.drive/sess',
    't2',
    makeFakeFile(1000),
    0,
    1000,
    (done, total, lbl, stage) => proxyEvents.push({{ done, total, stage }}),
    proxyRef,
    null
  );
  assert.strictEqual(proxyRef.proxy, true);
  assert.strictEqual(proxyRes.fileId, 'DRV_PROXY_1');
  assert.strictEqual(xhrInstances[1].url, '/api/drive/upload-chunk');
  assert.strictEqual(xhrInstances[1].method, 'POST');
  assert.strictEqual(proxyEvents[0].done, 950); // 95% cap in proxy mode
  assert.strictEqual(proxyEvents[1].stage, 'forwarding');
  assert.strictEqual(proxyEvents[2].done, 1000);

  // 4. AbortSignal cancellation calls xhr.abort() and rejects with AbortError
  xhrInstances = [];
  const ac = new AbortController();
  nextXhrBehavior = [
    (xhr) => {{
      setTimeout(() => {{ ac.abort(); }}, 5);
    }}
  ];
  let abortCaught = null;
  try {{
    await _sendDriveChunkWithRetry(
      'https://upload.drive/sess',
      't3',
      makeFakeFile(1000),
      0,
      1000,
      () => {{}},
      {{ proxy: false }},
      ac.signal
    );
  }} catch (e) {{
    abortCaught = e;
  }}
  assert(abortCaught && abortCaught.name === 'AbortError');
  assert.strictEqual(xhrInstances[0].aborted, true);

  // 5. Byte-weighted batch overall progress (done + activeFrac)
  const ovBar = {{ style: {{ width: '0%' }} }};
  const ovPct = {{ textContent: '' }};
  const ovTxt = {{ textContent: '' }};
  global.document._els = {{
    'batch-drv-overall-bar': ovBar,
    'batch-drv-overall-pct': ovPct,
    'batch-drv-overall-txt': ovTxt,
  }};
  global.state = {{
    batchQueue: [
      {{ recordId: 'r1', status: 'done', pct: 100 }},
      {{ recordId: 'r2', status: 'uploading', pct: 50 }},
    ]
  }};
  updateBatchOverallProgress();
  assert.strictEqual(ovBar.style.width, '75%');
  assert.strictEqual(ovPct.textContent, '75%');

  console.log('NODE_VM_DRIVE_PROGRESS_OK');
}})().catch((err) => {{
  console.error(err);
  process.exit(1);
}});
"""
        with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as tf:
            tf.write(node_script)
            tf_path = tf.name
        try:
            proc = subprocess.run(["node", tf_path], capture_output=True, text=True, check=False)
            self.assertEqual(proc.returncode, 0, msg=f"Node VM failed:\nSTDOUT:{proc.stdout}\nSTDERR:{proc.stderr}")
            self.assertIn("NODE_VM_DRIVE_PROGRESS_OK", proc.stdout)
        finally:
            try:
                Path(tf_path).unlink(missing_ok=True)
            except Exception:
                pass


if __name__ == "__main__":
    unittest.main()
