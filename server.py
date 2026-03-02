"""
server.py - DCR Web Application
Render + Supabase Production  |  Local SQLite fallback
"""
import json, os, sys, uuid, csv, io, re, base64, datetime, socket
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote

sys.path.insert(0, str(Path(__file__).parent))
from modules import database as db
from modules.utils import (compute_expected_reply, compute_duration, is_overdue,
                            format_date, get_next_doc_no, extract_rev, extract_seq,
                            build_doc_no, STATUS_COLORS)

PORT      = int(os.environ.get("PORT", 5000))
IS_RENDER = bool(os.environ.get("RENDER", ""))
STATIC_DIR   = Path(__file__).parent / "static"
TEMPLATE_DIR = Path(__file__).parent / "templates"


def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]; s.close(); return ip
    except: return "127.0.0.1"


def compute_row_value(row, col_key):
    """Compute display value for auto fields."""
    if col_key == 'expectedReplyDate':
        return format_date(compute_expected_reply(row.get('issuedDate'), row.get('docNo')))
    if col_key == 'duration':
        v = compute_duration(row.get('issuedDate'), row.get('actualReplyDate'))
        return str(v) if v is not None else ''
    if col_key in ('issuedDate','actualReplyDate'):
        return format_date(row.get(col_key,''))
    return str(row.get(col_key,'') or '')


class DCRHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # suppress default logging

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', len(body))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def send_html(self, html, status=200):
        body = html.encode('utf-8') if isinstance(html, str) else html
        self.send_response(status)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', len(body))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path):
        ext = Path(path).suffix.lower()
        mime = {'.css':'text/css','.js':'application/javascript',
                '.png':'image/png','.jpg':'image/jpeg','.ico':'image/x-icon',
                '.svg':'image/svg+xml'}.get(ext, 'application/octet-stream')
        try:
            with open(path, 'rb') as f: body = f.read()
            self.send_response(200)
            self.send_header('Content-Type', mime)
            self.send_header('Content-Length', len(body))
            self.send_header('Cache-Control', 'max-age=3600')
            self.end_headers()
            self.wfile.write(body)
        except FileNotFoundError:
            self.send_response(404); self.end_headers()

    def read_body(self):
        length = int(self.headers.get('Content-Length', 0))
        return self.rfile.read(length) if length else b''

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip('/')
        qs = parse_qs(parsed.query)

        # Static files
        if path.startswith('/static/'):
            self.send_file(STATIC_DIR / path[8:]); return

        # ── Ping endpoint — keeps Render alive via UptimeRobot ──
        if path == '/ping':
            self.send_json({"status": "ok", "time": datetime.datetime.now().isoformat()}); return

        # Main app
        if path in ('', '/'):
            self.send_html(build_main_page()); return

        # API routes
        if path == '/api/project':
            self.send_json(db.get_project()); return

        if path == '/api/doc_types':
            self.send_json(db.get_doc_types()); return

        if path.startswith('/api/records/'):
            dt_id = path.split('/')[-1]
            search = qs.get('search', [''])[0]
            records = db.get_records(dt_id, search=search)
            cols = db.get_columns(dt_id)
            # Enrich with computed fields
            for row in records:
                row['_expectedReplyDate'] = format_date(compute_expected_reply(row.get('issuedDate'), row.get('docNo')))
                row['_duration'] = str(compute_duration(row.get('issuedDate'), row.get('actualReplyDate')) or '')
                row['_overdue'] = is_overdue(row.get('issuedDate'), row.get('docNo'), row.get('actualReplyDate'))
                row['_isRev'] = extract_rev(row.get('docNo','')) > 0
                row['_issuedDateFmt'] = format_date(row.get('issuedDate',''))
                row['_actualReplyDateFmt'] = format_date(row.get('actualReplyDate',''))
            self.send_json({'records': records, 'columns': cols,
                            'count': db.get_record_count(dt_id)}); return

        if path == '/api/columns':
            dt_id = qs.get('dt', [''])[0]
            self.send_json(db.get_columns(dt_id)); return

        if path == '/api/dropdown_lists':
            self.send_json(db.get_all_dropdown_lists()); return

        if path.startswith('/api/next_doc_no/'):
            parts = path.split('/')
            dt_id = parts[-1]
            dt = next((d for d in db.get_doc_types() if d['id'] == dt_id), None)
            prefix = dt['code'] if dt else dt_id
            records = db.get_records(dt_id)
            self.send_json({'next': get_next_doc_no(prefix, records)}); return

        if path == '/api/counts':
            counts = {dt['id']: db.get_record_count(dt['id']) for dt in db.get_doc_types()}
            self.send_json(counts); return

        if path.startswith('/api/logo/'):
            key = unquote(path.split('/')[-1])
            data = db.get_logo(key)
            if data:
                raw = base64.b64decode(data)
                self.send_response(200)
                self.send_header('Content-Type', 'image/png')
                self.send_header('Content-Length', len(raw))
                self.end_headers()
                self.wfile.write(raw)
            else:
                self.send_response(404); self.end_headers()
            return

        if path.startswith('/api/export/'):
            dt_id = path.split('/')[-1]
            self._handle_export(dt_id); return

        self.send_response(404); self.end_headers()

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip('/')
        body = self.read_body()

        try:
            data = json.loads(body) if body else {}
        except Exception:
            data = {}

        if path == '/api/project':
            db.save_project(data)
            self.send_json({'ok': True}); return

        if path.startswith('/api/records/'):
            dt_id = path.split('/')[-1]
            rec_id = data.pop('_id', None) or str(uuid.uuid4())
            # Remove computed fields
            for k in list(data.keys()):
                if k.startswith('_'): del data[k]
            db.save_record(dt_id, rec_id, data)
            self.send_json({'ok': True, 'id': rec_id}); return

        if path.startswith('/api/delete_record/'):
            rec_id = path.split('/')[-1]
            db.delete_record(rec_id)
            self.send_json({'ok': True}); return

        if path == '/api/doc_types':
            name = data.get('name','').strip()
            code = data.get('code','').strip().upper()
            if name and code:
                db.add_doc_type(code, name, code)
                self.send_json({'ok': True}); return
            self.send_json({'ok': False, 'error': 'Name and code required'}, 400); return

        if path.startswith('/api/delete_doc_type/'):
            dt_id = path.split('/')[-1]
            db.delete_doc_type(dt_id)
            self.send_json({'ok': True}); return

        if path == '/api/dropdown_lists/add':
            db.add_dropdown_item(data.get('list_name',''), data.get('item',''))
            self.send_json({'ok': True}); return

        if path == '/api/dropdown_lists/remove':
            db.remove_dropdown_item(data.get('list_name',''), data.get('item',''))
            self.send_json({'ok': True}); return

        if path == '/api/columns/add':
            db.add_column(data['dt_id'], data['col_key'], data['label'],
                          data['col_type'], data.get('list_name'))
            self.send_json({'ok': True}); return

        if path.startswith('/api/columns/visibility/'):
            col_id = int(path.split('/')[-1])
            db.update_col_visibility(col_id, data.get('visible', True))
            self.send_json({'ok': True}); return

        if path.startswith('/api/columns/delete/'):
            col_id = int(path.split('/')[-1])
            db.delete_column(col_id)
            self.send_json({'ok': True}); return

        if path == '/api/logo':
            key = data.get('key','')
            img_data = data.get('data','')
            db.save_logo(key, img_data)
            self.send_json({'ok': True}); return

        if path == '/api/import_csv':
            self._handle_import(data.get('dt_id',''), data.get('csv_text','')); return

        self.send_response(404); self.end_headers()

    def _handle_export(self, dt_id):
        dt = next((d for d in db.get_doc_types() if d['id'] == dt_id), None)
        cols = [c for c in db.get_columns(dt_id) if c['visible']]
        records = db.get_records(dt_id)
        proj = db.get_project()

        out = io.StringIO()
        w = csv.writer(out)
        w.writerow(['Document Control Register'])
        w.writerow(['Project Code:', proj.get('code','')])
        w.writerow(['Project Name:', proj.get('name','')])
        w.writerow(['Client:', proj.get('client','')])
        w.writerow(['Main Consultant:', proj.get('mainConsultant','')])
        w.writerow(['Contractor:', proj.get('contractor','')])
        w.writerow(['Document Type:', dt['name'] if dt else dt_id])
        w.writerow(['Export Date:', datetime.datetime.now().strftime('%d/%b/%Y %H:%M')])
        w.writerow(['Total Records:', str(len(records))])
        w.writerow([])
        w.writerow(['Sr.'] + [c['label'] for c in cols])

        sr = 1
        for row in records:
            is_rev = extract_rev(row.get('docNo','')) > 0
            vals = ['' if is_rev else str(sr)]
            for col in cols:
                key = col['col_key']
                if key == 'expectedReplyDate':
                    vals.append(format_date(compute_expected_reply(row.get('issuedDate'), row.get('docNo'))))
                elif key == 'duration':
                    v = compute_duration(row.get('issuedDate'), row.get('actualReplyDate'))
                    vals.append(str(v) if v is not None else '')
                elif key in ('issuedDate','actualReplyDate'):
                    vals.append(format_date(row.get(key,'')))
                else:
                    vals.append(str(row.get(key,'') or ''))
            w.writerow(vals)
            if not is_rev: sr += 1

        csv_text = '\ufeff' + out.getvalue()
        body = csv_text.encode('utf-8')
        fname = f"{proj.get('code','DCR')}_{dt_id}_Register.csv"
        self.send_response(200)
        self.send_header('Content-Type', 'text/csv; charset=utf-8')
        self.send_header('Content-Disposition', f'attachment; filename="{fname}"')
        self.send_header('Content-Length', len(body))
        self.end_headers()
        self.wfile.write(body)

    def _handle_import(self, dt_id, csv_text):
        cols = db.get_columns(dt_id)
        col_map = {c['label']: c['col_key'] for c in cols}
        imported = 0
        reader = csv.reader(io.StringIO(csv_text))
        header = None
        for line in reader:
            if not header:
                if any(c in line for c in ['Document No.','Sr.','docNo']):
                    header = [col_map.get(h.strip(), h.strip()) for h in line]
                continue
            if not any(line): continue
            row_data = {}
            for i, val in enumerate(line):
                if i < len(header) and header[i] and header[i] not in ('sr','Sr.',''):
                    row_data[header[i]] = val.strip()
            if row_data:
                db.save_record(dt_id, str(uuid.uuid4()), row_data)
                imported += 1
        self.send_json({'ok': True, 'imported': imported})


# ============================================================
# BUILD MAIN HTML PAGE
# ============================================================
def build_main_page():
    proj = db.get_project()
    doc_types = db.get_doc_types()
    all_lists = db.get_all_dropdown_lists()

    status_colors_js = json.dumps(STATUS_COLORS)

    tabs_html = ''.join(
        f'<button class="tab-btn" data-id="{dt["id"]}" onclick="switchTab(\'{dt["id"]}\')">'
        f'<span class="tab-code">{dt["code"]}</span>'
        f'<span class="tab-count" id="cnt-{dt["id"]}">0</span></button>'
        for dt in doc_types
    )

    proj_fields_html = ''.join(
        f'<div class="pf"><span class="pf-lbl">{lbl}</span><span class="pf-val" id="pf-{key}">{proj.get(key,"—")}</span></div>'
        for key, lbl in [('code','Code'),('name','Project Name'),('client','Client'),
                          ('mainConsultant','Consultant'),('mepConsultant','MEP'),
                          ('contractor','Contractor'),('startDate','Start'),('endDate','End'),
                          ('pmo','PMO'),('landlord','Landlord')]
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Document Control Register</title>
<style>
:root{{
  --primary:#1a3a5c; --primary-lt:#2563a8; --accent:#f0a500;
  --bg:#f0f4f8; --sidebar:#0f2640; --white:#fff; --border:#dde3ed;
  --text:#1e2a3a; --muted:#6b7a94; --success:#16a34a; --danger:#ef4444;
  --warning:#f59e0b; --radius:6px;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--text);font-size:13px;height:100vh;display:flex;flex-direction:column;overflow:hidden}}

/* TOP BAR */
#topbar{{background:var(--primary);color:#fff;height:46px;display:flex;align-items:center;padding:0 16px;gap:10px;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,.25);z-index:100}}
#topbar .logo{{font-size:20px}}
#topbar .app-title{{font-weight:700;font-size:14px;letter-spacing:.3px}}
#topbar .spacer{{flex:1}}
.tb-btn{{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;padding:5px 12px;border-radius:var(--radius);cursor:pointer;font-size:12px;font-family:inherit}}
.tb-btn:hover{{background:rgba(255,255,255,.25)}}

/* PROJECT BAR */
#projbar{{background:var(--white);border-bottom:2px solid var(--primary);padding:6px 14px;display:flex;align-items:center;gap:0;flex-wrap:wrap;flex-shrink:0}}
.pf{{display:flex;flex-direction:column;padding:0 10px;border-right:1px solid var(--border)}}
.pf:last-of-type{{border-right:none}}
.pf-lbl{{font-size:9px;font-weight:700;color:var(--primary);text-transform:uppercase;letter-spacing:.4px}}
.pf-val{{font-size:11px;color:var(--text);font-weight:500;white-space:nowrap}}
#proj-edit-btn{{margin-left:auto;background:var(--primary);color:#fff;border:none;padding:5px 12px;border-radius:var(--radius);cursor:pointer;font-size:11px;font-family:inherit;flex-shrink:0}}
#proj-edit-btn:hover{{background:var(--primary-lt)}}

/* TABS */
#tabs-bar{{background:var(--sidebar);display:flex;align-items:center;overflow-x:auto;flex-shrink:0;padding:0 8px;scrollbar-width:thin}}
#tabs-bar::-webkit-scrollbar{{height:3px}}
#tabs-bar::-webkit-scrollbar-thumb{{background:rgba(255,255,255,.3)}}
.tab-btn{{display:flex;align-items:center;gap:6px;padding:9px 13px;background:transparent;border:none;border-bottom:2px solid transparent;color:rgba(255,255,255,.55);cursor:pointer;font-family:inherit;font-size:11px;font-weight:600;white-space:nowrap;transition:all .15s}}
.tab-btn:hover{{color:#fff;background:rgba(255,255,255,.08)}}
.tab-btn.active{{color:#fff;border-bottom-color:var(--accent)}}
.tab-code{{font-size:11px}}
.tab-count{{background:rgba(255,255,255,.2);border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700}}
.tab-btn.active .tab-count{{background:var(--accent);color:#000}}
.tab-add{{padding:6px 10px;background:rgba(255,255,255,.1);border:1px dashed rgba(255,255,255,.35);color:rgba(255,255,255,.7);border-radius:4px;cursor:pointer;font-size:16px;margin-left:6px;flex-shrink:0}}
.tab-add:hover{{background:rgba(255,255,255,.2)}}

/* TOOLBAR */
#toolbar{{background:var(--white);border-bottom:1px solid var(--border);padding:6px 12px;display:flex;align-items:center;gap:6px;flex-shrink:0;flex-wrap:wrap}}
.tool-btn{{display:flex;align-items:center;gap:4px;padding:5px 11px;background:var(--bg);border:1px solid var(--border);border-radius:var(--radius);cursor:pointer;font-size:11px;font-family:inherit;color:var(--text);transition:all .15s;white-space:nowrap}}
.tool-btn:hover{{background:var(--primary);color:#fff;border-color:var(--primary)}}
.tool-btn.green:hover{{background:#059669;border-color:#059669}}
.tool-btn.purple:hover{{background:#7c3aed;border-color:#7c3aed}}
.tool-btn.teal:hover{{background:#0891b2;border-color:#0891b2}}
#search-box{{flex:1;min-width:180px;max-width:280px;padding:5px 10px 5px 30px;border:1px solid var(--border);border-radius:var(--radius);font-family:inherit;font-size:12px;outline:none;background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='%236b7a94' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E") no-repeat 8px center}}
#search-box:focus{{border-color:var(--primary-lt);box-shadow:0 0 0 2px rgba(37,99,168,.12)}}

/* MAIN TABLE AREA */
#main{{flex:1;overflow:hidden;display:flex;flex-direction:column}}
#table-wrap{{flex:1;overflow:auto}}

table{{width:100%;border-collapse:collapse;min-width:1100px;font-size:12px}}
thead{{position:sticky;top:0;z-index:10}}
th{{background:var(--primary);color:#fff;padding:8px 8px;text-align:left;font-weight:600;white-space:nowrap;border-right:1px solid rgba(255,255,255,.12);cursor:pointer;user-select:none}}
th:hover{{background:var(--primary-lt)}}
th.sort-asc::after{{content:" ↑"}}
th.sort-desc::after{{content:" ↓"}}
.filter-row th{{background:#eef1f7;padding:3px 4px;cursor:default}}
.filter-row th:hover{{background:#eef1f7}}
.filter-row input,.filter-row select{{width:100%;padding:3px 5px;border:1px solid var(--border);border-radius:3px;font-size:10px;font-family:inherit;background:#fff;outline:none}}
td{{padding:5px 8px;border-bottom:1px solid #edf0f5;border-right:1px solid #edf0f5;vertical-align:top;max-width:220px;word-break:break-word}}
tr:hover td{{background:rgba(37,99,168,.035)}}
tr.overdue td{{background:#fff5f5}}
tr.overdue:hover td{{background:#fee2e2}}
tr.rev-row td{{color:var(--muted)}}
tr.alt td{{background:#fafbfd}}
.sr-cell{{text-align:center;color:var(--muted);font-size:10px;width:36px}}
.actions-cell{{white-space:nowrap;width:70px}}
.act-btn{{padding:2px 7px;border:1px solid var(--border);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.act-btn:hover{{background:var(--primary);color:#fff;border-color:var(--primary)}}
.act-btn.del:hover{{background:var(--danger);border-color:var(--danger)}}
.status-badge{{display:inline-block;border-radius:3px;padding:2px 7px;font-size:10px;font-weight:700;margin:1px}}
.file-link{{color:var(--primary-lt);text-decoration:underline;cursor:pointer;font-size:11px}}
.overdue-date{{color:#dc2626;font-weight:700}}

/* STATUS BAR */
#statusbar{{background:var(--primary);color:rgba(255,255,255,.75);padding:3px 14px;font-size:10px;display:flex;gap:18px;flex-shrink:0}}
#status-overdue{{color:#fca5a5}}
#status-clock{{margin-left:auto}}

/* MODALS */
.overlay{{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:1000;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px)}}
.overlay.hidden{{display:none}}
.modal{{background:#fff;border-radius:10px;box-shadow:0 24px 64px rgba(0,0,0,.3);width:90%;max-width:680px;max-height:90vh;display:flex;flex-direction:column;overflow:hidden;animation:modalIn .2s ease}}
@keyframes modalIn{{from{{transform:translateY(-20px);opacity:0}}to{{transform:none;opacity:1}}}}
.modal.wide{{max-width:850px}}
.modal-hdr{{background:var(--primary);color:#fff;padding:13px 18px;display:flex;align-items:center;justify-content:space-between;font-weight:700;font-size:13px;flex-shrink:0}}
.modal-close{{background:none;border:none;color:#fff;font-size:20px;cursor:pointer;opacity:.7;line-height:1}}
.modal-close:hover{{opacity:1}}
.modal-body{{padding:18px;overflow-y:auto;flex:1}}
.modal-footer{{padding:10px 18px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:8px;background:var(--bg);flex-shrink:0}}

/* FORMS */
.form-grid{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
.form-group{{display:flex;flex-direction:column;gap:4px}}
.form-group.full{{grid-column:1/-1}}
.form-group label{{font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.4px}}
.form-group input,.form-group select,.form-group textarea{{padding:7px 10px;border:1px solid var(--border);border-radius:var(--radius);font-family:inherit;font-size:12px;outline:none}}
.form-group input:focus,.form-group select:focus,.form-group textarea:focus{{border-color:var(--primary-lt);box-shadow:0 0 0 2px rgba(37,99,168,.1)}}
.form-group textarea{{min-height:60px;resize:vertical}}
.seq-warn{{background:#fef3c7;border:1px solid #f59e0b;border-radius:var(--radius);padding:7px 10px;font-size:11px;color:#92400e;display:none;margin-bottom:8px}}
.auto-info{{background:#f0f4fa;border-radius:var(--radius);padding:8px 12px;font-size:11px;color:var(--primary);margin-top:8px}}

/* BUTTONS */
.btn{{padding:7px 16px;border-radius:var(--radius);cursor:pointer;font-family:inherit;font-size:12px;font-weight:600;border:1px solid transparent}}
.btn-primary{{background:var(--primary);color:#fff}}
.btn-primary:hover{{background:var(--primary-lt)}}
.btn-secondary{{background:var(--bg);color:var(--text);border-color:var(--border)}}
.btn-secondary:hover{{background:var(--border)}}
.btn-danger{{background:var(--danger);color:#fff}}
.btn-success{{background:var(--success);color:#fff}}
.btn-sm{{padding:4px 10px;font-size:11px}}

/* TOAST */
#toast{{position:fixed;bottom:40px;right:20px;background:var(--primary);color:#fff;padding:10px 18px;border-radius:var(--radius);font-size:12px;z-index:9999;box-shadow:0 4px 16px rgba(0,0,0,.2);transform:translateY(80px);opacity:0;transition:all .3s;pointer-events:none}}
#toast.show{{transform:none;opacity:1}}
#toast.success{{background:#16a34a}}
#toast.error{{background:var(--danger)}}
#toast.warning{{background:var(--warning);color:#000}}

/* Settings list */
.settings-list{{list-style:none;display:flex;flex-direction:column;gap:4px;max-height:200px;overflow-y:auto;border:1px solid var(--border);border-radius:var(--radius);padding:4px}}
.settings-item{{display:flex;align-items:center;gap:8px;padding:4px 8px;background:var(--bg);border-radius:3px;font-size:11px}}
.settings-item .name{{flex:1}}
.settings-item button{{padding:2px 8px;font-size:10px;border:1px solid var(--border);background:#fff;border-radius:3px;cursor:pointer}}
.settings-item button:hover{{background:var(--danger);color:#fff;border-color:var(--danger)}}
.section-title{{font-size:11px;font-weight:700;color:var(--primary);text-transform:uppercase;letter-spacing:.5px;margin:14px 0 6px;padding-bottom:4px;border-bottom:2px solid var(--primary)}}
.add-item-row{{display:flex;gap:6px;margin-top:6px}}
.add-item-row input{{flex:1;padding:5px 8px;border:1px solid var(--border);border-radius:3px;font-size:11px;font-family:inherit;outline:none}}

/* Multi-select tags */
.ms-container{{border:1px solid var(--border);border-radius:var(--radius);min-height:36px;padding:4px;background:#fff;cursor:pointer;position:relative}}
.ms-container:focus-within{{border-color:var(--primary-lt)}}
.ms-tag{{display:inline-flex;align-items:center;gap:3px;background:var(--primary);color:#fff;border-radius:3px;padding:2px 7px;font-size:10px;margin:2px}}
.ms-tag .rm{{cursor:pointer;opacity:.7}}
.ms-tag .rm:hover{{opacity:1}}
.ms-placeholder{{color:var(--muted);font-size:11px;padding:2px 4px}}
.ms-dropdown{{position:absolute;left:0;right:0;top:100%;background:#fff;border:1px solid var(--border);border-radius:var(--radius);z-index:200;max-height:200px;overflow-y:auto;box-shadow:0 4px 16px rgba(0,0,0,.12);margin-top:2px}}
.ms-option{{padding:7px 10px;cursor:pointer;font-size:11px;display:flex;align-items:center;gap:7px}}
.ms-option:hover{{background:var(--bg)}}
.ms-option.sel{{background:#eff6ff}}

/* Empty state */
.empty{{text-align:center;padding:60px 20px;color:var(--muted)}}
.empty .ico{{font-size:48px;margin-bottom:12px}}
</style>
</head>
<body>

<!-- TOP BAR -->
<div id="topbar">
  <span class="logo">📋</span>
  <span class="app-title">Document Control Register</span>
  <div class="spacer"></div>
  <button class="tb-btn" onclick="openSettings()">⚙ Settings</button>
  <button class="tb-btn" onclick="openProjectModal()">🏗 Project</button>
</div>

<!-- PROJECT BAR -->
<div id="projbar">
  {proj_fields_html}
  <button id="proj-edit-btn" onclick="openProjectModal()">✏ Edit</button>
</div>

<!-- TABS -->
<div id="tabs-bar">
  {tabs_html}
  <button class="tab-add" onclick="openAddDocType()" title="Add Document Type">＋</button>
</div>

<!-- TOOLBAR -->
<div id="toolbar">
  <button class="tool-btn" onclick="openAddRecord()">➕ Add Document</button>
  <button class="tool-btn green" onclick="doExport()">📥 Export Excel</button>
  <button class="tool-btn teal" onclick="openImport()">📤 Import CSV</button>
  <button class="tool-btn purple" onclick="manageColumns()">⚙ Columns</button>
  <input type="text" id="search-box" placeholder="Search all fields..." oninput="applySearch()">
</div>

<!-- MAIN -->
<div id="main">
  <div id="table-wrap">
    <table id="reg-table">
      <thead id="t-head"></thead>
      <tbody id="t-body"></tbody>
    </table>
    <div class="empty hidden" id="empty-state">
      <div class="ico">📁</div>
      <h3>No records found</h3>
      <p>Add your first document using the button above.</p>
    </div>
  </div>
</div>

<!-- STATUS BAR -->
<div id="statusbar">
  <span id="s-total">Total: 0</span>
  <span id="s-showing">Showing: 0</span>
  <span id="s-overdue" id="status-overdue">Overdue: 0</span>
  <span id="status-clock" style="margin-left:auto"></span>
</div>

<div id="toast"></div>

<!-- ADD/EDIT RECORD MODAL -->
<div class="overlay hidden" id="rec-modal">
  <div class="modal wide">
    <div class="modal-hdr">
      <span id="rec-modal-title">Add Document</span>
      <button class="modal-close" onclick="closeModal('rec-modal')">✕</button>
    </div>
    <div class="modal-body">
      <div class="seq-warn" id="seq-warn">⚠ Sequence gap detected — previous document number is missing.</div>
      <div class="form-grid" id="rec-form"></div>
      <div class="auto-info" id="auto-info" style="display:none"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('rec-modal')">Cancel</button>
      <button class="btn btn-primary" onclick="saveRecord()">Save Document</button>
    </div>
  </div>
</div>

<!-- PROJECT MODAL -->
<div class="overlay hidden" id="proj-modal">
  <div class="modal wide">
    <div class="modal-hdr">
      <span>🏗 Project Information</span>
      <button class="modal-close" onclick="closeModal('proj-modal')">✕</button>
    </div>
    <div class="modal-body" id="proj-modal-body"></div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('proj-modal')">Cancel</button>
      <button class="btn btn-primary" onclick="saveProject()">Save Project</button>
    </div>
  </div>
</div>

<!-- SETTINGS MODAL -->
<div class="overlay hidden" id="settings-modal">
  <div class="modal wide">
    <div class="modal-hdr">
      <span>⚙ Settings</span>
      <button class="modal-close" onclick="closeModal('settings-modal')">✕</button>
    </div>
    <div class="modal-body" id="settings-body"></div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('settings-modal')">Close</button>
    </div>
  </div>
</div>

<!-- ADD DOC TYPE MODAL -->
<div class="overlay hidden" id="addtype-modal">
  <div class="modal" style="max-width:420px">
    <div class="modal-hdr">
      <span>Add Document Type</span>
      <button class="modal-close" onclick="closeModal('addtype-modal')">✕</button>
    </div>
    <div class="modal-body">
      <div class="form-grid">
        <div class="form-group"><label>Type Name</label><input id="dt-name" placeholder="e.g. Method Statement"></div>
        <div class="form-group"><label>Code / Prefix</label><input id="dt-code" placeholder="e.g. MS" style="text-transform:uppercase"></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('addtype-modal')">Cancel</button>
      <button class="btn btn-primary" onclick="saveDocType()">Add Type</button>
    </div>
  </div>
</div>

<!-- COLUMNS MODAL -->
<div class="overlay hidden" id="cols-modal">
  <div class="modal">
    <div class="modal-hdr">
      <span>⚙ Manage Columns</span>
      <button class="modal-close" onclick="closeModal('cols-modal')">✕</button>
    </div>
    <div class="modal-body" id="cols-body"></div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="openAddColumn()">+ Add Column</button>
      <button class="btn btn-primary" onclick="closeModal('cols-modal');loadRecords()">Done</button>
    </div>
  </div>
</div>

<!-- ADD COLUMN MODAL -->
<div class="overlay hidden" id="addcol-modal">
  <div class="modal" style="max-width:420px">
    <div class="modal-hdr">
      <span>Add Column</span>
      <button class="modal-close" onclick="closeModal('addcol-modal')">✕</button>
    </div>
    <div class="modal-body">
      <div class="form-grid">
        <div class="form-group full"><label>Column Name</label><input id="col-name"></div>
        <div class="form-group full"><label>Type</label>
          <select id="col-type" onchange="document.getElementById('col-list-grp').style.display=this.value==='dropdown'?'flex':'none'">
            <option value="text">Text</option><option value="number">Number</option>
            <option value="date">Date</option><option value="dropdown">Dropdown</option>
            <option value="link">Hyperlink</option>
          </select>
        </div>
        <div class="form-group full" id="col-list-grp" style="display:none">
          <label>Dropdown List Source</label>
          <select id="col-list-src"></select>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('addcol-modal')">Cancel</button>
      <button class="btn btn-primary" onclick="saveAddColumn()">Add Column</button>
    </div>
  </div>
</div>

<!-- IMPORT MODAL -->
<div class="overlay hidden" id="import-modal">
  <div class="modal" style="max-width:500px">
    <div class="modal-hdr">
      <span>📤 Import from CSV</span>
      <button class="modal-close" onclick="closeModal('import-modal')">✕</button>
    </div>
    <div class="modal-body">
      <p style="font-size:12px;color:var(--muted);margin-bottom:12px">Select a CSV file exported from the old Excel register. Column headers must match.</p>
      <input type="file" id="import-file" accept=".csv" style="font-size:12px">
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('import-modal')">Cancel</button>
      <button class="btn btn-primary" onclick="doImport()">Import</button>
    </div>
  </div>
</div>

<script>
// ============================================================
// STATE
// ============================================================
const STATUS_COLORS = {json.dumps(STATUS_COLORS)};
let state = {{
  activeTab: null,
  docTypes: [],
  columns: [],
  allRecords: [],
  sortCol: null,
  sortDir: 'asc',
  colFilters: {{}},
  editingId: null,
  allLists: {{}},
}};

// ============================================================
// INIT
// ============================================================
async function init() {{
  await loadDocTypes();
  await loadLists();
  updateClock();
  setInterval(updateClock, 60000);
}}

function updateClock() {{
  document.getElementById('status-clock').textContent = new Date().toLocaleString('en-GB');
}}

async function api(url, opts={{}}) {{
  const r = await fetch(url, {{headers:{{'Content-Type':'application/json'}},...opts}});
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}}

async function loadDocTypes() {{
  const types = await api('/api/doc_types');
  state.docTypes = types;
  renderTabs(types);
  const counts = await api('/api/counts');
  types.forEach(dt => {{
    const el = document.getElementById('cnt-'+dt.id);
    if(el) el.textContent = counts[dt.id]||0;
  }});
  if(!state.activeTab && types.length) switchTab(types[0].id);
}}

async function loadLists() {{
  state.allLists = await api('/api/dropdown_lists');
}}

// ============================================================
// TABS
// ============================================================
function renderTabs(types) {{
  const bar = document.getElementById('tabs-bar');
  bar.querySelectorAll('.tab-btn').forEach(b=>b.remove());
  const addBtn = bar.querySelector('.tab-add');
  types.forEach(dt => {{
    const btn = document.createElement('button');
    btn.className = 'tab-btn' + (dt.id===state.activeTab?' active':'');
    btn.dataset.id = dt.id;
    btn.innerHTML = `<span class="tab-code">${{dt.code}}</span><span class="tab-count" id="cnt-${{dt.id}}">0</span>`;
    btn.title = dt.name;
    btn.onclick = ()=>switchTab(dt.id);
    btn.oncontextmenu = e=>{{ e.preventDefault(); tabContextMenu(dt.id,e); }};
    bar.insertBefore(btn, addBtn);
  }});
}}

function switchTab(id) {{
  state.activeTab = id;
  state.colFilters = {{}};
  state.sortCol = null;
  document.getElementById('search-box').value = '';
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active', b.dataset.id===id));
  loadRecords();
}}

function tabContextMenu(id, e) {{
  const old = document.getElementById('tab-ctx');
  if(old) old.remove();
  const menu = document.createElement('div');
  menu.id = 'tab-ctx';
  menu.style.cssText=`position:fixed;left:${{e.clientX}}px;top:${{e.clientY}}px;background:#fff;border:1px solid var(--border);border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.15);z-index:9999;min-width:140px;overflow:hidden`;
  menu.innerHTML=`
    <div onclick="renameTab('${{id}}')" style="padding:8px 14px;cursor:pointer;font-size:12px" onmouseover="this.style.background='#f0f4fa'" onmouseout="this.style.background=''">✏ Rename</div>
    <div onclick="deleteTab('${{id}}')" style="padding:8px 14px;cursor:pointer;font-size:12px;color:#ef4444" onmouseover="this.style.background='#fef2f2'" onmouseout="this.style.background=''">🗑 Delete Type</div>`;
  document.body.appendChild(menu);
  setTimeout(()=>document.addEventListener('click',()=>menu.remove(),{{once:true}}),10);
}}

async function renameTab(id) {{
  const name = prompt('New name:'); if(!name) return;
  // For now just reload - rename can be added as API endpoint
  toast('Rename: refresh after implementing rename API','warning');
}}

async function deleteTab(id) {{
  if(!confirm('Delete this document type and ALL its records?')) return;
  await api('/api/delete_doc_type/'+id, {{method:'POST',body:'{{}}'}});
  if(state.activeTab===id) state.activeTab=null;
  await loadDocTypes();
  toast('Document type deleted','warning');
}}

// ============================================================
// LOAD & RENDER RECORDS
// ============================================================
async function loadRecords() {{
  if(!state.activeTab) return;
  const search = document.getElementById('search-box').value.trim();
  const url = '/api/records/'+state.activeTab+(search?'?search='+encodeURIComponent(search):'');
  const data = await api(url);
  state.allRecords = data.records;
  state.columns = data.columns.filter(c=>c.visible);

  // Update tab count
  const cnt = document.getElementById('cnt-'+state.activeTab);
  if(cnt) cnt.textContent = data.count;

  buildTableHead();
  renderRows();
}}

function buildTableHead() {{
  const head = document.getElementById('t-head');
  head.innerHTML = '';

  // Main header row
  const hr = document.createElement('tr');
  const srTh = document.createElement('th');
  srTh.textContent = 'Sr.'; srTh.style.width='38px'; srTh.style.cursor='default';
  hr.appendChild(srTh);

  state.columns.forEach(col => {{
    const th = document.createElement('th');
    const sortable = !['auto_date','auto_num'].includes(col.col_type);
    th.innerHTML = col.label + (state.sortCol===col.col_key ? (state.sortDir==='asc'?' ↑':' ↓'):'');
    if(sortable) th.onclick = ()=>sortBy(col.col_key);
    else th.style.cursor='default';
    hr.appendChild(th);
  }});

  const actTh = document.createElement('th');
  actTh.textContent='Actions'; actTh.style.width='72px'; actTh.style.cursor='default';
  hr.appendChild(actTh);
  head.appendChild(hr);

  // Filter row
  const fr = document.createElement('tr');
  fr.className = 'filter-row';
  fr.appendChild(document.createElement('th')); // sr

  state.columns.forEach(col => {{
    const th = document.createElement('th');
    if(['auto_date','auto_num'].includes(col.col_type)) {{ fr.appendChild(th); return; }}
    if(col.col_type==='dropdown' && col.list_name) {{
      const sel = document.createElement('select');
      const opts = state.allLists[col.list_name]||[];
      sel.innerHTML='<option value="">All</option>'+opts.map(o=>`<option ${{state.colFilters[col.col_key]===o?'selected':''}}>${{o}}</option>`).join('');
      sel.onchange = ()=>{{ state.colFilters[col.col_key]=sel.value; renderRows(); }};
      th.appendChild(sel);
    }} else {{
      const inp = document.createElement('input');
      inp.value = state.colFilters[col.col_key]||'';
      inp.oninput = ()=>{{ state.colFilters[col.col_key]=inp.value; renderRows(); }};
      th.appendChild(inp);
    }}
    fr.appendChild(th);
  }});
  fr.appendChild(document.createElement('th'));
  head.appendChild(fr);
}}

function renderRows() {{
  const body = document.getElementById('t-body');
  body.innerHTML = '';
  const emptyState = document.getElementById('empty-state');

  // Apply column filters
  let rows = state.allRecords.filter(row => {{
    for(const [k,v] of Object.entries(state.colFilters)) {{
      if(v && !String(row[k]||'').toLowerCase().includes(v.toLowerCase())) return false;
    }}
    return true;
  }});

  // Sort
  if(state.sortCol) {{
    rows.sort((a,b) => {{
      const va=String(a[state.sortCol]||'').toLowerCase();
      const vb=String(b[state.sortCol]||'').toLowerCase();
      return state.sortDir==='asc'?(va>vb?1:-1):(va<vb?1:-1);
    }});
  }}

  emptyState.classList.toggle('hidden', rows.length > 0);

  let sr = 1;
  rows.forEach((row, idx) => {{
    const tr = document.createElement('tr');
    if(row._overdue) tr.classList.add('overdue');
    else if(row._isRev) tr.classList.add('rev-row');
    else if(idx%2===1) tr.classList.add('alt');

    // Sr
    const tdSr = document.createElement('td');
    tdSr.className = 'sr-cell';
    tdSr.textContent = row._isRev ? '' : sr;
    tr.appendChild(tdSr);

    // Data columns
    state.columns.forEach(col => {{
      const td = document.createElement('td');
      const key = col.col_key;
      let val = '';

      if(key==='expectedReplyDate') {{
        val = row._expectedReplyDate||'';
        if(row._overdue && val) td.classList.add('overdue-date');
      }} else if(key==='duration') {{
        val = row._duration||'';
      }} else if(key==='issuedDate') {{
        val = row._issuedDateFmt||'';
      }} else if(key==='actualReplyDate') {{
        val = row._actualReplyDateFmt||'';
      }} else if(key==='status') {{
        val = row[key]||'';
        if(val) {{
          td.innerHTML = val.split(',').map(s=>{{
            s=s.trim();
            const [bg,fg]=STATUS_COLORS[s]||['e5e7eb','374151'];
            return `<span class="status-badge" style="background:#${{bg}};color:#${{fg}}">${{s}}</span>`;
          }}).join('');
          tr.appendChild(td); return;
        }}
      }} else if(key==='fileLocation') {{
        const url = row[key]||'';
        if(url) {{ td.innerHTML=`<a class="file-link" href="${{url}}" target="_blank">View File</a>`; tr.appendChild(td); return; }}
      }} else {{
        val = String(row[key]||'');
      }}

      td.textContent = val;
      tr.appendChild(td);
    }});

    // Actions
    const tdAct = document.createElement('td');
    tdAct.className = 'actions-cell';
    tdAct.innerHTML = `
      <button class="act-btn" onclick="editRecord('${{row._id}}')">✏</button>
      <button class="act-btn del" onclick="deleteRecord('${{row._id}}')">🗑</button>`;
    tr.appendChild(tdAct);
    body.appendChild(tr);
    if(!row._isRev) sr++;
  }});

  // Update status bar
  const overdue = state.allRecords.filter(r=>r._overdue).length;
  document.getElementById('s-total').textContent = 'Total: '+state.allRecords.length;
  document.getElementById('s-showing').textContent = 'Showing: '+rows.length;
  document.getElementById('s-overdue').textContent = 'Overdue: '+overdue;
}}

function sortBy(key) {{
  state.sortDir = state.sortCol===key ? (state.sortDir==='asc'?'desc':'asc') : 'asc';
  state.sortCol = key;
  buildTableHead();
  renderRows();
}}

function applySearch() {{ loadRecords(); }}

// ============================================================
// ADD / EDIT RECORD
// ============================================================
function openAddRecord() {{
  state.editingId = null;
  document.getElementById('rec-modal-title').textContent = 'Add Document';
  document.getElementById('seq-warn').style.display='none';
  document.getElementById('auto-info').style.display='none';
  buildRecordForm(null);
  openModal('rec-modal');
}}

function editRecord(id) {{
  state.editingId = id;
  const row = state.allRecords.find(r=>r._id===id);
  if(!row) return;
  document.getElementById('rec-modal-title').textContent = 'Edit Document';
  buildRecordForm(row);
  openModal('rec-modal');
}}

async function buildRecordForm(row) {{
  const allCols = await api('/api/columns?dt='+state.activeTab);
  const AUTO = new Set(['expectedReplyDate','duration']);
  const editableCols = allCols.filter(c=>!AUTO.has(c.col_key));
  const dt = state.docTypes.find(d=>d.id===state.activeTab);
  const prefix = dt?.code||state.activeTab;

  const grid = document.getElementById('rec-form');
  grid.innerHTML = '';

  // Suggest next doc no
  let nextDocNo = '';
  if(!row) {{
    const r = await api('/api/next_doc_no/'+state.activeTab);
    nextDocNo = r.next||'';
  }}

  for(const col of editableCols) {{
    const key = col.col_key;
    const full = ['title','remarks','fileLocation','itemRef'].includes(key);
    const grp = document.createElement('div');
    grp.className = 'form-group'+(full?' full':'');

    const lbl = document.createElement('label');
    lbl.textContent = col.label;
    grp.appendChild(lbl);

    const existing = row?.[key]||'';

    if(col.col_type==='date') {{
      const inp = document.createElement('input');
      inp.type='date'; inp.id='f-'+key; inp.value=existing;
      if(key==='issuedDate') inp.onchange = ()=>updateAutoInfo();
      grp.appendChild(inp);
    }} else if(col.col_type==='dropdown' && col.list_name) {{
      const opts = state.allLists[col.list_name]||[];
      const ms = buildMultiSelect(key, opts, existing);
      grp.appendChild(ms);
    }} else if(col.col_type==='docno') {{
      const inp = document.createElement('input');
      inp.id='f-'+key; inp.value=row?existing:nextDocNo;
      inp.style.fontFamily='Consolas,monospace'; inp.style.fontWeight='600';
      inp.onblur = ()=>checkSeqGap(inp.value, prefix);
      grp.appendChild(inp);
    }} else if(key==='remarks') {{
      const ta = document.createElement('textarea');
      ta.id='f-'+key; ta.value=existing;
      grp.appendChild(ta);
    }} else {{
      const inp = document.createElement('input');
      inp.id='f-'+key; inp.value=existing;
      if(col.col_type==='link') inp.placeholder='https://drive.google.com/...';
      grp.appendChild(inp);
    }}

    grid.appendChild(grp);
  }}

  if(row) updateAutoInfo(row);
}}

function buildMultiSelect(key, options, initialValue) {{
  const selected = initialValue ? initialValue.split(',').map(s=>s.trim()).filter(Boolean) : [];
  const container = document.createElement('div');
  container.className='ms-container'; container.id='f-'+key;
  container.dataset.value = initialValue||'';

  function render() {{
    container.innerHTML='';
    selected.forEach(v=>{{
      const tag=document.createElement('span');
      tag.className='ms-tag';
      tag.innerHTML=`${{v}} <span class="rm" data-v="${{v}}">✕</span>`;
      tag.querySelector('.rm').onclick=e=>{{ e.stopPropagation(); selected.splice(selected.indexOf(v),1); container.dataset.value=selected.join(', '); render(); }};
      container.appendChild(tag);
    }});
    if(!selected.length) container.innerHTML='<span class="ms-placeholder">Select...</span>';
    container.dataset.value=selected.join(', ');
  }}

  container.onclick=e=>{{
    if(e.target.classList.contains('rm')) return;
    const existing=document.querySelector('.ms-dropdown');
    if(existing){{existing.remove();return;}}
    const dd=document.createElement('div');
    dd.className='ms-dropdown';
    options.forEach(opt=>{{
      const item=document.createElement('div');
      item.className='ms-option'+(selected.includes(opt)?' sel':'');
      item.innerHTML=`<input type="checkbox" ${{selected.includes(opt)?'checked':''}} style="pointer-events:none"> ${{opt}}`;
      item.onclick=ev=>{{ ev.stopPropagation();
        if(selected.includes(opt)) selected.splice(selected.indexOf(opt),1); else selected.push(opt);
        container.dataset.value=selected.join(', '); render();
        item.classList.toggle('sel',selected.includes(opt));
        item.querySelector('input').checked=selected.includes(opt);
      }};
      dd.appendChild(item);
    }});
    container.style.position='relative'; container.appendChild(dd);
  }};
  document.addEventListener('click',e=>{{ if(!container.contains(e.target)) container.querySelector('.ms-dropdown')?.remove(); }},true);
  render();
  return container;
}}

function checkSeqGap(docNo, prefix) {{
  const m = docNo.match(new RegExp(`^${{prefix}}-(\\\\d+)\\\\s+REV(\\\\d+)$`,'i'));
  if(!m||parseInt(m[2])!==0||parseInt(m[1])<=1) {{ document.getElementById('seq-warn').style.display='none'; return; }}
  const prevNo = `${{prefix}}-${{String(parseInt(m[1])-1).padStart(3,'0')}} REV00`;
  const exists = state.allRecords.some(r=>r.docNo===prevNo);
  document.getElementById('seq-warn').style.display = exists?'none':'block';
}}

function updateAutoInfo(row) {{
  const issuedEl = document.getElementById('f-issuedDate');
  const docNoEl = document.getElementById('f-docNo');
  const issued = issuedEl?.value || row?.issuedDate;
  const docNo = docNoEl?.value || row?.docNo || '';
  if(!issued) return;
  // Compute expected: 14 days REV00, 7 days REV01+
  const rev = parseInt((docNo.match(/REV(\\d+)/i)||[0,0])[1]);
  const infoEl = document.getElementById('auto-info');
  infoEl.style.display='block';
  infoEl.innerHTML='📅 Expected Reply: calculating... | Duration: will compute when actual reply is added';
}}

async function saveRecord() {{
  const allCols = await api('/api/columns?dt='+state.activeTab);
  const AUTO = new Set(['expectedReplyDate','duration']);
  const data = {{}};

  for(const col of allCols) {{
    if(AUTO.has(col.col_key)) continue;
    const el = document.getElementById('f-'+col.col_key);
    if(!el) continue;
    if(el.classList.contains('ms-container')) {{
      data[col.col_key] = el.dataset.value||'';
    }} else if(el.tagName==='TEXTAREA') {{
      data[col.col_key] = el.value.trim();
    }} else {{
      data[col.col_key] = el.value.trim();
    }}
  }}

  if(!data.docNo) {{ toast('Document No. is required','error'); return; }}
  if(state.editingId) data._id = state.editingId;

  await api('/api/records/'+state.activeTab, {{method:'POST',body:JSON.stringify(data)}});
  closeModal('rec-modal');
  await loadRecords();
  await loadDocTypes();
  toast(state.editingId?'Record updated':'Record added','success');
}}

async function deleteRecord(id) {{
  if(!confirm('Delete this record?')) return;
  await api('/api/delete_record/'+id, {{method:'POST',body:'{{}}'}});
  await loadRecords();
  await loadDocTypes();
  toast('Record deleted','warning');
}}

// ============================================================
// PROJECT MODAL
// ============================================================
const PROJ_FIELDS = [
  ['code','Project Code'],['name','Project Name'],
  ['startDate','Start Date (YYYY-MM-DD)'],['endDate','End Date (YYYY-MM-DD)'],
  ['client','Client'],['landlord','Landlord'],['pmo','PMO'],
  ['mainConsultant','Main Consultant'],['mepConsultant','MEP Consultant'],['contractor','Contractor']
];

async function openProjectModal() {{
  const proj = await api('/api/project');
  let extra = [];
  try {{ extra = JSON.parse(proj.extraFields||'[]'); }} catch(e){{}}

  const body = document.getElementById('proj-modal-body');
  body.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'form-grid';

  PROJ_FIELDS.forEach(([key,lbl])=>{{
    const grp=document.createElement('div'); grp.className='form-group';
    grp.innerHTML=`<label>${{lbl}}</label><input id="pf-${{key}}" value="${{proj[key]||''}}">`;
    grid.appendChild(grp);
  }});
  body.appendChild(grid);
  body.innerHTML += '<div class="section-title" style="margin-top:16px">Additional Fields</div>';

  const extraDiv = document.createElement('div');
  extraDiv.id='extra-fields';
  extra.forEach((ef,i)=>{{
    const row=document.createElement('div');
    row.style.cssText='display:flex;gap:6px;margin-bottom:6px;align-items:center';
    row.innerHTML=`<span style="font-size:11px;font-weight:600;min-width:100px">${{ef.label}}</span>
      <input id="ef-${{i}}" value="${{ef.value||''}}" style="flex:1;padding:6px;border:1px solid var(--border);border-radius:4px;font-family:inherit;font-size:12px">
      <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">✕</button>`;
    extraDiv.appendChild(row);
  }});
  body.appendChild(extraDiv);
  body.innerHTML += `<button class="btn btn-secondary btn-sm" style="margin-top:6px" onclick="addExtraField()">+ Add Field</button>`;
  openModal('proj-modal');
}}

function addExtraField() {{
  const lbl = prompt('Field Label:'); if(!lbl) return;
  const extraDiv = document.getElementById('extra-fields');
  const i = extraDiv.children.length;
  const row=document.createElement('div');
  row.style.cssText='display:flex;gap:6px;margin-bottom:6px;align-items:center';
  row.innerHTML=`<span style="font-size:11px;font-weight:600;min-width:100px">${{lbl}}</span>
    <input id="ef-${{i}}" style="flex:1;padding:6px;border:1px solid var(--border);border-radius:4px;font-family:inherit;font-size:12px">
    <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">✕</button>`;
  extraDiv.appendChild(row);
}}

async function saveProject() {{
  const data = {{}};
  PROJ_FIELDS.forEach(([key])=>{{ const el=document.getElementById('pf-'+key); if(el) data[key]=el.value.trim(); }});
  const extra = [];
  document.querySelectorAll('#extra-fields > div').forEach((row,i)=>{{
    const lbl=row.querySelector('span')?.textContent;
    const val=row.querySelector('input')?.value||'';
    if(lbl) extra.push({{label:lbl,value:val}});
  }});
  data.extraFields = JSON.stringify(extra);
  await api('/api/project',{{method:'POST',body:JSON.stringify(data)}});
  // Refresh project bar
  PROJ_FIELDS.forEach(([key])=>{{
    const el=document.getElementById('pf-'+key);
    if(el) {{const pfEl=document.getElementById('pf-'+key+'_bar')||document.querySelector(`#projbar .pf-val[data-key="${{key}}"]`); }}
  }});
  // Simple: reload projbar
  location.reload();
}}

// ============================================================
// SETTINGS
// ============================================================
async function openSettings() {{
  await loadLists();
  const body = document.getElementById('settings-body');
  body.innerHTML='';

  for(const [listName, items] of Object.entries(state.allLists)) {{
    body.innerHTML += `<div class="section-title">${{listName.charAt(0).toUpperCase()+listName.slice(1)}}</div>`;
    const ul=document.createElement('ul'); ul.className='settings-list';
    items.forEach(item=>{{
      const li=document.createElement('li'); li.className='settings-item';
      li.innerHTML=`<span class="name">${{item}}</span><button onclick="removeListItem('${{listName}}','${{item.replace(/'/g,"\\\\'")}}',this)">Remove</button>`;
      ul.appendChild(li);
    }});
    body.appendChild(ul);
    const addRow=document.createElement('div'); addRow.className='add-item-row';
    addRow.innerHTML=`<input id="new-${{listName}}" placeholder="New item...">
      <button class="btn btn-success btn-sm" onclick="addListItem('${{listName}}')">Add</button>`;
    body.appendChild(addRow);
  }}

  body.innerHTML+=`<div class="section-title">Add New Dropdown List</div>
    <div class="add-item-row">
      <input id="new-list-name" placeholder="List name (e.g. category)">
      <button class="btn btn-primary btn-sm" onclick="createNewList()">Create</button>
    </div>`;

  openModal('settings-modal');
}}

async function addListItem(listName) {{
  const inp=document.getElementById('new-'+listName);
  const val=inp?.value.trim(); if(!val) return;
  await api('/api/dropdown_lists/add',{{method:'POST',body:JSON.stringify({{list_name:listName,item:val}})}});
  await loadLists();
  openSettings();
}}

async function removeListItem(listName, item, btn) {{
  await api('/api/dropdown_lists/remove',{{method:'POST',body:JSON.stringify({{list_name:listName,item:item}})}});
  btn.closest('li').remove();
  await loadLists();
}}

async function createNewList() {{
  const name=document.getElementById('new-list-name')?.value.trim().toLowerCase().replace(/\\s+/g,'_');
  if(!name) return;
  await api('/api/dropdown_lists/add',{{method:'POST',body:JSON.stringify({{list_name:name,item:'Item 1'}})}});
  await loadLists();
  openSettings();
}}

// ============================================================
// DOC TYPE
// ============================================================
function openAddDocType() {{
  document.getElementById('dt-name').value='';
  document.getElementById('dt-code').value='';
  openModal('addtype-modal');
}}

async function saveDocType() {{
  const name=document.getElementById('dt-name').value.trim();
  const code=document.getElementById('dt-code').value.trim().toUpperCase();
  if(!name||!code) {{ toast('Name and code required','error'); return; }}
  await api('/api/doc_types',{{method:'POST',body:JSON.stringify({{name,code}})}});
  closeModal('addtype-modal');
  await loadDocTypes();
  switchTab(code);
  toast('Document type added','success');
}}

// ============================================================
// COLUMNS
// ============================================================
async function manageColumns() {{
  const cols = await api('/api/columns?dt='+state.activeTab);
  const body = document.getElementById('cols-body');
  body.innerHTML='';
  const PROTECTED=new Set(['docNo','issuedDate','status','expectedReplyDate']);
  const ul=document.createElement('ul'); ul.className='settings-list'; ul.style.maxHeight='350px';
  cols.forEach(col=>{{
    const li=document.createElement('li'); li.className='settings-item';
    li.innerHTML=`
      <input type="checkbox" ${{col.visible?'checked':''}} onchange="toggleColVisibility(${{col.id}},this.checked)">
      <span class="name">${{col.label}}</span>
      <span style="font-size:9px;background:#e0e7ff;color:#3730a3;padding:1px 6px;border-radius:3px">${{col.col_type}}</span>
      ${{!PROTECTED.has(col.col_key)?`<button onclick="deleteColumn(${{col.id}},this)">Del</button>`:'<span style="font-size:9px;color:var(--muted)">core</span>'}}`;
    ul.appendChild(li);
  }});
  body.appendChild(ul);
  openModal('cols-modal');
}}

async function toggleColVisibility(colId, visible) {{
  await api('/api/columns/visibility/'+colId, {{method:'POST',body:JSON.stringify({{visible}})}});
}}

async function deleteColumn(colId, btn) {{
  if(!confirm('Delete this column?')) return;
  await api('/api/columns/delete/'+colId, {{method:'POST',body:'{{}}'}});
  btn.closest('li').remove();
}}

async function openAddColumn() {{
  await loadLists();
  const sel=document.getElementById('col-list-src');
  sel.innerHTML=Object.keys(state.allLists).map(k=>`<option value="${{k}}">${{k}}</option>`).join('');
  openModal('addcol-modal');
}}

async function saveAddColumn() {{
  const name=document.getElementById('col-name').value.trim();
  const type=document.getElementById('col-type').value;
  const listSrc=type==='dropdown'?document.getElementById('col-list-src').value:null;
  if(!name) {{ toast('Name required','error'); return; }}
  const key='custom_'+name.toLowerCase().replace(/\\s+/g,'_')+'_'+Date.now();
  await api('/api/columns/add',{{method:'POST',body:JSON.stringify({{dt_id:state.activeTab,col_key:key,label:name,col_type:type,list_name:listSrc}})}});
  closeModal('addcol-modal');
  closeModal('cols-modal');
  await loadRecords();
  toast('Column added','success');
}}

// ============================================================
// EXPORT / IMPORT
// ============================================================
function doExport() {{
  if(!state.activeTab) return;
  window.location='/api/export/'+state.activeTab;
}}

function openImport() {{ openModal('import-modal'); }}

async function doImport() {{
  const file=document.getElementById('import-file').files[0];
  if(!file) return;
  const text=await file.text();
  const r=await api('/api/import_csv',{{method:'POST',body:JSON.stringify({{dt_id:state.activeTab,csv_text:text}})}});
  closeModal('import-modal');
  await loadRecords();
  await loadDocTypes();
  toast('Imported '+r.imported+' records','success');
}}

// ============================================================
// MODAL HELPERS
// ============================================================
function openModal(id) {{ document.getElementById(id).classList.remove('hidden'); }}
function closeModal(id) {{ document.getElementById(id).classList.add('hidden'); }}
document.querySelectorAll('.overlay').forEach(o=>o.addEventListener('click',function(e){{if(e.target===this)this.classList.add('hidden')}}));

// ============================================================
// TOAST
// ============================================================
function toast(msg, type='info') {{
  const t=document.getElementById('toast');
  t.textContent=msg; t.className='show '+(type||'');
  clearTimeout(t._t); t._t=setTimeout(()=>t.className='',3200);
}}

// Start
init();
</script>
</body>
</html>"""


# ============================================================
# START SERVER
# ============================================================
if __name__ == '__main__':
    import threading, webbrowser

    db.init_db()
    local_ip = get_local_ip()

    if IS_RENDER:
        # ── Running on Render ──────────────────────────────────────────
        print(f"[DCR] Starting on Render — PORT={PORT}")
        print(f"[DCR] Database: {'PostgreSQL/Supabase' if db.USE_POSTGRES else 'SQLite (no DATABASE_URL)'}")
    else:
        # ── Running locally ────────────────────────────────────────────
        print("=" * 52)
        print("  📋 Document Control Register - Web App")
        print("=" * 52)
        print(f"  ✅ Local:   http://localhost:{PORT}")
        print(f"  🌐 Network: http://{local_ip}:{PORT}")
        print(f"  📂 DB: {getattr(db, 'DB_PATH', 'PostgreSQL')}")
        print("=" * 52)
        def open_browser():
            import time; time.sleep(1)
            webbrowser.open(f'http://localhost:{PORT}')
        threading.Thread(target=open_browser, daemon=True).start()

    server = HTTPServer(('0.0.0.0', PORT), DCRHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n\nServer stopped.")
