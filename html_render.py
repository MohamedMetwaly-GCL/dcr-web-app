"""html_render.py - DCR HTML page rendering.

Contains all server-side HTML generation for the three application pages
and their shared building blocks:

  _user_info_html(u)      -- topbar buttons/labels for the current user
  BASE_CSS                -- shared <style> block injected into every page
  SHARED_JS               -- shared toast/modal/apiFetch JS injected into every page
  render_login()          -- the login page (no auth required)
  render_dashboard(u)     -- the main dashboard / analytics page
  render_register(u,proj) -- the per-project document register page

Dependency map (no circular imports):
  html_render -> auth  (for can_edit)
  html_render -> db    (for get_doc_types, get_logo)
  html_render -> utils (for STATUS_COLORS)
  auth        -> db
  db          -> (psycopg2 only)

Step 3 of the incremental refactor. Only HTML rendering lives here.
No routes, no business logic, no database writes.
"""
import json

import db
from auth import can_edit
from utils import STATUS_COLORS

def _user_info_html(u):
    if not u:
        return ('<a href="/login"><button class="tb-btn glow">🔐 Login</button></a>', "guest", "GUEST", "#fff3")
    role = u["role"]
    colors = {"superadmin":"rgba(240,165,0,.35)","admin":"rgba(255,255,255,.2)",
              "editor":"rgba(99,102,241,.3)","viewer":"rgba(255,255,255,.15)","guest":"rgba(255,255,255,.1)"}
    rbg = colors.get(role, "#fff3")
    labels = {"superadmin":"SUPER ADMIN","admin":"ADMIN","editor":"EDITOR","viewer":"VIEWER"}
    rlbl = labels.get(role, role.upper())
    btns = ""
    if role == "superadmin":
        btns += '<button class="tb-btn" onclick="openAdmin()">⚙ Admin</button>'
    btns += '<button class="tb-btn" onclick="changePw()">🔑</button>'
    btns += '<form action="/logout" method="post" style="display:inline"><button type="submit" class="tb-btn">⏻</button></form>'
    name = u["username"]
    return btns, name, rlbl, rbg

BASE_CSS = """
<style>
:root{--pr:#1a3a5c;--pl:#2563a8;--ac:#f0a500;--bg:#f0f4f8;--wh:#fff;--bd:#dde3ed;
  --tx:#1e2a3a;--mu:#6b7a94;--ok:#16a34a;--er:#ef4444;--wa:#f59e0b;--rd:6px}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--tx);font-size:13px}
#topbar{background:var(--pr);color:#fff;height:46px;display:flex;align-items:center;
  padding:0 14px;gap:8px;box-shadow:0 2px 8px rgba(0,0,0,.25);flex-shrink:0;position:relative;z-index:100}
#topbar .sp{flex:1}
.tb-btn{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;
  padding:5px 11px;border-radius:var(--rd);cursor:pointer;font-size:12px;font-family:inherit;
  text-decoration:none;display:inline-block;transition:background .15s}
.tb-btn:hover{background:rgba(255,255,255,.28)}
.tb-btn.glow{background:rgba(240,165,0,.3);border-color:rgba(240,165,0,.7);font-weight:700}
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:1000;
  display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px)}
.overlay.hidden{display:none!important}
.modal{background:#fff;border-radius:10px;box-shadow:0 24px 64px rgba(0,0,0,.3);
  width:92%;max-width:600px;max-height:90vh;display:flex;flex-direction:column;
  animation:mIn .18s ease}
@keyframes mIn{from{transform:translateY(-14px);opacity:0}}
.mhdr{background:var(--pr);color:#fff;padding:12px 18px;font-weight:700;font-size:13px;
  display:flex;justify-content:space-between;align-items:center;flex-shrink:0;border-radius:10px 10px 0 0}
.mbody{padding:16px 18px;overflow-y:auto;flex:1}
.mfoot{padding:10px 18px;border-top:1px solid var(--bd);display:flex;justify-content:flex-end;
  gap:8px;background:var(--bg);flex-shrink:0;border-radius:0 0 10px 10px}
.xbtn{background:none;border:none;color:#fff;font-size:20px;cursor:pointer;opacity:.7;line-height:1}
.xbtn:hover{opacity:1}
.fgrid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.fg{display:flex;flex-direction:column;gap:4px}
.fg.full{grid-column:1/-1}
.fg label{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px}
.fg input,.fg select,.fg textarea{padding:7px 10px;border:1.5px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none;transition:border-color .2s}
.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:var(--pl);box-shadow:0 0 0 2px rgba(37,99,168,.1)}
.btn{padding:7px 16px;border-radius:var(--rd);cursor:pointer;font-family:inherit;
  font-size:12px;font-weight:600;border:1px solid transparent;transition:all .15s}
.btn-pr{background:var(--pr);color:#fff}.btn-pr:hover{background:var(--pl)}
.btn-sc{background:var(--bg);color:var(--tx);border-color:var(--bd)}.btn-sc:hover{background:var(--bd)}
.btn-ok{background:var(--ok);color:#fff}
.btn-er{background:var(--er);color:#fff}
.btn-sm{padding:4px 10px;font-size:11px}
.stitle{font-size:11px;font-weight:700;color:var(--pr);text-transform:uppercase;
  letter-spacing:.5px;margin:14px 0 6px;padding-bottom:4px;border-bottom:2px solid var(--pr)}
.badge{display:inline-block;border-radius:10px;padding:2px 9px;font-size:10px;font-weight:700}
#toast{position:fixed;bottom:28px;right:18px;background:var(--pr);color:#fff;
  padding:10px 18px;border-radius:var(--rd);font-size:12px;z-index:9999;
  box-shadow:0 4px 16px rgba(0,0,0,.2);transform:translateY(80px);opacity:0;
  transition:all .3s;pointer-events:none;max-width:320px}
#toast.show{transform:none;opacity:1}
#toast.ok{background:#16a34a}#toast.er{background:#ef4444}#toast.wa{background:#f59e0b;color:#000}
</style>"""

SHARED_JS = """
<div id="toast"></div>
<script>
function toast(msg,type=''){
  const t=document.getElementById('toast');
  t.textContent=msg;t.className='show '+(type||'');
  clearTimeout(t._t);t._t=setTimeout(()=>t.className='',3200);
}
function openM(id){document.getElementById(id).classList.remove('hidden')}
function closeM(id){document.getElementById(id).classList.add('hidden')}
async function apiFetch(url,opts={}){
  const r=await fetch(url,{credentials:'include',headers:{'Content-Type':'application/json'},...opts});
  if(r.status===403){const d=await r.json().catch(()=>({}));
    if(d.error==='LOGIN_REQUIRED'){window.location='/login';return null;}
    throw new Error(d.error||'Forbidden');}
  if(!r.ok)throw new Error(await r.text());
  return r.json();
}
async function changePw(){
  const pw=prompt('New password (min 4 chars):');
  if(!pw||pw.length<4)return;
  const r=await apiFetch('/api/change_password',{method:'POST',body:JSON.stringify({password:pw})});
  if(r&&r.ok)toast('✔ Password changed','ok');else toast((r&&r.error)||'Error','er');
}
</script>"""


def render_login():
    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR — Login</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;min-height:100vh;display:flex;align-items:center;
  justify-content:center;background:linear-gradient(135deg,#0f2640,#1a3a5c 60%,#2563a8)}}
.card{{background:#fff;border-radius:16px;box-shadow:0 24px 80px rgba(0,0,0,.4);width:100%;max-width:400px;overflow:hidden}}
.chdr{{background:linear-gradient(135deg,#1a3a5c,#2563a8);padding:32px;text-align:center}}
.chdr h1{{color:#fff;font-size:20px;margin-top:8px}}
.chdr p{{color:rgba(255,255,255,.6);font-size:12px;margin-top:4px}}
.cbody{{padding:28px 32px 32px}}
.fld{{margin-bottom:16px}}
.fld label{{display:block;font-size:11px;font-weight:700;color:#6b7a94;text-transform:uppercase;letter-spacing:.4px;margin-bottom:5px}}
.fld input{{width:100%;padding:11px 14px;border:1.5px solid #dde3ed;border-radius:8px;font-family:inherit;font-size:13px;outline:none;transition:border-color .2s}}
.fld input:focus{{border-color:#2563a8;box-shadow:0 0 0 3px rgba(37,99,168,.12)}}
.err{{background:#fef2f2;border:1px solid #fecaca;color:#dc2626;padding:9px 12px;border-radius:6px;font-size:12px;margin-bottom:14px;display:none}}
.btn-login{{width:100%;padding:13px;background:linear-gradient(135deg,#1a3a5c,#2563a8);color:#fff;border:none;
  border-radius:8px;font-family:inherit;font-size:14px;font-weight:700;cursor:pointer;transition:all .2s}}
.btn-login:hover{{transform:translateY(-1px);box-shadow:0 4px 16px rgba(26,58,92,.4)}}
.hint{{text-align:center;color:#9ca3af;font-size:11px;margin-top:16px}}
</style></head><body>
<div class="card">
  <div class="chdr"><div style="font-size:40px">📋</div>
    <h1>Document Control Register</h1><p>Sign in to continue</p></div>
  <div class="cbody">
    <div class="err" id="err"></div>
    <div class="fld"><label>Username</label><input id="un" type="text" autofocus autocomplete="username"></div>
    <div class="fld"><label>Password</label><input id="pw" type="password" autocomplete="current-password"></div>
    <button class="btn-login" onclick="login()">Sign In →</button>
    <p class="hint">Contact your administrator for credentials</p>
  </div>
</div>
<script>
document.getElementById('pw').onkeydown=e=>{{if(e.key==='Enter')login()}};
document.getElementById('un').onkeydown=e=>{{if(e.key==='Enter')document.getElementById('pw').focus()}};
async function login(){{
  const un=document.getElementById('un').value.trim();
  const pw=document.getElementById('pw').value;
  const err=document.getElementById('err');
  err.style.display='none';
  if(!un||!pw){{err.textContent='Please enter username and password';err.style.display='block';return;}}
  const r=await fetch('/login',{{method:'POST',credentials:'include',
    headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{username:un,password:pw}})}});
  const d=await r.json();
  if(d.ok) window.location='/';
  else{{err.textContent=d.error||'Invalid credentials';err.style.display='block';}}
}}
</script></body></html>"""




def render_dashboard(u):
    btns, uname, rlbl, rbg = _user_info_html(u)
    role = u["role"] if u else "guest"
    audit_btn = '<button class="tbtn" onclick="showTab(&#39;audit&#39;)">📜 Audit Log</button>' if role in ('superadmin','admin') else ''

    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR — Dashboard</title>
{BASE_CSS}
<style>
/* ── Layout ── */
body{{display:flex;flex-direction:column;min-height:100vh}}
.wrap{{max-width:1480px;margin:0 auto;padding:14px 12px;flex:1;width:100%}}

/* ── Dark Mode ── */
body.dark{{--bg:#0f172a;--wh:#1e293b;--tx:#e2e8f0;--mu:#94a3b8;--bd:#334155;--pr:#3b82f6;--pl:#60a5fa}}
body.dark .kpi,body.dark .ccard,body.dark .pcard,body.dark .panel,
body.dark .psel-bar,body.dark .tbl-wrap{{background:#1e293b;color:#e2e8f0}}
body.dark .kval{{color:#93c5fd}}
body.dark .kpi.ok .kval{{color:#86efac}}
body.dark .kpi.wa .kval{{color:#fcd34d}}
body.dark .kpi.er .kval{{color:#fca5a5}}
body.dark .kpi.pu .kval{{color:#c4b5fd}}
body.dark .klbl{{color:#94a3b8}}
body.dark .clbl{{color:#93c5fd}}
body.dark .stitle{{color:#93c5fd;border-color:#93c5fd}}
body.dark .dt-tbl th{{background:#1e3a5f;color:#e2e8f0}}
body.dark .dt-tbl td{{border-color:#334155;color:#cbd5e1}}
body.dark .dt-tbl tr:hover td{{background:#0f172a}}
body.dark .dt-tbl .alt td{{background:#162032}}
body.dark #topbar{{background:#0f2640}}
body.dark .pchdr{{background:#1e3a5f}}
body.dark .pcbody{{color:#cbd5e1}}
body.dark select,body.dark input{{background:#1e293b;color:#e2e8f0;border-color:#334155}}
body.dark .tbtn{{background:#1e293b;color:#e2e8f0;border-color:#334155}}
body.dark .ov-row{{border-color:#334155;color:#cbd5e1}}
body.dark .prog{{background:#334155}}

/* ── KPIs ── */
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin-bottom:12px}}
.kpi{{background:var(--wh);border-radius:10px;padding:12px 14px;
  box-shadow:0 1px 4px rgba(0,0,0,.07);border-left:4px solid var(--pr);cursor:default;
  transition:transform .15s,box-shadow .15s}}
.kpi:hover{{transform:translateY(-2px);box-shadow:0 4px 14px rgba(0,0,0,.12)}}
.kpi.ok{{border-left-color:var(--ok)}}.kpi.wa{{border-left-color:var(--wa)}}
.kpi.er{{border-left-color:var(--er)}}.kpi.pu{{border-left-color:#7c3aed}}
.kval{{font-size:28px;font-weight:800;color:var(--pr);line-height:1}}
.kpi.ok .kval{{color:var(--ok)}}.kpi.wa .kval{{color:var(--wa)}}
.kpi.er .kval{{color:var(--er)}}.kpi.pu .kval{{color:#7c3aed}}
.klbl{{font-size:9px;color:var(--mu);font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:3px}}
.ktrend{{font-size:10px;color:var(--mu);margin-top:2px}}

/* ── Toolbar ── */
.psel-bar{{display:flex;align-items:center;gap:8px;background:var(--wh);padding:8px 12px;
  border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:12px;flex-wrap:wrap}}
.psel-bar label{{font-size:11px;font-weight:700;color:var(--mu);white-space:nowrap}}
.psel-bar select,.psel-bar input{{padding:5px 10px;border:1.5px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none}}
.tbtn{{padding:5px 11px;border:1.5px solid var(--bd);border-radius:var(--rd);
  background:var(--wh);cursor:pointer;font-size:11px;font-weight:600;font-family:inherit;
  transition:all .15s;color:var(--tx)}}
.tbtn:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.tbtn.active{{background:var(--pr);color:#fff;border-color:var(--pr)}}

/* ── Charts grid ── */
.charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px}}
.charts-grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:12px}}
.ccard{{background:var(--wh);border-radius:10px;padding:14px;
  box-shadow:0 1px 4px rgba(0,0,0,.07)}}
.clbl{{font-size:10px;font-weight:700;color:var(--pr);text-transform:uppercase;
  letter-spacing:.5px;margin-bottom:10px;display:flex;align-items:center;gap:6px}}
canvas{{max-height:200px}}

/* ── Project cards ── */
.pgrid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:10px;margin-bottom:14px}}
.pcard{{background:var(--wh);border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,.08);
  overflow:hidden;text-decoration:none;color:inherit;display:block;
  transition:transform .15s,box-shadow .15s;position:relative}}
.pcard:hover{{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,.12)}}
.pchdr{{background:var(--pr);padding:10px 12px;display:flex;align-items:center;justify-content:space-between}}
.pcbody{{padding:10px 12px}}
.prow{{display:flex;justify-content:space-between;align-items:center;margin-bottom:4px}}
.prog{{height:5px;background:#eef1f7;border-radius:99px;overflow:hidden;margin-top:5px}}
.progf{{height:100%;border-radius:99px;transition:width .6s ease}}
.addcard{{background:var(--wh);border-radius:10px;border:2px dashed var(--bd);min-height:130px;
  display:flex;align-items:center;justify-content:center;flex-direction:column;gap:6px;
  cursor:pointer;transition:all .15s;color:var(--mu);font-size:12px}}
.addcard:hover{{border-color:var(--pr);color:var(--pr)}}

/* ── Tables ── */
.tbl-wrap{{background:var(--wh);border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.07);
  overflow:hidden;margin-bottom:14px}}
.dt-tbl{{width:100%;border-collapse:collapse;font-size:12px}}
.dt-tbl th{{background:var(--pr);color:#fff;padding:8px 10px;
  text-align:left;font-weight:600;white-space:nowrap;font-size:11px}}
.dt-tbl td{{padding:6px 10px;border-bottom:1px solid #edf0f5}}
.dt-tbl tr:hover td{{background:#f0f4f8}}
.dt-tbl .alt td{{background:#fafbfd}}
.pr-toggle{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.pr-toggle:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.pr-items-row{{display:none}}
.pr-items-row.open{{display:table-row}}
.pr-items-row td{{padding:0;border-bottom:1px solid #edf0f5}}
.pr-items-wrap{{padding:10px 12px;background:#f8fafc}}
.pr-items-table{{width:100%;border-collapse:collapse;font-size:11px}}
.pr-items-table th{{background:#e2e8f0;color:#334155;padding:6px 8px;text-align:left;font-weight:700}}
.pr-items-table td{{padding:6px 8px;border-bottom:1px solid #e5e7eb}}
.pr-items-table .pr-section td{{background:#e8eef6;color:var(--pr);font-weight:800;text-transform:uppercase;letter-spacing:.35px}}
.pr-items-empty{{color:var(--mu);font-size:11px;padding:6px 0}}
.pr-items-editor table{{width:100%;border-collapse:collapse;font-size:11px}}
.pr-items-editor th{{background:#e2e8f0;color:#334155;padding:6px 8px;text-align:left;font-weight:700}}
.pr-items-editor td{{padding:6px 8px;border-bottom:1px solid #e5e7eb}}
.pr-items-editor input{{width:100%;padding:6px 8px;border:1.5px solid var(--bd);border-radius:var(--rd);font-family:inherit;font-size:12px}}
.pr-items-editor tr.pr-head-edit td{{background:#f8fafc}}
.pr-items-editor .pr-head-label{{font-size:10px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.35px;margin-bottom:4px}}

/* ── Overdue panel ── */
.ov-row{{display:flex;align-items:center;gap:8px;padding:7px 10px;
  border-bottom:1px solid var(--bd);font-size:11px}}
.ov-row:last-child{{border-bottom:none}}
.ov-badge{{background:#fef2f2;color:#ef4444;border-radius:6px;
  padding:2px 8px;font-weight:700;font-size:10px;white-space:nowrap}}
.ov-badge.warn{{background:#fffbeb;color:#f59e0b}}

/* ── Executive Summary panel ── */
.panel{{background:var(--wh);border-radius:10px;padding:16px;
  box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:14px}}
.panel-title{{font-size:12px;font-weight:700;color:var(--pr);
  text-transform:uppercase;letter-spacing:.5px;margin-bottom:12px;
  border-bottom:2px solid var(--pr);padding-bottom:6px}}

/* ── Tabs ── */
.tab-bar{{display:flex;gap:4px;margin-bottom:12px;border-bottom:2px solid var(--bd);padding-bottom:0}}
.tab{{padding:8px 16px;cursor:pointer;font-size:12px;font-weight:600;color:var(--mu);
  border-bottom:2px solid transparent;margin-bottom:-2px;transition:all .15s}}
.tab.active{{color:var(--pr);border-bottom-color:var(--pr)}}
.tab:hover:not(.active){{color:var(--tx)}}
.tab-pane{{display:none}}.tab-pane.active{{display:block}}

/* ── Loading spinner ── */
#ld{{position:fixed;inset:0;background:rgba(15,38,64,.92);z-index:500;
  display:flex;align-items:center;justify-content:center;flex-direction:column;gap:14px}}
.spin{{width:40px;height:40px;border:4px solid rgba(255,255,255,.2);border-top-color:#f0a500;
  border-radius:50%;animation:spin .6s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}

/* ── Mobile ── */
@media(max-width:768px){{
  .charts-grid,.charts-grid-3{{grid-template-columns:1fr}}
  .pgrid{{grid-template-columns:1fr}}
  .kpi-grid{{grid-template-columns:repeat(2,1fr)}}
  .wrap{{padding:8px}}
  .tab{{padding:6px 10px;font-size:11px}}
  .tbl-wrap{{overflow-x:auto;-webkit-overflow-scrolling:touch}}
  .dt-tbl{{min-width:600px}}
  .psel-bar{{flex-wrap:wrap;gap:6px}}
  .psel-bar select{{min-width:140px;flex:1}}
  .kpi{{padding:8px 10px}}
  .kval{{font-size:22px}}
}}
@media(max-width:480px){{
  .kpi-grid{{grid-template-columns:repeat(2,1fr)}}
  .psel-bar{{flex-direction:column;align-items:stretch}}
  .kpi-grid{{gap:5px}}
  .charts-grid{{gap:8px}}
}}

/* ── Del button ── */
.del-pbtn{{background:none;border:none;cursor:pointer;color:#ef4444;
  font-size:11px;padding:2px 5px;border-radius:3px}}
.del-pbtn:hover{{background:#fef2f2}}
</style>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
</head><body>

<div id="ld">
  <div class="spin"></div>
  <div style="color:rgba(255,255,255,.7);font-size:13px;font-weight:600">Loading dashboard...</div>
</div>

<div id="topbar">
  <span style="font-size:20px">📋</span>
  <span style="font-weight:700;font-size:14px">Document Control Register</span>
  <div class="sp"></div>
  {btns}
  <button class="tb-btn" onclick="toggleDark()" id="darkBtn" title="Toggle dark mode">🌙</button>
  <span style="color:rgba(255,255,255,.45);padding:0 4px">|</span>
  <span style="color:rgba(255,255,255,.8);font-size:11px">👤 {uname}
    <span style="background:{rbg};border-radius:3px;padding:1px 7px;font-size:9px;font-weight:700">{rlbl}</span>
  </span>
</div>

<div class="wrap">

  <!-- KPIs -->
  <div class="kpi-grid">
    <div class="kpi"><div class="kval" id="kpi-total">—</div><div class="klbl">Total Docs</div></div>
    <div class="kpi ok"><div class="kval" id="kpi-approved">—</div><div class="klbl">Approved</div></div>
    <div class="kpi wa"><div class="kval" id="kpi-pending">—</div><div class="klbl">Under Review</div></div>
    <div class="kpi er"><div class="kval" id="kpi-overdue">—</div><div class="klbl">Overdue</div></div>
    <div class="kpi pu"><div class="kval" id="kpi-rejected">—</div><div class="klbl">Rejected/Revise</div></div>
    <div class="kpi"><div class="kval" id="kpi-pct">—</div><div class="klbl">Completion %</div></div>
    <div class="kpi"><div class="kval" id="kpi-projs">—</div><div class="klbl">Projects</div></div>
  </div>

  <!-- Toolbar -->
  <div class="psel-bar">
    <label>🔍 Project:</label>
    <select id="proj-sel" onchange="filterProject(this.value)">
      <option value="">All Projects</option>
    </select>
    <label>🏗 Discipline:</label>
    <select id="disc-sel" onchange="filterDisc(this.value)">
      <option value="">All Disciplines</option>
    </select>
    <div class="sp" style="flex:1"></div>
    <button class="tbtn" onclick="showTab('overview')">📊 Overview</button>
    <button class="tbtn" onclick="showTab('analytics')">📈 Analytics</button>
    <button class="tbtn" onclick="showTab('overdue')">⚠️ Overdue</button>
    <button class="tbtn" onclick="showTab('executive')">📋 Executive</button>
    {audit_btn}
  </div>

    <!-- TAB: OVERVIEW -->
    <div id="tab-overview" class="tab-pane active">
      <div class="stitle">🗂 Projects</div>
      <div class="pgrid" id="pgrid"></div>
      <div class="panel" style="padding:12px 14px;margin-bottom:12px">
        <div class="panel-title" style="margin-bottom:10px">PR Analytics</div>
        <div id="pr-panel" style="display:grid;grid-template-columns:1fr;gap:12px">
          <div style="font-size:11px;color:var(--mu)">Loading...</div>
        </div>
      </div>

    <div class="charts-grid">
      <div class="ccard"><div class="clbl">📊 Documents by Project</div><canvas id="cProj"></canvas></div>
      <div class="ccard"><div class="clbl">🥧 Status Distribution</div><canvas id="cStatus"></canvas></div>
    </div>

    <div class="stitle">📋 Document Types Summary</div>
    <div class="tbl-wrap">
      <table class="dt-tbl">
        <thead><tr>
          <th>Project</th><th>Code</th><th>Type</th>
          <th style="text-align:center">Total</th>
          <th style="text-align:center">Approved</th>
          <th style="text-align:center">Pending</th>
          <th style="text-align:center">Rejected</th>
          <th style="text-align:center">Overdue</th>
        </tr></thead>
        <tbody id="dt-tbody"></tbody>
      </table>
      <div id="dt-empty" style="text-align:center;padding:24px;color:var(--mu);display:none">No data</div>
    </div>

    <div class="stitle">🏗 Discipline Breakdown</div>
    <div class="tbl-wrap">
      <table class="dt-tbl">
        <thead><tr>
          <th>Project</th><th>Doc Type</th><th>Discipline</th>
          <th style="text-align:center">Total</th>
          <th style="text-align:center">Approved</th>
          <th style="text-align:center">Pending</th>
          <th style="text-align:center">Rejected</th>
          <th style="text-align:center">Overdue</th>
        </tr></thead>
        <tbody id="disc-tbody"></tbody>
      </table>
      <div id="disc-empty" style="text-align:center;padding:24px;color:var(--mu);display:none">No data</div>
    </div>
  </div>

  <!-- TAB: ANALYTICS -->
  <div id="tab-analytics" class="tab-pane">
    <div class="charts-grid">
      <div class="ccard">
        <div class="clbl">📅 Monthly Trend (last 6 months)</div>
        <div style="position:relative;height:200px"><canvas id="cTrend"></canvas></div>
      </div>
      <div class="ccard">
        <div class="clbl">⏳ Aging — Pending Docs (Days Lapsed)</div>
        <div style="position:relative;height:200px"><canvas id="cAging"></canvas></div>
      </div>
    </div>
    <div class="charts-grid">
      <div class="ccard">
        <div class="clbl">🔄 Document Quality (Revisions Needed)</div>
        <div style="position:relative;height:200px"><canvas id="cQuality"></canvas></div>
      </div>
      <div class="ccard">
        <div class="clbl">✅ Approval Rate by Doc Type</div>
        <div style="position:relative;height:200px"><canvas id="cApprRate"></canvas></div>
      </div>
    </div>
  </div>

  <!-- TAB: OVERDUE -->
  <div id="tab-overdue" class="tab-pane">
    <div class="panel">
      <div class="panel-title">⚠️ Overdue Documents — Awaiting Reply</div>
      <div id="ov-filter" style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap">
        <input id="ov-search" placeholder="🔍 Search..." oninput="filterOverdue()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none;flex:1;min-width:160px">
        <select id="ov-sort" onchange="filterOverdue()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none">
          <option value="days">Sort: Most Overdue</option>
          <option value="doc">Sort: Doc No.</option>
          <option value="disc">Sort: Discipline</option>
        </select>
      </div>
      <div id="ov-list" style="max-height:450px;overflow-y:auto"></div>
      <div id="ov-count" style="font-size:11px;color:var(--mu);margin-top:8px;text-align:right"></div>
    </div>
  </div>

  <!-- TAB: EXECUTIVE SUMMARY -->
  <div id="tab-executive" class="tab-pane">
    <div id="exec-content">
      <div style="text-align:center;padding:40px;color:var(--mu)">Loading...</div>
    </div>
  </div>

  <!-- TAB: AUDIT LOG -->
  <div id="tab-audit" class="tab-pane">
    <div class="panel">
      <div class="panel-title">📜 Activity Log — All Changes</div>
      <!-- Filters -->
      <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center">
        <select id="aud-pid" onchange="loadAudit()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none;min-width:160px">
          <option value="">All Projects</option>
        </select>
        <select id="aud-user" onchange="loadAudit()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none">
          <option value="">All Users</option>
        </select>
        <select id="aud-action" onchange="loadAudit()"
          style="padding:6px 10px;border:1.5px solid var(--bd);border-radius:var(--rd);font-size:12px;outline:none">
          <option value="">All Actions</option>
        </select>
        <button class="tbtn" onclick="loadAudit(true)" style="margin-left:auto">🔄 Refresh</button>
      </div>
      <!-- Table -->
      <div style="overflow-x:auto">
        <table class="dt-tbl" id="aud-tbl" style="min-width:800px">
          <thead><tr>
            <th style="width:140px">Time</th>
            <th style="width:100px">User</th>
            <th style="width:80px">Action</th>
            <th style="width:100px">Project</th>
            <th style="width:120px">Document</th>
            <th style="width:100px">Field</th>
            <th>Old Value</th>
            <th>New Value</th>
            <th>Detail</th>
          </tr></thead>
          <tbody id="aud-tbody"></tbody>
        </table>
      </div>
      <div id="aud-footer" style="display:flex;justify-content:space-between;align-items:center;margin-top:10px;font-size:11px;color:var(--mu)">
        <span id="aud-count"></span>
        <div style="display:flex;gap:8px">
          <button id="aud-prev" class="tbtn" onclick="auditPage(-1)" disabled>← Prev</button>
          <span id="aud-page" style="padding:5px 10px;font-size:11px">Page 1</span>
          <button id="aud-next" class="tbtn" onclick="auditPage(1)">Next →</button>
        </div>
      </div>
    </div>
  </div>

</div>

<!-- Modals -->
<div class="overlay hidden" id="newproj-modal">
  <div class="modal" style="max-width:420px">
    <div class="mhdr"><span>➕ New Project</span><button class="xbtn" onclick="closeM('newproj-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg"><label>Project ID</label><input id="np-id" placeholder="e.g. 24CP01"></div>
        <div class="fg"><label>Short Code</label><input id="np-code" placeholder="e.g. CP01"></div>
        <div class="fg full"><label>Project Name</label><input id="np-name" placeholder="Full project name"></div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('newproj-modal')">Cancel</button>
      <button class="btn btn-pr" id="cpbtn" onclick="createProject()">Create</button>
    </div>
  </div>
</div>

<!-- ADMIN MODAL (dashboard) -->
<div class="overlay hidden" id="admin-modal">
  <div class="modal" style="max-width:780px">
    <div class="mhdr"><span>⚙ Admin Panel</span><button class="xbtn" onclick="closeM('admin-modal')">✕</button></div>
    <div class="mbody" id="admin-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('admin-modal')">Close</button></div>
  </div>
</div>

{SHARED_JS}
<script>
const ROLE='{role}';
let STATS=[],OVERDUE_DATA=[],EXEC_DATA=null;
let PR_ANALYTICS=null;
let pChart,sChart,tChart,aChart,qChart,arChart,prTradeChart;
let _currentDisc='',_currentPid='';

// ── Dark mode ──────────────────────────────────────────────
function toggleDark(){{
  document.body.classList.toggle('dark');
  const on=document.body.classList.contains('dark');
  document.getElementById('darkBtn').textContent=on?'☀️':'🌙';
  localStorage.setItem('dcr_dark',on?'1':'');
}}
if(localStorage.getItem('dcr_dark')){{
  document.body.classList.add('dark');
  document.getElementById('darkBtn').textContent='☀️';
}}

// ── Tab switching ──────────────────────────────────────────
function showTab(name){{
  document.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));
  const pane=document.getElementById('tab-'+name);
  if(pane)pane.classList.add('active');
  if(name==='analytics'){{analyticsLoaded=false;setTimeout(()=>loadAnalytics(),50);}}
  if(name==='overdue'){{overdueLoaded=false;loadOverdue();}}
  if(name==='executive'){{EXEC_DATA=null;loadExecutive();}}
  if(name==='audit'){{loadAudit(true);}}
}}

// ── Init ──────────────────────────────────────────────────
  async function init(){{
    try{{
    const [stats]=await Promise.all([
        apiFetch('/api/dashboard_stats'),
      ]);
      STATS=stats;
      if(!STATS)return;
      // Populate project filter
      document.getElementById('proj-sel').innerHTML=
        '<option value="">All Projects</option>'+
      STATS.map(p=>`<option value="${{p.id}}">${{p.name}} (${{p.code}})</option>`).join('');
    // Populate discipline filter
    const discs=new Set();
    STATS.forEach(p=>(p.dt_stats||[]).forEach(dt=>
      (dt.disc_breakdown||[]).forEach(d=>discs.add(d.disc))));
        document.getElementById('disc-sel').innerHTML=
          '<option value="">All Disciplines</option>'+
          [...discs].sort().map(d=>`<option value="${{d}}">${{d}}</option>`).join('');
        await refreshOverviewPanels('');
        renderAll('','');
    }}finally{{document.getElementById('ld').style.display='none';}}
  }}

function filterProject(pid){{
  _currentPid=pid;refreshOverviewPanels(pid);renderAll(pid,_currentDisc);
  // Reload analytics/overdue if their tab is active
  const activeTab=document.querySelector('.tab-pane.active');
  if(activeTab?.id==='tab-analytics'){{analyticsLoaded=false;loadAnalytics();}}
  if(activeTab?.id==='tab-overdue'){{overdueLoaded=false;loadOverdue();}}
  if(activeTab?.id==='tab-executive'){{EXEC_DATA=null;loadExecutive();}}
}}
function filterDisc(disc){{_currentDisc=disc;renderAll(_currentPid,disc);}}

async function refreshOverviewPanels(pid){{
  const qs=pid?`?project_id=${{encodeURIComponent(pid)}}`:'';
  PR_ANALYTICS=await apiFetch('/api/pr_analytics_summary'+qs).catch(()=>null);
  renderPrAnalytics(PR_ANALYTICS);
}}

function getFiltered(pid,disc){{
  let d=pid?STATS.filter(s=>s.id===pid):STATS;
  if(disc){{
    d=d.map(p=>{{
      const dtFiltered=(p.dt_stats||[]).map(dt=>{{
        const discF=(dt.disc_breakdown||[]).filter(ds=>ds.disc===disc);
        const t=discF.reduce((s,r)=>s+r.total,0);
        const ap=discF.reduce((s,r)=>s+r.approved,0);
        const pe=discF.reduce((s,r)=>s+r.pending,0);
        const rj=discF.reduce((s,r)=>s+r.rejected,0);
        const ov=discF.reduce((s,r)=>s+r.overdue,0);
        return {{...dt,total:t,approved:ap,pending:pe,rejected:rj,overdue:ov,disc_breakdown:discF}};
      }}).filter(dt=>dt.total>0);
      const t=dtFiltered.reduce((s,dt)=>s+dt.total,0);
      const ap=dtFiltered.reduce((s,dt)=>s+dt.approved,0);
      return {{...p,dt_stats:dtFiltered,total:t,approved:ap,
        pct:t?Math.round(ap/t*100):0}};
    }}).filter(p=>p.total>0);
  }}
  return d;
}}

function renderAll(pid,disc){{
  const d=getFiltered(pid,disc);
  updateKPIs(d);renderCards(d,pid);renderCharts(d);renderDTTable(d);renderDiscTable(d);
}}

// ── KPIs ──────────────────────────────────────────────────

function renderPrAnalytics(data){{
  const el=document.getElementById('pr-panel');
  if(!el)return;
  if(!data){{
    el.innerHTML='<div style="font-size:11px;color:var(--mu)">Unavailable</div>';
    if(prTradeChart){{prTradeChart.destroy();prTradeChart=null;}}
    return;
  }}
  const topProjects=(data.top_projects||[]).length
    ? (data.top_projects||[]).map(p=>`<div style="display:flex;justify-content:space-between;gap:8px;padding:8px 0;border-bottom:1px solid var(--bd)">
        <span style="font-size:11px;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{p.project_name}}</span>
        <span style="font-size:11px;font-weight:700;color:var(--pr);white-space:nowrap">${{p.pr_count}}</span>
      </div>`).join('')
    : `<div style="font-size:11px;color:var(--mu)">No project data</div>`;
  el.innerHTML=`<div style="display:grid;grid-template-columns:minmax(180px,.24fr) minmax(280px,.36fr) minmax(340px,.4fr);gap:12px;align-items:stretch">
    <div style="background:var(--bg);border-radius:8px;padding:12px 14px;display:flex;flex-direction:column;justify-content:center">
      <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Total PRs</div>
      <div style="font-size:30px;font-weight:800;color:var(--pr);line-height:1.1;margin-top:6px">${{data.total_pr_records||0}}</div>
      <div style="font-size:11px;color:var(--mu);margin-top:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{data.top_project_name||'No project data'}}</div>
    </div>
    <div style="background:var(--bg);border-radius:8px;padding:12px 14px;display:flex;flex-direction:column;min-height:220px">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:10px">
        <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Top Projects</div>
        <div style="font-size:11px;color:var(--mu);white-space:nowrap">${{(data.top_projects||[]).length}} shown</div>
      </div>
      <div style="flex:1;overflow:auto;padding-right:2px">
        ${{topProjects}}
      </div>
    </div>
    <div style="background:var(--bg);border-radius:8px;padding:12px 14px;display:flex;flex-direction:column;min-height:220px">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:10px">
        <div>
          <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Top Trades by PR Count</div>
          <div style="font-size:18px;font-weight:800;color:var(--pr);line-height:1.15;margin-top:6px">${{data.top_trade_name||'No trade data'}}</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:18px;font-weight:800;color:var(--pr);line-height:1">${{data.top_trade_count||0}}</div>
          <div style="font-size:10px;color:var(--mu);margin-top:3px">${{data.trade_count_total||0}} trades</div>
        </div>
      </div>
      <div style="flex:1;min-height:220px"><canvas id="pr-trades-chart"></canvas></div>
    </div>
  </div>`;
  const tradeCanvas=document.getElementById('pr-trades-chart');
  if(prTradeChart){{prTradeChart.destroy();prTradeChart=null;}}
  if(!tradeCanvas)return;
  const trades=data.top_trades||[];
  prTradeChart=new Chart(tradeCanvas,{{
    type:'bar',
    data:{{
      labels:trades.map(t=>t.trade),
      datasets:[{{
        label:'PR Count',
        data:trades.map(t=>t.pr_count),
        backgroundColor:'#2563a8',
        borderRadius:4
      }}]
    }},
    options:{{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:'y',
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{beginAtZero:true,ticks:{{precision:0}}}},
        y:{{grid:{{display:false}}}}
      }}
    }}
  }});
}}

function updateKPIs(d){{
  const t=d.reduce((s,p)=>s+p.total,0),
    ap=d.reduce((s,p)=>s+p.approved,0),
    pe=d.reduce((s,p)=>s+p.pending,0),
    ov=d.reduce((s,p)=>s+p.overdue,0),
    rj=d.reduce((s,p)=>s+(p.rejected||0),0);
  document.getElementById('kpi-total').textContent=t;
  document.getElementById('kpi-approved').textContent=ap;
  document.getElementById('kpi-pending').textContent=pe;
  document.getElementById('kpi-overdue').textContent=ov;
  const rjEl=document.getElementById('kpi-rejected');
  if(rjEl)rjEl.textContent=rj;
  document.getElementById('kpi-pct').textContent=(t?Math.round(ap/t*100):0)+'%';
  document.getElementById('kpi-projs').textContent=d.length;
}}

// ── Project Cards ─────────────────────────────────────────
function renderCards(d,pid){{
  const g=document.getElementById('pgrid');g.innerHTML='';
  d.forEach(p=>{{
    const col=p.pct>=80?'#16a34a':p.pct>=50?'#f59e0b':'#ef4444';
    const a=document.createElement('a');a.href='/app?p='+p.id;a.className='pcard';
    a.innerHTML=`<div class="pchdr">
      <div><div style="color:rgba(255,255,255,.6);font-size:10px;font-weight:700">${{p.code}}</div>
        <div style="color:#fff;font-weight:700;font-size:13px">${{p.name}}</div></div>
      ${{p.overdue>0?`<span style="background:#ef4444;color:#fff;border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700">⚠${{p.overdue}}</span>`:''}}</div>
    <div class="pcbody">
      ${{p.client?`<div style="font-size:10px;color:var(--mu);margin-bottom:5px">👤 ${{p.client}}</div>`:''}}<div class="prow">
        <span style="font-size:11px;color:var(--mu)">Total</span><b>${{p.total}}</b></div>
      <div class="prow"><span style="font-size:11px;color:var(--mu)">Approved</span>
        <b style="color:#16a34a">${{p.approved}}</b></div>
      <div class="prow"><span style="font-size:11px;color:var(--mu)">Pending</span>
        <b style="color:#f59e0b">${{p.pending}}</b></div>
      <div class="prog"><div class="progf" style="width:${{p.pct}}%;background:${{col}}"></div></div>
      <div style="font-size:10px;color:${{col}};font-weight:700;margin-top:3px">${{p.pct}}%</div>
      ${{ROLE==='superadmin'?`<button class="del-pbtn" onclick="event.preventDefault();delProject('${{p.id}}','${{p.name}}')">🗑</button>`:''}}
    </div>`;
    g.appendChild(a);
  }});
  if(ROLE==='superadmin'&&!pid){{
    const add=document.createElement('div');add.className='addcard';
    add.innerHTML='<span style="font-size:28px">➕</span><span>New Project</span>';
    add.onclick=()=>openM('newproj-modal');g.appendChild(add);
  }}
}}

// ── Overview Charts ───────────────────────────────────────
function renderCharts(d){{
  if(pChart)pChart.destroy();
  pChart=new Chart(document.getElementById('cProj'),{{type:'bar',
    data:{{labels:d.map(s=>s.code),datasets:[
      {{label:'Approved',data:d.map(s=>s.approved),backgroundColor:'#16a34a',borderRadius:4}},
      {{label:'Pending', data:d.map(s=>s.pending), backgroundColor:'#f59e0b',borderRadius:4}},
      {{label:'Overdue', data:d.map(s=>s.overdue), backgroundColor:'#ef4444',borderRadius:4}}]}},
    options:{{responsive:true,plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}}}}}}}},
      scales:{{y:{{beginAtZero:true,grid:{{color:'rgba(0,0,0,.05)'}}}},
        x:{{grid:{{display:false}}}}}}}}}});
  const t=d.reduce((s,p)=>s+p.total,0),ap=d.reduce((s,p)=>s+p.approved,0),
    pe=d.reduce((s,p)=>s+p.pending,0),ov=d.reduce((s,p)=>s+p.overdue,0),
    rj=d.reduce((s,p)=>s+(p.rejected||0),0);
  if(sChart)sChart.destroy();
  sChart=new Chart(document.getElementById('cStatus'),{{type:'doughnut',
    data:{{labels:['Approved','Pending','Rejected','Overdue'],
      datasets:[{{data:[ap,pe,rj,ov],
        backgroundColor:['#16a34a','#f59e0b','#7c3aed','#ef4444'],
        borderWidth:3,borderColor:'#fff',hoverOffset:6}}]}},
    options:{{responsive:true,cutout:'65%',
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}}}}}},
        tooltip:{{callbacks:{{label:ctx=>` ${{ctx.label}}: ${{ctx.raw}} (${{t?Math.round(ctx.raw/t*100):0}}%)`}}}}}}}}}});
}}

// ── DT Table ──────────────────────────────────────────────
function renderDTTable(d){{
  const tbody=document.getElementById('dt-tbody'),empty=document.getElementById('dt-empty');
  tbody.innerHTML='';let rows=[],i=0;
  d.forEach(p=>{{(p.dt_stats||[]).forEach(dt=>{{rows.push({{...dt,pcode:p.code,pid:p.id}});}});}});
  if(!rows.length){{empty.style.display='block';return;}}
  empty.style.display='none';
  rows.forEach(r=>{{
    const pct=r.total?Math.round(r.approved/r.total*100):0;
    const col=pct>=80?'#16a34a':pct>=50?'#f59e0b':'#ef4444';
    const tr=document.createElement('tr');if(i%2)tr.className='alt';i++;
    tr.innerHTML=`<td style="font-size:10px;color:var(--mu)">${{r.pcode}}</td>
      <td style="font-weight:700;color:var(--pr)">${{r.code}}</td>
      <td>${{r.name}}</td>
      <td style="text-align:center;font-weight:700">${{r.total}}</td>
      <td style="text-align:center;color:#16a34a;font-weight:700">${{r.approved}}</td>
      <td style="text-align:center;color:#f59e0b;font-weight:700">${{r.pending}}</td>
      <td style="text-align:center;color:#7c3aed;font-weight:700">${{r.rejected||0}}</td>
      <td style="text-align:center;color:#ef4444;font-weight:700">${{r.overdue}}</td>`;
    tbody.appendChild(tr);
  }});
}}

// ── Disc Table ────────────────────────────────────────────
function renderDiscTable(data){{
  const tbody=document.getElementById('disc-tbody'),empty=document.getElementById('disc-empty');
  tbody.innerHTML='';let rows=0;
  data.forEach(p=>{{(p.dt_stats||[]).forEach(dt=>{{
    const disc=dt.disc_breakdown||[];if(!disc.length)return;
    disc.forEach((ds,di)=>{{
      rows++;
      const tr=document.createElement('tr');tr.className=rows%2===0?'alt':'';
      if(di===0){{
        const tdc=document.createElement('td');tdc.rowSpan=disc.length;
        tdc.style.cssText='font-size:10px;color:var(--mu);vertical-align:top;padding:6px 10px';
        tdc.textContent=p.code;tr.appendChild(tdc);
        const tdt=document.createElement('td');tdt.rowSpan=disc.length;
        tdt.style.cssText='font-weight:600;vertical-align:top;padding:6px 10px;font-size:11px';
        tdt.textContent=dt.code;tr.appendChild(tdt);
      }}
      const mk=(v,c)=>{{const td=document.createElement('td');
        td.style.cssText='text-align:center;padding:6px 8px;font-weight:600;color:'+c;
        td.textContent=v||'0';return td;}};
      const tdisc=document.createElement('td');tdisc.style.cssText='padding:6px 10px';
      tdisc.textContent=ds.disc;tr.appendChild(tdisc);
      tr.appendChild(mk(ds.total,'var(--pr)'));tr.appendChild(mk(ds.approved,'#16a34a'));
      tr.appendChild(mk(ds.pending,'#f59e0b'));tr.appendChild(mk(ds.rejected,'#7c3aed'));
      tr.appendChild(mk(ds.overdue,'#ef4444'));
      tbody.appendChild(tr);
    }});
  }});}});
  if(empty)empty.style.display=rows?'none':'block';
}}

// ── Analytics Tab ─────────────────────────────────────────
let analyticsLoaded=false;
async function loadAnalytics(forceReload=false){{
  if(analyticsLoaded&&!forceReload)return;
  analyticsLoaded=true;
  const pid=_currentPid;
  const [trend,aging,quality]=await Promise.all([
    apiFetch('/api/analytics/trend'+(pid?`?pid=${{pid}}`:'') ),
    apiFetch('/api/analytics/aging'+(pid?`?pid=${{pid}}`:'') ),
    apiFetch('/api/analytics/quality'+(pid?`?pid=${{pid}}`:'') ),
  ]);

  // Monthly Trend
  if(tChart)tChart.destroy();
  tChart=new Chart(document.getElementById('cTrend'),{{type:'bar',
    data:{{labels:trend.map(t=>t.month),datasets:[
      {{label:'Submitted',data:trend.map(t=>t.submitted),backgroundColor:'rgba(37,99,168,.25)',
        borderColor:'#2563a8',borderWidth:2,borderRadius:4,type:'bar'}},
      {{label:'Approved',data:trend.map(t=>t.approved),backgroundColor:'#16a34a',
        borderColor:'#16a34a',borderWidth:2,borderRadius:4,type:'bar'}}]}},
    options:{{responsive:true,plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}}}}}}}},
      scales:{{y:{{beginAtZero:true}},x:{{grid:{{display:false}}}}}}}}}});

  // Aging
  if(aChart)aChart.destroy();
  aChart=new Chart(document.getElementById('cAging'),{{type:'bar',
    data:{{labels:aging.map(a=>a.range+' Days'),
      datasets:[{{label:'Pending Docs',data:aging.map(a=>a.count),
        backgroundColor:['#16a34a','#f59e0b','#f97316','#ef4444'],borderRadius:6}}]}},
    options:{{responsive:true,plugins:{{legend:{{display:false}},
      tooltip:{{callbacks:{{label:ctx=>`${{ctx.raw}} documents`}}}}}},
      scales:{{y:{{beginAtZero:true}},x:{{grid:{{display:false}}}}}}}}}});

  // Quality
  if(qChart)qChart.destroy();
  qChart=new Chart(document.getElementById('cQuality'),{{type:'doughnut',
    data:{{labels:quality.map(q=>'Rev '+q.revisions),
      datasets:[{{data:quality.map(q=>q.count),
        backgroundColor:['#16a34a','#f59e0b','#f97316','#ef4444'],
        borderWidth:3,borderColor:'#fff',hoverOffset:6}}]}},
    options:{{responsive:true,cutout:'60%',
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}}}}}}}}}}}});

  // Approval Rate per Doc Type
  const d=getFiltered(_currentPid,_currentDisc);
  const dtLabels=[],dtRates=[];
  d.forEach(p=>(p.dt_stats||[]).forEach(dt=>{{
    if(dt.total>0){{dtLabels.push(dt.code);dtRates.push(dt.total?Math.round(dt.approved/dt.total*100):0);}}
  }}));
  if(arChart)arChart.destroy();
  arChart=new Chart(document.getElementById('cApprRate'),{{type:'bar',
    data:{{labels:dtLabels,datasets:[{{label:'Approval %',data:dtRates,
      backgroundColor:dtRates.map(r=>r>=80?'#16a34a':r>=50?'#f59e0b':'#ef4444'),
      borderRadius:4}}]}},
    options:{{responsive:true,indexAxis:'y',
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>`${{ctx.raw}}%`}}}}}},
      scales:{{x:{{beginAtZero:true,max:100,ticks:{{callback:v=>v+'%'}}}},y:{{grid:{{display:false}}}}}}}}}});
}}

// ── Overdue Tab ───────────────────────────────────────────
let overdueLoaded=false;
async function loadOverdue(){{
  if(overdueLoaded)return;
  overdueLoaded=true;
  const pid=_currentPid;
  OVERDUE_DATA=await apiFetch('/api/analytics/overdue'+(pid?`?pid=${{pid}}`:'') )||[];
  filterOverdue();
}}

function filterOverdue(){{
  const q=(document.getElementById('ov-search')?.value||'').toLowerCase();
  const sort=document.getElementById('ov-sort')?.value||'days';
  let data=[...OVERDUE_DATA];
  if(q)data=data.filter(r=>
    (r.docNo||'').toLowerCase().includes(q)||
    (r.title||'').toLowerCase().includes(q)||
    (r.discipline||'').toLowerCase().includes(q)||
    (r.dt_code||'').toLowerCase().includes(q));
  data.sort((a,b)=>sort==='days'?b.days_overdue-a.days_overdue:
    sort==='doc'?(a.docNo||'').localeCompare(b.docNo||''):
    (a.discipline||'').localeCompare(b.discipline||''));
  const list=document.getElementById('ov-list');
  if(!data.length){{list.innerHTML='<div style="text-align:center;padding:32px;color:var(--mu)">✅ No overdue documents</div>';return;}}
  list.innerHTML=data.map(r=>{{
    const bg=r.days_overdue>21?'er':r.days_overdue>14?'warn':'warn';
    return `<div class="ov-row">
      <span class="ov-badge ${{bg}}">${{r.days_overdue}}d</span>
      <div style="flex:1;min-width:0">
        <div style="font-weight:700;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{r.docNo}}</div>
        <div style="color:var(--mu);font-size:10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{r.title||'—'}}</div>
      </div>
      <span style="font-size:10px;color:var(--mu);white-space:nowrap">${{r.discipline||'—'}}</span>
      <span style="font-size:10px;background:#f0f4f8;padding:2px 7px;border-radius:4px;white-space:nowrap">${{r.dt_code}}</span>
      <span style="font-size:10px;color:var(--mu);white-space:nowrap">${{(r.issuedDate||'').slice(0,10)}}</span>
    </div>`;
  }}).join('');
  document.getElementById('ov-count').textContent=`Showing ${{data.length}} overdue document${{data.length!==1?'s':''}}`;
}}

// ── Executive Summary ─────────────────────────────────────
async function loadExecutive(){{
  if(EXEC_DATA)return;
  EXEC_DATA=await apiFetch('/api/executive_summary');
  if(!EXEC_DATA)return;
  const s=EXEC_DATA.summary;
  const col=s.completion_pct>=80?'#16a34a':s.completion_pct>=50?'#f59e0b':'#ef4444';
  document.getElementById('exec-content').innerHTML=`
    <div class="panel">
      <div class="panel-title">📋 Executive Summary — Generated ${{EXEC_DATA.generated_at}}</div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin-bottom:16px">
        ${{[['Total Docs',s.total,'var(--pr)'],['Approved',s.approved,'#16a34a'],
           ['Pending',s.pending,'#f59e0b'],['Rejected',s.rejected,'#7c3aed'],
           ['Overdue',s.overdue,'#ef4444'],['Completion',s.completion_pct+'%',col]]
          .map(([l,v,c])=>`<div style="background:var(--bg);border-radius:8px;padding:10px;text-align:center">
            <div style="font-size:22px;font-weight:800;color:${{c}}">${{v}}</div>
            <div style="font-size:9px;color:var(--mu);font-weight:700;text-transform:uppercase">${{l}}</div>
          </div>`).join('')}}
      </div>
      <table class="dt-tbl" style="margin-bottom:16px">
        <thead><tr><th>Project</th><th>Code</th>
          <th style="text-align:center">Total</th>
          <th style="text-align:center">Approved</th>
          <th style="text-align:center">%</th>
          <th style="text-align:center">Overdue</th></tr></thead>
        <tbody>
          ${{EXEC_DATA.projects.map((p,i)=>{{
            const col2=p.pct>=80?'#16a34a':p.pct>=50?'#f59e0b':'#ef4444';
            return `<tr class="${{i%2?'alt':''}}">
              <td style="font-weight:600">${{p.name}}</td>
              <td style="color:var(--pr);font-weight:700">${{p.code}}</td>
              <td style="text-align:center;font-weight:700">${{p.total}}</td>
              <td style="text-align:center;color:#16a34a;font-weight:700">${{p.approved}}</td>
              <td style="text-align:center;font-weight:700;color:${{col2}}">${{p.pct}}%</td>
              <td style="text-align:center;color:#ef4444;font-weight:700">${{p.overdue||0}}</td>
            </tr>`;
          }}).join('')}}
        </tbody>
      </table>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div>
          <div class="panel-title" style="margin-top:0">⏳ Aging Summary</div>
          ${{EXEC_DATA.aging.map(a=>`
            <div style="display:flex;justify-content:space-between;align-items:center;
              padding:6px 0;border-bottom:1px solid var(--bd);font-size:12px">
              <span>${{a.range}} Days</span>
              <span style="font-weight:700;color:${{a.range==='>21'?'#ef4444':'#f59e0b'}}">${{a.count}} docs</span>
            </div>`).join('')}}
        </div>
        <div>
          <div class="panel-title" style="margin-top:0">⚠️ Top Overdue</div>
          ${{EXEC_DATA.top_overdue.slice(0,5).map(r=>`
            <div style="padding:5px 0;border-bottom:1px solid var(--bd);font-size:11px">
              <div style="font-weight:700">${{r.docNo}}</div>
              <div style="color:var(--mu)">${{r.discipline||'—'}} · <span style="color:#ef4444;font-weight:700">${{r.days_overdue}}d overdue</span></div>
            </div>`).join('')||'<div style="color:var(--mu);padding:10px 0">No overdue documents ✅</div>'}}
        </div>
      </div>
    </div>`;
}}

// ── Project management ────────────────────────────────────
async function createProject(){{
  const id=document.getElementById('np-id')?.value.trim();
  const code=document.getElementById('np-code')?.value.trim();
  const name=document.getElementById('np-name')?.value.trim();
  if(!id||!name){{toast('Fill in Project ID and Name','er');return;}}
  const btn=document.getElementById('cpbtn');
  btn.disabled=true;btn.textContent='Creating...';
  try{{
    const r=await fetch('/api/projects/create',{{method:'POST',credentials:'include',
      headers:{{'Content-Type':'application/json','X-Requested-With':'XMLHttpRequest'}},
      body:JSON.stringify({{id,code:code||id,name}})}});
    const d=await r.json();
    if(d.ok){{toast('✔ Project created','ok');closeM('newproj-modal');
      setTimeout(()=>location.reload(),800);}}
    else toast(d.error||'Error','er');
  }}catch(e){{toast('Network error','er');}}
  finally{{btn.disabled=false;btn.textContent='Create';}}
}}

async function delProject(pid,name){{
  if(!confirm(`Delete project "${{name}}"?\nThis will delete ALL records, documents and settings. This cannot be undone.`))return;
  const r=await apiFetch('/api/projects/delete/'+pid,{{method:'POST'}});
  if(r&&r.ok){{toast('Project deleted','ok');setTimeout(()=>location.reload(),800);}}
  else toast((r&&r.error)||'Error','er');
}}

// ── Admin Panel ─────────────────────────────────────────
async function openAdmin(){{
  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  if(!users||!projects)return;
  const body=document.getElementById('admin-body');body.innerHTML='';
  const utitle=document.createElement('div');utitle.className='stitle';utitle.textContent='👥 Users';body.appendChild(utitle);
  for(const u of users){{
    const row=document.createElement('div');row.className='urow';
    row.innerHTML=`<span style="flex:1;font-weight:600">👤 ${{u.username}}</span>
      <span class="badge" style="background:#fef3c7;color:#92400e">${{u.role.toUpperCase()}}</span>
      ${{u.username!=='admin'?`<button class="btn btn-sc btn-sm" onclick="chgPw('${{u.username}}')">🔑 PW</button>
        <button class="btn btn-er btn-sm" onclick="delUsr('${{u.username}}')">✕</button>`:
        '<span style="font-size:10px;color:var(--mu)">(protected)</span>'}}`;
    body.appendChild(row);
    if(u.role!=='superadmin'){{
      const ad=document.createElement('div');
      ad.style.cssText='padding:4px 10px 10px 32px;border-bottom:1px solid var(--bd);margin-bottom:4px';
      ad.innerHTML='<div style="font-size:10px;color:var(--mu);margin-bottom:6px">Project access:</div>';
      const assigned=await apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]);
      const pl=document.createElement('div');pl.style.cssText='display:flex;flex-wrap:wrap;gap:5px';
      projects.forEach(p=>{{
        const isOn=assigned.includes(p.id);
        const btn=document.createElement('button');
        btn.style.cssText='padding:3px 10px;border-radius:4px;cursor:pointer;font-size:11px;font-weight:700;font-family:inherit;transition:all .15s;border:2px solid '+(isOn?'#f0a500':'#e2e8f0')+';background:'+(isOn?'#1a3a5c':'#f8fafc')+';color:'+(isOn?'#fff':'#94a3b8');
        btn.textContent=p.code;btn.title=p.name;
        if(isOn)btn.dataset.on='1';
        btn.onclick=async()=>{{
          const on=!!btn.dataset.on;
          await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:on?'unassign':'assign',username:u.username,project_id:p.id}})}});
          if(on){{delete btn.dataset.on;btn.style.borderColor='#e2e8f0';btn.style.background='#f8fafc';btn.style.color='#94a3b8';}}
          else{{btn.dataset.on='1';btn.style.borderColor='#f0a500';btn.style.background='#1a3a5c';btn.style.color='#fff';}}
          toast((btn.dataset.on?'✔ Assigned: ':'Removed: ')+p.name,'ok');
        }};
        pl.appendChild(btn);
      }});
      ad.appendChild(pl);body.appendChild(ad);
    }}
  }}
  const at=document.createElement('div');at.className='stitle';at.textContent='➕ Add User';body.appendChild(at);
  const ar=document.createElement('div');
  ar.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:8px;align-items:end">
    <div class="fg"><label>Username</label><input id="nu-name" placeholder="username"></div>
    <div class="fg"><label>Role</label><select id="nu-role">
      <option value="editor">Editor</option><option value="viewer">Viewer</option>
      <option value="admin">Admin</option></select></div>
    <div class="fg"><label>Password</label><input id="nu-pw" type="password"></div>
    <button class="btn btn-pr btn-sm" onclick="addUsr()">Add</button></div>`;
  body.appendChild(ar);openM('admin-modal');
}}
async function addUsr(){{
  const name=document.getElementById('nu-name')?.value.trim().toLowerCase();
  const role=document.getElementById('nu-role')?.value;
  const pw=document.getElementById('nu-pw')?.value;
  if(!name||!pw){{toast('Username and password required','er');return;}}
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'add',username:name,role,password:pw}})}});
  if(r&&r.ok){{toast('✔ User added','ok');closeM('admin-modal');openAdmin();}}
  else toast((r&&r.error)||'Error','er');
}}
async function delUsr(u){{
  if(!confirm('Delete user: '+u+'?'))return;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'delete',username:u}})}});
  if(r&&r.ok){{toast('Deleted','wa');closeM('admin-modal');openAdmin();}}
}}
async function chgPw(u){{
  const pw=prompt('New password for '+u+':');if(!pw)return;
  await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'change_password',username:u,password:pw}})}});
  toast('✔ Password changed','ok');
}}

// ── Audit Log ────────────────────────────────────────────────
let _auditOffset=0,_auditHasMore=true;
const ACTION_COLORS={{
  'ADD':'#166534','EDIT':'#1d4ed8','DELETE':'#991b1b',
  'LOGIN':'#6b7280','LOGIN_FAIL':'#dc2626','EXPORT_EXCEL':'#7c3aed'
}};

async function loadAudit(reset=false){{
  if(reset)_auditOffset=0;
  const pid=document.getElementById('aud-pid')?.value||'';
  const user=document.getElementById('aud-user')?.value||'';
  const action=document.getElementById('aud-action')?.value||'';
  const params=new URLSearchParams();
  if(pid)params.set('pid',pid);
  if(user)params.set('username',user);
  if(action)params.set('action',action);
  params.set('offset',_auditOffset);
  const data=await apiFetch('/api/audit?'+params);
  if(!data)return;

  // Populate filters on first load
  if(reset||_auditOffset===0){{
    const pSel=document.getElementById('aud-pid');
    const curPid=pSel.value;
    pSel.innerHTML='<option value="">All Projects</option>'+
      (data.projects||[]).map(p=>`<option value="${{p.id}}"${{p.id===curPid?' selected':''}}>${{p.code}} — ${{p.name}}</option>`).join('');

    const uSel=document.getElementById('aud-user');
    const curU=uSel.value;
    uSel.innerHTML='<option value="">All Users</option>'+
      (data.users||[]).map(u=>`<option value="${{u}}"${{u===curU?' selected':''}}>${{u}}</option>`).join('');

    const aSel=document.getElementById('aud-action');
    const curA=aSel.value;
    aSel.innerHTML='<option value="">All Actions</option>'+
      (data.actions||[]).map(a=>`<option value="${{a}}"${{a===curA?' selected':''}}>${{a}}</option>`).join('');
  }}

  const tbody=document.getElementById('aud-tbody');
  tbody.innerHTML='';
  const rows=data.rows||[];
  _auditHasMore=rows.length===100;

  if(!rows.length){{
    tbody.innerHTML='<tr><td colspan="9" style="text-align:center;padding:24px;color:var(--mu)">No activity found</td></tr>';
  }}else{{
    rows.forEach((r,i)=>{{
      const tr=document.createElement('tr');
      tr.className=i%2===0?'':'alt';
      const ts=new Date(r.ts);
      const tsStr=ts.toLocaleDateString('en-GB')+' '+ts.toLocaleTimeString('en-GB',{{hour:'2-digit',minute:'2-digit'}});
      const actionColor=ACTION_COLORS[r.action]||'#374151';
      const actionBg=r.action==='ADD'?'#bbf7d0':r.action==='EDIT'?'#dbeafe':r.action==='DELETE'?'#fee2e2':r.action==='LOGIN'?'#f3f4f6':'#fef3c7';
      tr.innerHTML=`
        <td style="font-size:10px;color:var(--mu);white-space:nowrap">${{tsStr}}</td>
        <td style="font-weight:600;font-size:11px">${{r.username||''}}</td>
        <td><span style="background:${{actionBg}};color:${{actionColor}};font-size:9px;font-weight:700;padding:2px 7px;border-radius:10px;white-space:nowrap">${{r.action||''}}</span></td>
        <td style="font-size:10px;color:var(--mu)">${{r.project_id||''}}</td>
        <td style="font-size:11px;font-weight:600;color:var(--pr)">${{r.doc_no||''}}</td>
        <td style="font-size:10px;color:var(--mu)">${{r.field_name||''}}</td>
        <td style="font-size:10px;color:#dc2626;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.old_value||''}}">${{r.old_value||''}}</td>
        <td style="font-size:10px;color:#166534;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.new_value||''}}">${{r.new_value||''}}</td>
        <td style="font-size:10px;color:var(--mu)">${{r.detail||''}}</td>`;
      tbody.appendChild(tr);
    }});
  }}

  const page=Math.floor(_auditOffset/100)+1;
  document.getElementById('aud-count').textContent=`Showing ${{rows.length}} entries`;
  document.getElementById('aud-page').textContent=`Page ${{page}}`;
  document.getElementById('aud-prev').disabled=_auditOffset===0;
  document.getElementById('aud-next').disabled=!_auditHasMore;
}}

function auditPage(dir){{
  _auditOffset=Math.max(0,_auditOffset+dir*100);
  loadAudit();
}}


init();
</script>
</body></html>"""



def render_register(u, proj):
    pid      = proj["id"]
    uname    = u["username"] if u else "guest"
    role     = u["role"]     if u else "guest"
    editable = can_edit(pid)
    btns, _, rlbl, rbg = _user_info_html(u)
    sc_json  = json.dumps(STATUS_COLORS)

    dts       = db.get_doc_types(pid)
    logo_l    = db.get_logo(pid,"logo_left")
    logo_r    = db.get_logo(pid,"logo_right")
    logo_html = ""
    if logo_l: logo_html += f'<img src="/api/logo/{pid}/logo_left" style="height:48px;object-fit:contain;flex-shrink:0">'
    if logo_r: logo_html += f'<img src="/api/logo/{pid}/logo_right" style="height:48px;object-fit:contain;flex-shrink:0;margin-left:8px">'

    DEFAULT_PROJ_FIELDS = [("code","Code"),("name","Project Name"),("startDate","Start"),("endDate","End"),
                   ("client","Client"),("landlord","Landlord"),("pmo","PMO"),
                   ("mainConsultant","Consultant"),("mepConsultant","MEP"),("contractor","Contractor")]
    custom_labels = proj.get("_labels") or {}
    PROJ_FIELDS = [(k, custom_labels.get(k, lbl)) for k, lbl in DEFAULT_PROJ_FIELDS]

    projbar = "".join(
        f'<div class="pf"><span class="pf-lbl">{lbl}</span>'
        f'<span class="pf-val" data-key="{key}">{proj.get(key,"") or "—"}</span></div>'
        for key,lbl in PROJ_FIELDS
        if proj.get(key,"").strip())  # hide empty fields

    tabs_html = "".join(
        f'<button class="tab-btn" data-id="{dt["id"]}" onclick="switchTab(\'{dt["id"]}\')">'
        f'<span>{dt["code"]}</span><span class="tcnt" id="cnt-{dt["id"]}">0</span></button>'
        for dt in dts)

    _sa_only = role == 'superadmin'
    _hol_btn = " <button class='tool-btn purple' onclick='openSettings()'>🗓 Holidays</button>" if _sa_only else ''
    _col_btn = "<button class='tool-btn purple' onclick='manageColumns()'>⚙ Columns</button>" if _sa_only else ''
    _lst_btn = "<button class='tool-btn' onclick='openLists()'>📋 Lists</button>" if _sa_only else ''
    edit_btns = (f'<button class="tool-btn" onclick="addRecord()">➕ Add</button>'
                 f'{_col_btn}'
                 f'{_hol_btn}'
                 f'{_lst_btn}'
                 f'<button class="tool-btn" onclick="editProject()">🏗 Project</button>'
                 if editable else
                 '<span style="font-size:11px;color:rgba(255,255,255,.5);padding:4px 8px">'
                 '👁 Read-only — <a href="/login" style="color:#f0a500">login to edit</a></span>')

    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DCR — {proj.get("name","Register")}</title>
{BASE_CSS}
<style>
.hidden{{display:none!important}}
body{{height:100vh;display:flex;flex-direction:column;overflow:hidden}}
@media print{{
  #topbar,#tabbar,.toolrow,.bulkbar,#statusbar,.acts{{display:none!important}}
  body{{height:auto;overflow:visible}}
  #main{{overflow:visible;height:auto}}
  #tblwrap{{overflow:visible;height:auto}}
  #regtbl{{font-size:10px}}
  #regtbl th,#regtbl td{{padding:4px 6px!important;white-space:normal!important}}
  @page{{size:A4 landscape;margin:10mm}}
}}
#projbar{{background:#fff;border-bottom:2px solid var(--pr);padding:3px 12px;
  display:flex;align-items:center;overflow-x:auto;flex-shrink:0;gap:4px}}
.pf{{display:flex;flex-direction:column;padding:0 10px;border-right:1px solid var(--bd)}}
.pf:last-of-type{{border-right:none}}
.pf-lbl{{font-size:9px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.4px}}
.pf-val{{font-size:11px;white-space:nowrap}}
#tabsbar{{background:#0f2640;display:flex;align-items:center;overflow-x:auto;flex-shrink:0;
  padding:0 8px;scrollbar-width:thin}}
#tabsbar::-webkit-scrollbar{{height:3px}}
#tabsbar::-webkit-scrollbar-thumb{{background:rgba(255,255,255,.3)}}
.tab-btn{{display:flex;align-items:center;gap:5px;padding:9px 12px;background:transparent;
  border:none;border-bottom:3px solid transparent;color:rgba(255,255,255,.55);cursor:pointer;
  font-family:inherit;font-size:11px;font-weight:600;white-space:nowrap;transition:all .15s}}
.tab-btn:hover{{color:#fff;background:rgba(255,255,255,.07)}}
.tab-btn.active{{color:#fff;border-bottom-color:var(--ac)}}
.tcnt{{background:rgba(255,255,255,.2);border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700}}
.tab-btn.active .tcnt{{background:var(--ac);color:#000}}
.tab-add{{padding:5px 10px;background:rgba(255,255,255,.1);border:1px dashed rgba(255,255,255,.35);
  color:rgba(255,255,255,.7);border-radius:4px;cursor:pointer;font-size:16px;margin-left:6px}}
.tab-add:hover{{background:rgba(255,255,255,.2)}}
#toolbar{{background:#fff;border-bottom:1px solid var(--bd);padding:5px 10px;
  display:flex;align-items:center;gap:5px;flex-shrink:0;flex-wrap:wrap}}
.tool-btn{{display:flex;align-items:center;gap:4px;padding:5px 10px;background:var(--bg);
  border:1px solid var(--bd);border-radius:var(--rd);cursor:pointer;font-size:11px;
  font-family:inherit;color:var(--tx);transition:all .15s;white-space:nowrap}}
.tool-btn:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.tool-btn.purple:hover{{background:#7c3aed;border-color:#7c3aed}}
.tool-btn.teal:hover{{background:#0891b2;border-color:#0891b2}}
.tool-dd{{position:relative}}
.tool-dd-menu{{position:absolute;top:calc(100% + 4px);left:0;background:#fff;border:1.5px solid var(--bd);border-radius:6px;box-shadow:0 8px 24px rgba(0,0,0,.15);z-index:300;min-width:210px;overflow:hidden}}
.tool-dd-menu button{{display:block;width:100%;text-align:left;padding:9px 14px;border:none;background:none;cursor:pointer;font-size:12px;font-family:inherit;color:#1e2a3a;white-space:nowrap}}
.tool-dd-menu button:hover{{background:#f0f4f8;color:var(--pr)}}
#srchbox{{flex:1;min-width:150px;max-width:260px;padding:5px 10px 5px 28px;border:1px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none;
  background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='%236b7a94' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E") no-repeat 7px center}}
#srchbox:focus{{border-color:var(--pl);box-shadow:0 0 0 2px rgba(37,99,168,.1)}}
#main{{flex:1;overflow:hidden;display:flex;flex-direction:column}}
#tblwrap{{flex:1;overflow:auto}}
table{{width:100%;border-collapse:collapse;min-width:900px;font-size:12px}}
thead{{position:sticky;top:0;z-index:10}}
th{{background:var(--pr);color:#fff;padding:8px;text-align:left;font-weight:600;
  white-space:nowrap;border-right:1px solid rgba(255,255,255,.1);cursor:pointer;user-select:none;position:relative}}
th:hover{{background:var(--pl)}}
.frow th{{background:#eef1f7;padding:2px 4px;cursor:default;position:sticky;top:33px;z-index:9}}
.frow th:hover{{background:#eef1f7}}
.frow input,.frow select{{width:100%;padding:3px 6px;border:1px solid var(--bd);
  border-radius:3px;font-size:10px;font-family:inherit;background:#fff;outline:none}}
td{{padding:5px 8px;border-bottom:1px solid #edf0f5;border-right:1px solid #f3f4f6;
  vertical-align:middle;max-width:220px;word-break:break-word}}
tr:hover td{{background:rgba(37,99,168,.10);transition:background .1s}}
tr.ov td{{background:#fff5f5}}
tr.rv td{{color:var(--mu)}}
tr.alt td{{background:#fafbfd}}
.sr{{text-align:center;color:var(--mu);font-size:10px;min-width:28px}}
.chkcell{{text-align:center;width:28px;padding:4px!important}}
.chkcell input{{width:14px;height:14px;cursor:pointer;accent-color:var(--pr)}}
.acts{{white-space:nowrap;width:64px}}
.act{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.act:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.act.del:hover{{background:var(--er);border-color:var(--er)}}
.sbadge{{display:inline-block;border-radius:10px;padding:2px 9px;font-size:10px;font-weight:700}}
.flink{{color:var(--pl);text-decoration:underline;cursor:pointer;font-size:11px}}
.ovdate{{color:#dc2626;font-weight:700}}
.mlcell{{white-space:pre-line!important;word-break:break-word}}
#bulkbar{{display:none;background:#1a3a5c;color:#fff;padding:5px 14px;align-items:center;gap:10px;font-size:12px;flex-shrink:0}}
#bulkbar.show{{display:flex}}
#sbar{{background:var(--pr);color:rgba(255,255,255,.75);padding:3px 14px;font-size:10px;display:flex;gap:16px;flex-shrink:0}}
.rz{{position:absolute;right:0;top:0;bottom:0;width:6px;cursor:col-resize;z-index:1}}
.rz:hover,.rz.rzg{{background:var(--ac)}}
.ms-con{{border:1px solid var(--bd);border-radius:var(--rd);min-height:34px;padding:3px;
  background:#fff;cursor:pointer;position:relative}}
.ms-tag{{display:inline-flex;align-items:center;gap:3px;background:var(--pr);color:#fff;
  border-radius:3px;padding:2px 7px;font-size:10px;margin:2px}}
.ms-rm{{cursor:pointer;opacity:.7}}
.ms-ph{{color:var(--mu);font-size:11px;padding:3px 5px}}
.ms-dd{{position:absolute;left:0;right:0;top:100%;background:#fff;border:1.5px solid var(--bd);
  border-radius:var(--rd);z-index:500;max-height:260px;overflow-y:auto;
  box-shadow:0 8px 24px rgba(0,0,0,.12);margin-top:2px}}
.ms-opt{{padding:9px 12px;cursor:pointer;font-size:12px;display:flex;align-items:center;gap:8px;border-bottom:1px solid #f0f4f8}}
.ms-opt:last-child{{border-bottom:none}}
.ms-opt:hover{{background:#f0f4f8;color:var(--pr)}}
.ms-opt.sel{{background:#eff6ff}}
.empty{{text-align:center;padding:60px 20px;color:var(--mu)}}
.slist{{list-style:none;display:flex;flex-direction:column;gap:3px;
  max-height:190px;overflow-y:auto;border:1px solid var(--bd);border-radius:var(--rd);padding:4px}}
.sitem{{display:flex;align-items:center;gap:8px;padding:4px 8px;background:var(--bg);
  border-radius:3px;font-size:11px}}
.sitem .nm{{flex:1}}
.sitem button{{padding:2px 8px;font-size:10px;border:1px solid var(--bd);background:#fff;
  border-radius:3px;cursor:pointer}}
.sitem button:hover{{background:var(--er);color:#fff;border-color:var(--er)}}
.addrow{{display:flex;gap:6px;margin-top:6px}}
.addrow input{{flex:1;padding:5px 8px;border:1px solid var(--bd);border-radius:3px;
  font-size:11px;font-family:inherit;outline:none}}
@media(max-width:768px){{
  .pf{{padding:0 6px;min-width:70px}}
  .tab-btn{{padding:7px 8px;font-size:10px}}
  .tool-btn{{padding:4px 7px;font-size:10px}}
  table{{font-size:11px}} td,th{{padding:4px 5px}}
  .modal{{width:97%;max-height:96vh}}
  .fgrid{{grid-template-columns:1fr}}
}}
</style></head><body>

<div id="topbar">
  <span style="font-size:20px">📋</span>
  <span style="font-weight:700;font-size:14px;letter-spacing:.3px">Document Control Register</span>
  <div class="sp"></div>
  <a href="/" class="tb-btn">📊 Dashboard</a>
  {btns}
  <span style="color:rgba(255,255,255,.45);padding:0 4px">|</span>
  <span style="color:rgba(255,255,255,.8);font-size:11px">👤 {uname}
    <span style="background:{rbg};border-radius:3px;padding:1px 7px;font-size:9px;font-weight:700">{rlbl}</span>
  </span>
</div>

<div id="projbar">
  {logo_html}
  {projbar}
  {'<button onclick="editProject()" style="margin-left:auto;background:var(--pr);color:#fff;border:none;padding:5px 12px;border-radius:var(--rd);cursor:pointer;font-size:11px;font-family:inherit;flex-shrink:0">✏ Edit</button>' if editable else ''}
</div>

<div id="tabsbar">
  {tabs_html}
  {'<button class="tab-add" onclick="addDocType()" title="Add Type">＋</button>' if editable else ''}
</div>

<div id="toolbar">
  {edit_btns}
  <div class="tool-dd" id="exp-dd">
    <button class="tool-btn teal" onclick="toggleExpDD(event)">📥 Export ▾</button>
    <div class="tool-dd-menu hidden" id="exp-menu">
      <button onclick="doExport();closeExpDD()">📊 Excel — This Tab</button>
      <button onclick="doExportAll();closeExpDD()">📊 Excel — All Tabs</button>
      <button onclick="doExportPDF();closeExpDD()">📄 PDF — This Tab</button>
      <button onclick="doExportAllPDF();closeExpDD()">📄 PDF — All Tabs</button>
    </div>
  </div>
  <button class="tool-btn teal" onclick="doPrint()">🖨 Print</button>
  {'<button class="tool-btn teal" onclick="openImport()">📤 Import</button>' if editable else ''}
  <input type="text" id="srchbox" placeholder="Search..." oninput="doSearch()">
</div>

<div id="bulkbar">
  <span id="bulkcnt">0 selected</span>
  <button class="btn btn-er btn-sm" onclick="bulkDel()">🗑 Delete</button>
  <button class="btn btn-sm" style="background:rgba(255,255,255,.15);color:#fff;border-color:rgba(255,255,255,.3)" onclick="clearSel()">✕ Cancel</button>
</div>

<div id="main">
  <div id="tblwrap">
    <table id="regtbl"><thead id="thead"></thead><tbody id="tbody"></tbody></table>
    <div class="empty hidden" id="empty" style="display:none"><div style="font-size:48px;margin-bottom:10px">📁</div><p style="color:var(--mu)">No records yet — click ➕ Add to create one</p></div>
  </div>
</div>

<div id="sbar">
  <span id="s-total">Total: 0</span>
  <span id="s-show">Showing: 0</span>
  <span id="s-ov">Overdue: 0</span>
  <span style="margin-left:auto" id="s-clock"></span>
</div>

<!-- ADD/EDIT RECORD -->
<div class="overlay hidden" id="rec-modal">
  <div class="modal" style="max-width:820px">
    <div class="mhdr"><span id="rec-title">Add Document</span>
      <button class="xbtn" onclick="closeM('rec-modal')">✕</button></div>
    <div class="mbody"><div class="fgrid" id="rec-form"></div></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('rec-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveRecord()">Save</button>
    </div>
  </div>
</div>

<!-- PROJECT MODAL -->
<div class="overlay hidden" id="proj-modal">
  <div class="modal" style="max-width:820px">
    <div class="mhdr"><span>🏗 Project Information</span>
      <button class="xbtn" onclick="closeM('proj-modal')">✕</button></div>
    <div class="mbody" id="proj-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('proj-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveProject()">💾 Save Project</button>
    </div>
  </div>
</div>

<!-- LISTS -->
<div class="overlay hidden" id="lists-modal">
  <div class="modal">
    <div class="mhdr"><span>📋 Dropdown Lists</span>
      <button class="xbtn" onclick="closeM('lists-modal')">✕</button></div>
    <div class="mbody" id="lists-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('lists-modal')">Close</button></div>
  </div>
</div>

<!-- ADD DOC TYPE -->
<div class="overlay hidden" id="dt-modal">
  <div class="modal" style="max-width:420px">
    <div class="mhdr"><span>Add Document Type</span>
      <button class="xbtn" onclick="closeM('dt-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg"><label>Code</label><input id="dt-code" placeholder="e.g. MS"></div>
        <div class="fg"><label>Name</label><input id="dt-name" placeholder="e.g. Method Statement"></div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('dt-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveDocType()">Add</button>
    </div>
  </div>
</div>

<!-- COLUMNS -->
<div class="overlay hidden" id="col-modal">
  <div class="modal">
    <div class="mhdr"><span>⚙ Manage Columns</span>
      <button class="xbtn" onclick="closeM('col-modal')">✕</button></div>
    <div class="mbody" id="col-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="openAddCol()">+ Add Column</button>
      <button class="btn btn-pr" onclick="closeM('col-modal');loadRecords()">Done</button>
    </div>
  </div>
</div>

<!-- ADD COLUMN -->
<div class="overlay hidden" id="addcol-modal">
  <div class="modal" style="max-width:460px">
    <div class="mhdr"><span>Add Column</span>
      <button class="xbtn" onclick="closeM('addcol-modal')">✕</button></div>
    <div class="mbody">
      <div class="fgrid">
        <div class="fg full"><label>Column Name</label><input id="col-name" placeholder="e.g. Review Duration"></div>
        <div class="fg full"><label>Type</label>
          <select id="col-type" onchange="onColType(this.value)">
            <option value="text">📝 Text</option>
            <option value="number">🔢 Number</option>
            <option value="date">📅 Date</option>
            <option value="dropdown">📋 Dropdown</option>
            <option value="link">🔗 Hyperlink</option>
            <option value="duration_calc">⏱ Duration (working days)</option>
          </select>
        </div>
        <div class="fg full" id="cg-list" style="display:none">
          <label>Dropdown Source</label><select id="col-list"></select></div>
        <div class="fg full" id="cg-ds" style="display:none">
          <label>Start Date Column</label><select id="col-ds"></select></div>
        <div class="fg full" id="cg-de" style="display:none">
          <label>End Date Column</label><select id="col-de"></select></div>
        <div class="fg full" id="cg-info" style="display:none">
          <div style="background:#f0f4fa;border-radius:6px;padding:9px 12px;font-size:11px;color:#1a3a5c">
            ⏱ Working days between two date columns — excludes Fridays + Egyptian holidays
          </div>
        </div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('addcol-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveAddCol()">Add</button>
    </div>
  </div>
</div>

<!-- IMPORT -->
<div class="overlay hidden" id="import-modal">
  <div class="modal" style="max-width:480px">
    <div class="mhdr"><span>📤 Import Excel / CSV</span>
      <button class="xbtn" onclick="closeM('import-modal')">✕</button></div>
    <div class="mbody">
      <p style="font-size:12px;color:var(--mu);margin-bottom:12px">Select .xlsx or .csv file. Headers must match register columns. For full-workbook import, Excel sheets will be matched to document types automatically.</p>
      <input type="file" id="imp-file" accept=".csv,.xlsx,.xls" style="font-size:12px">
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('import-modal')">Cancel</button>
      <button class="btn btn-sc" id="imp-project-btn" onclick="doImportProject()">Import Full Workbook</button>
      <button class="btn btn-pr" onclick="doImport()">Import</button>
    </div>
  </div>
</div>

<!-- ADMIN MODAL (register page) -->
<div class="overlay hidden" id="admin-modal">
  <div class="modal" style="max-width:780px">
    <div class="mhdr"><span>⚙ Admin Panel</span><button class="xbtn" onclick="closeM('admin-modal')">✕</button></div>
    <div class="mbody" id="admin-body"></div>
    <div class="mfoot"><button class="btn btn-sc" onclick="closeM('admin-modal')">Close</button></div>
  </div>
</div>

{SHARED_JS}
<script>
const PID='{pid}', ROLE='{role}', CAN_EDIT={'true' if editable else 'false'};
const SC={sc_json};
const PROJ_FIELDS=[['code','Code'],['name','Project Name'],['startDate','Start Date'],['endDate','End Date'],
  ['client','Client'],['landlord','Landlord'],['pmo','PMO'],['mainConsultant','Consultant'],
  ['mepConsultant','MEP'],['contractor','Contractor']];
const state={{tab:null,cols:[],recs:null,sortCol:null,sortDir:'asc',filters:{{}},editId:null,lists:{{}},prItemsCache:{{}}}};

function isPRTab(){{
  const dt=state.dtList?.find(d=>d.id===state.tab)||{{}};
  const code=(dt.code||dt.id||'').toString().toUpperCase();
  const name=(dt.name||'').toString().toLowerCase();
  return code==='PR' || name.includes('requisition') || name.includes('purchase request');
}}

function getPrDetailsColKey(){{
  const cols=state.cols||[];
  const norm=v=>String(v||'').toLowerCase().replace(/[^a-z0-9]+/g,'');
  for(const c of cols){{
    const key=String(c.col_key||'');
    const label=String(c.label||'').trim();
    if(label==='PR Details'||key==='prDetails'||key==='pr_details')return c.col_key;
  }}
  for(const c of cols){{
    const nk=norm(c.col_key),nl=norm(c.label);
    if(nk==='prdetails'||nl==='prdetails')return c.col_key;
  }}
  for(const c of cols){{
    const nk=norm(c.col_key),nl=norm(c.label);
    if(nk.includes('pr')&&nk.includes('detail'))return c.col_key;
    if(nl.includes('pr')&&nl.includes('detail'))return c.col_key;
  }}
  return null;
}}

function getPrManualDetails(row){{
  const detailsKey=getPrDetailsColKey();
  return detailsKey?String(row?.[detailsKey]||'').replace(/\s+/g,' ').trim():'';
}}

function getPrAutoSummary(row){{
  const items=state.prItemsCache[row?._id]||[];
  if(items.length){{
    const names=items
      .filter(it=>String(it?.row_type||'item').toLowerCase()!=='header')
      .map(it=>String(it?.item_name||'').trim())
      .filter(Boolean);
    if(names.length){{
      let summary=names.slice(0,2).join(', ');
      if(names.length>2)summary+=' ...';
      return summary;
    }}
  }}
  return '';
}}

function getPrSummary(row){{
  const manual=getPrManualDetails(row);
  if(manual)return manual;
  const auto=getPrAutoSummary(row);
  if(!auto)return '';
  return auto.length>80?auto.slice(0,77)+'...':auto;
}}

function escHtml(v){{
  return String(v==null?'':v)
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'",'&#39;');
}}

// Init
(async()=>{{
  await Promise.all([loadDTs(), loadLists()]);
  updateClock(); setInterval(updateClock,60000);
}})();

function updateClock(){{document.getElementById('s-clock').textContent=new Date().toLocaleString('en-GB');}}

async function loadDTs(keepTab=false){{
  const dts=await apiFetch('/api/doc_types/'+PID); if(!dts)return;
  renderTabs(dts);
  await refreshCounts();
  if(!keepTab){{
    const tab=new URLSearchParams(location.search).get('tab');
    if(dts.length) switchTab(tab&&dts.find(d=>d.id===tab)?tab:dts[0].id);
  }}
}}

async function refreshCounts(){{
  const cnts=await apiFetch('/api/counts/'+PID); if(!cnts)return;
  Object.entries(cnts).forEach(([id,n])=>{{const el=document.getElementById('cnt-'+id);if(el)el.textContent=n;}});
}}

async function loadLists(force=false){{
  if(!force&&Object.keys(state.lists).length) return;
  const d=await apiFetch('/api/lists/'+PID); if(d) state.lists=d;
}}

function renderTabs(dts){{
  state.dtList=dts;
  const bar=document.getElementById('tabsbar');
  bar.querySelectorAll('.tab-btn').forEach(b=>b.remove());
  const addBtn=bar.querySelector('.tab-add');
  dts.forEach(dt=>{{
    const btn=document.createElement('button');
    btn.className='tab-btn'+(dt.id===state.tab?' active':'');
    btn.dataset.id=dt.id; btn.title=dt.name;
    btn.innerHTML=`<span>${{dt.code}}</span><span class="tcnt" id="cnt-${{dt.id}}">0</span>`;
    btn.onclick=()=>switchTab(dt.id);
    if(CAN_EDIT) btn.oncontextmenu=e=>{{e.preventDefault();tabMenu(dt.id,e);}};
    bar.insertBefore(btn,addBtn||null);
  }});
}}

function switchTab(id){{
  state.tab=id;state.recs=null;state.filters={{}};state.sortCol=null;
  document.getElementById('srchbox').value='';
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active',b.dataset.id===id));
  loadRecords();
}}

function tabMenu(id,e){{
  const old=document.getElementById('tabctx');if(old)old.remove();
  const m=document.createElement('div');m.id='tabctx';
  m.style.cssText=`position:fixed;left:${{e.clientX}}px;top:${{e.clientY}}px;background:#fff;
    border:1px solid var(--bd);border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.15);z-index:9999;min-width:140px;overflow:hidden`;
  const dt=state.dtList?.find(d=>d.id===id)||{{}};
  m.innerHTML=`
    <div onclick="renameDT('${{id}}','${{dt.code||id}}','${{(dt.name||'').replace(/'/g,\"\\\\'\")}}')"
      style="padding:8px 14px;cursor:pointer;font-size:12px"
      onmouseover="this.style.background='#f0f4f8'" onmouseout="this.style.background=''">✏ Rename</div>
    <div onclick="delDT('${{id}}')" style="padding:8px 14px;cursor:pointer;font-size:12px;color:#ef4444"
      onmouseover="this.style.background='#fef2f2'" onmouseout="this.style.background=''">🗑 Delete Type</div>`;
  document.body.appendChild(m);
  setTimeout(()=>document.addEventListener('click',()=>m.remove(),{{once:true}}),10);
}}

async function renameDT(id,oldCode,oldName){{
  const newCode=prompt('Tab code (short):',oldCode);
  if(newCode===null)return;
  const newName=prompt('Full name:',oldName);
  if(newName===null)return;
  await apiFetch('/api/doc_types/'+PID+'/'+id,{{method:'PATCH',
    body:JSON.stringify({{code:newCode.trim().toUpperCase(),name:newName.trim()}})}});
  await loadDTs(true);toast('✔ Renamed','ok');
}}

async function delDT(id){{
  if(!confirm('Delete type and ALL its records?'))return;
  await apiFetch('/api/doc_types/'+PID+'/'+id,{{method:'DELETE'}});
  if(state.tab===id)state.tab=null;
  await loadDTs(); toast('Deleted','wa');
}}

async function loadRecords(){{
  if(!state.tab)return;
  const search=document.getElementById('srchbox').value.trim();
  const [data, widths]=await Promise.all([
    apiFetch('/api/records/'+PID+'/'+state.tab+(search?'?search='+encodeURIComponent(search):'')),
    apiFetch('/api/col_width/'+PID+'/'+state.tab)
  ]);
  if(!data)return;
  state.recs=data.records; state.cols=data.columns.filter(c=>c.visible);
  state.prItemsCache=data.pr_items_map||{{}};
  state.colWidths=widths||{{}};
  const cnt=document.getElementById('cnt-'+state.tab); if(cnt)cnt.textContent=data.count;
  buildHead(); requestAnimationFrame(()=>initRz()); renderRows();
}}

function buildHead(){{
  const head=document.getElementById('thead'); head.innerHTML='';
  const hr=document.createElement('tr');
  const chk=document.createElement('th'); chk.className='chkcell';
  if(CAN_EDIT)chk.innerHTML='<input type="checkbox" id="chkall" onchange="selAll(this.checked)">';
  hr.appendChild(chk);
  const sr=document.createElement('th');sr.textContent='Sr.';
  const _srW=state.colWidths&&state.colWidths['_sr'];
  sr.dataset.key='_sr';
  sr.style.cssText=(_srW?'width:'+_srW+'px;min-width:'+_srW+'px;max-width:'+_srW+'px;':'width:34px;')
    +'white-space:normal;word-break:break-word;cursor:default';
  hr.appendChild(sr);
  state.cols.forEach(col=>{{
    const th=document.createElement('th');th.dataset.key=col.col_key;
    const w=state.colWidths&&state.colWidths[col.col_key];
    if(w)th.style.cssText='width:'+w+'px;min-width:'+w+'px;max-width:'+w+'px;white-space:normal;word-break:break-word';
    else th.style.cssText='white-space:normal;word-break:break-word';
    const sortInd=state.sortCol===col.col_key?(state.sortDir==='asc'?' ↑':' ↓'):'';
    th.innerHTML='<span class="th-lbl">'+col.label+sortInd+'</span>';
    if(!['auto_date','auto_num'].includes(col.col_type))th.onclick=()=>sortBy(col.col_key);
    else th.style.cursor='default';
    hr.appendChild(th);
  }});
  const at=document.createElement('th');at.textContent='Actions';at.style.cssText='width:64px;cursor:default';hr.appendChild(at);
  head.appendChild(hr);
  const fr=document.createElement('tr');fr.className='frow';
  fr.appendChild(document.createElement('th'));fr.appendChild(document.createElement('th'));
  state.cols.forEach(col=>{{
    const th=document.createElement('th');
    if(['auto_date','auto_num'].includes(col.col_type)){{fr.appendChild(th);return;}}
    if(col.col_type==='dropdown'&&col.list_name){{
      const sel=document.createElement('select');
      sel.innerHTML='<option value="">All</option>'+(state.lists[col.list_name]||[]).map(o=>`<option ${{state.filters[col.col_key]===o?'selected':''}}>${{o}}</option>`).join('');
      sel.onchange=()=>{{state.filters[col.col_key]=sel.value;renderRows();}};
      th.appendChild(sel);
    }}else{{
      const inp=document.createElement('input');inp.value=state.filters[col.col_key]||'';
      inp.oninput=()=>{{state.filters[col.col_key]=inp.value;renderRows();}};
      th.appendChild(inp);
    }}
    fr.appendChild(th);
  }});
  fr.appendChild(document.createElement('th'));
  head.appendChild(fr);
}}

function parseDocNo(docNo){{
  // Extract base number and rev from "DS-001 REV00" format
  const m=(docNo||'').match(/^([A-Za-z]+)-([0-9]+) REV([0-9]+)$/i);
  if(m)return{{prefix:m[1],num:parseInt(m[2]),rev:parseInt(m[3])}};
  return{{prefix:docNo||'',num:0,rev:0}};
}}

function sortByDocNo(rows){{
  // Group by base (prefix+num), then sort revs within group
  const groups={{}};
  rows.forEach(r=>{{
    const p=parseDocNo(r.docNo||'');
    const key=p.prefix+'-'+String(p.num).padStart(6,'0');
    if(!groups[key])groups[key]={{baseNum:p.num,prefix:p.prefix,rows:[]}};
    groups[key].rows.push({{...r,_parsedNum:p.num,_parsedRev:p.rev}});
  }});
  // Sort groups by base number, then rows within group by rev
  const sorted=[];
  Object.values(groups)
    .sort((a,b)=>a.prefix.localeCompare(b.prefix)||a.baseNum-b.baseNum)
    .forEach(g=>{{
      g.rows.sort((a,b)=>a._parsedRev-b._parsedRev);
      g.rows.forEach(r=>sorted.push(r));
    }});
  return sorted;
}}

function validateDocNo(docNo,existingRecs,editId){{
  if(!docNo)return'Document No. is required';
  const p=parseDocNo(docNo);
  if(!p.num&&!docNo.includes('-'))return'Invalid format. Use: CODE-001 REV00';
  // Duplicate check
  const dup=existingRecs.find(r=>r._id!==editId&&(r.docNo||'').toLowerCase()===docNo.toLowerCase());
  if(dup)return'Document No. already exists: '+docNo;
  // REV check: base must exist
  if(p.rev>0){{
    const hasBase=existingRecs.some(r=>{{
      const bp=parseDocNo(r.docNo||'');
      return bp.prefix===p.prefix&&bp.num===p.num&&bp.rev===0&&r._id!==editId;
    }});
    if(!hasBase)return'Cannot add REV'+String(p.rev).padStart(2,'0')+' — REV00 not found for this document';
  }}
  // Sequence gap check (warning only)
  if(p.rev===0&&p.num>1){{
    const nums=existingRecs
      .filter(r=>{{const bp=parseDocNo(r.docNo||'');return bp.prefix===p.prefix&&bp.rev===0&&r._id!==editId;}})
      .map(r=>parseDocNo(r.docNo||'').num);
    const maxExist=nums.length?Math.max(...nums):0;
    if(p.num>maxExist+1)return'GAP: '+p.prefix+'-'+String(maxExist+1).padStart(3,'0')+' to '+p.prefix+'-'+String(p.num-1).padStart(3,'0')+' are missing';
  }}
  return null;
}}

function checkGap(docNo,recs,editId){{
  const p=parseDocNo(docNo);if(!p||p.rev>0)return null;
  // Find max existing base num for same prefix
  let maxN=0;
  recs.forEach(r=>{{
    if(r._id===editId)return;
    const rp=parseDocNo(r.docNo||'');
    if(rp&&rp.prefix===p.prefix&&rp.rev===0)maxN=Math.max(maxN,rp.num);
  }});
  if(p.num>maxN+1)return`Sequence gap detected: Next expected is ${{p.prefix}}-${{String(maxN+1).padStart(3,'0')}} REV00, but you entered ${{p.prefix}}-${{String(p.num).padStart(3,'0')}} REV00`;
  return null;
}}

function renderRows(){{
  const body=document.getElementById('tbody');body.innerHTML='';
  const isPrTab=isPRTab();
  const prDetailsKey=isPrTab?getPrDetailsColKey():null;
  let rows=state.recs.filter(r=>{{
    for(const[k,v]of Object.entries(state.filters)){{if(v&&!String(r[k]||'').toLowerCase().includes(v.toLowerCase()))return false;}}
    return true;
  }});
  if(state.sortCol){{
    rows.sort((a,b)=>{{
      const va=String(a[state.sortCol]||'').toLowerCase(),vb=String(b[state.sortCol]||'').toLowerCase();
      return state.sortDir==='asc'?(va>vb?1:-1):(va<vb?1:-1);
    }});
  }}else{{
    rows=sortByDocNo(rows);
  }}
  const emptyEl=document.getElementById('empty');
  const tblEl=document.getElementById('regtbl');
  if(rows.length===0){{
    emptyEl.style.display='none';
    tblEl.style.display='';
    const tr=document.createElement('tr');
    const td=document.createElement('td');
    td.colSpan=state.cols.length+3;
    td.style.cssText='padding:28px 12px;text-align:center;color:var(--mu);font-size:12px';
    td.textContent=(state.recs&&state.recs.length)?'No records found':'No records found';
    tr.appendChild(td);
    body.appendChild(tr);
  }}else{{
    emptyEl.style.display='none';
    tblEl.style.display='';
  }}
  if(rows.length===0){{
    const ov=state.recs.filter(r=>r._overdue).length;
    document.getElementById('s-total').textContent='Total: '+state.recs.length;
    document.getElementById('s-show').textContent='Showing: 0';
    document.getElementById('s-ov').textContent='Overdue: '+ov;
    return;
  }}
  let sr=1;
  rows.forEach((row,idx)=>{{
    const tr=document.createElement('tr');
    if(row._overdue)tr.classList.add('ov');
    else if(row._isRev)tr.classList.add('rv');
    else if(idx%2===1)tr.classList.add('alt');
    const tc=document.createElement('td');tc.className='chkcell';
    if(CAN_EDIT){{const cb=document.createElement('input');cb.type='checkbox';cb.dataset.id=row._id;cb.onchange=updBulk;tc.appendChild(cb);}}
    tr.appendChild(tc);
    const tsr=document.createElement('td');tsr.className='sr';
    const _srTdW=state.colWidths&&state.colWidths['_sr'];
    if(_srTdW)tsr.style.cssText='width:'+_srTdW+'px;min-width:'+_srTdW+'px;max-width:'+_srTdW+'px';
    tsr.textContent=row._isRev?'':sr;tr.appendChild(tsr);
    state.cols.forEach(col=>{{
      const td=document.createElement('td');const key=col.col_key;let val='';
      const k=key.toLowerCase();
      const longTextMeta=getLongTextMeta(col);
      if(longTextMeta||isFloorField(col))td.classList.add('mlcell');
      if(key==='expectedReplyDate'){{val=row._expectedReplyDate||'';if(row._overdue&&val)td.classList.add('ovdate');}}
      else if(key==='duration')val=row._duration||'';
      else if(col.col_type==='duration_calc'){{
        const[ds,de]=(col.list_name||'issuedDate,actualReplyDate').split(',');
        val=calcWD(row[ds.trim()]||'',row[de.trim()]||'');
      }}
      else if(key==='issuedDate')val=row._issuedFmt||'';
      else if(key==='actualReplyDate')val=row._replyFmt||'';
      else if(col.col_type==='date'||col.col_type==='auto_date'){{
        val=row['_fmt_'+key]||row[key]||'';  // use pre-formatted version
      }}
      else if(key==='status'){{
        val=row[key]||'';
        if(val){{
          td.innerHTML=val.split(',').map(s=>{{s=s.trim();const[bg,fg]=SC[s]||['e5e7eb','374151'];
            return `<span class="sbadge" style="background:#${{bg}};color:#${{fg}}">${{s}}</span>`;}}).join('');
          tr.appendChild(td);return;
        }}
      }}
      else if(key==='fileLocation'){{
        const url=row[key]||'';
        if(url){{td.innerHTML=`<a class="flink" href="${{url}}" target="_blank">View</a>`;tr.appendChild(td);return;}}
      }}
      else if(isPrTab&&prDetailsKey&&key===prDetailsKey)val=getPrSummary(row);
      else val=String(row[key]||'');
      let displayVal=formatDisplayValue(col,val);
      if(typeof displayVal==='string'&&longTextMeta){{
        const NL=String.fromCharCode(10);
        displayVal=displayVal
          .replaceAll(' / ',NL)
          .replaceAll(' /',NL)
          .replaceAll('/ ',NL);
      }}
      td.textContent=displayVal;tr.appendChild(td);
    }});
    const ta=document.createElement('td');ta.className='acts';
    let acts='';
    if(isPrTab)acts+=`<button class="pr-toggle" onclick="togglePrItems('${{row._id}}', this)">Items</button> `;
    if(CAN_EDIT)acts+=`<button class="act" onclick="editRec('${{row._id}}')">✏</button> <button class="act del" onclick="delRec('${{row._id}}')">🗑</button>`;
    else if(!acts)acts='<span style="color:var(--mu);font-size:10px">—</span>';
    ta.innerHTML=acts;
    tr.appendChild(ta);body.appendChild(tr);
    if(!row._isRev)sr++;
  }});
  const ov=state.recs.filter(r=>r._overdue).length;
  document.getElementById('s-total').textContent='Total: '+state.recs.length;
  document.getElementById('s-show').textContent='Showing: '+rows.length;
  document.getElementById('s-ov').textContent='Overdue: '+ov;
}}

function calcWD(s,e){{
  if(!s||!e)return'';
  try{{let a=new Date(s),b=new Date(e);if(isNaN(a)||isNaN(b)||b<=a)return'0';
    let n=0,c=new Date(a);c.setDate(c.getDate()+1);
    while(c<=b){{if(c.getDay()!==5)n++;c.setDate(c.getDate()+1);}}return String(n);}}
  catch{{return'';}}
}}

function isFloorField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  return key==='floor'||label==='floor'||label==='floors';
}}

function isItemRefField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase().replace(/[_./-]+/g,' ');
  return key==='itemref'||(label.includes('item ref')&&label.includes('dwg'));
}}

function formatDisplayValue(col,val){{
  let text=String(val??'');
  if(!text)return '';
  if(isFloorField(col))return text.split(',').map(s=>s.trim()).filter(Boolean).join('\n');
  if(isItemRefField(col))return text.replace(/\\r\\n/g,'\\n').replace(/\\r/g,'\\n');
  return text;
}}

function getLongTextMeta(col){{
  const key=String(col?.col_key||'');
  const label=String(col?.label||'');
  const k=key.toLowerCase();
  const l=label.toLowerCase();
  const text=(k+' '+l).replace(/[_-]+/g,' ').replace(/\\s+/g,' ').trim();
  const isContentLike=['content','description'].some(v=>text.includes(v));
  const isRemarksLike=['remarks','notes'].some(v=>text.includes(v));
  const isMsLike=text.includes('ms ref')||text.includes('msref')||text.includes('ms reference')||
    (/\\bms\\b/.test(text)&&(text.includes('ref')||text.includes('reference')||text.includes('no')));
  const isItemRefLike=isItemRefField(col);
  if(isContentLike)return {{
    rows:5,
    style:'resize:vertical; min-height:120px',
    placeholder:'Use Enter to put each item on a separate line'
  }};
  if(isRemarksLike)return {{
    rows:3,
    style:'resize:vertical; min-height:80px',
    placeholder:'Use Enter for multiline remarks'
  }};
  if(isMsLike)return {{
    rows:3,
    style:'resize:vertical; min-height:80px',
    placeholder:'Use Enter to put each MS on a separate line'
  }};
  if(isItemRefLike)return {{
    rows:3,
    style:'resize:vertical; min-height:80px',
    placeholder:'Use Enter to put each item reference / DWG No. on a separate line'
  }};
  return null;
}}

function sortBy(key){{
  state.sortDir=state.sortCol===key?(state.sortDir==='asc'?'desc':'asc'):'asc';
  state.sortCol=key;buildHead();renderRows();
}}

async function fetchPrItems(recordId){{
  if(Object.prototype.hasOwnProperty.call(state.prItemsCache,recordId))return state.prItemsCache[recordId];
  try{{
    const r=await apiFetch('/api/pr_items/'+recordId);
    state.prItemsCache[recordId]=(r&&r.ok&&Array.isArray(r.items))?r.items:[];
  }}catch(e){{
    console.warn('PR items load failed for',recordId,e);
    state.prItemsCache[recordId]=[];
  }}
  return state.prItemsCache[recordId];
}}

function renderPrItemsTable(items, legacyText){{
  if(!items||!items.length){{
    const legacy=legacyText?`<div style="margin-top:8px"><div style="font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px">Legacy PR Details</div><div style="color:var(--mu);font-size:11px;white-space:pre-line">${{escHtml(legacyText)}}</div></div>`:'';
    return `<div class="pr-items-empty">No items added</div>${{legacy}}`;
  }}
  const rows=items.map(it=>String(it?.row_type||'item').toLowerCase()==='header'
    ? `<tr class="pr-section"><td colspan="4" style="padding:8px 8px;border-bottom:1px solid #d7dee8">
        <div style="font-weight:800;font-size:12px;color:var(--pr);background:#e8eef6;padding:8px 10px;border-radius:4px;letter-spacing:.3px">
          ${{escHtml(it.item_name||'')}}
        </div>
      </td></tr>`
    : `<tr>
        <td>${{escHtml(it.item_name||'')}}</td>
        <td>${{escHtml(it.unit||'')}}</td>
        <td>${{escHtml(it.quantity??'')}}</td>
        <td>${{escHtml(it.remarks||'')}}</td>
      </tr>`).join('');
  return `<table class="pr-items-table">
    <thead><tr><th>Item</th><th>Unit</th><th>Qty</th><th>Remarks</th></tr></thead>
    <tbody>${{rows}}</tbody>
  </table>`;
}}

async function togglePrItems(recordId, btn){{
  const tr=btn.closest('tr');
  if(!tr)return;
  let nxt=tr.nextElementSibling;
  if(!nxt||!nxt.classList||!nxt.classList.contains('pr-items-row')||nxt.dataset.id!==recordId){{
    const colSpan=state.cols.length+1;
    const wrap=document.createElement('div');wrap.className='pr-items-wrap';
    wrap.innerHTML='<div class="pr-items-empty">Loading...</div>';
    const td=document.createElement('td');td.colSpan=colSpan;td.appendChild(wrap);
    const row=document.createElement('tr');row.className='pr-items-row';row.dataset.id=recordId;row.appendChild(td);
    row.style.display='none';
    tr.parentNode.insertBefore(row, tr.nextSibling);
    nxt=row;
  }}
  const open=!nxt.classList.contains('open');
  if(open){{
    const items=await fetchPrItems(recordId);
    const rowData=state.recs.find(r=>r._id===recordId)||{{}};
    const detailsKey=getPrDetailsColKey();
    const colIdx=state.cols.findIndex(c=>c.col_key===detailsKey);
    if(colIdx>=0){{
      const summaryCell=tr.children[2+colIdx];
      if(summaryCell)summaryCell.textContent=getPrSummary(rowData);
    }}
    const legacyKey=getPrDetailsColKey();
    const legacyText=legacyKey?String((state.recs.find(r=>r._id===recordId)||{{}})[legacyKey]||'').trim():'';
    nxt.querySelector('.pr-items-wrap').innerHTML=renderPrItemsTable(items, legacyText);
    nxt.style.display='table-row';
    nxt.classList.add('open');
    btn.textContent='Hide';
  }}else{{
    nxt.classList.remove('open');
    nxt.style.display='none';
    btn.textContent='Items';
  }}
}}

let _rzDrag=null;
document.addEventListener('mousemove',e=>{{
  if(!_rzDrag)return;
  const w=Math.max(40,_rzDrag.sw+e.clientX-_rzDrag.sx);
  _rzDrag.th.style.width=w+'px';_rzDrag.th.style.minWidth=w+'px';
}});
document.addEventListener('mouseup',()=>{{
  if(!_rzDrag)return;
  _rzDrag.rz.classList.remove('rzg');
  document.body.style.cursor='';document.body.style.userSelect='none';
  _rzDrag=null;
}});
let _rzActive=null;
function initRz(){{
  document.querySelectorAll('#regtbl thead tr:first-child th[data-key]').forEach(th=>{{
    if(th.querySelector('.rz'))return;
    const rz=document.createElement('div');rz.className='rz';th.appendChild(rz);
    rz.addEventListener('mousedown',e=>{{
      e.stopPropagation();e.preventDefault();
      _rzActive={{th,sx:e.clientX,sw:th.offsetWidth,key:th.dataset.key}};
      rz.classList.add('rzg');
      document.body.style.cursor='col-resize';
      document.body.style.userSelect='none';
    }});
  }});
}}
document.addEventListener('mousemove',e=>{{
  if(!_rzActive)return;
  const w=Math.max(40,_rzActive.sw+(e.clientX-_rzActive.sx));
  _rzActive.th.style.width=w+'px';
  _rzActive.th.style.minWidth=w+'px';
  _rzActive.th.style.maxWidth=w+'px';
}});
let _rzSaveTimer=null;
document.addEventListener('mouseup',()=>{{
  if(!_rzActive)return;
  const {{th,key}}=_rzActive;
  const w=th.offsetWidth;
  th.querySelector('.rz')?.classList.remove('rzg');
  _rzActive=null;
  document.body.style.cursor='';
  document.body.style.userSelect='';
  // Save width if superadmin
  if(ROLE==='superadmin'&&key){{
    clearTimeout(_rzSaveTimer);
    _rzSaveTimer=setTimeout(()=>{{
      apiFetch('/api/col_width/'+PID+'/'+state.tab,{{
        method:'POST',body:JSON.stringify({{col_key:key,width_px:w}})
      }}).then(r=>{{if(r&&r.ok){{if(!state.colWidths)state.colWidths={{}};state.colWidths[key]=w;}}}});
    }},400);
  }}
}});

let _st;
function doSearch(){{clearTimeout(_st);_st=setTimeout(()=>loadRecords(),250);}}

// Add/Edit Record
function addRecord(){{state.editId=null;document.getElementById('rec-title').textContent='Add Document';buildForm(null);openM('rec-modal');}}
function editRec(id){{state.editId=id;const row=state.recs.find(r=>r._id===id);if(!row)return;document.getElementById('rec-title').textContent='Edit Document';buildForm(row);openM('rec-modal');}}

async function buildForm(row){{
  const allCols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!allCols)return;
  const AUTO=new Set(['expectedReplyDate','duration','_duration','_duration_today']);
  const grid=document.getElementById('rec-form');grid.innerHTML='';
  let nextNo='';
  if(!row){{const r=await apiFetch('/api/next_doc_no/'+PID+'/'+state.tab);nextNo=r?.next||'';}}
  const isPrTab=isPRTab();
  const prevCols=state.cols;
  state.cols=allCols.filter(c=>c.visible);
  const prDetailsKey=isPrTab?getPrDetailsColKey():null;
  state.cols=prevCols;
  for(const col of allCols){{
    if(AUTO.has(col.col_key))continue;
    const key=col.col_key;
    const longTextMeta=getLongTextMeta(col);
    const full=['title','fileLocation','itemRef'].includes(key)||!!longTextMeta;
    const grp=document.createElement('div');grp.className='fg'+(full?' full':'');
    const lbl=document.createElement('label');lbl.textContent=col.label;grp.appendChild(lbl);
    const val=row?.[key]||'';
    if(isPrTab&&prDetailsKey&&key===prDetailsKey){{
      const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val;
      ta.rows=3;
      ta.style.cssText='resize:vertical;min-height:80px';
      ta.placeholder='Leave blank to auto-generate from items';
      grp.appendChild(ta);
      const hint=document.createElement('div');
      hint.style.cssText='font-size:10px;color:var(--mu);margin-top:4px';
      hint.textContent='Leave blank to auto-generate from PR items.';
      grp.appendChild(hint);
      grid.appendChild(grp);
      continue;
    }}
    if(col.col_type==='date'){{const inp=document.createElement('input');inp.type='date';inp.id='f-'+key;inp.value=val;grp.appendChild(inp);}}
    else if(col.col_type==='dropdown'&&col.list_name){{
      // Add free-text input below multiselect
      const wrapper=document.createElement('div');
      wrapper.appendChild(buildMS(key,state.lists[col.list_name]||[],val));
      const freeInp=document.createElement('input');
      freeInp.id='f-free-'+key;freeInp.placeholder='Or type custom value...';
      freeInp.style.cssText='margin-top:4px;width:100%;padding:5px 8px;border:1px dashed var(--bd);border-radius:var(--rd);font-size:11px;color:var(--mu);outline:none';
      freeInp.title='Type a custom value not in the list';
      freeInp.onchange=()=>{{
        const v=freeInp.value.trim();
        if(!v)return;
        const ms=document.getElementById('f-'+key);
        if(ms&&ms.classList.contains('ms-con')){{
          const cur=ms.dataset.value||'';
          const arr=cur?cur.split(', ').filter(Boolean):[];
          if(!arr.includes(v))arr.push(v);
          ms.dataset.value=arr.join(', ');
          // Trigger re-render
          ms.dispatchEvent(new Event('_update'));
          freeInp.value='';
        }}
      }};
      wrapper.appendChild(freeInp);
      grp.appendChild(wrapper);
    }}
    else if(col.col_type==='docno'){{
      const inp=document.createElement('input');inp.id='f-'+key;
      if(row){{inp.value=val;}}
      else{{inp.value=nextNo;inp.placeholder=nextNo?'':state.tab+'-001 REV00';}}
      inp.style.cssText='font-family:Consolas,monospace;font-weight:600';
      grp.appendChild(inp);
    }}
    else if(longTextMeta){{
      const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val;
      ta.rows=longTextMeta.rows;
      ta.style.cssText=longTextMeta.style;
      ta.placeholder=longTextMeta.placeholder;
      grp.appendChild(ta);
    }}
    else{{const inp=document.createElement('input');inp.id='f-'+key;inp.value=val;if(col.col_type==='link')inp.placeholder='https://...';grp.appendChild(inp);}}
    grid.appendChild(grp);
  }}
  if(isPrTab){{
    const grp=document.createElement('div');grp.className='fg full pr-items-editor';
    const lbl=document.createElement('label');lbl.textContent='PR Items';grp.appendChild(lbl);
    const wrap=document.createElement('div');wrap.id='pr-items-editor';
    wrap.innerHTML=`
      <table>
        <thead><tr><th>Item Name</th><th>Unit</th><th>Qty</th><th>Remarks</th><th></th></tr></thead>
        <tbody id="pr-items-body"></tbody>
      </table>
      <div style="margin-top:8px;display:flex;gap:8px;align-items:center">
        <button class="btn btn-sc btn-sm" onclick="addPrItemRow()">+ Add Item</button>
        <button class="btn btn-sc btn-sm" onclick="addPrHeaderRow()">+ Add Section Header</button>
        <div id="pr-legacy" style="font-size:11px;color:var(--mu);white-space:pre-line;display:none"></div>
      </div>`;
    grp.appendChild(wrap);grid.appendChild(grp);
    const legacyText=prDetailsKey?String(row?.[prDetailsKey]||'').trim():'';
    initPrItemsEditor([], legacyText);
    if(row&&row._id)loadPrItemsForEdit(row._id, legacyText);
  }}
}}

function addPrItemRow(item={{}}){{
  const body=document.getElementById('pr-items-body');if(!body)return;
  const tr=document.createElement('tr');
  tr.dataset.rowType='item';
  tr.innerHTML=`
    <td><input class="pri-item"></td>
    <td><input class="pri-unit"></td>
    <td><input class="pri-qty"></td>
    <td><input class="pri-remarks"></td>
    <td><button type="button" class="btn btn-er btn-sm">✕</button></td>`;
  tr.querySelector('.pri-item').value=item.item_name||'';
  tr.querySelector('.pri-unit').value=item.unit||'';
  tr.querySelector('.pri-qty').value=item.quantity??'';
  tr.querySelector('.pri-remarks').value=item.remarks||'';
  tr.querySelector('button').onclick=()=>tr.remove();
  body.appendChild(tr);
}}

function addPrHeaderRow(item={{}}){{
  const body=document.getElementById('pr-items-body');if(!body)return;
  const tr=document.createElement('tr');
  tr.className='pr-head-edit';
  tr.dataset.rowType='header';
  tr.innerHTML=`
    <td colspan="4">
      <div class="pr-head-label">Section Header</div>
      <input class="pri-header" placeholder="Section title / description">
    </td>
    <td><button type="button" class="btn btn-er btn-sm">✕</button></td>`;
  tr.querySelector('.pri-header').value=item.item_name||'';
  tr.querySelector('button').onclick=()=>tr.remove();
  body.appendChild(tr);
}}

function initPrItemsEditor(items, legacyText){{
  const body=document.getElementById('pr-items-body');if(!body)return;
  body.innerHTML='';
  if(items&&items.length)items.forEach(it=>String(it?.row_type||'item').toLowerCase()==='header'?addPrHeaderRow(it):addPrItemRow(it));
  else addPrItemRow();
  const legacy=document.getElementById('pr-legacy');
  if(legacy){{
    legacy.textContent=legacyText?('Legacy Details:\\n'+legacyText):'';
    legacy.style.display=(!items||!items.length)&&legacyText?'block':'none';
  }}
}}

async function loadPrItemsForEdit(recordId, legacyText){{
  const items=await fetchPrItems(recordId);
  initPrItemsEditor(items, legacyText);
}}

function getPrItemsFromEditor(){{
  const body=document.getElementById('pr-items-body');if(!body)return [];
  const rows=[...body.querySelectorAll('tr')];
  return rows.map(r=>{{
    const row_type=(r.dataset.rowType||'item').toLowerCase();
    if(row_type==='header'){{
      const item_name=r.querySelector('.pri-header')?.value.trim()||'';
      return {{row_type:'header', item_name}};
    }}
    const item_name=r.querySelector('.pri-item')?.value.trim()||'';
    const unit=r.querySelector('.pri-unit')?.value.trim()||'';
    const quantity=r.querySelector('.pri-qty')?.value.trim()||'';
    const remarks=r.querySelector('.pri-remarks')?.value.trim()||'';
    return {{row_type:'item', item_name, unit, quantity, remarks}};
  }}).filter(it=>it.row_type==='header'?it.item_name:(it.item_name||it.unit||it.quantity||it.remarks));
}}

function buildMS(key,options,init){{
  const sel=init?init.split(',').map(s=>s.trim()).filter(Boolean):[];
  const con=document.createElement('div');con.className='ms-con';con.id='f-'+key;con.dataset.value=init||'';
  function render(){{
    con.innerHTML='';
    sel.forEach(v=>{{
      const t=document.createElement('span');t.className='ms-tag';
      t.innerHTML=`${{v}} <span class="ms-rm" data-v="${{v}}">✕</span>`;
      t.querySelector('.ms-rm').onclick=e=>{{e.stopPropagation();sel.splice(sel.indexOf(v),1);con.dataset.value=sel.join(', ');render();}};
      con.appendChild(t);
    }});
    if(!sel.length)con.innerHTML='<span class="ms-ph">Select or type below...</span>';
    con.dataset.value=sel.join(', ');
  }}
  con.onclick=e=>{{
    if(e.target.classList.contains('ms-rm'))return;
    const ex=document.querySelector('.ms-dd');if(ex){{ex.remove();return;}}
    const dd=document.createElement('div');dd.className='ms-dd';
    dd.style.cssText='position:absolute;left:0;right:0;top:calc(100% + 2px);background:#fff;border:1.5px solid var(--bd);border-radius:6px;z-index:500;box-shadow:0 8px 24px rgba(0,0,0,.15);max-height:260px;overflow-y:auto';
    options.forEach(opt=>{{
      const it=document.createElement('div');
      it.style.cssText='padding:10px 14px;display:flex;align-items:center;gap:10px;cursor:pointer;border-bottom:1px solid #f1f5f9;font-size:12px'+(sel.includes(opt)?';background:#eff6ff':'');
      const chk=document.createElement('input');chk.type='checkbox';chk.checked=sel.includes(opt);chk.style.cssText='flex-shrink:0;width:14px;height:14px;pointer-events:none';
      const lbl=document.createElement('span');lbl.textContent=opt;lbl.style.flex='1';
      it.appendChild(chk);it.appendChild(lbl);
      it.onclick=ev=>{{
        ev.stopPropagation();
        if(sel.includes(opt))sel.splice(sel.indexOf(opt),1);else sel.push(opt);
        con.dataset.value=sel.join(', ');render();
        chk.checked=sel.includes(opt);
        it.style.background=sel.includes(opt)?'#eff6ff':'';
      }};
      dd.appendChild(it);
    }});
    con.style.position='relative';con.appendChild(dd);
  }};
  document.addEventListener('click',e=>{{if(!con.contains(e.target))con.querySelector('.ms-dd')?.remove();}},true);
  render();return con;
}}

function durChoice(docNo){{
  return new Promise(resolve=>{{
    const ov=document.createElement('div');ov.className='overlay';ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center';
    ov.innerHTML=`<div style="background:#fff;border-radius:12px;padding:28px 24px;max-width:380px;width:90%;box-shadow:0 20px 60px rgba(0,0,0,.3)">
      <div style="font-size:16px;font-weight:700;margin-bottom:8px">⏱ Duration</div>
      <p style="font-size:13px;color:#64748b;margin-bottom:20px">Issue date is set but no reply date yet.<br>How should Duration be shown?</p>
      <div style="display:flex;flex-direction:column;gap:8px">
        <button id="dur-today" style="padding:10px 16px;background:#1a3a5c;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px">📅 Calculate from issue date → Today</button>
        <button id="dur-empty" style="padding:10px 16px;background:#f1f5f9;color:#374151;border:none;border-radius:6px;cursor:pointer;font-size:13px">— Leave Duration empty</button>
        <button id="dur-cancel" style="padding:8px;background:none;border:none;cursor:pointer;font-size:12px;color:#94a3b8">Cancel</button>
      </div></div>`;
    document.body.appendChild(ov);
    ov.querySelector('#dur-today').onclick=()=>{{document.body.removeChild(ov);resolve('today');}};
    ov.querySelector('#dur-empty').onclick=()=>{{document.body.removeChild(ov);resolve('empty');}};
    ov.querySelector('#dur-cancel').onclick=()=>{{document.body.removeChild(ov);resolve(null);}};
  }});
}}

async function saveRecord(){{
  const allCols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!allCols)return;
  const AUTO=new Set(['expectedReplyDate','duration','_duration','_duration_today']);
  const data={{}};
  for(const col of allCols){{
    if(AUTO.has(col.col_key))continue;
    const el=document.getElementById('f-'+col.col_key);if(!el)continue;
    if(el.classList.contains('ms-con'))data[col.col_key]=el.dataset.value||'';
    else if(el.tagName==='TEXTAREA')data[col.col_key]=el.value.replace(/\\r\\n/g,'\\n').replace(/\\r/g,'\\n');
    else data[col.col_key]=el.value.trim();
  }}
  if(!data.docNo){{toast('Document No. required','er');return;}}
    // Duration is computed server-side automatically
  const valErr=validateDocNo(data.docNo,state.recs||[],state.editId);
  if(valErr&&valErr.startsWith('GAP')){{
    if(!confirm('⚠ Sequence gap detected: '+valErr.replace('GAP:','')+'. Continue anyway?'))return;
  }}else if(valErr){{toast('⚠ '+valErr,'er');return;}}
  if(state.editId)data._id=state.editId;
  const r=await apiFetch('/api/records/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify(data)}});
  if(r&&r.ok){{
    if(isPRTab()){{
      const items=getPrItemsFromEditor();
      const recId=r.id||state.editId;
      try{{
        const prRes=await apiFetch('/api/pr_items/'+recId,{{method:'POST',body:JSON.stringify({{items}})}});
        if(prRes&&prRes.ok)state.prItemsCache[recId]=items;
        else {{toast('Items save failed','er');return;}}
      }}catch(e){{
        toast('Items save failed: '+e.message,'er');return;
      }}
    }}
    closeM('rec-modal');const savedTab=state.tab;await loadRecords();await refreshCounts();toast(state.editId?'Updated':'Added','ok');
  }}
  else toast('Error saving','er');
}}

async function delRec(id){{
  if(!confirm('Delete this record?'))return;
  await apiFetch('/api/records/'+id,{{method:'DELETE'}});
  await loadRecords();await refreshCounts();toast('Deleted','wa');
}}

// Bulk
function updBulk(){{
  const checked=document.querySelectorAll('.chkcell input[data-id]:checked');
  document.getElementById('bulkcnt').textContent=checked.length+' selected';
  document.getElementById('bulkbar').classList.toggle('show',checked.length>0);
  const all=document.querySelectorAll('.chkcell input[data-id]');
  const ca=document.getElementById('chkall');if(ca)ca.checked=all.length>0&&checked.length===all.length;
}}
function selAll(v){{document.querySelectorAll('.chkcell input[data-id]').forEach(cb=>cb.checked=v);updBulk();}}
function clearSel(){{document.querySelectorAll('.chkcell input').forEach(cb=>cb.checked=false);updBulk();}}
async function bulkDel(){{
  const ids=[...document.querySelectorAll('.chkcell input[data-id]:checked')].map(cb=>cb.dataset.id);
  if(!ids.length||!confirm('Delete '+ids.length+' records?'))return;
  let ok=0;for(const id of ids){{const r=await apiFetch('/api/records/'+id,{{method:'DELETE'}});if(r&&r.ok)ok++;}}
  clearSel();await loadRecords();await refreshCounts();toast('✔ Deleted '+ok,'ok');
}}

// Project Modal
async function editProject(){{
  const proj=await apiFetch('/api/project/'+PID);if(!proj)return;
  const body=document.getElementById('proj-body');body.innerHTML='';
  const customLabels=proj._labels||{{}};
  // Field visibility toggles
  const visNote=document.createElement('div');
  visNote.style.cssText='font-size:10px;color:var(--mu);margin-bottom:8px;padding:6px 8px;background:var(--bg);border-radius:4px';
  visNote.textContent='💡 Leave any field empty to hide it from the project bar';
  body.appendChild(visNote);
  const grid=document.createElement('div');grid.className='fgrid';
  PROJ_FIELDS.forEach(([key,defaultLbl])=>{{
    const curLbl=customLabels[key]||defaultLbl;
    const grp=document.createElement('div');grp.className='fg';
    // Label row: custom label input + value input
    const lblRow=document.createElement('div');lblRow.style.cssText='display:flex;gap:4px;margin-bottom:3px;align-items:center';
    const lblInp=document.createElement('input');
    lblInp.id='lbl-'+key;lblInp.value=curLbl;
    lblInp.placeholder='Field label';
    lblInp.title='Edit field label';
    lblInp.style.cssText='font-size:9px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.4px;border:1px dashed var(--bd);border-radius:3px;padding:2px 5px;width:100%;background:transparent;outline:none';
    lblRow.appendChild(lblInp);
    const valInp=document.createElement('input');valInp.id='pf-'+key;valInp.value=proj[key]||'';
    valInp.placeholder='(leave empty to hide)';
    grp.appendChild(lblRow);grp.appendChild(valInp);
    grid.appendChild(grp);
  }});
  body.appendChild(grid);
  // Logo section
  const lt=document.createElement('div');lt.className='stitle';lt.textContent='Company Logos';body.appendChild(lt);
  const lg=document.createElement('div');lg.style.cssText='display:grid;grid-template-columns:1fr 1fr;gap:12px';
  for(const[k,lbl]of[['logo_left','Left Logo'],['logo_right','Right Logo']]){{
    const d=document.createElement('div');d.style.cssText='border:1px solid var(--bd);border-radius:6px;padding:10px;text-align:center;background:var(--bg)';
    const lbel=document.createElement('div');lbel.style.cssText='font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;margin-bottom:6px';lbel.textContent=lbl;d.appendChild(lbel);
    const img=document.createElement('img');img.id='lp-'+k;img.style.cssText='max-height:52px;max-width:100%;object-fit:contain;display:block;margin:0 auto 6px';d.appendChild(img);
    fetch('/api/logo/'+PID+'/'+k).then(r=>r.ok?r.blob():null).then(b=>{{if(b)img.src=URL.createObjectURL(b);}});
    const fi=document.createElement('input');fi.type='file';fi.accept='image/*';fi.style.cssText='width:100%;font-size:10px;margin-bottom:4px';
    fi.onchange=async e=>{{const f=e.target.files[0];if(!f)return;const b64=await new Promise(res=>{{const fr=new FileReader();fr.onload=e2=>res(e2.target.result);fr.readAsDataURL(f);}});img.src=b64;await apiFetch('/api/logo/'+PID+'/'+k,{{method:'POST',body:JSON.stringify({{data:b64.split(',')[1]}})}});toast('Logo saved','ok');}};
    d.appendChild(fi);
    const clr=document.createElement('button');clr.className='btn btn-sc btn-sm';clr.textContent='Remove';clr.style.fontSize='9px';
    clr.onclick=async()=>{{await apiFetch('/api/logo/'+PID+'/'+k,{{method:'POST',body:JSON.stringify({{data:''}})}});img.src='';toast('Removed','wa');}};
    d.appendChild(clr);lg.appendChild(d);
  }}
  body.appendChild(lg);
  openM('proj-modal');
}}

async function saveProject(){{
  const data={{}};
  const labels={{}};
  PROJ_FIELDS.forEach(([key,defaultLbl])=>{{
    const el=document.getElementById('pf-'+key);
    if(el)data[key]=el.value.trim();
    const lblEl=document.getElementById('lbl-'+key);
    if(lblEl&&lblEl.value.trim()&&lblEl.value.trim()!==defaultLbl)
      labels[key]=lblEl.value.trim();
  }});
  data._labels=labels;
  const r=await apiFetch('/api/project/'+PID,{{method:'POST',body:JSON.stringify(data)}});
  if(r===null)return;
  if(r.ok){{
    closeM('proj-modal');toast('✔ Project saved!','ok');
    setTimeout(()=>location.reload(),400);
  }}else toast('Save failed','er');
}}

// Doc Types
function addDocType(){{document.getElementById('dt-code').value='';document.getElementById('dt-name').value='';openM('dt-modal');}}
async function saveDocType(){{
  const code=document.getElementById('dt-code').value.trim().toUpperCase();
  const name=document.getElementById('dt-name').value.trim();
  if(!code||!name){{toast('Code and name required','er');return;}}
  await apiFetch('/api/doc_types/'+PID,{{method:'POST',body:JSON.stringify({{code,name}})}});
  closeM('dt-modal');await loadDTs();switchTab(code);toast('✔ Type added','ok');
}}

// Lists
async function openLists(){{
  await loadLists(true);
  const metaData=await apiFetch('/api/lists_meta/'+PID)||{{}};
  const META_LABELS={{approved:{{lbl:'Approved',bg:'#bbf7d0',fg:'#166534'}},rejected:{{lbl:'Rejected',bg:'#fce7f3',fg:'#831843'}},pending:{{lbl:'Pending',bg:'#fef9c3',fg:'#713f12'}},cancelled:{{lbl:'Cancelled',bg:'#f1f5f9',fg:'#94a3b8'}}}};
  const body=document.getElementById('lists-body');body.innerHTML='';
  for(const[ln,items]of Object.entries(state.lists)){{
    const isStatus=ln.toLowerCase().startsWith('status');
    const t=document.createElement('div');t.className='stitle';
    t.textContent=ln.charAt(0).toUpperCase()+ln.slice(1)+(isStatus?' 🏷':'');
    body.appendChild(t);
    const ul=document.createElement('ul');ul.className='slist';
    const metaItems=metaData[ln]||[];
    items.forEach(item=>{{
      const mi=metaItems.find(m=>m.value===item);
      const curMeta=mi?.meta||'pending';
      const ml=META_LABELS[curMeta]||META_LABELS.pending;
      const li=document.createElement('li');li.className='sitem';
      li.style.cssText='display:flex;align-items:center;gap:6px;padding:5px 8px';
      const nm=document.createElement('span');nm.style.flex='1';nm.textContent=item;
      li.appendChild(nm);
      if(isStatus){{
        const badge=document.createElement('span');
        badge.style.cssText=`background:${{ml.bg}};color:${{ml.fg}};font-size:9px;font-weight:700;padding:2px 8px;border-radius:10px;white-space:nowrap`;
        badge.textContent=ml.lbl;
        const sel=document.createElement('select');
        sel.style.cssText='font-size:10px;padding:2px 5px;border:1.5px solid #e2e8f0;border-radius:4px;background:#fff;cursor:pointer';
        ['approved','rejected','pending','cancelled'].forEach(m=>{{
          const o=document.createElement('option');o.value=m;
          o.textContent=META_LABELS[m].lbl;
          if(m===curMeta)o.selected=true;
          sel.appendChild(o);
        }});
        sel.onchange=async()=>{{
          const nm2=META_LABELS[sel.value];
          badge.style.background=nm2.bg;badge.style.color=nm2.fg;badge.textContent=nm2.lbl;
          await apiFetch('/api/lists_meta/'+PID,{{method:'POST',body:JSON.stringify({{list_name:ln,item_value:item,meta:sel.value}})}});
          toast('✔ Saved','ok');
        }};
        li.appendChild(badge);li.appendChild(sel);
      }}
      const rb=document.createElement('button');rb.textContent='Remove';
      rb.onclick=()=>rmItem(ln,item,rb);li.appendChild(rb);
      ul.appendChild(li);
    }});
    body.appendChild(ul);
    const ar=document.createElement('div');ar.className='addrow';
    const ni=document.createElement('input');ni.id='new-'+ln;ni.placeholder='New item...';
    const ab=document.createElement('button');ab.className='btn btn-ok btn-sm';ab.textContent='Add';ab.onclick=()=>addItem(ln);
    ar.appendChild(ni);ar.appendChild(ab);
    if(isStatus){{
      const hint=document.createElement('div');
      hint.style.cssText='font-size:10px;color:#94a3b8;margin-top:4px';
      hint.textContent='🏷 Set category for each status — affects Dashboard counts';
      ar.appendChild(hint);
    }}
    body.appendChild(ar);
  }}
  const nl=document.createElement('div');nl.className='stitle';nl.textContent='New List';
  const nar=document.createElement('div');nar.className='addrow';
  nar.innerHTML=`<input id="new-list" placeholder="List name"><button class="btn btn-pr btn-sm" onclick="mkList()">Create</button>`;
  body.appendChild(nl);body.appendChild(nar);
  openM('lists-modal');
}}

async function addItem(ln){{
  const inp=document.getElementById('new-'+ln);const val=inp?.value.trim();if(!val)return;
  await apiFetch('/api/lists/'+PID,{{method:'POST',body:JSON.stringify({{list_name:ln,item:val}})}});
  // Auto-assign pending meta for new status items
  if(ln.toLowerCase().startsWith('status')){{
    await apiFetch('/api/lists_meta/'+PID,{{method:'POST',body:JSON.stringify({{list_name:ln,item_value:val,meta:'pending'}})}});
  }}
  await loadLists(true);openLists();
}}
async function rmItem(ln,item,btn){{
  await apiFetch('/api/lists/'+PID,{{method:'DELETE',body:JSON.stringify({{list_name:ln,item}})}});
  btn.closest('li').remove();await loadLists(true);
}}
async function mkList(){{
  const name=document.getElementById('new-list')?.value.trim().toLowerCase().replace(/[\\s]+/g,'_');if(!name)return;
  await apiFetch('/api/lists/'+PID,{{method:'POST',body:JSON.stringify({{list_name:name,item:'Item 1'}})}});
  await loadLists(true);openLists();
}}

// Column drag-to-reorder
function initColDrag(ul){{
  let drag=null,over=null;
  ul.querySelectorAll('li').forEach(li=>{{
    li.draggable=true;
    li.ondragstart=e=>{{drag=li;li.style.opacity='.4';e.dataTransfer.effectAllowed='move';}};
    li.ondragend=()=>{{drag.style.opacity='';drag=null;ul.querySelectorAll('li').forEach(l=>l.style.background='');}};
    li.ondragover=e=>{{e.preventDefault();if(li===drag)return;over=li;
      ul.querySelectorAll('li').forEach(l=>l.style.background='');
      li.style.background='#e0f2fe';}};
    li.ondrop=e=>{{e.preventDefault();if(!drag||drag===li)return;
      const items=[...ul.querySelectorAll('li')];
      const di=items.indexOf(drag),oi=items.indexOf(li);
      if(di<oi)ul.insertBefore(drag,li.nextSibling);else ul.insertBefore(drag,li);
      saveColOrder();}};
  }});
}}
async function saveColOrder(){{
  const ids=[...document.querySelectorAll('#col-sortable li')].map(li=>parseInt(li.dataset.id));
  await apiFetch('/api/columns/reorder/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{order:ids}})}});
  toast('✔ Order saved','ok');
}}

// Columns
async function manageColumns(){{
  const cols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!cols)return;
  const body=document.getElementById('col-body');body.innerHTML='';
  const ul=document.createElement('ul');ul.id='col-sortable';ul.className='slist';ul.style.maxHeight='400px';
  cols.forEach(col=>{{
    const li=document.createElement('li');li.className='sitem';li.dataset.id=col.id;li.draggable=true;
    li.style.cssText='cursor:default;align-items:center;gap:7px';
    const grip=document.createElement('span');grip.textContent='⠿';
    grip.style.cssText='cursor:grab;color:var(--mu);font-size:16px;flex-shrink:0';
    grip.title='Drag to reorder';
    const chk=document.createElement('input');chk.type='checkbox';chk.checked=col.visible;
    chk.onchange=()=>toggleCol(col.id,chk.checked);
    const nm=document.createElement('span');nm.textContent=col.label;nm.style.cssText='flex:1;font-size:12px';
    const tp=document.createElement('span');tp.textContent=col.col_type;
    tp.style.cssText='font-size:9px;background:#e0e7ff;color:#3730a3;padding:1px 6px;border-radius:3px;flex-shrink:0';
    const ren_btn=document.createElement('button');ren_btn.textContent='✏';
    ren_btn.title='Rename';ren_btn.className='btn btn-sc btn-sm';
    ren_btn.style.cssText='flex-shrink:0;padding:2px 6px;font-size:11px';
    ren_btn.onclick=async()=>{{
      const newLbl=prompt('New label for column:',col.label);
      if(!newLbl||!newLbl.trim())return;
      const r=await apiFetch('/api/columns/rename/'+col.id,{{method:'POST',body:JSON.stringify({{label:newLbl.trim()}})}});
      if(r&&r.ok){{nm.textContent=newLbl.trim();toast('✔ Renamed','ok');}}
      else toast('Error','er');
    }};
    const db_btn=document.createElement('button');db_btn.textContent='🗑';
    db_btn.title='Delete';db_btn.className='btn btn-er btn-sm';
    db_btn.style.cssText='flex-shrink:0;padding:2px 6px;font-size:11px';
    db_btn.onclick=()=>deleteCol(col.id,col.col_key,db_btn);
    li.appendChild(grip);li.appendChild(chk);li.appendChild(nm);li.appendChild(tp);li.appendChild(ren_btn);li.appendChild(db_btn);
    ul.appendChild(li);
  }});
  initColDrag(ul);
  body.appendChild(ul);openM('col-modal');
}}
async function toggleCol(id,v){{await apiFetch('/api/columns/visibility/'+id,{{method:'POST',body:JSON.stringify({{visible:v}})}});}}
async function deleteCol(id,key,btn){{const warn=key==='docNo'?'⚠ WARNING: Deleting Document No. column will break the register! Are you sure?':'Delete this column?';if(!confirm(warn))return;await apiFetch('/api/columns/'+id,{{method:'DELETE'}});btn.closest('li').remove();}}

function onColType(v){{
  document.getElementById('cg-list').style.display=v==='dropdown'?'flex':'none';
  document.getElementById('cg-ds').style.display=v==='duration_calc'?'flex':'none';
  document.getElementById('cg-de').style.display=v==='duration_calc'?'flex':'none';
  document.getElementById('cg-info').style.display=v==='duration_calc'?'block':'none';
}}
async function openAddCol(){{
  await loadLists();
  const ls=document.getElementById('col-list');
  ls.innerHTML=Object.keys(state.lists).map(k=>`<option value="${{k}}">${{k}}</option>`).join('');
  const all=await apiFetch('/api/columns/'+PID+'/'+state.tab);
  const dates=(all||[]).filter(c=>['date','auto_date'].includes(c.col_type));
  const dopts=dates.map(c=>`<option value="${{c.col_key}}">${{c.label}}</option>`).join('')||'<option value="issuedDate">Issued Date</option>';
  document.getElementById('col-ds').innerHTML=dopts;
  document.getElementById('col-de').innerHTML=dopts;
  document.getElementById('col-name').value='';document.getElementById('col-type').value='text';onColType('text');
  openM('addcol-modal');
}}
async function saveAddCol(){{
  const name=document.getElementById('col-name').value.trim();
  const type=document.getElementById('col-type').value;
  const list=type==='dropdown'?document.getElementById('col-list').value:null;
  const ds=type==='duration_calc'?document.getElementById('col-ds').value:null;
  const de=type==='duration_calc'?document.getElementById('col-de').value:null;
  if(!name){{toast('Name required','er');return;}}
  if(type==='duration_calc'&&ds===de){{toast('Start and end must differ','er');return;}}
  await apiFetch('/api/columns/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{label:name,col_type:type,list_name:type==='duration_calc'?(ds+','+de):list}})}});
  closeM('addcol-modal');closeM('col-modal');await loadRecords();toast('✔ Column added','ok');
}}

// Admin Panel (same as dashboard)
async function openAdmin(){{
  const [users,projects]=await Promise.all([apiFetch('/api/users'),apiFetch('/api/projects')]);
  if(!users||!projects) return;
  const body=document.getElementById('admin-body'); body.innerHTML='';
  const utitle=document.createElement('div');utitle.className='stitle';utitle.textContent='👥 Users';body.appendChild(utitle);
  for(const u of users){{
    const row=document.createElement('div');row.className='urow';
    row.innerHTML=`<span style="flex:1;font-weight:600">👤 ${{u.username}}</span>
      <span class="badge" style="background:#fef3c7;color:#92400e">${{u.role.toUpperCase()}}</span>
      ${{u.username!=='admin'?`<button class="btn btn-sc btn-sm" onclick="chgPw('${{u.username}}')">🔑 PW</button>
        <button class="btn btn-er btn-sm" onclick="delUsr('${{u.username}}')">✕</button>`:
        '<span style="font-size:10px;color:var(--mu)">(protected)</span>'}}`;
    body.appendChild(row);
    if(u.role!=='superadmin'){{
      const ad=document.createElement('div');ad.style.cssText='padding:4px 10px 10px 32px;border-bottom:1px solid var(--bd);margin-bottom:4px';
      ad.innerHTML='<div style="font-size:10px;color:var(--mu);margin-bottom:4px">Project access:</div>';
      const assigned=await apiFetch('/api/users/'+u.username+'/projects').catch(()=>[]);
      const pl=document.createElement('div');pl.style.cssText='display:flex;flex-wrap:wrap;gap:4px';
      projects.forEach(p=>{{
        const isOn=assigned.includes(p.id);
        const btn=document.createElement('button');
        btn.style.cssText=`padding:3px 10px;border-radius:4px;border:2px solid;cursor:pointer;font-size:11px;font-weight:700;font-family:inherit;transition:all .15s;background:${{isOn?'#1a3a5c':'#f1f5f9'}};color:${{isOn?'#fff':'#94a3b8'}};border-color:${{isOn?'#f0a500':'#e2e8f0'}}`;
        btn.textContent=p.code;btn.title=p.name;
        if(isOn) btn.dataset.on='1';
        btn.onclick=async()=>{{
          const on=!!btn.dataset.on;
          await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:on?'unassign':'assign',username:u.username,project_id:p.id}})}});
          if(on){{delete btn.dataset.on;btn.style.background='#f1f5f9';btn.style.color='#94a3b8';btn.style.borderColor='#e2e8f0';}}
          else{{btn.dataset.on='1';btn.style.background='#1a3a5c';btn.style.color='#fff';btn.style.borderColor='#f0a500';}}
          toast((btn.dataset.on?'✔ Assigned: ':'Removed: ')+p.name,'ok');
        }};
        pl.appendChild(btn);}});
      ad.appendChild(pl);body.appendChild(ad);
    }}
  }}
  const at=document.createElement('div');at.className='stitle';at.textContent='➕ Add User';body.appendChild(at);
  const ar=document.createElement('div');
  ar.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:8px;align-items:end">
    <div class="fg"><label>Username</label><input id="nu-name" placeholder="username"></div>
    <div class="fg"><label>Role</label><select id="nu-role">
      <option value="editor">Editor</option><option value="viewer">Viewer</option>
      <option value="superadmin">Super Admin</option></select></div>
    <div class="fg"><label>Password</label><input id="nu-pw" type="password"></div>
    <button class="btn btn-pr btn-sm" style="margin-bottom:1px" onclick="addUsr()">Add</button></div>`;
  body.appendChild(ar);openM('admin-modal');
}}
async function addUsr(){{
  const name=document.getElementById('nu-name').value.trim().toLowerCase();
  const role=document.getElementById('nu-role').value;
  const pw=document.getElementById('nu-pw').value;
  if(!name||!pw){{toast('Username and password required','er');return;}}
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'add',username:name,role,password:pw}})}});
  if(r&&r.ok){{toast('✔ User added','ok');closeM('admin-modal');openAdmin();}}else toast((r&&r.error)||'Error','er');
}}
async function delUsr(u){{if(!confirm('Delete user: '+u+'?'))return;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'delete',username:u}})}});
  if(r&&r.ok){{toast('Deleted','wa');closeM('admin-modal');openAdmin();}}
}}
async function chgPw(u){{const pw=prompt('New password for '+u+':');if(!pw)return;
  await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'change_password',username:u,password:pw}})}});
  toast('✔ Password changed','ok');
}}

// Export/Import
function doExport(){{if(state.tab)window.location='/api/export/'+PID+'/'+state.tab;}}
function doExportPDF(){{if(state.tab)window.location='/api/export_pdf/'+PID+'/'+state.tab;}}
function doExportAllPDF(){{window.location='/api/export_pdf_all/'+PID;}}
function toggleExpDD(e){{e.stopPropagation();document.getElementById('exp-menu').classList.toggle('hidden');}}
function closeExpDD(){{document.getElementById('exp-menu').classList.add('hidden');}}
document.addEventListener('click',e=>{{if(!e.target.closest('#exp-dd'))closeExpDD();}});
function doPrint(){{
  const orig=document.title;
  document.title='DCR Print';
  window.print();
  document.title=orig;
}}
// Holidays Settings
async function openSettings(){{
  const r=await apiFetch('/api/settings/holidays');if(!r)return;
  let hols=[...(r.holidays||[])];
  const ov=document.createElement('div');ov.className='overlay';
  ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:9999;display:flex;align-items:center;justify-content:center';
  function buildUI(){{
    ov.innerHTML=`<div style="background:#fff;border-radius:12px;padding:24px;max-width:520px;width:95%;max-height:85vh;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.3)">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
        <div style="font-size:16px;font-weight:800;color:#1a3a5c">🗓 Public Holidays</div>
        <button id="hol-close" style="background:none;border:none;font-size:20px;cursor:pointer;color:#64748b">✕</button>
      </div>
      <p style="font-size:12px;color:#64748b;margin-bottom:12px">These dates are excluded from Duration and Expected Reply calculations (along with Fridays).</p>
      <div style="display:flex;gap:8px;margin-bottom:12px">
        <input type="date" id="hol-inp" style="flex:1;padding:6px 10px;border:1.5px solid #e2e8f0;border-radius:6px;font-family:inherit;font-size:12px">
        <button id="hol-add-btn" style="padding:6px 14px;background:#1a3a5c;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600">+ Add</button>
      </div>
      <div id="hol-list" style="flex:1;overflow-y:auto;border:1px solid #e2e8f0;border-radius:6px;padding:8px;min-height:180px;max-height:340px"></div>
      <div style="margin-top:12px;font-size:11px;color:#94a3b8">${{hols.length}} holidays total</div>
      <div style="display:flex;gap:8px;margin-top:12px;justify-content:flex-end">
        <button id="hol-cancel" style="padding:7px 16px;border:1.5px solid #e2e8f0;background:#fff;border-radius:6px;cursor:pointer;font-size:12px">Cancel</button>
        <button id="hol-save" style="padding:7px 16px;background:#16a34a;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600">💾 Save Holidays</button>
      </div></div>`;
    renderHL();
    ov.querySelector('#hol-close').onclick=()=>ov.remove();
    ov.querySelector('#hol-cancel').onclick=()=>ov.remove();
    ov.querySelector('#hol-add-btn').onclick=()=>{{
      const v=ov.querySelector('#hol-inp').value;
      if(!v)return;if(hols.includes(v)){{toast('Already added','wa');return;}}
      hols.push(v);renderHL();ov.querySelector('#hol-inp').value='';
    }};
    ov.querySelector('#hol-save').onclick=async()=>{{
      const btn=ov.querySelector('#hol-save');btn.disabled=true;btn.textContent='Saving...';
      const res=await apiFetch('/api/settings/holidays',{{method:'POST',body:JSON.stringify({{holidays:hols}})}});
      btn.disabled=false;btn.textContent='💾 Save Holidays';
      if(res&&res.ok){{toast('✔ '+res.count+' holidays saved','ok');ov.remove();}}
      else toast('Error','er');
    }};
  }}
  function renderHL(){{
    const ul=ov.querySelector('#hol-list');if(!ul)return;ul.innerHTML='';
    if(!hols.length){{ul.innerHTML='<div style="color:#94a3b8;text-align:center;padding:30px;font-size:12px">No holidays — click + Add to add dates</div>';return;}}
    [...hols].sort().forEach(d=>{{
      const row=document.createElement('div');
      row.style.cssText='display:flex;align-items:center;justify-content:space-between;padding:5px 8px;border-radius:4px;margin-bottom:3px;background:#f8fafc;font-size:12px';
      const dt=new Date(d+'T00:00:00');
      const fmt=dt.getDate().toString().padStart(2,'0')+'-'+(dt.getMonth()+1).toString().padStart(2,'0')+'-'+dt.getFullYear();
      const days=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
      const day=days[dt.getDay()];
      row.innerHTML=`<span style="color:#1e293b">${{fmt}} <span style="color:#94a3b8;font-size:10px">(${{day}})</span></span>
        <button data-d="${{d}}" style="background:none;border:none;color:#ef4444;cursor:pointer;font-size:15px;line-height:1;padding:2px 6px">✕</button>`;
      row.querySelector('button').onclick=e=>{{const dv=e.currentTarget.dataset.d;hols=hols.filter(h=>h!==dv);renderHL();}};
      ul.appendChild(row);
    }});
  }}
  buildUI();
  document.body.appendChild(ov);
}}

function doExportAll(){{window.location='/api/export_all/'+PID;}}
function openImport(){{openM('import-modal');}}
async function doImport(){{
  const file=document.getElementById('imp-file').files[0];if(!file)return;
  const ext=file.name.split('.').pop().toLowerCase();
  const btn=document.querySelector('#import-modal .btn-pr');
  btn.disabled=true;btn.textContent='⏳ Importing...';
  try{{
    const b64=await new Promise((res,rej)=>{{const fr=new FileReader();fr.onload=e=>res(e.target.result);fr.onerror=rej;fr.readAsDataURL(file);}});
    const r=await apiFetch('/api/import/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify({{file_b64:b64,ext}})}});
    if(r===null)return;
    closeM('import-modal');switchTab(state.tab);await loadDTs();toast('✔ Imported '+r.imported+' records','ok');
  }}catch(e){{toast('Error: '+e.message,'er');}}
  finally{{btn.disabled=false;btn.textContent='Import';}}
}}
async function doImportProject(){{
  const file=document.getElementById('imp-file').files[0];if(!file)return;
  const ext=file.name.split('.').pop().toLowerCase();
  if(!['xlsx','xls'].includes(ext)){{toast('Full workbook import supports Excel only','er');return;}}
  const btn=document.getElementById('imp-project-btn');
  btn.disabled=true;btn.textContent='⏳ Importing...';
  try{{
    const b64=await new Promise((res,rej)=>{{const fr=new FileReader();fr.onload=e=>res(e.target.result);fr.onerror=rej;fr.readAsDataURL(file);}});
    const r=await apiFetch('/api/import_project/'+PID,{{method:'POST',body:JSON.stringify({{file_b64:b64,ext}})}});
    if(r===null)return;
    closeM('import-modal');switchTab(state.tab);await loadDTs();
    toast('✔ Imported '+r.imported_total+' records from '+(r.matched_sheets||[]).length+' sheet(s)','ok');
  }}catch(e){{toast('Error: '+e.message,'er');}}
  finally{{btn.disabled=false;btn.textContent='Import Full Workbook';}}
}}
</script></body></html>"""


