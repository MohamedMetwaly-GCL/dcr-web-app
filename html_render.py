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
import os

import db
from flask import url_for
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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;600&display=swap');

/* ═══════════════════════════════════════════════════════
   GAS CHILL DCR — Enterprise Design Token System v2.0
   Brand: Navy #1a2f4e | Green #7ec832 | Teal #00b4a6
   Mode:  Light Data / Dark Shell
═══════════════════════════════════════════════════════ */
:root{
  /* ── Brand ── */
  --brand-navy:   #1a2f4e;
  --brand-navy-2: #243347;
  --brand-navy-3: #0f1d2e;
  --brand-green:  #7ec832;
  --brand-green-d:#5ea01e;
  --brand-teal:   #00b4a6;

  /* ── Shell (Topbar, Sidebar, Modal headers) ── */
  --shell-bg:     #1a2f4e;
  --shell-bg2:    #243347;
  --shell-border: rgba(255,255,255,.1);
  --shell-text:   #f1f5f9;
  --shell-muted:  rgba(255,255,255,.6);

  /* ── Data Surface (Tables, Cards, Forms) ── */
  --pr:#1a2f4e;
  --pl:#2563a8;
  --ac:#7ec832;
  --bg:#f4f6f9;
  --bg2:#eef1f6;
  --wh:#ffffff;
  --bd:#dde3ed;
  --tx:#1a2636;
  --mu:#6b7a94;

  /* ── Semantic ── */
  --ok:#16a34a;
  --er:#ef4444;
  --wa:#f59e0b;
  --pu:#7c3aed;
  --info:#00b4a6;

  /* ── KPI accent borders ── */
  --kpi-total:   #1a2f4e;
  --kpi-ok:      #16a34a;
  --kpi-warn:    #f59e0b;
  --kpi-danger:  #ef4444;
  --kpi-purple:  #7c3aed;
  --kpi-teal:    #00b4a6;

  /* ── Shape ── */
  --rd:6px;
  --rd-lg:10px;
  --shadow-sm:0 1px 4px rgba(0,0,0,.07);
  --shadow-md:0 4px 14px rgba(0,0,0,.10);
  --shadow-lg:0 24px 64px rgba(0,0,0,.20);
}

*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{font-family:'Inter','Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--tx);font-size:13px}
body.dark{--pr:#60a5fa;--pl:#93c5fd;--ac:#7ec832;--bg:#0f172a;--wh:#162132;--bd:#304257;
  --tx:#e2e8f0;--mu:#9fb0c6;--shell-bg:#0a1929;--shell-bg2:#0f2238;--bg2:#0f172a}
/* ── TOPBAR: Dark Navy Shell ── */
#topbar{
  background:var(--shell-bg);
  color:var(--shell-text);
  height:52px;
  display:flex;
  align-items:center;
  padding:0 18px;
  gap:10px;
  box-shadow:0 2px 12px rgba(0,0,0,.35);
  flex-shrink:0;
  position:relative;
  z-index:100;
  border-bottom:1px solid rgba(126,200,50,.2);
}
body.dark #topbar{background:var(--shell-bg,#0a1929);border-bottom-color:rgba(126,200,50,.15)}
#topbar .sp{flex:1}

/* Topbar brand line — subtle green accent at bottom */
#topbar::after{
  content:'';
  position:absolute;
  bottom:0;left:0;right:0;
  height:2px;
  background:linear-gradient(90deg,var(--brand-green) 0%,var(--brand-teal) 60%,transparent 100%);
  opacity:.7;
}

/* Topbar logo area */
.topbar-brand{
  display:flex;align-items:center;gap:8px;text-decoration:none;flex-shrink:0;
}
.topbar-brand-icon{
  width:32px;height:32px;border-radius:8px;
  background:linear-gradient(135deg,var(--brand-green),var(--brand-teal));
  display:flex;align-items:center;justify-content:center;
  font-size:16px;font-weight:900;color:#fff;flex-shrink:0;
  box-shadow:0 2px 8px rgba(126,200,50,.35);
}
.topbar-brand-text{
  font-size:12px;font-weight:700;color:var(--shell-text);
  letter-spacing:.3px;line-height:1.2;
}
.topbar-brand-sub{
  font-size:9px;color:var(--shell-muted);font-weight:400;letter-spacing:.5px;
}

/* Topbar buttons */
.tb-btn{
  background:rgba(255,255,255,.08);
  border:1px solid rgba(255,255,255,.15);
  color:rgba(255,255,255,.9);
  padding:6px 12px;
  border-radius:var(--rd);
  cursor:pointer;
  font-size:11.5px;
  font-weight:500;
  font-family:inherit;
  text-decoration:none;
  display:inline-flex;
  align-items:center;
  gap:5px;
  transition:all .15s;
  letter-spacing:.2px;
}
.tb-btn:hover{background:rgba(255,255,255,.18);border-color:rgba(255,255,255,.3);transform:translateY(-1px)}
.tb-btn:active{transform:translateY(0)}
.tb-btn.glow{
  background:rgba(126,200,50,.2);
  border-color:rgba(126,200,50,.5);
  color:#c7f078;
  font-weight:700;
}
.tb-btn.glow:hover{background:rgba(126,200,50,.3);border-color:rgba(126,200,50,.7)}

/* Topbar user chip */
.topbar-user{
  display:inline-flex;align-items:center;gap:7px;
  padding:4px 10px 4px 5px;
  border-radius:20px;
  background:rgba(255,255,255,.08);
  border:1px solid rgba(255,255,255,.12);
  color:rgba(255,255,255,.9);
  font-size:11.5px;
  font-weight:500;
}
.topbar-user-avatar{
  width:24px;height:24px;border-radius:50%;
  background:linear-gradient(135deg,var(--brand-green),var(--brand-teal));
  display:flex;align-items:center;justify-content:center;
  font-size:10px;font-weight:800;color:#fff;flex-shrink:0;
}
.topbar-title-short{display:none}
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:1000;
  display:flex;align-items:center;justify-content:center;backdrop-filter:blur(3px)}
.overlay.hidden{display:none!important}
.modal{background:#fff;border-radius:10px;box-shadow:0 24px 64px rgba(0,0,0,.3);
  width:92%;max-width:600px;max-height:90vh;display:flex;flex-direction:column;
  animation:mIn .18s ease}
body.dark .modal{background:var(--wh);color:var(--tx);box-shadow:0 24px 64px rgba(0,0,0,.48)}
@keyframes mIn{from{transform:translateY(-14px);opacity:0}}
.mhdr{
  background:linear-gradient(135deg,var(--brand-navy) 0%,var(--brand-navy-2) 100%);
  color:#fff;
  padding:13px 18px;
  font-weight:700;font-size:13px;
  display:flex;justify-content:space-between;align-items:center;
  flex-shrink:0;
  border-radius:10px 10px 0 0;
  border-bottom:2px solid rgba(126,200,50,.25);
  letter-spacing:.2px;
}
body.dark .mhdr{background:linear-gradient(135deg,#0a1929 0%,#0f2238 100%)}
.mbody{padding:16px 18px;overflow-y:auto;flex:1}
.mfoot{padding:10px 18px;border-top:1px solid var(--bd);display:flex;justify-content:flex-end;
  gap:8px;background:var(--bg);flex-shrink:0;border-radius:0 0 10px 10px}
body.dark .mfoot{background:#101a29}
.xbtn{background:none;border:none;color:#fff;font-size:20px;cursor:pointer;opacity:.7;line-height:1}
.xbtn:hover{opacity:1}
.fgrid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.fg{display:flex;flex-direction:column;gap:4px}
.fg.full{grid-column:1/-1}
.fg label{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px}
.fg input,.fg select,.fg textarea{width:100%;padding:7px 10px;border:1.5px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none;transition:border-color .2s;
  background:var(--wh);color:var(--tx)}
.fg input:focus,.fg select:focus,.fg textarea:focus{
  border-color:var(--brand-navy);
  box-shadow:0 0 0 3px rgba(26,47,78,.12);
  outline:none;
}
/* KPI counter animation */
@keyframes kpi-count-up{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
.kval.animated{animation:kpi-count-up .4s ease forwards}
/* Topbar separator line */
.tb-sep{width:1px;height:22px;background:rgba(255,255,255,.15);flex-shrink:0}
.btn{padding:7px 16px;border-radius:var(--rd);cursor:pointer;font-family:inherit;
  font-size:12px;font-weight:600;border:1px solid transparent;transition:all .15s}
.btn-pr{
  background:linear-gradient(135deg,var(--brand-navy) 0%,#1e3d6e 100%);
  color:#fff;
  box-shadow:0 2px 6px rgba(26,47,78,.25);
}
.btn-pr:hover{
  background:linear-gradient(135deg,#243347 0%,var(--brand-navy) 100%);
  box-shadow:0 4px 12px rgba(26,47,78,.35);
  transform:translateY(-1px);
}
.btn-sc{background:var(--bg);color:var(--tx);border-color:var(--bd)}.btn-sc:hover{background:var(--bd)}
.btn-ok{background:var(--ok);color:#fff}
.btn-er{background:var(--er);color:#fff}
.btn-sm{padding:4px 10px;font-size:11px}
.stitle{
  font-size:10.5px;font-weight:700;
  color:var(--brand-navy);
  text-transform:uppercase;
  letter-spacing:.6px;
  margin:12px 0 7px;
  padding-bottom:5px;
  padding-left:8px;
  border-bottom:2px solid var(--brand-green);
  border-left:3px solid var(--brand-navy);
  background:linear-gradient(90deg,rgba(126,200,50,.07) 0%,transparent 80%);
  border-radius:0 3px 3px 0;
}
#tab-overview > .stitle:first-child{display:none}
.record-form-shell{display:flex;flex-direction:column;gap:9px}
.form-section{background:linear-gradient(180deg,#fcfdff,#f7fafe);border:1px solid var(--bd);border-radius:12px;
  padding:9px 10px 8px;box-shadow:0 2px 8px rgba(15,23,42,.04)}
.form-section-header{display:flex;justify-content:space-between;gap:8px;align-items:flex-start;margin-bottom:7px}
.form-section-title{font-size:10px;font-weight:800;color:var(--pr);text-transform:uppercase;letter-spacing:.45px}
.form-section-sub{font-size:9px;color:var(--mu);line-height:1.3;margin-top:2px}
.form-section-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:6px 10px;align-items:start}
.form-section-grid .fg{min-width:0}
.form-section-grid .fg label{margin-bottom:2px}
.form-section-grid .fg.span-2{grid-column:span 2}
.form-section.compact-one{padding:7px 9px 6px;box-shadow:none}
.form-section.compact-one .form-section-header{margin-bottom:4px}
.form-section.compact-one .form-section-sub{display:none}
.form-section.compact-one .form-section-grid{gap:4px 8px}
.form-section-grid textarea{min-height:52px;resize:vertical;overflow:hidden}
.form-section-grid .fg.full textarea{min-height:58px}
body.dark .form-section{background:linear-gradient(180deg,#162132,#101a28);border-color:#2b3c4f;box-shadow:none}
body.dark .stitle,body.dark .form-section-title{color:#93c5fd}
body.dark .form-section-sub,body.dark .fg label{color:#cbd5e1}
body.dark .modal .mbody,body.dark .modal .fg label,body.dark .modal .sitem .nm{color:#e2e8f0}
body.dark .modal .stitle{border-bottom-color:#3b82f6}
body.dark .btn-sc{background:#101a29}
body.dark .btn-sc:hover{background:#223246}
body.dark .fg input,body.dark .fg select,body.dark .fg textarea{background:#0f1a29;border-color:#314558;color:#e2e8f0}
body.dark .fg input::placeholder,body.dark .fg textarea::placeholder{color:#7f93ac}
@media(max-width:768px){
  .modal{width:94vw;max-height:88vh;border-radius:10px}
  .mhdr{padding:10px 14px;font-size:12px}
  .mbody{padding:10px 12px}
  .mfoot{padding:8px 12px;gap:7px;position:sticky;bottom:0}
  .btn{padding:7px 12px;font-size:11px}
  .form-section{padding:7px 8px 6px;border-radius:10px}
  .record-form-shell{gap:6px}
  .form-section-header{margin-bottom:5px}
  .form-section-title{font-size:9px}
  .form-section-sub{font-size:9px;line-height:1.25;margin-top:2px}
  .form-section-grid{grid-template-columns:1fr;gap:8px 10px}
  .form-section-grid .fg.span-2,.form-section-grid .fg.full{grid-column:1/-1}
  .form-section-grid .fg label{font-size:9px;margin-bottom:2px}
  .fg input,.fg select,.fg textarea{padding:6px 8px;font-size:11px}
  .form-section-grid textarea{min-height:48px}
  .form-section-grid .fg.full textarea{min-height:54px}
  .stitle{font-size:9px;margin:7px 0 4px;padding-bottom:3px}
  .slist{max-height:150px;padding:3px;gap:2px}
  .sitem{gap:5px;padding:3px 5px;font-size:10px}
  .sitem button{padding:2px 6px;font-size:9px}
  .addrow{gap:5px;margin-top:5px}
  .addrow input{padding:5px 7px;font-size:10px}
}
.badge{display:inline-block;border-radius:10px;padding:2px 9px;font-size:10px;font-weight:700}
#toast{position:fixed;bottom:28px;right:18px;background:var(--pr);color:#fff;
  padding:10px 18px;border-radius:var(--rd);font-size:12px;z-index:9999;
  box-shadow:0 4px 16px rgba(0,0,0,.2);transform:translateY(80px);opacity:0;
  transition:all .3s;pointer-events:none;max-width:320px}
#toast.show{transform:none;opacity:1}
#toast.ok{background:#16a34a}#toast.er{background:#ef4444}#toast.wa{background:#f59e0b;color:#000}/* === BULLETPROOF DATA TABLE UX HOTFIX === */
/* 1. Remove Vertical Borders & Force Zebra Striping */
table, .dt-tbl, #regtbl { border-collapse: collapse !important; }
table td, table th, .dt-tbl td, .dt-tbl th, #regtbl td, #regtbl th { border-left: none !important; border-right: none !important; border-bottom: 1px solid #e2e8f0 !important; }
#regtbl thead tr:first-child th { border-bottom: none !important; }

/* Proper background inherit for sticky columns */
#regtbl tbody tr td { background-color: inherit !important; }
#regtbl tbody tr { background-color: #ffffff; }
#regtbl tbody tr:nth-child(even) { background-color: #f8fafc; }
#regtbl tbody tr:hover { background-color: #f1f5f9; filter: brightness(0.97); }

body.dark #regtbl tbody tr { background-color: var(--bg); }
body.dark #regtbl tbody tr:nth-child(even) { background-color: #1e293b; }
body.dark #regtbl tbody tr:hover { background-color: #334155; }

/* Overdue Styling for rows */
#regtbl tbody tr.ov td { background-color: inherit !important; }
#regtbl tbody tr.ov, #regtbl tbody tr.ov:nth-child(even) { background-color: #fef2f2 !important; }
#regtbl tbody tr.ov:hover { background-color: #ffe7e7 !important; }
body.dark #regtbl tbody tr.ov, body.dark #regtbl tbody tr.ov:nth-child(even) { background-color: #3a231f !important; }
body.dark #regtbl tbody tr.ov:hover { background-color: #4a2a24 !important; }

/* 2. Frozen Columns (Checkbox, Sr, Document No) */
#regtbl th:nth-child(1), #regtbl td:nth-child(1) { width: 32px !important; min-width: 32px !important; max-width: 32px !important; position: sticky !important; left: 0 !important; z-index: 5 !important; background-color: inherit !important; }
#regtbl th:nth-child(2), #regtbl td:nth-child(2) { width: 34px !important; min-width: 34px !important; max-width: 34px !important; position: sticky !important; left: 32px !important; z-index: 5 !important; background-color: inherit !important; }
#regtbl th.docno-cell, #regtbl td.docno-cell, #regtbl th:nth-child(3), #regtbl td:nth-child(3) { 
    position: sticky !important; left: 66px !important; z-index: 5 !important; background-color: inherit !important; 
    box-shadow: 2px 0 5px -2px rgba(0,0,0,0.1) !important;
}

/* Specific styling for the data cells (not header) */
#regtbl td.docno-cell, #regtbl td:nth-child(3) {
    font-family: 'Consolas', 'Courier New', monospace !important; font-weight: 600 !important; color: var(--brand-teal) !important;
}

/* Header z-index for Frozen Columns */
#regtbl th:nth-child(1), #regtbl th:nth-child(2), #regtbl th:nth-child(3), #regtbl th.docno-cell { z-index: 10 !important; background-color: var(--brand-navy) !important; }
body.dark #regtbl th:nth-child(1), body.dark #regtbl th:nth-child(2), body.dark #regtbl th:nth-child(3), body.dark #regtbl th.docno-cell { background-color: #0b131e !important; }

/* 3. Status Pill Badges & Resizing (Allow Wrapping) */
.sbadge { 
    border-radius: 9999px !important; 
    padding: 4px 12px !important; 
    display: inline-flex !important; 
    align-items: center !important; 
    justify-content: center !important;
    gap: 6px !important; 
    max-width: 100% !important; 
    white-space: normal !important; /* Allow text to wrap */
    word-break: break-word !important;
    text-align: left !important;
    line-height: 1.2 !important;
}
#regtbl td { max-width: 400px; /* Provides a flexible upper bound for resizer plugin */ }

/* === PHASE 3: TOP BAR & TABS REDESIGN === */
#projbar { background: #ffffff !important; border-bottom: 1px solid #e2e8f0 !important; box-shadow: 0 4px 12px rgba(0,0,0,0.03) !important; padding: 8px 16px !important; display: flex !important; flex-wrap: nowrap !important; align-items: flex-start !important; gap: 20px !important; overflow-x: auto !important; scrollbar-width: thin !important; scrollbar-color: rgba(0,0,0,0.2) transparent !important; }
#projbar::-webkit-scrollbar { height: 6px !important; display: block !important; }
#projbar::-webkit-scrollbar-thumb { background-color: rgba(0,0,0,0.2) !important; border-radius: 4px !important; }
body.dark #projbar { background: #111b2a !important; border-bottom: 1px solid #1e293b !important; box-shadow: 0 4px 12px rgba(0,0,0,0.2) !important; scrollbar-color: rgba(255,255,255,0.2) transparent !important; }
body.dark #projbar::-webkit-scrollbar-thumb { background-color: rgba(255,255,255,0.2) !important; }

/* Force Logo Size */
#projbar img { height: 45px !important; width: auto !important; max-width: none !important; object-fit: contain !important; margin: 0 !important; flex: 0 0 auto !important; }

/* Force Single Row for Main Container */
#projbar-main { display: flex !important; flex-wrap: nowrap !important; gap: 20px !important; align-items: flex-start !important; justify-content: flex-start !important; flex: 0 0 auto !important; max-width: none !important; width: auto !important; }

/* Destroy artificial boundaries so all items flow naturally in projbar-main */
#projbar-primary, #projbar-extra { display: contents !important; }

/* Force Flex Items to not shrink or wrap */
#projbar .pf { display: flex !important; flex-direction: column !important; gap: 2px !important; border: none !important; padding: 0 !important; flex: 0 0 auto !important; max-width: none !important; width: auto !important; }

#projbar .pf-lbl { font-size: 9.5px !important; font-weight: 700 !important; text-transform: uppercase !important; color: #64748b !important; letter-spacing: 0.3px !important; }
body.dark #projbar .pf-lbl { color: #94a3b8 !important; }

/* Prevent Text Truncation and Wrapping entirely */
#projbar .pf-val { 
    font-size: 13px !important; font-weight: 600 !important; color: #0f172a !important; 
    white-space: nowrap !important; 
    display: block !important;
}
body.dark #projbar .pf-val { color: #f8fafc !important; }
#projbar .pf.primary:last-child .pf-val { font-size: 14px !important; font-weight: 800 !important; color: var(--brand-teal) !important; }

.proj-edit-btn { background: #f1f5f9 !important; color: #334155 !important; border: 1px solid #cbd5e1 !important; padding: 6px 12px !important; border-radius: 9999px !important; font-size: 11.5px !important; font-weight: 600 !important; transition: all 0.2s ease !important; box-shadow: none !important; cursor: pointer !important; align-self: flex-start !important; }
.proj-edit-btn:hover { background: #e2e8f0 !important; color: #0f172a !important; }
body.dark .proj-edit-btn { background: #1e293b !important; color: #cbd5e1 !important; border-color: #334155 !important; }
body.dark .proj-edit-btn:hover { background: #334155 !important; color: #f8fafc !important; }

/* Tabs Bar */
#tabsbar { background: var(--brand-navy) !important; padding: 8px 16px !important; display: flex !important; flex-wrap: nowrap !important; gap: 8px !important; align-items: center !important; overflow-x: auto !important; scrollbar-width: thin !important; scrollbar-color: rgba(255,255,255,0.2) transparent !important; border-bottom: none !important; }
#tabsbar::-webkit-scrollbar { height: 4px !important; display: block !important; }
#tabsbar::-webkit-scrollbar-thumb { background-color: rgba(255,255,255,0.2) !important; border-radius: 4px !important; }
body.dark #tabsbar { background: #0a1420 !important; }

.tab-btn { flex-shrink: 0 !important; background: rgba(255,255,255,0.06) !important; border: 1px solid rgba(255,255,255,0.1) !important; color: rgba(255,255,255,0.7) !important; border-radius: 9999px !important; padding: 8px 16px !important; font-size: 13px !important; font-weight: 600 !important; display: inline-flex !important; align-items: center !important; gap: 10px !important; transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important; cursor: pointer !important; }
.tab-btn:hover { background: rgba(255,255,255,0.15) !important; color: #ffffff !important; transform: translateY(-1px) !important; }
.tab-btn.active { background: var(--brand-teal) !important; color: #ffffff !important; border-color: var(--brand-teal) !important; box-shadow: 0 4px 14px rgba(0, 180, 166, 0.35) !important; }

.tcnt { background: rgba(0,0,0,0.25) !important; color: #ffffff !important; padding: 2px 8px !important; border-radius: 9999px !important; font-size: 11px !important; font-weight: 700 !important; transition: all 0.2s !important; }
.tab-btn.active .tcnt { background: #ffffff !important; color: var(--brand-teal) !important; }

.tab-add { flex-shrink: 0 !important; background: transparent !important; border: 1px dashed rgba(255,255,255,0.3) !important; color: rgba(255,255,255,0.6) !important; border-radius: 50% !important; width: 32px !important; height: 32px !important; display: flex !important; align-items: center !important; justify-content: center !important; font-size: 18px !important; cursor: pointer !important; transition: all 0.2s !important; }
.tab-add:hover { background: rgba(255,255,255,0.1) !important; color: #ffffff !important; border-color: rgba(255,255,255,0.6) !important; transform: scale(1.05) !important; }
/* === PHASE 4: MODALS, FORMS & BUTTONS REDESIGN === */
.overlay { backdrop-filter: blur(5px) !important; background: rgba(15, 23, 42, 0.6) !important; }
body.dark .overlay { background: rgba(0, 0, 0, 0.75) !important; }

.modal { border-radius: 16px !important; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.35) !important; border: 1px solid rgba(255, 255, 255, 0.8) !important; overflow: hidden !important; }
body.dark .modal { box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7) !important; border-color: rgba(255, 255, 255, 0.1) !important; }

.mhdr { background: #ffffff !important; color: #0f172a !important; padding: 20px 24px !important; font-size: 16px !important; border-bottom: 1px solid #e2e8f0 !important; font-weight: 800 !important; }
body.dark .mhdr { background: #111b2a !important; color: #f8fafc !important; border-bottom-color: #1e293b !important; }

.mfoot { padding: 16px 24px !important; background: #f8fafc !important; border-top: 1px solid #e2e8f0 !important; gap: 12px !important; border-radius: 0 0 16px 16px !important; }
body.dark .mfoot { background: #0b131e !important; border-top-color: #1e293b !important; }

.xbtn { color: #64748b !important; background: rgba(0,0,0,0.04) !important; border-radius: 50% !important; width: 32px !important; height: 32px !important; display: flex !important; align-items: center !important; justify-content: center !important; font-size: 16px !important; transition: all 0.2s !important; }
.xbtn:hover { background: rgba(239, 68, 68, 0.1) !important; color: #ef4444 !important; transform: rotate(90deg) !important; opacity: 1 !important; }
body.dark .xbtn { color: #94a3b8 !important; background: rgba(255,255,255,0.05) !important; }

.form-section { background: #ffffff !important; border: 1px solid #e2e8f0 !important; border-radius: 12px !important; padding: 16px !important; box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important; }
body.dark .form-section { background: #111b2a !important; border-color: #1e293b !important; }

.fg label { font-size: 11px !important; font-weight: 700 !important; color: #475569 !important; text-transform: uppercase !important; letter-spacing: 0.5px !important; margin-bottom: 6px !important; }
body.dark .fg label { color: #94a3b8 !important; }

.fg input, .fg select, .fg textarea { padding: 10px 14px !important; font-size: 13px !important; border-radius: 8px !important; border: 1.5px solid #cbd5e1 !important; background: #f8fafc !important; color: #0f172a !important; font-weight: 500 !important; transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important; box-shadow: inset 0 1px 2px rgba(0,0,0,0.02) !important; }
.fg input:hover, .fg select:hover, .fg textarea:hover { background: #ffffff !important; border-color: #94a3b8 !important; }
.fg input:focus, .fg select:focus, .fg textarea:focus { background: #ffffff !important; border-color: var(--brand-teal) !important; box-shadow: 0 0 0 4px rgba(0, 180, 166, 0.15) !important; outline: none !important; }

body.dark .fg input, body.dark .fg select, body.dark .fg textarea { background: #0f172a !important; border-color: #334155 !important; color: #f8fafc !important; }
body.dark .fg input:hover, body.dark .fg select:hover, body.dark .fg textarea:hover { background: #1e293b !important; border-color: #475569 !important; }
body.dark .fg input:focus, body.dark .fg select:focus, body.dark .fg textarea:focus { background: #1e293b !important; border-color: var(--brand-teal) !important; box-shadow: 0 0 0 4px rgba(0, 180, 166, 0.2) !important; }

.btn { padding: 9px 20px !important; font-size: 13px !important; border-radius: 8px !important; font-weight: 600 !important; letter-spacing: 0.3px !important; transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important; display: inline-flex !important; align-items: center !important; justify-content: center !important; gap: 6px !important; }

.btn-pr { background: linear-gradient(135deg, var(--brand-teal) 0%, #00998c 100%) !important; color: #ffffff !important; border: none !important; box-shadow: 0 4px 12px rgba(0, 180, 166, 0.25) !important; }
.btn-pr:hover { transform: translateY(-2px) !important; box-shadow: 0 6px 16px rgba(0, 180, 166, 0.35) !important; background: linear-gradient(135deg, #00d2c2 0%, var(--brand-teal) 100%) !important; }
.btn-pr:active { transform: translateY(0) !important; box-shadow: 0 2px 8px rgba(0, 180, 166, 0.25) !important; }

.btn-sc { background: #ffffff !important; color: #475569 !important; border: 1px solid #cbd5e1 !important; box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important; }
.btn-sc:hover { background: #f8fafc !important; color: #0f172a !important; border-color: #94a3b8 !important; }
body.dark .btn-sc { background: #1e293b !important; color: #cbd5e1 !important; border-color: #334155 !important; }
body.dark .btn-sc:hover { background: #334155 !important; color: #f8fafc !important; border-color: #475569 !important; }

.stitle { font-size: 12px !important; font-weight: 800 !important; color: var(--brand-navy) !important; text-transform: uppercase !important; letter-spacing: 0.8px !important; margin: 16px 0 10px !important; padding-bottom: 6px !important; padding-left: 10px !important; border-bottom: 2px solid var(--brand-green) !important; border-left: 4px solid var(--brand-navy) !important; background: linear-gradient(90deg, rgba(126,200,50,0.07) 0%, transparent 100%) !important; border-radius: 0 4px 4px 0 !important; }
body.dark .stitle { color: #f8fafc !important; background: linear-gradient(90deg, rgba(126,200,50,0.1) 0%, transparent 100%) !important; border-left-color: var(--brand-teal) !important; }

/* === PHASE 4 POLISH: Labels & Inline fixes === */
#proj-modal input[id^="lbl-"] {
    background-color: transparent !important;
    border: none !important;
    box-shadow: none !important;
    font-weight: 700 !important;
    color: #64748b !important;
    font-size: 11px !important;
    letter-spacing: 0.5px !important;
    padding: 0 !important;
    margin-bottom: 8px !important;
}
body.dark #proj-modal input[id^="lbl-"] { color: #94a3b8 !important; }
#proj-modal input[id^="lbl-"]:focus { border-bottom: 1px dashed var(--brand-teal) !important; border-radius: 0 !important; outline: none !important; }

#proj-modal input[id^="pf-"] {
    border: 1px solid #e2e8f0 !important;
    border-radius: 6px !important;
}
body.dark #proj-modal input[id^="pf-"] { border-color: #334155 !important; }

/* Force Save Holidays to behave like Primary Button */
#hol-save { background: linear-gradient(135deg, var(--brand-teal) 0%, #00998c 100%) !important; color: #ffffff !important; border: none !important; box-shadow: 0 4px 12px rgba(0, 180, 166, 0.25) !important; padding: 9px 20px !important; font-size: 13px !important; border-radius: 8px !important; transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important; }
#hol-save:hover { transform: translateY(-2px) !important; box-shadow: 0 6px 16px rgba(0, 180, 166, 0.35) !important; background: linear-gradient(135deg, #00d2c2 0%, var(--brand-teal) 100%) !important; }

/* === PHASE 4 POLISH: Buttons & Scrollbars === */
.btn-er { background: transparent !important; color: #94a3b8 !important; border: 1px solid transparent !important; box-shadow: none !important; }
.btn-er:hover { background: #ef4444 !important; color: #ffffff !important; border-color: #ef4444 !important; box-shadow: 0 4px 12px rgba(239, 68, 68, 0.25) !important; transform: translateY(-1px) !important; }

.btn-ok { background: linear-gradient(135deg, var(--brand-teal) 0%, #00998c 100%) !important; color: #ffffff !important; border: none !important; box-shadow: 0 4px 12px rgba(0, 180, 166, 0.25) !important; }
.btn-ok:hover { transform: translateY(-2px) !important; box-shadow: 0 6px 16px rgba(0, 180, 166, 0.35) !important; background: linear-gradient(135deg, #00d2c2 0%, var(--brand-teal) 100%) !important; }
.btn-ok:active { transform: translateY(0) !important; box-shadow: 0 2px 8px rgba(0, 180, 166, 0.25) !important; }

.mbody::-webkit-scrollbar, .slist::-webkit-scrollbar, .tool-dd-menu::-webkit-scrollbar { width: 6px !important; height: 6px !important; display: block !important; }
.mbody::-webkit-scrollbar-thumb, .slist::-webkit-scrollbar-thumb, .tool-dd-menu::-webkit-scrollbar-thumb { background-color: rgba(0,0,0,0.2) !important; border-radius: 4px !important; }
body.dark .mbody::-webkit-scrollbar-thumb, body.dark .slist::-webkit-scrollbar-thumb, body.dark .tool-dd-menu::-webkit-scrollbar-thumb { background-color: rgba(255,255,255,0.2) !important; }
.mbody, .slist, .tool-dd-menu { scrollbar-width: thin !important; scrollbar-color: rgba(0,0,0,0.2) transparent !important; }
body.dark .mbody, body.dark .slist, body.dark .tool-dd-menu { scrollbar-color: rgba(255,255,255,0.2) transparent !important; }

/* === PHASE 5 UI POLISH === */
.sbadge::before { display: none !important; }
#projbar .pf { flex: 1 1 auto !important; min-width: 0 !important; }
#projbar .pf-val { white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important; max-width: none !important; }

/* Topbar Project Info (Moved from projbar) */
#topbar-proj-info:hover { background: rgba(255,255,255,0.08); }
#topbar-proj-info .pf { display: flex; flex-direction: column; gap: 2px; }
#topbar-proj-info .pf-lbl { font-size: 9px; color: rgba(255,255,255,0.6); text-transform: uppercase; letter-spacing: 0.5px; font-weight: 700; }
#topbar-proj-info .pf-val { font-size: 15px; color: #ffffff; font-weight: 800; white-space: nowrap; max-width: none !important; }
#topbar-proj-info .pf.primary .pf-val { font-size: 14px; color: var(--brand-teal); }
</style>"""

SHARED_JS = """
<div id="toast"></div>
<script>
function applyDarkMode(on){
  document.body.classList.toggle('dark',!!on);
  const btn=document.getElementById('darkBtn');
  if(btn)btn.textContent=on?'☀️':'🌙';
}
function toggleDark(){
  const on=!document.body.classList.contains('dark');
  applyDarkMode(on);
  localStorage.setItem('dcr_dark',on?'1':'');
}
if(localStorage.getItem('dcr_dark'))applyDarkMode(true);
function toast(msg,type=''){
  const t=document.getElementById('toast');
  t.textContent=msg;t.className='show '+(type||'');
  clearTimeout(t._t);t._t=setTimeout(()=>t.className='',3200);
}
function openM(id){document.getElementById(id).classList.remove('hidden')}
function closeM(id){document.getElementById(id).classList.add('hidden')}
function toggleProjectInfo(){
  const bar=document.getElementById('projbar');
  const btn=document.getElementById('projbar-toggle');
  if(!bar||!btn)return;
  const open=bar.classList.toggle('show-details');
  btn.textContent=open?'Hide Info':'Project Info';
}
async function apiFetch(url,opts={}){
  const r=await fetch(url,{credentials:'include',headers:{'Content-Type':'application/json'},cache:'no-store',...opts});
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
function escHtml(v){
  return String(v==null?'':v)
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'",'&#39;');
}
</script>"""


def render_login():
    static_dir = os.path.join(os.path.dirname(__file__), "static")
    logo_name = "logo-login.png"
    logo_file = os.path.join(static_dir, logo_name)
    import time
    logo_ver = str(int(os.path.getmtime(logo_file))) if os.path.exists(logo_file) else "1"
    logo_src = url_for("static", filename=logo_name) + f"?v={logo_ver}"
    return f"""<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Gas Chill DCR — Sign In</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{height:100%;font-family:'Inter','Segoe UI',Arial,sans-serif;overflow:hidden}}

/* ── Split Layout ── */
.login-wrap{{
  display:flex;height:100vh;width:100vw;
}}

/* ── LEFT PANEL: Dark Navy Brand Shell ── */
.login-left{{
  width:42%;
  background:linear-gradient(160deg,#0f1d2e 0%,#1a2f4e 45%,#0d2a1f 100%);
  display:flex;flex-direction:column;justify-content:center;align-items:center;
  padding:48px 40px;position:relative;overflow:hidden;flex-shrink:0;
}}

/* Dot grid pattern overlay */
.login-left::before{{
  content:'';position:absolute;inset:0;
  background-image:radial-gradient(rgba(126,200,50,.18) 1px, transparent 1px);
  background-size:28px 28px;pointer-events:none;
}}
/* Bottom gradient fade */
.login-left::after{{
  content:'';position:absolute;bottom:0;left:0;right:0;height:40%;
  background:linear-gradient(to top,rgba(15,29,46,.8),transparent);
  pointer-events:none;
}}

.ll-content{{position:relative;z-index:1;text-align:center;width:100%}}

.ll-logo{{
  display:flex;align-items:center;justify-content:center;
  margin:0 auto 32px;
}}
.ll-logo img{{
  width:140px;height:auto;
  object-fit:contain;
  filter:drop-shadow(0 0 10px rgba(255, 255, 255, 0.2));
}}

.ll-brand{{font-size:28px;font-weight:800;color:#fff;letter-spacing:-.3px;margin-bottom:4px}}
.ll-tagline{{font-size:12px;color:rgba(126,200,50,.85);font-weight:600;letter-spacing:2px;text-transform:uppercase;margin-bottom:24px}}

.valmore-logo {{ display: flex; align-items: center; justify-content: center; gap: 8px; margin-bottom: 36px; opacity: 0.8; }}
.valmore-icon {{ width: 18px; height: 18px; background: #fff; color: #1a2f4e; display: flex; align-items: center; justify-content: center; font-family: 'Inter', sans-serif; font-weight: 800; font-size: 14px; }}
.valmore-text {{ display: flex; align-items: baseline; }}
.v-main {{ font-family: Georgia, serif; font-size: 16px; color: #fff; letter-spacing: 0.5px; }}
.v-sub {{ font-family: 'Inter', sans-serif; font-size: 8px; color: #fff; letter-spacing: 2px; margin-left: 4px; font-weight: 600; text-transform: uppercase; }}

/* Divider */
.ll-divider{{
  width:48px;height:2px;margin:0 auto 28px;
  background:linear-gradient(90deg,transparent,#7ec832,transparent);
}}

.ll-desc{{font-size:13px;color:rgba(255,255,255,.55);line-height:1.7;max-width:260px;margin:0 auto 40px}}

/* Stats */
.ll-stats{{display:flex;justify-content:center;gap:28px}}
.ll-stat{{text-align:center}}
.ll-stat-num{{font-size:22px;font-weight:800;color:#7ec832;letter-spacing:-1px;font-variant-numeric:tabular-nums}}
.ll-stat-lbl{{font-size:9px;color:rgba(255,255,255,.45);font-weight:600;letter-spacing:.8px;text-transform:uppercase;margin-top:2px}}

/* Bottom brand text */
.ll-footer{{
  position:absolute;bottom:24px;left:0;right:0;z-index:1;
  text-align:center;font-size:10px;color:rgba(255,255,255,.25);letter-spacing:.5px;
}}

/* ── RIGHT PANEL: Clean Form ── */
.login-right{{
  flex:1;display:flex;align-items:center;justify-content:center;
  background:#eef1f6;padding:40px 32px;
  position:relative;
}}

/* Subtle top accent bar */
.login-right::before{{
  content:'';position:absolute;top:0;left:0;right:0;height:3px;
  background:linear-gradient(90deg,#1a2f4e 0%,#7ec832 50%,#00b4a6 100%);
}}

/* White card wrapping the form for depth */
.login-form-wrap{{
  width:100%;max-width:380px;
  background:#fff;
  border-radius:16px;
  padding:36px 32px 28px;
  box-shadow:0 8px 32px rgba(26,47,78,.12),0 1px 4px rgba(0,0,0,.06);
  border:1px solid rgba(221,227,237,.8);
}}

.lf-title{{
  font-size:22px;font-weight:800;color:#1a2f4e;margin-bottom:4px;letter-spacing:-.3px;
}}
.lf-sub{{font-size:12px;color:#6b7a94;margin-bottom:32px;}}

.fld{{margin-bottom:20px}}
.fld label{{
  display:block;font-size:10.5px;font-weight:700;color:#4a5568;
  text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px;
}}
/* Always show border — no invisible fields */
.fld input{{
  width:100%;padding:13px 15px;
  border:1.5px solid #d0d7e3;
  border-radius:10px;
  font-family:inherit;font-size:14px;
  outline:none;transition:all .2s;
  background:#fafbfd;color:#1a2636;
}}
.fld input:focus{{border-color:#1a2f4e;box-shadow:0 0 0 3px rgba(26,47,78,.1);background:#fff}}
.fld input::placeholder{{color:#a0aec0}}

.err{{
  background:#fef2f2;border:1px solid #fecaca;color:#b91c1c;
  padding:10px 14px;border-radius:8px;font-size:12px;margin-bottom:18px;
  align-items:center;gap:6px;
}}

.btn-login{{
  width:100%;padding:14px 16px;
  background:linear-gradient(135deg,#1a2f4e 0%,#243347 100%);
  color:#fff;border:none;border-radius:10px;
  font-family:inherit;font-size:14px;font-weight:700;
  letter-spacing:.3px;cursor:pointer;
  transition:all .2s;
  box-shadow:0 4px 14px rgba(26,47,78,.3);
  display:flex;align-items:center;justify-content:center;gap:8px;
}}
.btn-login:hover{{
  background:linear-gradient(135deg,#243347 0%,#1a2f4e 100%);
  transform:translateY(-2px);
  box-shadow:0 8px 24px rgba(26,47,78,.4);
}}
.btn-login:active{{transform:translateY(0)}}

.lf-footer{{
  margin-top:20px;text-align:center;font-size:11px;color:#94a3b8;
}}
.lf-footer a{{color:#7ec832;text-decoration:none;font-weight:600}}

/* ── Mobile: Stack vertically ── */
@media(max-width:700px){{
  .login-wrap{{flex-direction:column}}
  .login-left{{width:100%;min-height:220px;padding:32px 20px;}}
  .ll-logo{{margin-bottom:20px;}}
  .ll-logo img{{width:100px;height:auto;}}
  .ll-brand{{font-size:20px}}
  .ll-stats{{gap:20px;margin-top:0}}
  .ll-stat-num{{font-size:18px}}
  .login-right{{padding:28px 20px}}
  .ll-desc{{display:none}}
}}
</style></head><body>
<div class="login-wrap">

  <!-- LEFT: Dark Navy Brand Panel -->
  <div class="login-left">
    <div class="ll-content">
      <div class="ll-logo">
        <img src="{logo_src}" alt="Gas Chill">
      </div>
      <div class="ll-brand">GAS CHILL</div>
      <div class="ll-tagline">Energy. Redefined.</div>
      <div class="valmore-logo">
        <div class="valmore-icon"><span>V</span></div>
        <div class="valmore-text">
          <span class="v-main">Valmore</span>
          <span class="v-sub">HOLDING</span>
        </div>
      </div>
      <div class="ll-divider"></div>
      <div class="ll-desc">
        Enterprise Document Control System — Managing engineering submissions across all active projects with precision.
      </div>
      <div class="ll-stats">
        <div class="ll-stat">
          <div class="ll-stat-num">DCR</div>
          <div class="ll-stat-lbl">Platform</div>
        </div>
        <div class="ll-stat">
          <div class="ll-stat-num">v2</div>
          <div class="ll-stat-lbl">Version</div>
        </div>
        <div class="ll-stat">
          <div class="ll-stat-num">24/7</div>
          <div class="ll-stat-lbl">Access</div>
        </div>
      </div>
    </div>
    <div class="ll-footer">Gas Chill &nbsp;|&nbsp; Document Control System</div>
  </div>

  <!-- RIGHT: Clean Login Form -->
  <div class="login-right">
    <div class="login-form-wrap">
      <div class="lf-title">Welcome Back</div>
      <div class="lf-sub">Sign in to your workspace to continue</div>
      <div class="err" id="err" style="display:none"></div>
      <div class="fld">
        <label>Username</label>
        <input id="un" type="text" autofocus autocomplete="username" placeholder="Enter your username">
      </div>
      <div class="fld">
        <label>Password</label>
        <input id="pw" type="password" autocomplete="current-password" placeholder="Enter your password">
      </div>
      <button class="btn-login" onclick="login()">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/></svg>
        Sign In
      </button>
      <div class="lf-footer">
        Gas Chill Document Control System &nbsp;&mdash;&nbsp; Secure Access
      </div>
    </div>
  </div>

</div>
<script>
document.getElementById('pw').onkeydown=e=>{{if(e.key==='Enter')login();}};
document.getElementById('un').onkeydown=e=>{{if(e.key==='Enter')document.getElementById('pw').focus();}};
async function login(){{
  const un=document.getElementById('un').value.trim();
  const pw=document.getElementById('pw').value;
  const err=document.getElementById('err');
  err.style.display='none';
  if(!un||!pw){{err.textContent='Please enter username and password';err.style.display='flex';return;}}
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
body{{display:flex;flex-direction:column;min-height:100vh;overflow-x:hidden}}
.wrap{{max-width:1480px;margin:0 auto;padding:12px 12px;flex:1;width:100%}}

/* ── Dark Mode ── */
body.dark{{--bg:#0f172a;--wh:#162132;--tx:#e2e8f0;--mu:#9fb0c6;--bd:#304257;--pr:#60a5fa;--pl:#93c5fd}}
body.dark .kpi,body.dark .ccard,body.dark .pcard,body.dark .panel,
body.dark .psel-bar,body.dark .tbl-wrap{{background:#162132;color:#e2e8f0}}
body.dark .kval{{color:#93c5fd}}
body.dark .kpi.ok .kval{{color:#86efac}}
body.dark .kpi.wa .kval{{color:#fcd34d}}
body.dark .kpi.er .kval{{color:#fca5a5}}
body.dark .kpi.pu .kval{{color:#c4b5fd}}
body.dark .klbl,body.dark .ktrend,body.dark .psel-bar label{{color:#b8c8da}}
body.dark .clbl{{color:#93c5fd}}
body.dark .stitle{{color:#93c5fd;border-color:#93c5fd}}
body.dark .dt-tbl th{{background:#1e3a5f;color:#e2e8f0}}
body.dark .dt-tbl td{{border-color:#334155;color:#cbd5e1}}
body.dark .dt-tbl tr:hover td{{background:#142235}}
body.dark .dt-tbl .alt td{{background:#132031}}
body.dark #topbar{{background:#0f2640}}
body.dark .pchdr{{background:#1e3a5f}}
body.dark .pcbody{{color:#dbe7f3}}
body.dark select,body.dark input{{background:#1e293b;color:#e2e8f0;border-color:#334155}}
body.dark .tbtn{{background:#1e293b;color:#e2e8f0;border-color:#334155}}
body.dark .ov-row{{border-color:#334155;color:#cbd5e1}}
body.dark .prog{{background:#334155}}
body.dark .addcard{{background:linear-gradient(180deg,#132031,#0f1a29);border-color:#304257;color:#b8c8da}}
body.dark .addcard:hover{{background:linear-gradient(180deg,#17263a,#122033);border-color:#60a5fa;color:#dbe7f3}}
body.dark .pr-block,body.dark .ltr-summary-card,body.dark .ltr-parties-wrap,body.dark .ltr-party-card{{border-color:#304257!important}}
body.dark .pr-block,body.dark .ltr-summary-card,body.dark .ltr-parties-wrap{{background:#101a29!important}}
body.dark .ltr-party-card{{background:#162132!important}}

/* ── KPIs — Enterprise Edition ── */
.kpi-grid{{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(130px,1fr));
  gap:8px;
  margin-bottom:12px;
}}
.kpi{{
  background:var(--wh);
  border-radius:var(--rd-lg);
  padding:12px 14px 10px;
  min-height:88px;
  box-shadow:var(--shadow-sm);
  border-top:3px solid var(--kpi-total);
  cursor:default;
  transition:transform .2s,box-shadow .2s;
  display:flex;flex-direction:column;justify-content:center;
  position:relative;overflow:hidden;
}}
/* Subtle background glow on hover */
.kpi::before{{
  content:'';
  position:absolute;inset:0;
  background:linear-gradient(135deg,rgba(126,200,50,.04) 0%,transparent 60%);
  opacity:0;transition:opacity .2s;pointer-events:none;
}}
.kpi:hover{{transform:translateY(-3px);box-shadow:var(--shadow-md)}}
.kpi:hover::before{{opacity:1}}
.kpi.ok{{border-top-color:var(--kpi-ok)}}
.kpi.wa{{border-top-color:var(--kpi-warn)}}
.kpi.er{{border-top-color:var(--kpi-danger)}}
.kpi.pu{{border-top-color:var(--kpi-purple)}}
.kpi.tl{{border-top-color:var(--kpi-teal)}}
.kval{{
  font-size:30px;font-weight:800;
  color:var(--brand-navy);
  line-height:1;margin-bottom:5px;
  font-variant-numeric:tabular-nums;
  font-family:'IBM Plex Mono','Inter',monospace;
  letter-spacing:-1px;
}}
.kpi.ok .kval{{color:var(--kpi-ok)}}
.kpi.wa .kval{{color:var(--kpi-warn)}}
.kpi.er .kval{{color:var(--kpi-danger)}}
.kpi.pu .kval{{color:var(--kpi-purple)}}
.kpi.tl .kval{{color:var(--kpi-teal)}}
.klbl{{
  font-size:9.5px;color:var(--mu);font-weight:700;
  text-transform:uppercase;letter-spacing:.6px;
  margin-top:0;line-height:1.3;
}}
.ktrend{{
  font-size:10px;color:var(--mu);margin-top:4px;
  display:flex;align-items:center;gap:3px;
}}
/* Dark mode KPIs */
body.dark .kpi{{background:var(--wh);}}
body.dark .kval{{color:#93c5fd}}
body.dark .kpi.ok .kval{{color:#86efac}}
body.dark .kpi.wa .kval{{color:#fcd34d}}
body.dark .kpi.er .kval{{color:#fca5a5}}
body.dark .kpi.pu .kval{{color:#c4b5fd}}
body.dark .kpi.tl .kval{{color:#5eead4}}
body.dark .klbl{{color:#b8c8da}}

/* ── Toolbar ── */
.psel-bar{{display:flex;align-items:center;gap:4px;background:var(--wh);padding:5px 8px;
  border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:10px;flex-wrap:nowrap;overflow-x:auto}}
.psel-bar label{{font-size:11px;font-weight:700;color:var(--mu);white-space:nowrap}}
.psel-bar select,.psel-bar input{{padding:4px 8px;border:1.5px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:11.5px;outline:none}}
.tbtn{{padding:4px 8px;border:1.5px solid var(--bd);border-radius:var(--rd);
  background:var(--wh);cursor:pointer;font-size:10.5px;font-weight:600;font-family:inherit;
  transition:all .15s;color:var(--tx);white-space:nowrap}}
.tbtn:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.tbtn.active{{background:var(--pr);color:#fff;border-color:var(--pr)}}

/* ── Charts grid ── */
.charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:10px}}
.charts-grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:9px;margin-bottom:10px}}
.ccard{{background:var(--wh);border-radius:10px;padding:10px 11px;
  box-shadow:0 1px 4px rgba(0,0,0,.07);transition:box-shadow .15s,transform .15s}}
.ccard:hover{{box-shadow:0 4px 14px rgba(0,0,0,.08);transform:translateY(-1px)}}
.clbl{{font-size:10px;font-weight:700;color:var(--pr);text-transform:uppercase;
  letter-spacing:.5px;margin-bottom:6px;display:flex;align-items:center;gap:6px}}
canvas{{max-height:174px}}

/* ── Project cards ── */
.overview-stack{{display:grid;gap:10px;margin-bottom:10px}}
.overview-projects-panel{{padding:11px 12px;margin-bottom:0;min-width:0}}
.overview-projects-panel .panel-title{{margin-bottom:8px}}
.overview-pr-panel{{padding:11px 12px;margin-bottom:0;min-width:0}}
.overview-pr-panel .panel-title{{margin-bottom:8px}}
.overview-wide-panel{{padding:11px 12px;margin-bottom:10px}}
.overview-wide-panel .panel-title{{margin-bottom:8px}}
.pr-analytics-grid{{display:grid;grid-template-columns:minmax(170px,.85fr) minmax(240px,1.18fr) minmax(290px,1.5fr);gap:10px;align-items:stretch}}
.pr-analytics-grid > div{{min-width:0}}
.pr-block,.ltr-summary-card,.ltr-parties-wrap,.ltr-party-card,.charts-grid > *,.charts-grid-3 > *,#pr-panel,#ltr-panel{{min-width:0}}
.ltr-summary-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:9px}}
.ltr-parties-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px}}
.pgrid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:8px;margin-bottom:0}}
.pcard{{background:var(--wh);border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,.08);
  overflow:hidden;text-decoration:none;color:inherit;display:block;
  transition:transform .15s,box-shadow .15s;position:relative;min-height:132px}}
.pcard:hover{{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,.12)}}
.pchdr{{
  background:linear-gradient(135deg,var(--brand-navy) 0%,var(--brand-navy-2) 100%);
  padding:10px 12px;
  display:flex;align-items:center;justify-content:space-between;
  border-bottom:2px solid rgba(126,200,50,.3);
}}
.pcbody{{padding:7px 10px 8px;display:flex;flex-direction:column;gap:2px}}
.prow{{display:flex;justify-content:space-between;align-items:center;margin-bottom:1px}}
.prog{{height:5px;background:#eef1f7;border-radius:99px;overflow:hidden;margin-top:4px}}
.progf{{height:100%;border-radius:99px;transition:width .6s ease}}
.addcard{{background:linear-gradient(180deg,#fbfcfe,#f5f8fc);border-radius:10px;border:1.5px dashed #c9d4e2;min-height:132px;
  display:flex;align-items:center;justify-content:center;flex-direction:column;gap:6px;
  cursor:pointer;transition:all .15s;color:#71829a;font-size:12px}}
.addcard:hover{{border-color:var(--pr);color:var(--pr);background:linear-gradient(180deg,#ffffff,#f3f8ff)}}

/* ── Tables ── */
.tbl-wrap{{background:var(--wh);border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.07);
  overflow-x:auto;overflow-y:hidden;margin-bottom:12px;transition:box-shadow .15s;
  -webkit-overflow-scrolling:touch}}
.tbl-wrap:hover{{box-shadow:0 4px 14px rgba(0,0,0,.07)}}
.dt-tbl{{width:100%;min-width:760px;border-collapse:collapse;font-size:12px}}
.dt-tbl th{{
  background:var(--brand-navy);color:#fff;padding:12px 14px;
  text-align:left;font-weight:600;white-space:nowrap;
  letter-spacing:.3px;text-transform:uppercase;font-size:11px;
  position:sticky;top:0;z-index:2;
}}
.dt-tbl th:first-child{{border-radius:0;position:sticky;left:0;z-index:3}}
.dt-tbl td{{
  padding:10px 14px;
  border-bottom:1px solid #e2e8f0;
  font-variant-numeric:tabular-nums;
  transition:background .12s ease;
  line-height:1.5;
}}
.dt-tbl tr:nth-child(even) td{{background:#f8fafc}}
.dt-tbl tr:hover td{{background:#eef5ff}}
.dt-tbl td:first-child{{
  position:sticky;left:0;z-index:1;background:#fff;
  font-family:'Consolas','Courier New',monospace;
  font-weight:600;color:var(--brand-teal);
}}
.dt-tbl tr:nth-child(even) td:first-child{{background:#f8fafc}}
.dt-tbl tr:hover td:first-child{{background:#eef5ff}}
.dt-summary-row.warn td{{background:#fff4e8}}
.dt-summary-row.warn:hover td{{background:#feeccc}}
body.dark .dt-summary-row.warn td{{background:#35271e}}
body.dark .dt-summary-row.warn:hover td{{background:#453224}}
#overview-pane-discipline .dt-tbl{{table-layout:fixed}}
#overview-pane-discipline .dt-tbl th:first-child,
#overview-pane-discipline .dt-tbl td:first-child{{white-space:nowrap}}
#overview-pane-discipline .dt-tbl th:nth-child(2),
#overview-pane-discipline .dt-tbl td:nth-child(2){{white-space:nowrap}}
.disc-group-row td{{background:#f9fbfd;font-weight:600}}
.disc-group-row.alt td{{background:#f6f9fc}}
.disc-group-row:hover td{{background:#edf3f8}}
.disc-group-row.warn td{{background:#fff4e8}}
.disc-group-row.warn:hover td{{background:#feeccc}}
.disc-child-row{{display:none}}
.disc-child-row.open{{display:table-row}}
.disc-child-row td{{background:#f7fafe}}
.disc-child-row:hover td{{background:#eef3f8}}
.disc-child-spacer{{width:14px;padding:0 0 0 10px!important}}
.disc-child-label{{padding-left:18px!important;position:relative;font-weight:600;color:#2f4f64}}
.disc-child-label::before{{content:'';position:absolute;left:8px;top:50%;width:6px;height:6px;border-radius:999px;background:#c7d2de;transform:translateY(-50%)}}
.disc-badge{{display:inline-flex;align-items:center;justify-content:center;min-width:26px;padding:2px 8px;border-radius:999px;background:#e8eef6;color:#2f4f64;font-size:10px;font-weight:700}}
.disc-expander{{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border:1px solid #d7e0ea;border-radius:999px;background:#fff;color:#2f4f64;cursor:pointer;font-size:11px;transition:all .15s}}
.disc-expander:hover{{border-color:#2f4f64;background:#f5f8fc}}
.disc-expander.open{{background:#2f4f64;color:#fff;border-color:#2f4f64}}
.disc-meta{{font-size:10px;color:var(--mu);font-weight:600}}
.disc-num-cell{{text-align:center!important;font-weight:700;font-variant-numeric:tabular-nums}}
.disc-mobile-list{{display:none}}
.disc-mobile-card{{border:1px solid var(--bd);border-radius:8px;background:var(--wh);overflow:hidden;margin-bottom:8px}}
.disc-mobile-card.warn{{border-color:#f2c48d;background:#fff8ef}}
.disc-mobile-head{{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;padding:9px 10px;border-bottom:1px solid var(--bd)}}
.disc-mobile-title{{font-weight:800;color:var(--pr);font-size:12px}}
.disc-mobile-sub{{font-size:10px;color:var(--mu);margin-top:2px}}
.disc-mobile-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:6px 10px;padding:9px 10px}}
.disc-mobile-cell{{min-width:0}}
.disc-mobile-lbl{{font-size:9px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.24px}}
.disc-mobile-val{{font-size:13px;font-weight:800;margin-top:1px;font-variant-numeric:tabular-nums;color:var(--tx)}}
.disc-mobile-children{{display:none;border-top:1px solid var(--bd);background:rgba(255,255,255,.55)}}
.disc-mobile-card.open .disc-mobile-children{{display:block}}
.disc-mobile-child{{padding:8px 10px;border-bottom:1px solid var(--bd)}}
.disc-mobile-child:last-child{{border-bottom:none}}
#overview-pane-discipline .disc-group-row td,#overview-pane-discipline .disc-child-row td{{vertical-align:middle}}
#overview-pane-discipline .disc-group-row td:nth-child(3) > div{{display:flex;align-items:center;justify-content:flex-start;gap:8px;min-width:0}}
#overview-pane-discipline .disc-group-row td:nth-child(3),#overview-pane-discipline .disc-child-row td:nth-child(3){{min-width:170px}}
#overview-pane-discipline .disc-num-cell{{min-width:84px}}
body.dark .disc-group-row td{{background:#162132;color:#dbe7f3}}
body.dark .disc-group-row.alt td{{background:#132031}}
body.dark .disc-group-row:hover td{{background:#1a2a3d}}
body.dark .disc-group-row.warn td{{background:#3a2a1f}}
body.dark .disc-group-row.warn:hover td{{background:#4a3425}}
body.dark .disc-child-row td{{background:#101a29}}
body.dark .disc-child-row:hover td{{background:#162132}}
body.dark .disc-badge{{background:#223246;color:#dbe7f3}}
body.dark .disc-expander{{background:#101a29;border-color:#304257;color:#dbe7f3}}
body.dark .disc-expander.open{{background:#3b82f6;border-color:#3b82f6;color:#fff}}
body.dark .disc-meta,body.dark .disc-child-label{{color:#b8c8da}}
body.dark .disc-mobile-card{{background:#162132;border-color:#304257}}
body.dark .disc-mobile-card.warn{{background:#2d241a;border-color:#5b3b20}}
body.dark .disc-mobile-children{{background:#101a29;border-color:#304257}}
body.dark .disc-mobile-child{{border-color:#304257}}
body.dark .disc-mobile-sub,body.dark .disc-mobile-lbl{{color:#b8c8da}}
body.dark .disc-mobile-val{{color:#e2e8f0}}
@media(max-width:640px){{
  .disc-mobile-grid{{grid-template-columns:repeat(4,minmax(0,1fr));gap:5px;padding:8px}}
  .disc-mobile-cell{{display:flex;flex-direction:column;gap:1px;min-width:0;background:rgba(241,245,249,.62);border-radius:6px;padding:5px 6px}}
  .disc-mobile-lbl{{font-size:7.5px;line-height:1;white-space:nowrap}}
  .disc-mobile-val{{font-size:11px;line-height:1.05;margin-top:0;white-space:nowrap}}
  .disc-mobile-title{{font-size:11px}}
  .disc-mobile-sub{{font-size:9px}}
  body.dark .disc-mobile-cell{{background:#101a29}}
}}
.overview-table-switch{{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin:8px 0 8px}}
.overview-table-btn{{padding:6px 11px;border:1.5px solid var(--bd);border-radius:999px;background:var(--bg);
  color:var(--mu);cursor:pointer;font-size:10px;font-weight:700;letter-spacing:.35px;font-family:inherit;
  transition:all .15s}}
.overview-table-btn:hover{{border-color:#2F4F64;color:#2F4F64;background:#fff}}
.overview-table-btn.active{{background:#2F4F64;color:#fff;border-color:#2F4F64;box-shadow:0 2px 6px rgba(47,79,100,.16)}}
.overview-table-pane{{display:none}}
.overview-table-pane.active{{display:block}}
.overview-table-shell{{max-height:350px;overflow:auto;-webkit-overflow-scrolling:touch;padding-right:2px;scroll-behavior:smooth}}
.overview-table-shell .dt-tbl th{{position:sticky;top:0;z-index:2;box-shadow:0 1px 0 rgba(255,255,255,.08)}}
.overview-table-shell .tbl-wrap{{margin-bottom:0}}
.pr-toggle{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.pr-toggle:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.pr-items-row{{display:none}}
.pr-items-row.open{{display:table-row}}
.pr-items-row>td{{padding:0!important;border-bottom:1px solid #d9e2ec!important;border-right:none!important;max-width:none!important;background:transparent!important}}
.pr-items-panel{{display:block;width:100%;box-sizing:border-box;padding:16px;background:#f8fafc;border-top:1px solid #d9e2ec;border-bottom:1px solid #d9e2ec;position:relative;z-index:1}}
.pr-items-grid-wrap{{display:block;width:100%;max-width:100%;overflow-x:auto;padding-bottom:2px}}
.pr-items-title{{display:flex;align-items:center;gap:8px;margin-bottom:8px;font-size:11px;font-weight:800;color:var(--pr);text-transform:uppercase;letter-spacing:.45px}}
.pr-items-title:before{{content:"";width:4px;height:16px;border-radius:999px;background:var(--ac);display:inline-block}}
.pr-items-grid{{display:grid;grid-template-columns:minmax(520px,1fr) 120px 100px minmax(240px,.6fr);min-width:1100px;width:100%;border:1px solid #dbe3ed;border-radius:6px;overflow:hidden;background:#fff;font-size:11px}}
.pr-items-grid-head{{background:#173f66;color:#fff;font-weight:800;padding:10px 12px;border-right:1px solid rgba(255,255,255,.18)}}
.pr-items-grid-head:last-child{{border-right:none}}
.pr-items-cell{{padding:10px 12px;border-top:1px solid #edf2f7;border-right:1px solid #edf2f7;background:#fff;color:#10233a;white-space:normal;word-break:normal;line-height:1.35}}
.pr-items-cell:nth-child(4n){{border-right:none}}
.pr-items-cell.qty{{text-align:center;font-variant-numeric:tabular-nums}}
.pr-items-cell.unit{{white-space:nowrap}}
.pr-items-section{{grid-column:1/-1;background:#e8eef6;color:var(--pr);font-weight:800;text-transform:uppercase;letter-spacing:.35px;padding:9px 12px;border-top:1px solid #d7dee8}}
.pr-items-empty{{color:var(--mu);font-size:11px;padding:6px 0}}
body.dark .pr-items-row>td{{background:transparent!important;border-color:#304257!important}}
body.dark .pr-items-panel{{background:#101a29;border-color:#304257}}
body.dark .pr-items-grid{{background:#162132;border-color:#304257}}
body.dark .pr-items-grid-head{{background:#17314d;color:#dbeafe;border-color:#304257}}
body.dark .pr-items-cell{{background:#162132;color:#d7e1ec;border-color:#253648}}
body.dark .pr-items-section{{background:#1e3147;color:#dbeafe;border-color:#304257}}
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
.panel{{background:var(--wh);border-radius:10px;padding:12px 14px;
  box-shadow:0 1px 4px rgba(0,0,0,.07);margin-bottom:12px;transition:box-shadow .15s,transform .15s}}
.panel:hover{{box-shadow:0 4px 14px rgba(0,0,0,.08);transform:translateY(-1px)}}
.panel-title{{font-size:11px;font-weight:700;color:var(--pr);
  text-transform:uppercase;letter-spacing:.48px;margin-bottom:8px;
  border-bottom:1.5px solid var(--pr);padding-bottom:5px}}

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
@media(max-width:1200px){{
  .wrap{{padding:12px 10px}}
  .pgrid{{grid-template-columns:repeat(auto-fill,minmax(190px,1fr))}}
  .pr-analytics-grid{{grid-template-columns:minmax(160px,.9fr) minmax(220px,1.08fr) minmax(250px,1.32fr)}}
}}
@media(min-width:901px) and (max-width:1080px){{
  .pr-analytics-grid{{grid-template-columns:minmax(160px,.95fr) minmax(210px,1.08fr) minmax(240px,1.28fr)}}
}}
@media(max-width:900px){{
  .wrap{{padding:10px}}
  .overview-stack{{gap:10px}}
  .pr-analytics-grid{{grid-template-columns:1fr}}
  .charts-grid,.charts-grid-3{{grid-template-columns:1fr}}
  .kpi-grid{{grid-template-columns:repeat(3,minmax(0,1fr))}}
  .tab{{padding:7px 10px;font-size:11px}}
  canvas{{max-height:210px}}
  .overview-table-shell{{max-height:320px}}
}}
@media(max-width:640px){{
  .wrap{{padding:8px}}
  .kpi-grid{{grid-template-columns:repeat(2,minmax(0,1fr));gap:6px}}
  .kpi{{min-height:74px}}
  .kval{{font-size:22px}}
  .pgrid{{grid-template-columns:1fr}}
  .overview-projects-panel,.overview-pr-panel,.overview-wide-panel{{padding:10px 10px}}
  .psel-bar{{align-items:stretch;gap:6px}}
  .psel-bar select{{min-width:0;flex:1 1 160px}}
  .overview-table-shell{{max-height:300px}}
  .dt-tbl{{min-width:760px}}
  #overview-pane-discipline .tbl-wrap{{display:none}}
  #overview-pane-discipline .disc-mobile-list{{display:block}}
  .ltr-summary-grid,.ltr-parties-grid{{grid-template-columns:1fr}}
  .pr-block{{padding:10px!important}}
}}
@media(max-width:768px){{
  .dt-tbl{{min-width:600px}}
  .psel-bar{{flex-wrap:wrap;gap:6px}}
  .psel-bar select{{min-width:140px;flex:1}}
}}
@media(max-width:480px){{
  .psel-bar{{flex-direction:column;align-items:stretch}}
  .psel-bar .tbtn{{width:100%;justify-content:center}}
  .kpi-grid{{grid-template-columns:repeat(2,1fr);gap:5px}}
  .charts-grid{{gap:8px}}
  canvas{{max-height:190px}}
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
    <div style="display:flex; gap:4px; flex-wrap:nowrap; justify-content:flex-end">
      <button class="tbtn" onclick="showTab('overview')">📊 Overview</button>
      <button class="tbtn" onclick="showTab('analytics')">📈 Analytics</button>
      <button class="tbtn" onclick="showTab('overdue')">⚠️ Overdue</button>
      <button class="tbtn" onclick="showTab('executive')">📋 Executive</button>
      <button class="tbtn" onclick="showTab('daily-digest')">🌟 Daily Digest</button>
      {audit_btn}
    </div>
  </div>


    <!-- TAB: DAILY DIGEST -->
    <div id="tab-daily-digest" class="tab-pane">
      <div class="stitle" style="display:flex; justify-content:space-between; align-items:center;">
        <div>🌟 Daily Digest</div>
        <div>
          <input type="date" id="digest-date" style="padding:4px 8px; border:1px solid var(--bd); border-radius:4px; font-size:12px; background:var(--bg); color:var(--tx); letter-spacing:normal; text-transform:none; font-weight:normal;" onchange="loadDailyDigest()">
        </div>
      </div>
      <div id="daily-digest-content" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:20px;">
        <div style="color:var(--mu);font-size:14px;">Loading Daily Digest...</div>
      </div>
    </div>

    <!-- TAB: OVERVIEW -->
    <div id="tab-overview" class="tab-pane active">
      <div class="stitle">🗂 Projects</div>
      <div class="overview-stack">
        <div class="panel overview-projects-panel">
          <div class="panel-title">Projects</div>
          <div class="pgrid" id="pgrid"></div>
        </div>
        <div class="panel overview-pr-panel">
          <div class="panel-title">PR Analytics</div>
          <div id="pr-panel" style="display:grid;grid-template-columns:1fr;gap:10px">
            <div style="font-size:11px;color:var(--mu)">Loading...</div>
          </div>
        </div>
      </div>

      <div class="panel overview-wide-panel">
        <div class="panel-title">Letters Overview</div>
        <div id="ltr-panel" style="display:grid;grid-template-columns:1fr;gap:10px">
          <div style="font-size:11px;color:var(--mu)">Loading...</div>
        </div>
      </div>

    <div class="charts-grid">
      <div class="ccard"><div class="clbl">Documents by Project</div><canvas id="cProj"></canvas></div>
      <div class="ccard"><div class="clbl">Status Distribution</div><canvas id="cStatus"></canvas></div>
    </div>

    <div class="panel overview-wide-panel">
      <div class="panel-title">Operational Breakdown</div>
      <div class="overview-table-switch">
        <button type="button" id="ovtab-doc-types" class="overview-table-btn active" onclick="setOverviewTableTab('docTypes')">Document Types Summary</button>
        <button type="button" id="ovtab-discipline" class="overview-table-btn" onclick="setOverviewTableTab('discipline')">Discipline Breakdown</button>
      </div>
      <div class="overview-table-shell">
        <div id="overview-pane-docTypes" class="overview-table-pane active">
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
        </div>
        <div id="overview-pane-discipline" class="overview-table-pane">
          <div class="tbl-wrap">
            <table class="dt-tbl">
              <colgroup>
                <col style="width:120px">
                <col style="width:90px">
                <col>
                <col style="width:88px">
                <col style="width:88px">
                <col style="width:88px">
                <col style="width:88px">
                <col style="width:88px">
                <col style="width:52px">
              </colgroup>
              <thead><tr>
                <th>Project</th><th>Doc Type</th><th>Disciplines</th>
                <th style="text-align:center">Total</th>
                <th style="text-align:center">Approved</th>
                <th style="text-align:center">Pending</th>
                <th style="text-align:center">Rejected</th>
                <th style="text-align:center">Overdue</th>
                <th style="text-align:center;width:42px">View</th>
              </tr></thead>
              <tbody id="disc-tbody"></tbody>
            </table>
            <div id="disc-empty" style="text-align:center;padding:24px;color:var(--mu);display:none">No data</div>
          </div>
          <div id="disc-mobile-list" class="disc-mobile-list"></div>
        </div>
      </div>
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

// ── Tab switching ──────────────────────────────────────────
function showTab(name){{
  document.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));
  const pane=document.getElementById('tab-'+name);
  if(pane)pane.classList.add('active');
  if(name==='analytics'){{analyticsLoaded=false;setTimeout(()=>loadAnalytics(),50);}}
  if(name==='overdue'){{overdueLoaded=false;loadOverdue();}}
  if(name==='executive'){{EXEC_DATA=null;loadExecutive();}}
  if(name==='daily-digest'){{loadDailyDigest();}}
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
  if(activeTab?.id==='tab-daily-digest'){{loadDailyDigest();}}
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
  const lettersScope=pid?STATS.filter(s=>s.id===pid):STATS;
  updateKPIs(d);renderCards(d,pid);renderLettersOverview(lettersScope);renderCharts(d);renderDTTable(d);renderDiscTable(d);
  ensureOverviewTableControls();setOverviewTableTab(window._overviewTableTab||'docTypes');
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
    ? (data.top_projects||[]).map((p,idx)=>`<div style="display:flex;justify-content:space-between;gap:10px;padding:${{idx===0?'7px 0 8px':'6px 0'}};border-bottom:1px solid var(--bd);align-items:center">
        <div style="min-width:0">
          <div style="font-size:${{idx===0?'11.5px':'11px'}};font-weight:${{idx===0?'800':'600'}};color:${{idx===0?'var(--pr)':'var(--tx)'}};white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{p.project_name}}</div>
          ${{idx===0?'<div style="font-size:9px;color:var(--mu);margin-top:2px;text-transform:uppercase;letter-spacing:.28px">Most active project</div>':''}}
        </div>
        <span style="font-size:${{idx===0?'12px':'11px'}};font-weight:800;color:var(--pr);white-space:nowrap">${{p.pr_count}}</span>
      </div>`).join('')
    : `<div style="font-size:11px;color:var(--mu)">No project data</div>`;
  el.innerHTML=`<div class="pr-analytics-grid">
    <div class="pr-block pr-total" style="background:var(--bg);border-radius:8px;padding:10px 12px;border:1px solid #e3eaf2;display:flex;flex-direction:column;justify-content:center">
      <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Total PRs</div>
      <div style="font-size:24px;font-weight:800;color:var(--pr);line-height:1.05;margin-top:4px">${{data.total_pr_records||0}}</div>
      <div style="font-size:9px;color:var(--mu);margin-top:6px;text-transform:uppercase;letter-spacing:.28px">Current dashboard scope</div>
      <div style="font-size:10px;color:var(--mu);margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">Top project: ${{data.top_project_name||'No project data'}}</div>
    </div>
    <div class="pr-block pr-projects" style="background:var(--bg);border-radius:8px;padding:10px 12px;border:1px solid #e3eaf2;display:flex;flex-direction:column;min-height:156px">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:8px">
        <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Most Active Projects</div>
        <div style="font-size:10px;color:var(--mu);white-space:nowrap">${{(data.top_projects||[]).length}} shown</div>
      </div>
      <div style="flex:1;overflow:auto;padding-right:2px">
        ${{topProjects}}
      </div>
    </div>
    <div class="pr-block pr-trades" style="background:var(--bg);border-radius:8px;padding:10px 12px;border:1px solid #e3eaf2;display:flex;flex-direction:column;min-height:156px">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:8px">
        <div>
          <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.4px">Most Active Trades</div>
          <div style="font-size:15px;font-weight:800;color:var(--pr);line-height:1.1;margin-top:5px">${{data.top_trade_name||'No trade data'}}</div>
          <div style="font-size:9px;color:var(--mu);margin-top:2px;text-transform:uppercase;letter-spacing:.28px">Most active trade</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:15px;font-weight:800;color:var(--pr);line-height:1">${{data.top_trade_count||0}}</div>
          <div style="font-size:10px;color:var(--mu);margin-top:3px">${{data.trade_count_total||0}} trades</div>
        </div>
      </div>
      <div style="flex:1;min-height:150px;max-height:160px"><canvas id="pr-trades-chart"></canvas></div>
    </div>
  </div>`;
  const tradeCanvas=document.getElementById('pr-trades-chart');
  if(prTradeChart){{prTradeChart.destroy();prTradeChart=null;}}
  if(!tradeCanvas)return;
  const trades=data.top_trades||[];
  const dark=document.body.classList.contains('dark');
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
        x:{{beginAtZero:true,grid:{{color:dark?'rgba(159,176,198,.16)':'rgba(0,0,0,.06)'}},ticks:{{precision:0,font:{{size:10}},color:dark?'#cbd5e1':'#54657d'}}}},
        y:{{grid:{{display:false}},ticks:{{font:{{size:10}},color:dark?'#dbe7f3':'#334155'}}}}
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
      ${{(ROLE==='superadmin'||ROLE==='admin')&&!pid?`<div style="position:absolute;right:6px;bottom:6px;display:flex;gap:4px">
        <button class="del-pbtn" title="Move project up" onclick="event.preventDefault();event.stopPropagation();moveProject('${{p.id}}',-1)">Up</button>
        <button class="del-pbtn" title="Move project down" onclick="event.preventDefault();event.stopPropagation();moveProject('${{p.id}}',1)">Down</button>
      </div>`:''}}
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
async function moveProject(id,delta){{
  const idx=STATS.findIndex(p=>p.id===id);
  const next=idx+delta;
  if(idx<0||next<0||next>=STATS.length)return;
  [STATS[idx],STATS[next]]=[STATS[next],STATS[idx]];
  const order=STATS.map(p=>p.id);
  await apiFetch('/api/projects/reorder',{{method:'POST',body:JSON.stringify({{order}})}});
  document.getElementById('proj-sel').innerHTML=
    '<option value="">All Projects</option>'+
    STATS.map(p=>`<option value="${{p.id}}" ${{p.id===_currentPid?'selected':''}}>${{p.name}} (${{p.code}})</option>`).join('');
  renderAll(_currentPid,_currentDisc);
  toast('Project order updated','ok');
}}

function renderCharts(d){{
  const dark=document.body.classList.contains('dark');
  const gridColor=dark?'rgba(159,176,198,.16)':'rgba(0,0,0,.05)';
  const tickColor=dark?'#cbd5e1':'#54657d';
  const legendColor=dark?'#dbe7f3':'#334155';
  if(pChart)pChart.destroy();
  pChart=new Chart(document.getElementById('cProj'),{{type:'bar',
    data:{{labels:d.map(s=>s.code),datasets:[
      {{label:'Approved',data:d.map(s=>s.approved),backgroundColor:'#16a34a',borderRadius:4}},
      {{label:'Pending', data:d.map(s=>s.pending), backgroundColor:'#f59e0b',borderRadius:4}},
      {{label:'Overdue', data:d.map(s=>s.overdue), backgroundColor:'#ef4444',borderRadius:4}}]}},
    options:{{responsive:true,plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}},color:legendColor}}}}}},
      scales:{{y:{{beginAtZero:true,grid:{{color:gridColor}},ticks:{{color:tickColor}}}},
        x:{{grid:{{display:false}},ticks:{{color:tickColor}}}}}}}}}});
  const t=d.reduce((s,p)=>s+p.total,0),ap=d.reduce((s,p)=>s+p.approved,0),
    pe=d.reduce((s,p)=>s+p.pending,0),ov=d.reduce((s,p)=>s+p.overdue,0),
    rj=d.reduce((s,p)=>s+(p.rejected||0),0);
  if(sChart)sChart.destroy();
  sChart=new Chart(document.getElementById('cStatus'),{{type:'doughnut',
    data:{{labels:['Approved','Pending','Rejected','Overdue'],
      datasets:[{{data:[ap,pe,rj,ov],
        backgroundColor:['#16a34a','#f59e0b','#7c3aed','#ef4444'],
        borderWidth:3,borderColor:dark?'#162132':'#fff',hoverOffset:6}}]}},
    options:{{responsive:true,cutout:'65%',
      plugins:{{legend:{{position:'bottom',labels:{{boxWidth:10,font:{{size:10}},color:legendColor}}}},
        tooltip:{{callbacks:{{label:ctx=>` ${{ctx.label}}: ${{ctx.raw}} (${{t?Math.round(ctx.raw/t*100):0}}%)`}}}}}}}}}});
}}

// ── DT Table ──────────────────────────────────────────────
function renderLettersOverview(d){{
  const el=document.getElementById('ltr-panel');
  if(!el)return;
  try{{
    const safeHtml=v=>String(v??'').replace(/[&<>"']/g,m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]));
    const partyMap={{}};
    const stats=(Array.isArray(d)?d:[]).reduce((acc,p)=>{{
      const l=(p&&typeof p==='object'&&p.ltr&&typeof p.ltr==='object')?p.ltr:{{}};
      acc.total+=Number(l.total||0);
      acc.sent+=Number(l.sent||0);
      acc.received+=Number(l.received||0);
      if(Number(l.total||0)>0)acc.project_count+=1;
      (Array.isArray(l.party_stats)?l.party_stats:[]).forEach(row=>{{
        const party=String((row&&row.party)||'').trim();
        if(!party)return;
        if(!partyMap[party])partyMap[party]={{party,total:0,sent:0,received:0}};
        partyMap[party].total+=Number((row&&row.total)||0);
        partyMap[party].sent+=Number((row&&row.sent)||0);
        partyMap[party].received+=Number((row&&row.received)||0);
      }});
      return acc;
    }},{{total:0,sent:0,received:0,project_count:0}});
    if(!stats.total){{
      el.innerHTML='<div style="font-size:11px;color:var(--mu)">No letter records in the current project scope.</div>';
      return;
    }}
    const topParties=Object.values(partyMap)
      .sort((a,b)=>(b.total-a.total)||(b.sent-a.sent)||(b.received-a.received)||a.party.localeCompare(b.party))
      .slice(0,6);
    const cards=[
      ['Total Letters',stats.total,'#2F4F64','Across '+stats.project_count+' project'+(stats.project_count===1?'':'s')],
      ['Sent',stats.sent,'#2563a8','Outgoing correspondence'],
      ['Received',stats.received,'#16a34a','Incoming correspondence'],
    ];
    const summaryHtml=`<div class="ltr-summary-grid">
      ${{cards.map(([label,value,color,hint])=>`<div class="ltr-summary-card" style="background:var(--bg);border-radius:8px;padding:11px 13px;border:1px solid #e3eaf2;border-left:4px solid ${{color}};box-shadow:0 1px 2px rgba(0,0,0,.03)">
        <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.45px">${{label}}</div>
        <div style="font-size:26px;font-weight:800;color:${{color}};line-height:1.08;margin-top:5px">${{value}}</div>
        <div style="font-size:10px;color:var(--mu);margin-top:3px">${{hint}}</div>
      </div>`).join('')}}
    </div>`;
    const partiesHtml=topParties.length
      ? `<div class="ltr-parties-wrap" style="background:var(--bg);border-radius:8px;padding:9px 11px;border:1px solid #e3eaf2">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:8px">
            <div style="font-size:10px;font-weight:700;color:var(--tx);text-transform:uppercase;letter-spacing:.45px">Most Active Parties</div>
            <div style="font-size:10px;color:var(--mu);font-weight:700">Top ${{topParties.length}} Parties</div>
          </div>
          <div class="ltr-parties-grid">
            ${{topParties.map((row, idx)=>{{
              const pct=stats.total?Math.round((Number(row.total||0)/stats.total)*100):0;
              const accent=idx<3?'rgba(47,79,100,.12)':'rgba(221,227,237,.95)';
              const bg=idx<3?'linear-gradient(180deg,rgba(255,255,255,.98),rgba(245,249,255,.94))':'rgba(255,255,255,.82)';
              const shadow=idx<3?'0 3px 10px rgba(47,79,100,.06)':'0 1px 3px rgba(15,23,42,.04)';
              return `<div class="ltr-party-card" style="border:1px solid ${{accent}};border-radius:9px;background:${{bg}};padding:9px 10px;box-shadow:${{shadow}};transition:background .15s,border-color .15s,transform .15s,box-shadow .15s"
                onmouseenter="this.style.background='linear-gradient(180deg,rgba(255,255,255,1),rgba(244,248,253,.98))';this.style.borderColor='#c7d2de';this.style.transform='translateY(-1px)';this.style.boxShadow='0 4px 12px rgba(47,79,100,.08)'"
                onmouseleave="this.style.background='${{bg}}';this.style.borderColor='${{accent}}';this.style.transform='translateY(0)';this.style.boxShadow='${{shadow}}'">
                <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start">
                  <div style="min-width:0">
                    <div style="font-size:12px;font-weight:800;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{safeHtml(row.party)}}</div>
                    <div style="font-size:9px;color:var(--mu);margin-top:2px">Share of letters: ${{pct}}%</div>
                  </div>
                  <div style="text-align:right;flex-shrink:0">
                    <div style="font-size:14px;font-weight:800;color:#2F4F64;line-height:1">${{row.total}}</div>
                    <div style="font-size:9px;color:var(--mu);margin-top:2px">Total</div>
                  </div>
                </div>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:8px">
                  <div style="background:#eef4fb;border-radius:7px;padding:6px 8px;border:1px solid #dbe8f7">
                    <div style="font-size:10px;font-weight:800;color:#2563a8">Sent: ${{row.sent}}</div>
                    <div style="font-size:9px;color:#2563a8;letter-spacing:.2px;margin-top:2px">Outgoing</div>
                  </div>
                  <div style="background:#edf9f0;border-radius:7px;padding:6px 8px;border:1px solid #d7efdc">
                    <div style="font-size:10px;font-weight:800;color:#16a34a">Received: ${{row.received}}</div>
                    <div style="font-size:9px;color:#16a34a;letter-spacing:.2px;margin-top:2px">Incoming</div>
                  </div>
                </div>
              </div>`;
            }}).join('')}}
          </div>
        </div>`
      : '<div style="font-size:11px;color:var(--mu)">No party activity available in the current project scope.</div>';
    el.innerHTML=`<div style="display:grid;gap:10px">
      ${{summaryHtml}}
      ${{partiesHtml}}
    </div>`;
  }}catch(err){{
    console.error('renderLettersOverview failed', err);
    el.innerHTML='<div style="font-size:11px;color:var(--mu)">Letters overview is temporarily unavailable.</div>';
  }}
}}

function ensureOverviewTableControls(){{
  const overview=document.getElementById('tab-overview');
  if(!overview)return;
  const shell=overview.querySelector('.overview-table-shell');
  const docPane=document.getElementById('overview-pane-docTypes');
  const discPane=document.getElementById('overview-pane-discipline');
  if(!shell||!docPane||!discPane)return;
}}

function setOverviewTableTab(tab){{
  window._overviewTableTab=(tab==='discipline')?'discipline':'docTypes';
  const docPane=document.getElementById('overview-pane-docTypes');
  const discPane=document.getElementById('overview-pane-discipline');
  const docBtn=document.getElementById('ovtab-doc-types');
  const discBtn=document.getElementById('ovtab-discipline');
  if(docPane)docPane.classList.toggle('active',window._overviewTableTab==='docTypes');
  if(discPane)discPane.classList.toggle('active',window._overviewTableTab==='discipline');
  if(docBtn)docBtn.classList.toggle('active',window._overviewTableTab==='docTypes');
  if(discBtn)discBtn.classList.toggle('active',window._overviewTableTab==='discipline');
}}

function renderDTTable(d){{
  const tbody=document.getElementById('dt-tbody'),empty=document.getElementById('dt-empty');
  tbody.innerHTML='';let rows=[],i=0;
  d.forEach(p=>{{(p.dt_stats||[]).forEach(dt=>{{rows.push({{...dt,pcode:p.code,pid:p.id}});}});}});
  if(!rows.length){{empty.style.display='block';return;}}
  empty.style.display='none';
  rows.forEach(r=>{{
    const pct=r.total?Math.round(r.approved/r.total*100):0;
    const col=pct>=80?'#16a34a':pct>=50?'#f59e0b':'#ef4444';
    const tr=document.createElement('tr');
    tr.className=`dt-summary-row${{r.overdue>0?' warn':''}}${{i%2?' alt':''}}`;
    i++;
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
  const mobile=document.getElementById('disc-mobile-list');
  tbody.innerHTML='';let rows=0,groupIdx=0;
  if(mobile)mobile.innerHTML='';
  const mk=(label,v,c)=>`<td class="disc-num-cell" data-label="${{label}}" style="color:${{c}}">${{v||0}}</td>`;
  const safe=v=>String(v??'').replace(/[&<>"']/g,m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]));
  const metric=(label,value,color)=>`<div class="disc-mobile-cell"><div class="disc-mobile-lbl">${{label}}</div><div class="disc-mobile-val" style="color:${{color||'var(--tx)'}}">${{value||0}}</div></div>`;
  data.forEach(p=>{{(p.dt_stats||[]).forEach(dt=>{{
    const disc=dt.disc_breakdown||[];if(!disc.length)return;
    rows++;groupIdx++;
    const groupId=`discgrp-${{groupIdx}}`;
    const warn=(Number(dt.overdue||0)>0)||disc.some(ds=>Number(ds.overdue||0)>0);
    if(mobile){{
      const card=document.createElement('div');
      card.className=`disc-mobile-card${{warn?' warn':''}}`;
      card.dataset.group=groupId;
      card.innerHTML=`<div class="disc-mobile-head">
        <div>
          <div class="disc-mobile-title">${{safe(p.code)}} | ${{safe(dt.code)}}</div>
          <div class="disc-mobile-sub">${{disc.length}} ${{disc.length===1?'discipline':'disciplines'}}</div>
        </div>
        <button type="button" class="disc-expander" aria-expanded="false" onclick="toggleDiscMobileGroup('${{groupId}}', this)">â–¼</button>
      </div>
      <div class="disc-mobile-grid">
        ${{metric('Project',safe(p.code),'var(--pr)')}}
        ${{metric('Doc Type',safe(dt.code),'var(--pr)')}}
        ${{metric('Disciplines',disc.length,'#2f4f64')}}
        ${{metric('Total',dt.total,'var(--pr)')}}
        ${{metric('Approved',dt.approved,'#16a34a')}}
        ${{metric('Pending',dt.pending,'#f59e0b')}}
        ${{metric('Rejected',dt.rejected||0,'#7c3aed')}}
        ${{metric('Overdue',dt.overdue,'#ef4444')}}
      </div>
      <div class="disc-mobile-children">
        ${{disc.map(ds=>`<div class="disc-mobile-child">
          <div class="disc-mobile-title">${{safe(ds.disc)}}</div>
          <div class="disc-mobile-grid">
            ${{metric('Total',ds.total,'var(--pr)')}}
            ${{metric('Approved',ds.approved,'#16a34a')}}
            ${{metric('Pending',ds.pending,'#f59e0b')}}
            ${{metric('Rejected',ds.rejected||0,'#7c3aed')}}
            ${{metric('Overdue',ds.overdue,'#ef4444')}}
          </div>
        </div>`).join('')}}
      </div>`;
      mobile.appendChild(card);
    }}
    const tr=document.createElement('tr');
    tr.className=`disc-group-row${{warn?' warn':''}}${{rows%2===0?' alt':''}}`;
    tr.innerHTML=`<td data-label="Project" style="font-size:10px;color:var(--mu)">${{p.code}}</td>
      <td data-label="Doc Type" style="font-weight:700;color:var(--pr)">${{dt.code}}</td>
      <td data-label="Disciplines"><div style="display:flex;align-items:center;gap:8px"><span class="disc-badge">${{disc.length}}</span><span class="disc-meta">${{disc.length===1?'1 discipline':`${{disc.length}} disciplines`}}</span></div></td>
      ${{mk('Total',dt.total,'var(--pr)')}}
      ${{mk('Approved',dt.approved,'#16a34a')}}
      ${{mk('Pending',dt.pending,'#f59e0b')}}
      ${{mk('Rejected',dt.rejected||0,'#7c3aed')}}
      ${{mk('Overdue',dt.overdue,'#ef4444')}}
      <td style="text-align:center"><button type="button" class="disc-expander" data-group="${{groupId}}" aria-expanded="false" onclick="toggleDiscGroup('${{groupId}}', this)">▼</button></td>`;
    tbody.appendChild(tr);
    disc.forEach(ds=>{{
      const child=document.createElement('tr');
      child.className='disc-child-row';
      child.dataset.group=groupId;
      child.innerHTML=`<td class="disc-child-spacer"></td>
        <td style="font-size:10px;color:var(--mu)">${{dt.code}}</td>
        <td class="disc-child-label">${{ds.disc}}</td>
        ${{mk('Total',ds.total,'var(--pr)')}}
        ${{mk('Approved',ds.approved,'#16a34a')}}
        ${{mk('Pending',ds.pending,'#f59e0b')}}
        ${{mk('Rejected',ds.rejected||0,'#7c3aed')}}
        ${{mk('Overdue',ds.overdue,'#ef4444')}}
        <td></td>`;
      tbody.appendChild(child);
    }});
  }});}});
  [...tbody.querySelectorAll('.disc-group-row')].forEach(tr=>{{
    const cells=tr.children;
    if(cells[0])cells[0].dataset.label='Project';
    if(cells[1])cells[1].dataset.label='Doc Type';
    if(cells[2])cells[2].dataset.label='Disciplines';
    if(cells[8])cells[8].dataset.label='View';
  }});
  [...tbody.querySelectorAll('.disc-child-row')].forEach(tr=>{{
    const cells=tr.children;
    if(cells[0])cells[0].dataset.label='';
    if(cells[1])cells[1].dataset.label='Doc Type';
    if(cells[2])cells[2].dataset.label='Discipline';
    if(cells[8])cells[8].dataset.label='';
  }});
  if(empty)empty.style.display=rows?'none':'block';
}}

function toggleDiscGroup(groupId, btn){{
  const rows=document.querySelectorAll(`#disc-tbody tr[data-group="${{groupId}}"]`);
  const shouldOpen=!btn.classList.contains('open');
  rows.forEach(row=>row.classList.toggle('open', shouldOpen));
  btn.classList.toggle('open', shouldOpen);
  btn.setAttribute('aria-expanded', shouldOpen ? 'true' : 'false');
  btn.textContent=shouldOpen ? '▲' : '▼';
}}

// ── Analytics Tab ─────────────────────────────────────────
function toggleDiscMobileGroup(groupId, btn){{
  const card=document.querySelector(`#disc-mobile-list .disc-mobile-card[data-group="${{groupId}}"]`);
  if(!card)return;
  const shouldOpen=!card.classList.contains('open');
  card.classList.toggle('open', shouldOpen);
  btn.classList.toggle('open', shouldOpen);
  btn.setAttribute('aria-expanded', shouldOpen ? 'true' : 'false');
  btn.textContent=shouldOpen ? 'â–²' : 'â–¼';
}}

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
async function loadDailyDigest() {{
  const container = document.getElementById('daily-digest-content');
  if(!container) return;
  const dateInput = document.getElementById('digest-date');
  if (dateInput && !dateInput.value) {{
    const today = new Date();
    today.setMinutes(today.getMinutes() - today.getTimezoneOffset());
    dateInput.value = today.toISOString().split('T')[0];
  }}
  const selectedDate = dateInput ? dateInput.value : '';
  container.innerHTML = '<div style="color:var(--mu);font-size:14px;">Loading...</div>';
  try {{
    let targetPid = typeof window.PID !== 'undefined' ? window.PID : (typeof _currentPid !== 'undefined' && _currentPid ? _currentPid : 'all');
    const qs = selectedDate ? '?date=' + selectedDate : '';
    const r = await apiFetch('/api/daily_digest/' + targetPid + qs);
    if(!r) {{ container.innerHTML = '<div style="color:red">Failed to load digest (Null response).</div>'; return; }}
    if(r.error) {{
      console.error("Daily Digest Error:", r.error, r.trace);
      container.innerHTML = '<div style="color:red;white-space:pre-wrap;font-size:12px;font-family:monospace;background:#fee2e2;padding:10px;border-radius:6px;">' + 
        '<b>Error:</b> ' + r.error + '<br><br><b>Trace:</b><br>' + (r.trace || 'No trace available') + '</div>';
      return;
    }}
    
    let html = '';
    
    // Received Card
    html += '<div class="panel" style="border-top:4px solid #3b82f6;"><div class="panel-title" style="color:#1e40af;font-size:16px;">📥 Received (' + (r.received||[]).length + ')</div><div style="margin-top:12px;">';
    if((r.received||[]).length === 0) html += '<div style="color:var(--mu);font-size:13px;font-style:italic;">No documents received.</div>';
    (r.received||[]).forEach(doc => {{
      html += '<div style="padding:4px 8px;border-bottom:1px solid var(--bd);display:flex;flex-direction:column;gap:2px;">';
      html += '<div style="display:flex;justify-content:space-between;align-items:center;gap:6px;">';
      html += '<div style="font-weight:700;color:var(--tx);font-size:11px;word-break:break-all;line-height:1;">'+doc.docNo+'</div>';
      if((doc.project_name || doc.project_id)) html += '<div style="font-size:8px;background:#f1f5f9;color:#475569;padding:1px 4px;border-radius:3px;white-space:nowrap;border:1px solid #cbd5e1;flex-shrink:0;">'+(doc.project_name||doc.project_id)+'</div>';
      html += '</div><div title="' + escHtml(doc.title||'') + '" style="color:var(--mu);font-size:10px;line-height:1.1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">'+doc.title+'</div></div>';
    }});
    html += '</div></div>';

    // Issued Card
    html += '<div class="panel" style="border-top:4px solid #f97316;"><div class="panel-title" style="color:#c2410c;font-size:16px;">📤 Issued (' + (r.issued||[]).length + ')</div><div style="margin-top:12px;">';
    if((r.issued||[]).length === 0) html += '<div style="color:var(--mu);font-size:13px;font-style:italic;">No documents issued.</div>';
    (r.issued||[]).forEach(doc => {{
      html += '<div style="padding:4px 8px;border-bottom:1px solid var(--bd);display:flex;flex-direction:column;gap:2px;">';
      html += '<div style="display:flex;justify-content:space-between;align-items:center;gap:6px;">';
      html += '<div style="font-weight:700;color:var(--tx);font-size:11px;word-break:break-all;line-height:1;">'+doc.docNo+'</div>';
      if((doc.project_name || doc.project_id)) html += '<div style="font-size:8px;background:#fef3c7;color:#b45309;padding:1px 4px;border-radius:3px;white-space:nowrap;border:1px solid #fde68a;flex-shrink:0;">'+(doc.project_name||doc.project_id)+'</div>';
      html += '</div><div title="' + escHtml(doc.title||'') + '" style="color:var(--mu);font-size:10px;line-height:1.1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">'+doc.title+'</div></div>';
    }});
    html += '</div></div>';

    // Replied Card
    html += '<div class="panel" style="border-top:4px solid #22c55e;"><div class="panel-title" style="color:#166534;font-size:16px;">✅ Replied (' + (r.replied||[]).length + ')</div><div style="margin-top:12px;">';
    if((r.replied||[]).length === 0) html += '<div style="color:var(--mu);font-size:13px;font-style:italic;">No documents replied.</div>';
    (r.replied||[]).forEach(doc => {{
      html += '<div style="padding:4px 8px;border-bottom:1px solid var(--bd);display:flex;flex-direction:column;gap:2px;">';
      html += '<div style="display:flex;justify-content:space-between;align-items:center;gap:6px;">';
      html += '<div style="font-weight:700;color:var(--tx);font-size:11px;word-break:break-all;line-height:1;">'+doc.docNo+'</div>';
      if((doc.project_name || doc.project_id)) html += '<div style="font-size:8px;background:#dcfce7;color:#15803d;padding:1px 4px;border-radius:3px;white-space:nowrap;border:1px solid #bbf7d0;flex-shrink:0;">'+(doc.project_name||doc.project_id)+'</div>';
      html += '</div><div title="' + escHtml(doc.title||'') + '" style="color:var(--mu);font-size:10px;line-height:1.1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">'+doc.title+'</div>';
      if(doc.status) html += '<div style="font-size:9px;font-weight:700;color:var(--tx2);">Status: ' + doc.status + '</div>';
      html += '</div>';
    }});
    html += '</div></div>';

    container.innerHTML = html;
  }} catch(e) {{
    console.error(e);
    container.innerHTML = '<div style="color:red;white-space:pre-wrap;font-size:12px;font-family:monospace;background:#fee2e2;padding:10px;border-radius:6px;"><b>Fatal UI Error:</b> ' + e.message + '</div>';
  }}
}}

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
    const assigned_cached = u.projects || [];
    const row=document.createElement('div');row.className='urow';
    row.innerHTML=`<span style="flex:1;font-weight:600">👤 ${{u.username}}</span>
      <input id="email-${{u.username}}" placeholder="Email (e.g. dc@company.com)" value="${{u.email || ''}}" style="flex:1; margin-right: 8px; padding: 2px 6px; font-size: 11px; border: 1px solid var(--bd); border-radius: 4px;" onblur="updUsrEmail('${{u.username}}')">
      <span class="badge" style="background:#fef3c7;color:#92400e">${{u.role.toUpperCase()}}</span>
      ${{u.username!=='admin'?`<button class="btn btn-sc btn-sm" onclick="chgPw('${{u.username}}')">🔑 PW</button>
        <button class="btn btn-er btn-sm" onclick="delUsr('${{u.username}}')">✕</button>`:
        '<span style="font-size:10px;color:var(--mu)">(protected)</span>'}}`;
    if(u.username!=='admin'){{
      const roleSel=document.createElement('select');
      roleSel.id='role-'+u.username;
      roleSel.className='btn btn-sc btn-sm';
      roleSel.style.cssText='height:auto;padding:4px 8px;outline:none;';
      roleSel.innerHTML=['viewer','editor','admin','superadmin'].map(r=>`<option value="${{r}}" ${{u.role===r?'selected':''}}>${{r==='superadmin'?'Super Admin':r.charAt(0).toUpperCase()+r.slice(1)}}</option>`).join('');
      const roleBtn=document.createElement('button');
      roleBtn.className='btn btn-pr btn-sm';
      roleBtn.textContent='Save Role';
      roleBtn.onclick=()=>updUsrRole(u.username);
      row.insertBefore(roleSel,row.children[1]||null);
      row.insertBefore(roleBtn,row.children[2]||null);
    }}
    body.appendChild(row);
    if(u.role!=='superadmin'){{
      const ad=document.createElement('div');
      ad.style.cssText='padding:4px 10px 10px 32px;border-bottom:1px solid var(--bd);margin-bottom:4px';
      ad.innerHTML='<div style="font-size:10px;color:var(--mu);margin-bottom:6px">Project access:</div>';
      const assigned=assigned_cached;
      const pl=document.createElement('div');pl.style.cssText='display:flex;flex-wrap:wrap;gap:5px';
      projects.forEach(p=>{{
        const projAccess = assigned.find(a => a.project_id === p.id);
        const isOn = !!projAccess;
        const isDC = isOn && projAccess.is_dc;
        const btn=document.createElement('button');
        btn.style.cssText='padding:3px 10px;border-radius:4px;cursor:pointer;font-size:11px;font-weight:700;font-family:inherit;transition:all .15s;border:2px solid '+(isDC?'#22c55e':(isOn?'#f0a500':'#e2e8f0'))+';background:'+(isOn?'#1a3a5c':'#f8fafc')+';color:'+(isOn?'#fff':'#94a3b8');
        btn.textContent=p.code + (isDC ? ' (DC)' : '');
        btn.title=p.name + " | Left-click to assign access. Right-click to set as Project DC";
        if(isOn)btn.dataset.on='1';
        
        btn.onclick=async()=>{{
          const on=!!btn.dataset.on;
          btn.dataset.on = on ? '' : '1';
          const isOn = !on;
          btn.style.borderColor = isDC?'#22c55e':(isOn?'#f0a500':'#e2e8f0');
          btn.style.background = isOn?'#1a3a5c':'#f8fafc';
          btn.style.color = isOn?'#fff':'#94a3b8';
          btn.textContent = p.code + (isDC ? ' (DC)' : '');
          try {{
            await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:on?'unassign':'assign',username:u.username,project_id:p.id,is_dc:false}})}});
          }} catch(e) {{ openAdmin(); }}
        }};
        
        btn.oncontextmenu = async (e) => {{
          e.preventDefault();
          const currIsOn = !!btn.dataset.on;
          if(!currIsOn) return toast('User must be assigned to project first','wa');
          await apiFetch('/api/users', {{method:'POST', body:JSON.stringify({{action:'assign',username:u.username,project_id:p.id, is_dc: !isDC}})}});
          openAdmin(); // Reload to refresh isDC constant state accurately
        }};
        
        pl.appendChild(btn);
      }});
      ad.appendChild(pl);body.appendChild(ad);
    }}
  }}
  const at=document.createElement('div');at.className='stitle';at.textContent='➕ Add User';body.appendChild(at);
  const ar=document.createElement('div');
  ar.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1.5fr 1fr 1fr auto;gap:8px;align-items:end">
    <div class="fg"><label>Username</label><input id="nu-name" placeholder="username"></div>
    <div class="fg"><label>Email</label><input id="nu-email" placeholder="Email"></div>
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
  const email=document.getElementById('nu-email')?.value;
  if(!name||!pw){{toast('Username and password required','er');return;}}
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'add',username:name,role,password:pw,email:email}})}});
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
async function updUsrEmail(u){{
  const email=document.getElementById('email-'+u)?.value;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'update_email',username:u,email}})}});
  if(r&&r.ok) toast('Email saved','ok'); else toast('Failed to save email','er');
}}
async function updUsrRole(u){{
  const role=document.getElementById('role-'+u)?.value;
  if(!role)return;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'update_role',username:u,role}})}});
  if(r&&r.ok){{toast('Role updated','ok');closeM('admin-modal');openAdmin();}}
  else toast((r&&r.error)||'Role update failed','er');
}}

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
    custom_sc_json = json.dumps(proj.get("status_colors", {}))

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

    primary_keys = {"code", "name"}
    projbar_primary = "".join(
        f'<div class="pf primary"><span class="pf-lbl">{lbl}</span>'
        f'<span class="pf-val" data-key="{key}">{proj.get(key,"") or "—"}</span></div>'
        for key,lbl in PROJ_FIELDS
        if key in primary_keys and proj.get(key,"").strip())
    projbar_secondary = "".join(
        f'<div class="pf secondary"><span class="pf-lbl">{lbl}</span>'
        f'<span class="pf-val" data-key="{key}">{proj.get(key,"") or "—"}</span></div>'
        for key,lbl in PROJ_FIELDS
        if key not in primary_keys and proj.get(key,"").strip())
    projbar_toggle = ('<button id="projbar-toggle" class="tool-btn" type="button" '
                      'onclick="toggleProjectInfo()">Project Info</button>') if projbar_secondary else ''

    tabs_html = "".join(
        f'<button class="tab-btn" data-id="{dt["id"]}" onclick="switchTab(\'{dt["id"]}\')">'
        f'<span>{dt["code"]}</span><span class="tcnt" id="cnt-{dt["id"]}">0</span></button>'
        for dt in dts)

    _sa_only = role == 'superadmin'
    _hol_btn = " <button class='tool-btn purple' onclick='openSettings()'>🗓 Holidays</button>" if _sa_only else ''
    _col_btn = "<button class='tool-btn purple' onclick='manageColumns()'>⚙ Columns</button>" if _sa_only else ''
    _lst_btn = "<button class='tool-btn' onclick='openLists()'>📋 Lists</button>" if _sa_only else ''
    _color_btn = "<button class='tool-btn purple' onclick='openProjectSettings()'>🎨 Status Colors</button>" if role in ('admin', 'superadmin') else ''
    edit_btns = (f'<button class="tool-btn" onclick="addRecord()">➕ Add</button>'
                 f'{_col_btn}'
                 f'{_hol_btn}'
                 f'{_lst_btn}'
                 f'{_color_btn}'
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
body.dark #projbar,body.dark #toolbar,body.dark #ltr-quickbar,body.dark #noc-summary{{background:#162132;color:#e2e8f0;border-color:#304257}}
body.dark #tabsbar{{background:#0d1f33}}
body.dark #tblwrap{{background:#111b2a}}
body.dark .pf-lbl{{color:#93c5fd}}
body.dark .pf-val{{color:#dbe7f3}}
body.dark #regtbl th{{background:#17314d}}
body.dark .frow th{{background:#101a29}}
body.dark .frow input,body.dark .frow select{{background:#0f1a29;border-color:#314558;color:#e2e8f0}}
body.dark #regtbl td{{border-color:#253648;color:#d7e1ec}}
body.dark #regtbl tr.alt td{{background:#132031}}
body.dark #regtbl tr.ov td{{background:#3a231f;color:#f8d8c7}}
body.dark #regtbl tr.rv td{{background:#101a29;color:#b8c8da}}
body.dark #regtbl tbody tr:hover td{{background:#1e3147;color:#eef5ff}}
body.dark #regtbl tbody tr.ov:hover td{{background:#4a2a24;color:#ffe1d2}}
body.dark #regtbl tbody tr.rv:hover td{{background:#1e3147;color:#eef5ff}}
body.dark #regtbl tbody tr.row-selected td{{background:#1d3a5a!important;color:#f3f8ff!important;box-shadow:inset 0 1px 0 rgba(147,197,253,.2),inset 0 -1px 0 rgba(147,197,253,.2)}}
body.dark #regtbl tbody tr.row-selected td:first-child{{box-shadow:inset 4px 0 0 #93c5fd,inset 0 1px 0 rgba(147,197,253,.2),inset 0 -1px 0 rgba(147,197,253,.2)}}
body.dark #regtbl tbody tr.row-selected:hover td{{background:#25496d!important;color:#fff!important}}
body.dark .flink{{color:#93c5fd}}
body.dark .tool-dd-menu{{background:#162132;border-color:#304257}}
body.dark .tool-dd-menu button{{color:#e2e8f0}}
body.dark .tool-dd-menu button:hover{{background:#101a29;color:#93c5fd}}
body.dark .tool-btn{{color:#e2e8f0;border-color:#304257;background:#101a29}}
body.dark .tab-btn{{color:rgba(255,255,255,.72)}}
body.dark .tab-btn:hover{{background:rgba(255,255,255,.08)}}
body.dark .tcnt{{background:rgba(255,255,255,.14)}}
body.dark #sbar{{background:#0d1f33;color:rgba(255,255,255,.72)}}
body.dark .sitem{{background:#101a29;color:#e2e8f0;border-color:#253648}}
body.dark .sitem button{{background:#f8fafc;color:#0f172a;border-color:#94a3b8}}
body.dark .addrow input{{background:#f8fafc;color:#0f172a;border-color:#cbd5e1}}
@media print{{
  #topbar,#tabbar,.toolrow,.bulkbar,#statusbar,.acts{{display:none!important}}
  body{{height:auto;overflow:visible}}
  #main{{overflow:visible;height:auto}}
  #tblwrap{{overflow:visible;height:auto}}
  #regtbl{{font-size:10px}}
  #regtbl th,#regtbl td{{padding:4px 6px!important;white-space:normal!important}}
  @page{{size:A4 landscape;margin:10mm}}
}}
#projbar{{background:#fff;border-bottom:2px solid var(--pr);padding:5px 10px;
  display:flex;align-items:center;flex-wrap:wrap;flex-shrink:0;gap:6px}}
#projbar-main{{display:flex;align-items:center;flex:1 1 auto;gap:4px;min-width:0;flex-wrap:wrap}}
#projbar-primary,#projbar-extra{{display:flex;align-items:center;gap:0;flex-wrap:wrap;min-width:0}}
#projbar-toggle{{display:none}}
#projbar img{{max-height:30px;max-width:108px;object-fit:contain;flex-shrink:0}}
.pf{{display:flex;flex-direction:column;padding:0 8px;border-right:1px solid var(--bd);min-width:0}}
.pf:last-of-type{{border-right:none}}
.pf-lbl{{font-size:9px;font-weight:700;color:var(--pr);text-transform:uppercase;letter-spacing:.4px}}
.pf-val{{font-size:11px;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px}}
#tabsbar{{background:#0f2640;display:flex;align-items:center;overflow-x:auto;flex-shrink:0;
  padding:4px 8px;scrollbar-width:thin;gap:6px}}
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
#toolbar{{background:#fff;border-bottom:1px solid var(--bd);padding:4px 8px;
  display:flex;align-items:center;gap:5px;flex-shrink:0;flex-wrap:wrap}}
#toolbar-actions{{display:flex;align-items:center;gap:5px;flex-wrap:wrap;min-width:0}}
.tool-btn{{display:flex;align-items:center;gap:4px;padding:4px 9px;background:var(--bg);
  border:1px solid var(--bd);border-radius:var(--rd);cursor:pointer;font-size:11px;
  font-family:inherit;color:var(--tx);transition:all .15s;white-space:nowrap}}
.tool-btn:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.tool-btn.purple:hover{{background:#7c3aed;border-color:#7c3aed}}
.tool-btn.teal:hover{{background:#0891b2;border-color:#0891b2}}
.tool-dd{{position:relative}}
.tool-dd-menu{{position:absolute;top:calc(100% + 4px);left:0;background:#fff;border:1.5px solid var(--bd);border-radius:6px;box-shadow:0 8px 24px rgba(0,0,0,.15);z-index:300;min-width:210px;overflow:hidden}}
.tool-dd-menu button{{display:block;width:100%;text-align:left;padding:9px 14px;border:none;background:none;cursor:pointer;font-size:12px;font-family:inherit;color:#1e2a3a;white-space:nowrap}}
.tool-dd-menu button:hover{{background:#f0f4f8;color:var(--pr)}}
#srchbox{{flex:1;min-width:150px;max-width:260px;padding:4px 10px 4px 28px;border:1px solid var(--bd);
  border-radius:var(--rd);font-family:inherit;font-size:12px;outline:none;
  background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='%236b7a94' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E") no-repeat 7px center}}
#srchbox:focus{{border-color:var(--pl);box-shadow:0 0 0 2px rgba(37,99,168,.1)}}
#main{{flex:1;overflow:hidden;display:flex;flex-direction:column;min-width:0}}
#tblwrap{{flex:1;overflow:auto;min-width:0;-webkit-overflow-scrolling:touch}}
#regtbl{{width:100%;border-collapse:collapse;min-width:980px;font-size:12px}}
.sr{{text-align:center;color:var(--mu);font-size:10px;min-width:40px;white-space:nowrap}}
#regtbl thead{{position:sticky;top:0;z-index:20}}
#regtbl th{{background:var(--pr);color:#fff;padding:8px;text-align:left;font-weight:600;
  white-space:nowrap;border-right:1px solid rgba(255,255,255,.1);cursor:pointer;user-select:none;position:relative}}
#thead tr:first-child th{{position:relative;top:auto;z-index:auto;border-bottom:2px solid rgba(255,255,255,.24);box-shadow:none}}
#regtbl th:hover{{background:var(--pl)}}
.frow{{position:static}}
.frow th{{background:#eef1f7;padding:3px 4px;cursor:default;position:static;top:auto;z-index:auto;box-shadow:0 1px 0 rgba(221,227,237,.95)}}
.frow th:hover{{background:#eef1f7}}
.frow input,.frow select{{width:100%;padding:3px 6px;border:1px solid var(--bd);
  border-radius:3px;font-size:10px;font-family:inherit;background:#fff;outline:none}}
#regtbl td{{padding:5px 8px;border-bottom:1px solid #edf0f5;border-right:1px solid #f3f4f6;
  vertical-align:middle;max-width:220px;word-break:break-word}}
#regtbl tr:hover td{{background:rgba(37,99,168,.10);transition:background .1s}}
#regtbl tr.ov td{{background:#fff5f5}}
#regtbl tr.rv td{{color:var(--mu)}}
#regtbl tr.alt td{{background:#fafbfd}}
#regtbl tbody tr:hover td{{background:#dce8f5;color:var(--tx)}}
#regtbl tbody tr.ov:hover td{{background:#ffe7e7;color:#3b1f1f}}
#regtbl tbody tr.rv:hover td{{background:#fff4d8;color:#3b2a12}}
#regtbl tbody tr.row-selected td{{background:#d7e8fb!important;color:#10233a!important;box-shadow:inset 0 1px 0 rgba(37,99,168,.18),inset 0 -1px 0 rgba(37,99,168,.18)}}
#regtbl tbody tr.row-selected td:first-child{{box-shadow:inset 4px 0 0 #2563a8,inset 0 1px 0 rgba(37,99,168,.18),inset 0 -1px 0 rgba(37,99,168,.18)}}
#regtbl tbody tr.row-selected:hover td{{background:#c8def6!important;color:#10233a!important}}
.chkcell{{text-align:center;width:28px;padding:4px!important}}
.chkcell input{{width:14px;height:14px;cursor:pointer;accent-color:var(--pr)}}
.acts{{white-space:nowrap;width:64px}}
.act{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.act:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.act.del:hover{{background:var(--er);border-color:var(--er)}}
.pr-toggle{{padding:2px 7px;border:1px solid var(--bd);background:#fff;border-radius:3px;cursor:pointer;font-size:11px}}
.pr-toggle:hover{{background:var(--pr);color:#fff;border-color:var(--pr)}}
.pr-items-row{{display:none}}
.pr-items-row.open{{display:table-row}}
.pr-items-row>td{{padding:0!important;border-bottom:1px solid #d9e2ec!important;border-right:none!important;max-width:none!important;background:transparent!important}}
.pr-items-panel{{display:block;width:100%;box-sizing:border-box;padding:16px;background:#f8fafc;border-top:1px solid #d9e2ec;border-bottom:1px solid #d9e2ec;position:relative;z-index:1}}
.pr-items-grid-wrap{{display:block;width:100%;max-width:100%;overflow-x:auto;padding-bottom:2px}}
.pr-items-title{{display:flex;align-items:center;gap:8px;margin-bottom:8px;font-size:11px;font-weight:800;color:var(--pr);text-transform:uppercase;letter-spacing:.45px}}
.pr-items-title:before{{content:"";width:4px;height:16px;border-radius:999px;background:var(--ac);display:inline-block}}
.pr-items-grid{{display:grid;grid-template-columns:minmax(520px,1fr) 120px 100px minmax(240px,.6fr);min-width:1100px;width:100%;border:1px solid #dbe3ed;border-radius:6px;overflow:hidden;background:#fff;font-size:11px}}
.pr-items-grid-head{{background:#173f66;color:#fff;font-weight:800;padding:10px 12px;border-right:1px solid rgba(255,255,255,.18)}}
.pr-items-grid-head:last-child{{border-right:none}}
.pr-items-cell{{padding:10px 12px;border-top:1px solid #edf2f7;border-right:1px solid #edf2f7;background:#fff;color:#10233a;white-space:normal;word-break:normal;line-height:1.35}}
.pr-items-cell:nth-child(4n){{border-right:none}}
.pr-items-cell.qty{{text-align:center;font-variant-numeric:tabular-nums}}
.pr-items-cell.unit{{white-space:nowrap}}
.pr-items-section{{grid-column:1/-1;background:#e8eef6;color:var(--pr);font-weight:800;text-transform:uppercase;letter-spacing:.35px;padding:9px 12px;border-top:1px solid #d7dee8}}
.pr-items-empty{{color:var(--mu);font-size:11px;padding:6px 0}}
body.dark .pr-items-row>td{{background:transparent!important;border-color:#304257!important}}
body.dark .pr-items-panel{{background:#101a29;border-color:#304257}}
body.dark .pr-items-grid{{background:#162132;border-color:#304257}}
body.dark .pr-items-grid-head{{background:#17314d;color:#dbeafe;border-color:#304257}}
body.dark .pr-items-cell{{background:#162132;color:#d7e1ec;border-color:#253648}}
body.dark .pr-items-section{{background:#1e3147;color:#dbeafe;border-color:#304257}}
.sbadge{{display:inline-flex;align-items:center;gap:6px;border-radius:9999px;padding:4px 10px;font-size:11px;font-weight:600;white-space:nowrap}}
.sbadge::before{{content:'';width:6px;height:6px;border-radius:50%;background:currentColor;display:block}}
.flink{{color:var(--pl);text-decoration:underline;cursor:pointer;font-size:11px}}
.ovdate{{color:#dc2626;font-weight:700}}
.mlcell{{white-space:pre-line!important;word-break:break-word}}
#bulkbar{{display:none;background:#1a3a5c;color:#fff;padding:5px 14px;align-items:center;gap:10px;font-size:12px;flex-shrink:0}}
#bulkbar.show{{display:flex}}
#sbar{{background:var(--pr);color:rgba(255,255,255,.75);padding:2px 12px;font-size:10px;display:flex;gap:12px;flex-wrap:wrap;flex-shrink:0}}
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
body.dark .ms-con{{background:#101a29;border-color:#304257;color:#e2e8f0}}
body.dark .ms-ph{{color:#9fb0c6}}
body.dark .ms-dd{{background:#101a29;border-color:#304257;box-shadow:0 14px 36px rgba(0,0,0,.42)}}
body.dark .ms-opt{{background:#101a29;color:#dbe7f3;border-bottom-color:#253648}}
body.dark .ms-opt:hover{{background:#1e3147;color:#eef5ff}}
body.dark .ms-opt.sel{{background:#17314d;color:#eef5ff}}
body.dark .ms-opt input{{accent-color:#60a5fa}}
.letter-flow{{position:relative;padding-left:10px}}
.letter-flow-line{{position:absolute;left:13px;top:0;bottom:0;width:2px;background:linear-gradient(180deg,#d7e0e6,#c7d2da)}}
.letter-node{{position:relative;padding-left:30px;margin-bottom:14px}}
.letter-dot{{position:absolute;left:8px;top:14px;width:12px;height:12px;border-radius:999px;background:#2F4F64;border:3px solid #fff;box-shadow:0 0 0 2px rgba(47,79,100,.18)}}
.letter-dot.current{{background:#8BC34A;box-shadow:0 0 0 2px rgba(139,195,74,.35)}}
.letter-card{{background:#fff;border:1px solid #d7e0e6;border-left:4px solid #d7e0e6;border-radius:12px;padding:12px 14px;box-shadow:0 6px 18px rgba(15,23,42,.05)}}
.letter-card.current{{background:#f6fbef;border-color:#8BC34A;border-left-color:#8BC34A;box-shadow:0 10px 26px rgba(139,195,74,.14)}}
.letter-card-kicker{{font-size:9px;color:#7d8a95;text-transform:uppercase;letter-spacing:.35px;margin-bottom:6px}}
.letter-card-ref{{font-size:12px;font-weight:800;color:var(--pr)}}
.letter-card-subject{{font-size:12px;color:var(--tx);margin-top:6px;line-height:1.45}}
.letter-card-meta{{font-size:10px;color:var(--mu);margin-top:8px;line-height:1.5}}
.letter-date-badge{{display:inline-flex;align-items:center;background:#eef4f8;color:#2F4F64;border-radius:999px;padding:4px 10px;font-size:10px;font-weight:800;letter-spacing:.2px}}
.letter-current-badge{{flex-shrink:0;background:#8BC34A;color:#17351b;border-radius:999px;padding:3px 8px;font-size:9px;font-weight:800;letter-spacing:.35px;text-transform:uppercase}}
.letter-branch-cue{{color:#7d8a95;font-weight:800;margin-right:8px}}
.letter-connector{{position:absolute;left:-24px;top:0;bottom:22px;width:2px;background:#cbd5df;border-radius:2px}}
.letter-connector-arm{{position:absolute;left:-24px;top:24px;width:18px;height:2px;background:#cbd5df;border-radius:2px}}
body.dark .letter-flow-line{{background:linear-gradient(180deg,#3a526d,#253648)}}
body.dark .letter-dot{{background:#60a5fa;border-color:#101a29;box-shadow:0 0 0 2px rgba(96,165,250,.28)}}
body.dark .letter-dot.current{{background:#8BC34A;box-shadow:0 0 0 2px rgba(139,195,74,.42)}}
body.dark .letter-card{{background:#162132;border-color:#304257;border-left-color:#496780;box-shadow:0 10px 28px rgba(0,0,0,.28)}}
body.dark .letter-card.current{{background:#1d2b1d;border-color:#8BC34A;border-left-color:#8BC34A;box-shadow:0 10px 28px rgba(0,0,0,.32)}}
body.dark .letter-card-kicker,body.dark .letter-card-meta{{color:#b8c8da}}
body.dark .letter-card-ref{{color:#dbeafe}}
body.dark .letter-card-subject{{color:#eef5ff}}
body.dark .letter-date-badge{{background:#233850;color:#bfdbfe}}
body.dark .letter-branch-cue{{color:#93c5fd}}
body.dark .letter-connector,body.dark .letter-connector-arm{{background:#496780}}
body.dark #thread-body [style*="background:#fff"],
body.dark #timeline-body [style*="background:#fff"]{{background:#162132!important;border-color:#304257!important;box-shadow:0 10px 28px rgba(0,0,0,.28)!important}}
body.dark #thread-body [style*="background:#f6fbef"],
body.dark #timeline-body [style*="background:#f6fbef"]{{background:#1d2b1d!important;border-color:#8BC34A!important;box-shadow:0 10px 28px rgba(0,0,0,.32)!important}}
body.dark #thread-body [style*="color:#7d8a95"],
body.dark #timeline-body [style*="color:#7d8a95"]{{color:#b8c8da!important}}
body.dark #thread-body [style*="background:#eef4f8"],
body.dark #timeline-body [style*="background:#eef4f8"]{{background:#233850!important;color:#bfdbfe!important}}
body.dark #thread-body [style*="border:3px solid #fff"],
body.dark #timeline-body [style*="border:3px solid #fff"]{{border-color:#101a29!important}}
body.dark #thread-body [style*="linear-gradient(180deg,#d7e0e6"],
body.dark #timeline-body [style*="linear-gradient(180deg,#d7e0e6"]{{background:linear-gradient(180deg,#3a526d,#253648)!important}}
.empty{{text-align:center;padding:60px 20px;color:var(--mu)}}
.slist{{list-style:none;display:flex;flex-direction:column;gap:3px;
  max-height:190px;overflow-y:auto;border:1px solid var(--bd);border-radius:var(--rd);padding:4px}}
.sitem{{display:flex;align-items:center;gap:8px;padding:4px 8px;background:var(--bg);
  border-radius:3px;font-size:11px}}
.sitem .nm{{flex:1}}
.sitem button{{padding:2px 8px;font-size:10px;border:1px solid var(--bd);background:#fff;
  border-radius:3px;cursor:pointer}}
.sitem button:hover{{background:var(--er);color:#fff;border-color:var(--er)}}
.rtl-txt{{direction:rtl;text-align:right}}
.addrow{{display:flex;gap:6px;margin-top:6px}}
.addrow input{{flex:1;padding:5px 8px;border:1px solid var(--bd);border-radius:3px;
  font-size:11px;font-family:inherit;outline:none}}
#rec-modal .record-modal{{width:min(96vw,980px)!important;max-width:980px!important}}
#rec-modal .mbody{{padding:14px 16px 12px}}
#rec-modal .record-modal-actions{{gap:10px;flex-wrap:wrap;border-top:1px solid rgba(148,163,184,.34);background:linear-gradient(180deg,rgba(248,250,252,.96),#eef4fb);box-shadow:0 -10px 24px rgba(15,23,42,.10);z-index:5}}
body.dark #rec-modal .record-modal-actions{{border-top-color:#304257;background:linear-gradient(180deg,#142033,#101a29);box-shadow:0 -12px 26px rgba(0,0,0,.28)}}
.record-modal-actions .btn{{min-height:38px}}
#pr-items-editor{{overflow-x:auto;padding-bottom:2px}}
.pr-items-editor table{{min-width:560px}}
@media(max-width:1200px){{
  #projbar{{flex-wrap:wrap;gap:6px;padding:5px 10px}}
  .pf{{min-width:110px;flex:1 1 120px}}
  #toolbar{{gap:6px}}
  #toolbar-actions{{flex:1 1 auto}}
  #srchbox{{max-width:none;flex:1 1 220px}}
}}
@media(max-width:900px){{
  .tool-btn{{min-height:34px}}
  #tabsbar{{padding:6px 8px}}
  .tab-btn{{border:1px solid rgba(255,255,255,.16);border-bottom:none;border-radius:999px;padding:8px 11px;color:rgba(255,255,255,.82)}}
  .tab-btn.active{{background:#fff;color:#0f2640}}
  .tab-add{{margin-left:0;border-radius:999px;padding:6px 11px;font-size:13px}}
  #srchbox{{flex:1 1 100%;min-width:0}}
  #regtbl{{font-size:11px}}
  #regtbl td,#regtbl th{{padding:4px 5px}}
  #tblwrap{{padding-bottom:4px}}
}}
@media(max-width:768px){{
  #topbar{{height:34px;padding:0 8px;gap:5px;overflow:hidden}}
  #topbar .topbar-title-full{{display:none}}
  #topbar .topbar-title-short{{display:inline!important;font-size:13px!important}}
  #topbar .topbar-title-short::after{{content:" Register"}}
  #topbar .topbar-mark{{font-size:16px!important}}
  #topbar .sp{{flex:0 1 6px;min-width:2px}}
  #topbar .tb-btn{{padding:3px 6px;font-size:9.5px;line-height:1.1;border-radius:7px;min-height:25px;display:inline-flex;align-items:center;justify-content:center;white-space:nowrap}}
  #topbar form .tb-btn{{width:auto;padding:3px 6px}}
  #topbar > span[style*="padding:0 4px"]{{display:none}}
  #topbar .topbar-user{{font-size:8px;gap:0;flex-shrink:0}}
  #topbar .topbar-user-name{{display:none}}
  #topbar .topbar-user span:last-child{{padding:2px 5px!important;font-size:8px!important;line-height:1.05;max-width:62px;text-align:center}}
  #projbar{{padding:2px 8px;gap:5px;align-items:flex-start;border-bottom-width:1px}}
  #projbar img{{max-height:34px;max-width:82px;margin-top:0}}
  #projbar-main{{display:grid;grid-template-columns:minmax(130px,.9fr) minmax(0,1.1fr);gap:2px 10px;flex:1 1 auto;width:auto;min-width:0}}
  #projbar-primary,#projbar-extra{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:1px 7px;width:100%;min-width:0}}
  #projbar-primary{{grid-template-columns:minmax(0,1fr)}}
  #projbar-primary .pf,#projbar-extra .pf{{padding:0;border-right:none;min-width:0;display:block}}
  #projbar-primary .pf-lbl,#projbar-extra .pf-lbl{{font-size:6px;line-height:.9;color:var(--mu);letter-spacing:.16px;margin-bottom:0}}
  #projbar-primary .pf-val,#projbar-extra .pf-val{{font-size:9px;line-height:1;max-width:none;white-space:normal;display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical;overflow:hidden}}
  #projbar-primary .pf.primary:first-child .pf-val{{font-size:10px;font-weight:800;color:var(--pr);text-transform:uppercase;-webkit-line-clamp:1}}
  #projbar-primary .pf.primary:last-child{{grid-column:1/-1;border-bottom:none;padding-bottom:0;margin-bottom:0}}
  #projbar-primary .pf.primary:last-child .pf-val{{font-size:10.5px;font-weight:800;line-height:1;-webkit-line-clamp:1}}
  #projbar-extra{{padding-top:0;border-top:none}}
  #projbar-toggle,.proj-edit-btn{{display:none!important}}
  #tabsbar{{padding:2px 7px;gap:3px;scroll-snap-type:x proximity}}
  .tab-btn{{padding:3px 7px;font-size:8.5px;min-height:22px;border-radius:999px;scroll-snap-align:start}}
  .tcnt{{padding:0 4px;font-size:7.5px;line-height:1.25}}
  .tab-add{{padding:2px 7px;font-size:10px;min-height:22px}}
  #toolbar{{padding:2px 7px;gap:3px;align-items:stretch}}
  #toolbar-actions{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));width:100%;gap:3px}}
  .tool-btn{{justify-content:center;min-width:0;padding:2px 3px;font-size:8.3px;min-height:22px;border-radius:6px;gap:2px}}
  .tool-dd{{min-width:0}}
  .tool-dd .tool-btn{{width:100%}}
  #srchbox{{width:100%;flex:1 1 100%;max-width:none;min-height:25px;padding:3px 8px 3px 27px;font-size:10px;border-radius:7px}}
  #tblwrap{{min-height:56vh;padding-bottom:2px}}
  #regtbl{{min-width:1040px;font-size:10.5px}}
  #regtbl th{{padding:5px 6px!important;font-size:10px;line-height:1.15}}
  .frow th{{padding:2px 4px!important}}
  .frow input,.frow select{{height:27px;padding:2px 5px;font-size:10px;border-radius:4px}}
  #regtbl td{{padding:4px 6px!important;line-height:1.18;max-width:190px}}
  .chkcell{{width:26px;padding:3px!important}}
  .chkcell input{{width:13px;height:13px}}
  .sr{{font-size:9px;min-width:34px;white-space:nowrap}}
  .sbadge{{padding:2px 7px;font-size:9px;line-height:1.15}}
  #sbar{{padding:2px 8px;font-size:9px;gap:9px;line-height:1.2;justify-content:space-between}}
  #ltr-quickbar{{padding:4px 8px!important;gap:4px!important;display:flex!important;flex-direction:row!important}}
  #ltr-quickbar .tool-btn{{min-height:22px!important;font-size:8.5px!important;padding:2px 6px!important;flex:1 1 0!important}}
}}
@media(max-width:640px){{
  #topbar{{height:32px;padding:0 5px;gap:3px}}
  #topbar .tb-btn{{padding:3px 5px;font-size:9px;min-height:24px}}
  #topbar .topbar-title-full{{display:none}}
  #topbar .topbar-title-short{{display:inline}}
  #topbar .topbar-mark{{font-size:15px}}
  #topbar .topbar-user{{font-size:9px;gap:4px}}
  #topbar .topbar-user-name{{display:none}}
  #tabsbar{{padding:2px 6px;gap:3px}}
  .tab-btn{{padding:3px 7px;font-size:8.5px;min-height:22px}}
  .tab-add{{padding:2px 7px;font-size:10px;min-height:22px}}
  #projbar{{padding:2px 7px;align-items:flex-start;gap:4px}}
  #projbar img{{max-height:32px;max-width:78px}}
  #projbar-main{{display:grid;grid-template-columns:minmax(120px,.9fr) minmax(0,1.1fr);align-items:start;gap:2px 8px;width:auto;flex:1 1 auto}}
  #projbar-primary,#projbar-extra{{display:grid;grid-template-columns:1fr 1fr;gap:1px 6px;width:100%;min-width:0}}
  #projbar-primary .pf,#projbar-extra .pf{{padding:0;border-right:none;border-bottom:none;min-width:0;display:block;flex:none}}
  #projbar-primary .pf-lbl,#projbar-extra .pf-lbl{{display:block;font-size:6px;line-height:.9;color:var(--mu);letter-spacing:.18px}}
  #projbar-primary .pf-val,#projbar-extra .pf-val{{font-size:9px;line-height:1;max-width:none;white-space:normal;display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical;overflow:hidden}}
  #projbar-primary .pf.primary:first-child .pf-val{{font-size:10px;font-weight:800;color:var(--pr);text-transform:uppercase;letter-spacing:.3px;-webkit-line-clamp:1}}
  #projbar-primary .pf.primary:last-child{{grid-column:1/-1}}
  #projbar-primary .pf.primary:last-child .pf-val{{font-size:10px;font-weight:800}}
  #projbar-extra{{padding-top:0;border-top:none}}
  #projbar-toggle{{display:none}}
  .proj-edit-btn{{display:none}}
  #toolbar{{padding:2px 6px;align-items:stretch;gap:3px}}
  #toolbar-actions{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));width:100%;gap:3px}}
  .tool-btn{{justify-content:center;min-width:0;padding:2px 3px;font-size:8px;min-height:21px}}
  #rec-modal .record-modal{{width:96vw!important;max-width:96vw!important}}
  #rec-modal .mbody{{padding:12px}}
  .record-modal-actions{{padding:10px 12px;justify-content:stretch}}
  .record-modal-actions .btn{{flex:1 1 100%;width:100%;min-height:40px}}
  .modal{{width:96vw;max-height:96vh}}
  #srchbox{{width:100%;flex:1 1 100%;max-width:none;padding-top:2px;padding-bottom:2px;font-size:10px;min-height:24px}}
  #regtbl{{min-width:1040px;font-size:10.5px}}
  #regtbl th{{white-space:nowrap!important;word-break:normal!important}}
  #tblwrap{{min-height:56vh}}
  #sbar{{padding:2px 8px;font-size:9px;gap:8px}}
  #ltr-quickbar{{display:flex!important;flex-direction:row!important;padding:3px 6px!important;gap:3px!important}}
  #ltr-quickbar .tool-btn{{flex:1 1 0!important;min-height:21px!important;font-size:8px!important;padding:2px 5px!important}}
  .overview-pr-panel{{overflow:hidden}}
  .pr-analytics-grid{{grid-template-columns:minmax(0,1fr)!important;width:100%}}
  .pr-block{{width:100%;max-width:100%;min-width:0;overflow:hidden}}
  .pr-projects,.pr-trades{{min-height:0!important}}
  .pr-trades > div:last-child{{min-width:0;height:180px;max-height:none!important}}
  #overview-pane-discipline .tbl-wrap{{display:none}}
  #overview-pane-discipline .disc-mobile-list{{display:block}}
}}
@media(max-width:520px){{
  #topbar{{height:31px;padding:0 5px;gap:3px}}
  #topbar .tb-btn{{font-size:8.5px;width:auto;padding:2px 4px;overflow:visible;min-height:23px}}
  #topbar a.tb-btn{{width:auto;font-size:8.5px;padding:2px 5px}}
  #topbar .topbar-user span:last-child{{max-width:50px;font-size:7px!important;padding:2px 4px!important}}
  #projbar{{padding:2px 6px}}
  #projbar img{{max-height:32px;max-width:76px}}
  #projbar-primary,#projbar-extra{{gap:1px 6px}}
  #projbar-primary .pf-lbl,#projbar-extra .pf-lbl{{font-size:6px}}
  #projbar-primary .pf-val,#projbar-extra .pf-val{{font-size:8.8px;line-height:1;-webkit-line-clamp:1}}
  #projbar-primary .pf.primary:last-child .pf-val{{font-size:10px}}
  #tabsbar{{padding:2px 6px}}
  .tab-btn{{padding:3px 7px;font-size:8.5px;gap:3px}}
  #toolbar{{padding:2px 5px}}
  #toolbar-actions{{grid-template-columns:repeat(4,minmax(0,1fr));gap:3px}}
  .tool-btn{{font-size:7.8px;min-height:20px;padding:2px 3px}}
  #srchbox{{min-height:23px;font-size:9.8px}}
  #regtbl{{min-width:980px;font-size:10px}}
  #regtbl th{{padding:4px 5px!important;font-size:9.5px}}
  #regtbl td{{padding:3px 5px!important;line-height:1.15}}
  .frow input,.frow select{{height:25px;font-size:9.5px}}
  #sbar{{font-size:8.5px;gap:6px}}
}}
@media(max-width:480px){{
  #topbar{{height:31px}}
  .tab-btn{{padding:5px 8px;font-size:9px}}
  .tool-btn{{padding:3px 3px;font-size:8.5px}}
  #toolbar-actions{{grid-template-columns:repeat(4,minmax(0,1fr))}}
  #ltr-quickbar .tool-btn{{flex:1 1 0!important}}
  #projbar-extra{{grid-template-columns:1fr 1fr}}
}}
@media(max-width:900px) and (orientation:landscape){{
  #topbar{{height:28px;padding:0 6px;gap:4px}}
  #topbar .topbar-title-full{{display:none}}
  #topbar .topbar-title-short{{display:inline!important;font-size:12px!important}}
  #topbar .topbar-title-short::after{{content:" Register"}}
  #topbar .tb-btn{{min-height:21px;padding:2px 5px;font-size:8px}}
  #topbar .topbar-user span:last-child{{font-size:7px!important;padding:1px 5px!important;max-width:58px}}
  #projbar{{padding:2px 8px;gap:5px;min-height:0;align-items:center}}
  #projbar img{{max-height:30px;max-width:72px}}
  #projbar-main{{gap:1px;display:grid;grid-template-columns:minmax(0,1fr);flex:1 1 auto}}
  #projbar-primary,#projbar-extra{{grid-template-columns:repeat(3,minmax(0,1fr));gap:1px 8px}}
  #projbar-primary .pf-lbl,#projbar-extra .pf-lbl{{font-size:5.8px;line-height:.9;margin:0}}
  #projbar-primary .pf-val,#projbar-extra .pf-val{{font-size:8px;line-height:.95;-webkit-line-clamp:1}}
  #projbar-primary .pf.primary:last-child{{grid-column:auto;border-bottom:none;padding-bottom:0;margin-bottom:0}}
  #projbar-primary .pf.primary:last-child .pf-val{{font-size:8.5px}}
  #tabsbar{{padding:1px 7px;gap:3px}}
  .tab-btn{{min-height:20px;padding:2px 7px;font-size:8px}}
  .tcnt{{font-size:7px;padding:0 4px}}
  .tab-add{{min-height:20px;padding:1px 7px;font-size:9px}}
  #toolbar{{padding:2px 6px;gap:3px;display:grid;grid-template-columns:minmax(0,1fr) minmax(150px,220px);align-items:center}}
  #toolbar-actions{{grid-template-columns:repeat(8,minmax(0,1fr));gap:3px;width:auto}}
  .tool-btn{{min-height:20px;padding:1px 3px;font-size:7.5px;gap:1px}}
  #srchbox{{min-height:22px;padding-top:1px;padding-bottom:1px;font-size:9px;flex:0 0 auto;width:100%;max-width:none}}
  #ltr-quickbar{{display:flex!important;flex-direction:row!important;padding:2px 6px!important;gap:3px!important}}
  #ltr-quickbar .tool-btn{{min-height:19px!important;font-size:7.5px!important;padding:1px 4px!important;flex:1 1 0!important}}
  #tblwrap{{min-height:50vh;height:calc(100vh - 188px)}}
  #sbar{{padding:1px 8px;font-size:8px;line-height:1.1}}
}}
</style></head><body>

<div id="topbar">
  <span class="topbar-mark" style="font-size:20px">📋</span>
  <span class="topbar-title-full" style="font-weight:700;font-size:14px;letter-spacing:.3px">Document Control Register</span>
  <span class="topbar-title-short" style="font-weight:800;font-size:13px;letter-spacing:.35px">DCR</span>
  
  <div style="flex:1"></div>
  
  <div id="topbar-proj-info" onclick="if(typeof editProject === 'function') editProject()" style="display:flex;align-items:center;gap:24px;cursor:pointer;padding:4px 16px;border-radius:8px;transition:all 0.2s" title="Edit Project Information">
    {projbar_primary}
  </div>

  <div style="flex:1"></div>

  <a href="/" class="tb-btn">📊 Dashboard</a>
  <button class="tb-btn" onclick="toggleDark()" id="darkBtn" title="Toggle dark mode">🌙</button>
  {btns}
  <span style="color:rgba(255,255,255,.45);padding:0 4px">|</span>
  <span class="topbar-user"><span class="topbar-user-name">👤 {uname}</span>
    <span style="background:{rbg};border-radius:3px;padding:1px 7px;font-size:9px;font-weight:700">{rlbl}</span>
  </span>
</div>

<div id="projbar">
  {logo_html}
  <div id="projbar-main">
    {projbar_toggle}
    <div id="projbar-extra" style="display:flex;gap:20px;flex:1 1 auto;overflow-x:auto">{projbar_secondary}</div>
  </div>
  {'<button class="proj-edit-btn" onclick="editProject()" style="margin-left:auto;background:var(--pr);color:#fff;border:none;padding:5px 12px;border-radius:var(--rd);cursor:pointer;font-size:11px;font-family:inherit;flex-shrink:0">✏ Edit</button>' if editable else ''}
</div>

<div id="tabsbar">
  {tabs_html}
  {'<button class="tab-add" onclick="addDocType()" title="Add Type">＋</button>' if editable else ''}
</div>

  <div id="toolbar">
    <div id="toolbar-actions">
      {edit_btns}
      <button id="btn-bulk-dl" class="tool-btn" style="display:none;background:#2563eb;border-color:#1d4ed8;color:#fff" onclick="bulkDownload()">⬇️ Download Selected</button>
      <button class="tool-btn teal" onclick="openM('export-modal')">📥 Export ▾</button>
    <button class="tool-btn teal" onclick="doPrint()">🖨 Print</button>
      {'<button class="tool-btn teal" onclick="openImport()">📤 Import</button>' if editable else ''}
    </div>
    <input type="text" id="srchbox" placeholder="Search..." oninput="doSearch()">
  </div>
  <div id="ltr-quickbar" style="display:none;padding:8px 10px;border-bottom:1px solid var(--bd);background:#f8fafc;gap:6px;align-items:center;flex-wrap:wrap">
    <span style="font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px">Letter Views</span>
    <button class="tool-btn" type="button" data-ltr-view="" onclick="setLTRQuickView('')">All</button>
    <button class="tool-btn" type="button" data-ltr-view="sent" onclick="setLTRQuickView('sent')">Sent Letters</button>
    <button class="tool-btn" type="button" data-ltr-view="received" onclick="setLTRQuickView('received')">Received Letters</button>
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
  <div id="noc-summary" style="display:none;padding:10px 14px;border-top:1px solid var(--bd);background:#f8fafc">
    <div id="noc-summary-inner" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px"></div>
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
  <div class="modal record-modal">
    <div class="mhdr"><span id="rec-title">Add Document</span>
      <button class="xbtn" onclick="closeM('rec-modal')">✕</button></div>
    <div class="mbody"><div class="record-form-shell" id="rec-form"></div></div>
    <div class="mfoot record-modal-actions">
      <button class="btn btn-sc" onclick="closeM('rec-modal')">Cancel</button>
      <button class="btn btn-pr" id="rec-save-btn" onclick="saveRecord()">Save</button>
    </div>
  </div>
</div>

<!-- LETTER THREAD -->
<div class="overlay hidden" id="thread-modal">
  <div class="modal" style="max-width:760px">
    <div class="mhdr"><span id="thread-title">Letter Thread</span>
      <button class="xbtn" onclick="closeM('thread-modal')">✕</button></div>
    <div class="mbody" id="thread-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('thread-modal')">Close</button>
    </div>
  </div>
</div>

<!-- LETTER TIMELINE -->
<div class="overlay hidden" id="timeline-modal">
  <div class="modal" style="max-width:760px">
    <div class="mhdr"><span id="timeline-title">Letter Timeline</span>
      <button class="xbtn" onclick="closeM('timeline-modal')">✕</button></div>
    <div class="mbody" id="timeline-body"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('timeline-modal')">Close</button>
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

<!-- DISTRIBUTION MATRIX MODAL -->
<div class="overlay hidden" id="dist-modal">
  <div class="modal" style="max-width:900px;max-height:90vh">
    <div class="mhdr"><span>📇 Distribution Matrix</span>
      <button class="xbtn" onclick="closeM('dist-modal')">✕</button></div>
    <div class="mbody" id="dist-body" style="overflow-y:auto;max-height:calc(90vh - 130px)"></div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('dist-modal')">Close</button>
    </div>
  </div>
</div>

<!-- EXPORT MODAL -->
<div class="overlay hidden" id="export-modal">
  <div class="modal" style="max-width:500px;">
    <div class="mhdr"><span>📊 Export Center</span>
      <button class="xbtn" onclick="closeM('export-modal')">✕</button></div>
    <div class="mbody" style="padding:20px;">
      <div style="margin-bottom:15px;">
        <label style="display:block;font-weight:600;margin-bottom:5px;font-size:13px;">Format</label>
        <select id="export-format" class="inp" style="width:100%;" onchange="document.getElementById('export-scope-wrap').style.display = this.value === 'excel' ? 'block' : 'none';">
          <option value="excel">📊 Excel (.xlsx)</option>
          <option value="pdf">📄 PDF Executive Summary</option>
        </select>
      </div>
      <div id="export-scope-wrap" style="margin-bottom:15px;">
        <label style="display:block;font-weight:600;margin-bottom:5px;font-size:13px;">Scope</label>
        <select id="export-scope" class="inp" style="width:100%;">
          <option value="all">All Documents</option>
          <option value="current">Current Tab Only</option>
        </select>
      </div>
      <div style="margin-bottom:15px;">
        <label style="display:block;font-weight:600;margin-bottom:5px;font-size:13px;">Time Range</label>
        <select id="export-time" class="inp" style="width:100%;">
          <option value="all">All Time</option>
          <option value="30">Last 30 Days</option>
          <option value="7">Last 7 Days</option>
        </select>
      </div>
      <div style="margin-bottom:15px;">
        <label style="display:block;font-weight:600;margin-bottom:5px;font-size:13px;">Status</label>
        <select id="export-status" class="inp" style="width:100%;">
          <option value="all">All Statuses</option>
          <option value="overdue">Overdue Only</option>
        </select>
      </div>
    </div>
    <div class="mfoot" style="justify-content:space-between;">
      <button class="btn btn-sc" onclick="closeM('export-modal')">Cancel</button>
      <button class="btn btn-pr" style="background:#2563eb;" onclick="executeAdvancedExport()">Export Now 🚀</button>
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
    <div class="mhdr"><span id="dt-modal-title">Add Document Type</span>
      <button class="xbtn" onclick="closeM('dt-modal')">✕</button></div>
    <div class="mbody">
      <input type="hidden" id="dt-edit-id">
      <div class="fgrid">
        <div class="fg"><label>Code</label><input id="dt-code" placeholder="e.g. MS"></div>
        <div class="fg"><label>Name</label><input id="dt-name" placeholder="e.g. Method Statement"></div>
      </div>
      <div class="stitle" style="margin-top:12px">Expected Reply Override</div>
      <div style="font-size:10px;color:var(--mu);margin:-4px 0 8px;padding:6px 8px;background:var(--bg);border-radius:4px">
        Optional day-count override for this document type only. Calculation mode, weekend rule, and holidays still inherit from the project.
      </div>
      <label style="display:flex;align-items:center;gap:8px;font-size:12px;font-weight:700;margin-bottom:8px">
        <input type="checkbox" id="dt-er-use" onchange="toggleDocTypeReplyOverride()">
        Use Custom Reply Days For This Document Type
      </label>
      <div class="fgrid" id="dt-er-fields" style="display:none">
        <div class="fg"><label>REV00 Reply Days</label><input id="dt-er-rev0" type="number" min="0" step="1" value="14"></div>
        <div class="fg"><label>REV&gt;00 Reply Days</label><input id="dt-er-rev" type="number" min="0" step="1" value="7"></div>
      </div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('dt-modal')">Cancel</button>
      <button class="btn btn-pr" id="dt-save-btn" onclick="saveDocType()">Add</button>
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

<!-- PROJECT SETTINGS MODAL -->
<div class="overlay hidden" id="project-settings-modal">
  <div class="modal" style="max-width:600px">
    <div class="mhdr"><span>⚙ Project Settings</span><button class="xbtn" onclick="closeM('project-settings-modal')">✕</button></div>
    <div class="mbody">
      <div class="stitle">🎨 Status Colors</div>
      <div style="font-size:11px;color:var(--mu);margin-bottom:12px">Configure custom background and text colors for document statuses.</div>
      <div id="status-colors-list" style="display:flex;flex-direction:column;gap:8px"></div>
    </div>
    <div class="mfoot">
      <button class="btn btn-sc" onclick="closeM('project-settings-modal')">Cancel</button>
      <button class="btn btn-pr" onclick="saveProjectSettings()">Save Settings</button>
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
const SC_DEFAULTS={{...{sc_json},
  'Approved':['166534','ffffff'],
  'Accepted & to proceed with Part C':['bbf7d0','166534'],
  'Rejected':['fecaca','7f1d1d'],
  'Cancelled':['e5e7eb','475569'],
  'Under Review':['fef9c3','854d0e'],
  'Information Required':['e0e7ff','312e81'],
  'Pending':['fde68a','92400e']
}};
const SC={{...SC_DEFAULTS, ...{custom_sc_json}}};
const PROJ_FIELDS=[['code','Code'],['name','Project Name'],['startDate','Start Date'],['endDate','End Date'],
  ['client','Client'],['landlord','Landlord'],['pmo','PMO'],['mainConsultant','Consultant'],
  ['mepConsultant','MEP'],['contractor','Contractor']];
const DEFAULT_EXPECTED_REPLY_RULE={{
  rev0_reply_days:14,
  rev_reply_days:7,
  calculation_mode:'working_days',
  weekend_mode:'friday_only',
  exclude_official_holidays:true
}};
const state={{tab:null,cols:[],recs:null,visibleRows:[],selectedRowId:null,sortCol:null,sortDir:'asc',filters:{{}},editId:null,revisionDraftActive:false,savingRecord:false,lists:{{}},prItemsCache:{{}}}};

function isPRTab(){{
  const dt=state.dtList?.find(d=>d.id===state.tab)||{{}};
  const code=(dt.code||dt.id||'').toString().toUpperCase();
  const name=(dt.name||'').toString().toLowerCase();
  return code==='PR' || name.includes('requisition') || name.includes('purchase request');
}}

function isNOCTab(){{
  const dt=state.dtList?.find(d=>d.id===state.tab)||{{}};
  const code=(dt.code||dt.id||'').toString().toUpperCase();
  const name=(dt.name||'').toString().toLowerCase();
  return code==='NOC' || name.includes('notice of change');
}}

function isLTRTab(){{
  const dt=state.dtList?.find(d=>d.id===state.tab)||{{}};
  const code=(dt.code||dt.id||'').toString().toUpperCase();
  const name=(dt.name||'').toString().toLowerCase();
  return code==='LTR' || name.includes('letter');
}}

function isHiddenSystemListName(name){{
  const ln=String(name||'').trim().toLowerCase();
  return ['letter_direction','letter_status'].includes(ln);
}}

function projectHasLTR(){{
  return (state.dtList||[]).some(dt=>{{
    const code=String(dt?.code||dt?.id||'').trim().toUpperCase();
    const name=String(dt?.name||'').trim().toLowerCase();
    return code==='LTR' || name.includes('letter') || name.includes('correspondence');
  }});
}}

function ensureLTRProjectLists(){{
  if(!projectHasLTR())return;
  if(!state.lists||typeof state.lists!=='object')state.lists={{}};
  if(!Object.prototype.hasOwnProperty.call(state.lists,'correspondence_parties'))state.lists.correspondence_parties=[];
}}

function isLTRSingleSelectField(col){{
  if(!isLTRTab())return false;
  const key=String(col?.col_key||'').trim().toLowerCase();
  return ['direction','fromparty','toparty'].includes(key);
}}

function isLTRExcludedField(col){{
  if(!isLTRTab())return false;
  const role=getLTRFieldRole(col);
  const key=String(col?.col_key||'').trim().toLowerCase();
  return ['description','status'].includes(role)||['description','status'].includes(key);
}}

function isLTRInternalField(col){{
  if(!isLTRTab())return false;
  const key=String(col?.col_key||col||'').trim().toLowerCase();
  return ['parentletterid'].includes(key);
}}

function normLTRText(v){{
  return String(v||'').trim().toLowerCase().replace(/[^a-z0-9]+/g,'');
}}

function getLTRFieldRole(col){{
  const key=normLTRText(col?.col_key||'');
  const label=normLTRText(col?.label||'');
  if(key==='docno'||label==='letterref')return 'docNo';
  if(key==='title'||key==='subject'||label==='subject')return 'title';
  if(key==='description'||label==='description')return 'description';
  if(key==='filelocation'||label.includes('attachment')||label.includes('filelocation'))return 'fileLocation';
  if(key==='remarks'||label==='remarks')return 'remarks';
  if(key==='direction'||label==='direction')return 'direction';
  if(key==='fromparty'||key==='from'||label==='fromparty'||label==='from')return 'fromParty';
  if(key==='toparty'||key==='to'||label==='toparty'||label==='to')return 'toParty';
  if(key==='issueddate'||label==='issuedate')return 'issuedDate';
  if(key==='receiveddate'||label==='receiveddate')return 'receivedDate';
  if(key==='parentletterid'||label==='parentletterid')return 'parentLetterId';
  if(key==='parentletterref'||key==='responseref'||label==='responseref'||label==='parentletter')return 'parentLetterRef';
  if(key==='status'||label==='status')return 'status';
  return '';
}}

function getLTRRegisterKeys(){{
  return ['docNo','direction','fromParty','toParty','issuedDate','receivedDate','title','parentLetterRef'];
}}

function getLTRFormKeys(){{
  return ['docNo','title','direction','fromParty','toParty','issuedDate','receivedDate','remarks','parentLetterId','parentLetterRef'];
}}

function getLTRColsByRole(cols, roles, visibleOnly=false){{
  const out=[];
  const used=new Set();
  for(const role of roles){{
    const col=(cols||[]).find(c=>{{
      if(visibleOnly&&!c.visible)return false;
      const r=getLTRFieldRole(c);
      if(r!==role||used.has(c.col_key))return false;
      return role!=='parentLetterId';
    }});
    if(col){{
      out.push(col);
      used.add(col.col_key);
    }}
  }}
  return out;
}}

function getLTRColForRole(cols, role){{
  return (cols||[]).find(c=>getLTRFieldRole(c)===role)||null;
}}

function getLTRValue(row, cols, role){{
  const matches=(cols||[]).filter(c=>getLTRFieldRole(c)===role);
  for(const col of matches){{
    const val=row?.[col.col_key];
    if(String(val??'').trim())return val;
  }}
  const first=matches[0];
  return first?row?.[first.col_key]:'';
}}

function formatLTRParentLabel(opt){{
  const doc=String(opt?.doc_no||'').trim();
  const subject=String(opt?.subject||'').trim();
  if(doc&&subject)return `${{doc}} | ${{subject}}`;
  return doc||subject||'Untitled letter';
}}

function formatLTRParentMeta(opt){{
  if(!opt)return 'No parent letter selected';
  const parts=[];
  if(opt.direction)parts.push(opt.direction);
  if(opt.from_party)parts.push('From: '+opt.from_party);
  if(opt.to_party)parts.push('To: '+opt.to_party);
  if(opt.date)parts.push(opt.date);
  return parts.join(' • ')||'No additional details';
}}

function getLTRVisibleColsFallback(cols, roles){{
  const picked=getLTRColsByRole(cols, roles, true);
  return picked.length?picked:(cols||[]).filter(c=>c.visible&&!isLTRInternalField(c)&&!isLTRExcludedField(c));
}}

function getLTRFormColsFallback(cols, roles){{
  const picked=getLTRColsByRole(cols, roles, false);
  return picked.length?picked:(cols||[]).filter(c=>!isLTRInternalField(c)&&!isLTRExcludedField(c));
}}

function updateLTRQuickBar(){{
  const bar=document.getElementById('ltr-quickbar');
  if(!bar)return;
  const on=isLTRTab();
  bar.style.display=on?'flex':'none';
  if(!on)delete state.filters._ltrQuick;
  bar.querySelectorAll('[data-ltr-view]').forEach(btn=>{{
    const active=(state.filters._ltrQuick||'')===btn.dataset.ltrView;
    btn.style.background=active?'var(--pr)':'';
    btn.style.color=active?'#fff':'';
    btn.style.borderColor=active?'var(--pr)':'';
  }});
}}

function setLTRQuickView(view){{
  state.filters._ltrQuick=view||'';
  updateLTRQuickBar();
  renderRows();
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
  return detailsKey?String(row?.[detailsKey]||'').replace(/\\s+/g,' ').trim():'';
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

// _WHOAMI: must be declared BEFORE the IIFE that calls loadDTs()
let _WHOAMI = null;

// Init
(async()=>{{
  await Promise.all([loadDTs(), loadLists()]);
  updateClock(); setInterval(updateClock,60000);
}})();

function updateClock(){{document.getElementById('s-clock').textContent=new Date().toLocaleString('en-GB');}}
async function loadDTs(keepTab=false){{
  // Fetch whoami (for DC button) and doc types in parallel — no sequential delay
  const [whoamiResult, dts] = await Promise.all([
    _WHOAMI ? Promise.resolve(_WHOAMI) : apiFetch('/api/whoami').catch(()=>null),
    apiFetch('/api/doc_types/'+PID)
  ]);
  if(!_WHOAMI) _WHOAMI = whoamiResult;
  if(!dts)return;
  renderTabs(dts);
  ensureLTRProjectLists();
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
  ensureLTRProjectLists();
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
  state.tab=id;state.recs=null;state.visibleRows=[];state.selectedRowId=null;state.filters={{}};state.sortCol=null;
  document.getElementById('srchbox').value='';
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active',b.dataset.id===id));
  updateLTRQuickBar();
  loadRecords();
}}

function isMobileRegister(){{
  return window.matchMedia&&window.matchMedia('(max-width: 768px)').matches;
}}

function mobileColumnWidth(col){{
  if(!isMobileRegister()||!col)return 0;
  const key=String(col.col_key||'').toLowerCase();
  const label=String(col.label||'').toLowerCase();
  const text=(key+' '+label).replace(/[_./-]+/g,' ');
  const role=isLTRTab()?getLTRFieldRole(col):'';
  if(role==='title'||role==='subject'||role==='description'||text.includes('subject'))return 230;
  if(role==='fromParty'||role==='toParty')return 112;
  if(role==='parentLetterRef'||role==='parentLetterId')return 138;
  if(role==='direction')return 82;
  if(role==='issuedDate'||role==='receivedDate')return 106;
  if(text.includes('document no')||key==='docno'||key==='doc_no')return 152;
  if(text.includes('title')||text.includes('description')||text.includes('scope'))return 220;
  if(text.includes('remarks')||text.includes('comment')||text.includes('note'))return 180;
  if(text.includes('prepared')||text.includes('company')||text.includes('client')||text.includes('consultant')||text.includes('contractor')||text.includes('person'))return 132;
  if(text.includes('from')||text.includes(' to ')||text==='to'||key==='to'||key==='from')return 112;
  if(text.includes('ms ref')||text.includes('dwg')||text.includes('item ref')||text.includes('reference'))return 150;
  if(text.includes('discipline'))return 112;
  if(text.includes('sub trade')||text.includes('sub-trade'))return 122;
  if(text.includes('status'))return 118;
  if(text.includes('date'))return 104;
  if(text.includes('duration'))return 72;
  if(text.includes('file')||text.includes('link'))return 82;
  if(text.includes('floor')||text.includes('location'))return 116;
  if(text.includes('revision')||text.includes('rev'))return 82;
  if(text.includes('qty')||text.includes('quantity')||text.includes('amount'))return 88;
  return 108;
}}

function mobileColumnStyle(col){{
  const w=mobileColumnWidth(col);
  return w?`width:${{w}}px;min-width:${{w}}px;`: '';
}}

function tabMenu(id,e){{
  const old=document.getElementById('tabctx');if(old)old.remove();
  const m=document.createElement('div');m.id='tabctx';
  m.style.cssText=`position:fixed;left:${{e.clientX}}px;top:${{e.clientY}}px;background:#fff;
    border:1px solid var(--bd);border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.15);z-index:9999;min-width:140px;overflow:hidden`;
  const dt=state.dtList?.find(d=>d.id===id)||{{}};
  const idx=(state.dtList||[]).findIndex(d=>d.id===id);
  m.innerHTML=`
    <div onclick="renameDT('${{id}}','${{dt.code||id}}','${{(dt.name||'').replace(/'/g,\"\\\\'\")}}')"
      style="padding:8px 14px;cursor:pointer;font-size:12px"
      onmouseover="this.style.background='#f0f4f8'" onmouseout="this.style.background=''">✏ Type Settings</div>
    <div onclick="moveDT('${{id}}',-1)" style="padding:8px 14px;cursor:${{idx>0?'pointer':'not-allowed'}};font-size:12px;opacity:${{idx>0?'1':'.45'}}"
      onmouseover="if(${{idx>0}})this.style.background='#f0f4f8'" onmouseout="this.style.background=''">⬅ Move Left</div>
    <div onclick="moveDT('${{id}}',1)" style="padding:8px 14px;cursor:${{idx>=0&&idx<(state.dtList||[]).length-1?'pointer':'not-allowed'}};font-size:12px;opacity:${{idx>=0&&idx<(state.dtList||[]).length-1?'1':'.45'}}"
      onmouseover="if(${{idx>=0&&idx<(state.dtList||[]).length-1}})this.style.background='#f0f4f8'" onmouseout="this.style.background=''">➡ Move Right</div>
    <div onclick="delDT('${{id}}')" style="padding:8px 14px;cursor:pointer;font-size:12px;color:#ef4444"
      onmouseover="this.style.background='#fef2f2'" onmouseout="this.style.background=''">🗑 Delete Type</div>`;
  document.body.appendChild(m);
  setTimeout(()=>document.addEventListener('click',()=>m.remove(),{{once:true}}),10);
}}

function renameDT(id,oldCode,oldName){{
  const dt=(state.dtList||[]).find(d=>d.id===id)||{{id,code:oldCode,name:oldName}};
  setDocTypeModalMode(dt);
  openM('dt-modal');
}}

async function delDT(id){{
  if(!confirm('Delete type and ALL its records?'))return;
  await apiFetch('/api/doc_types/'+PID+'/'+id,{{method:'DELETE'}});
  if(state.tab===id)state.tab=null;
  await loadDTs(); toast('Deleted','wa');
}}

async function moveDT(id,delta){{
  const order=(state.dtList||[]).map(d=>d.id);
  const idx=order.indexOf(id);
  const next=idx+delta;
  if(idx<0||next<0||next>=order.length)return;
  [order[idx],order[next]]=[order[next],order[idx]];
  await apiFetch('/api/doc_types/'+PID+'/reorder',{{method:'POST',body:JSON.stringify({{order}})}});
  await loadDTs(true);
  toast('✔ Tab order updated','ok');
}}

async function loadRecords(){{
  if(!state.tab)return;
  const search=document.getElementById('srchbox').value.trim();
  const [data, widths]=await Promise.all([
    apiFetch('/api/records/'+PID+'/'+state.tab+(search?'?search='+encodeURIComponent(search):'')),
    apiFetch('/api/col_width/'+PID+'/'+state.tab)
  ]);
  if(!data)return;
  state.recs=data.records; state.allTabCols=data.columns||[]; state.cols=data.columns.filter(c=>c.visible);
  if(isLTRTab())state.cols=state.cols.filter(c=>!isLTRInternalField(c)&&!isLTRExcludedField(c));
  const drvIdx = state.cols.findIndex(c => c.col_type === 'url' || c.col_key === 'fileLocation');
  if(drvIdx > -1) {{
    const drvCol = state.cols.splice(drvIdx, 1)[0];
    const targetIdx = state.cols.findIndex(c => c.col_key === 'docNo' || String(c.label).toLowerCase().includes('doc. no'));
    if(targetIdx > -1) state.cols.splice(targetIdx + 1, 0, drvCol);
    else state.cols.unshift(drvCol);
  }}
  if(!state.recs.some(r=>String(r._id)===String(state.selectedRowId)))state.selectedRowId=null;
  state.prItemsCache=data.pr_items_map||{{}};
  state.colWidths=widths||{{}};
  const cnt=document.getElementById('cnt-'+state.tab); if(cnt)cnt.textContent=data.count;
  updateLTRQuickBar();
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
  sr.style.cssText=(isMobileRegister()?'width:34px;min-width:34px;max-width:34px;':(_srW?'width:'+_srW+'px;min-width:'+_srW+'px;max-width:'+_srW+'px;':'width:40px;min-width:40px;'))
    +'white-space:nowrap;cursor:default';
  hr.appendChild(sr);
  state.cols.forEach(col=>{{
    const th=document.createElement('th');th.dataset.key=col.col_key;
    const w=state.colWidths&&state.colWidths[col.col_key];
    const labelTxt=String(col.label||'').toLowerCase();
    const minW=labelTxt.includes('sub-trade')?150:
      labelTxt.includes('discipline')?130:
      labelTxt.includes('title')?220:
      labelTxt.includes('subject')?220:0;
    const mw=mobileColumnWidth(col);
    if(mw)th.style.cssText='width:'+mw+'px;min-width:'+mw+'px;white-space:normal;word-break:normal';
    else if(w)th.style.cssText='width:'+w+'px;min-width:'+w+'px;max-width:'+w+'px;white-space:normal;word-break:normal';
    else th.style.cssText=(minW?'min-width:'+minW+'px;':'')+'white-space:normal;word-break:normal';
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
    const mw=mobileColumnWidth(col);
    if(mw)th.style.cssText='width:'+mw+'px;min-width:'+mw+'px';
    if(['auto_date','auto_num'].includes(col.col_type)){{fr.appendChild(th);return;}}
    if(col.col_type==='dropdown'&&col.list_name){{
      const sel=document.createElement('select');
      const ltrRole=isLTRTab()?getLTRFieldRole(col):'';
      const listName=ltrRole==='fromParty'||ltrRole==='toParty'
        ? 'correspondence_parties'
        : (ltrRole==='direction' ? 'letter_direction' : (ltrRole==='status' ? 'letter_status' : col.list_name));
      sel.innerHTML='<option value="">All</option>'+(state.lists[listName]||[]).map(o=>`<option ${{state.filters[col.col_key]===o?'selected':''}}>${{o}}</option>`).join('');
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
  const raw=String(docNo||'').trim();
  const revMatch=raw.match(/\\bREV\\s*([0-9]+)\\b/i);
  const rev=revMatch?parseInt(revMatch[1],10):0;
  const hasRev=!!revMatch;
  const withoutRev=raw.replace(/\\s*\\bREV\\s*[0-9]+\\b\\s*$/i,'').trim();
  const m=withoutRev.match(/^(.*?)-(\\d+)$/);
  if(m)return{{prefix:m[1],num:parseInt(m[2],10),width:m[2].length,rev,base:withoutRev,raw,hasRev}};
  return{{prefix:raw,num:0,width:3,rev,base:withoutRev||raw,raw,hasRev}};
}}

function buildDocNoFromParts(p,num,rev){{
  let res = `${{p.prefix}}-${{String(num).padStart(p.width||3,'0')}}`;
  if(p.hasRev) res += ` REV${{String(rev).padStart(2,'0')}}`;
  return res;
}}

function incrementRevisionNumber(docNo){{
  const raw=String(docNo||'').trim();
  const m=raw.match(/\\bREV\\s*([0-9]+)\\b/i);
  if(!m)return '';
  const next=String(parseInt(m[1],10)+1).padStart(m[1].length,'0');
  return raw.replace(/\\bREV\\s*[0-9]+\\b/i,'REV'+next);
}}

function suggestNextDocNoFromRows(rows){{
  const parsed=(rows||[]).map((r,idx)=>({{...parseDocNo(r.docNo||''),_visibleIdx:idx}})).filter(p=>p&&p.num>0);
  if(!parsed.length)return '';
  const families=new Map();
  parsed.forEach(p=>{{
    const key=String(p.prefix||'').toLowerCase();
    if(!families.has(key))families.set(key,[]);
    families.get(key).push(p);
  }});
  const lastVisible=parsed.reduce((a,b)=>b._visibleIdx>a._visibleIdx?b:a,parsed[0]);
  const family=families.get(String(lastVisible.prefix||'').toLowerCase())||[];
  if(!family.length)return '';
  family.sort((a,b)=>a.num-b.num);
  const base=family[family.length-1];
  return buildDocNoFromParts(base,base.num+1,0);
}}

function suggestNextDocNoFromVisibleRows(){{
  return suggestNextDocNoFromRows(state.visibleRows||[]);
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

function getSelectedRegisterRow(){{
  if(!state.selectedRowId)return null;
  return (state.recs||[]).find(r=>String(r._id)===String(state.selectedRowId))||null;
}}

function setSelectedRegisterRow(id){{
  state.selectedRowId=id||null;
  document.querySelectorAll('#regtbl tbody tr.row-selected').forEach(tr=>tr.classList.remove('row-selected'));
  if(state.selectedRowId){{
    const tr=document.querySelector(`#regtbl tbody tr[data-rec-id="${{CSS.escape(String(state.selectedRowId))}}"]`);
    if(tr)tr.classList.add('row-selected');
  }}
}}

function isInteractiveRowTarget(target){{
  return !!target.closest('button,a,input,select,textarea,label,.pr-items-row');
}}

function getCheckedRecordIds(){{
  return [...document.querySelectorAll('.chkcell input[data-id]:checked')].map(cb=>cb.dataset.id);
}}

function getCheckedRecordRows(){{
  const ids=new Set(getCheckedRecordIds().map(String));
  return (state.recs||[]).filter(r=>ids.has(String(r._id)));
}}

function syncCheckboxRowHighlights(){{
  document.querySelectorAll('#regtbl tbody tr[data-rec-id]').forEach(tr=>{{
    const cb=tr.querySelector('.chkcell input[data-id]');
    tr.classList.toggle('row-selected',!!cb?.checked);
  }});
}}

function isVolatileRevisionDraftField(col){{
  const key=String(col?.col_key||'').trim().toLowerCase();
  const label=String(col?.label||'').trim().toLowerCase();
  const type=String(col?.col_type||'').trim().toLowerCase();
  const text=(key+' '+label).replace(/[_-]+/g,' ');
  if(!key||key==='docno')return false;
  if(type==='auto_date'||type==='auto_num'||type==='duration_calc')return true;
  if(type==='date')return true;
  const volatileTerms=[
    'issued','issue date','submitted','submission','submittal','received','receive date',
    'expected reply','actual reply','reply date','reply','response','status','duration',
    'remark','comment','file location','attachment','link','transmittal','return date',
    'approval','approved','review','code','reference no','reply reference','parent letter'
  ];
  return volatileTerms.some(term=>text.includes(term));
}}

function sanitizeRevisionDraftData(oldRecord,newDocNo){{
  const draft={{}};
  const cols=state.allTabCols||state.cols||[];
  cols.forEach(col=>{{
    const key=col.col_key;
    if(!key||key.startsWith('_'))return;
    draft[key]=isVolatileRevisionDraftField(col)?'':(oldRecord?.[key]??'');
  }});
  draft.docNo=newDocNo;
  draft._cloneSourceId=oldRecord?._id;
  return draft;
}}

function renderRows(){{
  try {{
  const body=document.getElementById('tbody');body.innerHTML='';
  const isPrTab=isPRTab();
  const isLtrTab=isLTRTab();
  const prDetailsKey=isPrTab?getPrDetailsColKey():null;
  let rows=state.recs.filter(r=>{{
    for(const[k,v]of Object.entries(state.filters)){{
      if(k.startsWith('_'))continue;
      if(v&&!String(r[k]||'').toLowerCase().includes(v.toLowerCase()))return false;
    }}
    if(isLtrTab){{
      const quick=String(state.filters._ltrQuick||'').toLowerCase();
      const direction=String(getLTRValue(r,state.allTabCols,'direction')||'').toLowerCase();
      if(quick==='sent'&&direction!=='sent')return false;
      if(quick==='received'&&direction!=='received')return false;
    }}
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
  state.visibleRows=rows;
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
    renderNocSummary(rows);
    return;
  }}
  let sr=1;
  rows.forEach((row,idx)=>{{
    const tr=document.createElement('tr');
    tr.dataset.recId=row._id;
    if(row._overdue)tr.classList.add('ov');
    else if(row._isRev)tr.classList.add('rv');
    else if(idx%2===1)tr.classList.add('alt');
    const tc=document.createElement('td');tc.className='chkcell';
    if(CAN_EDIT){{const cb=document.createElement('input');cb.type='checkbox';cb.dataset.id=row._id;cb.onchange=updBulk;tc.appendChild(cb);}}
    tr.appendChild(tc);
    const tsr=document.createElement('td');tsr.className='sr';
    const _srTdW=state.colWidths&&state.colWidths['_sr'];
    if(isMobileRegister())tsr.style.cssText='width:34px;min-width:34px;max-width:34px;white-space:nowrap;';
    else if(_srTdW)tsr.style.cssText='width:'+_srTdW+'px;min-width:'+_srTdW+'px;max-width:'+_srTdW+'px;white-space:nowrap;';
    tsr.textContent=row._isRev?'':sr;tr.appendChild(tsr);
    state.cols.forEach(col=>{{
      const td=document.createElement('td');const key=col.col_key||'';let val='';
      const mcs=mobileColumnStyle(col);
      if(mcs)td.style.cssText=mcs;
      const ltrRole=isLtrTab?getLTRFieldRole(col):'';
      
      if(col.col_type==='url' || key==='fileLocation'){{
        const url=String(row[key]||'').trim();
        const hasValidLink = url && url !== '#' && url.toLowerCase() !== 'null' && url.toLowerCase() !== 'none';
        if(hasValidLink){{td.innerHTML=`<a class="flink" href="${{url}}" target="_blank" title="View Document" style="text-decoration:none;font-size:15px;color:#0ea5e9;">👁️</a>`;}}
        else {{td.innerHTML=`<span title="No Document Attached" style="font-size:15px;color:#cbd5e1;cursor:not-allowed;opacity:0.5;">👁️</span>`;}}
        td.style.textAlign='center';
        tr.appendChild(td);return;
      }}
      
      const k=key.toLowerCase();
      const longTextMeta=getLongTextMeta(col);
      if(longTextMeta||isFloorField(col))td.classList.add('mlcell');
      if(col.col_type==='docno'||key==='docNo')td.classList.add('docno-cell');
      if(key==='expectedReplyDate'){{val=row._expectedReplyDate||'';if(row._overdue&&val)td.classList.add('ovdate');}}
      else if(key==='duration')val=row._duration||'';
      else if(col.col_type==='duration_calc'){{
        const[ds,de]=String(col.list_name||'issuedDate,actualReplyDate').split(',');
        val=calcWD(row[ds.trim()]||'',row[de.trim()]||'');
      }}
      else if(key==='issuedDate')val=row._issuedFmt||'';
      else if(key==='actualReplyDate')val=row._replyFmt||'';
      else if(col.col_type==='date'||col.col_type==='auto_date'){{
        val=row['_fmt_'+key]||row[key]||'';  // use pre-formatted version
      }}
      else if(key==='status'||isStatusLikeField(col)){{
        val=row[key]||((isLtrTab&&ltrRole)?getLTRValue(row,state.allTabCols,ltrRole):'');
        if(val){{
          td.innerHTML=String(val).split(',').map(s=>{{
            s=s.trim();
            let [bg,fg]=SC[s]||['e5e7eb','374151'];
            if(!SC[s]){{
              const up=s.toUpperCase();
              if(up.includes('CANCELLED')) [bg,fg]=['6c757d','ffffff'];
              else if(/^A(-|\\s|$)/.test(up) || up.includes('APPROVED')) [bg,fg]=['bbf7d0','166534'];
              else if(/^B(-|\\s|$)/.test(up) || up.includes('NOTED')) [bg,fg]=['dcfce7','166534'];
              else if(/^C(-|\\s|$)/.test(up) || up.includes('REVISE')) [bg,fg]=['fed7aa','9a3412'];
              else if(/^D(-|\\s|$)/.test(up) || up.includes('REJECTED')) [bg,fg]=['fecaca','7f1d1d'];
            }}
            return `<span class="sbadge" style="background:#${{bg}};color:#${{fg}}">${{s}}</span>`;
          }}).join('');
          tr.appendChild(td);return;
        }}
      }}
      else if(col.col_type==='url' || key==='fileLocation'){{
        const url=row[key]||'';
        if(url){{td.innerHTML=`<a class="flink" href="${{url}}" target="_blank">View File</a>`;tr.appendChild(td);return;}}
      }}
      else if(isPrTab&&prDetailsKey&&key===prDetailsKey)val=getPrSummary(row);
      else val=String(row[key]||'');
      if(isLtrTab&&ltrRole&&!String(val||'').trim())val=String(getLTRValue(row,state.allTabCols,ltrRole)||'');
      let displayVal=formatDisplayValue(col,val);
      if(typeof displayVal==='string'&&longTextMeta){{
        const NL=String.fromCharCode(10);
        displayVal=displayVal
          .replaceAll(' / ',NL)
          .replaceAll(' /',NL)
          .replaceAll('/ ',NL);
      }}
      td.textContent=displayVal;
      applyDisplayDirection(td,displayVal);
      tr.appendChild(td);
    }});
    const ta=document.createElement('td');ta.className='acts';
    let acts='';
    if(isPrTab)acts+=`<button class="pr-toggle" onclick="togglePrItems('${{row._id}}', this)">Items</button> `;
    if(isLtrTab)acts+=`<button class="pr-toggle" onclick="openLetterThread('${{row._id}}')">Thread</button> <button class="pr-toggle" onclick="openLetterTimeline('${{row._id}}')">Timeline</button> `;
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
  renderNocSummary(rows);
  if(isPrTab && rows.length > 0) {{
    const searchTxt = document.getElementById('srchbox')?.value.trim().toLowerCase();
    if(searchTxt) {{
      setTimeout(() => {{
        rows.forEach(r => {{
          const items = state.prItemsCache[r._id] || [];
          const hasMatch = items.some(it => 
            String(it.item_name||'').toLowerCase().includes(searchTxt) || 
            String(it.remarks||'').toLowerCase().includes(searchTxt) ||
            String(it.item_code||'').toLowerCase().includes(searchTxt)
          );
          if(hasMatch) {{
            const rowTr = document.querySelector(`tr[data-rec-id="${{CSS.escape(r._id)}}"]`);
            if(rowTr) {{
              const btn = rowTr.querySelector('.pr-toggle');
              if(btn && btn.textContent === 'Items') {{
                btn.click();
              }}
            }}
          }}
        }});
      }}, 50);
    }}
  }}
  }} catch(err) {{
    const errStr = 'JS ERROR: ' + err.message;
    document.getElementById('s-total').textContent=errStr;
    document.getElementById('s-total').style.color='red';
    document.getElementById('s-total').style.fontWeight='bold';
  }}
}}

function calcWD(s,e){{
  if(!s||!e)return'';
  try{{let a=new Date(s),b=new Date(e);if(isNaN(a)||isNaN(b)||b<=a)return'0';
    let n=0,c=new Date(a);c.setDate(c.getDate()+1);
    while(c<=b){{if(c.getDay()!==5)n++;c.setDate(c.getDate()+1);}}return String(n);}}
  catch{{return'';}}
}}
function getLongTextMeta(col){{
  if(col.col_type==='long_text')return {{type:'long_text',rows:col.rows||3}};
  return null;
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

function isStatusLikeField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  return key==='status'||key.includes('status')||label.includes('status');
}}

async function openLetterThread(id){{
  const data=await apiFetch('/api/letters/thread/'+PID+'/'+id);if(!data)return;
  renderLetterThread(data);
  openM('thread-modal');
}}

async function openLetterTimeline(id){{
  const data=await apiFetch('/api/letters/timeline/'+PID+'/'+id);if(!data)return;
  renderLetterTimeline(data);
  openM('timeline-modal');
}}

function renderLetterThread(data){{
  const body=document.getElementById('thread-body');
  const title=document.getElementById('thread-title');
  const items=Array.isArray(data?.items)?data.items:[];
  const currentId=String(data?.current_id||'').trim();
  const current=items.find(it=>it.id===currentId)||items[0]||null;
  title.textContent=current?.doc_no?('Letter Thread — '+current.doc_no):'Letter Thread';
  if(!items.length){{
    body.innerHTML='<div style="padding:12px;color:var(--mu);font-size:12px">This letter has no linked thread yet.</div>';
    return;
  }}
  body.innerHTML=items.map(it=>{{
    const isCurrent=it.id===currentId;
    const margin=it.level*42;
    const border=isCurrent?'#8BC34A':'#d7e0e6';
    const bg=isCurrent?'#f6fbef':'#fff';
    const shadow=isCurrent?'0 10px 26px rgba(139,195,74,.14)':'0 6px 18px rgba(15,23,42,.05)';
    const titleText=escHtml(it.subject||'');
    const refText=escHtml(it.doc_no||'Untitled');
    const metaParts=[
      it.direction?escHtml(it.direction):'',
      it.from_party?('From: '+escHtml(it.from_party)):'',
      it.to_party?('To: '+escHtml(it.to_party)):'',
      it.date?escHtml(it.date):''
    ].filter(Boolean);
    const currentBadge=isCurrent?'<span class="letter-current-badge">Current</span>':'';
    const subjectHtml=titleText||'<span style="color:var(--mu)">No subject</span>';
    const branchCue=it.level>0?'<span style="color:#7d8a95;font-weight:800;margin-right:8px">↳</span>':'';
    const levelLabel=it.level>0?('Reply level '+it.level):'Root letter';
    const metaHtml=metaParts.length?'<div style="font-size:10px;color:var(--mu);margin-top:8px;line-height:1.5">'+metaParts.join(' • ')+'</div>':'';
    return `<div style="margin-left:${{margin}}px;margin-bottom:10px;position:relative">
      ${{it.level>0?`<div style="position:absolute;left:-24px;top:0;bottom:22px;width:2px;background:${{connectorColor}};border-radius:2px"></div>
      <div style="position:absolute;left:-24px;top:24px;width:18px;height:2px;background:${{connectorColor}};border-radius:2px"></div>`:''}}
      <div style="background:${{bg}};border:1px solid ${{border}};border-left:6px solid ${{border}};border-radius:12px;padding:12px 14px;box-shadow:${{shadow}}">
        <div style="font-size:9px;color:#7d8a95;text-transform:uppercase;letter-spacing:.35px;margin-bottom:6px">${{levelLabel}}</div>
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start">
          <div style="min-width:0">
            <div style="font-size:12px;font-weight:800;color:var(--pr);display:flex;align-items:center" dir="auto">${{branchCue}}<span>${{refText}}</span></div>
            <div style="font-size:12px;color:var(--tx);margin-top:4px;line-height:1.45" dir="auto">${{subjectHtml}}</div>
          </div>
          ${{currentBadge}}
        </div>
        ${{metaHtml}}
      </div>
    </div>`;
  }}).join('');
}}

function renderLetterTimeline(data){{
  const body=document.getElementById('timeline-body');
  const title=document.getElementById('timeline-title');
  const items=Array.isArray(data?.items)?data.items:[];
  const currentId=String(data?.current_id||'').trim();
  const current=items.find(it=>it.id===currentId)||items[0]||null;
  title.textContent=current?.doc_no?('Letter Timeline — '+current.doc_no):'Letter Timeline';
  if(!items.length){{
    body.innerHTML='<div style="padding:12px;color:var(--mu);font-size:12px">This letter has no linked timeline yet.</div>';
    return;
  }}
  const cards=items.map(it=>{{
    const isCurrent=String(it.id||'')===currentId;
    const border=isCurrent?'#8BC34A':'#d7e0e6';
    const bg=isCurrent?'#f6fbef':'#fff';
    const shadow=isCurrent?'0 10px 26px rgba(139,195,74,.14)':'0 6px 18px rgba(15,23,42,.05)';
    const refText=escHtml(it.doc_no||'Untitled');
    const subjectHtml=escHtml(it.subject||'')||'<span style="color:var(--mu)">No subject</span>';
    const dateText=escHtml(it.date||it.created_at||'');
    const metaParts=[
      it.direction?escHtml(it.direction):'',
      it.from_party?('From: '+escHtml(it.from_party)):'',
      it.to_party?('To: '+escHtml(it.to_party)):''
    ].filter(Boolean);
    const metaHtml=metaParts.length?'<div style="font-size:10px;color:var(--mu);margin-top:8px;line-height:1.5">'+metaParts.join(' • ')+'</div>':'';
    const badge=isCurrent?'<span style="flex-shrink:0;background:#8BC34A;color:#17351b;border-radius:999px;padding:3px 8px;font-size:9px;font-weight:800;letter-spacing:.35px;text-transform:uppercase">Current</span>':'';
    const dateBadge=dateText?'<div style="display:inline-flex;align-items:center;background:#eef4f8;color:#2F4F64;border-radius:999px;padding:4px 10px;font-size:10px;font-weight:800;letter-spacing:.2px">'+dateText+'</div>':'';
    return `<div style="position:relative;padding-left:30px;margin-bottom:14px">
      <div style="position:absolute;left:8px;top:14px;width:12px;height:12px;border-radius:999px;background:${{isCurrent?'#8BC34A':'#2F4F64'}};border:3px solid #fff;box-shadow:0 0 0 2px ${{isCurrent?'rgba(139,195,74,.35)':'rgba(47,79,100,.18)'}}"></div>
      <div style="background:${{bg}};border:1px solid ${{border}};border-left:4px solid ${{border}};border-radius:12px;padding:12px 14px;box-shadow:${{shadow}}">
        <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start">
          <div style="min-width:0">
            <div style="font-size:12px;font-weight:800;color:var(--pr)" dir="auto">${{refText}}</div>
            <div style="margin-top:5px">${{dateBadge}}</div>
            <div style="font-size:12px;color:var(--tx);margin-top:6px;line-height:1.45" dir="auto">${{subjectHtml}}</div>
          </div>
          ${{badge}}
        </div>
        ${{metaHtml}}
      </div>
    </div>`;
  }}).join('');
  body.innerHTML=`<div style="position:relative;padding-left:10px">
    <div style="position:absolute;left:13px;top:0;bottom:0;width:2px;background:linear-gradient(180deg,#d7e0e6,#c7d2da)"></div>
    ${{cards}}
  </div>`;
}}

function isCurrencyField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  return key.includes('cost')||key.includes('value')||label.includes('(egp)');
}}

function formatCurrencyValue(val){{
  const raw=String(val??'').trim();
  if(!raw)return '';
  const num=Number(raw.toString().replace(/,/g,''));
  if(Number.isNaN(num))return raw;
  return new Intl.NumberFormat('en-US',{{minimumFractionDigits:2,maximumFractionDigits:2}}).format(num)+' EGP';
}}

function parseSafeNumber(val){{
  const raw=String(val??'').trim();
  if(!raw)return null;
  const num=Number(raw.replace(/,/g,''));
  return Number.isFinite(num)?num:null;
}}

function getNocStageProgress(data){{
  const d=data||{{}};
  if(d.partDReturnDate||d.partDStatus||d.finalApprovedCost||d.voNo||d.voIssueDate||d.voValueWithSIAndVAT)return 'Stage D';
  if(d.partCIssueDate||d.submittedCost)return 'Stage C';
  if(d.partBReturnDate||d.partBStatus)return 'Stage B';
  if(d.partAIssueDate||d.docNo||d.title||d.nocDescription||d.originatingDocument)return 'Stage A';
  return 'Not Started';
}}

function getNocAutoBaseValue(data){{
  const finalVal=String(data?.voValueWithSIAndVAT??'').trim();
  if(!finalVal)return '';
  const num=Number(finalVal.replace(/,/g,''));
  if(Number.isNaN(num))return '';
  return (num/1.05).toFixed(2);
}}

function formatDisplayValue(col,val){{
  let text=String(val??'');
  if(!text)return '';
  const NL=String.fromCharCode(10);
  if(isFloorField(col))return text.split(',').map(s=>s.trim()).filter(Boolean).join(NL);
  if(isItemRefField(col))return text.replace(/\\r\\n/g,NL).replace(/\\r/g,NL);
  if(isCurrencyField(col))return formatCurrencyValue(text);
  return text;
}}

function getTextDirectionMeta(val){{
  const text=String(val??'').trim();
  if(!text)return null;
  const arabic=(text.match(/[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]/g)||[]).length;
  const latin=(text.match(/[A-Za-z]/g)||[]).length;
  if(arabic&&arabic>=latin)return {{dir:'rtl',align:'right'}};
  return null;
}}

function applyDisplayDirection(el,val){{
  const meta=getTextDirectionMeta(val);
  el.classList.remove('rtl-txt');
  el.style.direction='';
  el.style.textAlign='';
  if(meta){{
    el.classList.add('rtl-txt');
    el.style.direction=meta.dir;
    el.style.textAlign=meta.align;
  }}
}}

function bindDirectionalInput(el){{
  if(!el)return;
  const sync=()=>applyDisplayDirection(el,el.value);
  el.addEventListener('input',sync);
  el.addEventListener('change',sync);
  sync();
}}

function bindAutoGrowTextarea(ta){{
  if(!ta||ta.tagName!=='TEXTAREA')return;
  const grow=()=>{{
    ta.style.height='auto';
    ta.style.height=Math.min(Math.max(ta.scrollHeight,42),220)+'px';
  }};
  ta.addEventListener('input',grow);
  ta.addEventListener('change',grow);
  requestAnimationFrame(grow);
}}

function bindSmartTextarea(ta){{
  bindDirectionalInput(ta);
  bindAutoGrowTextarea(ta);
}}

function isCompactCoreFormField(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  const text=(key+' '+label).replace(/[_-]+/g,' ').replace(/\\s+/g,' ').trim();
  return ['brand','origin','originator','prepared by','preparedby','discipline','sub trade','sub-trade','floor','document no','docno','letter ref','letterref'].some(v=>text.includes(v));
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

function highlightText(text, search) {{
  if (!search) return escHtml(text || '');
  const str = String(text || '');
  const escapedSearch = search.replace(/[.*+?^${{}}()|[\\]\\\\]/g, '\\\\$&');
  const regex = new RegExp(`(${{escapedSearch}})`, 'gi');
  const parts = str.split(regex);
  return parts.map((p, i) => i % 2 === 1 ? `<mark style="background:#fef08a;color:#854d0e;border-radius:2px;padding:0 2px">${{escHtml(p)}}</mark>` : escHtml(p)).join('');
}}

function renderPrItemsTable(items, legacyText, search){{
  if(!items||!items.length){{
    const legacy=legacyText?`<div style="margin-top:8px"><div style="font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px">Legacy PR Details</div><div style="color:var(--mu);font-size:11px;white-space:pre-line">${{escHtml(legacyText)}}</div></div>`:'';
    return `<div class="pr-items-title">PR Items Breakdown</div><div class="pr-items-empty">No items added</div>${{legacy}}`;
  }}
  const rows=items.map(it=>String(it?.row_type||'item').toLowerCase()==='header'
    ? `<div class="pr-items-section">${{highlightText(it.item_name||'', search)}}</div>`
    : `<div class="pr-items-cell item">${{highlightText(it.item_name||'', search)}}</div>
       <div class="pr-items-cell unit">${{highlightText(it.unit||'', search)}}</div>
       <div class="pr-items-cell qty">${{escHtml(it.quantity??'')}}</div>
       <div class="pr-items-cell remarks">${{highlightText(it.remarks||'', search)}}</div>`).join('');
  return `<div class="pr-items-title">PR Items Breakdown</div>
    <div class="pr-items-grid-wrap">
      <div class="pr-items-grid">
        <div class="pr-items-grid-head">Item</div>
        <div class="pr-items-grid-head">Unit</div>
        <div class="pr-items-grid-head">Qty</div>
        <div class="pr-items-grid-head">Remarks</div>
        ${{rows}}
      </div>
    </div>`;
}}

function getRenderedRegisterColumnCount(tr){{
  const headRow=document.querySelector('#regtbl thead tr:first-child');
  const headCount=headRow?headRow.children.length:0;
  const rowCount=tr?tr.children.length:0;
  return Math.max(headCount,rowCount,state.cols.length+3);
}}

function sizePrItemsPanel(panel){{
  const wrap=document.getElementById('tblwrap');
  if(!panel||!wrap)return;
  panel.style.width=Math.max(wrap.clientWidth,320)+'px';
}}

async function togglePrItems(recordId, btn){{
  const tr=btn.closest('tr');
  if(!tr)return;
  let nxt=tr.nextElementSibling;
  if(!nxt||!nxt.classList||!nxt.classList.contains('pr-items-row')||nxt.dataset.id!==recordId){{
    const colSpan=getRenderedRegisterColumnCount(tr);
    const panel=document.createElement('div');panel.className='pr-items-panel';
    panel.innerHTML='<div class="pr-items-empty">Loading...</div>';
    const td=document.createElement('td');td.colSpan=colSpan;td.appendChild(panel);
    const row=document.createElement('tr');row.className='pr-items-row';row.dataset.id=recordId;row.appendChild(td);
    row.style.display='none';
    tr.parentNode.insertBefore(row, tr.nextSibling);
    sizePrItemsPanel(panel);
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
    const panel=nxt.querySelector('.pr-items-panel');
    if(panel){{
      const searchTxt = document.getElementById('srchbox')?.value.trim() || '';
      panel.innerHTML=renderPrItemsTable(items, legacyText, searchTxt);
      sizePrItemsPanel(panel);
    }}
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
function buildRevisionDraftFromSelected(){{
  if(isLTRTab())return null;
  const selectedRows=getCheckedRecordRows();
  if(selectedRows.length>1){{
    alert('Please select only one document to create a new revision.');
    return false;
  }}
  const selected=selectedRows[0]||null;
  if(!selected)return null;
  const nextDocNo=incrementRevisionNumber(selected.docNo||'');
  if(!nextDocNo)return null;
  return sanitizeRevisionDraftData(selected,nextDocNo);
}}

function addRecord(){{
  state.editId=null;
  state.revisionDraftActive=false;
  state.savingRecord=false;setRecordSaveState(false);
  const draft=buildRevisionDraftFromSelected();
  if(draft===false)return;
  if(draft){{
    state.revisionDraftActive=true;
    document.getElementById('rec-title').textContent='Add Revision Draft';
    buildForm(draft,{{mode:'revisionDraft'}});
  }}else{{
    document.getElementById('rec-title').textContent='Add Document';
    buildForm(null,{{suggestedDocNo:suggestNextDocNoFromVisibleRows()}});
  }}
  openM('rec-modal');
}}
function editRec(id){{state.editId=id;state.revisionDraftActive=false;state.savingRecord=false;setRecordSaveState(false);const row=state.recs.find(r=>r._id===id);if(!row)return;document.getElementById('rec-title').textContent='Edit Document';buildForm(row);openM('rec-modal');}}

function setRecordSaveState(isSaving){{
  const btn=document.getElementById('rec-save-btn');
  if(!btn)return;
  btn.disabled=!!isSaving;
  btn.textContent=isSaving?'Saving...':'Save';
  btn.style.opacity=isSaving?'.72':'';
  btn.style.cursor=isSaving?'wait':'';
}}

function releaseRecordSave(){{
  state.savingRecord=false;
  setRecordSaveState(false);
}}

function isCalculatedFormField(col){{
  const key=String(col?.col_key||'');
  const type=String(col?.col_type||'').toLowerCase();
  return ['expectedReplyDate','duration','_duration','_duration_today'].includes(key)||
    ['auto_date','auto_num','duration_calc'].includes(type);
}}

function normalizeNocFieldText(col){{
  const rawKey=String(col?.col_key||'');
  const rawLabel=String(col?.label||'');
  const spacedKey=rawKey.replace(/([a-z])([A-Z])/g,'$1 $2');
  const text=(spacedKey+' '+rawLabel).toLowerCase().replace(/[_./()&+\\-]+/g,' ').replace(/\\s+/g,' ').trim();
  const compact=text.replace(/\\s+/g,'');
  const keyCompact=spacedKey.toLowerCase().replace(/[_./()&+\\-]+/g,' ').replace(/\\s+/g,'').trim();
  return {{text,compact,keyCompact}};
}}

function getNocFieldProfile(col){{
  const meta=normalizeNocFieldText(col);
  const has=(...terms)=>terms.some(term=>meta.text.includes(term)||meta.compact.includes(term.replace(/\\s+/g,'')));
  const profile=(section,rank,span='span-1')=>({{section,rank,span}});
  const exact={{
    docno:profile('NOC Basic Info',10,'span-1'),
    nocno:profile('NOC Basic Info',10,'span-1'),
    title:profile('NOC Basic Info',20,'span-2'),
    nocsubject:profile('NOC Basic Info',20,'span-2'),
    nocdescription:profile('NOC Basic Info',30,'full'),
    originatingdocument:profile('NOC Basic Info',40,'span-2'),
    partaissuedate:profile('Part A / B Tracking',110,'span-1'),
    partbreturndate:profile('Part A / B Tracking',120,'span-1'),
    partbstatus:profile('Part A / B Tracking',130,'span-1'),
    partcissuedate:profile('Part C / D & Cost Tracking',210,'span-1'),
    submittedcost:profile('Part C / D & Cost Tracking',220,'span-1'),
    partdreturndate:profile('Part C / D & Cost Tracking',230,'span-1'),
    partdstatus:profile('Part C / D & Cost Tracking',240,'span-1'),
    finalapprovedcost:profile('Part C / D & Cost Tracking',250,'span-1'),
    vono:profile('VO Details',310,'span-1'),
    voissuedate:profile('VO Details',320,'span-1'),
    vobasevalue:profile('VO Details',330,'span-1'),
    vovaluewithsiandvat:profile('VO Details',340,'span-1'),
    originalfile:profile('Remarks / Files',510,'span-2'),
    filelocation:profile('Remarks / Files',520,'span-2'),
    remarks:profile('Remarks / Files',540,'full')
  }};
  if(exact[meta.keyCompact])return exact[meta.keyCompact];

  if(has('noc no','document no'))return profile('NOC Basic Info',10,'span-1');
  if(has('noc subject','subject','title'))return profile('NOC Basic Info',20,'span-2');
  if(has('noc description','description','scope'))return profile('NOC Basic Info',30,'full');
  if(has('originating document','originating doc','origin document'))return profile('NOC Basic Info',40,'span-2');

  if(has('part a issue'))return profile('Part A / B Tracking',110,'span-1');
  if(has('part b return'))return profile('Part A / B Tracking',120,'span-1');
  if(has('part b status'))return profile('Part A / B Tracking',130,'span-1');

  if(has('part c issue'))return profile('Part C / D & Cost Tracking',210,'span-1');
  if(has('submitted cost'))return profile('Part C / D & Cost Tracking',220,'span-1');
  if(has('part d return'))return profile('Part C / D & Cost Tracking',230,'span-1');
  if(has('part d status'))return profile('Part C / D & Cost Tracking',240,'span-1');
  if(has('final approved cost','approved cost'))return profile('Part C / D & Cost Tracking',250,'span-1');

  if(has('vo no','variation order no'))return profile('VO Details',310,'span-1');
  if(has('vo issue'))return profile('VO Details',320,'span-1');
  if(has('vo base value','base value'))return profile('VO Details',330,'span-1');
  if(has('vo value','including si','incl si','vat'))return profile('VO Details',340,'span-1');
  if(meta.text.includes('vo ')||meta.text.includes('variation order'))return profile('VO Details',390,'span-1');

  if(has('original file'))return profile('Remarks / Files',510,'span-2');
  if(has('file link','file location','attachment','attach','link','url','file'))return profile('Remarks / Files',520,'span-2');
  if(has('remarks','remark','comment','notes'))return profile('Remarks / Files',540,'full');

  return profile('Other NOC Fields',900,'span-1');
}}

function buildNocFormFields(allCols,isLtrTab){{
  return buildDynamicOrderedFormFields(allCols,isLtrTab)
    .map((col,idx)=>({{col,idx,profile:getNocFieldProfile(col),calculated:isCalculatedFormField(col)}}))
    .sort((a,b)=>(Number(a.calculated)-Number(b.calculated))||(a.profile.rank-b.profile.rank)||(a.idx-b.idx))
    .map(x=>x.col);
}}

function buildDynamicOrderedFormFields(allCols,isLtrTab){{
  const byKey=new Map((allCols||[]).map(c=>[c.col_key,c]));
  const used=new Set();
  const ordered=[], calculated=[];
  const addCol=(col)=>{{
    if(!col||used.has(col.col_key))return;
    if(isLtrTab&&(isLTRInternalField(col)||isLTRExcludedField(col)))return;
    (isCalculatedFormField(col)?calculated:ordered).push(col);used.add(col.col_key);
  }};
  (state.cols||[]).forEach(vc=>addCol(byKey.get(vc.col_key)||vc));
  (allCols||[]).forEach(col=>addCol(col));
  return [...ordered,...calculated];
}}

function getOrderedRecordFormCols(allCols,isLtrTab,isNocTab=false){{
  return isNocTab?buildNocFormFields(allCols,isLtrTab):buildDynamicOrderedFormFields(allCols,isLtrTab);
}}

async function buildForm(row,opts={{}}){{
  const allCols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!allCols){{releaseRecordSave();return;}}
  const AUTO=new Set(['expectedReplyDate','duration','_duration','_duration_today']);
  const formRoot=document.getElementById('rec-form');formRoot.innerHTML='';
  const sectionBodies={{}};
  let nextNo='';
  const isAddMode=!state.editId;
  if(!row){{
    nextNo=String(opts.suggestedDocNo||'').trim();
    if(!nextNo){{const r=await apiFetch('/api/next_doc_no/'+PID+'/'+state.tab);nextNo=r?.next||'';}}
  }}
  const isPrTab=isPRTab();
  const isNocTab=isNOCTab();
  const isLtrTab=isLTRTab();
  const formCols=getOrderedRecordFormCols(allCols,isLtrTab,isNocTab);
  const ltrParentIdCol=isLtrTab?getLTRColForRole(allCols,'parentLetterId'):null;
  let ltrParentOptions=[];
  if(isLtrTab){{
    const q=row?('?record_id='+encodeURIComponent(row._id)):'';
    const parentData=await apiFetch('/api/letters/parent-options/'+PID+q);
    ltrParentOptions=Array.isArray(parentData?.options)?parentData.options:[];
  }}
  const prevCols=state.cols;
  state.cols=formCols.filter(c=>c.visible);
  const prDetailsKey=isPrTab?getPrDetailsColKey():null;
  state.cols=prevCols;

  let nocStageInp=null,nocBaseInp=null,nocTotalInp=null;
  if(isNocTab){{
    const sf=makeReadOnlyField('Stage Progress',getNocStageProgress(row||{{}}));
    nocStageInp=sf.inp;
    if(!isAddMode)getOrCreateFormSection(formRoot,sectionBodies,'Calculated / System Fields').appendChild(sf.grp);
  }}
  if(isLtrTab){{
    const hid=document.createElement('input');
    hid.type='hidden';
    hid.id='f-'+(ltrParentIdCol?.col_key||'parentLetterId');
    hid.value=String(ltrParentIdCol?(row?.[ltrParentIdCol.col_key]||''):(getLTRValue(row,allCols,'parentLetterId')||''));
    formRoot.appendChild(hid);
  }}
  for(const col of formCols){{
    const key=col.col_key;
    const longTextMeta=getLongTextMeta(col);
    const ctx={{isPrTab,isLtrTab,isNocTab}};
    const span=getDynamicFieldSpan(col,ctx);
    const grp=document.createElement('div');grp.className='fg'+(span==='full'?' full':(span==='span-2'?' span-2':''));
    const targetSection=getDynamicFormSection(col,{{isPrTab,isLtrTab,isNocTab}});
    const lbl=document.createElement('label');lbl.textContent=col.label;grp.appendChild(lbl);
    const ltrRole=isLtrTab?getLTRFieldRole(col):'';
    const val=(isLtrTab&&ltrRole)?(row?.[key]||getLTRValue(row,allCols,ltrRole)||''):(row?.[key]||'');
    if(AUTO.has(col.col_key)||isCalculatedFormField(col)){{
      if(isAddMode)continue;
      const inp=document.createElement('input');inp.id='f-'+key;inp.value=getCalculatedDisplayValue(col,row);
      inp.readOnly=true;
      inp.placeholder='';
      inp.style.cssText='background:var(--bg);font-weight:700;color:var(--mu)';
      grp.appendChild(inp);
      getOrCreateFormSection(formRoot,sectionBodies,getDynamicFormSection(col,ctx)).appendChild(grp);
      continue;
    }}
    if(isPrTab&&prDetailsKey&&key===prDetailsKey){{
      const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val;
      ta.rows=2;
      ta.style.cssText='resize:vertical;min-height:52px;overflow:hidden';
      ta.placeholder='Leave blank to auto-generate from items';
      bindSmartTextarea(ta);
      grp.appendChild(ta);
      const hint=document.createElement('div');
      hint.style.cssText='font-size:10px;color:var(--mu);margin-top:4px';
      hint.textContent='Leave blank to auto-generate from PR items.';
      grp.appendChild(hint);
      getOrCreateFormSection(formRoot,sectionBodies,targetSection).appendChild(grp);
      continue;
    }}
    if(isLtrTab&&getLTRFieldRole(col)==='parentLetterRef'){{
      lbl.textContent='Parent Letter';
      const currentParentId=String(ltrParentIdCol?(row?.[ltrParentIdCol.col_key]||''):(getLTRValue(row,allCols,'parentLetterId')||'')).trim();
      const currentParentRef=String(val||'').trim();
      const options=[...ltrParentOptions];
      if(currentParentId&&!options.some(o=>o.id===currentParentId)){{
        options.unshift({{
          id:currentParentId,
          doc_no:currentParentRef,
          subject:'',
          direction:'',
          from_party:'',
          to_party:'',
          date:''
        }});
      }}
      const sel=document.createElement('select');sel.id='f-'+key+'-selector';
      sel.innerHTML='<option value="">No parent letter</option>'+options.map(o=>`<option value="${{escHtml(o.id)}}">${{escHtml(formatLTRParentLabel(o))}}</option>`).join('');
      sel.value=currentParentId||'';
      grp.appendChild(sel);
      const meta=document.createElement('div');
      meta.style.cssText='font-size:10px;color:var(--mu);margin-top:5px;line-height:1.4';
      grp.appendChild(meta);
      const syncParent=()=>{{
        const picked=options.find(o=>o.id===sel.value)||null;
        const hid=document.getElementById('f-'+(ltrParentIdCol?.col_key||'parentLetterId'));
        if(hid)hid.value=picked?picked.id:'';
        meta.textContent=formatLTRParentMeta(picked);
      }};
      sel.onchange=syncParent;
      syncParent();
      getOrCreateFormSection(formRoot,sectionBodies,targetSection).appendChild(grp);
      continue;
    }}
    if(col.col_type==='date'){{const inp=document.createElement('input');inp.type='date';inp.id='f-'+key;inp.value=val;grp.appendChild(inp);}}
    else if(isLtrTab&&col.col_type==='dropdown'&&col.list_name){{
      const sel=document.createElement('select');sel.id='f-'+key;
      const ltrRole=getLTRFieldRole(col);
      const listName=ltrRole==='fromParty'||ltrRole==='toParty'
        ? 'correspondence_parties'
        : (ltrRole==='direction' ? 'letter_direction' : (ltrRole==='status' ? 'letter_status' : col.list_name));
      const opts=[...(state.lists[listName]||[])];
      if(val&&!opts.includes(val))opts.push(val);
      sel.innerHTML='<option value=""></option>'+opts.map(o=>`<option value="${{escHtml(o)}}">${{escHtml(o)}}</option>`).join('');
      sel.value=val||'';
      grp.appendChild(sel);
    }}
    else if(isLtrTab&&['title','description','remarks'].includes(getLTRFieldRole(col))){{
      const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val||'';
      const role=getLTRFieldRole(col);
      ta.rows=2;
      ta.style.cssText=role==='description'?'resize:vertical; min-height:58px; overflow:hidden':'resize:vertical; min-height:52px; overflow:hidden';
      ta.placeholder=role==='title'?'Use Enter for a multiline subject':'Use Enter for multiline text';
      bindSmartTextarea(ta);
      grp.appendChild(ta);
    }}
    else if(col.col_type==='dropdown'&&col.list_name){{
      // Add free-text input below multiselect
      const wrapper=document.createElement('div');
      wrapper.appendChild(buildMS(key,state.lists[col.list_name]||[],val));
      const freeInp=document.createElement('input');
      freeInp.id='f-free-'+key;freeInp.placeholder='Or type custom value...';
      freeInp.style.cssText='margin-top:4px;width:100%;padding:5px 8px;border:1px dashed var(--bd);border-radius:var(--rd);font-size:11px;color:var(--mu);outline:none';
      freeInp.title='Type a custom value not in the list';
      bindDirectionalInput(freeInp);
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
      bindDirectionalInput(inp);
      grp.appendChild(inp);
    }}
    else if(col.col_type==='number'){{
      const inp=document.createElement('input');inp.type='number';inp.step='0.01';inp.id='f-'+key;inp.value=val;
      if(isCurrencyField(col))inp.placeholder='0.00';
      grp.appendChild(inp);
      if(key==='voBaseValue')nocBaseInp=inp;
      if(key==='voValueWithSIAndVAT')nocTotalInp=inp;
    }}
    else if(longTextMeta){{
      const ta=document.createElement('textarea');ta.id='f-'+key;ta.value=val;
      ta.rows=longTextMeta.rows;
      ta.style.cssText=longTextMeta.style;
      ta.placeholder=longTextMeta.placeholder;
      bindSmartTextarea(ta);
      grp.appendChild(ta);
    }}
    else{{const inp=document.createElement('input');inp.id='f-'+key;inp.value=val;if(col.col_type==='link')inp.placeholder='https://...';bindDirectionalInput(inp);grp.appendChild(inp);}}
    if(isNocTab&&isCurrencyField(col))grp.style.gridColumn='span 1';
    getOrCreateFormSection(formRoot,sectionBodies,targetSection).appendChild(grp);
  }}
  if(isNocTab){{
    const syncNoc=()=>{{
      const data={{}};
      allCols.forEach(c=>{{
        const el=document.getElementById('f-'+c.col_key);
        if(!el)return;
        data[c.col_key]=el.classList?.contains?.('ms-con')?(el.dataset.value||''):el.value;
      }});
      if(nocStageInp)nocStageInp.value=getNocStageProgress(data);
      if(nocBaseInp&&nocTotalInp&&!String(nocBaseInp.value||'').trim())nocBaseInp.value=getNocAutoBaseValue(data);
    }};
    formRoot.querySelectorAll('input, textarea, select').forEach(el=>el.addEventListener('input',syncNoc));
    formRoot.querySelectorAll('input, textarea, select').forEach(el=>el.addEventListener('change',syncNoc));
    syncNoc();
  }}
  if(isPrTab){{
    const prSection=getOrCreateFormSection(formRoot,sectionBodies,'PR Items');
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
    grp.appendChild(wrap);prSection.appendChild(grp);
    const legacyText=prDetailsKey?String(row?.[prDetailsKey]||'').trim():'';
    initPrItemsEditor([], legacyText);
    if(row&&row._id)loadPrItemsForEdit(row._id, legacyText);
  }}
  if(!isNocTab)compactSingleFieldFormSections(formRoot);
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
  bindDirectionalInput(tr.querySelector('.pri-item'));
  bindDirectionalInput(tr.querySelector('.pri-unit'));
  bindDirectionalInput(tr.querySelector('.pri-remarks'));
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
  bindDirectionalInput(tr.querySelector('.pri-header'));
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
    dd.style.cssText='position:absolute;left:0;right:0;top:calc(100% + 2px);z-index:500;max-height:260px;overflow-y:auto';
    options.forEach(opt=>{{
      const it=document.createElement('div');
      it.className='ms-opt'+(sel.includes(opt)?' sel':'');
      const chk=document.createElement('input');chk.type='checkbox';chk.checked=sel.includes(opt);chk.style.cssText='flex-shrink:0;width:14px;height:14px;pointer-events:none';
      const lbl=document.createElement('span');lbl.textContent=opt;lbl.style.flex='1';
      it.appendChild(chk);it.appendChild(lbl);
      it.onclick=ev=>{{
        ev.stopPropagation();
        if(sel.includes(opt))sel.splice(sel.indexOf(opt),1);else sel.push(opt);
        con.dataset.value=sel.join(', ');render();
        chk.checked=sel.includes(opt);
        it.classList.toggle('sel',sel.includes(opt));
      }};
      dd.appendChild(it);
    }});
    con.style.position='relative';con.appendChild(dd);
  }};
  document.addEventListener('click',e=>{{if(!con.contains(e.target))con.querySelector('.ms-dd')?.remove();}},true);
  render();return con;
}}

function getFormSectionHint(title){{
  const hints={{
    'Core Register Fields':'Primary register identity and classification fields',
    'References / Technical Fields':'Drawing, MS, parent, and technical reference details',
    'Dates / Timeline':'Submission, issue, received, reply, and timeline dates',
    'Status / Workflow':'Status, direction, review, approval, and response controls',
    'Files / Notes':'Attachments, file links, remarks, comments, and descriptions',
    'Calculated / System Fields':'Readonly calculated values shown for context',
    'Other Dynamic Fields':'Additional fields configured for this document type',
    'Commercial & Quantities':'Values, quantities, and related numeric fields',
    'PR Items':'Line items and grouped procurement details',
    'NOC Basic Info':'NOC number, subject, description, and originating document',
    'Part A / B Tracking':'Part A issue and Part B return/status tracking',
    'Part C / D & Cost Tracking':'Part C/D tracking with submitted and final approved costs',
    'VO Details':'Variation order references, dates, and values',
    'Remarks / Files':'Remarks, original files, links, and attachments',
    'Other NOC Fields':'Additional NOC fields configured for this project',
    'Letter Information':'Reference and headline details for the letter',
    'Correspondence Routing':'Direction, parties, and linked parent correspondence',
    'Additional Details':'Remaining project-specific fields'
  }};
  return hints[title]||'Structured fields for this document section';
}}

const FORM_SECTION_ORDER=[
  'NOC Basic Info',
  'Part A / B Tracking',
  'Part C / D & Cost Tracking',
  'VO Details',
  'Remarks / Files',
  'Other NOC Fields',
  'Core Register Fields',
  'References / Technical Fields',
  'Dates / Timeline',
  'Status / Workflow',
  'Commercial & Quantities',
  'Files / Notes',
  'Other Dynamic Fields',
  'Calculated / System Fields',
  'PR Items'
];

function getFormSectionOrder(title){{
  const idx=FORM_SECTION_ORDER.indexOf(title);
  return idx>=0?idx:FORM_SECTION_ORDER.length;
}}

function compactSingleFieldFormSections(root){{
  (root||document).querySelectorAll('.form-section').forEach(sec=>{{
    const fields=[...sec.querySelectorAll('.form-section-grid > .fg')];
    const compact=fields.length===1&&!fields[0].classList.contains('full')&&!fields[0].classList.contains('span-2');
    sec.classList.toggle('compact-one',compact);
  }});
}}

function getFormFieldText(col){{
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  return (key+' '+label).replace(/[_./()&+\\-]+/g,' ').replace(/\\s+/g,' ').trim();
}}

function hasFormTerm(text, terms){{
  const compact=text.replace(/\\s+/g,'');
  return terms.some(term=>text.includes(term)||compact.includes(String(term).replace(/\\s+/g,'')));
}}

function isGenericTimelineField(col){{
  const type=String(col?.col_type||'').toLowerCase();
  const text=getFormFieldText(col);
  if(['date','datetime'].includes(type))return true;
  if(hasFormTerm(text,['date','issued','issue date','submitted','submission','received','received date','reply date','return date','expected reply','actual reply','consultant reply','pmo reply']))return true;
  return hasFormTerm(text,['rec from','rec by','received from','received by']);
}}

function isGenericWorkflowField(col){{
  const text=getFormFieldText(col);
  return hasFormTerm(text,['status','consultant status','pmo status','direction','approval','approved','review','workflow','response','reply status','decision','stage','part b status','part d status']);
}}

function isGenericFileNoteField(col){{
  const text=getFormFieldText(col);
  if(hasFormTerm(text,['details of instruction','verbal instruction']))return false;
  return hasFormTerm(text,['original file','file location','file link','attachment','attach','link','url','remarks','remark','comment','comments','notes','note','description','narrative','content']);
}}

function isGenericReferenceField(col){{
  const text=getFormFieldText(col);
  return hasFormTerm(text,['item ref','dwg','drawing','approved dwg','ms ref','ms reference','spec ref','technical ref','tech ref','parent ref','parent technical','parent letter','reference','ref no','code ref','originating document','details of instruction','verbal instruction']);
}}

function isGenericCommercialField(col){{
  const text=getFormFieldText(col);
  return hasFormTerm(text,['qty','quantity','unit','amount','value','cost','price','rate','total']);
}}

function isGenericCoreField(col){{
  const text=getFormFieldText(col);
  if(isGenericTimelineField(col)||isGenericWorkflowField(col)||isGenericFileNoteField(col)||isGenericCommercialField(col)||isCalculatedFormField(col))return false;
  if(hasFormTerm(text,['originating document']))return false;
  return hasFormTerm(text,['docno','document no','letter ref','noc no','pr no','discipline','sub trade','trade','title','subject','floor','location','brand','origin','originator','prepared by','preparedby','company','responsible party','client','consultant','contractor','recipient']);
}}

function classifyFormFieldSemantic(col, ctx={{}}){{
  const longTextMeta=getLongTextMeta(col);
  const ltrRole=ctx.isLtrTab?getLTRFieldRole(col):'';
  if(isCalculatedFormField(col))return 'calculated';
  if(ctx.isLtrTab){{
    if(['docNo','title'].includes(ltrRole))return 'core';
    if(['direction','fromParty','toParty','issuedDate','receivedDate','parentLetterRef','parentLetterId'].includes(ltrRole))return 'workflow';
    if(['description','remarks'].includes(ltrRole)||longTextMeta)return 'notes';
  }}
  if(isGenericTimelineField(col))return 'timeline';
  if(isGenericWorkflowField(col))return 'workflow';
  if(isGenericReferenceField(col))return 'reference';
  if(isGenericCommercialField(col))return 'commercial';
  if(isGenericFileNoteField(col)||longTextMeta)return 'notes';
  if(isGenericCoreField(col))return 'core';
  return 'other';
}}
function getDynamicFormSection(col, ctx){{
  if(ctx?.isNocTab)return getNocFieldProfile(col).section;
  const semantic=classifyFormFieldSemantic(col,ctx);
  const map={{
    core:'Core Register Fields',
    reference:'References / Technical Fields',
    timeline:'Dates / Timeline',
    workflow:'Status / Workflow',
    notes:'Files / Notes',
    calculated:'Calculated / System Fields',
    commercial:'Commercial & Quantities',
    other:'Other Dynamic Fields'
  }};
  return map[semantic]||'Other Dynamic Fields';
}}

function getDynamicFieldSpan(col, ctx={{}}){{
  if(ctx?.isNocTab)return getNocFieldProfile(col).span;
  const key=String(col?.col_key||'').toLowerCase();
  const label=String(col?.label||'').toLowerCase();
  const type=String(col?.col_type||'').toLowerCase();
  const text=(key+' '+label).replace(/[_-]+/g,' ');
  const semantic=classifyFormFieldSemantic(col,ctx);
  if(isCompactCoreFormField(col))return 'span-1';
  if(getLongTextMeta(col)||['textarea','long_text'].includes(type))return 'full';
  if(semantic==='notes')return 'full';
  if(['title','subject','description','remarks','filelocation','file location','pr details','item ref','dwg','ms ref','parent letter'].some(v=>text.includes(v)))return 'span-2';
  if(semantic==='reference'&&text.length>24)return 'span-2';
  return 'span-1';
}}

function getCalculatedDisplayValue(col,row){{
  const key=String(col?.col_key||'');
  const label=String(col?.label||'').toLowerCase();
  const type=String(col?.col_type||'').toLowerCase();
  if(key==='expectedReplyDate')return row?._expectedReplyDate||row?.expectedReplyDate||'';
  if(key==='duration'||key==='_duration'||label.includes('duration')||label==='dur.'||key.toLowerCase().includes('duration'))return row?._duration||row?.[key]||row?.duration||'';
  if(type==='duration_calc'){{
    const[ds,de]=String(col.list_name||'issuedDate,actualReplyDate').split(',');
    return calcWD(row?.[ds.trim()]||'',row?.[de.trim()]||'');
  }}
  if(key.toLowerCase().includes('overdue'))return row?._overdue?'Yes':(row?.[key]||'');
  if(key.toLowerCase().includes('aging'))return row?.[key]||'';
  return row?.[key]||'';
}}

function getOrCreateFormSection(root, sectionMap, title){{
  if(sectionMap[title])return sectionMap[title];
  const sec=document.createElement('section');
  sec.className='form-section';
  sec.dataset.sectionOrder=String(getFormSectionOrder(title));
  const hdr=document.createElement('div');
  hdr.className='form-section-header';
  hdr.innerHTML=`<div><div class="form-section-title">${{escHtml(title)}}</div><div class="form-section-sub">${{escHtml(getFormSectionHint(title))}}</div></div>`;
  const body=document.createElement('div');
  body.className='form-section-grid';
  sec.appendChild(hdr);
  sec.appendChild(body);
  const order=getFormSectionOrder(title);
  const before=[...root.children].find(el=>Number(el.dataset.sectionOrder||999)>order);
  if(before)root.insertBefore(sec,before);else root.appendChild(sec);
  sectionMap[title]=body;
  return body;
}}

function appendSectionTitle(grid,title){{
  return getOrCreateFormSection(grid,{{}},title);
}}

function makeReadOnlyField(label,value){{
  const grp=document.createElement('div');grp.className='fg full';
  const lbl=document.createElement('label');lbl.textContent=label;grp.appendChild(lbl);
  const inp=document.createElement('input');inp.value=value||'';inp.readOnly=true;
  inp.style.cssText='background:var(--bg);font-weight:700;color:var(--pr)';
  grp.appendChild(inp);
  return {{grp,inp}};
}}

function renderNocSummary(rows){{
  const wrap=document.getElementById('noc-summary');
  const inner=document.getElementById('noc-summary-inner');
  if(!wrap||!inner)return;
  if(!isNOCTab()){{wrap.style.display='none';inner.innerHTML='';return;}}
  const totals={{
    submittedCost:0,
    finalApprovedCost:0,
    voBaseValue:0,
    voValueWithSIAndVAT:0,
  }};
  for(const row of (rows||[])){{
    for(const key of Object.keys(totals)){{
      const n=parseSafeNumber(row?.[key]);
      if(n!==null)totals[key]+=n;
    }}
  }}
  const items=[
    ['Visible NOCs', String((rows||[]).length), 'Filtered rows'],
    ['Submitted Cost', formatCurrencyValue(totals.submittedCost), 'Visible total'],
    ['Final Approved Cost', formatCurrencyValue(totals.finalApprovedCost), 'Visible total'],
    ['VO Base Value', formatCurrencyValue(totals.voBaseValue), 'Visible total'],
    ['VO Value Incl. SI & VAT', formatCurrencyValue(totals.voValueWithSIAndVAT), 'Visible total'],
  ];
  inner.innerHTML=items.map(([label,val,sub])=>`<div style="background:#fff;border:1px solid var(--bd);border-radius:8px;padding:10px 12px">
    <div style="font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.4px">${{label}}</div>
    <div style="font-size:18px;font-weight:800;color:var(--pr);margin-top:4px;line-height:1.2">${{val}}</div>
    <div style="font-size:10px;color:var(--mu);margin-top:3px">${{sub}}</div>
  </div>`).join('');
  wrap.style.display='';
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
  if(state.savingRecord)return;
  state.savingRecord=true;setRecordSaveState(true);
  try{{
  const allCols=await apiFetch('/api/columns/'+PID+'/'+state.tab);if(!allCols){{releaseRecordSave();return;}}
  const AUTO=new Set(['expectedReplyDate','duration','_duration','_duration_today']);
  const data={{}};
  for(const col of allCols){{
    if(AUTO.has(col.col_key))continue;
    const el=document.getElementById('f-'+col.col_key);if(!el)continue;
      if(el.classList.contains('ms-con'))data[col.col_key]=el.dataset.value||'';
      else if(el.tagName==='TEXTAREA')data[col.col_key]=el.value.replace(/\\r\\n/g,String.fromCharCode(10)).replace(/\\r/g,String.fromCharCode(10));
      else data[col.col_key]=el.value.trim();
  }}
  if(isLTRTab()){{
    const direction=String(getLTRValue(data,allCols,'direction')||'').trim().toLowerCase();
    const issued=String(getLTRValue(data,allCols,'issuedDate')||'').trim();
    const received=String(getLTRValue(data,allCols,'receivedDate')||'').trim();
    const parentId=String(getLTRValue(data,allCols,'parentLetterId')||'').trim();
    if(direction==='sent'&&!issued){{toast('Issue Date is required for Sent letters','er');releaseRecordSave();return;}}
    if(direction==='received'&&!received){{toast('Received Date is required for Received letters','er');releaseRecordSave();return;}}
    if(state.editId&&parentId&&parentId===state.editId){{toast('A letter cannot reference itself as parent','er');releaseRecordSave();return;}}
  }}
  if(isNOCTab()&&!String(data.voBaseValue||'').trim()&&String(data.voValueWithSIAndVAT||'').trim()){{
    const autoBase=getNocAutoBaseValue(data);
    if(autoBase)data.voBaseValue=autoBase;
  }}
  if(!data.docNo){{toast('Document No. required','er');releaseRecordSave();return;}}
    // Duration is computed server-side automatically
  let valErr=validateDocNo(data.docNo,state.recs||[],state.editId);
  if(valErr&&valErr.startsWith('GAP')){{
    const ok=confirm('Sequence gap detected: '+valErr.replace('GAP:','')+'. Continue anyway?');
    if(!ok){{releaseRecordSave();return;}}
    valErr=null;
  }}
  if(valErr){{toast('Validation error: '+valErr,'er');releaseRecordSave();return;}}
  if(state.editId)data._id=state.editId;
  const r=await apiFetch('/api/records/'+PID+'/'+state.tab,{{method:'POST',body:JSON.stringify(data)}});
  if(r&&r.ok){{
    const wasRevisionDraft=state.revisionDraftActive&&!state.editId;
    if(isPRTab()){{
      const items=getPrItemsFromEditor();
      const recId=r.id||state.editId;
      try{{
        const prRes=await apiFetch('/api/pr_items/'+recId,{{method:'POST',body:JSON.stringify({{items}})}});
        if(prRes&&prRes.ok)state.prItemsCache[recId]=items;
        else {{toast('Items save failed','er');releaseRecordSave();return;}}
      }}catch(e){{
        toast('Items save failed: '+e.message,'er');releaseRecordSave();return;
      }}
    }}
    if(wasRevisionDraft)clearSel();
    state.revisionDraftActive=false;
    closeM('rec-modal');const savedTab=state.tab;await loadRecords();await refreshCounts();toast(state.editId?'Updated':'Added','ok');releaseRecordSave();
  }}
  else {{toast('Error saving','er');releaseRecordSave();}}
  }}catch(e){{
    toast('Error saving: '+e.message,'er');
    releaseRecordSave();
  }}
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
  const dlBtn = document.getElementById('btn-bulk-dl');
  if(dlBtn) dlBtn.style.display = checked.length>0 ? 'inline-block' : 'none';
  const all=document.querySelectorAll('.chkcell input[data-id]');
  const ca=document.getElementById('chkall');if(ca)ca.checked=all.length>0&&checked.length===all.length;
  syncCheckboxRowHighlights();
}}
function selAll(v){{document.querySelectorAll('.chkcell input[data-id]').forEach(cb=>cb.checked=v);updBulk();}}
function clearSel(){{document.querySelectorAll('.chkcell input').forEach(cb=>cb.checked=false);updBulk();}}
async function bulkDel(){{
  const ids=[...document.querySelectorAll('.chkcell input[data-id]:checked')].map(cb=>cb.dataset.id);
  if(!ids.length||!confirm('Delete '+ids.length+' records?'))return;
  const r=await apiFetch('/api/records/bulk_delete',{{method:'POST',body:JSON.stringify({{ids}})}});
  if(r&&r.ok){{
    clearSel();await loadRecords();await refreshCounts();toast('✔ Deleted '+(r.deleted||0),'ok');
  }}
}}

async function bulkDownload(){{
  const ids=[...document.querySelectorAll('.chkcell input[data-id]:checked')].map(cb=>cb.dataset.id);
  if(!ids.length)return;
  
  const drvCol = state.cols.find(c => c.col_type === 'url' || c.col_key === 'fileLocation');
  if(!drvCol) return toast('Drive Link column not found', 'er');
  const key = drvCol.col_key;
  
  const links = [];
  ids.forEach(id => {{
    const rec = state.recs.find(r => String(r._id) === String(id));
    if(rec) {{
      const url = String(rec[key]||'').trim();
      const hasValidLink = url && url !== '#' && url.toLowerCase() !== 'null' && url.toLowerCase() !== 'none';
      if(hasValidLink) {{
        let fileId = null;
        let m = url.match(/d\\/([a-zA-Z0-9_-]+)/);
        if(m) fileId = m[1];
        else {{
          m = url.match(/id=([a-zA-Z0-9_-]+)/);
          if(m) fileId = m[1];
        }}
        if(fileId) {{
          links.push(`https://drive.google.com/uc?export=download&id=${{fileId}}`);
        }} else {{
           links.push(url);
        }}
      }}
    }}
  }});
  
  if(!links.length) return toast('No valid documents found in selected rows', 'wa');
  
  toast(`Starting download of ${{links.length}} file(s)...`, 'ok');
  
  links.forEach((link, idx) => {{
    setTimeout(() => {{
      const a = document.createElement('a');
      a.href = link;
      a.download = '';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }}, idx * 800);
  }});
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
  
  // Google Drive
  const drvTitle=document.createElement('div');drvTitle.className='stitle';drvTitle.textContent='Google Drive Integration';body.appendChild(drvTitle);
  const drvGrid=document.createElement('div');drvGrid.className='fgrid';
  drvGrid.innerHTML=`
    <div class="fg"><label>Drive Root Folder ID</label><input id="proj-drive-id" placeholder="ID from URL" value="${{proj.drive_folder_id||''}}"></div>
    <div class="fg" style="display:flex;align-items:flex-end">
      <button class="btn btn-sc" type="button" onclick="syncDriveLinks(this)" style="width:100%;height:32px;display:flex;justify-content:center;align-items:center;gap:6px;background:var(--bg)">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 0 0-9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M3 12a9 9 0 0 0 9 9 9.75 9.75 0 0 0 6.74-2.74L21 16"/><path d="M16 21v-5h5"/></svg>
        Sync Drive Links
      </button>
    </div>
  `;
  body.appendChild(drvGrid);

  // ── Distribution Matrix (visible to DC & Admins) ──────────────
  try {{
    // Use cached whoami (loaded once at page start) — no extra API call
    const isDCOrAdmin = _WHOAMI && (_WHOAMI.role==='superadmin'||_WHOAMI.role==='admin'||(_WHOAMI.dc_projects||[]).includes(PID));
    if(isDCOrAdmin) {{
      const distTitle=document.createElement('div');distTitle.className='stitle';distTitle.style.marginTop='18px';
      distTitle.innerHTML='📇 Distribution Matrix';body.appendChild(distTitle);
      const distNote=document.createElement('div');
      distNote.style.cssText='font-size:10px;color:var(--mu);margin:-4px 0 10px;padding:6px 8px;background:var(--bg);border-radius:4px';
      distNote.textContent='Configure who receives email notifications per document type and event. Changes are saved immediately.';
      body.appendChild(distNote);
      const distBtn=document.createElement('button');
      distBtn.className='btn btn-sc';
      distBtn.style.cssText='width:100%;display:flex;align-items:center;justify-content:center;gap:8px;padding:10px';
      distBtn.innerHTML='📇 Open Distribution Matrix';
      distBtn.onclick=()=>openDistributionMatrix(PID);
      body.appendChild(distBtn);
    }}
  }} catch(e){{ console.error('[DistMatrix UI]',e); }}

  const er={{...DEFAULT_EXPECTED_REPLY_RULE,...(proj.expected_reply_rule||{{}})}};
  const erTitle=document.createElement('div');erTitle.className='stitle';erTitle.textContent='Expected Reply Rule';body.appendChild(erTitle);
  const erNote=document.createElement('div');
  erNote.style.cssText='font-size:10px;color:var(--mu);margin:-4px 0 8px;padding:6px 8px;background:var(--bg);border-radius:4px';
  erNote.textContent='Project-specific rule for calculated Expected Reply dates. Existing saved records are not modified.';
  body.appendChild(erNote);
  const erGrid=document.createElement('div');erGrid.className='fgrid';
  erGrid.innerHTML=`
    <div class="fg"><label>REV00 Reply Days</label><input id="er-rev0-days" type="number" min="0" step="1" value="${{Number.isFinite(Number(er.rev0_reply_days))?Number(er.rev0_reply_days):14}}"></div>
    <div class="fg"><label>REV&gt;00 Reply Days</label><input id="er-rev-days" type="number" min="0" step="1" value="${{Number.isFinite(Number(er.rev_reply_days))?Number(er.rev_reply_days):7}}"></div>
    <div class="fg"><label>Calculation Mode</label><select id="er-calc-mode">
      <option value="calendar_days" ${{er.calculation_mode==='calendar_days'?'selected':''}}>Calendar Days</option>
      <option value="working_days" ${{er.calculation_mode!=='calendar_days'?'selected':''}}>Working Days</option>
    </select></div>
    <div class="fg"><label>Weekend Rule</label><select id="er-weekend-mode">
      <option value="friday_only" ${{er.weekend_mode==='friday_only'?'selected':''}}>Friday Only</option>
      <option value="friday_saturday" ${{er.weekend_mode==='friday_saturday'?'selected':''}}>Friday + Saturday</option>
      <option value="none" ${{er.weekend_mode==='none'?'selected':''}}>None</option>
    </select></div>
    <div class="fg"><label>Exclude Official Holidays</label><select id="er-exclude-holidays">
      <option value="yes" ${{er.exclude_official_holidays!==false?'selected':''}}>Yes</option>
      <option value="no" ${{er.exclude_official_holidays===false?'selected':''}}>No</option>
    </select></div>
  `;
  body.appendChild(erGrid);
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
  const rev0Days=parseInt(document.getElementById('er-rev0-days')?.value||'14',10);
  const revDays=parseInt(document.getElementById('er-rev-days')?.value||'7',10);
  data.expected_reply_rule={{
    rev0_reply_days:Number.isFinite(rev0Days)&&rev0Days>=0?rev0Days:14,
    rev_reply_days:Number.isFinite(revDays)&&revDays>=0?revDays:7,
    calculation_mode:document.getElementById('er-calc-mode')?.value||'working_days',
    weekend_mode:document.getElementById('er-weekend-mode')?.value||'friday_only',
    exclude_official_holidays:(document.getElementById('er-exclude-holidays')?.value||'yes')==='yes'
  }};
  const r=await apiFetch('/api/project/'+PID,{{method:'POST',body:JSON.stringify(data)}});
  if(r===null)return;
  if(r.ok){{
    closeM('proj-modal');toast('✔ Project saved!','ok');
    setTimeout(()=>location.reload(),400);
  }}else toast('Save failed','er');
}}

// ── Distribution Matrix ──────────────────────────────────────
async function openDistributionMatrix(pid) {{
  if(!pid){{ toast('No project selected','er'); return; }}

  try {{
    const [docTypes, savedDist, projUsers] = await Promise.all([
      apiFetch('/api/doc_types/'+pid).catch(()=>[]),
      apiFetch('/api/distribution/'+pid).catch(()=>({{}})),
      apiFetch('/api/project_users/'+pid).catch(()=>[])
    ]);
    if(!docTypes || !docTypes.length){{ toast('No doc types found for this project','wa'); return; }}
    
    const body = document.getElementById('dist-body');
    body.innerHTML = '';

    function makeUserSelector(initialUsers, pid, dtId) {{
      const wrap = document.createElement('div');
      wrap.style.cssText = 'display:flex;flex-wrap:wrap;gap:4px;padding:6px 8px;border:1px solid var(--bd);border-radius:6px;min-height:36px;background:var(--bg);cursor:text;';
      
      let users = [...(initialUsers||[])];
      
      function renderTags() {{
        wrap.innerHTML='';
        users.forEach((em, i) => {{
          const chip = document.createElement('span');
          chip.style.cssText = 'display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:99px;background:#dbeafe;color:#1e40af;font-size:11px;font-weight:600';
          chip.innerHTML = em;
          const cross = document.createElement('span');
          cross.innerHTML = '✕';
          cross.style.cssText = 'cursor:pointer;font-size:14px;line-height:1';
          cross.onclick = async () => {{
            users = users.filter(u => u !== em);
            await saveDistRow();
            renderTags();
          }};
          chip.appendChild(cross);
          wrap.appendChild(chip);
        }});
        
        const sel=document.createElement('select');
        sel.style.cssText='border:none;outline:none;background:transparent;font-size:12px;min-width:180px;flex:1;color:var(--tx)';
        sel.innerHTML = '<option value="">+ Add User...</option>';
        projUsers.forEach(u => {{
          if(!users.includes(u.username)) {{
            sel.innerHTML += `<option value="${{u.username}}">${{u.username}} (${{u.role}})</option>`;
          }}
        }});
        
        sel.onchange=async(e)=>{{
          const val = sel.value;
          if(val && !users.includes(val)) {{
            users.push(val);
            await saveDistRow();
            renderTags();
          }}
        }};
        wrap.appendChild(sel);
      }}
      
      async function saveDistRow() {{
        try {{
          const r = await apiFetch('/api/distribution/'+pid, {{
            method:'POST',
            body: JSON.stringify({{doc_type_id:dtId, event_type:'access', emails:users}})
          }});
          if(r&&r.ok) toast('✔ Saved','ok');
          else toast('Save failed','er');
        }} catch(e){{ toast('Save error','er'); console.error(e); }}
      }}
      
      renderTags();
      return wrap;
    }}

    docTypes.forEach(dt=>{{
      const dtSect=document.createElement('div');
      dtSect.style.cssText='margin-bottom:20px;border:1px solid var(--bd);border-radius:8px;overflow:hidden';

      const dtHdr=document.createElement('div');
      dtHdr.style.cssText='padding:10px 14px;background:var(--bg2,#f1f5f9);font-weight:700;font-size:13px;display:flex;align-items:center;justify-content:space-between;gap:8px';
      
      const titleSpan = document.createElement('div');
      titleSpan.innerHTML='<span style="font-size:16px">📄</span> '+dt.name+' <code style="font-size:10px;padding:1px 6px;border-radius:4px;background:#e2e8f0;color:#475569">'+dt.code+'</code>';
      
      const magicBtn = document.createElement('button');
      magicBtn.className = 'btn-ok';
      magicBtn.innerHTML = 'Generate Magic Link 🔗';
      magicBtn.style.padding = '4px 10px';
      magicBtn.style.fontSize = '12px';
      magicBtn.onclick = async () => {{
        try {{
          const r = await apiFetch(`/api/magic/generate/${{pid}}/${{dt.id}}`, {{method:'POST'}});
          if(r && r.ok) {{
            navigator.clipboard.writeText(r.link);
            toast('Magic Link copied to clipboard!','ok');
          }} else toast('Failed to generate link', 'er');
        }} catch(e){{ toast('Error','er'); }}
      }};
      
      dtHdr.appendChild(titleSpan);
      dtHdr.appendChild(magicBtn);
      dtSect.appendChild(dtHdr);

      const distRows=document.createElement('div');
      distRows.style.padding='12px 14px';
      
      const evtRow=document.createElement('div');
      evtRow.style.cssText='display:grid;grid-template-columns:200px 1fr;gap:10px;align-items:start;margin-bottom:12px;padding-bottom:12px;border-bottom:none';
      
      const evtLabel=document.createElement('div');
      evtLabel.innerHTML='<span style="font-size:12px;font-weight:700;color:#0f172a">Assigned Engineers</span><div style="font-size:10px;color:var(--mu);margin-top:2px">Users with access to daily digest</div>';
      
      const savedUsers = (savedDist[dt.id]||{{}})['access']||[];
      const tagInput = makeUserSelector(savedUsers, pid, dt.id);
      
      evtRow.appendChild(evtLabel);
      evtRow.appendChild(tagInput);
      distRows.appendChild(evtRow);
      
      dtSect.appendChild(distRows);
      body.appendChild(dtSect);
    }});

    openM('dist-modal');
  }} catch(e){{
    console.error('[DistMatrix]', e);
    toast('Error loading distribution matrix','er');
  }}
}}

async function syncDriveLinks(btn) {{
  const folderId = document.getElementById('proj-drive-id')?.value.trim();
  if (!folderId) {{ toast('Please enter a Drive Folder ID first.', 'er'); return; }}
  
  const originalHtml = btn.innerHTML;
  btn.innerHTML = '⏳ Syncing in background...';
  btn.disabled = true;
  
  const r = await apiFetch('/api/drive/sync/'+PID, {{
    method: 'POST',
    body: JSON.stringify({{ folder_id: folderId }})
  }});
  
  btn.innerHTML = originalHtml;
  btn.disabled = false;
  
  if (r && r.ok) {{
    toast('✔ Drive Sync completed in background', 'ok');
  }} else {{
    toast('Sync request failed', 'er');
  }}
}}

// Doc Types
function normalizeDocTypeReplyOverride(dt){{
  const o=dt?.expected_reply_override||{{}};
  return {{
    use_expected_reply_override:!!(o.use_expected_reply_override||o.enabled),
    rev0_reply_days_override:Number.isFinite(Number(o.rev0_reply_days_override))?Number(o.rev0_reply_days_override):14,
    rev_reply_days_override:Number.isFinite(Number(o.rev_reply_days_override))?Number(o.rev_reply_days_override):7
  }};
}}
function toggleDocTypeReplyOverride(){{
  const on=document.getElementById('dt-er-use')?.checked;
  const fields=document.getElementById('dt-er-fields');
  if(fields)fields.style.display=on?'grid':'none';
}}
function setDocTypeModalMode(dt=null){{
  const edit=!!dt;
  document.getElementById('dt-modal-title').textContent=edit?'Document Type Settings':'Add Document Type';
  document.getElementById('dt-save-btn').textContent=edit?'Save':'Add';
  document.getElementById('dt-edit-id').value=edit?dt.id:'';
  document.getElementById('dt-code').value=edit?(dt.code||dt.id||''):'';
  document.getElementById('dt-name').value=edit?(dt.name||''):'';
  document.getElementById('dt-code').disabled=false;
  const o=normalizeDocTypeReplyOverride(dt||{{}});
  document.getElementById('dt-er-use').checked=o.use_expected_reply_override;
  document.getElementById('dt-er-rev0').value=o.rev0_reply_days_override;
  document.getElementById('dt-er-rev').value=o.rev_reply_days_override;
  toggleDocTypeReplyOverride();
}}
function collectDocTypeReplyOverride(){{
  const rev0=parseInt(document.getElementById('dt-er-rev0')?.value||'14',10);
  const rev=parseInt(document.getElementById('dt-er-rev')?.value||'7',10);
  return {{
    use_expected_reply_override:!!document.getElementById('dt-er-use')?.checked,
    rev0_reply_days_override:Number.isFinite(rev0)&&rev0>=0?rev0:14,
    rev_reply_days_override:Number.isFinite(rev)&&rev>=0?rev:7
  }};
}}
function addDocType(){{setDocTypeModalMode(null);openM('dt-modal');}}
async function saveDocType(){{
  const editId=document.getElementById('dt-edit-id').value.trim();
  const code=document.getElementById('dt-code').value.trim().toUpperCase();
  const name=document.getElementById('dt-name').value.trim();
  if(!code||!name){{toast('Code and name required','er');return;}}
  const payload={{code,name,expected_reply_override:collectDocTypeReplyOverride()}};
  if(editId){{
    await apiFetch('/api/doc_types/'+PID+'/'+editId,{{method:'PATCH',body:JSON.stringify(payload)}});
    closeM('dt-modal');await loadDTs(true);toast('Type settings saved','ok');
  }}else{{
    await apiFetch('/api/doc_types/'+PID,{{method:'POST',body:JSON.stringify(payload)}});
    closeM('dt-modal');await loadDTs();switchTab(code);toast('Type added','ok');
  }}
}}

// Lists
async function openLists(){{
  await loadLists(true);
  ensureLTRProjectLists();
  const metaData=await apiFetch('/api/lists_meta/'+PID)||{{}};
  const META_LABELS={{approved:{{lbl:'Approved',bg:'#bbf7d0',fg:'#166534'}},rejected:{{lbl:'Rejected',bg:'#fce7f3',fg:'#831843'}},pending:{{lbl:'Pending',bg:'#fef9c3',fg:'#713f12'}},cancelled:{{lbl:'Cancelled',bg:'#f1f5f9',fg:'#94a3b8'}}}};
  const body=document.getElementById('lists-body');body.innerHTML='';
  for(const[ln,items]of Object.entries(state.lists).filter(([ln])=>!isHiddenSystemListName(ln))){{
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
      const idx=items.indexOf(item);
      const ren=document.createElement('button');ren.textContent='Rename';
      ren.onclick=()=>renameItem(ln,item);li.appendChild(ren);
      const up=document.createElement('button');up.textContent='↑';
      up.disabled=idx===0;
      up.style.opacity=idx===0?'.45':'1';
      up.onclick=()=>moveListItem(ln,item,-1);li.appendChild(up);
      const down=document.createElement('button');down.textContent='↓';
      down.disabled=idx===items.length-1;
      down.style.opacity=idx===items.length-1?'.45':'1';
      down.onclick=()=>moveListItem(ln,item,1);li.appendChild(down);
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
async function renameItem(ln,item){{
  const next=prompt('Rename list item:',item);
  if(next===null)return;
  const new_item=next.trim();
  if(!new_item||new_item===item)return;
  await apiFetch('/api/lists/'+PID,{{method:'PATCH',body:JSON.stringify({{list_name:ln,old_item:item,new_item}})}});
  await loadLists(true);openLists();toast('✔ Item renamed','ok');
}}
async function moveListItem(ln,item,delta){{
  const items=[...(state.lists[ln]||[])];
  const idx=items.indexOf(item),next=idx+delta;
  if(idx<0||next<0||next>=items.length)return;
  [items[idx],items[next]]=[items[next],items[idx]];
  await apiFetch('/api/lists/'+PID+'/reorder',{{method:'POST',body:JSON.stringify({{list_name:ln,order:items}})}});
  await loadLists(true);openLists();toast('✔ List order updated','ok');
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
  const visibleCols=isLTRTab()?cols.filter(c=>!isLTRInternalField(c)&&!isLTRExcludedField(c)):cols;
  visibleCols.forEach(col=>{{
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
  ensureLTRProjectLists();
  const ls=document.getElementById('col-list');
  ls.innerHTML=Object.keys(state.lists).filter(k=>!isHiddenSystemListName(k)).map(k=>`<option value="${{k}}">${{k}}</option>`).join('');
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
    const assigned_cached = u.projects || [];
    const row=document.createElement('div');row.className='urow';
    row.innerHTML=`<span style="flex:1;font-weight:600">👤 ${{u.username}}</span>
      <input id="email-${{u.username}}" placeholder="Email (e.g. dc@company.com)" value="${{u.email || ''}}" style="flex:1; margin-right: 8px; padding: 2px 6px; font-size: 11px; border: 1px solid var(--bd); border-radius: 4px;" onblur="updUsrEmail('${{u.username}}')">
      <span class="badge" style="background:#fef3c7;color:#92400e">${{u.role.toUpperCase()}}</span>
      ${{u.username!=='admin'?`<button class="btn btn-sc btn-sm" onclick="chgPw('${{u.username}}')">🔑 PW</button>
        <button class="btn btn-er btn-sm" onclick="delUsr('${{u.username}}')">✕</button>`:
        '<span style="font-size:10px;color:var(--mu)">(protected)</span>'}}`;
    body.appendChild(row);
    if(u.username!=='admin'){{
      const roleSel=document.createElement('select');
      roleSel.id='role-'+u.username;
      roleSel.className='btn btn-sc btn-sm';
      roleSel.style.cssText='height:auto;padding:4px 8px;outline:none;';
      roleSel.innerHTML=['viewer','editor','admin','superadmin'].map(r=>`<option value="${{r}}" ${{u.role===r?'selected':''}}>${{r==='superadmin'?'Super Admin':r.charAt(0).toUpperCase()+r.slice(1)}}</option>`).join('');
      const roleBtn=document.createElement('button');
      roleBtn.className='btn btn-pr btn-sm';
      roleBtn.textContent='Save Role';
      roleBtn.onclick=()=>updUsrRole(u.username);
      row.insertBefore(roleSel,row.children[1]||null);
      row.insertBefore(roleBtn,row.children[2]||null);
    }}
    if(u.role!=='superadmin'){{
      const ad=document.createElement('div');ad.style.cssText='padding:4px 10px 10px 32px;border-bottom:1px solid var(--bd);margin-bottom:4px';
      ad.innerHTML='<div style="font-size:10px;color:var(--mu);margin-bottom:4px">Project access:</div>';
      const assigned=assigned_cached;
      const pl=document.createElement('div');pl.style.cssText='display:flex;flex-wrap:wrap;gap:5px';
      projects.forEach(p=>{{
        const projAccess = assigned.find(a => a.project_id === p.id);
        const isOn = !!projAccess;
        const isDC = isOn && projAccess.is_dc;
        const btn=document.createElement('button');
        btn.style.cssText='padding:3px 10px;border-radius:4px;cursor:pointer;font-size:11px;font-weight:700;font-family:inherit;transition:all .15s;border:2px solid '+(isDC?'#22c55e':(isOn?'#f0a500':'#e2e8f0'))+';background:'+(isOn?'#1a3a5c':'#f8fafc')+';color:'+(isOn?'#fff':'#94a3b8');
        btn.textContent=p.code + (isDC ? ' (DC)' : '');
        btn.title=p.name + " | Left-click to assign access. Right-click to set as Project DC";
        if(isOn)btn.dataset.on='1';
        
        btn.onclick=async()=>{{
          const on=!!btn.dataset.on;
          btn.dataset.on = on ? '' : '1';
          const isOn = !on;
          btn.style.borderColor = isDC?'#22c55e':(isOn?'#f0a500':'#e2e8f0');
          btn.style.background = isOn?'#1a3a5c':'#f8fafc';
          btn.style.color = isOn?'#fff':'#94a3b8';
          btn.textContent = p.code + (isDC ? ' (DC)' : '');
          try {{
            await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:on?'unassign':'assign',username:u.username,project_id:p.id,is_dc:false}})}});
          }} catch(e) {{ openAdmin(); }}
        }};
        
        btn.oncontextmenu = async (e) => {{
          e.preventDefault();
          const currIsOn = !!btn.dataset.on;
          if(!currIsOn) return toast('User must be assigned to project first','wa');
          await apiFetch('/api/users', {{method:'POST', body:JSON.stringify({{action:'assign',username:u.username,project_id:p.id, is_dc: !isDC}})}});
          openAdmin(); // Reload to refresh isDC constant state accurately
        }};
        
        pl.appendChild(btn);
      }});
      ad.appendChild(pl);body.appendChild(ad);
    }}
  }}
  const at=document.createElement('div');at.className='stitle';at.textContent='➕ Add User';body.appendChild(at);
  const ar=document.createElement('div');
  ar.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1.5fr 1fr 1fr auto;gap:8px;align-items:end">
    <div class="fg"><label>Username</label><input id="nu-name" placeholder="username"></div>
    <div class="fg"><label>Email</label><input id="nu-email" placeholder="Email"></div>
    <div class="fg"><label>Role</label><select id="nu-role">
      <option value="editor">Editor</option><option value="viewer">Viewer</option>
      <option value="superadmin">Super Admin</option></select></div>
    <div class="fg"><label>Password</label><input id="nu-pw" type="password"></div>
    <button class="btn btn-pr btn-sm" style="margin-bottom:1px" onclick="addUsr()">Add</button></div>`;
  body.appendChild(ar);openM('admin-modal');
}}
async function addUsr(){{
  const name=document.getElementById('nu-name')?.value.trim().toLowerCase();
  const role=document.getElementById('nu-role')?.value;
  const pw=document.getElementById('nu-pw')?.value;
  const email=document.getElementById('nu-email')?.value || '';
  if(!name||!pw){{toast('Username and password required','er');return;}}
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'add',username:name,role,password:pw,email:email}})}});
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

async function updUsrEmail(u){{
  const email=document.getElementById('email-'+u)?.value;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'update_email',username:u,email}})}});
  if(r&&r.ok) toast('Email saved','ok'); else toast('Failed to save email','er');
}}
async function updUsrRole(u){{
  const role=document.getElementById('role-'+u)?.value;
  if(!role)return;
  const r=await apiFetch('/api/users',{{method:'POST',body:JSON.stringify({{action:'update_role',username:u,role}})}});
  if(r&&r.ok){{toast('Role updated','ok');closeM('admin-modal');openAdmin();}}
  else toast((r&&r.error)||'Role update failed','er');
}}

function executeAdvancedExport(){{
  const format = document.getElementById('export-format').value;
  const scope = document.getElementById('export-scope').value;
  const time = document.getElementById('export-time').value;
  const status = document.getElementById('export-status').value;
  
  let targetTab = (scope === 'current' && state.tab) ? state.tab : 'all';
  
  let baseUrl = '';
  if(format === 'excel') {{
    baseUrl = (targetTab === 'all') ? `/api/export_all/${{PID}}` : `/api/export/${{PID}}/${{targetTab}}`;
  }} else if (format === 'pdf') {{
    baseUrl = (targetTab === 'all') ? `/api/export_pdf_all/${{PID}}` : `/api/export_pdf/${{PID}}/${{targetTab}}`;
  }}

  const params = new URLSearchParams();
  if(time !== 'all') params.set('days', time);
  if(status !== 'all') params.set('status', status);
  
  const query = params.toString();
  const finalUrl = query ? `${{baseUrl}}?${{query}}` : baseUrl;
  
  window.location = finalUrl;
  closeM('export-modal');
}}
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

async function openProjectSettings() {{
  if(!CAN_EDIT || !['admin','superadmin'].includes(ROLE)) return toast('Forbidden', 'er');
  const lst = document.getElementById('status-colors-list');
  lst.innerHTML = '';
  
  await loadLists(false);
  const statuses = new Set();
  const statusListKey = Object.keys(state.lists || {{}}).find(k => k.toLowerCase().startsWith('status'));
  
  if(statusListKey && state.lists[statusListKey] && state.lists[statusListKey].length > 0) {{
    state.lists[statusListKey].forEach(s => {{
      if(s.trim()) statuses.add(s.trim());
    }});
  }} else if(state.recs) {{
    state.recs.forEach(r => {{
      if(r.status) {{
        String(r.status).split(',').forEach(s => {{
          const st = s.trim();
          if(st) statuses.add(st);
        }});
      }}
    }});
  }}
  
  if(statuses.size === 0) {{
    lst.innerHTML = '<div style="color:var(--mu);font-size:11px;padding:10px">No statuses found in any records yet.</div>';
  }} else {{
    [...statuses].sort().forEach(s => {{
      let [bg, fg] = SC[s] || ['e5e7eb', '374151'];
      bg = bg.startsWith('#') ? bg : '#' + bg;
      fg = fg.startsWith('#') ? fg : '#' + fg;
      
      const row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;gap:10px;padding:6px 12px;background:var(--bg);border:1px solid var(--bd);border-radius:6px';
      
      const lbl = document.createElement('div');
      lbl.style.cssText = 'flex:1;font-weight:600;font-size:12px;color:var(--tx)';
      lbl.textContent = s;
      
      const bgLbl = document.createElement('div');
      bgLbl.style.cssText = 'font-size:10px;color:var(--mu)';
      bgLbl.textContent = 'Background:';
      
      const bgInp = document.createElement('input');
      bgInp.type = 'color';
      bgInp.value = bg;
      bgInp.dataset.status = s;
      bgInp.dataset.type = 'bg';
      bgInp.title = 'Background Color';
      bgInp.style.cssText = 'width:28px;height:24px;padding:0;border:none;cursor:pointer;background:none';
      
      const fgLbl = document.createElement('div');
      fgLbl.style.cssText = 'font-size:10px;color:var(--mu);margin-left:10px';
      fgLbl.textContent = 'Text:';
      
      const fgInp = document.createElement('input');
      fgInp.type = 'color';
      fgInp.value = fg;
      fgInp.dataset.status = s;
      fgInp.dataset.type = 'fg';
      fgInp.title = 'Text Color';
      fgInp.style.cssText = 'width:28px;height:24px;padding:0;border:none;cursor:pointer;background:none';
      
      row.appendChild(lbl);
      row.appendChild(bgLbl);
      row.appendChild(bgInp);
      row.appendChild(fgLbl);
      row.appendChild(fgInp);
      lst.appendChild(row);
    }});
  }}
  openM('project-settings-modal');
}}

async function saveProjectSettings() {{
  const btn = document.querySelector('#project-settings-modal .btn-pr');
  btn.disabled = true; btn.textContent = 'Saving...';
  
  try {{
    const lst = document.getElementById('status-colors-list');
    const status_colors = {{}};
    const rows = lst.children;
    for(let i=0; i<rows.length; i++) {{
      const bgInp = rows[i].querySelector('input[data-type="bg"]');
      const fgInp = rows[i].querySelector('input[data-type="fg"]');
      if(bgInp && fgInp) {{
        const s = bgInp.dataset.status;
        const bg = bgInp.value.replace('#','').toLowerCase();
        const fg = fgInp.value.replace('#','').toLowerCase();
        
        let [dbg, dfg] = SC_DEFAULTS[s] || ['e5e7eb', '374151'];
        dbg = dbg.toLowerCase();
        dfg = dfg.toLowerCase();
        
        if(bg !== dbg || fg !== dfg) {{
          status_colors[s] = [bg, fg];
        }}
      }}
    }}
    
    const r = await apiFetch('/api/project/'+PID+'/settings', {{method:'POST', body:JSON.stringify({{status_colors}})}});
    if(r && r.ok) {{
      toast('✔ Settings saved successfully','ok');
      setTimeout(() => location.reload(), 600);
    }} else {{
      toast(r?.error || 'Error saving settings','er');
    }}
  }} catch(e) {{
    toast('Error: ' + e.message, 'er');
  }} finally {{
    btn.disabled = false; btn.textContent = 'Save Settings';
  }}
}}
</script></body></html>"""


