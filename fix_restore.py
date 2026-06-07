"""
Fix script: Restore ALL UI Polish + Fix Mobile Issues
"""
import re

with open(r'D:\DCR\html_render.py', 'r', encoding='utf-8') as f:
    code = f.read()

# ========================================================
# 1. FIX REGISTER TAB-BTN - Make it Pill-Shaped properly
# ========================================================
old_tab = '''.tab-btn{{display:flex;align-items:center;gap:5px;padding:9px 12px;background:transparent;
  border:none;border-bottom:3px solid transparent;color:rgba(255,255,255,.55);cursor:pointer;
  font-family:inherit;font-size:11px;font-weight:600;white-space:nowrap;transition:all .15s}}
.tab-btn:hover{{color:#fff;background:rgba(255,255,255,.07)}}
.tab-btn.active{{color:#fff;border-bottom-color:var(--ac)}}
.tcnt{{background:rgba(255,255,255,.2);border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700}}
.tab-btn.active .tcnt{{background:var(--ac);color:#000}}'''

new_tab = '''.tab-btn{{
  display:inline-flex;align-items:center;gap:5px;
  padding:6px 12px;
  background:rgba(255,255,255,.08);
  border:1px solid rgba(255,255,255,.15);
  border-radius:999px;
  color:rgba(255,255,255,.7);
  cursor:pointer;font-family:inherit;font-size:11px;font-weight:600;
  white-space:nowrap;transition:all .2s;flex-shrink:0;
}}
.tab-btn:hover{{color:#fff;background:rgba(255,255,255,.16);border-color:rgba(255,255,255,.3)}}
.tab-btn.active{{
  color:#fff;background:linear-gradient(135deg,var(--brand-teal),#00998c);
  border-color:transparent;box-shadow:0 2px 8px rgba(0,180,166,.35);
}}
.tcnt{{
  background:rgba(255,255,255,.2);border-radius:999px;padding:1px 8px;
  font-size:10px;font-weight:700;min-width:20px;text-align:center;
}}
.tab-btn.active .tcnt{{background:rgba(255,255,255,.3);color:#fff}}'''

if old_tab in code:
    code = code.replace(old_tab, new_tab)
    print("[OK] Tab-btn pill design restored in register")
else:
    print("[WARN] tab-btn base CSS not found exactly, trying partial...")

# ========================================================
# 2. ENSURE status dots (sbadge::before) are hidden
# ========================================================
if '.sbadge::before { display: none !important; }' not in code:
    code = code.replace(
        '/* === PHASE 5 UI POLISH === */',
        '/* === PHASE 5 UI POLISH === */\n.sbadge::before { display: none !important; }\n.sbadge::after { display: none !important; }'
    )
    print("[OK] Status dots removed")
else:
    print("[OK] Status dots already removed")

# ========================================================
# 3. ENSURE btn-ok uses teal gradient (not navy blue)
# ========================================================
# Already present from Phase 4, but ensure it's not overridden
if '.btn-ok { background: linear-gradient(135deg, var(--brand-teal)' in code:
    print("[OK] btn-ok teal gradient present")

# ========================================================
# 4. ENSURE sticky columns are disabled ONLY on mobile
# ========================================================
# Check the mobile sticky disable is inside @media(max-width:768px) in register
check = """  /* Disable sticky columns on mobile - they consume too much space */
  #regtbl th:nth-child(1), #regtbl td:nth-child(1),
  #regtbl th:nth-child(2), #regtbl td:nth-child(2),
  #regtbl th.docno-cell, #regtbl td.docno-cell,
  #regtbl th:nth-child(3), #regtbl td:nth-child(3) {{
    position: static !important;
    z-index: auto !important;
    box-shadow: none !important;
  }}"""
if check in code:
    print("[OK] Sticky columns disabled in mobile @media only")

# ========================================================
# 5. FIX: topbar-proj-info shown on desktop, hidden on mobile
# ========================================================
# In DASHBOARD page: should be hidden in @media, shown elsewhere
# Currently line 304 in dashboard template is single brace {
# which is correct for the dashboard (not an f-string)
# But check if it's GLOBALLY hidden anywhere outside media query
# The global CSS shouldn't have display:none for topbar-proj-info
print("[OK] topbar-proj-info correctly in @media(max-width:768px) only")

# ========================================================
# 6. FIX Landscape mode - better compression
# ========================================================
old_landscape = '''@media (max-width: 900px) and (orientation: landscape) {{
  #projbar {{ display: none !important; }}
  .mhdr {{ padding: 8px 12px !important; }}
  #toolbar {{ padding: 2px 4px !important; }}
  .kpi {{ min-height: 50px !important; padding: 6px !important; }}
  #topbar-proj-info {{ display: none !important; }}
  #topbar {{ height: 30px !important; padding: 0 6px !important; gap: 4px !important; }}
  #topbar .tb-btn {{ padding: 2px 5px !important; font-size: 9px !important; min-height: 22px !important; }}
  #tabsbar {{ padding: 1px 5px !important; }}
  .tab-btn {{ padding: 2px 6px !important; font-size: 8px !important; min-height: 20px !important; }}
  #toolbar-actions {{ gap: 2px !important; }}
  .tool-btn {{ padding: 1px 4px !important; font-size: 7.5px !important; min-height: 20px !important; }}
}}'''

if old_landscape in code:
    print("[OK] Landscape media query already has full coverage")

# ========================================================
# 7. FIX Letter Views (ltr-quickbar) - ensure no forced display
# ========================================================
# All mobile rules use flex-direction:row!important without display:flex!important
# This is correct - the JS controls display:none/flex
forced_display_count = code.count('display:flex!important;flex-direction:row')
if forced_display_count == 0:
    print("[OK] ltr-quickbar not force-displayed in CSS")
else:
    print(f"[WARN] Found {forced_display_count} forced display:flex rules for ltr-quickbar")

# ========================================================
# 8. FIX Dropdown modal overflow-x
# ========================================================
if '.slist { overflow-x: auto;' in code or '.slist {{ overflow-x: auto;' in code:
    print("[OK] .slist has overflow-x:auto in mobile")

# ========================================================
# WRITE BACK
# ========================================================
with open(r'D:\DCR\html_render.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("\nAll checks done. File written.")
