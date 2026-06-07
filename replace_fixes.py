with open('D:\\DCR\\html_render.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_topbar_css = '''#topbar-proj-info .pf-val { font-size: 15px; color: #ffffff; font-weight: 800; white-space: nowrap; }
#topbar-proj-info .pf.primary:first-child .pf-val { font-size: 14px; color: var(--brand-teal); }'''
new_topbar_css = '''#topbar-proj-info .pf-val { font-size: 15px; color: #ffffff; font-weight: 800; white-space: nowrap; max-width: none !important; }
#topbar-proj-info .pf.primary .pf-val { font-size: 14px; color: var(--brand-teal); }'''

old_projbar_css = '''#projbar .pf-val { white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important; }'''
new_projbar_css = '''#projbar .pf-val { white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important; max-width: none !important; }'''

old_regtbl_css = '''/* 1. Remove Vertical Borders & Force Zebra Striping */
table, .dt-tbl, #regtbl { border-collapse: collapse !important; }
table td, table th, .dt-tbl td, .dt-tbl th, #regtbl td, #regtbl th { border-left: none !important; border-right: none !important; border-bottom: 1px solid #e2e8f0 !important; }'''
new_regtbl_css = '''/* 1. Remove Vertical Borders & Force Zebra Striping */
table, .dt-tbl, #regtbl { border-collapse: collapse !important; }
table td, table th, .dt-tbl td, .dt-tbl th, #regtbl td, #regtbl th { border-left: none !important; border-right: none !important; border-bottom: 1px solid #e2e8f0 !important; }
#regtbl thead tr:first-child th { border-bottom: none !important; }'''

old_dark_css = '''body.dark #regtbl tbody tr:hover { background-color: #334155; }'''
new_dark_css = '''body.dark #regtbl tbody tr:hover { background-color: #334155; }

/* Overdue Styling for rows */
#regtbl tbody tr.ov td { background-color: inherit !important; }
#regtbl tbody tr.ov, #regtbl tbody tr.ov:nth-child(even) { background-color: #fef2f2 !important; }
#regtbl tbody tr.ov:hover { background-color: #ffe7e7 !important; }
body.dark #regtbl tbody tr.ov, body.dark #regtbl tbody tr.ov:nth-child(even) { background-color: #3a231f !important; }
body.dark #regtbl tbody tr.ov:hover { background-color: #4a2a24 !important; }'''

old_logo_py = '''    logo_ver = str(int(os.path.getmtime(logo_file))) + "_v3_" + str(int(time.time())) if os.path.exists(logo_file) else "1"'''
new_logo_py = '''    logo_ver = str(int(os.path.getmtime(logo_file))) if os.path.exists(logo_file) else "1"'''

c1 = code.count(old_topbar_css)
c2 = code.count(old_projbar_css)
c3 = code.count(old_regtbl_css)
c4 = code.count(old_dark_css)
c5 = code.count(old_logo_py)

if c1==1 and c2==1 and c3==1 and c4==1 and c5==1:
    code = code.replace(old_topbar_css, new_topbar_css)
    code = code.replace(old_projbar_css, new_projbar_css)
    code = code.replace(old_regtbl_css, new_regtbl_css)
    code = code.replace(old_dark_css, new_dark_css)
    code = code.replace(old_logo_py, new_logo_py)
    with open('D:\\DCR\\html_render.py', 'w', encoding='utf-8') as f:
        f.write(code)
    print("SUCCESS")
else:
    print(f"FAILED: counts = {c1}, {c2}, {c3}, {c4}, {c5}")
