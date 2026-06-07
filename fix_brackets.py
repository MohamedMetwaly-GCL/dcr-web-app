import re

with open("html_render.py", "r", encoding="utf-8") as f:
    c = f.read()

# Fix openDistributionMatrix
start_idx = c.find("async function openDistributionMatrix(pid) {")
if start_idx != -1:
    end_idx = c.find("async function syncDriveLinks(btn) {{", start_idx)
    if end_idx != -1:
        block = c[start_idx:end_idx]
        # Replace { with {{ and } with }}, EXCEPT if they are already {{ or }}
        # We can just un-double everything, then double everything.
        block = block.replace("{{", "{").replace("}}", "}")
        block = block.replace("{", "{{").replace("}", "}}")
        c = c[:start_idx] + block + c[end_idx:]

# Fix loadDailyDigest
start_idx2 = c.find("async function loadDailyDigest() {")
if start_idx2 != -1:
    end_idx2 = c.find("async function loadExecutive() {{", start_idx2)
    if end_idx2 != -1:
        block2 = c[start_idx2:end_idx2]
        block2 = block2.replace("{{", "{").replace("}}", "}")
        block2 = block2.replace("{", "{{").replace("}", "}}")
        c = c[:start_idx2] + block2 + c[end_idx2:]

with open("html_render.py", "w", encoding="utf-8") as f:
    f.write(c)
