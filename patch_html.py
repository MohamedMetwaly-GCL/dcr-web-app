import os

with open("html_render.py", "r", encoding="utf-8") as f:
    c = f.read()

c = c.replace(
    "return ['description','fileLocation','status'].includes(role)||['description','filelocation','status'].includes(key);",
    "return ['description','status'].includes(role)||['description','status'].includes(key);"
)

with open("html_render.py", "w", encoding="utf-8") as f:
    f.write(c)
