#!/usr/bin/env python3
"""Inject sales_trend_data.js content into index.html's inline salesTrendData declaration."""
import re, os, json

REPO = os.path.expanduser('~/P0122001-Inventory-Dashboard')
os.chdir(REPO)

# Read generated data
js = open('data/sales_trend_data.js', encoding='utf-8').read()
# Extract the JSON part (after "const salesTrendData = ")
m = re.search(r'const salesTrendData = (\{.*\});', js, re.S)
if not m:
    print("!! cannot extract JSON from sales_trend_data.js")
    raise SystemExit(1)
data_json = m.group(1)
# validate
json.loads(data_json)
print("extracted JSON:", len(data_json), "chars")

h = open('index.html', encoding='utf-8').read()

# Replace the inline declaration block (match any preceding comment line)
old = re.search(r'    // Sales Trend Data[^\n]*\n    const salesTrendData = \{.*?\};\n', h, re.S)
if not old:
    print("!! inline salesTrendData declaration not found")
    raise SystemExit(1)

new_block = "    // Sales Trend Data (from the GP Report monthly SKU data)\n    const salesTrendData = " + data_json + ";\n"
h = h[:old.start()] + new_block + h[old.end():]
open('index.html', 'w', encoding='utf-8').write(h)
print("injected. new index.html:", len(h), "chars")

# verify (best-effort; /tmp always exists)
import subprocess
blocks = re.findall(r'<script>(.*?)</script>', h, re.S)
for i, b in enumerate(blocks):
    p = f'/tmp/inj_block{i}.js'
    try:
        open(p, 'w').write(b)
    except OSError as e:
        print("warn: cannot write", p, e)
print("script blocks:", len(blocks))
