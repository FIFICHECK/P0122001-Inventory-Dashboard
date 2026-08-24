#!/usr/bin/env python3
"""Build data/report_manifest.json for the Report tab download tables.

Scans reports/ for inventory CSVs and reports/order_reports/ for order XLSX files,
computes date/time/SKU-count/GMV where derivable, and emits:
{
  "inventory_reports": [{"date": "2026-08-21", "time": "1200", "skus": 40436, "file": "inventory_report_20260821_1200.csv", "size": 12252314}],
  "order_reports": [{"date": "2026-08-23", "time": "235959", "gmv": 12345.67, "file": "ECOM-...xlsx", "size": 87184}],
  "updated_at": "..."
}
"""
import os, re, glob, json, csv, datetime

REPO = os.path.expanduser('~/P0122001-Inventory-Dashboard')
os.chdir(REPO)

def parse_inventory_file(name):
    m = re.match(r'inventory_report_(\d{8})_(\d{4})\.csv$', name)
    if not m:
        return None
    date_raw, time_raw = m.group(1), m.group(2)
    date = f"{date_raw[:4]}-{date_raw[4:6]}-{date_raw[6:8]}"
    # count SKU rows (strip 5 header lines)
    path = os.path.join('reports', name)
    try:
        lines = open(path, encoding='utf-8-sig', errors='replace').read().split('\n')
        # data rows: after header line (index 5). Count non-empty
        skus = sum(1 for l in lines[6:] if l.strip())
    except Exception:
        skus = 0
    return {'date': date, 'time': time_raw, 'skus': skus, 'file': name, 'size': os.path.getsize(path)}

def parse_order_file(name):
    m = re.match(r'ECOM-MMSNG_DAILY_ORDER_P0122001_(\d{8})(\d{6})\.xlsx$', name)
    if not m:
        m = re.match(r'ECOM-MMSNG_DAILY_ORDER_P0122001_(\d{8})(\d{6})\.xls$', name)
    if not m:
        return None
    date_raw, time_raw = m.group(1), m.group(2)
    date = f"{date_raw[:4]}-{date_raw[4:6]}-{date_raw[6:8]}"
    path = os.path.join('reports/order_reports', name)
    return {'date': date, 'time': time_raw, 'gmv': None, 'file': name, 'size': os.path.getsize(path)}

inventory = []
for f in sorted(glob.glob('reports/inventory_report_*.csv')):
    e = parse_inventory_file(os.path.basename(f))
    if e:
        inventory.append(e)
inventory.sort(key=lambda x: x['date'], reverse=True)

order = []
for f in sorted(glob.glob('reports/order_reports/*.xlsx') + glob.glob('reports/order_reports/*.xls')):
    e = parse_order_file(os.path.basename(f))
    if e:
        order.append(e)
order.sort(key=lambda x: (x['date'], x['time']), reverse=True)

# GMV by SKU By Month report (GP Report export)
gmv_reports = []
for f in sorted(glob.glob('reports/P0122001_GMV_By_SKU_By_Month_*.xlsx')):
    name = os.path.basename(f)
    m = re.match(r'P0122001_GMV_By_SKU_By_Month_(\d{6})-(\d{6})\.xlsx$', name)
    period = f"{m.group(1)[:4]}-{m.group(1)[4:6]} 至 {m.group(2)[:4]}-{m.group(2)[4:6]}" if m else '—'
    gmv_reports.append({'period': period, 'file': name, 'size': os.path.getsize(f)})
gmv_reports.sort(key=lambda x: x['period'], reverse=True)

out = {
    'inventory_reports': inventory,
    'order_reports': order,
    'gmv_reports': gmv_reports,
    'updated_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
}
json.dump(out, open('data/report_manifest.json', 'w'), ensure_ascii=False, indent=1)
print("inventory reports:", len(inventory), "| order reports:", len(order), "| gmv reports:", len(gmv_reports))
for e in inventory[:3]:
    print("  inv:", e['date'], e['time'], e['skus'], 'SKUs')
for e in order[:3]:
    print("  ord:", e['date'], e['time'], e['size'], 'bytes')
for e in gmv_reports:
    print("  gmv:", e['period'], e['size'], 'bytes')
