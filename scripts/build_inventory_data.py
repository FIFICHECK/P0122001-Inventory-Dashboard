#!/usr/bin/env python3
"""P0122001 Inventory Data builder.

Reads data/inventory_all.csv (40,436 size-variant SKU rows) and aggregates at
STYLE level (Merchant Product ID, e.g. 104624-GRY) → data/inventory_data.json.

Output shape:
{
  "report_date": "2026-08-21",
  "download_time": "2026-08-24 00:20",
  "styles": [
    {
      "style": "104624-GRY",            # Merchant Product ID
      "sku": "P0122001_S_104624-GRY",   # style-level display SKU
      "link_sku": "P0122001_S_104624-GRY-7",  # first online size-variant SKU for HKTVmall link
      "name_en": "...", "name_chi": "...",
      "brand": "SKECHERS",
      "category": "AA...", "category_name": "...",
      "stock": 43, "sizes": 5, "online_sizes": 2,
      "online": true/false, "invisible": true/false, "foos": true/false,
      "rsp": 699.0, "psp": 279.0,
      "size_list": ["5","6","7","8","9"],   # sizes with stock>0
      "zero_size_list": [...],
      "create_date": "25-Oct-2022",
      "status": "high|normal|low|zero"  # based on total stock
    }, ...
  ],
  "counts": {"total": 8102, "online": 1136, "zero": 6829, "low": N, "normal": N, "high": N, "invisible": N, "foos": N, "new": N}
}
"""
import csv, json, os, re, sys
from collections import Counter

REPO = os.path.expanduser('~/P0122001-Inventory-Dashboard')
os.chdir(REPO)

def load_rows():
    rows = []
    with open('data/inventory_all.csv', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for r in reader:
            sku = (r.get('Merchant SKU ID') or '').strip()
            if not sku:
                continue
            rows.append(r)
    return rows

def status_of(stock):
    if stock <= 0: return 'zero'
    if stock < 10: return 'low'
    if stock < 50: return 'normal'
    return 'high'

def main():
    rows = load_rows()
    print("raw rows:", len(rows))
    # report date: prefer MMS live meta (fetched today), else Exchange raw CSV
    report_date = ''
    import json as _json, os as _os
    if _os.path.exists('data/inventory_mms_meta.json'):
        _m = _json.load(open('data/inventory_mms_meta.json'))
        if _m.get('source') == 'mms-api' and _m.get('fetched_at'):
            report_date = _m['fetched_at'][:10]
    if not report_date:
        import glob
        for f in sorted(glob.glob('reports/inventory_report_*.csv')):
            head = open(f, encoding='utf-8-sig').read().split('\n')
            for ln in head[:6]:
                if ln.startswith('Date,'):
                    report_date = ln.split(',', 1)[1].strip()
                    break
            if report_date:
                break
    # fallback: YYYY/MM/DD
    m = re.search(r'(\d{4})/(\d{2})/(\d{2})', report_date)
    if m:
        report_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    print("report date:", report_date)

    styles = {}
    for r in rows:
        style = (r.get('Merchant Product ID') or '').strip()
        if not style:
            continue
        # Exclude non-SKECHERS test/sample products (e.g. BIOTHERM Z9000033-AST 試用三件套裝)
        brand_en = (r.get('Brand Name (EN)') or '').strip()
        brand_chi = (r.get('Brand Name (CHI)') or '').strip()
        if brand_en and brand_en.upper() != 'SKECHERS' and brand_chi and brand_chi.upper() != 'SKECHERS':
            continue
        e = styles.setdefault(style, {
            'style': style,
            'sku': 'P0122001_S_' + style,
            'name_en': (r.get('SKU Name') or '').strip(),
            'name_chi': (r.get('SKU Name (Chi)') or '').strip(),
            'brand': (r.get('Brand Name (EN)') or r.get('Brand Name (CHI)') or '').strip(),
            'category': (r.get('Primary Category Code') or '').strip(),
            'category_name': (r.get('Primary Category Name (CHI)') or '').strip(),
            'stock': 0, 'sizes': 0, 'online_sizes': 0,
            'online': False, 'invisible': False, 'foos': False,
            'rsp': None, 'psp': None,
            'size_list': [], 'zero_size_list': [],
            'create_date': (r.get('Create Date') or '').strip(),
            'link_sku': None,
        })
        sku_id = (r.get('Merchant SKU ID') or '').strip()
        # extract size suffix (after style code)
        size = ''
        if sku_id.startswith('P0122001_S_' + style + '-'):
            size = sku_id[len('P0122001_S_' + style + '-'):]
        try:
            stock = int((r.get('StockLevel') or '0').strip() or 0)
        except:
            stock = 0
        online = (r.get('Online Status') or '').strip().lower() == 'online'
        invisible = (r.get('Invisible') or '').strip().upper() == 'Y'
        foos = (r.get('Force Out Of Stock') or '').strip().upper() == 'Y'
        e['sizes'] += 1
        e['stock'] += stock
        if online:
            e['online_sizes'] += 1
            e['online'] = True
            if not e['link_sku']:
                e['link_sku'] = sku_id
        if invisible: e['invisible'] = True
        if foos: e['foos'] = True
        if stock > 0:
            if size: e['size_list'].append(size)
        else:
            if size: e['zero_size_list'].append(size)
        rsp = (r.get('Original Price') or '').strip()
        psp = (r.get('Discount Price') or '').strip()
        try:
            if rsp and e['rsp'] is None: e['rsp'] = float(rsp)
        except: pass
        try:
            if psp and e['psp'] is None: e['psp'] = float(psp)
        except: pass
        if not e['create_date'] and (r.get('Create Date') or '').strip():
            e['create_date'] = (r.get('Create Date') or '').strip()

    # post-process: sort size lists numerically, compute status
    slist = []
    for e in styles.values():
        e['size_list'] = sorted(set(e['size_list']), key=lambda x: (len(x), x))
        e['zero_size_list'] = sorted(set(e['zero_size_list']), key=lambda x: (len(x), x))
        e['status'] = status_of(e['stock'])
        if not e['link_sku']:
            e['link_sku'] = e['sku']  # fallback
        slist.append(e)
    print("styles:", len(slist))

    counts = {
        'total': len(slist),
        'online': sum(1 for e in slist if e['online']),
        'zero': sum(1 for e in slist if e['status'] == 'zero'),
        'low': sum(1 for e in slist if e['status'] == 'low'),
        'normal': sum(1 for e in slist if e['status'] == 'normal'),
        'high': sum(1 for e in slist if e['status'] == 'high'),
        'invisible': sum(1 for e in slist if e['invisible']),
        'foos': sum(1 for e in slist if e['foos']),
        'total_stock': sum(e['stock'] for e in slist),
    }
    # new SKUs: created within last 14 days (parse dd-MMM-yyyy)
    from datetime import datetime, timedelta
    today = datetime.now()
    new_cutoff = today - timedelta(days=14)
    nnew = 0
    for e in slist:
        try:
            d = datetime.strptime(e['create_date'], '%d-%b-%Y')
            if d >= new_cutoff:
                e['is_new'] = True
                nnew += 1
            else:
                e['is_new'] = False
        except:
            e['is_new'] = False
    counts['new'] = nnew
    print("counts:", counts)

    out = {
        'report_date': report_date,
        'counts': counts,
        'styles': slist,
    }
    with open('data/inventory_data.json', 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False)
    print("wrote data/inventory_data.json:", os.path.getsize('data/inventory_data.json'), "bytes")

if __name__ == '__main__':
    main()
