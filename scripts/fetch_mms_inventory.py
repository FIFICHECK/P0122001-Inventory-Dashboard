#!/usr/bin/env python3
"""Fetch P0122001 (SKECHERS) inventory from MMS API and merge into data/inventory_all.csv.

Replaces the Exchange portal CSV download (3-4 day cadence) with MMS live inventory
(same-day updates). MMS API lacks static fields (Category Code/Name, Warehouse ID,
dates, prices) so those are PRESERVED from the last Exchange CSV — only live fields
are updated: StockLevel (merchant + 3PL + consignment), Online Status, Invisible.

Usage: python3 scripts/fetch_mms_inventory.py <accessToken>
  token = value of the `accessToken` cookie after MMS 2.0 login (merchant.shoalter.com)
Writes: data/inventory_all.csv (merged) + data/inventory_mms_meta.json
"""
import csv, json, os, sys, urllib.request, datetime

TOKEN = sys.argv[1] if len(sys.argv) > 1 else os.environ.get('MMS_TOKEN', '')
if not TOKEN:
    print("!! usage: fetch_mms_inventory.py <accessToken>"); sys.exit(1)

API = 'https://merchant-inventory-api.shoalter.com/inventory/api/v2/product-inventory'
STORE_ID = 46383   # P0122001 (SKECHERS) numeric store ID — from MMS UI request body
PAGE_SIZE = 1000
REPO = os.path.expanduser('~/P0122001-Inventory-Dashboard')
os.chdir(REPO)


def fetch_page(n):
    body = json.dumps({"pageNumber": n, "pageSize": PAGE_SIZE,
                       "buCodeList": ["HKTV"], "storeId": [STORE_ID]}).encode()
    req = urllib.request.Request(API, data=body, headers={
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + TOKEN,
        'Accept': 'application/json, text/plain, */*',
    })
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)


# ---- 1. Fetch all pages ----
rows, page, total = [], 1, None
while page <= 100:
    d = fetch_page(page)
    if d.get('code') != 'SUCCESS':
        print(f"!! API error page {page}: {d.get('response')}"); sys.exit(1)
    resp = d['response']
    total = resp.get('totalElement', 0)
    content = resp.get('content', [])
    rows.extend(content)
    if len(rows) >= total or not content:
        break
    page += 1
print(f"fetched {len(rows)} / {total} SKUs from MMS API (storeId {STORE_ID})")

# ---- 2. Build MMS lookup: normalized SKU -> live fields ----
mms = {}
for it in rows:
    sku = it.get('skuId', '')
    if not sku:
        continue
    bu = (it.get('buProductDetail') or [{}])[0]
    qty = int(it.get('merchantInventoryQty') or 0) + int(it.get('tplInventoryQty') or 0) \
        + int(it.get('consignmentInventoryQty') or 0)
    mms[sku] = {
        'stock': qty,
        'online': 'online' if str(bu.get('status', '')).upper() == 'ONLINE' else 'offline',
        'invisible': 'Y' if bu.get('isVisible') is False else 'N',
        'sku_name_en': it.get('skuNameEn', ''),
        'sku_name_ch': it.get('skuNameCh', ''),
        'product_id': it.get('productId', ''),
        'merchant_id': it.get('merchantId', ''),
        'update_time': it.get('updateTime'),
        'is_bundle': it.get('isBundle'),
    }
print(f"MMS lookup keys: {len(mms)}")

# ---- 3. Merge into inventory_all.csv ----
CSV_PATH = 'data/inventory_all.csv'
if not os.path.exists(CSV_PATH):
    print(f"!! {CSV_PATH} not found — run once from Exchange first"); sys.exit(1)

with open(CSV_PATH, encoding='utf-8-sig', newline='') as f:
    reader = csv.DictReader(f)
    fieldnames = reader.fieldnames
    db_rows = list(reader)

updated = 0
missing = 0
for r in db_rows:
    sku = (r.get('Merchant SKU ID') or '').replace('P0122001_S_', '')
    m = mms.get(sku)
    if m:
        r['StockLevel'] = m['stock']
        r['Online Status'] = m['online']
        r['Invisible'] = m['invisible']
        if m['sku_name_en']:
            r['SKU Name'] = m['sku_name_en']
        if m['sku_name_ch']:
            r['SKU Name (Chi)'] = m['sku_name_ch']
        updated += 1
    else:
        missing += 1
print(f"merged: {updated} updated, {missing} SKUs not found in MMS response")

with open(CSV_PATH, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=fieldnames)
    w.writeheader()
    w.writerows(db_rows)

# ---- 4. Write reports/inventory_report_<date>_<time>.csv (Report tab download table) ----
now = datetime.datetime.now()
date_raw = now.strftime('%Y%m%d')
time_raw = now.strftime('%H%M')
rep_path = f'reports/inventory_report_{date_raw}_{time_raw}.csv'
with open(rep_path, 'w', encoding='utf-8-sig', newline='') as f:
    f.write('﻿﻿Stock Level Summary Report\n')
    f.write('Merchant ID,P0122001\n')
    f.write('Merchant Name,SKECHERS\n')
    f.write(f'Date,{now.strftime("%Y/%m/%d")}\n')
    f.write(',,,,,,,,,,,,,,,Packing Dimension\n')
    f.write(open(CSV_PATH, encoding='utf-8-sig').read())
print(f"written {rep_path} ({os.path.getsize(rep_path)} bytes)")

meta = {
    'source': 'mms-api',
    'fetched_at': datetime.datetime.now().isoformat(timespec='seconds'),
    'store_id': STORE_ID,
    'total_skus': total,
    'updated_skus': updated,
    'missing_skus': missing,
}
with open('data/inventory_mms_meta.json', 'w', encoding='utf-8') as f:
    json.dump(meta, f, ensure_ascii=False, indent=2)
print(f"written {CSV_PATH} + data/inventory_mms_meta.json ({updated} live fields merged)")
