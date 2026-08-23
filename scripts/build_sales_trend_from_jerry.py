#!/usr/bin/env python3
"""Build P0122001 sales_trend_data.js from jerry-dashboard's sku_data_full.json.

jerry-dashboard (https://fificheck.github.io/jerry-dashboard/) has 8 months of
P0122001 SKU-level monthly data (2026-01 .. 2026-08). This script:
1. Downloads sku_data_full.json (or uses local cache)
2. Extracts P0122001_S_* records
3. Builds the salesTrendData structure the Sales Trend tab expects:
   - gmv_by_month (8 months GMV)
   - gmv_by_sku_monthly / qty_by_sku_monthly (top 50 SKUs, monthly series)
   - gmv_by_brand_monthly (SKECHERS single brand)
   - available_months, sku_name_map, sku_brand_map, summary
   - gmv_by_date / gmv_by_date_hour left EMPTY (jerry-dashboard has no daily data;
     daily series will be populated by the cron job once order reports accumulate)

NOTE: The Sales Trend tab's chart buttons 1,3,4,5,7 (daily) will show empty
until daily order reports exist. Buttons 2,6,8,9,10 (monthly) work now.
"""
import json, os, re, urllib.request
from collections import defaultdict

REPO = os.path.expanduser('~/P0122001-Inventory-Dashboard')
os.chdir(REPO)
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36'}

CACHE = '/tmp/p0122001/sku_data_full.json'

def fetch():
    if os.path.exists(CACHE) and os.path.getsize(CACHE) > 1000000:
        print("using cache", CACHE)
        return json.load(open(CACHE))
    req = urllib.request.Request('https://fificheck.github.io/jerry-dashboard/sku_data_full.json', headers=UA)
    d = json.load(urllib.request.urlopen(req, timeout=120))
    json.dump(d, open(CACHE, 'w'))
    return d

def main():
    full = fetch()
    print("total records:", len(full))
    # P0122001 records
    p = [e for e in full if str(e.get('sc','')).startswith('P0122001_S_')]
    print("P0122001 records:", len(p))
    if not p:
        print("NO P0122001 data — abort")
        return

    # Group by month
    monthly_gmv = defaultdict(float)
    monthly_qty = defaultdict(int)
    sku_monthly_gmv = defaultdict(lambda: defaultdict(float))   # sku -> month -> gmv
    sku_monthly_qty = defaultdict(lambda: defaultdict(int))     # sku -> month -> qty
    sku_names = {}
    sku_brand = {}
    for e in p:
        m = e.get('m')
        gmv = e.get('gmv') or 0
        qty = e.get('qty') or 0
        sn = (e.get('sn') or '').strip()
        if m:
            monthly_gmv[m] += gmv
            monthly_qty[m] += qty
            sku_monthly_gmv[e['sc']][m] += gmv
            sku_monthly_qty[e['sc']][m] += qty
        if sn and e['sc'] not in sku_names:
            sku_names[e['sc']] = sn
        sku_brand[e['sc']] = 'SKECHERS'

    months = sorted(monthly_gmv.keys())
    print("months:", months)

    # Summary
    total_gmv = sum(monthly_gmv.values())
    total_orders = sum(monthly_qty.values())
    # this_month / last_month / month_before_last
    def month_label(m):
        return f"{int(m[5:7])}月 {m[0:4]}"
    summary = {}
    if len(months) >= 3:
        summary = {
            'this_month': {'label': month_label(months[-1]), 'gmv': round(monthly_gmv[months[-1]], 2), 'orders': monthly_qty[months[-1]], 'avg': round(monthly_gmv[months[-1]] / monthly_qty[months[-1]], 2) if monthly_qty[months[-1]] else 0},
            'last_month': {'label': month_label(months[-2]), 'gmv': round(monthly_gmv[months[-2]], 2), 'orders': monthly_qty[months[-2]], 'avg': round(monthly_gmv[months[-2]] / monthly_qty[months[-2]], 2) if monthly_qty[months[-2]] else 0},
            'month_before_last': {'label': month_label(months[-3]), 'gmv': round(monthly_gmv[months[-3]], 2), 'orders': monthly_qty[months[-3]], 'avg': round(monthly_gmv[months[-3]] / monthly_qty[months[-3]], 2) if monthly_qty[months[-3]] else 0},
        }
    summary['total_gmv'] = round(total_gmv, 2)
    summary['total_orders'] = total_orders
    summary['avg_order_value'] = round(total_gmv / total_orders, 2) if total_orders else 0

    # Top 50 SKUs by total GMV
    sku_totals = {sku: sum(d.values()) for sku, d in sku_monthly_gmv.items()}
    top_skus = sorted(sku_totals, key=lambda x: -sku_totals[x])[:50]

    gmv_by_month = {'labels': months, 'data': [round(monthly_gmv[m], 2) for m in months]}
    qty_by_month = {'labels': months, 'data': [monthly_qty[m] for m in months]}
    # B0961005-format: data is Month×SKU 2D array (each row = one month's values across SKUs)
    sku_monthly_data = [[round(sku_monthly_gmv[s].get(m, 0), 2) for s in top_skus] for m in months]
    qty_sku_monthly_data = [[sku_monthly_qty[s].get(m, 0) for s in top_skus] for m in months]
    gmv_sku_monthly = {
        'labels': months,
        'skus': top_skus,
        'data': sku_monthly_data,
    }
    gmv_brand_monthly = {
        'labels': months,
        'brands': ['SKECHERS'],
        'data': [[round(monthly_gmv[m], 2)] for m in months],
    }

    out = {
        'gmv_by_date': {'labels': [], 'data': []},
        'gmv_by_month': gmv_by_month,
        'gmv_by_hour': {'labels': [], 'hours': [f"{h:02d}" for h in range(24)], 'data': {}},
        'gmv_by_day_of_week': {'labels': [], 'data': []},
        'gmv_by_sku_daily': {'labels': [], 'skus': []},
        'qty_by_sku_daily': {'labels': [], 'skus': []},
        'gmv_by_sku_monthly': gmv_sku_monthly,
        'qty_by_sku_monthly': gmv_sku_monthly,
        'gmv_by_brand_daily': {'labels': [], 'brands': []},
        'gmv_by_brand_monthly': gmv_brand_monthly,
        'sku_name_map': sku_names,
        'sku_product_name_map': sku_names,
        'sku_brand_map': sku_brand,
        'available_months': months,
        'summary': summary,
    }

    out_js = "// Auto-generated sales trend data (from jerry-dashboard sku_data_full.json)\n// Generated: %s\n// Source: jerry-dashboard monthly SKU data (P0122001)\n\nconst salesTrendData = %s;\n" % (
        __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        json.dumps(out, ensure_ascii=False)
    )
    open('data/sales_trend_data.js', 'w', encoding='utf-8').write(out_js)
    print("wrote data/sales_trend_data.js:", os.path.getsize('data/sales_trend_data.js'), "bytes")
    print("months:", months)
    print("summary:", json.dumps(summary, ensure_ascii=False)[:300])
    print("top SKUs:", len(top_skus))

if __name__ == '__main__':
    main()
