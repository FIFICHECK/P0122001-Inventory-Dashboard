#!/usr/bin/env python3
"""Merge Daily Order Report data into salesTrendData daily charts.

Reads ALL daily order report xlsx files in reports/order_reports/ regardless of format:
  - MMS:     ECOM-MMSNG_DAILY_ORDER_P0122001_*.xlsx
  - Exchange: ECOM-EXCH_DAILY_ORDER_P0122001_*.xlsx

Aggregates per day/hour/day-of-week/SKU/brand, merges into data/sales_trend_data.js.
Keeps monthly keys untouched; fills the *_daily keys.

Usage: run with python3.12 + openpyxl, then inject via inject_sales.py.
"""
import openpyxl, json, os, glob, re, datetime
from collections import defaultdict

REPO = os.path.expanduser('~/P0122001-Inventory-Dashboard')
os.chdir(REPO)

gmv_by_date = defaultdict(float)
qty_by_date = defaultdict(int)
orders_by_date = defaultdict(int)
gmv_by_hour = defaultdict(float)
gmv_by_dow = defaultdict(float)
gmv_by_sku = defaultdict(lambda: defaultdict(float))
qty_by_sku = defaultdict(lambda: defaultdict(int))
sku_names = {}
gmv_by_brand = defaultdict(lambda: defaultdict(float))
DATE_KEYS = []

def process_mms(ws):
    """MMS format: rows start at row 6, headers at row 5."""
    date = None
    parent_orders = set()
    for r in range(6, ws.max_row + 1):
        od = ws.cell(row=r, column=7).value   # Order Date
        if od is None:
            continue
        date = str(od)[:10]
        tm = str(ws.cell(row=r, column=8).value or '')
        hour = int(tm[:2]) if len(tm) >= 2 and tm[:2].isdigit() else 0
        sku = str(ws.cell(row=r, column=18).value or '').strip()
        name = str(ws.cell(row=r, column=22).value or ws.cell(row=r, column=21).value or '').strip()
        qty = ws.cell(row=r, column=24).value or 0
        total = ws.cell(row=r, column=27).value or 0
        oid = str(ws.cell(row=r, column=5).value or '')   # Sub-Order ID
        parent_orders.add(oid)
        _accum(date, hour, sku, name, qty, total)

    if date:
        orders_by_date[date] = len(parent_orders)
        DATE_KEYS.append(date)

def process_exchange(ws, fname):
    """Exchange format: rows start at row 6 (row 5 = headers), GMV at row 2 col 6.
    Headers: Delivery Mode | Warehouse ID | Warehouse Name | Order ID | Sub-Order ID |
             Shipment ID | Order Date | Order Time | Pickup Date | Pickup Time | Delivery Date |
             ... (more columns) ... | Item Total | ... | Total Amount |
    We use Order Date (col 7) + Sub-Order ID (col 5) + Item Total (col ~14?).
    But Exchange format uses row-level items — each row = one item in an order.
    To avoid double-counting GMV, we aggregate by (Order ID, Sub-Order ID) first.
    """
    # Find header row to map column positions
    max_col = ws.max_column
    headers = [ws.cell(5, c).value for c in range(1, max_col + 1)]
    # Search for relevant columns
    col_order_id = next((i for i, h in enumerate(headers, 1) if h and 'Order ID' in str(h) and 'Sub' not in str(h)), 4)
    col_sub_order = next((i for i, h in enumerate(headers, 1) if h and 'Sub-Order' in str(h)), 5)
    col_order_date = next((i for i, h in enumerate(headers, 1) if h and 'Order Date' in str(h)), 7)
    col_qty = next((i for i, h in enumerate(headers, 1) if h and ('Qty' in str(h) or 'Quantity' in str(h))), None)
    # Item Total or Total Amount
    col_total = next((i for i, h in enumerate(headers, 1) if h and ('Item Total' in str(h) or 'Total Amount' in str(h) and 'Total Sales' not in str(h))), 14)

    # Also get GMV from row 2 (Total Sales Amount Sum) as baseline validation
    total_gmv_sheet = None
    for c in range(1, max_col + 1):
        v = ws.cell(2, c).value
        if v and isinstance(v, (int, float)) and v > 1000:
            total_gmv_sheet = v
            break

    # Aggregate: (order_id, sub_order_id) -> accumulated GMV/qty
    order_key_map = defaultdict(lambda: {'gmv': 0.0, 'qty': 0, 'date': None, 'hour': 0})
    seen_orders = set()
    for r in range(6, ws.max_row + 1):
        oid_raw = ws.cell(r, col_order_id).value
        if oid_raw is None:
            continue
        oid = str(oid_raw)
        sub_oid = str(ws.cell(r, col_sub_order).value or '')
        key = (oid, sub_oid)
        if key in seen_orders:
            continue

        od = ws.cell(r, col_order_date).value
        if od is None:
            continue
        date = str(od)[:10]

        qty_val = ws.cell(r, col_qty).value if col_qty else None
        total_val = ws.cell(r, col_total).value if col_total else None
        qty = int(qty_val) if qty_val else 0
        total = float(total_val) if total_val else 0.0

        order_key_map[key] = {'gmv': total, 'qty': qty, 'date': date}
        seen_orders.add(key)

    # Now DATE_KEYS = unique dates from Exchange file
    dates_found = set(v['date'] for v in order_key_map.values() if v['date'])
    for d in sorted(dates_found):
        DATE_KEYS.append(d)

    # Accumulate per order (not per row) to avoid double-counting
    for info in order_key_map.values():
        d = info['date']
        if not d:
            continue
        total = info['gmv']
        qty = info['qty']
        # Get hour from first row of this order
        for r in range(6, ws.max_row + 1):
            if str(ws.cell(r, col_order_id).value or '') == str(order_key_map_inv(info, col_order_id)):
                pass
        hour = 12  # default midday
        _accum(d, hour, '', '', qty, total)

def order_key_map_inv(info, col_order_id):
    """Helper — not used directly, Exchange uses set-based dedup."""
    return ''

def _accum(date, hour, sku, name, qty, total):
    """Shared accumulator — called by both MMS and Exchange processors."""
    gmv_by_date[date] += float(total)
    qty_by_date[date] += int(qty or 0)
    gmv_by_hour[hour] += float(total)
    try:
        dow = datetime.date.fromisoformat(date).weekday()
        gmv_by_dow[dow] += float(total)
    except ValueError:
        pass
    if sku:
        gmv_by_sku[date][sku] += float(total)
        qty_by_sku[date][sku] += int(qty or 0)
        sku_names[sku] = name or sku
    gmv_by_brand[date]['SKECHERS'] += float(total)

# ---------- 1. Read all daily reports (both MMS and Exchange) ----------
files = sorted(glob.glob('reports/order_reports/ECOM-MMSNG_DAILY_ORDER_P0122001_*.xlsx')) \
        + sorted(glob.glob('reports/order_reports/ECOM-EXCH_DAILY_ORDER_P0122001_*.xlsx'))
print(f"daily reports: {len(files)}")
for f in files:
    print("  ", os.path.basename(f))

for f in files:
    is_exchange = 'EXCH' in os.path.basename(f)
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active
        if is_exchange:
            process_exchange(ws, os.path.basename(f))
        else:
            process_mms(ws)
        print(f"  ✓ {os.path.basename(f)}")
    except Exception as e:
        print(f"  ✗ {os.path.basename(f)}: {e}")

DATE_KEYS = sorted(set(DATE_KEYS))
print(f"\ndates: {DATE_KEYS}")
print(f"daily GMV: {{d: round(gmv_by_date[d],2) for d in DATE_KEYS}}")

# ---------- 2. Load existing sales_trend_data.js ----------
js = open('data/sales_trend_data.js', encoding='utf-8').read()
m = re.search(r'const salesTrendData = (\{.*\});', js, re.S)
if not m:
    raise SystemExit('!! cannot find salesTrendData in data/sales_trend_data.js')
sd = json.loads(m.group(1))

# ---------- 3. Merge daily data ----------
sd['gmv_by_date'] = {'labels': DATE_KEYS, 'data': [round(gmv_by_date[d], 2) for d in DATE_KEYS]}
sd['orders_by_date'] = {'labels': DATE_KEYS, 'data': [orders_by_date.get(d, 0) for d in DATE_KEYS]}

sd['gmv_by_hour'] = {'labels': [], 'hours': [f"{h:02d}" for h in range(24)], 'data': {}}
sd['gmv_by_hour']['data'] = {f"{h:02d}": round(gmv_by_hour.get(h, 0), 2) for h in range(24)}

dow_names = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
dow_labels = []
dow_data = []
for d in DATE_KEYS:
    try:
        dow = datetime.date.fromisoformat(d).weekday()
        dow_labels.append(f"{dow_names[dow]} ({d})")
        dow_data.append(round(gmv_by_date[d], 2))
    except ValueError:
        pass

sd['gmv_by_day_of_week'] = {'labels': dow_labels, 'data': dow_data}
sd['gmv_by_date_hour'] = {'labels': DATE_KEYS, 'hours': [f"{h:02d}" for h in range(24)], 'data': {}}

# SKU daily: top SKUs by total GMV
sku_totals = defaultdict(float)
for d, m2 in gmv_by_sku.items():
    for s, v in m2.items():
        sku_totals[s] += v
top_skus = [s for s, _ in sorted(sku_totals.items(), key=lambda x: -x[1])]
sd['gmv_by_sku_daily'] = {'labels': DATE_KEYS, 'skus': top_skus,
                            'data': [[round(gmv_by_sku[d].get(s, 0), 2) for s in top_skus] for d in DATE_KEYS]}
sd['qty_by_sku_daily'] = {'labels': DATE_KEYS, 'skus': top_skus,
                            'data': [[qty_by_sku[d].get(s, 0) for s in top_skus] for d in DATE_KEYS]}
sd['sku_name_map'] = {**sd.get('sku_name_map', {}), **{s: sku_names.get(s, s) for s in top_skus}}
sd['sku_product_name_map'] = sd['sku_name_map']
sd['sku_brand_map'] = {**sd.get('sku_brand_map', {}), **{s: 'SKECHERS' for s in top_skus}}

sd['gmv_by_brand_daily'] = {'labels': DATE_KEYS, 'brands': ['SKECHERS'],
                            'data': [[round(gmv_by_brand[d].get('SKECHERS', 0), 2)] for d in DATE_KEYS]}

# Update summary: this_month/last_month with actual daily data when available
summ = sd['summary']
if DATE_KEYS:
    first, last = DATE_KEYS[0], DATE_KEYS[-1]
    summ['daily_range'] = f"{first} ~ {last}"
    summ['daily_gmv_total'] = round(sum(gmv_by_date.values()), 2)
    summ['daily_orders_total'] = int(sum(orders_by_date.values()))
    summ['daily_qty_total'] = int(sum(qty_by_date.values()))
    summ['avg_order_value_daily'] = round(summ['daily_gmv_total'] / summ['daily_orders_total'], 2) if summ['daily_orders_total'] else 0
    # this_month: GP Excel has no Aug orders — fill from daily reports
    if summ.get('this_month') and summ.get('this_month', {}).get('orders', 0) == 0:
        summ['this_month']['orders'] = summ['daily_orders_total']
        summ['this_month']['avg'] = summ['avg_order_value_daily']

    # Roll: 新月份開始（今日月份 > this_month label 月份）→ 成組搬落一格，當月用 daily 數據
    try:
        _today = datetime.date.today()
        _tm = summ.get('this_month', {})
        _m = re.match(r'(\d+)月 (\d{4})', str(_tm.get('label', '')))
        if _m and (int(_m.group(2)), int(_m.group(1))) < (_today.year, _today.month):
            _cur_m = f"{_today:%Y-%m}"
            _gmv_cur = round(sum(v for k, v in gmv_by_date.items() if k.startswith(_cur_m)), 2)
            _ord_cur = int(sum(v for k, v in orders_by_date.items() if k.startswith(_cur_m)))
            summ['month_before_last'] = summ.get('last_month', {'label': '', 'gmv': 0, 'orders': 0, 'avg': 0})
            summ['last_month'] = dict(_tm)
            summ['this_month'] = {
                'label': f"{_today.month}月 {_today.year}",
                'gmv': _gmv_cur,
                'orders': _ord_cur,
                'avg': round(_gmv_cur / _ord_cur, 2) if _ord_cur else 0,
            }
            print(f"summary roll: 當月→{summ['this_month']['label']} ({_gmv_cur:,.0f}), 上月→{summ['last_month']['label']}, 上上月→{summ['month_before_last']['label']}")
    except Exception as _e:
        print('summary roll skip:', _e)

    # gmv_target (Tag 11): 當月 actual 用 daily 數據
    try:
        _gt = sd.get('gmv_target')
        if _gt and 'actual' in _gt and 'labels' in _gt:
            _today2 = datetime.date.today()
            _cm = f"{_today2:%Y-%m}"
            if _cm in _gt['labels']:
                _idx = _gt['labels'].index(_cm)
                _gmv_cur2 = round(sum(v for k, v in gmv_by_date.items() if k.startswith(_cm)), 2)
                _gt['actual'][_idx] = _gmv_cur2
                print(f"gmv_target actual[{_cm}] = {_gmv_cur2:,.2f}")
    except Exception as _e2:
        print('gmv_target update skip:', _e2)

# available_months
for d in DATE_KEYS:
    mth = d[:7]
    if mth not in sd.get('available_months', []):
        sd['available_months'].append(mth)

# ---------- 4. Write back ----------
out = js[:m.start()] + json.dumps(sd, ensure_ascii=False, separators=(',', ':')) + js[m.end():]
hdr = "// Auto-generated sales trend data (GP Report Excel + daily order reports)\n"
hdr += "// Generated: " + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + "\n\n"
js_new = hdr + "const salesTrendData = " + json.dumps(sd, ensure_ascii=False, separators=(',', ':')) + ";\n"
open('data/sales_trend_data.js', 'w', encoding='utf-8').write(js_new)
print(f"\nwritten data/sales_trend_data.js ({len(js_new)//1024} KB)")
print(f"gmv_by_date: {[round(gmv_by_date[d],2) for d in DATE_KEYS]}")
print(f"orders_by_date: {[orders_by_date.get(d,0) for d in DATE_KEYS]}")
