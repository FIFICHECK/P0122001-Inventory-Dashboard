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
gmv_date_hour = defaultdict(lambda: defaultdict(float))   # date -> hour -> gmv（chart 3 filter 用）
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

def _accum(date, hour, sku, name, qty, total):
    """Shared accumulator — called by both MMS and Exchange processors."""
    gmv_by_date[date] += float(total)
    qty_by_date[date] += int(qty or 0)
    gmv_by_hour[hour] += float(total)
    gmv_date_hour[date][hour] += float(total)
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
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active
        # MMS + Exchange 格式 layout 完全相同（row 5 headers, col 7 Order Date, col 27 Total）
        # — 實測 2026-09-28：Exchange 檔用 MMS parser 得出 $14,926.80，同檔案自身總數一致
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
# gmv_by_date_hour: {date: [24 個鐘嘅 GMV array]} — chart 3 filter by date/month 用
sd['gmv_by_date_hour'] = {
    'labels': DATE_KEYS,
    'hours': [f"{h:02d}" for h in range(24)],
    'data': {d: [round(gmv_date_hour[d].get(h, 0), 2) for h in range(24)] for d in DATE_KEYS},
}
print(f"gmv_by_date_hour: {len(DATE_KEYS)} dates x 24 hours")

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

    # 當月 KPI 每次由 daily 重新計（唔好停留喺第一次 roll 嘅舊值）
    try:
        _now4 = datetime.date.today()
        _lbl4 = f"{_now4.month}月 {_now4.year}"
        if summ.get('this_month', {}).get('label') == _lbl4:
            _cm4 = f"{_now4:%Y-%m}"
            _gmv4 = round(sum(v for k, v in gmv_by_date.items() if k.startswith(_cm4)), 2)
            _ord4 = int(sum(v for k, v in orders_by_date.items() if k.startswith(_cm4)))
            summ['this_month']['gmv'] = _gmv4
            summ['this_month']['orders'] = _ord4
            summ['this_month']['avg'] = round(_gmv4 / _ord4, 2) if _ord4 else 0
            print(f"this_month refresh: {_lbl4} GMV={_gmv4:,.2f} orders={_ord4}")
    except Exception as _e4:
        print('this_month refresh skip:', _e4)

    # gmv_by_month: 加入/更新當月 bucket（GP Excel 未有當月 → 用 daily 累計）
    # 令 monthly chart + month filter 揀當月時唔會空白
    try:
        _now3 = datetime.date.today()
        _cm3 = f"{_now3:%Y-%m}"
        _gmv_m3 = round(sum(v for k, v in gmv_by_date.items() if k.startswith(_cm3)), 2)
        _gbm = sd.get('gmv_by_month') or {'labels': [], 'data': []}
        _pairs = dict(zip(_gbm.get('labels', []), _gbm.get('data', [])))
        _pairs[_cm3] = _gmv_m3
        _sorted = sorted(_pairs.items())
        sd['gmv_by_month'] = {'labels': [p[0] for p in _sorted], 'data': [p[1] for p in _sorted]}
        print(f"gmv_by_month[{_cm3}] = {_gmv_m3:,.2f}")
    except Exception as _e3:
        print('gmv_by_month update skip:', _e3)

# ── Calendar-correct KPI cards + month backfill (2026-10-01 month-boundary fix) ──
# The summary roll above advances by ONE data-month (from_excel's latest GP month = 2026-08), so once
# today is >=2 months past the latest GP month, the previous calendar month (2026-09 — which has a
# COMPLETE month of daily reports but no GP row) is skipped: its GMV vanishes from the KPI cards, the
# monthly chart and Tag 11. Fix: recompute the three cards by REAL calendar month (當月/上月/上上月),
# preferring the GP monthly value when present else the daily sum; backfill any missing calendar month
# into gmv_by_month; fill gmv_target.actual where it is still 0.
try:
    _now5 = datetime.date.today()
    _p1d = _now5.replace(day=1) - datetime.timedelta(days=1)      # 上月
    _p2d = _p1d.replace(day=1) - datetime.timedelta(days=1)       # 上上月
    _gbm_lk = dict(zip(sd.get('gmv_by_month', {}).get('labels', []),
                       sd.get('gmv_by_month', {}).get('data', [])))
    _obm_lk = dict(zip(summ.get('orders_by_month', {}).get('labels', []),
                       summ.get('orders_by_month', {}).get('data', [])))

    def _day_gmv5(ym):
        return round(sum(v for k, v in gmv_by_date.items() if k.startswith(ym)), 2)

    def _day_ord5(ym):
        return int(sum(orders_by_date.get(k, 0) for k in orders_by_date if k.startswith(ym)))

    def _card5(dt):
        ym = f"{dt:%Y-%m}"
        g = _gbm_lk.get(ym)
        if g:
            o = int(_obm_lk.get(ym) or 0)
            return {'label': f"{dt.month}月 {dt.year}", 'gmv': round(float(g), 2), 'orders': o,
                    'avg': round(float(g) / o, 2) if o else 0}
        g = _day_gmv5(ym); o = _day_ord5(ym)
        return {'label': f"{dt.month}月 {dt.year}", 'gmv': g, 'orders': o,
                'avg': round(g / o, 2) if o else 0}

    _rolled_last = dict(summ.get('last_month', {}))   # post-roll == from_excel's this_month (GP, Aug)
    summ['this_month'] = _card5(_now5)
    summ['last_month'] = _card5(_p1d)
    if _rolled_last.get('label') == f"{_p2d.month}月 {_p2d.year}":
        summ['month_before_last'] = _rolled_last
    else:
        summ['month_before_last'] = _card5(_p2d)
    print("KPI cards (calendar): 當月->{0} {1:,.0f} | 上月->{2} {3:,.0f} | 上上月->{4} {5:,.0f}".format(
        summ['this_month']['label'], summ['this_month']['gmv'],
        summ['last_month']['label'], summ['last_month']['gmv'],
        summ['month_before_last']['label'], summ['month_before_last']['gmv']))

    # backfill calendar months (current + previous) into gmv_by_month from daily when GP has no row
    _pairs5 = dict(zip(sd['gmv_by_month'].get('labels', []), sd['gmv_by_month'].get('data', [])))
    for _dt5 in (_now5, _p1d):
        _ym5 = f"{_dt5:%Y-%m}"
        if not _pairs5.get(_ym5):
            _dv5 = _day_gmv5(_ym5)
            if _dv5 > 0:
                _pairs5[_ym5] = _dv5
    _sorted5 = sorted(_pairs5.items())
    sd['gmv_by_month'] = {'labels': [p[0] for p in _sorted5], 'data': [p[1] for p in _sorted5]}
    print("gmv_by_month months: {0}".format(sd['gmv_by_month']['labels']))

    # Tag 11: fill actual for a month (current + previous) that is still 0, from daily
    _gt5 = sd.get('gmv_target')
    if _gt5 and 'actual' in _gt5 and 'labels' in _gt5:
        for _dt5 in (_now5, _p1d):
            _ym5 = f"{_dt5:%Y-%m}"
            if _ym5 in _gt5['labels']:
                _ix5 = _gt5['labels'].index(_ym5)
                if not _gt5['actual'][_ix5]:
                    _gt5['actual'][_ix5] = _day_gmv5(_ym5)
        print("gmv_target.actual: {0}".format(_gt5['actual']))
except Exception as _e5:
    print('calendar KPI fix skip:', _e5)

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
