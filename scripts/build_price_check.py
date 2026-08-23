#!/usr/bin/env python3
"""P0122001 Price Check data builder.

Compares HKTVmall store prices (PSP preferred, RSP fallback) against
SKECHERS official site (skechers.com.hk) prices, ONLINE styles only.

Matching strategy (strict, zero false positives):
  1. Numeric styles: HKTVmall "104624-GRY" <-> official "104624/GRY" (bare + color, case-insensitive)
  2. Letter styles:  HKTVmall "P224U043-0018" <-> official "P224U043/0018" (dash->slash normalization)
  3. Single-color bare fallback: official has exactly ONE color for a bare code and
     the HKTVmall style's bare matches (any HKTVmall color of that bare = same product)
No match -> official_price = null, diff = null (dashboard shows '—')

Output: data/price_check_data.json + data/psp_history.json (appends today's snapshot)
"""
import json, os, re, csv, datetime

REPO = os.path.expanduser('~/P0122001-Inventory-Dashboard')
os.chdir(REPO)

def valid_price(v):
    try:
        f = float(v)
        return f if f > 0 else None
    except:
        return None

def load_inventory():
    rows = []
    with open('data/inventory_all.csv', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for r in reader:
            sku = (r.get('Merchant SKU ID') or '').strip()
            if not sku:
                continue
            style = (r.get('Merchant Product ID') or '').strip()
            rows.append({
                'style': style,
                'sku': sku,
                'name_en': (r.get('SKU Name') or '').strip(),
                'name_chi': (r.get('SKU Name (Chi)') or '').strip(),
                'brand': (r.get('Brand Name (EN)') or r.get('Brand Name (CHI)') or '').strip(),
                'rsp': valid_price(r.get('Original Price')),
                'psp': valid_price(r.get('Discount Price')),
                'stock': (r.get('StockLevel') or '0').strip(),
                'status': (r.get('Online Status') or '').strip().lower(),
            })
    return rows

def load_official():
    with open('data/skechers_official_prices.json', encoding='utf-8') as f:
        return json.load(f)

def main():
    rows = load_inventory()
    official = load_official()
    print("inventory rows:", len(rows), "official styles:", len(official))

    # ---- build official lookup ----
    # numeric: (bare, color) -> official style key
    off_numeric = {}
    for k in official:
        m = re.match(r'^(\d+)[/-]([A-Za-z0-9]+)', k)
        if m:
            off_numeric[(m.group(1), m.group(2).upper())] = k
    # letter: normalized key (dash -> slash)
    off_letter = {}
    for k in official:
        m = re.match(r'^([A-Z]\d{3}[A-Z]\d{3})[/-]([A-Z0-9]+)$', k)
        if m:
            off_letter[m.group(1) + '/' + m.group(2)] = k
    # bare -> set of official colors (for single-color fallback)
    off_bare_colors = {}
    for k in official:
        m = re.match(r'^(\d+)[/-]([A-Za-z0-9]+)', k)
        if m:
            off_bare_colors.setdefault(m.group(1), set()).add(m.group(2).upper())
    single_color_bare = {b for b, cs in off_bare_colors.items() if len(cs) == 1}
    print("numeric official pairs:", len(off_numeric), "letter:", len(off_letter),
          "single-color bare:", len(single_color_bare))

    # ---- aggregate inventory at style level (online-only filter) ----
    style_map = {}  # style -> agg
    for r in rows:
        if r['status'] != 'online':
            continue
        e = style_map.setdefault(r['style'], {
            'style': r['style'],
            'name_en': r['name_en'], 'name_chi': r['name_chi'],
            'brand': r['brand'], 'rsp': None, 'psp': None,
            'stock': 0, 'link_sku': None,
        })
        if e['rsp'] is None and r['rsp']: e['rsp'] = r['rsp']
        if e['psp'] is None and r['psp']: e['psp'] = r['psp']
        try: e['stock'] += int(r['stock'] or 0)
        except: pass
        if not e['link_sku']:
            e['link_sku'] = r['sku']
    print("online styles:", len(style_map))

    # ---- match ----
    results = []
    matched = 0
    for style, e in style_map.items():
        rec = {
            'sku': 'P0122001_S_' + style,
            'link_sku': e['link_sku'],
            'style': style,
            'brand': e['brand'] or 'SKECHERS',
            'name_en': e['name_en'], 'name_chi': e['name_chi'],
            'rsp': e['rsp'], 'psp': e['psp'],
            'compare_price': e['psp'] if e['psp'] is not None else e['rsp'],
            'official_price': None, 'official_name': None, 'official_url': None,
            'diff': None, 'status': 'online', 'stock': e['stock'],
        }
        off_key = None
        m_num = re.match(r'^(\d+)-([A-Za-z0-9]+)$', style)
        m_let = re.match(r'^([A-Z]\d{3}[A-Z]\d{3})-([A-Z0-9]+)$', style)
        if m_num:
            k = (m_num.group(1), m_num.group(2).upper())
            if k in off_numeric:
                off_key = off_numeric[k]
            elif m_num.group(1) in single_color_bare:
                # official has single color for this bare — same product
                off_key = off_numeric.get((m_num.group(1), next(iter(off_bare_colors[m_num.group(1)]))))
        elif m_let:
            k = m_let.group(1) + '/' + m_let.group(2)
            if k in off_letter:
                off_key = off_letter[k]
        if off_key and off_key in official:
            o = official[off_key]
            op = valid_price(o.get('price'))
            if op:
                rec['official_price'] = op
                rec['official_name'] = (o.get('name') or '')[:60]
                rec['official_url'] = o.get('url') or ('https://www.skechers.com.hk/products/' + off_key.lower().replace('/', '-'))
                cp = rec['compare_price']
                if cp is not None:
                    rec['diff'] = round(cp - op, 2)
                matched += 1
        results.append(rec)

    results.sort(key=lambda r: r['style'])
    print("matched with official price:", matched, "/", len(results))

    # ---- psp_history append (for 走向 trend chart) ----
    hist_path = 'data/psp_history.json'
    hist = {}
    if os.path.exists(hist_path):
        try:
            hist = json.load(open(hist_path, encoding='utf-8'))
        except:
            hist = {}
    today = datetime.date.today().isoformat()
    for rec in results:
        trend_price = rec['psp'] if rec['psp'] is not None else rec['rsp']
        if trend_price is not None:
            entry = hist.setdefault(rec['sku'], {})
            entry[today] = trend_price
    json.dump(hist, open(hist_path, 'w', encoding='utf-8'), ensure_ascii=False)
    print("psp_history SKUs:", len(hist))

    # ---- write price_check_data.json ----
    json.dump(results, open('data/price_check_data.json', 'w', encoding='utf-8'), ensure_ascii=False)
    print("wrote data/price_check_data.json:", os.path.getsize('data/price_check_data.json'), "bytes")
    # summary
    up = sum(1 for r in results if r['diff'] is not None and r['diff'] > 0)
    down = sum(1 for r in results if r['diff'] is not None and r['diff'] < 0)
    flat = sum(1 for r in results if r['diff'] == 0)
    nodata = sum(1 for r in results if r['diff'] is None)
    print(f"diff: {up} up / {down} down / {flat} flat / {nodata} no-data")
    for r in results[:5]:
        print(" ", r['style'], "| rsp:", r['rsp'], "| psp:", r['psp'], "| off:", r['official_price'], "| diff:", r['diff'])

if __name__ == '__main__':
    main()
