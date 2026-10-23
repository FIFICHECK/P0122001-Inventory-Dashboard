#!/usr/bin/env python3
"""Scrape skechers.com.hk (Shopify) → data/skechers_official_prices.json

The official site exposes the standard Shopify /products.json API. Each product has
variants with SKUs like "104624/GRY-5" (style/color-size) and prices.

Output: data/skechers_official_prices.json
{
  "104624/GRY": {"name": "...", "price": 699.0, "url": "https://www.skechers.com.hk/products/vapor-foam-move-104624-gry", "colors": [...]},
  ...
}
Keyed by official style code (e.g. 104624/GRY). Also stores variant-level SKUs.
"""
import json, os, re, time, urllib.request

REPO = os.path.expanduser('~/P0122001-Inventory-Dashboard')
os.chdir(REPO)
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36'}

CACHE = '/tmp/p0122001/skechers/all_products.json'
MIN_PRODUCTS = 2000   # full scrape = ~3,949 products; anything smaller is partial
MIN_STYLES = 1500     # full run = ~7,000 styles
_COMPLETE = {'ok': True}   # flipped False if ANY page fails mid-pagination

def fetch_all():
    all_products = []
    if os.path.exists(CACHE):
        try:
            cached = json.load(open(CACHE))
        except Exception:
            cached = []
        # a poisoned/partial cache (0 or 250 products) must be ignored, not trusted
        if len(cached) >= MIN_PRODUCTS:
            print(f"using cache: {len(cached)} products")
            return cached
        print(f"ignoring poisoned cache ({len(cached)} products) — refetching")
    page = 1
    while True:
        req = urllib.request.Request(f'https://www.skechers.com.hk/products.json?limit=250&page={page}', headers=UA)
        try:
            d = json.load(urllib.request.urlopen(req, timeout=60))
        except Exception as e:
            print('ERR page', page, e)
            _COMPLETE['ok'] = False   # partial data — never publish it
            break
        prods = d['products']
        all_products.extend(prods)
        print(f'page {page}: +{len(prods)} (total {len(all_products)})')
        if len(prods) < 250:
            break
        page += 1
        time.sleep(0.4)
    # GUARD (2026-10-23): never cache an empty/failed scrape — a poisoned cache of
    # `[]` makes every later run "succeed" with 0 products and 0 styles.
    if not all_products:
        print('ERR scrape returned 0 products (rate-limit/outage) — cache NOT written')
        return all_products
    os.makedirs(os.path.dirname(CACHE), exist_ok=True)
    json.dump(all_products, open(CACHE, 'w'), ensure_ascii=False)
    return all_products

def build():
    all_products = fetch_all()
    styles = {}   # style_code -> {name, price, url, variants: {sku: price}, colors: set}
    for p in all_products:
        handle = p['handle']
        url = 'https://www.skechers.com.hk/products/' + handle
        title = (p.get('title') or '').strip()
        for v in p['variants']:
            s = v.get('sku') or ''
            price = v.get('price')
            try:
                price = float(price) if price else None
            except:
                price = None
            if not s or price is None:
                continue
            # official style code: strip size suffix
            m = re.match(r'^(.+?)-(?:XXS|XS|S|M|L|XL|XXL|XXXL|XLT|[0-9]+(?:\.[0-9]+)?|C\d+|Y\d+|W\d+|M\d+|5H|6H|7H|8H|9H)$', s)
            style = m.group(1) if m else s
            e = styles.setdefault(style, {'name': title, 'price': None, 'url': url, 'variants': {}, 'colors': set()})
            e['variants'][s] = price
            e['colors'].add(v.get('title') or '')
            if e['price'] is None or price < e['price']:
                e['price'] = price
    # normalize colors to list
    for e in styles.values():
        e['colors'] = sorted(c for c in e['colors'] if c)
    # GUARD (2026-10-23): a 0-style run must NEVER overwrite the good price file.
    # Observed: Shopify /products.json returned HTTP 429, the script wrote `{}` over
    # 6,942 styles, and the next build_price_check would have shown 0 matches / all
    # "暫無數據" on the live dashboard. Fail loudly and keep the previous snapshot.
    if not styles:
        print('ERR 0 styles parsed — data/skechers_official_prices.json left UNTOUCHED')
        print('    (previous snapshot kept; re-run later or wait out the Shopify 429)')
        raise SystemExit(1)
    # GUARD 2 (2026-10-23): a PARTIAL scrape (page 1 ok, page 2 → 429) produced 533
    # styles and happily overwrote the full 6,942-style snapshot. Require a complete
    # pagination run AND a sane style count before replacing the file.
    if not _COMPLETE['ok'] or len(styles) < MIN_STYLES:
        print(f"ERR incomplete scrape ({len(styles)} styles, complete={_COMPLETE['ok']}, "
              f"min={MIN_STYLES}) — data/skechers_official_prices.json left UNTOUCHED")
        print('    previous snapshot kept; re-run later (Shopify rate limit resets)')
        raise SystemExit(1)
    json.dump(styles, open('data/skechers_official_prices.json', 'w'), ensure_ascii=False, indent=1)
    print(f"styles: {len(styles)}, total variants: {sum(len(e['variants']) for e in styles.values())}")
    # stats: how many match bare numeric patterns
    numeric = sum(1 for k in styles if re.match(r'^\d+[/-]', k))
    print("numeric-prefixed styles:", numeric)
    # sample
    for k in list(styles)[:5]:
        print(" ", k, "->", styles[k]['price'], styles[k]['name'][:50])

if __name__ == '__main__':
    build()
