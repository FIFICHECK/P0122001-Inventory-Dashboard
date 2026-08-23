# P0122001 Inventory Dashboard (SKECHERS)

SKECHERS 香港官方網上商店 — inventory + price check dashboard for HKTVmall store **P0122001**.

Live: https://fificheck.github.io/P0122001-Inventory-Dashboard/

## Features

- **📋 All** — 8,102 style-level SKUs (aggregated from 40,436 size variants), searchable + brand/status filters
- **🔢 By Online SKU** — 1,136 online styles
- **🏷️ Brand** — brand summary
- **📦 By Category** — category type + full code breakdown
- **📊 By SKU Status** — online / invisible / force-OOS
- **⚠️ Alerts** — zero stock (6,829) + low stock (539)
- **🆕 New SKU** — created within 14 days
- **💰 Price Check** — compares HKTVmall price (PSP preferred, RSP fallback) vs **SKECHERS 官網** (skechers.com.hk):
  - 差價 / 走向 / 變動 three separate columns (H8391001 format)
  - 價格動向 button → daily PSP price line chart popup
  - 1,136 online styles, 753 matched with official price
- **📋 Report** — inventory report CSV downloads
- **📈 Sales Trend** — (ready for daily order report data)

## Data pipeline

| Script | Input | Output |
|---|---|---|
| `scripts/scrape_skechers.py` | skechers.com.hk Shopify /products.json | `data/skechers_official_prices.json` |
| `scripts/build_inventory_data.py` | `data/inventory_all.csv` (40,436 rows) | `data/inventory_data.json` (8,102 styles) |
| `scripts/build_price_check.py` | inventory CSV + official prices | `data/price_check_data.json` + `data/psp_history.json` |

The dashboard is **JS-driven** (fetches the JSON files) — unlike the source template's hardcoded rows,
because SKECHERS has 40K size-variant SKUs.

## Update flow (manual)

```bash
python3 scripts/scrape_skechers.py          # refresh official prices (optional)
python3 scripts/build_inventory_data.py     # from latest Exchange stocklevel CSV
python3 scripts/build_price_check.py        # rebuild price check + psp history
# then commit + push; GitHub Actions deploys Pages
```

Store data date: see `data/inventory_data.json` → `report_date`.
