#!/bin/bash
# sync_private.sh — copy P0122001 data/reports to the PRIVATE repo + push
# B pilot (2026-08-29): 由 crons 喺 build 完成後 call，確保 private repo 有最新數據
set -euo pipefail
SRC="${1:-/home/snkwok/P0122001-Inventory-Dashboard}"
DST="${2:-/home/snkwok/dashboard-private-data/P0122001}"

[ -d "$SRC/data" ] || { echo "no data dir: $SRC"; exit 1; }
mkdir -p "$DST/data" "$DST/reports"

# copy data files (json/js/csv only — 唔會抄 .py/allowlist 之外嘅嘢)
for f in "$SRC"/data/*; do
  b="$(basename "$f")"
  case "$b" in
    access_allowlist.json|*.py) continue ;;   # allowlist 留 public；py 唔係 data
  esac
  cp -f "$f" "$DST/data/" 2>/dev/null || true
done
# reports (xlsx/csv)
for f in "$SRC"/reports/*; do
  [ -e "$f" ] || continue
  cp -rf "$f" "$DST/reports/" 2>/dev/null || true
done
for sub in order_reports; do
  if [ -d "$SRC/reports/$sub" ]; then
    mkdir -p "$DST/reports/$sub"
    cp -f "$SRC/reports/$sub"/* "$DST/reports/$sub/" 2>/dev/null || true
  fi
done

cd "$DST"
git add -A
if ! git diff --cached --quiet; then
  git -c user.email="hermes@fificheck.local" -c user.name="Hermes" commit -q -m "P0122001 data sync $(date +'%Y-%m-%d %H:%M')"
  git push origin main
  echo "✅ private repo synced: $(date +'%H:%M:%S')"
else
  echo "ℹ️ 冇嘢改 — skip push"
fi
