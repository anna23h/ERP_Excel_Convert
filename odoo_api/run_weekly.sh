#!/bin/bash
# 周三备货汇总定时入口（launchd 调它；2026-11-08 起跑 restock.py，取代已搁置的 stock_report.py 库存周报）。装进 launchd 的做法见 odoo_api/README.md。
# 单独写个 wrapper 而不是让 launchd 直接调 python：需要固定 cwd（产出路径 output/ 是相对的）、
# 需要用仓库自己的 .venv、还需要把每次运行追加进日志好事后追溯。
set -u
cd "$(dirname "$0")/.." || exit 1
REPO="$(pwd)"
VPY="$REPO/.venv/bin/python"
[ -x "$VPY" ] || VPY="$(command -v python3)"
LOG="$REPO/output/odoo_api.log"
mkdir -p "$REPO/output"
{
  echo "===== $(date '+%F %T') 开始 ====="
  "$VPY" "$REPO/odoo_api/restock.py" "$@"
  echo "===== $(date '+%F %T') 结束 退出码=$? ====="
} >>"$LOG" 2>&1
