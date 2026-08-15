#!/bin/bash
# VO 拉单工具 · macOS 双击运行
# 首次运行：自动创建本地虚拟环境(.venv)并安装依赖；之后直接启动。
cd "$(dirname "$0")" || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3。请先安装 Python 3：https://www.python.org/downloads/"
  read -n 1 -s -r -p "按任意键关闭…"
  exit 1
fi

VENV=".venv"
VPY="$VENV/bin/python"
# 探活：目录在不代表能用(改过仓库名/升级过 Python 都会让它坏)，坏了就删掉重建
if [ -d "$VENV" ] && ! "$VPY" -c "import sys" >/dev/null 2>&1; then
  echo "运行环境已损坏(可能改过仓库名/升级过 Python)，正在重建…"
  rm -rf "$VENV"
fi
if [ ! -d "$VENV" ]; then
  echo "首次运行：正在创建运行环境并安装依赖（约 1–2 分钟，仅此一次）…"
  python3 -m venv "$VENV" || { echo "创建虚拟环境失败"; read -n 1 -s -r; exit 1; }
  "$VPY" -m pip install --upgrade pip >/dev/null 2>&1
  # 装依赖失败只警告不阻断：真缺包时 Python 会报清楚的 ImportError，好过「装不上就打不开」
  "$VPY" -m pip install -r requirements.txt || echo "⚠ 依赖安装失败(检查网络/代理)。仍尝试启动；若出现 ImportError 请把报错发给开发者。"
fi

"$VPY" vo_orders/gui.py
