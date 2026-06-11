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
if [ ! -d "$VENV" ]; then
  echo "首次运行：正在创建运行环境并安装依赖（约 1–2 分钟，仅此一次）…"
  python3 -m venv "$VENV" || { echo "创建虚拟环境失败"; read -n 1 -s -r; exit 1; }
  "$VENV/bin/pip" install --upgrade pip >/dev/null 2>&1
  "$VENV/bin/pip" install -r requirements.txt || { echo "安装依赖失败"; read -n 1 -s -r; exit 1; }
fi

"$VENV/bin/python" gui.py
