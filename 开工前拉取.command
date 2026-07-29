#!/bin/bash
# 开工前拉取 · 把公开库和私有库一起 pull 起来
#
# 为什么要这个脚本：docs/、ISSUES.md、CLAUDE.md、SPEC.md 都是软链到
# 独立的私有仓库 erp-private，在项目目录里 git pull 只拉公开库，
# 私有库那份纹丝不动、也看不出落后 —— 换机后极易在 push 时才发现。
#
# 双击运行，或在终端 `./开工前拉取.command`（加 -q 跳过结束时的暂停）。
# 只读：只 pull，不 push、不 reset、不 checkout。

cd "$(dirname "$0")" || exit 1
PUBLIC="$(pwd)"

# 私有库路径从软链解析，不硬编码 —— 两台 Mac 目录布局可以不同
PRIVATE=""
if [ -L docs ]; then
  PRIVATE="$(dirname "$(readlink docs)")"
fi

fail=0

pull_repo() {
  local label="$1" dir="$2"
  echo ""
  echo "── $label"
  if [ ! -d "$dir/.git" ]; then
    echo "   ✗ 不是 git 仓库：$dir"
    fail=1
    return
  fi
  # --autostash: 本地有未提交改动也不被挡下（收工前的常见状态）
  if git -C "$dir" pull --rebase --autostash; then
    echo "   ✓ $(git -C "$dir" rev-parse --abbrev-ref HEAD) · $(git -C "$dir" log --oneline -1)"
    local dirty
    dirty="$(git -C "$dir" status --porcelain | wc -l | tr -d ' ')"
    [ "$dirty" != "0" ] && echo "   · 本地有 $dirty 处未提交改动（未动）"
  else
    echo "   ✗ 拉取失败 —— 冲突或网络问题，需要手工处理：$dir"
    fail=1
  fi
}

echo "开工前拉取"
pull_repo "公开库 $(basename "$PUBLIC")" "$PUBLIC"

if [ -z "$PRIVATE" ]; then
  echo ""
  echo "── 私有库"
  echo "   ✗ docs 不是软链，找不到 erp-private —— 私有库没拉，请手工确认"
  fail=1
else
  pull_repo "私有库 $(basename "$PRIVATE")" "$PRIVATE"
fi

echo ""
if [ "$fail" = "0" ]; then
  echo "✓ 两个仓库都已是最新，可以开工"
else
  echo "⚠ 有仓库没拉成功，见上面的 ✗ —— 别急着改代码"
fi

if [ "$1" != "-q" ]; then
  echo ""
  read -n 1 -s -r -p "按任意键关闭…"
  echo ""
fi

exit "$fail"
