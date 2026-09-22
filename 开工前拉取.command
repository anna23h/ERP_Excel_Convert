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
PUBLIC_TOP="$(git rev-parse --show-toplevel 2>/dev/null)"

# 私有库定位：**问 git，不问 readlink**（2026-09-22 改）。
#
# 原先是 `[ -L docs ] && dirname "$(readlink docs)"`，只在「项目根有 docs 软链」这一种
# 布局下成立。Mac 是 5 条根软链，成立；**Windows 上一条都没有**——那边建不了文件级软链
# （mklink 不带 /J 要管理员或开发者模式），走的是 `原始开发文档` → erp-private 的整目录
# 联接，名字和路径都对不上 `docs`。于是 Windows 上私有库那半从来没被拉过，
# 2026-09-22 实测本地已落后 11 个提交，直到 push 被拒才发现。
#
# 改成逐个候选问 `git rev-parse --show-toplevel`：**是个 git 仓库、且不是公开库自己**，
# 就是它。不关心那个路径是软链、目录联接还是真目录，一份脚本两边通用，不写 if-Windows。
resolve_private() {
  local c top
  for c in "$ERP_PRIVATE_DIR" docs 原始开发文档 ../erp-private; do
    [ -n "$c" ] && [ -e "$c" ] || continue
    top="$(git -C "$c" rev-parse --show-toplevel 2>/dev/null)" || continue
    if [ -n "$top" ] && [ "$top" != "$PUBLIC_TOP" ]; then
      echo "$top"
      return 0
    fi
  done
  return 1
}

PRIVATE="$(resolve_private)" || PRIVATE=""

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
  echo "   ✗ 找不到 erp-private —— 私有库没拉，请手工确认"
  echo "     试过这几处（都不是独立 git 仓库）：\$ERP_PRIVATE_DIR / docs / 原始开发文档 / ../erp-private"
  echo "     Mac: 项目根应有 docs 等 5 条软链；Windows: 应有 原始开发文档 目录联接"
  echo "     临时可用 ERP_PRIVATE_DIR=/path/to/erp-private ./开工前拉取.command"
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
