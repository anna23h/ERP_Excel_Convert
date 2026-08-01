#!/bin/bash
# SessionStart hook：开会话时自动把公开库 + 私有库一起 pull 起来。
#
# 为什么需要：docs/、ISSUES.md、CLAUDE.md、SPEC.md 都是软链到 erp-private，
# 在项目目录里 git pull 只拉公开库，私有库落后了看不出来。CLAUDE.md 写了
# 「开工前跑 ./开工前拉取.command」，但靠人记会漏——2026-08-01 就漏了一次，
# 在落后的状态上干了一整天，收工 push 被两个远端同时拒绝，只能做一次
# 带重命名冲突的 rebase。这个 hook 把那条约定变成机器强制。
#
# 由 .claude/settings.json 的 SessionStart hook 调用。
# 只读：只 pull，不 push / 不 reset / 不 checkout。永远 exit 0，绝不阻断会话
# ——拉不动是「提醒人」的事，不该让人连会话都开不了。

cd "${CLAUDE_PROJECT_DIR:-$PWD}" 2>/dev/null

if [ ! -x ./开工前拉取.command ]; then
  # 会话可能在别的目录开的，此时静静提示一句就好
  echo '{"systemMessage":"⚠ 开工前拉取: 找不到 ./开工前拉取.command，本次未拉取"}'
  exit 0
fi

out=$(./开工前拉取.command -q 2>&1); rc=$?

if [ "$rc" = "0" ]; then
  msg="✓ 开工前拉取：公开库 + 私有库均已最新"
  note="两个仓库均已 pull 到最新，可以开工。"
else
  msg="⚠ 开工前拉取失败（退出码 ${rc}）—— 别急着改代码"
  note="拉取未全部成功。在动代码前先把这件事告诉用户并确认怎么处理，不要在落后的状态上工作。"
fi

# 用 jq 转义：拉取输出里有换行、引号、中文，手拼 JSON 必翻车
jq -Rs --arg msg "$msg" --arg note "$note" \
  '{systemMessage:$msg, suppressOutput:true,
    hookSpecificOutput:{hookEventName:"SessionStart",
      additionalContext:("【开工前拉取 · SessionStart】" + $note + "\n---\n" + .)}}' <<<"$out"
