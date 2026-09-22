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

# JSON 转义用 **Python 而非 jq**（2026-09-22 改）：拉取输出里有换行、引号、中文，
# 手拼必翻车，所以一定要有个转义器。原来用 jq —— 但 **Windows 的 Git Bash 不自带 jq**，
# 实测这台机器上 hook 每次都是 `jq: command not found`、退出码 127，什么都没注入，
# 于是"私有库落后了"从来没被提醒过。Python 是本仓最小集依赖，三台机器上必然有。
PY=""
for c in python3 python py; do
  # WindowsApps 下的 python3 可能只是 Microsoft Store 转发壳，真跑一下才知道能不能用
  command -v "$c" >/dev/null 2>&1 && "$c" -c "import sys" >/dev/null 2>&1 && { PY="$c"; break; }
done

if [ -z "$PY" ]; then
  # 连 Python 都没有：退回一条纯静态消息（不含拉取输出，因而无需转义）
  printf '%s\n' '{"systemMessage":"⚠ 开工前拉取: 没找到可用的 Python，本次结果无法注入；请手工跑 ./开工前拉取.command"}'
  exit 0
fi

# msg/note 走环境变量传进去，免得在 shell 与 Python 两层引号里来回转义。
# ensure_ascii 保持默认 True：产出纯 ASCII 的 \uXXXX 转义，绕开 Windows 控制台 cp1252
# 编码问题。读 stdin 也必须走 buffer 再显式 utf-8 解码 —— 同一天已经在
# test_sales_insight.py 上栽过一次（text=True 走 cp1252 解中文直接炸）。
MSG="$msg" NOTE="$note" "$PY" -c '
import json, os, sys
out = sys.stdin.buffer.read().decode("utf-8", "replace")
print(json.dumps({
    "systemMessage": os.environ["MSG"],
    "suppressOutput": True,
    "hookSpecificOutput": {
        "hookEventName": "SessionStart",
        "additionalContext": "【开工前拉取 · SessionStart】" + os.environ["NOTE"] + "\n---\n" + out,
    },
}))
' <<<"$out"
