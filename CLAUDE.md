# VO 拉单自动化

Python 脚本,自动化处理 VO 拉单相关的 Excel 表格。纯本地脚本,非产品级开发。

## 工作前必读
动手改之前,先读 `docs/journal/` 下最近几篇日志,了解此前的决策、踩过的坑和当前主线,避免重复探索或偏离方向。

## 工作后必写
每次会话结束、或完成一个有意义的改动后,在 `docs/journal/` 写一篇 `YYYY-MM-DD.md`,用摘要形式记录:改了什么、为什么、否决过哪些方案及原因、遗留问题/下一步、测试结果要点。保持摘要可读,不要贴完整对话。

## 多机协作
- 家里 iMac(开发,历史数据)/ 公司 MacBook Pro(开发 + 实时数据测试)/ Windows(仅测试目标,从不在此改逻辑)。
- 开工前 `git pull`,收工前 `git commit && git push`。
- 数据源路径因机器而异,写在 `config.py`(已 gitignore),不要硬编码进脚本、也不要提交。

## 不进 git
`test-data/`、`results/`、`config.py` 已 gitignore。测试数据和结果走 Syncthing 在两台 Mac 间同步,不走 git。
