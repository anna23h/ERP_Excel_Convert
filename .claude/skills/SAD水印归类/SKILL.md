---
name: SAD水印归类
description: 每月一次的 JD 直邮小包裹荷兰出口申报单（SAD PDF）批次处理。用户说「处理 SAD 水印」「JD 问题包裹归类」「挑出问题包裹 / 没水印的单」「做 POE 声明 / Export statement」「荷兰出口申报单」，或给出含 declaration.xlsx + SAD-JD…-….pdf 的文件夹、Export statement templete.xlsx 时触发。流程：跑 sad_export/sad_batch.py（先 verify 回归，再 run）→ 源文件夹只读，复制出 待处理/（declaration 里 TrackingId 只出现 1 次）、正常/（每单一份 Exit Confirmed 水印件）和填好的 Export statement；退出码 2 就停下把原因交给用户。规则已与用户确认，不必再复述需求或逐条征求意见。
---

# SAD 水印归类（JD 问题包裹 → POE 声明）

规则已于 2026-10-02 / 10-16 与用户定好，**写在脚本里**（`sad_export/sad_batch.py` 的 docstring 是权威说明），
不必再复述需求或逐条征求意见。来龙去脉见私有库 `ISSUES.md` 的 `[SAD] A`。

## 输入

- **源文件夹**：用户下载解压的一批，**只读**——不移动、不删除、不改 `declaration.xlsx`。
  里面是 `declaration.xlsx`（`TrackingId | DeclarationUcr | FileName`）+ 每行一份 `SAD-<TrackingId>-<时间戳>.pdf`
  （正常单 3 份，补发时 4~6 份）。可能是 `SAD_<批次号>\`，也可能是扁平的 `9月荷兰出口申报单\` 这种。
- **工作根**：通常 `Desktop\InProgress\JD问题包裹自动化处理\`，放着 `Export statement templete.xlsx` 和历次成品。
- **输出目录**：`<工作根>\SAD_<批次号>\`。源文件夹没有批次号时**问用户叫什么**（2026-10 那批用户定的是 `SAD_9月荷兰出口申报单`）。

## 第一步：verify（PDF 版式回归）

在仓库根目录跑。拿**最近一次已完成的批次**的 `待处理\` 和成品表做参考，必须输出「一致」：

```bash
py -3 -X utf8 -m sad_export.sad_batch verify "<工作根>/<上一批>/待处理" "<工作根>/<上一批>/Export statement - <上一批>.xlsx"
```

Mac 上把 `py -3 -X utf8` 换成 `python3`。不一致 = PDF 版式变了，**停下告诉用户**，按项目规则立项后再改解析。

## 第二步：run

```bash
py -3 -X utf8 -W ignore -m sad_export.sad_batch run "<源文件夹>" "<工作根>/SAD_<批次号>"
```

- 模板默认取输出目录的上级里的 `Export statement templete.xlsx`，别处的用 `--template`。
- 1.6 万份 PDF 约 5 分钟（每份解内容流查水印）。命令可能超过 2 分钟，放后台跑。
- 依赖：`openpyxl` + 系统 `pdftotext`（Windows 本机 Git 自带 mingw 版；Mac `brew install poppler`）。
  Windows 上 PATH 里的 `python` 是 WindowsApps 空壳，用 `py -3`；`-X utf8` 防中文打印崩。
- 输出目录已存在 → 拒绝（不覆盖）。要重做就让用户先挪走旧目录。

## 退出码 2 = 停下交给用户

脚本列出原因、**不写任何文件**。这些正是任务定义里「停下来报告、不要自行处理」的情况：

| 情况 | 交给用户的问题 |
|---|---|
| 没有 `declaration.xlsx` / 列名不同 / 与磁盘 PDF 对不上 | 源文件夹是不是不完整、导出格式变了？ |
| 「只出现 1 次」与「无水印」不一致 | 这批规则是否不同？把两边名单原样贴出 |
| 问题包裹 > 3%（往批 1.0%、0.38%） | 同上 |
| 问题包裹的 SAD `Artikelen` ≠ 1 | 多品名一行填不下，怎么填？ |
| 字段为空 / 格式异常 / layout 与内容流两种解析不一致 | 附上那份 PDF 的文件名，请用户看原件 |

**不要**手工补字段后另写脚本出表，也不要放宽阈值重跑——等用户决定。

## 交付时报告

一批一行：批次名、TrackingId 总数、问题包裹数、正常件数、成品路径。再加：

- 守恒：`待处理 + 正常 = TrackingId 总数`（脚本打印）；源文件夹份数未变
- 问题包裹清单（File Name / Parcel No. / Content / Value / Weight，脚本 stdout 里有）
- 份数分布里出现新形态（如 4、5 份）或「水印件不是同单最大」提示时，提一句
- 重复的水印件（同单多份水印）只收时间戳最新一份，报一下另有几份没收

## 规则要点（脚本已实现，改脚本时别破坏）

- **识别水印靠 PDF 内容**，看内容流里有没有 `(Exit Confirmed)Tj`。不靠文件名、不靠大小：
  2026-10 批水印件恒为按时间戳第 3 份、且是同单最大文件，但这只是规律，「最大」只作交叉核对。
- **不要用 `stream…endstream` 正则截取内容流**：压缩数据里可能碰巧有这串字节，2026-10-02 因此误判 3 单。
  从 `>> stream` 之后交给 `zlib.decompressobj()`，由 zlib 判断流结束。
- 品名在第 32 栏商品行上，但有时因纵坐标差一点被 pdftotext 排到上一行；收件人名隔一空行，不会混进来。
- Excel：H1 写 `File Name`；D、E 列文本、德式逗号；去掉模板的「重复值」条件格式（唛头本来就会重复），
  C 列 wine/champagne 文字规则保留；Sheet2 不动。动表前看有没有 `~$` 锁文件（Excel 正开着）。
- 多品名 SAD（`Artikelen` ≠ 1）字段位置不同，脚本不解析、直接停。2026-10-02 首测见过 13 单，都带水印。
