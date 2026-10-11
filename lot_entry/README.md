# 入库批次录入（`lot_entry/`）

操作员不再进 Odoo 逐条录批次/效期，改在一张排好版的 Excel 里填，脚本转成 Odoo 导入表。
立项与决策过程见 `ISSUES.md` 的 `[入库批次] R`（前身 Q：拍照识别，已搁置）。

```
Odoo 入库单导出 ──①──> Lot_Entry_*.xlsx ──操作员填──> ──②──> Lot_Import_*.xlsx + Lot_Check_*.xlsx ──> 导入 Odoo
```

一个入口两用，**按文件内容自动判断**走 ① 还是 ②（录入表里有隐藏 sheet `_meta` 作标记）：

- Windows：把 xlsx **拖到 `Lot-Entry-Windows.bat` 上**，或双击后选文件。跑完窗口停住看摘要，并自动打开产出目录。
- Mac：双击 `Lot-Entry-Mac.command` 选文件；或 `python3 lot_entry/lot_entry.py <xlsx> [--no-open]`。
- 产出一律写 `output/YYYYMMDD/`，文件名带时间戳，不覆盖。

## 输入：Odoo 导出（① 的输入）

在 Inventory → Transfers 勾选要入库的单 → Export。**存成导出模板**，操作员以后一键导出。
技术名（勾「导入兼容导出」）与英文标签两种列名都认；**德语界面的标签不认**，缺列会报出缺哪个。

| 必需字段 | 英文标签 | 技术名 |
|---|---|---|
| 入库单 External ID | `External ID` | `id` |
| 行 ID | `Operations/ID` | `move_line_ids/id` |
| Stock Move ID | `Operations/Stock Move/ID` | `move_line_ids/move_id/id` |
| 产品 | `Operations/Product` | `move_line_ids/product_id` |
| 应收量 | `Operations/Stock Move/Demand` | `move_line_ids/move_id/product_uom_qty` |
| 单位 / 来源库位 / 目标库位 | `Operations/Unit of Measure` / `From` / `To` | `…/product_uom_id` / `location_id` / `location_dest_id` |

选填：`Reference`（WH/IN 单号，**强烈建议**——操作员靠它对纸面；缺时用 `Source Document` 作单头）、
`Contact`、`Source Document`。另有数字 `ID` 列（数据库 ID）会被忽略。

## 产出 1：`Lot_Entry_*.xlsx`（录入表）

用法设想（用户 2026-10-11）：**打印**出来拿去对货、手写批次/MHD → 回到电脑由同一人照抄进这张表 → 喂脚本。

- 按入库单分组（深色单头行：WH/IN 单号 · 采购单 · 供应商），**单内按品名字母序**（对货时按包装上的名字找）。
- 列：`Bezeichnung`（品名，加粗打头）| `Produkt`（编码）| `Soll` | `Charge / Lot` | `MHD` | `Menge` | `Σ / Status`。
- **第 1 个批次就填在产品同一行**，第 2 个填下面一行；更多批次在该产品下**插行**往下续写，
  脚本按「上方最近的产品行」归属。行高按手写留足。
- `Menge` 不填 = 实到即应到（只限单批次；多批次每批都要写数量）。
- 包装上没有批次 → **整行留空即可**（不再有 Keine Charge 列）；核对清单会列为 `NO LOT` 供复查。
- 只有批次 / MHD / 数量三列可写，其余锁定（工作表保护**无密码**，防手滑不防人；允许插行、调行高）。
- 批次与 MHD 列是**文本格式**：防 Excel 把 `11.2027` 吞成小数、`1E5` 变科学计数。
- `Σ / Status` 实时显示合计：`✓ = Soll`（单批不写数量）、`✓ 360`（相符，绿）、`≠ 8 / 9`（黄）。
  块长度用 `MATCH` 现找，插在哪都算得进；**只在屏幕上看，不进打印区域**（打印到 `Menge` 为止，纵向 A4 适宽）。
- 隐藏列 I–P 存行 ID / move ID / 单位 / 库位等，**别动**。

## 产出 2：`Lot_Import_*.xlsx`（导回 Odoo）

列（技术名，2026-10-11 实测导入成功的格式）：
`id`, `move_line_ids/id`, `move_line_ids/move_id/id`, `move_line_ids/product_id`, `move_line_ids/lot_name`,
`move_line_ids/qty_done`, `move_line_ids/product_uom_id`, `move_line_ids/location_id`,
`move_line_ids/location_dest_id`, `move_line_ids/expiration_date`。

- 第 1 个批次写进已有行；第 2 个起另起行（行 ID 留空、抄 move ID → 同一 move 下新增）。
- 新批次写 `lot_name`（`Lot/Serial Number` 只认已存在批次，会报 "No matching record"）。
- **所选单的全部已有行都带上**，没录的原值写回（`qty_done` 0、无批次）——不赌 Odoo 对未列出的 o2m 行的处理。
- 一张单一个能写的都没有 → 整单不进导入表。一个能写的都没有 → 不产导入表，只出核对清单。

## 产出 3：`Lot_Check_*.xlsx`（核对清单）

每产品的结果：`OK`（绿）/ `WARN`（黄，照写）/ `ERROR` + `SKIP`（红，**该产品整体不进导入表**）/
`NO LOT`（灰，一个批次都没写 → 按「包装无批次」跳过，已有行原样写回；请确认不是漏录）。列含品名，方便对表找。

| 规则 | 级别 |
|---|---|
| MHD `MM.YYYY` → **当月最后一天**；`DD.MM.YYYY` 原样；也认 `/` `-` 分隔与 `YYYY-MM(-DD)` | — |
| MHD 两位年（`05.28`）、认不出、不是真日期、已过期 | ERROR |
| 同产品已写了别的批次、这一行却只有 MHD/数量没批次；批次含空格等异常字符、不在 3–20 位、同产品重复 | ERROR |
| 多批次但有批次没写数量；数量非正数 | ERROR |
| 各批合计 ≠ 应收（部分到货等） | WARN |
| 录入 MHD 与产品名里的 MHD（如 `MHD 2027/10`、`MHD05.28`）年月不符 | WARN |
| 有批次没写 MHD | WARN |
| 没写批次却写了 MHD 或数量（Odoo 14 效期挂在批次上，无批次无处存；也可能是漏录批次） | WARN，跳过 |

单批次不写数量 → 取应收量。批次照录不改大小写。

## 测试

```bash
python3 -m pytest lot_entry/test_lot_entry.py -q
```

全程测试用代码造假导出（真实导出含供应商信息，`*.xlsx` 一律不进 git）。
