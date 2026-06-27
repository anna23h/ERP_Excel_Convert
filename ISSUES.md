# 使用问题清单 / ISSUES

> 使用中遇到的问题往这里追加，每条固定写：现象 / 意图 / 期望 / 线索。
> 修完把「状态」从 TODO 改成 DONE，不要删除（保留历史）。
> 公开仓库：勿写真实进货价、供应商、客户信息、凭据等敏感内容。

---

## [阶段一] 多件装面单只标了 x2
- 状态: DONE (2026-06-27, build_excel.py is_multipack 正则改为 x\d+$)
- 现象: 阶段一生成的面单里，Internal Reference 结尾为 `x2` 的才被标色，`x3`/`x4` 等没被标
- 意图: 标色是提醒打包员"这是多件装商品，需要装 N 件"
- 期望: 所有 `x?`（x 后跟任意数字）都标色，不止 x2
- 线索: `build_excel.py:126` 的 `is_x2()` 正则硬编码为 `x2$`，应放宽为 `x\d+$`（函数名也可一并改为更通用的命名）
- 记录: 2026-06-27

---

## [阶段一] 捡货表/面单 字体与单元格格式调整
- 状态: DONE (2026-06-27, build_excel.py style_sheet 加 left_cols/small_cols 参数 + fix_merged_alignment；ROW_H 40→35)
- 现象: 捡货表、面单的内容单元格当前全部统一居中（水平+垂直 center）、字号 15、行高 40
- 意图: 让打包/捡货员更易读对齐，弱化次要列
- 期望:
  1. 【捡货表】`Internal Reference`、`Picking Name`、`Barcode` 三列内容 → 左对齐 + 垂直下沉(bottom)；其中 `Picking Name` 字号比其他小 2 号
  2. 【面单】前四列 `Order Reference`、`VO Tracking No`、`Internal Reference`、`Picking Name` 内容 → 左对齐 + 垂直下沉(bottom)；但若该单元格是合并单元格 → 改为 左对齐 + 垂直居中(center)。其中 `Internal Reference`、`Picking Name` 字号比其他小 2 号
  3. 捡货表、面单、无货勾选 三张表行高 40 → 35
- 线索:
  - 常量: `build_excel.py:22-25` `CENTER`(全居中) / `FONT`(15) / `HEAD_FONT` / `ROW_H=40`
  - 套格式: `style_sheet()` (build_excel.py:131-140) 目前**全表统一**套 CENTER+FONT，无按列区分 → 需新增「按列名差异化对齐/字号」逻辑，并判断单元格是否在合并区域（openpyxl `ws.merged_cells`）走不同 vertical
  - 行高: 把 `ROW_H = 40` 改为 35（三张表共用同一函数即一处生效；需确认无货勾选表也走 style_sheet）
  - 待确认: "下沉" = 垂直 bottom 对齐（如理解有误请先纠正）
- 记录: 2026-06-27

---

## [阶段一] 无货勾选表改为面单版式 + 0/1 标注列
- 状态: DONE (2026-06-27, build_nogoods_helper 改为复用 build_facesheet + 前置 0/1 列；无货勾选套用面单同款样式/合并/标色)
- 现象: 无货勾选表版式与面单不同（列不同、无合并）。仓库返回的纸质文档是面单版式，操作员对着纸面单在屏幕上核对无货勾选表时两边对不上，逐行错位，输入易错、效率低
- 意图: 0/1 标缺货（0=有货,1=缺货）方式有效，保留；但表要和纸质面单长一样，操作员逐行填 0/1 不错位
- 期望: 无货勾选表 = 面单内容/版式完全一致（同列、同合并、同标色），仅在最前面加一列 0/1 标注列
- 线索:
  - `build_excel.py` `build_nogoods_helper()` 改为 `build_facesheet()` + `insert(0, "无货(1=缺货)", 0)`
  - `_write_pickface` 里无货勾选 sheet 套和面单一样的 style_sheet(left/small) + highlight_facesheet + merge_multiproduct + fix_merged_alignment
  - stage2 兼容性: 活跃读取 `classify_return` 按表头『无货』前缀找标注列 + 行内 SCP 键（Order Reference 含 SCP），与结构解耦 → 不受影响（已端到端验证）
  - 遗留: 休眠函数 build_shortage()/read_marked() 仍按旧列名 SKU/商品名/数量 取数；未来复活缺货记录阶段需改读面单版列名
- 记录: 2026-06-27
