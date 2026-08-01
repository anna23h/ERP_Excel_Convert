# 新机器上:cp config.example.py config.py,再按本机实际路径填写。
# config.py 已 gitignore,不会被提交。
INPUT_DIR = "/path/to/excel/inputs"
OUTPUT_DIR = "/path/to/excel/outputs"

# 供应商简称覆盖: short_vendor() 的产出 → 最终写进 ERP `FS` 字段的代号。
# 放这里而不是 common/vendor.py: 公开库不能出现真实供应商名,
# 而用代号本来就是为了不点名。未列出的供应商保持 short_vendor() 原产出不变。
# 键 = short_vendor() 算出来的简称(不是采购单里的全名),值 = 你要的代号。
VENDOR_ALIAS = {
    # "某大批发商": "P",
    # "某连锁药房": "B",
}
