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

# --- Odoo XML-RPC 只读拉数（odoo_api/）---------------------------------
# 也可以改用环境变量 ODOO_URL / ODOO_DB / ODOO_USER / ODOO_API_KEY，环境变量优先。
# api_key 用 Odoo 的 **API Key**（用户头像 → 我的资料 → 账号安全 → 开发者 API 密钥），
# 不要用登录密码：密钥能单独吊销、不受双因素影响、泄露了不等于账号被接管。
# config.py 已 gitignore；**凭据不要贴进对话、不要写进任何 .py 以外的地方**。
ODOO = {
    "url": "http://erp.example.local:8888",
    "db": "your_db_name",
    "user": "script_bot@example.com",
    "api_key": "",
}

# 算在途/预测库存时要排除的供应商（名字子串，ilike 匹配）。填本机实际值，不进 git。
# 用途：并非所有确认状态的采购单都会真的收货——某些系统集成会用 PO 建技术性映射，
# 这类单永远不到货，却会生成未完成入库 move，把 virtual_available 彻底毁掉。
# 2026-08-23 实测：全局 7460 条未完成入库里 7350 条（98.5%）来自单一供应商的这类单，
# 真实在途只有 105 条 / 88 个产品。不排除的话「在途/预测」两列没有任何参考价值。
# 先跑 odoo_api/discover.py 或按供应商聚合一次 stock.move，确认本机该填谁。
ODOO_EXCLUDE_VENDORS = []
