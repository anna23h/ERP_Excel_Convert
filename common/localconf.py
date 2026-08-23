"""仓库根 `config.py` 的加载器（因机器而异的本地配置，已 gitignore）。

按**文件位置**直接加载，不靠 `sys.path` —— 各入口对 sys.path 的处理不一致，
靠 `import config` 会时灵时不灵（原注释见 vendor.py，2026-08-23 抽出共用）。

没有 config.py 时一律回落到默认值，不报错：同事那台机器只跑 Excel 流水线，
没有 config.py 也必须能跑。
"""
import importlib.util
import os

_CACHE = {}


def _config_path():
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "config.py")


def load_module():
    """加载仓库根 config.py 并返回模块对象；不存在返回 None。结果缓存，重复调用不重复执行。"""
    path = _config_path()
    if path in _CACHE:
        return _CACHE[path]
    mod = None
    if os.path.exists(path):
        spec = importlib.util.spec_from_file_location("_erp_config", path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
    _CACHE[path] = mod
    return mod


def get(name, default=None):
    """读 config.py 里的一个顶层变量；没有 config.py 或没定义该变量则返回 default。"""
    mod = load_module()
    if mod is None:
        return default
    val = getattr(mod, name, None)
    return default if val is None else val
