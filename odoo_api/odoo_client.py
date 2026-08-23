"""Odoo XML-RPC 连接与查询封装（只读）。

凭据来源，按优先级：
  1. 环境变量 ODOO_URL / ODOO_DB / ODOO_USER / ODOO_API_KEY
  2. 仓库根 config.py 里的 ODOO = {"url":…, "db":…, "user":…, "api_key":…}
两者都没有就报错退出。**凭据不进 git**（config.py 已 gitignore），也别写进本文件。

为什么封一层而不是直接用 xmlrpc.client：
  - 超时。默认 ServerProxy 没有超时，Odoo 一卡 launchd 任务就永远挂着。
  - 分页。search_read 一次拉一万多行会被截断或超时，必须 limit/offset 循环。
  - 字段名跨版本有出入。用 fields_get 认字段 + 候选别名兜底，别硬编码
    （沿用 packing_list 德文列名那次的「别名元组」做法）。
"""
import http.client
import os
import sys
import time
import xmlrpc.client


class OdooError(RuntimeError):
    pass


# --------------------------------------------------------------------------
# 凭据
# --------------------------------------------------------------------------
def load_credentials():
    """→ dict(url, db, user, api_key)。缺任何一项都抛 OdooError（带上怎么补的说明）。"""
    env = {
        "url": os.environ.get("ODOO_URL"),
        "db": os.environ.get("ODOO_DB"),
        "user": os.environ.get("ODOO_USER"),
        "api_key": os.environ.get("ODOO_API_KEY"),
    }
    if all(env.values()):
        return env

    from common import localconf
    cfg = localconf.get("ODOO", {}) or {}
    merged = {k: (env[k] or cfg.get(k)) for k in env}
    missing = [k for k, v in merged.items() if not v]
    if missing:
        raise OdooError(
            "缺少 Odoo 凭据：" + " / ".join(missing) + "\n"
            "补法二选一：\n"
            "  1) 环境变量 ODOO_URL / ODOO_DB / ODOO_USER / ODOO_API_KEY\n"
            "  2) 仓库根 config.py 加一段（config.py 已 gitignore，不会被提交）：\n"
            '     ODOO = {"url": "http://erp.example:8888", "db": "…",\n'
            '             "user": "…", "api_key": "…"}\n'
            "api_key 用 Odoo 的 API Key（用户设置 → 账号安全 → 开发者 API 密钥），\n"
            "不要用登录密码——密钥能单独吊销，密码不能。")
    return merged


# --------------------------------------------------------------------------
# 带超时的 Transport
# --------------------------------------------------------------------------
class _TimeoutTransport(xmlrpc.client.Transport):
    def __init__(self, timeout, use_https=False):
        super().__init__()
        self._timeout = timeout
        self._use_https = use_https

    def make_connection(self, host):
        if self._connection and host == self._connection[0]:
            return self._connection[1]
        chost, self._extra_headers, x509 = self.get_host_info(host)
        cls = http.client.HTTPSConnection if self._use_https else http.client.HTTPConnection
        kw = {"timeout": self._timeout}
        if self._use_https:
            kw.update(x509 or {})
        self._connection = host, cls(chost, **kw)
        return self._connection[1]


class _TimeoutSafeTransport(_TimeoutTransport):
    def __init__(self, timeout):
        super().__init__(timeout, use_https=True)


def _proxy(url, path, timeout):
    use_https = url.lower().startswith("https://")
    tr = _TimeoutSafeTransport(timeout) if use_https else _TimeoutTransport(timeout)
    return xmlrpc.client.ServerProxy(url.rstrip("/") + path, transport=tr, allow_none=True)


# --------------------------------------------------------------------------
# 客户端
# --------------------------------------------------------------------------
class Odoo:
    """一个已登录的只读会话。

    context: 传给每次 execute_kw 的上下文。多公司环境下用 allowed_company_ids
    锁定公司，否则读到的库存/销量是「当前用户默认公司」——换台机器跑可能不一样。
    """

    #: 只读白名单。本层不写 ERP，误调用直接在客户端拦下，不指望服务端权限兜底。
    ALLOWED_METHODS = {"search", "search_read", "read", "read_group", "search_count",
                       "fields_get", "default_get", "name_get", "name_search"}

    def __init__(self, url, db, uid, api_key, timeout=180, context=None, verbose=True):
        self.url, self.db, self.uid, self.api_key = url, db, uid, api_key
        self.timeout, self.verbose = timeout, verbose
        self.context = dict(context or {})
        self.models = _proxy(url, "/xmlrpc/2/object", timeout)
        self._fields_cache = {}
        self.call_count = 0

    # -- 连接 ---------------------------------------------------------------
    @classmethod
    def connect(cls, creds=None, timeout=180, context=None, verbose=True):
        c = creds or load_credentials()
        if c["url"].lower().startswith("http://") and verbose:
            print("⚠ 走的是明文 HTTP：API 密钥、会话 cookie 在链路上全程明文。"
                  "确认这段链路真在内网再跑；公网必须先上 HTTPS。", file=sys.stderr)
        common = _proxy(c["url"], "/xmlrpc/2/common", timeout)
        try:
            version = common.version()
        except Exception as e:
            raise OdooError(f"连不上 {c['url']}：{e}\n"
                            "先确认地址能通（curl -sS <url>/web/login -o /dev/null -w '%{http_code}\\n'）"
                            "，再确认该端口在本机可达。") from e
        uid = common.authenticate(c["db"], c["user"], c["api_key"], {})
        if not uid:
            raise OdooError(
                f"认证失败：db={c['db']} user={c['user']}。\n"
                "三种常见原因：① db 名写错（自托管多库时尤其容易）；"
                "② 用了登录密码而该账号开了双因素——必须用 API Key；"
                "③ 该账号被停用。")
        obj = cls(c["url"], c["db"], uid, c["api_key"], timeout, context, verbose)
        obj.server_version = version.get("server_version", "?")
        if verbose:
            print(f"已连接 Odoo {obj.server_version}  db={c['db']}  uid={uid}")
        return obj

    # -- 底层调用 -----------------------------------------------------------
    def execute(self, model, method, args=None, kwargs=None, retries=2):
        if method not in self.ALLOWED_METHODS:
            raise OdooError(f"本层只读，不允许调用 {model}.{method}。"
                            "要写 ERP 请走人工导入表（见 sales_insight 的回写表）。")
        kwargs = dict(kwargs or {})
        ctx = dict(self.context)
        ctx.update(kwargs.pop("context", {}) or {})
        if ctx:
            kwargs["context"] = ctx
        last = None
        for attempt in range(retries + 1):
            try:
                self.call_count += 1
                return self.models.execute_kw(self.db, self.uid, self.api_key,
                                              model, method, args or [], kwargs)
            except xmlrpc.client.Fault as e:
                # 业务/权限错误重试没意义，直接抛
                last_line = e.faultString.strip().splitlines()[-1]
                if "cannot marshal None" in e.faultString:
                    # Odoo 14 服务端的坑：read_group 聚合出 None（最典型是 groupby=[] 且
                    # 命中 0 行，sum 为 NULL）时，服务端用 allow_none=False 序列化响应，
                    # 直接在服务端抛 TypeError。客户端改不了，只能把报错翻译成人话。
                    raise OdooError(
                        f"{model}.{method} 失败：服务端聚合结果里有 None，无法序列化"
                        "（Odoo 14 的已知行为）。\n"
                        "最常见的原因是**这个条件下一行都没匹配上**——检查日期窗口、"
                        "state 过滤、渠道过滤是不是把数据全筛掉了。\n"
                        "按 product_id 之类分组时不会触发（空结果返回空列表）；"
                        "只在不分组的全局聚合上出现。") from e
                raise OdooError(f"{model}.{method} 失败：{last_line}") from e
            except (OSError, http.client.HTTPException, xmlrpc.client.ProtocolError) as e:
                last = e
                if attempt < retries:
                    wait = 2 ** attempt * 3
                    print(f"  网络错误（{type(e).__name__}），{wait}s 后重试 "
                          f"{attempt + 1}/{retries}…", file=sys.stderr)
                    time.sleep(wait)
        raise OdooError(f"{model}.{method} 连续失败：{last}")

    # -- 分页查询 -----------------------------------------------------------
    def search_read_all(self, model, domain, fields, batch=2000, order="id", label=None):
        """limit/offset 循环把结果全部拉回来。

        用 order='id' 而不是默认排序：翻页期间若有人改数据，非稳定排序会漏行或重行。
        """
        out, offset = [], 0
        while True:
            chunk = self.execute(model, "search_read",
                                 [domain, fields],
                                 {"limit": batch, "offset": offset, "order": order})
            out.extend(chunk)
            if self.verbose and label:
                print(f"  {label}: {len(out)} 行…", end="\r", flush=True)
            if len(chunk) < batch:
                break
            offset += batch
        if self.verbose and label:
            print(f"  {label}: {len(out)} 行  ")
        return out

    def read_group_all(self, model, domain, fields, groupby, lazy=False, **kw):
        """read_group 不分页（服务端一次返回全部分组）。lazy=False 才能按多字段分组。"""
        return self.execute(model, "read_group", [domain, fields, groupby],
                            dict(lazy=lazy, **kw))

    def count(self, model, domain):
        return self.execute(model, "search_count", [domain])

    # -- 字段发现 -----------------------------------------------------------
    def fields_of(self, model):
        if model not in self._fields_cache:
            self._fields_cache[model] = self.execute(
                model, "fields_get", [], {"attributes": ["string", "type", "store", "relation"]})
        return self._fields_cache[model]

    def pick_field(self, model, candidates, required=True, what=""):
        """从候选名里挑第一个该模型真有的字段。跨 Odoo 版本字段改名时不至于直接崩。"""
        have = self.fields_of(model)
        for c in candidates:
            if c in have:
                return c
        if required:
            raise OdooError(f"{model} 上找不到{what or '所需'}字段，试过：{candidates}。\n"
                            f"该模型实际字段可用 fields_get 查看（跑 odoo_api/discover.py）。")
        return None

    def field_by_label(self, models, label):
        """按**界面标签**反查技术字段名——自定义字段（x_studio_*）的技术名不可硬编码猜。

        返回 [(model, field_name, label), …]，按 models 给的顺序排。
        """
        rows = self.search_read_all(
            "ir.model.fields",
            [("model", "in", list(models)), ("field_description", "=", label)],
            ["model", "name", "field_description", "ttype", "store"])
        order = {m: i for i, m in enumerate(models)}
        rows.sort(key=lambda r: order.get(r["model"], 99))
        return rows


def m2o_id(val):
    """Odoo 的 many2one 读出来是 [id, "显示名"]，空是 False。取 id。"""
    return val[0] if isinstance(val, (list, tuple)) and val else None


def m2o_name(val):
    return val[1] if isinstance(val, (list, tuple)) and len(val) > 1 else ""
