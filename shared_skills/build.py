"""共享版 skill 打包：仓库源 → 仓库外独立目录里的 skill 文件夹 + zip + 使用说明 + 脱敏样例（ISSUES [发票改价] C）。

用法：
    python3 shared_skills/build.py                  # 默认输出到 <仓库>/../shared-skills/
    python3 shared_skills/build.py -o D:/share      # 指定输出目录
    python3 shared_skills/build.py --no-smoke       # 跳过冒烟（不建议）

每个共享 skill 在 SKILLS 里登记：源 SKILL.md 在 shared_skills/<name>/，脚本从仓库现有代码**原样复制**
（不分叉、不改写），所以仓库脚本改了重跑本脚本即可。产出（全部重新生成）：
    <out>/<name>/                 skill 文件夹（B 方式：整个拷到 ~/.claude/skills/<name>/）
    <out>/<name>.zip              zip 顶层即 <name>/（A 方式：claude.ai Customize > Skills 上传）
    <out>/<name>.使用说明.md / <name>.guide.md     给同事的说明（shared_skills/guides/）
    <out>/samples/<name>/         合成样例（不含任何真实客户 / 价格），给同事试跑
校验不过（name 格式、description 长度、脚本里出现 odoo_api / config / 绝对路径）或冒烟失败就退出码 2、不出 zip。
"""
import argparse
import datetime as dt
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

from openpyxl import Workbook, load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
MARKER = ".generated-by-shared_skills-build"     # 有它才允许整目录重建，防误删别人的目录

#: skill 名 → {skill 内路径: 仓库内源路径}
SKILLS = {
    "invoice-reprice": {
        "scripts/invoice_reprice.py": "procure/invoice_reprice.py",
        "scripts/common/__init__.py": "common/__init__.py",
        "scripts/common/xlsx.py": "common/xlsx.py",
    },
}

FORBIDDEN = [
    (re.compile(r"^\s*(from|import)\s+odoo_api", re.M), "import odoo_api（要连 Odoo）"),
    (re.compile(r"^\s*(from|import)\s+config\b", re.M), "import config（本机配置）"),
    (re.compile(r"/Users/|[A-Za-z]:\\\\|~/"), "绝对路径"),
    (re.compile(r"\.\./"), "指向 skill 外的相对路径"),
]


class BuildError(RuntimeError):
    pass


def say(msg=""):
    print(msg, flush=True)


# --------------------------------------------------------------------------
def frontmatter(path):
    txt = open(path, encoding="utf-8").read()
    m = re.match(r"^---\n(.*?)\n---\n", txt, re.S)
    if not m:
        raise BuildError(f"{path}：没有 frontmatter")
    fm = dict(re.findall(r"^(\w+):\s*(.*)$", m.group(1), re.M))
    return fm


def check_skill(name, src_dir, files):
    fm = frontmatter(os.path.join(src_dir, "SKILL.md"))
    if fm.get("name") != name:
        raise BuildError(f"{name}：frontmatter name 是「{fm.get('name')}」，应与文件夹同名")
    if not re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", name) or len(name) > 64:
        raise BuildError(f"{name}：name 只能是小写字母 / 数字 / 连字符，≤ 64 字符")
    desc = fm.get("description", "").strip()
    if not desc or len(desc) > 1024 or "<" in desc or ">" in desc:
        raise BuildError(f"{name}：description 为空、超过 1024 字符或含尖括号（{len(desc)} 字符）")
    for rel, src in files.items():
        txt = open(os.path.join(REPO, src), encoding="utf-8").read()
        for rx, what in FORBIDDEN:
            m = rx.search(txt)
            if m:
                line = txt[:m.start()].count("\n") + 1
                raise BuildError(f"{src}:{line} 含{what}，不能进共享 skill")
    return desc


def build_skill(name, files, out):
    src_dir = os.path.join(HERE, name)
    desc = check_skill(name, src_dir, files)
    target = os.path.join(out, name)
    if os.path.exists(target):
        if not os.path.exists(os.path.join(target, MARKER)):
            raise BuildError(f"{target} 已存在且不是本脚本生成的，不覆盖。请换输出目录或手动挪走。")
        shutil.rmtree(target)
    os.makedirs(target)
    shutil.copy2(os.path.join(src_dir, "SKILL.md"), os.path.join(target, "SKILL.md"))
    for rel, src in files.items():
        dst = os.path.join(target, *rel.split("/"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(os.path.join(REPO, src), dst)
    with open(os.path.join(target, MARKER), "w", encoding="utf-8") as f:
        f.write(f"generated {dt.datetime.now():%Y-%m-%d %H:%M} from ERP_Excel_Convert; do not edit here\n")
    say(f"✓ {name}：skill 文件夹 {target}（description {len(desc)} 字符）")
    return target


def zip_skill(name, target, out):
    zpath = os.path.join(out, f"{name}.zip")
    if os.path.exists(zpath):
        os.remove(zpath)
    n = 0
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, fnames in os.walk(target):
            dirs[:] = sorted(d for d in dirs if d != "__pycache__")
            for fn in sorted(fnames):
                if fn == MARKER or fn.endswith(".pyc"):
                    continue
                full = os.path.join(root, fn)
                z.write(full, name + "/" + os.path.relpath(full, target).replace(os.sep, "/"))
                n += 1
    say(f"✓ {name}：{zpath}（{n} 个文件，{os.path.getsize(zpath) / 1024:.0f} KB）")
    return zpath


def copy_guides(name, out):
    gdir = os.path.join(HERE, "guides")
    for fn in sorted(os.listdir(gdir)) if os.path.isdir(gdir) else []:
        if fn.startswith(name + "."):
            shutil.copy2(os.path.join(gdir, fn), os.path.join(out, fn))
            say(f"✓ {name}：说明 {os.path.join(out, fn)}")


# --------------------------------------------------------------------------
# invoice-reprice：合成样例 + 冒烟（全部虚构：客户、编码、价格都不是真的）
# --------------------------------------------------------------------------
INV_HEADER = ["Status", "ID", "Invoice/Bill Date", "Date", "Partner", "Untaxed Amount", "Total",
              "Invoice lines/ID", "Invoice lines/Display Type", "Invoice lines/Sequence",
              "Invoice lines/Product/Internal Reference", "Invoice lines/Label", "Invoice lines/Quantity",
              "Invoice lines/Unit Price", "Invoice lines/Discount (%)", "Invoice lines/Subtotal",
              "Invoice lines/Total"]
PL_HEADER = ["Pricelist Name", "Pricelist Items/Product/Internal Reference", "Pricelist Items/Product/Name",
             "Pricelist Items/Fixed Price", "Pricelist Items/Min. Quantity", "Pricelist Items/Start Date",
             "Pricelist Items/End Date", "Pricelist Items/Compute Price"]
MOVE = "__export__.account_move_900001_demo"
#: (行外部 ID, 编码, 品名, 数量, 原价) —— 第一行是备注行
DEMO_LINES = [
    ("__export__.account_move_line_1_demo", "", "Delivery note: demo order", 0, 0.0),
    ("__export__.account_move_line_2_demo", "DemoCream_01234567", "Demo Cream 50 ml", 10, 4.00),
    ("__export__.account_move_line_3_demo", "DemoDrops_07654321_GW", "Demo Drops 20 ml (GW)", 5, 6.00),
    ("__export__.account_move_line_4_demo", "DemoTablets_01111111", "Demo Tablets 30 pcs", 2, 9.90),
]
DEMO_EXPECT = {"__export__.account_move_line_2_demo": 3.50,    # 两条接续，取 09-01 起那条
               "__export__.account_move_line_3_demo": 5.40}    # 两条都有效，取 Start 最新（⚠ 多条有效）


def _save(path, header, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def demo_invoice(path, prices=None):
    prices = prices or {}
    rows, total = [], 0.0
    for k, (lid, code, label, qty, price) in enumerate(DEMO_LINES):
        price = prices.get(lid, price)
        sub = round(qty * price, 2)
        total += sub
        note = k == 0
        rows.append([None, None, None, None, None, None, None, lid, "line_note" if note else "product", k + 1,
                     code or None, label, None if note else qty, None if note else price,
                     None if note else 0, None if note else sub, None if note else sub])
    rows[0][:7] = ["Draft", MOVE, dt.datetime(2026, 11, 10), dt.datetime(2026, 11, 10),
                   "Sample Customer Ltd", round(total, 2), round(total, 2)]
    return _save(path, INV_HEADER, rows)


def demo_pricelist(path):
    d = dt.datetime
    rows = [
        ["001-Sample Customer Ltd", "DemoCream_01234567", "Demo Cream 50 ml", 3.80, 0, d(2026, 7, 1),
         d(2026, 8, 31), "Fixed Price"],
        [None, "DemoCream_01234567", "Demo Cream 50 ml", 3.50, 0, d(2026, 9, 1), None, "Fixed Price"],
        [None, "DemoDrops_07654321_GW", "Demo Drops 20 ml (GW)", 5.20, 0, d(2026, 7, 1), None, "Fixed Price"],
        [None, "DemoDrops_07654321_GW", "Demo Drops 20 ml (GW)", 5.40, 0, d(2026, 10, 1), None, "Fixed Price"],
        [None, "DemoDrops_07654321", "Demo Drops 20 ml", 4.90, 0, d(2026, 7, 1), None, "Fixed Price"],
    ]
    return _save(path, PL_HEADER, rows)


def _run(script, args, cwd):
    p = subprocess.run([sys.executable, "-E", script] + args, cwd=cwd, capture_output=True,
                       encoding="utf-8", errors="replace")
    return p.returncode, p.stdout + p.stderr


def smoke_invoice_reprice(target, out):
    samples = os.path.join(out, "samples", "invoice-reprice")
    os.makedirs(samples, exist_ok=True)
    inv = demo_invoice(os.path.join(samples, "demo_invoice_export.xlsx"))
    pl = demo_pricelist(os.path.join(samples, "demo_pricelist_export.xlsx"))
    script = os.path.join(target, "scripts", "invoice_reprice.py")
    with tempfile.TemporaryDirectory() as tmp:               # 仓库外的空目录里跑，证明不依赖仓库
        code, log = _run(script, ["build", inv, pl, "-o", os.path.join(tmp, "out")], tmp)
        if code != 0:
            raise BuildError(f"冒烟 build 退出码 {code}：\n{log}")
        files = os.listdir(os.path.join(tmp, "out"))
        imp = [f for f in files if "导入表" in f]
        snap = [f for f in files if "改前快照" in f]
        if len(imp) != 1 or len(snap) != 1 or not any("核对表" in f for f in files):
            raise BuildError(f"冒烟 build 产出不对：{files}")
        imp = os.path.join(tmp, "out", imp[0])
        rows = list(load_workbook(imp).active.iter_rows(values_only=True))
        got = {r[1]: r[2] for r in rows[1:]}
        if rows[0] != ("id", "invoice_line_ids/id", "invoice_line_ids/price_unit") or rows[1][0] != MOVE \
                or got != DEMO_EXPECT:
            raise BuildError(f"冒烟导入表内容不对：{rows}")
        say("  冒烟 build ✓（导入表 2 行：接续价取新、多条有效取 Start 最新；配不上的行保持原价）")
        after = demo_invoice(os.path.join(tmp, "after.xlsx"), DEMO_EXPECT)
        code, log2 = _run(script, ["verify", os.path.join(tmp, "out", snap[0]), after, "--expect", imp], tmp)
        if code != 0:
            raise BuildError(f"冒烟 verify 退出码 {code}：\n{log2}")
        say("  冒烟 verify ✓（模拟导入后的导出：只有单价与金额变化）")
    say(f"✓ invoice-reprice：样例 {samples}")
    return log


SMOKE = {"invoice-reprice": smoke_invoice_reprice}


# --------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description="共享版 skill 打包（仓库源 → 仓库外目录 + zip）")
    ap.add_argument("-o", "--outdir", default=os.path.join(os.path.dirname(REPO), "shared-skills"),
                    help="输出目录（默认 <仓库>/../shared-skills）")
    ap.add_argument("--no-smoke", action="store_true", help="跳过冒烟（不建议）")
    args = ap.parse_args(argv)
    out = os.path.abspath(args.outdir)
    if os.path.commonpath([out, REPO]) == REPO:
        say(f"✗ 输出目录 {out} 在仓库里，请放到仓库外")
        return 2
    os.makedirs(out, exist_ok=True)
    try:
        for name, files in SKILLS.items():
            target = build_skill(name, files, out)
            if not args.no_smoke and name in SMOKE:
                SMOKE[name](target, out)
                for root, dirs, _ in os.walk(target):           # 冒烟留下的字节码不随 skill 发出去
                    if "__pycache__" in dirs:
                        shutil.rmtree(os.path.join(root, "__pycache__"))
            zip_skill(name, target, out)
            copy_guides(name, out)
    except BuildError as e:
        say(f"✗ {e}")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
