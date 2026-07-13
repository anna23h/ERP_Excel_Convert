#!/usr/bin/env python3
"""Reorder Helper - standalone English GUI.

For colleagues who don't read Chinese. Pick two files, click Generate:
  - Product / demand list  (full shipping-detail workbook, or a plain PZN list)
  - Purchase order export   (Odoo purchase.order export)
-> one-row-per-product reorder decision sheet (English column names).

Kept intentionally separate from the main (Chinese) gui.py.
"""
import os
import sys
import threading
import traceback
import subprocess
from datetime import date

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# so PyInstaller one-file build can import sibling modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import reorder_helper  # noqa: E402

EXCEL_TYPES = [("Excel / CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")]

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def open_folder(path):
    """Open a folder cross-platform; create it first if missing."""
    try:
        os.makedirs(path, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(path)              # noqa: P204
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
    except Exception:
        pass


class App:
    LABEL_W = 20

    def __init__(self, root):
        self.root = root
        root.title("Reorder Helper")
        scale = 1.0
        try:
            dpi = root.winfo_fpixels("1i")
            if dpi and dpi > 96:
                scale = dpi / 96.0
        except tk.TclError:
            pass
        w, h = int(900 * scale), int(620 * scale)
        w = min(w, root.winfo_screenwidth() - 40)
        h = min(h, root.winfo_screenheight() - int(140 * scale))
        root.geometry(f"{w}x{h}")
        root.minsize(min(760, w), min(520, h))

        self.demand = tk.StringVar()
        self.po = tk.StringVar()
        # optional explicit output file; blank -> output/<today>/reorder-<today>.xlsx
        self.out = tk.StringVar()
        self.outdir = tk.StringVar(
            value=os.path.join(BASE_DIR, "output", date.today().strftime("%Y%m%d")))
        self._buttons = []
        self._build_ui()

    # ---------- UI ----------
    def _init_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Action.TButton", font=("", 11, "bold"), padding=(18, 8),
                        foreground="white", background="#2563eb", borderwidth=0)
        style.map("Action.TButton",
                  background=[("active", "#1d4ed8"), ("disabled", "#b6c2d6")])
        style.configure("Card.TLabelframe.Label", font=("", 13, "bold"), foreground="#111827")
        style.configure("Field.TLabel", font=("", 11), foreground="#111827")
        style.configure("Hint.TLabel", font=("", 9), foreground="#6b7280")

    def _file_row(self, parent, label, var, hint):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", pady=4)
        ttk.Label(fr, text=label, width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(fr, text="Browse…", width=9,
                   command=lambda: self._pick_file(var)).pack(side="left")
        ttk.Label(parent, text=hint, style="Hint.TLabel", wraplength=760,
                  justify="left").pack(anchor="w", padx=(self.LABEL_W * 7, 0), pady=(0, 4))

    def _build_ui(self):
        self._init_styles()
        root = self.root

        card = ttk.LabelFrame(root, text="Inputs", style="Card.TLabelframe", padding=12)
        card.pack(fill="x", padx=12, pady=(12, 0))
        self._file_row(card, "Product / demand list:", self.demand,
                       "Required. The full shipping-detail workbook, or a plain PZN list "
                       "(only a PZN column). Chinese-only columns (platform price / total "
                       "demand) are dropped automatically when absent.")
        self._file_row(card, "Purchase order export:", self.po,
                       "Required. The Odoo purchase.order export (row form). Provides last "
                       "vendor / price / qty / date and current stock on hand.")
        self._file_row(card, "Output file (optional):", self.out,
                       "Leave blank to save into the output folder as reorder-<date>.xlsx.")

        outrow = ttk.Frame(root)
        outrow.pack(fill="x", padx=12, pady=(8, 0))
        ttk.Label(outrow, text="Output folder:", width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(outrow, textvariable=self.outdir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(outrow, text="Browse…", width=9, command=self._pick_dir).pack(side="left")

        act = ttk.Frame(root)
        act.pack(fill="x", padx=12, pady=(12, 6))
        b = ttk.Button(act, text="▶  Generate", style="Action.TButton", command=self._run)
        b.pack(side="left")
        self._buttons.append(b)
        ttk.Button(act, text="📂 Open output folder",
                   command=lambda: open_folder(self.outdir.get())).pack(side="right")

        self.log = scrolledtext.ScrolledText(root, height=12, state="disabled",
                                             font=("Menlo", 10) if sys.platform == "darwin"
                                             else ("Consolas", 10))
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    # ---------- helpers ----------
    def _pick_file(self, var):
        p = filedialog.askopenfilename(filetypes=EXCEL_TYPES)
        if p:
            var.set(p)

    def _pick_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.outdir.set(p)

    def _write(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_busy(self, busy):
        for b in self._buttons:
            b.configure(state="disabled" if busy else "normal")

    def _bg(self, work):
        self._set_busy(True)

        def task():
            try:
                lines = work()
                self.root.after(0, lambda: self._done(lines))
            except Exception as e:
                tb = traceback.format_exc()
                self.root.after(0, lambda e=e, tb=tb: self._fail(e, tb))
        threading.Thread(target=task, daemon=True).start()

    def _done(self, lines):
        for ln in lines:
            self._write(ln)
        self._write("✅ Done\n")
        self._set_busy(False)

    def _fail(self, e, tb):
        self._write("❌ Error: " + str(e))
        self._write(tb)
        self._set_busy(False)
        messagebox.showerror("Error", str(e))

    # ---------- action ----------
    def _run(self):
        if not self.demand.get() or not self.po.get():
            messagebox.showwarning(
                "Missing files",
                "Please choose both the product/demand list and the purchase order export.")
            return
        self._write("Generating reorder sheet…")
        demand = self.demand.get().strip()
        po = self.po.get().strip()
        out = self.out.get().strip() or None
        outdir = self.outdir.get().strip()

        def work():
            # explicit output path wins; otherwise write into the chosen output folder
            out_path = out if out else os.path.join(
                outdir, f"reorder-{date.today():%Y%m%d}.xlsx")
            path, n, matched = reorder_helper.build(demand, po, out_path)
            return [f"Saved: {path}",
                    f"{n} product(s), {matched} with purchase records, "
                    f"{n - matched} without a match."]
        self._bg(work)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
