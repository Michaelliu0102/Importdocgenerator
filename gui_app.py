"""报关资料生成器 - 图形界面.

需要 Python + Tk 8.6+（推荐 Tk 9.0）。
使用 run_gui.sh 或 run_gui.command 启动。
"""

import threading
import subprocess
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from main import CustomsDocGenerator

APP_VERSION = "v4.4"

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TKINTERDND2_IMPORTED = True
except Exception:
    TkinterDnD = None
    DND_FILES = None
    TKINTERDND2_IMPORTED = False

# Set True in main() only if TkinterDnD.Tk() initializes (matches this Tk build).
HAS_DND = False

BG            = "#f0f4f8"
DROP_BG_IMPORT = "#d4e6ff"
DROP_BG_EXPORT = "#d8f0e0"
DROP_BD_IMPORT = "#5599dd"
DROP_BD_EXPORT = "#44aa77"
LIST_BG       = "#ffffff"
STAT_BG       = "#fffde6"
ACCENT        = "#2060b0"
BTN_GRAY      = "#aaaaaa"
FG            = "#1a1a1a"
FG_DIM        = "#666666"


class CustomsDocGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"报关资料生成器 {APP_VERSION}")
        self.root.geometry("1100x900")
        self.root.minsize(960, 720)
        self.root.configure(bg=BG)

        self.base_dir = Path(__file__).resolve().parent
        self.import_paths: list[str] = []
        self.import_display: list[str] = []
        self.export_paths: list[str] = []
        self.export_display: list[str] = []
        self.last_output_folder = None
        self._default_output_dir = str(self.base_dir / "output")
        self.output_dir_var = tk.StringVar(value=self._default_output_dir)
        self.ocr_enabled_var = tk.BooleanVar(value=True)
        self.ocr_lang_var = tk.StringVar(value="eng")
        self.export_fx_var = tk.StringVar(value="")
        self._want_eur_invoice_pdf = False  # toggled by Checkbutton command callback
        self.eur_standalone_pdf_var = tk.StringVar(value="")
        self._eur_only_btn_enabled = True
        self.status_var = tk.StringVar(
            value="左侧进口 Invoice PDF；右侧出口请同时添加 CustInvc 发票与 ItemShip 装箱单 PDF；选择输出目录后生成。")

        self._build_ui()

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=BG, padx=18, pady=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        row = 0

        tk.Label(
            outer,
            text=f"进口 / 出口 报关资料生成   {APP_VERSION}",
            font=("Helvetica", 17, "bold"),
            fg=FG, bg=BG, anchor="w",
        ).grid(row=row, column=0, sticky="ew", pady=(0, 6))
        row += 1

        # ── 双栏：进口 | 出口 ─────────────────────────────────
        dual = tk.Frame(outer, bg=BG)
        dual.grid(row=row, column=0, sticky="nsew", pady=(0, 8))
        dual.columnconfigure(0, weight=1)
        dual.columnconfigure(1, weight=1)
        dual.rowconfigure(1, weight=1)
        row += 1

        tk.Label(
            dual, text="进口资料（Invoice PDF）",
            font=("Helvetica", 11, "bold"), fg=FG, bg=BG, anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        tk.Label(
            dual, text="出口资料（发票/装箱单等，PDF）",
            font=("Helvetica", 11, "bold"), fg=FG, bg=BG, anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        left = tk.Frame(dual, bg=BG)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        right = tk.Frame(dual, bg=BG)
        right.grid(row=1, column=1, sticky="nsew", padx=(6, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        self.drop_zone_import = tk.Canvas(
            left, height=88, bg=DROP_BG_IMPORT,
            highlightthickness=2, highlightbackground=DROP_BD_IMPORT,
            cursor="hand2",
        )
        self.drop_zone_import.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self._bind_drop_redraw(
            self.drop_zone_import,
            "将进口 PDF / 文件夹拖到此区域",
        )

        imp_btns = tk.Frame(left, bg=BG)
        imp_btns.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        tk.Button(
            imp_btns, text="添加进口文件...",
            command=self._choose_import_files,
        ).pack(side="left")
        tk.Button(
            imp_btns, text="移除选中",
            command=self._remove_selected_import,
        ).pack(side="left", padx=(8, 0))
        tk.Button(
            imp_btns, text="清空",
            command=self._clear_import,
        ).pack(side="left", padx=(8, 0))

        limp = tk.Frame(left, bg=BG)
        limp.grid(row=2, column=0, sticky="nsew")
        limp.columnconfigure(0, weight=1)
        limp.rowconfigure(0, weight=1)
        self.file_listbox_import = tk.Listbox(
            limp, selectmode="extended", height=6,
            font=("Helvetica", 11),
            bg=LIST_BG, fg=FG,
            selectbackground=ACCENT, selectforeground="white",
            bd=1, relief="solid",
        )
        self.file_listbox_import.grid(row=0, column=0, sticky="nsew")
        vsb_i = tk.Scrollbar(
            limp, orient="vertical",
            command=self.file_listbox_import.yview,
        )
        vsb_i.grid(row=0, column=1, sticky="ns")
        self.file_listbox_import.config(yscrollcommand=vsb_i.set)
        self.file_listbox_import.bind(
            "<Double-Button-1>",
            lambda e: self._remove_selected_import(),
        )

        self.drop_zone_export = tk.Canvas(
            right, height=88, bg=DROP_BG_EXPORT,
            highlightthickness=2, highlightbackground=DROP_BD_EXPORT,
            cursor="hand2",
        )
        self.drop_zone_export.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self._bind_drop_redraw(
            self.drop_zone_export,
            "将出口 PDF / 文件夹拖到此区域",
        )

        exp_btns = tk.Frame(right, bg=BG)
        exp_btns.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        tk.Button(
            exp_btns, text="添加出口文件...",
            command=self._choose_export_files,
        ).pack(side="left")
        tk.Button(
            exp_btns, text="移除选中",
            command=self._remove_selected_export,
        ).pack(side="left", padx=(8, 0))
        tk.Button(
            exp_btns, text="清空",
            command=self._clear_export,
        ).pack(side="left", padx=(8, 0))

        lexp = tk.Frame(right, bg=BG)
        lexp.grid(row=2, column=0, sticky="nsew")
        lexp.columnconfigure(0, weight=1)
        lexp.rowconfigure(0, weight=1)
        self.file_listbox_export = tk.Listbox(
            lexp, selectmode="extended", height=6,
            font=("Helvetica", 11),
            bg=LIST_BG, fg=FG,
            selectbackground=ACCENT, selectforeground="white",
            bd=1, relief="solid",
        )
        self.file_listbox_export.grid(row=0, column=0, sticky="nsew")
        vsb_e = tk.Scrollbar(
            lexp, orient="vertical",
            command=self.file_listbox_export.yview,
        )
        vsb_e.grid(row=0, column=1, sticky="ns")
        self.file_listbox_export.config(yscrollcommand=vsb_e.set)
        self.file_listbox_export.bind(
            "<Double-Button-1>",
            lambda e: self._remove_selected_export(),
        )

        # ── 输出目录 ──────────────────────────────────────────
        out_f = tk.Frame(outer, bg=BG)
        out_f.grid(row=row, column=0, sticky="ew", pady=4)
        row += 1
        tk.Label(out_f, text="输出目录:", fg=FG, bg=BG,
                 font=("Helvetica", 11), anchor="w").pack(side="left")
        self.output_dir_entry = tk.Entry(
            out_f,
            textvariable=self.output_dir_var,
            fg=FG,
            bg="white",
            font=("Helvetica", 11),
        )
        self.output_dir_entry.pack(
            side="left", fill="x", expand=True, padx=(6, 8))
        tk.Button(out_f, text="选择目录...",
                  command=self._choose_output_dir).pack(side="left")

        # ── OCR 设置 ──────────────────────────────────────────
        ocr_f = tk.Frame(outer, bg=BG)
        ocr_f.grid(row=row, column=0, sticky="ew", pady=4)
        row += 1
        tk.Checkbutton(ocr_f, text="自动 OCR（扫描件推荐开启，仅进口解析）",
                       variable=self.ocr_enabled_var,
                       bg=BG, fg=FG, activebackground=BG,
                       font=("Helvetica", 11)).pack(side="left")
        tk.Label(ocr_f, text="OCR 语言:", fg=FG, bg=BG,
                 font=("Helvetica", 11)).pack(side="left", padx=(18, 4))
        tk.Entry(ocr_f, textvariable=self.ocr_lang_var, width=12,
                 fg=FG, bg="white",
                 font=("Helvetica", 11)).pack(side="left")
        tk.Label(ocr_f, text="如: eng / eng+ita",
                 fg=FG_DIM, bg=BG,
                 font=("Helvetica", 10)).pack(side="left", padx=(10, 0))

        # ── 出口：是否生成 EUR 版 Invoice PDF ─────────────────
        eur_pdf_f = tk.Frame(outer, bg=BG)
        eur_pdf_f.grid(row=row, column=0, sticky="ew", pady=2)
        row += 1
        self._eur_invoice_pdf_cb = tk.Checkbutton(
            eur_pdf_f,
            text="生成 EUR 版 Invoice PDF（CustInvc_…，在原 PDF 上替换为欧元；未勾选则不生成该 PDF）",
            command=self._toggle_eur_pdf,
            bg=BG,
            fg=FG,
            activebackground=BG,
            font=("Helvetica", 11),
        )
        self._eur_invoice_pdf_cb.pack(side="left", anchor="w")

        # ── 出口 EUR 汇率（可选）──────────────────────────────
        fx_f = tk.Frame(outer, bg=BG)
        fx_f.grid(row=row, column=0, sticky="ew", pady=4)
        row += 1
        tk.Label(
            fx_f,
            text="出口 EUR 汇率：1 EUR =",
            fg=FG, bg=BG,
            font=("Helvetica", 11),
        ).pack(side="left")
        # 保存引用：焦点仍在输入框时点「开始生成」时，StringVar 可能尚未同步，需用 Entry.get()。
        self.fx_entry = tk.Entry(
            fx_f,
            textvariable=self.export_fx_var,
            width=14,
            fg=FG,
            bg="white",
            font=("Helvetica", 11),
        )
        self.fx_entry.pack(side="left", padx=(6, 4))
        tk.Label(
            fx_f,
            text="单位发票货币（非 EUR 时填汇率；留空则出口合同/报关单等仍用原币种）",
            fg=FG_DIM,
            bg=BG,
            font=("Helvetica", 10),
        ).pack(side="left")

        # ── 独立：仅生成 EUR 版 CustInvc PDF ─────────────────
        eur_only_f = tk.Frame(outer, bg=BG)
        eur_only_f.grid(row=row, column=0, sticky="ew", pady=(10, 4))
        row += 1
        tk.Label(
            eur_only_f,
            text="EUR Invoice PDF（单独）：上传非 EUR 的 CustInvc 发票 PDF，按上方汇率生成欧元版（≤1MB 尽量压缩）",
            fg=FG,
            bg=BG,
            font=("Helvetica", 11, "bold"),
            anchor="w",
        ).pack(anchor="w")
        eur_only_row = tk.Frame(eur_only_f, bg=BG)
        eur_only_row.pack(fill="x", pady=(6, 0))
        eur_only_row.columnconfigure(0, weight=1)
        self.eur_standalone_entry = tk.Entry(
            eur_only_row,
            textvariable=self.eur_standalone_pdf_var,
            fg=FG,
            bg="white",
            insertbackground=FG,
            font=("Helvetica", 10),
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground="#cccccc",
            highlightcolor=ACCENT,
        )
        self.eur_standalone_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        tk.Button(
            eur_only_row,
            text="选择发票 PDF…",
            command=self._choose_eur_standalone_pdf,
            width=14,
        ).grid(row=0, column=1, sticky="e")
        # macOS 上 tk.Button 的 bg/fg 常被系统主题覆盖，改用 Label 保证蓝底白字
        self.eur_only_btn = tk.Label(
            eur_only_f,
            text="生成 EUR Invoice PDF",
            font=("Helvetica", 11, "bold"),
            fg="white",
            bg=ACCENT,
            padx=22,
            pady=10,
            cursor="hand2",
            relief="raised",
            bd=1,
        )
        self.eur_only_btn.pack(anchor="w", pady=(8, 0))
        self.eur_only_btn.bind("<Button-1>", self._on_eur_only_btn_click)
        self.eur_only_btn.bind(
            "<ButtonRelease-1>",
            lambda e: self.eur_only_btn.config(relief="raised"),
        )

        # ── 操作按钮 ──────────────────────────────────────────
        act_f = tk.Frame(outer, bg=BG)
        act_f.grid(row=row, column=0, sticky="ew", pady=(12, 4))
        row += 1
        self._run_btn_enabled = False
        self.run_btn = tk.Label(
            act_f, text="▶  开始生成",
            font=("Helvetica", 13, "bold"),
            fg="#dddddd", bg=BTN_GRAY,
            padx=30, pady=12,
            cursor="arrow",
            relief="raised", bd=1,
        )
        self.run_btn.pack(side="left")
        self.run_btn.bind("<Button-1>", self._on_run_btn_click)
        self.run_btn.bind("<ButtonRelease-1>",
                          lambda e: self.run_btn.config(relief="raised"))
        tk.Button(act_f, text="打开输出目录",
                  command=self._open_output_dir,
                  width=14).pack(side="left", padx=(12, 0))

        # ── 状态栏 ────────────────────────────────────────────
        stat_f = tk.Frame(outer, bg="#aaaaaa", padx=1, pady=1)
        stat_f.grid(row=row, column=0, sticky="ew", pady=(10, 0))
        self.status_label = tk.Label(
            stat_f, textvariable=self.status_var,
            fg=FG, bg=STAT_BG,
            font=("Helvetica", 11),
            anchor="w", justify="left",
            padx=10, pady=6,
        )
        self.status_label.pack(fill="x")

        # ── 注册 DnD ──────────────────────────────────────────
        if HAS_DND:
            for w in (
                self.drop_zone_import,
                self.drop_zone_export,
                self.file_listbox_import,
                self.file_listbox_export,
            ):
                try:
                    w.drop_target_register(DND_FILES)
                except tk.TclError:
                    continue
            self.drop_zone_import.dnd_bind(
                "<<Drop>>", self._on_drop_import)
            self.drop_zone_export.dnd_bind(
                "<<Drop>>", self._on_drop_export)
            self.file_listbox_import.dnd_bind(
                "<<Drop>>", self._on_drop_import)
            self.file_listbox_export.dnd_bind(
                "<<Drop>>", self._on_drop_export)
            self._set_status(
                "拖拽已启用。左栏进口、右栏出口；或点击「添加…文件」。")
        elif TKINTERDND2_IMPORTED:
            self._set_status(
                "请点击「添加…文件」导入 PDF。"
                "（已安装 tkinterdnd2，但 tkdnd 与当前 Tk 不兼容，拖拽不可用）")
        else:
            self._set_status(
                "请点击「添加…文件」导入 PDF。"
                "（安装 tkinterdnd2 可启用拖拽）")

        self._sync_output_dir_entry()
        self._update_run_btn()

    def _sync_output_dir_entry(self):
        """部分 macOS/Tk 下输出目录 Entry 不显示 StringVar 初值，需显式写入。"""
        od = (self.output_dir_var.get().strip() or self._default_output_dir)
        self.output_dir_var.set(od)
        ent = getattr(self, "output_dir_entry", None)
        if ent is not None:
            ent.delete(0, tk.END)
            ent.insert(0, od)

    def _bind_drop_redraw(self, canvas: tk.Canvas, main_text: str):
        def _redraw(event=None):
            canvas.delete("all")
            w = max(canvas.winfo_width(), 200)
            canvas.create_text(
                w // 2, 28,
                text=main_text,
                font=("Helvetica", 12, "bold"), fill="#1a4080",
            )
            if HAS_DND:
                hint, color = "或使用上方「添加…文件」", "#555555"
            elif TKINTERDND2_IMPORTED:
                hint, color = (
                    "⚠ 拖拽不可用 — 请用「添加…文件」",
                    "#cc0000",
                )
            else:
                hint, color = ("⚠ 拖拽不可用 — 请安装 tkinterdnd2",
                               "#cc0000")
            canvas.create_text(
                w // 2, 56,
                text=hint, font=("Helvetica", 9), fill=color,
            )

        canvas.bind("<Configure>", _redraw)
        self.root.after(80, _redraw)

    # ------------------------------------------------------------------
    def _update_run_btn(self):
        """Gray when not ready, blue when ready."""
        has_any = bool(self.import_paths or self.export_paths)
        od = self.output_dir_var.get().strip()
        if getattr(self, "output_dir_entry", None) is not None:
            od = self.output_dir_entry.get().strip() or od
        has_output = bool(od or self._default_output_dir)
        if has_any and has_output:
            self._run_btn_enabled = True
            self.run_btn.config(bg=ACCENT, fg="white", cursor="hand2")
        else:
            self._run_btn_enabled = False
            self.run_btn.config(bg=BTN_GRAY, fg="#dddddd", cursor="arrow")

    def _on_run_btn_click(self, event=None):
        if self._run_btn_enabled:
            self.run_btn.config(relief="sunken")
            self._start_generate()

    def _set_status(self, text: str):
        self.status_var.set(text)
        self.root.update_idletasks()

    def _choose_import_files(self):
        paths = filedialog.askopenfilenames(
            title="选择进口 Invoice PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        self._add_import_files(paths)

    def _choose_export_files(self):
        paths = filedialog.askopenfilenames(
            title="选择出口资料 PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        self._add_export_files(paths)

    def _choose_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            p = str(Path(path).expanduser().resolve())
            self.output_dir_var.set(p)
            ent = getattr(self, "output_dir_entry", None)
            if ent is not None:
                ent.delete(0, tk.END)
                ent.insert(0, p)
            self.root.update_idletasks()
            self._update_run_btn()

    def _expand_paths(self, paths):
        out = []
        for p in paths:
            p = str(Path(p).expanduser()).strip()
            if not p:
                continue
            if p.lower().startswith("file://"):
                p = p[7:]
            path_obj = Path(p)
            if not path_obj.exists():
                continue
            if path_obj.is_dir():
                for pdf in sorted(path_obj.rglob("*.pdf")):
                    out.append(str(pdf.resolve()))
            elif path_obj.suffix.lower() == ".pdf":
                out.append(str(path_obj.resolve()))
        seen = set()
        unique = []
        for x in out:
            if x not in seen:
                seen.add(x)
                unique.append(x)
        return unique

    def _add_import_files(self, paths):
        if not paths:
            return
        for path in self._expand_paths(paths):
            if path not in self.import_paths:
                self.import_paths.append(path)
                self.import_display.append(Path(path).name)
                self.file_listbox_import.insert("end", Path(path).name)
        self._update_run_btn()

    def _add_export_files(self, paths):
        if not paths:
            return
        for path in self._expand_paths(paths):
            if path not in self.export_paths:
                self.export_paths.append(path)
                self.export_display.append(Path(path).name)
                self.file_listbox_export.insert("end", Path(path).name)
        self._update_run_btn()

    def _remove_selected_import(self):
        selected = list(self.file_listbox_import.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            self.file_listbox_import.delete(idx)
            self.import_paths.pop(idx)
            self.import_display.pop(idx)
        self._update_run_btn()

    def _remove_selected_export(self):
        selected = list(self.file_listbox_export.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            self.file_listbox_export.delete(idx)
            self.export_paths.pop(idx)
            self.export_display.pop(idx)
        self._update_run_btn()

    def _clear_import(self):
        self.file_listbox_import.delete(0, "end")
        self.import_paths.clear()
        self.import_display.clear()
        self._update_run_btn()

    def _clear_export(self):
        self.file_listbox_export.delete(0, "end")
        self.export_paths.clear()
        self.export_display.clear()
        self._update_run_btn()

    def _parse_drop_paths(self, event):
        raw = getattr(event, "data", "") or ""
        paths = []
        try:
            paths = list(self.root.tk.splitlist(raw))
        except tk.TclError:
            paths = []
        if not paths:
            s = raw.strip()
            if s.startswith("{") and s.endswith("}"):
                s = s[1:-1]
            if s:
                paths = [s]
        return paths

    def _on_drop_import(self, event):
        self._add_import_files(self._parse_drop_paths(event))

    def _on_drop_export(self, event):
        self._add_export_files(self._parse_drop_paths(event))

    def _open_output_dir(self):
        out_dir = (Path(self.last_output_folder) if self.last_output_folder
                   else Path(self.output_dir_var.get()).expanduser())
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.run(["open", str(out_dir)], check=True)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开输出目录:\n{e}")

    def _toggle_eur_pdf(self):
        """Checkbutton command callback — flip the plain Python bool each click."""
        self._want_eur_invoice_pdf = not self._want_eur_invoice_pdf

    def _choose_eur_standalone_pdf(self):
        path = filedialog.askopenfilename(
            title="选择 CustInvc 发票 PDF",
            filetypes=[("PDF", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            p = str(Path(path).expanduser().resolve())
            self.eur_standalone_pdf_var.set(p)
            # 部分环境下仅靠 StringVar 不刷新显示，同步写入 Entry
            ent = getattr(self, "eur_standalone_entry", None)
            if ent is not None:
                ent.delete(0, tk.END)
                ent.insert(0, p)
                ent.xview_moveto(1.0)
            self.root.update_idletasks()

    def _on_eur_only_btn_click(self, event=None):
        if not self._eur_only_btn_enabled:
            return
        self._start_eur_standalone()

    def _start_eur_standalone(self):
        pdf = self.eur_standalone_pdf_var.get().strip()
        out = self.output_dir_var.get().strip() or self._default_output_dir
        if not pdf:
            messagebox.showwarning("提示", "请先选择发票 PDF")
            return
        if not Path(pdf).exists():
            messagebox.showerror("错误", "所选文件不存在")
            return
        fx_raw = (
            self.fx_entry.get().strip()
            if getattr(self, "fx_entry", None) is not None
            else self.export_fx_var.get().strip()
        )
        fx = None
        if fx_raw:
            try:
                fx = float(fx_raw.replace(",", "."))
                if fx <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "汇率无效",
                    "请填写大于 0 的汇率，或留空（仅当发票已为 EUR 时）。",
                )
                return
        self._eur_only_btn_enabled = False
        self.eur_only_btn.config(
            bg=BTN_GRAY, fg="#eeeeee", cursor="arrow", relief="raised")
        self._set_status("正在生成 EUR Invoice PDF…")
        # Tk 变量必须在主线程读取，不可在后台线程里 .get()
        ocr_lang = (self.ocr_lang_var.get().strip() or "eng")
        enable_ocr = self.ocr_enabled_var.get()
        threading.Thread(
            target=self._eur_standalone_worker,
            args=(pdf, out, fx, enable_ocr, ocr_lang),
            daemon=True,
        ).start()

    def _eur_standalone_worker(self, pdf, out, fx, enable_ocr, ocr_lang):
        try:
            from eur_invoice_standalone import convert_custinvc_to_eur_pdf_safe

            path, err = convert_custinvc_to_eur_pdf_safe(
                pdf,
                out,
                fx_units_per_eur=fx,
                enable_ocr=enable_ocr,
                ocr_lang=ocr_lang,
            )
            if path:
                msg = f"已生成:\n{path}"
                if err:
                    msg += f"\n\n提示: {err}"
                self.root.after(
                    0,
                    lambda m=msg: messagebox.showinfo("完成", m),
                )
                self.root.after(
                    0,
                    lambda p=path: self._set_status(f"EUR PDF: {p}"),
                )
                self.last_output_folder = str(Path(path).parent)
            else:
                self.root.after(
                    0,
                    lambda e=err: messagebox.showerror("失败", e or "未知错误"),
                )
                self.root.after(
                    0,
                    lambda: self._set_status("EUR PDF 生成失败"),
                )
        finally:
            self.root.after(0, self._eur_only_btn_reset)

    def _eur_only_btn_reset(self):
        self._eur_only_btn_enabled = True
        self.eur_only_btn.config(
            bg=ACCENT, fg="white", cursor="hand2", relief="raised")

    def _read_eur_invoice_pdf_wanted(self) -> bool:
        return self._want_eur_invoice_pdf

    def _start_generate(self):
        if getattr(self, "output_dir_entry", None) is not None:
            output_dir = self.output_dir_entry.get().strip()
        else:
            output_dir = self.output_dir_var.get().strip()
        output_dir = output_dir or self._default_output_dir
        if not self.import_paths and not self.export_paths:
            messagebox.showwarning(
                "提示",
                "请至少在左侧或右侧添加一个 PDF 文件",
            )
            return
        for label, paths in (
            ("进口", self.import_paths),
            ("出口", self.export_paths),
        ):
            missing = [p for p in paths if not Path(p).exists()]
            if missing:
                messagebox.showerror(
                    "错误",
                    f"{label}文件不存在，请检查:\n{missing[0]}",
                )
                return

        self.root.update_idletasks()
        # 以控件选中态为准（与汇率框同理，避免 Var 与 Tk 内部状态不一致）
        generate_eur_invoice_pdf = self._read_eur_invoice_pdf_wanted()
        fx_raw = (
            self.fx_entry.get().strip()
            if getattr(self, "fx_entry", None) is not None
            else self.export_fx_var.get().strip()
        )
        fx_units_per_eur = None
        if fx_raw:
            try:
                fx_units_per_eur = float(fx_raw.replace(",", "."))
                if fx_units_per_eur <= 0:
                    raise ValueError("rate must be positive")
            except ValueError:
                messagebox.showerror(
                    "汇率无效",
                    "请在「1 EUR = … 发票货币」中填写大于 0 的数字，或留空不换算。",
                )
                return

        enable_ocr = self.ocr_enabled_var.get()
        ocr_lang = self.ocr_lang_var.get().strip() or "eng"
        import_paths = list(self.import_paths)
        export_paths = list(self.export_paths)

        total = len(import_paths) + (1 if export_paths else 0)
        self._run_btn_enabled = False
        self.run_btn.config(bg=BTN_GRAY, fg="#dddddd", cursor="arrow")
        self._set_status(f"正在处理 {total} 个任务，请稍候...")
        threading.Thread(
            target=self._generate,
            args=(
                output_dir,
                enable_ocr,
                ocr_lang,
                import_paths,
                export_paths,
                fx_units_per_eur,
                generate_eur_invoice_pdf,
            ),
            daemon=True,
        ).start()

    def _generate(
        self,
        output_dir,
        enable_ocr,
        ocr_lang,
        import_paths,
        export_paths,
        fx_units_per_eur=None,
        generate_eur_invoice_pdf=False,
    ):
        try:
            all_files = []
            success_import = 0
            success_export = 0
            from datetime import datetime
            batch_folder = (
                Path(output_dir).expanduser()
                / f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
            batch_folder.mkdir(parents=True, exist_ok=True)
            self.last_output_folder = str(batch_folder)

            n_imp = len(import_paths)
            n_exp = len(export_paths)
            exp_batch = 1 if n_exp else 0
            total_tasks = n_imp + exp_batch
            task_i = 0

            for i, invoice_path in enumerate(import_paths, start=1):
                task_i += 1
                self.root.after(
                    0,
                    lambda ii=i, ti=task_i, ni=n_imp, tt=total_tasks:
                    self._set_status(
                        f"进口 {ii}/{ni}（总 {ti}/{tt}）..."),
                )
                raw_stem = Path(invoice_path).stem.strip() or f"invoice_{i}"
                invoice_stem = re.sub(r'[\\/:*?"<>|]+', "_", raw_stem)
                invoice_out = batch_folder / f"import_{i:02d}_{invoice_stem}"
                invoice_out.mkdir(parents=True, exist_ok=True)

                generator = CustomsDocGenerator(
                    config_path=str(
                        self.base_dir / "data"
                        / "supplier_product_mapping_import.yaml"
                    ),
                    templates_dir=str(self.base_dir / "templates"),
                )
                result = generator.process_invoice(
                    invoice_path=invoice_path,
                    output_dir=str(invoice_out),
                    enable_ocr=enable_ocr,
                    ocr_lang=ocr_lang,
                )
                files = []
                for key in ("contract", "customs_declaration",
                            "declaration_elements"):
                    files.extend(result.get(key, []))
                if files:
                    success_import += 1
                    all_files.extend(files)

            if export_paths:
                task_i += 1
                self.root.after(
                    0,
                    lambda ti=task_i, tt=total_tasks:
                    self._set_status(
                        f"出口（发票+装箱单，总 {ti}/{tt}）..."),
                )
                exp_out = batch_folder / "export_pack"
                exp_out.mkdir(parents=True, exist_ok=True)

                generator = CustomsDocGenerator(
                    config_path=str(
                        self.base_dir / "data"
                        / "supplier_product_mapping_export.yaml"
                    ),
                    templates_dir=str(self.base_dir / "templates"),
                    export_templates_dir=str(
                        self.base_dir / "export_templates"
                    ),
                )
                result = generator.process_export_documents(
                    source_paths=export_paths,
                    output_dir=str(exp_out),
                    enable_ocr=enable_ocr,
                    ocr_lang=ocr_lang,
                    fx_units_per_eur=fx_units_per_eur,
                    generate_eur_invoice_pdf=generate_eur_invoice_pdf,
                )
                ex_files = result.get("export", [])
                if ex_files:
                    success_export += 1
                    all_files.extend(ex_files)

            if not all_files:
                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "完成",
                        "处理完成，但未生成文件。请检查模板或解析结果。"),
                )
                self.root.after(
                    0,
                    lambda: self._set_status("未生成文件，请查看日志。"),
                )
            else:
                preview = "\n".join(all_files[:12])
                if len(all_files) > 12:
                    preview += f"\n... 还有 {len(all_files) - 12} 个文件"
                msg = (
                    f"处理完成\n"
                    f"进口成功: {success_import}/{n_imp}，"
                    f"出口成功: {success_export}/{exp_batch}\n"
                    f"输出目录: {self.last_output_folder}\n\n"
                    f"已生成文件:\n{preview}"
                )
                self.root.after(
                    0, lambda: messagebox.showinfo("生成成功", msg))
                self.root.after(
                    0,
                    lambda: self._set_status(
                        f"完成：进口 {success_import}/{n_imp}，"
                        f"出口 {success_export}/{exp_batch}，"
                        f"共 {len(all_files)} 个文件。"),
                )
                self.root.after(0, self._open_output_dir)
        except Exception as e:
            # Python 3.12+ 在离开 except 后会清除异常名 e，延后执行的 lambda 不能引用 e
            err_msg = str(e)
            self.root.after(
                0,
                lambda: messagebox.showerror("处理失败", f"生成失败:\n{err_msg}"),
            )
            self.root.after(
                0,
                lambda: self._set_status(f"处理失败: {err_msg}"),
            )
        finally:
            self.root.after(0, self._update_run_btn)


def main():
    global HAS_DND
    root = None
    if TKINTERDND2_IMPORTED:
        try:
            root = TkinterDnD.Tk()
            HAS_DND = True
        except Exception:
            HAS_DND = False
            root = None
    if root is None:
        root = tk.Tk()
    CustomsDocGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
