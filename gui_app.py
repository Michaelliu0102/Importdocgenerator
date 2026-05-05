"""ClearanceOS - 图形界面.

需要 Python + Tk 8.6+（推荐 Tk 9.0）。
使用 run_gui.sh 或 run_gui.command 启动。
"""

import threading
import subprocess
import re
import os
import platform
import sys
import types
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from main import CustomsDocGenerator

APP_NAME = "ClearanceOS"
APP_VERSION = "v4.5"


def _install_tix_stub():
    """Python 3.14 removed tkinter.tix; tkinterdnd2 still imports it."""
    if "tkinter.tix" in sys.modules:
        return
    stub = types.ModuleType("tkinter.tix")
    stub.Tk = tk.Tk
    sys.modules["tkinter.tix"] = stub
    setattr(tk, "tix", stub)


def _patch_tkinterdnd2_require(module):
    """Handle Python 3.14/tk 9 and macOS symbol-case quirks for tkdnd."""
    if getattr(module, "_custom_require_patched", False):
        return

    def _platform_dir() -> str:
        system = platform.system()
        machine = (
            os.environ.get("PROCESSOR_ARCHITECTURE", platform.machine())
            if system == "Windows"
            else platform.machine()
        )
        if system == "Darwin" and machine == "arm64":
            return "osx-arm64"
        if system == "Darwin" and machine == "x86_64":
            return "osx-x64"
        if system == "Linux" and machine == "aarch64":
            return "linux-arm64"
        if system == "Linux" and machine == "x86_64":
            return "linux-x64"
        if system == "Windows" and machine == "ARM64":
            return "win-arm64"
        if system == "Windows" and machine == "AMD64":
            return "win-x64"
        if system == "Windows" and machine == "x86":
            return "win-x86"
        raise RuntimeError("Platform not supported for tkdnd.")

    def _require(tkroot):
        module_path = (
            Path(module.__file__).resolve().parent / "tkdnd" / _platform_dir()
        )
        tkroot.tk.call("lappend", "auto_path", str(module_path))
        preferred_lib = next(
            (p.name for p in module_path.glob("libtcl9tkdnd*.dylib")),
            None,
        )
        if preferred_lib and tk.TkVersion >= 9.0:
            tkroot.tk.call("source", str(module_path / "tkdnd.tcl"))
            tkroot.tk.call(
                "tkdnd::initialise",
                str(module_path),
                preferred_lib,
                "Tkdnd",
            )
            version = "2.9.5"
            try:
                version = tkroot.tk.call("package", "provide", "tkdnd")
            except tk.TclError:
                pass
            if not version:
                version = "2.9.5"
                tkroot.tk.call("package", "provide", "tkdnd", version)
            module.TkdndVersion = version
            return version

        try:
            version = tkroot.tk.call("package", "require", "tkdnd")
            module.TkdndVersion = version
            return version
        except tk.TclError as exc:
            if platform.system() != "Darwin" or "tkdnd_Init" not in str(exc):
                raise RuntimeError("Unable to load tkdnd library.") from exc

        lib_name = next(
            (
                p.name
                for pattern in ("libtcl9tkdnd*.dylib", "libtkdnd*.dylib")
                for p in module_path.glob(pattern)
            ),
            None,
        )
        if not lib_name:
            raise RuntimeError("Unable to load tkdnd library.")

        tkroot.tk.call("source", str(module_path / "tkdnd.tcl"))
        tkroot.tk.call(
            "tkdnd::initialise",
            str(module_path),
            lib_name,
            "Tkdnd",
        )
        version = "2.9.3"
        try:
            version = tkroot.tk.call("package", "provide", "tkdnd")
        except tk.TclError:
            pass
        if not version:
            version = "2.9.3"
            tkroot.tk.call("package", "provide", "tkdnd", version)
        module.TkdndVersion = version
        return version

    module._require = _require
    module._custom_require_patched = True


try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _patch_tkinterdnd2_require(TkinterDnD)
    TKINTERDND2_IMPORTED = True
except ImportError as exc:
    if "tix" in str(exc).lower():
        try:
            _install_tix_stub()
            from tkinterdnd2 import TkinterDnD, DND_FILES
            _patch_tkinterdnd2_require(TkinterDnD)
            TKINTERDND2_IMPORTED = True
        except Exception:
            TkinterDnD = None
            DND_FILES = None
            TKINTERDND2_IMPORTED = False
    else:
        TkinterDnD = None
        DND_FILES = None
        TKINTERDND2_IMPORTED = False
except Exception:
    TkinterDnD = None
    DND_FILES = None
    TKINTERDND2_IMPORTED = False

# Set True in main() only if TkinterDnD.Tk() initializes (matches this Tk build).
HAS_DND = False
DND_INIT_ERROR = ""

BG            = "#eef2f6"
PANEL_BG      = "#fbfcfd"
PANEL_ALT_BG  = "#f6f8fa"
BORDER        = "#cfd7e2"
DROP_BG_IMPORT = "#e8f1ff"
DROP_BG_EXPORT = "#eaf5ee"
DROP_BD_IMPORT = "#6e9bd4"
DROP_BD_EXPORT = "#78a887"
LIST_BG       = "#ffffff"
STAT_BG       = "#fff9df"
ACCENT        = "#1764c8"
BTN_GRAY      = "#b8c0ca"
SECONDARY_BTN_BG = "#eef1f5"
FG            = "#1d2430"
FG_DIM        = "#657080"
FG_MUTED      = "#87909c"


class CustomsDocGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} {APP_VERSION}")
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
        outer = tk.Frame(self.root, bg=BG, padx=20, pady=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header = tk.Frame(outer, bg=BG)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        tk.Label(
            header,
            text=f"{APP_NAME} {APP_VERSION}",
            font=("Helvetica", 19, "bold"),
            fg=FG, bg=BG, anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="导入进口或出口 PDF，确认输出目录后生成。EUR Invoice PDF 在右下角单独处理。",
            font=("Helvetica", 11),
            fg=FG_DIM, bg=BG, anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        dual = tk.Frame(outer, bg=BG)
        dual.grid(row=1, column=0, sticky="nsew")
        dual.columnconfigure(0, weight=1)
        dual.columnconfigure(1, weight=1)
        dual.rowconfigure(0, weight=1)

        def file_panel(parent, title, subtitle, drop_bg, drop_bd, main_text,
                       choose_text, choose_cmd, remove_cmd, clear_cmd):
            shell = tk.Frame(parent, bg=BORDER, padx=1, pady=1)
            inner = tk.Frame(shell, bg=PANEL_BG, padx=12, pady=11)
            inner.pack(fill="both", expand=True)
            inner.columnconfigure(0, weight=1)
            inner.rowconfigure(4, weight=1)

            tk.Label(
                inner,
                text=title,
                font=("Helvetica", 12, "bold"),
                fg=FG, bg=PANEL_BG, anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            tk.Label(
                inner,
                text=subtitle,
                font=("Helvetica", 10),
                fg=FG_DIM, bg=PANEL_BG, anchor="w",
            ).grid(row=1, column=0, sticky="ew", pady=(2, 8))

            drop_zone = tk.Canvas(
                inner, height=82, bg=drop_bg,
                highlightthickness=1, highlightbackground=drop_bd,
                cursor="hand2",
            )
            drop_zone.grid(row=2, column=0, sticky="ew", pady=(0, 8))
            self._bind_drop_redraw(drop_zone, main_text)

            tools = tk.Frame(inner, bg=PANEL_BG)
            tools.grid(row=3, column=0, sticky="ew", pady=(0, 8))
            tk.Button(tools, text=choose_text, command=choose_cmd).pack(side="left")
            tk.Button(tools, text="移除选中", command=remove_cmd).pack(
                side="left", padx=(8, 0)
            )
            tk.Button(tools, text="清空列表", command=clear_cmd).pack(
                side="left", padx=(8, 0)
            )

            list_wrap = tk.Frame(inner, bg=BORDER, padx=1, pady=1)
            list_wrap.grid(row=4, column=0, sticky="nsew")
            list_wrap.columnconfigure(0, weight=1)
            list_wrap.rowconfigure(0, weight=1)
            listbox = tk.Listbox(
                list_wrap, selectmode="extended", height=7,
                font=("Helvetica", 11),
                bg=LIST_BG, fg=FG,
                selectbackground=ACCENT, selectforeground="white",
                bd=0, relief="flat",
            )
            listbox.grid(row=0, column=0, sticky="nsew")
            scroll = tk.Scrollbar(list_wrap, orient="vertical", command=listbox.yview)
            scroll.grid(row=0, column=1, sticky="ns")
            listbox.config(yscrollcommand=scroll.set)
            return shell, drop_zone, listbox

        left, self.drop_zone_import, self.file_listbox_import = file_panel(
            dual,
            "进口资料",
            "Invoice PDF，可拖入单个文件或文件夹。",
            DROP_BG_IMPORT,
            DROP_BD_IMPORT,
            "拖入进口 PDF / 文件夹",
            "添加进口 PDF...",
            self._choose_import_files,
            self._remove_selected_import,
            self._clear_import,
        )
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        right, self.drop_zone_export, self.file_listbox_export = file_panel(
            dual,
            "出口资料",
            "CustInvc 发票与 ItemShip 装箱单 PDF 放在这里。",
            DROP_BG_EXPORT,
            DROP_BD_EXPORT,
            "拖入出口 PDF / 文件夹",
            "添加出口 PDF...",
            self._choose_export_files,
            self._remove_selected_export,
            self._clear_export,
        )
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self.file_listbox_import.bind(
            "<Double-Button-1>", lambda e: self._remove_selected_import()
        )
        self.file_listbox_export.bind(
            "<Double-Button-1>", lambda e: self._remove_selected_export()
        )

        controls = tk.Frame(outer, bg=BG)
        controls.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        controls.columnconfigure(0, weight=3)
        controls.columnconfigure(1, weight=2)

        main_shell = tk.Frame(controls, bg=BORDER, padx=1, pady=1)
        main_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        main_f = tk.Frame(main_shell, bg=PANEL_BG, padx=14, pady=12)
        main_f.pack(fill="both", expand=True)
        main_f.columnconfigure(1, weight=1)

        tk.Label(
            main_f, text="主流程", fg=FG, bg=PANEL_BG,
            font=("Helvetica", 12, "bold"), anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="ew")
        tk.Label(
            main_f, text="导入资料后，从这里生成整批报关文件。",
            fg=FG_DIM, bg=PANEL_BG, font=("Helvetica", 10), anchor="w",
        ).grid(row=1, column=0, columnspan=3, sticky="ew", pady=(2, 10))

        tk.Label(
            main_f, text="输出目录", fg=FG, bg=PANEL_BG,
            font=("Helvetica", 11), anchor="w",
        ).grid(row=2, column=0, sticky="w")
        self.output_dir_entry = tk.Entry(
            main_f,
            textvariable=self.output_dir_var,
            fg=FG,
            bg="white",
            font=("Helvetica", 11),
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        self.output_dir_entry.grid(row=2, column=1, sticky="ew", padx=(10, 8))
        tk.Button(
            main_f, text="选择目录...",
            command=self._choose_output_dir,
            width=12,
        ).grid(row=2, column=2, sticky="e")

        option_row = tk.Frame(main_f, bg=PANEL_BG)
        option_row.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 12))
        tk.Checkbutton(
            option_row, text="自动 OCR（扫描件推荐开启，仅进口解析）",
            variable=self.ocr_enabled_var,
            bg=PANEL_BG, fg=FG, activebackground=PANEL_BG,
            font=("Helvetica", 11),
        ).pack(side="left")
        tk.Label(
            option_row, text="OCR 语言", fg=FG, bg=PANEL_BG,
            font=("Helvetica", 11),
        ).pack(side="left", padx=(18, 6))
        tk.Entry(
            option_row, textvariable=self.ocr_lang_var, width=10,
            fg=FG, bg="white", font=("Helvetica", 11),
            relief="solid", bd=1,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        ).pack(side="left")
        tk.Label(
            option_row, text="eng / eng+ita",
            fg=FG_MUTED, bg=PANEL_BG, font=("Helvetica", 10),
        ).pack(side="left", padx=(8, 0))

        action_row = tk.Frame(main_f, bg=PANEL_BG)
        action_row.grid(row=4, column=0, columnspan=3, sticky="ew")
        self._run_btn_enabled = False
        self.run_btn = tk.Label(
            action_row, text="开始生成报关资料",
            font=("Helvetica", 13, "bold"),
            fg="#e9edf3", bg=BTN_GRAY,
            padx=28, pady=12,
            cursor="arrow",
            relief="raised", bd=1,
        )
        self.run_btn.pack(side="left")
        self.run_btn.bind("<Button-1>", self._on_run_btn_click)
        self.run_btn.bind(
            "<ButtonRelease-1>",
            lambda e: self.run_btn.config(relief="raised"),
        )
        tk.Button(
            action_row, text="在 Finder 中查看输出",
            command=self._open_output_dir,
            width=18,
        ).pack(side="left", padx=(12, 0))

        eur_shell = tk.Frame(controls, bg=BORDER, padx=1, pady=1)
        eur_shell.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        eur_f = tk.Frame(eur_shell, bg=PANEL_ALT_BG, padx=14, pady=12)
        eur_f.pack(fill="both", expand=True)
        eur_f.columnconfigure(1, weight=1)

        tk.Label(
            eur_f, text="EUR Invoice PDF（单独功能）",
            fg=FG, bg=PANEL_ALT_BG,
            font=("Helvetica", 12, "bold"), anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="ew")
        tk.Label(
            eur_f, text="选择一张 CustInvc 发票 PDF，按汇率单独生成欧元版。",
            fg=FG_DIM, bg=PANEL_ALT_BG, font=("Helvetica", 10), anchor="w",
        ).grid(row=1, column=0, columnspan=3, sticky="ew", pady=(2, 10))

        tk.Label(
            eur_f, text="出口汇率", fg=FG, bg=PANEL_ALT_BG,
            font=("Helvetica", 11),
        ).grid(row=2, column=0, sticky="w")
        fx_row = tk.Frame(eur_f, bg=PANEL_ALT_BG)
        fx_row.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(10, 0))
        tk.Label(
            fx_row, text="1 EUR =", fg=FG, bg=PANEL_ALT_BG,
            font=("Helvetica", 11),
        ).pack(side="left")
        # 保存引用：焦点仍在输入框时点「开始生成」时，StringVar 可能尚未同步，需用 Entry.get()。
        self.fx_entry = tk.Entry(
            fx_row,
            textvariable=self.export_fx_var,
            width=12,
            fg=FG,
            bg="white",
            font=("Helvetica", 11),
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        self.fx_entry.pack(side="left", padx=(6, 6))
        tk.Label(
            fx_row,
            text="发票货币",
            fg=FG_DIM,
            bg=PANEL_ALT_BG,
            font=("Helvetica", 10),
        ).pack(side="left")

        tk.Label(
            eur_f, text="发票 PDF", fg=FG, bg=PANEL_ALT_BG,
            font=("Helvetica", 11),
        ).grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.eur_standalone_entry = tk.Entry(
            eur_f,
            textvariable=self.eur_standalone_pdf_var,
            fg=FG,
            bg="white",
            insertbackground=FG,
            font=("Helvetica", 10),
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        self.eur_standalone_entry.grid(
            row=3, column=1, sticky="ew", padx=(10, 8), pady=(10, 0)
        )
        tk.Button(
            eur_f,
            text="选择...",
            command=self._choose_eur_standalone_pdf,
            width=8,
        ).grid(row=3, column=2, sticky="e", pady=(10, 0))

        self._eur_invoice_pdf_cb = tk.Checkbutton(
            eur_f,
            text="出口批量生成时也生成 EUR Invoice PDF",
            command=self._toggle_eur_pdf,
            bg=PANEL_ALT_BG,
            fg=FG,
            activebackground=PANEL_ALT_BG,
            font=("Helvetica", 10),
        )
        self._eur_invoice_pdf_cb.grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(10, 8)
        )

        self.eur_only_btn = tk.Label(
            eur_f,
            text="单独生成 EUR Invoice PDF",
            font=("Helvetica", 11, "bold"),
            fg=FG,
            bg=SECONDARY_BTN_BG,
            padx=18,
            pady=9,
            cursor="hand2",
            relief="raised",
            bd=1,
        )
        self.eur_only_btn.grid(row=5, column=0, columnspan=3, sticky="w")
        self.eur_only_btn.bind("<Button-1>", self._on_eur_only_btn_click)
        self.eur_only_btn.bind(
            "<ButtonRelease-1>",
            lambda e: self.eur_only_btn.config(relief="raised"),
        )

        stat_f = tk.Frame(outer, bg=BORDER, padx=1, pady=1)
        stat_f.grid(row=3, column=0, sticky="ew", pady=(12, 0))
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
            detail = f"；原因：{DND_INIT_ERROR}" if DND_INIT_ERROR else ""
            self._set_status(
                "请点击「添加…文件」导入 PDF。"
                f"（已安装 tkinterdnd2，但 tkdnd 与当前 Tk 不兼容，拖拽不可用{detail}）")
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
                hint, color = "或点击「添加…PDF」", "#555555"
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
            self.run_btn.config(bg=BTN_GRAY, fg="#e9edf3", cursor="arrow")

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
            system = platform.system()
            if system == "Darwin":
                subprocess.run(["open", str(out_dir)], check=True)
            elif system == "Windows":
                os.startfile(str(out_dir))  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", str(out_dir)], check=True)
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
        if getattr(self, "output_dir_entry", None) is not None:
            out = self.output_dir_entry.get().strip()
        else:
            out = self.output_dir_var.get().strip()
        out = out or self._default_output_dir
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
            bg=BTN_GRAY, fg="#e9edf3", cursor="arrow", relief="raised")
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
            bg=SECONDARY_BTN_BG, fg=FG, cursor="hand2", relief="raised")

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
        self.run_btn.config(bg=BTN_GRAY, fg="#e9edf3", cursor="arrow")
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
    global HAS_DND, DND_INIT_ERROR
    root = None
    if TKINTERDND2_IMPORTED:
        try:
            root = TkinterDnD.Tk()
            HAS_DND = True
        except Exception as exc:
            HAS_DND = False
            DND_INIT_ERROR = str(exc)
            root = None
    if root is None:
        root = tk.Tk()
    CustomsDocGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
