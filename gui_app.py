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

APP_VERSION = "v4.0"

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TKINTERDND2_IMPORTED = True
except Exception:
    TkinterDnD = None
    DND_FILES = None
    TKINTERDND2_IMPORTED = False

# Set True in main() only if TkinterDnD.Tk() initializes (matches this Tk build).
HAS_DND = False

BG        = "#f0f4f8"
DROP_BG   = "#d4e6ff"
DROP_BD   = "#5599dd"
LIST_BG   = "#ffffff"
STAT_BG   = "#fffde6"
ACCENT    = "#2060b0"
BTN_GRAY  = "#aaaaaa"
FG        = "#1a1a1a"
FG_DIM    = "#666666"


class CustomsDocGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"报关资料生成器 {APP_VERSION}")
        self.root.geometry("960x740")
        self.root.minsize(860, 620)
        self.root.configure(bg=BG)

        self.base_dir = Path(__file__).resolve().parent
        self.invoice_paths: list[str] = []
        self.invoice_display: list[str] = []
        self.last_output_folder = None
        self.output_dir_var = tk.StringVar(value=str(self.base_dir / "output"))
        self.ocr_enabled_var = tk.BooleanVar(value=True)
        self.ocr_lang_var = tk.StringVar(value="eng")
        self.status_var = tk.StringVar(
            value="请先添加 PDF，再点击「开始生成」。")

        self._build_ui()

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=BG, padx=18, pady=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(5, weight=1)

        row = 0

        # ── 标题 ──────────────────────────────────────────────
        tk.Label(
            outer,
            text=f"Invoice 自动生成报关资料   {APP_VERSION}",
            font=("Helvetica", 17, "bold"),
            fg=FG, bg=BG, anchor="w",
        ).grid(row=row, column=0, sticky="ew", pady=(0, 10))
        row += 1

        # ── 拖拽区标题 ────────────────────────────────────────
        tk.Label(
            outer,
            text="▼  拖拽区域  ——  将 PDF / 文件夹拖到下方框中",
            font=("Helvetica", 11, "bold"),
            fg=FG, bg=BG, anchor="w",
        ).grid(row=row, column=0, sticky="ew", pady=(0, 3))
        row += 1

        # ── 拖拽区域（Canvas）─────────────────────────────────
        drop = tk.Canvas(
            outer, height=100, bg=DROP_BG,
            highlightthickness=2, highlightbackground=DROP_BD,
            cursor="hand2",
        )
        drop.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        self.drop_zone = drop

        def _redraw(event=None):
            drop.delete("all")
            w = max(drop.winfo_width(), 400)
            drop.create_text(
                w // 2, 36,
                text="将 PDF 文件 / 文件夹拖拽到此区域",
                font=("Helvetica", 14, "bold"), fill="#1a4080",
            )
            if HAS_DND:
                hint, color = "也可使用下方「添加文件」按钮", "#555555"
            elif TKINTERDND2_IMPORTED:
                hint, color = (
                    "⚠ 拖拽不可用 — tkdnd 与当前 Tk 不兼容，请用「添加文件」",
                    "#cc0000",
                )
            else:
                hint, color = ("⚠ 拖拽不可用 — 请安装 tkinterdnd2 后重启",
                               "#cc0000")
            drop.create_text(
                w // 2, 66,
                text=hint, font=("Helvetica", 10), fill=color,
            )

        drop.bind("<Configure>", _redraw)
        self.root.after(80, _redraw)
        row += 1

        # ── 文件列表标题 ──────────────────────────────────────
        tk.Label(
            outer,
            text="▼  已添加文件列表（双击行可删除）",
            font=("Helvetica", 11, "bold"),
            fg=FG, bg=BG, anchor="w",
        ).grid(row=row, column=0, sticky="ew", pady=(0, 3))
        row += 1

        # ── 按钮行 ────────────────────────────────────────────
        btn_f = tk.Frame(outer, bg=BG)
        btn_f.grid(row=row, column=0, sticky="ew", pady=(0, 4))
        row += 1
        tk.Button(btn_f, text="添加文件...",
                  command=self._choose_invoice_files).pack(side="left")
        tk.Button(btn_f, text="移除选中",
                  command=self._remove_selected).pack(
            side="left", padx=(8, 0))
        tk.Button(btn_f, text="清空列表",
                  command=self._clear_files).pack(
            side="left", padx=(8, 0))
        tk.Label(btn_f, text="（仅显示文件名）",
                 fg=FG_DIM, bg=BG,
                 font=("Helvetica", 10)).pack(side="left", padx=(14, 0))

        # ── 文件列表 ──────────────────────────────────────────
        list_f = tk.Frame(outer, bg=BG)
        list_f.grid(row=row, column=0, sticky="nsew", pady=(0, 8))
        row += 1
        list_f.columnconfigure(0, weight=1)
        list_f.rowconfigure(0, weight=1)
        self.file_listbox = tk.Listbox(
            list_f, selectmode="extended", height=7,
            font=("Helvetica", 12),
            bg=LIST_BG, fg=FG,
            selectbackground=ACCENT, selectforeground="white",
            bd=1, relief="solid",
        )
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        vsb = tk.Scrollbar(list_f, orient="vertical",
                           command=self.file_listbox.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.file_listbox.config(yscrollcommand=vsb.set)
        self.file_listbox.bind("<Double-Button-1>",
                               lambda e: self._remove_selected())

        # ── 输出目录 ──────────────────────────────────────────
        out_f = tk.Frame(outer, bg=BG)
        out_f.grid(row=row, column=0, sticky="ew", pady=4)
        row += 1
        tk.Label(out_f, text="输出目录:", fg=FG, bg=BG,
                 font=("Helvetica", 11), anchor="w").pack(side="left")
        tk.Entry(out_f, textvariable=self.output_dir_var,
                 fg=FG, bg="white",
                 font=("Helvetica", 11)).pack(
            side="left", fill="x", expand=True, padx=(6, 8))
        tk.Button(out_f, text="选择目录...",
                  command=self._choose_output_dir).pack(side="left")

        # ── OCR 设置 ──────────────────────────────────────────
        ocr_f = tk.Frame(outer, bg=BG)
        ocr_f.grid(row=row, column=0, sticky="ew", pady=4)
        row += 1
        tk.Checkbutton(ocr_f, text="自动 OCR（扫描件推荐开启）",
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
            for w in (self.drop_zone, self.file_listbox):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop_files)
                except tk.TclError:
                    pass
            self._set_status(
                "拖拽已启用。可将 PDF 拖入蓝色区域，或点击「添加文件」。")
        elif TKINTERDND2_IMPORTED:
            self._set_status(
                "请点击「添加文件」导入 PDF。"
                "（已安装 tkinterdnd2，但 tkdnd 与当前 Tk 不兼容，拖拽不可用）")
        else:
            self._set_status(
                "请点击「添加文件」导入 PDF。"
                "（安装 tkinterdnd2 可启用拖拽）")

    # ------------------------------------------------------------------
    def _update_run_btn(self):
        """Gray when not ready, blue when ready."""
        has_files = bool(self.invoice_paths)
        has_output = bool(self.output_dir_var.get().strip())
        if has_files and has_output:
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

    def _choose_invoice_files(self):
        paths = filedialog.askopenfilenames(
            title="选择一个或多个 Invoice PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        self._add_files(paths)

    def _choose_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir_var.set(path)
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

    def _add_files(self, paths):
        if not paths:
            return
        for path in self._expand_paths(paths):
            if path not in self.invoice_paths:
                self.invoice_paths.append(path)
                display = Path(path).name
                self.invoice_display.append(display)
                self.file_listbox.insert("end", display)
        self._update_run_btn()

    def _remove_selected(self):
        selected = list(self.file_listbox.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            self.file_listbox.delete(idx)
            self.invoice_paths.pop(idx)
            self.invoice_display.pop(idx)
        self._update_run_btn()

    def _clear_files(self):
        self.file_listbox.delete(0, "end")
        self.invoice_paths.clear()
        self.invoice_display.clear()
        self._update_run_btn()

    def _on_drop_files(self, event):
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
        self._add_files(paths)

    def _open_output_dir(self):
        out_dir = (Path(self.last_output_folder) if self.last_output_folder
                   else Path(self.output_dir_var.get()).expanduser())
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.run(["open", str(out_dir)], check=True)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开输出目录:\n{e}")

    def _start_generate(self):
        output_dir = self.output_dir_var.get().strip()
        if not self.invoice_paths:
            messagebox.showwarning("提示", "请先添加至少一个 Invoice PDF 文件")
            return
        missing = [p for p in self.invoice_paths if not Path(p).exists()]
        if missing:
            messagebox.showerror("错误",
                                 f"有文件不存在，请检查:\n{missing[0]}")
            return
        if not output_dir:
            messagebox.showwarning("提示", "请先选择输出目录")
            return

        enable_ocr = self.ocr_enabled_var.get()
        ocr_lang = self.ocr_lang_var.get().strip() or "eng"
        invoice_paths = list(self.invoice_paths)

        self._run_btn_enabled = False
        self.run_btn.config(bg=BTN_GRAY, fg="#dddddd", cursor="arrow")
        self._set_status(
            f"正在处理 {len(invoice_paths)} 份 PDF，请稍候...")
        threading.Thread(
            target=self._generate,
            args=(output_dir, enable_ocr, ocr_lang, invoice_paths),
            daemon=True,
        ).start()

    def _generate(self, output_dir, enable_ocr, ocr_lang, invoice_paths):

        try:
            all_files = []
            success_count = 0
            from datetime import datetime
            batch_folder = (
                Path(output_dir).expanduser()
                / f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
            batch_folder.mkdir(parents=True, exist_ok=True)
            self.last_output_folder = str(batch_folder)

            for i, invoice_path in enumerate(invoice_paths, start=1):
                self.root.after(
                    0,
                    lambda i=i, n=len(invoice_paths):
                        self._set_status(f"处理中 {i}/{n} ..."),
                )
                raw_stem = Path(invoice_path).stem.strip() or f"invoice_{i}"
                invoice_stem = re.sub(r'[\\/:*?"<>|]+', "_", raw_stem)
                invoice_out = batch_folder / f"{i:02d}_{invoice_stem}"
                invoice_out.mkdir(parents=True, exist_ok=True)

                generator = CustomsDocGenerator(
                    config_path=str(
                        self.base_dir / "data"
                        / "supplier_product_mapping.yaml"
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
                    success_count += 1
                    all_files.extend(files)

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
                    f"批量处理完成: {success_count}/{len(invoice_paths)} 成功\n"
                    f"输出目录: {self.last_output_folder}\n\n"
                    f"已生成文件:\n{preview}"
                )
                self.root.after(
                    0, lambda: messagebox.showinfo("生成成功", msg))
                self.root.after(
                    0,
                    lambda: self._set_status(
                        f"处理完成：成功 {success_count}/{len(invoice_paths)}，"
                        f"共生成 {len(all_files)} 个文件。"),
                )
                self.root.after(0, self._open_output_dir)
        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror("处理失败", f"生成失败:\n{e}"),
            )
            self.root.after(
                0, lambda: self._set_status(f"处理失败: {e}"))
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
