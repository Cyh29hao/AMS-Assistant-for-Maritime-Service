from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, END, LEFT, RIGHT, X, Y

from desktop_app.excel_sync_engine import SyncService, SyncTask, list_headers, list_sheets


SOURCE_MODE_LABELS = {
    "整张表": "whole_sheet",
    "自定义区域": "custom_range",
}
SOURCE_MODE_VALUES = {value: label for label, value in SOURCE_MODE_LABELS.items()}

TARGET_MODE_LABELS = {
    "覆盖目标表": "replace_sheet",
    "从指定单元格开始": "write_from_cell",
}
TARGET_MODE_VALUES = {value: label for label, value in TARGET_MODE_LABELS.items()}

COLUMN_MODE_LABELS = {
    "只保留勾选列（兼容旧任务）": "include",
    "复制整表，仅排除勾选列（推荐）": "exclude",
}
COLUMN_MODE_VALUES = {value: label for label, value in COLUMN_MODE_LABELS.items()}


class ExcelSyncPanel:
    def __init__(
        self,
        parent: ttk.Frame,
        window: ttk.Window,
        service: SyncService,
        open_path,
        open_help,
    ) -> None:
        self.parent = parent
        self.window = window
        self.service = service
        self.service.status_callback = self._schedule_refresh
        self.open_path = open_path
        self.open_help = open_help

        self.pending_refresh = False
        self.selected_task_id: str | None = None
        self.header_vars: dict[str, tk.BooleanVar] = {}
        self.columns_window = None

        self._build_vars()
        self._build_ui()
        self.service.start()
        self._refresh()
        self.window.after(1000, self._poll_ui)

    def _build_vars(self) -> None:
        self.name_var = tk.StringVar()
        self.enabled_var = tk.BooleanVar(value=True)
        self.source_file_var = tk.StringVar()
        self.source_sheet_var = tk.StringVar()
        self.source_mode_var = tk.StringVar(value="整张表")
        self.source_range_var = tk.StringVar()
        self.target_file_var = tk.StringVar()
        self.target_sheet_var = tk.StringVar()
        self.target_mode_var = tk.StringVar(value="覆盖目标表")
        self.target_start_cell_var = tk.StringVar(value="A1")
        self.column_mode_var = tk.StringVar(value="复制整表，仅排除勾选列（推荐）")
        self.header_row_var = tk.StringVar(value="1")
        self.data_start_row_var = tk.StringVar(value="2")
        self.formula_var = tk.StringVar(value="values")

    def _build_ui(self) -> None:
        intro = ttk.Frame(self.parent, padding=16, bootstyle="light", borderwidth=1, relief="solid")
        intro.pack(fill=X, pady=(0, 12))
        ttk.Label(intro, text="同步任务说明", font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w")
        ttk.Label(
            intro,
            text="把一个 Excel 的指定工作表同步到另一个 Excel。现在默认更适合“几乎整张表复制过去，只排除少数几列”的场景，同时会尽量保留隐藏行、空行和表格结构。",
            bootstyle="secondary",
            wraplength=980,
        ).pack(anchor="w", pady=(6, 0))
        intro_buttons = ttk.Frame(intro)
        intro_buttons.pack(fill=X, pady=(10, 0))
        ttk.Button(intro_buttons, text="打开功能说明", bootstyle="info", command=lambda: self.open_help("sync")).pack(side=LEFT, padx=4)
        ttk.Button(intro_buttons, text="打开示例文件夹", bootstyle="secondary", command=self._open_examples).pack(side=LEFT, padx=4)
        ttk.Button(intro_buttons, text="打开任务配置文件", bootstyle="secondary", command=self._open_task_file).pack(side=LEFT, padx=4)
        ttk.Button(intro_buttons, text="打开运行日志", bootstyle="secondary", command=self._open_log).pack(side=LEFT, padx=4)

        top = ttk.Frame(self.parent)
        top.pack(fill=BOTH, expand=True)
        top.columnconfigure(0, weight=11)
        top.columnconfigure(1, weight=16)
        top.rowconfigure(0, weight=1)

        left = ttk.Frame(top, padding=12, bootstyle="light", borderwidth=1, relief="solid")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        ttk.Label(left, text="同步任务", font=("Microsoft YaHei UI", 13, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(left, text="左边看任务列表，右边改规则和字段。", bootstyle="secondary").grid(row=1, column=0, sticky="w", pady=(4, 10))

        toolbar = ttk.Frame(left)
        toolbar.grid(row=0, column=1, sticky="e", pady=(0, 10))
        self.monitor_button = ttk.Button(toolbar, text="暂停监控", bootstyle="warning", command=self._toggle_monitoring)
        self.monitor_button.pack(side=LEFT)
        ttk.Button(toolbar, text="新建任务", bootstyle="success", command=self._new_task).pack(side=LEFT, padx=(8, 0))

        cols = ("enabled", "name", "status", "last_startsync", "last_synced")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", height=18)
        for key, title, width in (
            ("enabled", "启用", 70),
            ("name", "任务名", 180),
            ("status", "状态", 220),
            ("last_startsync", "上次手动同步", 160),
            ("last_synced", "上次完成", 160),
        ):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor="center" if key == "enabled" else "w")
        tree_scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.grid(row=2, column=0, sticky="nsew")
        tree_scroll.grid(row=2, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.status_label = ttk.Label(left, text="监控中", bootstyle="secondary")
        self.status_label.grid(row=3, column=0, sticky="w", pady=(10, 0))

        right = ttk.Frame(top, padding=12, bootstyle="light", borderwidth=1, relief="solid")
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)
        ttk.Label(right, text="任务编辑器", font=("Microsoft YaHei UI", 13, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(right, text="这里决定源表、目标表、列和同步方式。", bootstyle="secondary").grid(row=0, column=1, sticky="e")
        row = 1

        ttk.Label(right, text="任务名称").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(right, textvariable=self.name_var).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Checkbutton(right, text="启用这个任务", variable=self.enabled_var, bootstyle="round-toggle").grid(
            row=row, column=0, columnspan=2, sticky="w", pady=4
        )
        row += 1

        ttk.Label(right, text="源 Excel").grid(row=row, column=0, sticky="w", pady=4)
        source_row = ttk.Frame(right)
        source_row.grid(row=row, column=1, sticky="ew", pady=4)
        source_row.columnconfigure(0, weight=1)
        ttk.Entry(source_row, textvariable=self.source_file_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(source_row, text="选择文件", bootstyle="secondary", command=self._pick_source).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(source_row, text="加载工作表", bootstyle="info", command=self._load_sheets).grid(row=0, column=2, padx=(8, 0))
        row += 1

        ttk.Label(right, text="源工作表").grid(row=row, column=0, sticky="w", pady=4)
        self.sheet_combo = ttk.Combobox(right, textvariable=self.source_sheet_var, state="readonly")
        self.sheet_combo.grid(row=row, column=1, sticky="ew", pady=4)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_headers())
        row += 1

        ttk.Label(right, text="源区域模式").grid(row=row, column=0, sticky="w", pady=4)
        self.source_mode_combo = ttk.Combobox(
            right,
            textvariable=self.source_mode_var,
            state="readonly",
            values=list(SOURCE_MODE_LABELS.keys()),
        )
        self.source_mode_combo.grid(row=row, column=1, sticky="ew", pady=4)
        self.source_mode_combo.bind("<<ComboboxSelected>>", self._on_source_mode_change)
        row += 1

        self.source_range_label = ttk.Label(right, text="源区域")
        self.source_range_label.grid(row=row, column=0, sticky="w", pady=4)
        self.source_range_entry = ttk.Entry(right, textvariable=self.source_range_var)
        self.source_range_entry.grid(row=row, column=1, sticky="ew", pady=4)
        self.source_range_entry.bind("<FocusOut>", lambda _e: self._load_headers())
        self.source_range_entry.bind("<Return>", lambda _e: self._load_headers())
        row += 1

        ttk.Label(right, text="目标 Excel").grid(row=row, column=0, sticky="w", pady=4)
        target_row = ttk.Frame(right)
        target_row.grid(row=row, column=1, sticky="ew", pady=4)
        target_row.columnconfigure(0, weight=1)
        ttk.Entry(target_row, textvariable=self.target_file_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(target_row, text="选择文件", bootstyle="secondary", command=self._pick_target).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(target_row, text="读取目标表", bootstyle="info", command=self._load_target_sheets).grid(row=0, column=2, padx=(8, 0))
        row += 1

        ttk.Label(right, text="目标工作表").grid(row=row, column=0, sticky="w", pady=4)
        self.target_sheet_combo = ttk.Combobox(right, textvariable=self.target_sheet_var, state="normal")
        self.target_sheet_combo.grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(right, text="写入方式").grid(row=row, column=0, sticky="w", pady=4)
        self.target_mode_combo = ttk.Combobox(
            right,
            textvariable=self.target_mode_var,
            state="readonly",
            values=list(TARGET_MODE_LABELS.keys()),
        )
        self.target_mode_combo.grid(row=row, column=1, sticky="ew", pady=4)
        self.target_mode_combo.bind("<<ComboboxSelected>>", self._on_target_mode_change)
        row += 1

        self.target_start_label = ttk.Label(right, text="起始单元格")
        self.target_start_label.grid(row=row, column=0, sticky="w", pady=4)
        self.target_start_entry = ttk.Entry(right, textvariable=self.target_start_cell_var)
        self.target_start_entry.grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        advanced = ttk.Labelframe(right, text="高级选项", padding=8, bootstyle="secondary")
        advanced.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(8, 10))
        ttk.Label(advanced, text="表头所在行").grid(row=0, column=0, sticky="w", pady=2)
        self.header_row_entry = ttk.Entry(advanced, textvariable=self.header_row_var, width=10)
        self.header_row_entry.grid(row=0, column=1, sticky="w", pady=2)
        self.header_row_entry.bind("<FocusOut>", lambda _e: self._load_headers())
        self.header_row_entry.bind("<Return>", lambda _e: self._load_headers())
        ttk.Label(advanced, text="列处理方式").grid(row=1, column=0, sticky="w", pady=2)
        self.column_mode_combo = ttk.Combobox(
            advanced,
            textvariable=self.column_mode_var,
            state="readonly",
            values=list(COLUMN_MODE_LABELS.keys()),
            width=28,
        )
        self.column_mode_combo.grid(row=1, column=1, sticky="w", pady=2)
        self.column_mode_combo.bind("<<ComboboxSelected>>", self._on_column_mode_change)
        ttk.Label(advanced, text="数据起始行").grid(row=2, column=0, sticky="w", pady=2)
        self.data_start_row_entry = ttk.Entry(advanced, textvariable=self.data_start_row_var, width=10)
        self.data_start_row_entry.grid(row=2, column=1, sticky="w", pady=2)
        ttk.Label(advanced, text="公式输出").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Combobox(
            advanced,
            textvariable=self.formula_var,
            state="readonly",
            values=["values", "formulas"],
            width=12,
        ).grid(row=3, column=1, sticky="w", pady=2)
        row += 1

        self.columns_label = ttk.Label(right, text="要排除的列")
        self.columns_label.grid(row=row, column=0, sticky="nw", pady=4)
        columns_frame = ttk.Frame(right)
        columns_frame.grid(row=row, column=1, sticky="nsew")
        right.rowconfigure(row, weight=1)
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.rowconfigure(1, weight=1)

        selection_actions = ttk.Frame(columns_frame)
        selection_actions.grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Button(selection_actions, text="全选", bootstyle="secondary", command=self._select_all_headers).pack(side=LEFT)
        ttk.Button(selection_actions, text="清空", bootstyle="secondary", command=self._clear_headers).pack(side=LEFT, padx=(8, 0))
        ttk.Button(selection_actions, text="反选", bootstyle="secondary", command=self._invert_headers).pack(side=LEFT, padx=(8, 0))

        self.columns_canvas = tk.Canvas(columns_frame, height=240, highlightthickness=0)
        cols_scroll = ttk.Scrollbar(columns_frame, orient="vertical", command=self.columns_canvas.yview)
        self.columns_canvas.configure(yscrollcommand=cols_scroll.set)
        self.columns_canvas.grid(row=1, column=0, sticky="nsew")
        cols_scroll.grid(row=1, column=1, sticky="ns")
        background = self.columns_canvas.cget("background")
        self.columns_inner = tk.Frame(self.columns_canvas, bg=background)
        self.columns_inner.bind(
            "<Configure>",
            lambda _e: self.columns_canvas.configure(scrollregion=self.columns_canvas.bbox("all")),
        )
        self.columns_window = self.columns_canvas.create_window((0, 0), window=self.columns_inner, anchor="nw")
        self.columns_canvas.bind("<Configure>", self._resize_columns_area)
        row += 1

        actions = ttk.Frame(right)
        actions.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(actions, text="保存任务", bootstyle="success", command=self._save_task).pack(side=LEFT)
        ttk.Button(actions, text="保存并立即同步", bootstyle="primary", command=self._run_now).pack(side=LEFT, padx=(8, 0))
        ttk.Button(actions, text="复制任务", bootstyle="secondary", command=self._copy_task).pack(side=LEFT, padx=(8, 0))
        ttk.Button(actions, text="删除任务", bootstyle="danger", command=self._delete_task).pack(side=LEFT, padx=(8, 0))

        log_box = ttk.Frame(self.parent, padding=10, bootstyle="light", borderwidth=1, relief="solid")
        log_box.pack(fill=BOTH, expand=False, pady=(12, 0))
        ttk.Label(log_box, text="最近日志", font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w")
        self.log_text = tk.Text(log_box, height=10, wrap="word", state="disabled")
        log_scroll = ttk.Scrollbar(log_box, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.configure(background="#f5f9fc", foreground="#24384b", relief="flat", borderwidth=0, highlightthickness=0)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        log_scroll.pack(side=RIGHT, fill=Y)

        self._show_columns_placeholder("先选择源文件和工作表，然后在这里勾选需要同步的列。")
        self._update_source_mode_ui()
        self._update_target_mode_ui()
        self._update_column_mode_ui()

    def shutdown(self) -> None:
        self.service.stop()

    def _schedule_refresh(self) -> None:
        self.pending_refresh = True

    def _poll_ui(self) -> None:
        if self.pending_refresh:
            self.pending_refresh = False
            self._refresh()
        self.window.after(1000, self._poll_ui)

    def _refresh(self) -> None:
        rows = self.service.list_runtime_rows()
        ids = set()
        for task, runtime in rows:
            ids.add(task.id)
            values = (
                "是" if task.enabled else "否",
                task.name,
                runtime.status,
                runtime.last_startsync_at,
                runtime.last_synced_at,
            )
            if self.tree.exists(task.id):
                self.tree.item(task.id, values=values)
            else:
                self.tree.insert("", "end", iid=task.id, values=values)
        for item_id in self.tree.get_children():
            if item_id not in ids:
                self.tree.delete(item_id)
        running = self.service.is_running()
        self.status_label.configure(text="监控中" if running else "监控已暂停")
        self.monitor_button.configure(text="暂停监控" if running else "恢复监控", bootstyle="warning" if running else "success")
        self._refresh_log()

    def _refresh_log(self) -> None:
        tail = self.service.recent_log_lines()
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.insert("1.0", tail)
        self.log_text.configure(state="disabled")

    def _update_source_mode_ui(self) -> None:
        is_custom = SOURCE_MODE_LABELS.get(self.source_mode_var.get(), "whole_sheet") == "custom_range"
        if is_custom:
            self.source_range_label.grid()
            self.source_range_entry.grid()
            self.header_row_entry.state(["disabled"])
            self.data_start_row_entry.state(["disabled"])
        else:
            self.source_range_label.grid_remove()
            self.source_range_entry.grid_remove()
            self.header_row_entry.state(["!disabled"])
            self.data_start_row_entry.state(["!disabled"])

    def _update_target_mode_ui(self) -> None:
        write_from_cell = TARGET_MODE_LABELS.get(self.target_mode_var.get(), "replace_sheet") == "write_from_cell"
        if write_from_cell:
            self.target_start_label.grid()
            self.target_start_entry.grid()
        else:
            self.target_start_label.grid_remove()
            self.target_start_entry.grid_remove()

    def _update_column_mode_ui(self) -> None:
        exclude_mode = COLUMN_MODE_LABELS.get(self.column_mode_var.get(), "exclude") == "exclude"
        self.columns_label.configure(text="要排除的列" if exclude_mode else "要保留的列")

    def _on_source_mode_change(self, _event=None):
        self._update_source_mode_ui()
        self._load_headers()

    def _on_target_mode_change(self, _event=None):
        self._update_target_mode_ui()

    def _on_column_mode_change(self, _event=None):
        self._update_column_mode_ui()

    def _pick_source(self) -> None:
        path = filedialog.askopenfilename(
            title="选择源 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.source_file_var.set(path)
            self._load_sheets()

    def _pick_target(self) -> None:
        path = filedialog.asksaveasfilename(
            title="选择目标 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.target_file_var.set(path)
            self._load_target_sheets()
            if not self.target_sheet_var.get():
                self.target_sheet_var.set("Export")

    def _load_sheets(self) -> None:
        source_file = self.source_file_var.get().strip()
        if not source_file:
            return
        if Path(source_file).name.startswith("~$"):
            messagebox.showerror("读取工作表失败", "请选择真实的 Excel 文件，不要选择 Office 临时锁文件。")
            return
        try:
            sheets = list_sheets(source_file)
        except Exception as exc:
            messagebox.showerror("读取工作表失败", str(exc))
            return
        self.sheet_combo["values"] = sheets
        if sheets and self.source_sheet_var.get() not in sheets:
            self.source_sheet_var.set(sheets[0])
        self._load_headers()

    def _load_target_sheets(self) -> None:
        target_file = self.target_file_var.get().strip()
        if not target_file:
            self.target_sheet_combo["values"] = []
            return
        target_path = Path(target_file)
        if not target_path.exists():
            self.target_sheet_combo["values"] = []
            if not self.target_sheet_var.get():
                self.target_sheet_var.set("Export")
            return
        try:
            sheets = list_sheets(target_file)
        except Exception as exc:
            messagebox.showerror("读取目标表失败", str(exc))
            return
        self.target_sheet_combo["values"] = sheets
        if sheets and not self.target_sheet_var.get():
            self.target_sheet_var.set(sheets[0])

    def _show_columns_placeholder(self, text: str) -> None:
        for child in self.columns_inner.winfo_children():
            child.destroy()
        self.header_vars = {}
        background = self.columns_inner.cget("bg")
        label = tk.Label(
            self.columns_inner,
            text=text,
            anchor="w",
            justify="left",
            bg=background,
            fg="#666666",
            padx=4,
            pady=8,
        )
        label.grid(row=0, column=0, sticky="w")

    def _load_headers(self) -> None:
        source_file = self.source_file_var.get().strip()
        source_sheet = self.source_sheet_var.get().strip()
        if not source_file or not source_sheet:
            mode_text = "排除" if COLUMN_MODE_LABELS.get(self.column_mode_var.get(), "exclude") == "exclude" else "保留"
            self._show_columns_placeholder(f"先选择源文件和工作表，然后勾选要{mode_text}的列。")
            return
        try:
            headers = list_headers(
                source_file,
                source_sheet,
                int(self.header_row_var.get() or "1"),
                source_mode=SOURCE_MODE_LABELS.get(self.source_mode_var.get(), "whole_sheet"),
                source_range=self.source_range_var.get().strip(),
            )
        except Exception as exc:
            messagebox.showerror("读取字段失败", str(exc))
            return
        current = {header for header, var in self.header_vars.items() if var.get()}
        for child in self.columns_inner.winfo_children():
            child.destroy()
        self.header_vars = {}
        if not headers:
            self._show_columns_placeholder("当前工作表或区域没有读到有效表头。")
            return
        columns_per_row = 2
        background = self.columns_inner.cget("bg")
        for col_idx in range(columns_per_row):
            self.columns_inner.grid_columnconfigure(col_idx, weight=1)
        for idx, header in enumerate(headers):
            var = tk.BooleanVar(value=header in current)
            self.header_vars[header] = var
            checkbox = tk.Checkbutton(
                self.columns_inner,
                text=header,
                variable=var,
                onvalue=True,
                offvalue=False,
                anchor="w",
                justify="left",
                bg=background,
                activebackground=background,
                selectcolor="white",
                highlightthickness=0,
                relief="flat",
                padx=4,
                pady=2,
            )
            checkbox.grid(row=idx // columns_per_row, column=idx % columns_per_row, sticky="ew", padx=(0, 12), pady=2)
        self._resize_columns_area()

    def _resize_columns_area(self, event=None) -> None:
        if self.columns_window is None:
            return
        width = event.width if event is not None else self.columns_canvas.winfo_width()
        if width > 1:
            self.columns_canvas.itemconfigure(self.columns_window, width=width)

    def _select_all_headers(self) -> None:
        for var in self.header_vars.values():
            var.set(True)

    def _clear_headers(self) -> None:
        for var in self.header_vars.values():
            var.set(False)

    def _invert_headers(self) -> None:
        for var in self.header_vars.values():
            var.set(not var.get())

    def _on_select(self, _event) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        self.selected_task_id = selection[0]
        self._load_task_into_form(self.service.get_task(self.selected_task_id))

    def _load_task_into_form(self, task: SyncTask | None) -> None:
        if task is None:
            self.selected_task_id = None
            self.name_var.set("")
            self.enabled_var.set(True)
            self.source_file_var.set("")
            self.source_sheet_var.set("")
            self.source_mode_var.set("整张表")
            self.source_range_var.set("")
            self.target_file_var.set("")
            self.target_sheet_var.set("")
            self.target_mode_var.set("覆盖目标表")
            self.target_start_cell_var.set("A1")
            self.column_mode_var.set("复制整表，仅排除勾选列（推荐）")
            self.header_row_var.set("1")
            self.data_start_row_var.set("2")
            self.formula_var.set("values")
            self.sheet_combo["values"] = []
            self.target_sheet_combo["values"] = []
            self._update_source_mode_ui()
            self._update_target_mode_ui()
            self._update_column_mode_ui()
            self._show_columns_placeholder("先选择源文件和工作表，然后在这里勾选要排除的列。")
            return

        self.name_var.set(task.name)
        self.enabled_var.set(task.enabled)
        self.source_file_var.set(task.source_file)
        self.source_sheet_var.set(task.source_sheet)
        self.source_mode_var.set(SOURCE_MODE_VALUES.get(task.source_mode, "整张表"))
        self.source_range_var.set(task.source_range)
        self.target_file_var.set(task.target_file)
        self.target_sheet_var.set(task.target_sheet)
        self.target_mode_var.set(TARGET_MODE_VALUES.get(task.target_mode, "覆盖目标表"))
        self.target_start_cell_var.set(task.target_start_cell or "A1")
        self.column_mode_var.set(COLUMN_MODE_VALUES.get(task.column_selection_mode, "只保留勾选列（兼容旧任务）"))
        self.header_row_var.set(str(task.header_row))
        self.data_start_row_var.set(str(task.data_start_row))
        self.formula_var.set(task.formula_handling)
        self._update_source_mode_ui()
        self._update_target_mode_ui()
        self._update_column_mode_ui()
        try:
            sheets = list_sheets(task.source_file) if task.source_file else []
        except Exception:
            sheets = []
        self.sheet_combo["values"] = sheets
        self._load_target_sheets()
        self._load_headers()
        for header, var in self.header_vars.items():
            var.set(header in task.columns_by_header)

    def _build_task(self) -> SyncTask:
        headers = [header for header, var in self.header_vars.items() if var.get()]
        existing = self.service.get_task(self.selected_task_id) if self.selected_task_id else None
        return SyncTask(
            id=self.selected_task_id or SyncTask().id,
            name=self.name_var.get().strip() or "新任务",
            enabled=self.enabled_var.get(),
            source_file=self.source_file_var.get().strip(),
            source_sheet=self.source_sheet_var.get().strip(),
            source_mode=SOURCE_MODE_LABELS.get(self.source_mode_var.get(), "whole_sheet"),
            source_range=self.source_range_var.get().strip(),
            target_file=self.target_file_var.get().strip(),
            target_sheet=self.target_sheet_var.get().strip() or "Export",
            target_mode=TARGET_MODE_LABELS.get(self.target_mode_var.get(), "replace_sheet"),
            target_start_cell=self.target_start_cell_var.get().strip() or "A1",
            column_selection_mode=COLUMN_MODE_LABELS.get(self.column_mode_var.get(), "exclude"),
            columns_by_header=headers,
            header_row=int(self.header_row_var.get() or "1"),
            data_start_row=int(self.data_start_row_var.get() or "2"),
            formula_handling=self.formula_var.get().strip() or "values",
            last_target_rows=existing.last_target_rows if existing else 0,
            last_target_cols=existing.last_target_cols if existing else 0,
        )

    def _persist_task(self) -> SyncTask:
        task = self._build_task()
        tasks = list(self.service.tasks)
        for idx, existing in enumerate(tasks):
            if existing.id == task.id:
                tasks[idx] = task
                break
        else:
            tasks.append(task)
        self.service.set_tasks(tasks)
        self.selected_task_id = task.id
        self._refresh()
        if self.tree.exists(task.id):
            self.tree.selection_set(task.id)
            self.tree.focus(task.id)
        self._load_task_into_form(task)
        return task

    def _save_task(self) -> None:
        try:
            self._persist_task()
        except Exception as exc:
            messagebox.showerror("保存任务失败", str(exc))

    def _copy_task(self) -> None:
        source_task = self.service.get_task(self.selected_task_id) if self.selected_task_id else None
        if source_task is None:
            messagebox.showinfo("复制任务", "请先选中一个任务。")
            return
        copied = SyncTask(
            name=f"{source_task.name}-副本",
            enabled=source_task.enabled,
            source_file=source_task.source_file,
            source_sheet=source_task.source_sheet,
            source_mode=source_task.source_mode,
            source_range=source_task.source_range,
            target_file=source_task.target_file,
            target_sheet=source_task.target_sheet,
            target_mode=source_task.target_mode,
            target_start_cell=source_task.target_start_cell,
            column_selection_mode=source_task.column_selection_mode,
            columns_by_header=list(source_task.columns_by_header),
            header_row=source_task.header_row,
            data_start_row=source_task.data_start_row,
            copy_style=source_task.copy_style,
            copy_column_widths=source_task.copy_column_widths,
            copy_row_heights=source_task.copy_row_heights,
            include_header=source_task.include_header,
            drop_empty_rows=source_task.drop_empty_rows,
            formula_handling=source_task.formula_handling,
            last_target_rows=0,
            last_target_cols=0,
        )
        tasks = list(self.service.tasks)
        tasks.append(copied)
        self.service.set_tasks(tasks)
        self.selected_task_id = copied.id
        self._refresh()
        if self.tree.exists(copied.id):
            self.tree.selection_set(copied.id)
            self.tree.focus(copied.id)
        self._load_task_into_form(copied)

    def _delete_task(self) -> None:
        if not self.selected_task_id:
            return
        if not messagebox.askyesno("删除任务", "确定要删除这个同步任务吗？"):
            return
        tasks = [task for task in self.service.tasks if task.id != self.selected_task_id]
        self.service.set_tasks(tasks)
        self._new_task()
        self._refresh()

    def _new_task(self) -> None:
        for item in self.tree.selection():
            self.tree.selection_remove(item)
        self._load_task_into_form(None)

    def _run_now(self) -> None:
        try:
            task = self._persist_task()
        except Exception as exc:
            messagebox.showerror("同步失败", str(exc))
            return
        self.service.run_task_now(task.id)
        self._refresh()

    def _toggle_monitoring(self) -> None:
        if self.service.is_running():
            self.service.stop()
        else:
            self.service.start()
        self._refresh()

    def _open_log(self) -> None:
        self.open_path(self.service.paths.log_path)

    def _open_task_file(self) -> None:
        self.open_path(self.service.paths.data_path)

    def _open_examples(self) -> None:
        examples_dir = self.service.paths.log_dir.parent / "示例文件"
        self.open_path(examples_dir)
