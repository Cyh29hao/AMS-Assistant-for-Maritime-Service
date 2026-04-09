from __future__ import annotations

import queue
import subprocess
import threading
import traceback
import webbrowser
from pathlib import Path
import tkinter.ttk as tkttk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from typing import Any, Callable

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, END, LEFT, RIGHT, X, Y

from desktop_app.excel_sync_panel import ExcelSyncPanel
from desktop_app.runtime import (
    APP_NAME,
    APP_REPO_URL,
    APP_VERSION,
    CLEARANCE_FEATURE_NAME,
    CONTRACT_FEATURE_NAME,
    LINEUP_FEATURE_NAME,
    SYNC_FEATURE_NAME,
    AmsOperations,
    AppConfig,
    ConfigStore,
    help_assets_dir,
    open_in_file_explorer,
    portable_install_root,
)


ThemeChoices = [
    "flatly",
    "litera",
    "minty",
    "sandstone",
    "journal",
]

BrowserChoices = [
    "auto",
    "msedge",
    "chrome",
    "playwright",
]


class AmsDesktopApp:
    def __init__(self) -> None:
        self.config_store = ConfigStore()
        self.config = self.config_store.load()
        self.ops = AmsOperations(self.config)
        self.ops.ensure_workspace()

        self.window = ttk.Window(themename=self.config.theme_name)
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        default_width = min(1380, max(1180, screen_width - 120))
        default_height = min(980, max(860, screen_height - 120))
        self.window.title(f"{APP_NAME} Desktop")
        self.window.geometry(f"{default_width}x{default_height}")
        self.window.minsize(1180, 820)
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        self.window.option_add("*tearOff", False)

        self.status_text = ttk.StringVar(value="准备就绪")
        self.contract_input_var = ttk.StringVar()
        self.contract_result_var = ttk.StringVar()
        self.clearance_input_var = ttk.StringVar()
        self.clearance_result_var = ttk.StringVar()
        self.clearance_query_var = ttk.StringVar()
        self.clearance_iemark_var = ttk.StringVar(value="")
        self.clearance_mode_var = ttk.StringVar(value="auto")
        self.clearance_session_hint_var = ttk.StringVar(value="尚未检查网站登录状态")
        self.workspace_var = ttk.StringVar(value=str(self.ops.workspace_root))
        self.settings_path_var = ttk.StringVar(value=str(self.config_store.config_path))
        self.theme_var = ttk.StringVar(value=self.config.theme_name)
        self.auto_open_var = ttk.BooleanVar(value=self.config.auto_open_results)
        self.auto_update_check_var = ttk.BooleanVar(value=self.config.check_updates_on_launch)
        self.req2_browser_var = ttk.StringVar(value=self.config.req2_browser_preference)
        self.last_update_var = ttk.StringVar(value=f"当前版本：{APP_VERSION}")

        self.task_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.task_running = False
        self.log_widget: ScrolledText | None = None
        self.sync_panel: ExcelSyncPanel | None = None

        self.apply_visual_styles()
        self.build_ui()
        self.refresh_paths()
        self.refresh_clearance_session_hint()
        self.window.after(180, self.process_task_queue)

        if self.config.check_updates_on_launch:
            self.start_task("检查更新", self.ops.check_for_updates, self.on_check_updates_done)

    def apply_visual_styles(self) -> None:
        style = tkttk.Style(self.window)
        style.configure("TNotebook.Tab", padding=(18, 10), font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TButton", padding=(12, 8), font=("Microsoft YaHei UI", 10))

    def build_ui(self) -> None:
        outer = ttk.Frame(self.window, padding=(16, 14, 16, 14))
        outer.pack(fill=BOTH, expand=True)

        self.build_header(outer)

        notebook_shell = self.make_surface(outer, padding=8, bootstyle="light")
        notebook_shell.pack(fill=BOTH, expand=True, pady=(12, 10))

        self.notebook = ttk.Notebook(notebook_shell, bootstyle="primary")
        self.notebook.pack(fill=BOTH, expand=True)

        self.home_tab = ttk.Frame(self.notebook, padding=18)
        self.contract_tab = ttk.Frame(self.notebook, padding=18)
        self.clearance_tab = ttk.Frame(self.notebook, padding=18)
        self.lineup_tab = ttk.Frame(self.notebook, padding=18)
        self.sync_tab = ttk.Frame(self.notebook, padding=18)
        self.settings_tab = ttk.Frame(self.notebook, padding=18)

        self.notebook.add(self.home_tab, text="首页")
        self.notebook.add(self.contract_tab, text="合同生成")
        self.notebook.add(self.clearance_tab, text="通关查询")
        self.notebook.add(self.lineup_tab, text="船期矩阵")
        self.notebook.add(self.sync_tab, text="表格同步")
        self.notebook.add(self.settings_tab, text="设置")

        self.build_home_tab()
        self.build_contract_tab()
        self.build_clearance_tab()
        self.build_lineup_tab()
        self.build_sync_tab()
        self.build_settings_tab()

        log_shell = self.make_surface(outer, padding=0, bootstyle="light")
        log_shell.pack(fill=BOTH, expand=False)
        log_header = ttk.Frame(log_shell, padding=(16, 14, 16, 0))
        log_header.pack(fill=X)
        ttk.Label(log_header, text="运行日志", font=("Microsoft YaHei UI", 15, "bold")).pack(side=LEFT)
        ttk.Label(log_header, text="把执行过程留在下面，既方便验收，也方便排错。", bootstyle="secondary").pack(side=LEFT, padx=(10, 0))
        ttk.Button(log_header, text="清空日志", bootstyle="secondary", command=self.clear_log).pack(side=RIGHT)
        log_body = ttk.Frame(log_shell, padding=(16, 10, 16, 16))
        log_body.pack(fill=BOTH, expand=True)
        self.log_widget = ScrolledText(
            log_body,
            height=8,
            font=("Microsoft YaHei UI", 10),
            background="#f5f9fc",
            foreground="#24384b",
            insertbackground="#24384b",
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
        )
        self.log_widget.pack(fill=BOTH, expand=True)
        self.log("应用已启动。")

        status_shell = self.make_surface(outer, padding=(12, 8), bootstyle="light")
        status_shell.pack(fill=X)
        ttk.Label(status_shell, textvariable=self.status_text, anchor="w", bootstyle="secondary").pack(fill=X)

    def build_header(self, parent: ttk.Frame) -> None:
        header = ttk.Frame(parent)
        header.pack(fill=X)
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=0)

        left = self.make_surface(header, padding=18, bootstyle="light")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.make_pill(left, "LOCAL-FIRST MARITIME WORKSPACE", "primary").pack(anchor="w")
        ttk.Label(left, text="AMS Assistant Desktop", font=("Microsoft YaHei UI", 25, "bold"), bootstyle="primary").pack(anchor="w", pady=(8, 2))
        ttk.Label(
            left,
            text="把合同生成、通关查询、船期预留和 Excel 自动同步，整理进一个更好上手、也更愿意打开的本地工作台。",
            font=("Microsoft YaHei UI", 11),
            bootstyle="secondary",
            wraplength=760,
            justify="left",
        ).pack(anchor="w")
        tags = ttk.Frame(left)
        tags.pack(fill=X, pady=(12, 0))
        self.make_pill(tags, "合同", "success").pack(side=LEFT, padx=(0, 8))
        self.make_pill(tags, "通关", "info").pack(side=LEFT, padx=(0, 8))
        self.make_pill(tags, "表格同步", "primary").pack(side=LEFT, padx=(0, 8))
        self.make_pill(tags, "可更新", "warning").pack(side=LEFT)

        right = self.make_surface(header, padding=14, bootstyle="light")
        right.grid(row=0, column=1, sticky="nsew")
        ttk.Label(right, text=f"版本 {APP_VERSION}", font=("Microsoft YaHei UI", 10, "bold"), bootstyle="primary").pack(anchor="e")
        ttk.Label(
            right,
            text="不想翻文件夹，就从这里把事做完。",
            bootstyle="secondary",
            wraplength=260,
            justify="right",
        ).pack(anchor="e", pady=(4, 8))
        actions1 = ttk.Frame(right)
        actions1.pack(anchor="e")
        ttk.Button(actions1, text="打开工作区", bootstyle="secondary", command=self.open_workspace_root).pack(side=LEFT, padx=4)
        ttk.Button(actions1, text="帮助中心", bootstyle="info", command=self.open_help_center).pack(side=LEFT, padx=4)
        ttk.Button(actions1, text="反馈建议", bootstyle="success", command=self.open_feedback_mail).pack(side=LEFT, padx=4)
        actions2 = ttk.Frame(right)
        actions2.pack(anchor="e", pady=(8, 0))
        ttk.Button(actions2, text="检查更新", bootstyle="warning", command=self.check_updates_clicked).pack(side=LEFT, padx=4)
        ttk.Button(actions2, text="GitHub", bootstyle="secondary", command=lambda: webbrowser.open(APP_REPO_URL)).pack(side=LEFT, padx=4)

    def build_home_tab(self) -> None:
        self.home_tab.columnconfigure(0, weight=7)
        self.home_tab.columnconfigure(1, weight=4)

        hero = self.make_surface(self.home_tab, padding=22, bootstyle="primary")
        hero.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 12))
        self.make_pill(hero, "欢迎使用", "light").pack(anchor="w")
        ttk.Label(hero, text="如果你只是想做事，就从这里开始。", font=("Microsoft YaHei UI", 15, "bold"), bootstyle="inverse-primary").pack(anchor="w", pady=(10, 6))
        ttk.Label(
            hero,
            text="最顺手的方式通常是：先打开固定输入文件，填好并保存，再回到应用里执行。结果、帮助页和最新输出，都尽量保持在固定入口。",
            bootstyle="inverse-primary",
            wraplength=700,
            justify="left",
        ).pack(anchor="w")
        self.make_step_row(
            hero,
            [
                ("1", "打开输入文件", "先把 Excel 或工作簿准备好。"),
                ("2", "回到应用执行", "点主按钮，不用记命令。"),
                ("3", "直接看最新结果", "输出会保留固定入口。"),
            ],
            bootstyle="light",
        ).pack(fill=X, pady=(14, 0))
        hero_buttons = ttk.Frame(hero)
        hero_buttons.pack(fill=X, pady=(14, 0))
        ttk.Button(hero_buttons, text=CONTRACT_FEATURE_NAME, bootstyle="light", command=lambda: self.select_tab(self.contract_tab)).pack(side=LEFT, padx=4)
        ttk.Button(hero_buttons, text=CLEARANCE_FEATURE_NAME, bootstyle="light", command=lambda: self.select_tab(self.clearance_tab)).pack(side=LEFT, padx=4)
        ttk.Button(hero_buttons, text=SYNC_FEATURE_NAME, bootstyle="light", command=lambda: self.select_tab(self.sync_tab)).pack(side=LEFT, padx=4)
        ttk.Button(hero_buttons, text="打开帮助中心", bootstyle="info", command=self.open_help_center).pack(side=LEFT, padx=4)

        summary = self.make_surface(self.home_tab, padding=20, bootstyle="light")
        summary.grid(row=0, column=1, sticky="nsew", pady=(0, 12))
        ttk.Label(summary, text="当前环境", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        ttk.Label(summary, textvariable=self.last_update_var, bootstyle="primary", font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w", pady=(10, 0))
        self.make_path_block(summary, "当前工作区", textvariable=self.workspace_var, bootstyle="secondary", wraplength=360).pack(fill=X, pady=(12, 0))
        self.make_path_block(summary, "设置文件", textvariable=self.settings_path_var, bootstyle="secondary", wraplength=360).pack(fill=X, pady=(10, 0))
        self.make_path_block(summary, "网站登录态", textvariable=self.clearance_session_hint_var, bootstyle="secondary", wraplength=360).pack(fill=X, pady=(10, 0))

        feature_grid = ttk.Frame(self.home_tab)
        feature_grid.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 12))
        feature_grid.columnconfigure(0, weight=1)
        feature_grid.columnconfigure(1, weight=1)

        self.make_feature_tile(
            feature_grid,
            row=0,
            column=0,
            title=CONTRACT_FEATURE_NAME,
            status_text="已可用",
            description="填固定 Excel，一键生成 Word 合同。适合不想折腾路径和模板逻辑的人。",
            primary_label="打开输入文件",
            primary_command=lambda: self.open_path(self.ops.contract_input_path),
            secondary_label="功能说明",
            secondary_command=lambda: self.open_help_page("contract"),
            bootstyle="success",
        )
        self.make_feature_tile(
            feature_grid,
            row=0,
            column=1,
            title=CLEARANCE_FEATURE_NAME,
            status_text="已可用",
            description="保存登录态后，就能做单票测试和整表自动回填。",
            primary_label="打开输入文件",
            primary_command=lambda: self.open_path(self.ops.clearance_input_path),
            secondary_label="功能说明",
            secondary_command=lambda: self.open_help_page("clearance"),
            bootstyle="info",
        )
        self.make_feature_tile(
            feature_grid,
            row=1,
            column=0,
            title=SYNC_FEATURE_NAME,
            status_text="已可用",
            description="尽量按源表原样复制，只排除你勾掉的列，适合日报同步和系统间搬运。",
            primary_label="打开示例文件",
            primary_command=lambda: self.open_path(self.ops.sync_examples_dir),
            secondary_label="功能说明",
            secondary_command=lambda: self.open_help_page("sync"),
            bootstyle="primary",
        )
        self.make_feature_tile(
            feature_grid,
            row=1,
            column=1,
            title=LINEUP_FEATURE_NAME,
            status_text="预留入口",
            description="船期表和港区矩阵的入口已留好，后续会继续接入实际查询和报告。",
            primary_label="查看预留目录",
            primary_command=lambda: self.open_path(self.ops.lineup_dir),
            secondary_label="功能说明",
            secondary_command=lambda: self.open_help_page("lineup"),
            bootstyle="warning",
        )

        latest = self.make_surface(self.home_tab, padding=18, bootstyle="light")
        latest.grid(row=2, column=0, columnspan=2, sticky="ew")
        ttk.Label(latest, text="最近结果快捷入口", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        ttk.Label(latest, text="如果你只是想确认最新产物，下面这些按钮比翻目录更快。", bootstyle="secondary", wraplength=980).pack(anchor="w", pady=(6, 0))
        latest_buttons = ttk.Frame(latest)
        latest_buttons.pack(fill=X, pady=(12, 0))
        ttk.Button(latest_buttons, text="打开最新合同", bootstyle="success", command=self.open_latest_contract_document).pack(side=LEFT, padx=4)
        ttk.Button(latest_buttons, text="打开最新通关工作簿", bootstyle="info", command=self.open_latest_clearance_workbook).pack(side=LEFT, padx=4)
        ttk.Button(latest_buttons, text="打开最新网站报告", bootstyle="secondary", command=self.open_latest_clearance_site_report).pack(side=LEFT, padx=4)
        ttk.Button(latest_buttons, text="反馈建议", bootstyle="warning", command=self.open_feedback_mail).pack(side=RIGHT, padx=4)

    def build_contract_tab(self) -> None:
        self.contract_tab.columnconfigure(0, weight=7)
        self.contract_tab.columnconfigure(1, weight=4)

        left = self.make_surface(self.contract_tab, padding=22, bootstyle="success")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.make_pill(left, CONTRACT_FEATURE_NAME, "light").pack(anchor="w")
        ttk.Label(left, text="填固定 Excel，回到这里，一键出合同。", font=("Microsoft YaHei UI", 15, "bold"), bootstyle="inverse-success").pack(anchor="w", pady=(10, 6))
        ttk.Label(
            left,
            text="推荐普通用户就走固定模板这条线。这样路径、输出位置和验收口径都会更稳定。",
            bootstyle="inverse-success",
            wraplength=720,
            justify="left",
        ).pack(anchor="w")
        self.make_step_row(
            left,
            [
                ("1", "打开 Excel", "在固定输入文件里填数据。"),
                ("2", "保存并关闭", "避免被 Excel 占用。"),
                ("3", "点一键生成", "结果和摘要会一起生成。"),
            ],
            bootstyle="light",
        ).pack(fill=X, pady=(14, 0))
        self.make_path_block(left, "当前输入文件", textvariable=self.contract_input_var, bootstyle="inverse-success", wraplength=720).pack(fill=X, pady=(14, 0))
        self.make_path_block(left, "当前结果目录", textvariable=self.contract_result_var, bootstyle="inverse-success", wraplength=720).pack(fill=X, pady=(10, 0))

        row1 = ttk.Frame(left)
        row1.pack(fill=X, pady=(14, 0))
        ttk.Button(row1, text="打开输入 Excel", bootstyle="light", command=lambda: self.open_path(self.ops.contract_input_path)).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="重新生成空白模板", bootstyle="warning", command=self.contract_reset_template_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="一键生成合同", bootstyle="success", command=self.contract_generate_clicked).pack(side=LEFT, padx=4)

        row2 = ttk.Frame(left)
        row2.pack(fill=X, pady=(8, 0))
        ttk.Button(row2, text="打开结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.contract_result_dir)).pack(side=LEFT, padx=4)
        ttk.Button(row2, text="打开最新合同", bootstyle="info", command=self.open_latest_contract_document).pack(side=LEFT, padx=4)
        ttk.Button(row2, text="选择别的 Excel 并生成", bootstyle="secondary", command=self.contract_pick_other_excel_clicked).pack(side=LEFT, padx=4)

        right = ttk.Frame(self.contract_tab)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)

        tips = self.make_surface(right, padding=18, bootstyle="light")
        tips.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(tips, text="使用建议", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        ttk.Label(tips, text="这块最适合“固定格式、重复生成”的工作。", bootstyle="secondary", wraplength=360).pack(anchor="w", pady=(6, 0))
        self.make_pill(tips, "尽量先用固定输入文件", "success").pack(anchor="w", pady=(12, 0))
        ttk.Label(tips, text="如果 Word 没生成出来，先检查 Excel / Word / WPS 是否还开着。", bootstyle="secondary", wraplength=360).pack(anchor="w", pady=(10, 0))
        ttk.Button(tips, text="打开功能说明", bootstyle="info", command=lambda: self.open_help_page("contract")).pack(anchor="w", pady=(14, 0))

        latest = self.make_surface(right, padding=18, bootstyle="light")
        latest.grid(row=1, column=0, sticky="ew")
        ttk.Label(latest, text="结果入口", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        ttk.Label(latest, text="最新合同会维护一个固定文件名，方便你快速确认。", bootstyle="secondary", wraplength=360).pack(anchor="w", pady=(6, 0))
        ttk.Button(latest, text="打开最新合同", bootstyle="success", command=self.open_latest_contract_document).pack(anchor="w", pady=(14, 0))
        ttk.Button(latest, text="打开结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.contract_result_dir)).pack(anchor="w", pady=(8, 0))

    def build_clearance_tab(self) -> None:
        self.clearance_tab.columnconfigure(0, weight=1)
        self.clearance_tab.columnconfigure(1, weight=1)

        intro = self.make_surface(self.clearance_tab, padding=20, bootstyle="info")
        intro.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        self.make_pill(intro, CLEARANCE_FEATURE_NAME, "light").pack(anchor="w")
        ttk.Label(intro, text="把登录、测试和整表回填，收进同一个页面。", font=("Microsoft YaHei UI", 15, "bold"), bootstyle="inverse-info").pack(anchor="w", pady=(10, 6))
        ttk.Label(
            intro,
            text="先保存一次网站登录态，之后就可以先查一票，再整表回填。这样最容易排查是登录问题还是数据问题。",
            bootstyle="inverse-info",
            wraplength=980,
            justify="left",
        ).pack(anchor="w")
        self.make_path_block(intro, "当前输入文件", textvariable=self.clearance_input_var, bootstyle="inverse-info", wraplength=980).pack(fill=X, pady=(14, 0))
        self.make_path_block(intro, "当前结果目录", textvariable=self.clearance_result_var, bootstyle="inverse-info", wraplength=980).pack(fill=X, pady=(10, 0))

        session_card = self.make_surface(self.clearance_tab, padding=18, bootstyle="light")
        session_card.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=(0, 12))
        ttk.Label(session_card, text="网站登录", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        ttk.Label(session_card, text="首次使用时，先保存一次登录态。以后通常只要检查一下是否过期。", bootstyle="secondary", wraplength=460).pack(anchor="w", pady=(6, 0))
        self.make_path_block(session_card, "当前登录态提示", textvariable=self.clearance_session_hint_var, bootstyle="secondary", wraplength=420).pack(fill=X, pady=(12, 0))
        session_row = ttk.Frame(session_card)
        session_row.pack(fill=X, pady=(14, 0))
        ttk.Button(session_row, text="首次登录并保存登录态", bootstyle="primary", command=self.clearance_capture_session_clicked).pack(side=LEFT, padx=4)
        ttk.Button(session_row, text="检查登录态", bootstyle="info", command=self.clearance_check_session_clicked).pack(side=LEFT, padx=4)
        ttk.Button(session_row, text="打开登录态目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.clearance_site_session_dir)).pack(side=LEFT, padx=4)

        query_card = self.make_surface(self.clearance_tab, padding=18, bootstyle="light")
        query_card.grid(row=1, column=1, sticky="nsew", pady=(0, 12))
        ttk.Label(query_card, text="单票测试", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        ttk.Label(query_card, text="先查一票，确认网站查询是通的，再做整表更新会更安心。", bootstyle="secondary", wraplength=460).pack(anchor="w", pady=(6, 0))
        ttk.Label(query_card, text="识别值", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(12, 0))
        ttk.Entry(query_card, textvariable=self.clearance_query_var, width=42).pack(fill=X, pady=(6, 0))
        filters = ttk.Frame(query_card)
        filters.pack(fill=X, pady=(10, 0))
        ttk.Label(filters, text="模式", font=("Microsoft YaHei UI", 10, "bold")).pack(side=LEFT)
        ttk.Combobox(filters, textvariable=self.clearance_mode_var, values=["auto", "blNo", "entryNo", "ctnrNo"], width=12, state="readonly").pack(side=LEFT, padx=(8, 16))
        ttk.Label(filters, text="I/E", font=("Microsoft YaHei UI", 10, "bold")).pack(side=LEFT)
        ttk.Combobox(filters, textvariable=self.clearance_iemark_var, values=["", "E", "I"], width=8, state="readonly").pack(side=LEFT, padx=(8, 0))
        query_actions = ttk.Frame(query_card)
        query_actions.pack(fill=X, pady=(14, 0))
        ttk.Button(query_actions, text="查询这一票", bootstyle="success", command=self.clearance_query_one_clicked).pack(side=LEFT, padx=4)
        ttk.Button(query_actions, text="打开查询结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.clearance_site_query_dir)).pack(side=LEFT, padx=4)

        update_card = self.make_surface(self.clearance_tab, padding=20, bootstyle="light")
        update_card.grid(row=2, column=0, columnspan=2, sticky="ew")
        ttk.Label(update_card, text="整表自动回填", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        ttk.Label(
            update_card,
            text="等单票测试正常后，再让它读取当前工作簿、自动查询网站并回填结果。你也可以临时选别的 Excel 做一次性更新。",
            bootstyle="secondary",
            wraplength=980,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))
        update_row = ttk.Frame(update_card)
        update_row.pack(fill=X, pady=(14, 0))
        ttk.Button(update_row, text="打开输入 Excel", bootstyle="light", command=lambda: self.open_path(self.ops.clearance_input_path)).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="重新生成空白模板", bootstyle="warning", command=self.clearance_reset_template_clicked).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="用当前输入表自动回填", bootstyle="success", command=self.clearance_update_clicked).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="选择别的 Excel", bootstyle="secondary", command=self.clearance_pick_other_excel_clicked).pack(side=LEFT, padx=4)

        result_row = ttk.Frame(update_card)
        result_row.pack(fill=X, pady=(10, 0))
        ttk.Button(result_row, text="打开最新工作簿", bootstyle="info", command=self.open_latest_clearance_workbook).pack(side=LEFT, padx=4)
        ttk.Button(result_row, text="打开最新网站报告", bootstyle="secondary", command=self.open_latest_clearance_site_report).pack(side=LEFT, padx=4)
        ttk.Button(result_row, text="打开功能说明", bootstyle="info", command=lambda: self.open_help_page("clearance")).pack(side=RIGHT, padx=4)

    def build_lineup_tab(self) -> None:
        self.lineup_tab.columnconfigure(0, weight=3)
        self.lineup_tab.columnconfigure(1, weight=2)

        intro = self.make_surface(self.lineup_tab, padding=22, bootstyle="warning")
        intro.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.make_pill(intro, LINEUP_FEATURE_NAME, "light").pack(anchor="w")
        ttk.Label(intro, text="这里会接入后续的船期表与港区矩阵工作流。", font=("Microsoft YaHei UI", 15, "bold"), bootstyle="inverse-warning").pack(anchor="w", pady=(10, 6))
        ttk.Label(
            intro,
            text="现在先把入口、目录和说明整理好，等后续业务规则更明确时，尽量还是沿用同一套桌面应用体验去验收。",
            bootstyle="inverse-warning",
            wraplength=700,
            justify="left",
        ).pack(anchor="w")
        self.make_path_block(intro, "当前预留目录", text=str(self.ops.lineup_dir), bootstyle="inverse-warning", wraplength=680).pack(fill=X, pady=(14, 0))
        row = ttk.Frame(intro)
        row.pack(fill=X, pady=(14, 0))
        ttk.Button(row, text="打开预留目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.lineup_dir)).pack(side=LEFT, padx=4)
        ttk.Button(row, text="打开功能说明", bootstyle="info", command=lambda: self.open_help_page("lineup")).pack(side=LEFT, padx=4)

        roadmap = self.make_surface(self.lineup_tab, padding=18, bootstyle="light")
        roadmap.grid(row=0, column=1, sticky="nsew")
        ttk.Label(roadmap, text="后续会往这里放什么", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        self.make_step_row(
            roadmap,
            [
                ("A", "船期数据读取", "从原始表或站点拿到航次信息。"),
                ("B", "矩阵生成", "整理成港口 x 区域的可读结果。"),
                ("C", "报告输出", "给验收和业务查看固定结果入口。"),
            ],
            bootstyle="warning",
        ).pack(fill=X, pady=(12, 0))
        ttk.Label(roadmap, text="你以后验收这一块时，也尽量从桌面应用里走，而不是回到脚本层单独折腾。", bootstyle="secondary", wraplength=360).pack(anchor="w", pady=(14, 0))

    def build_sync_tab(self) -> None:
        self.sync_panel = ExcelSyncPanel(
            parent=self.sync_tab,
            window=self.window,
            service=self.ops.build_sync_service(status_callback=lambda: None),
            open_path=self.open_path,
            open_help=self.open_help_page,
        )

    def build_settings_tab(self) -> None:
        self.settings_tab.columnconfigure(0, weight=3)
        self.settings_tab.columnconfigure(1, weight=2)

        left = self.make_surface(self.settings_tab, padding=20, bootstyle="light")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        ttk.Label(left, text="工作区与界面", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        ttk.Label(left, text="主要控制工作区位置、主题和浏览器偏好。", bootstyle="secondary").pack(anchor="w", pady=(6, 0))

        row1 = ttk.Frame(left)
        row1.pack(fill=X, pady=(14, 0))
        ttk.Label(row1, text="工作区根目录", width=14).pack(side=LEFT)
        ttk.Entry(row1, textvariable=self.workspace_var, width=68).pack(side=LEFT, padx=6, fill=X, expand=True)
        ttk.Button(row1, text="选择文件夹", bootstyle="secondary", command=self.choose_workspace_clicked).pack(side=LEFT)

        row2 = ttk.Frame(left)
        row2.pack(fill=X, pady=(12, 0))
        ttk.Label(row2, text="界面主题", width=14).pack(side=LEFT)
        ttk.Combobox(row2, textvariable=self.theme_var, values=ThemeChoices, width=18, state="readonly").pack(side=LEFT, padx=6)

        row3 = ttk.Frame(left)
        row3.pack(fill=X, pady=(12, 0))
        ttk.Label(row3, text="通关查询浏览器", width=14).pack(side=LEFT)
        ttk.Combobox(row3, textvariable=self.req2_browser_var, values=BrowserChoices, width=18, state="readonly").pack(side=LEFT, padx=6)
        ttk.Label(row3, text="推荐保持 auto，会优先尝试 Edge。", bootstyle="secondary").pack(side=LEFT, padx=6)

        row4 = ttk.Frame(left)
        row4.pack(fill=X, pady=(14, 0))
        ttk.Checkbutton(row4, text="执行后自动打开结果目录", variable=self.auto_open_var, bootstyle="round-toggle").pack(side=LEFT, padx=4)
        ttk.Checkbutton(row4, text="启动时自动检查更新", variable=self.auto_update_check_var, bootstyle="round-toggle").pack(side=LEFT, padx=16)

        actions = ttk.Frame(left)
        actions.pack(fill=X, pady=(16, 0))
        ttk.Button(actions, text="保存设置", bootstyle="success", command=self.save_settings_clicked).pack(side=LEFT, padx=4)
        ttk.Button(actions, text="打开工作区", bootstyle="secondary", command=self.open_workspace_root).pack(side=LEFT, padx=4)
        ttk.Button(actions, text="打开用户数据目录", bootstyle="secondary", command=self.open_settings_root).pack(side=LEFT, padx=4)

        right = ttk.Frame(self.settings_tab)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)

        support = self.make_surface(right, padding=18, bootstyle="light")
        support.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(support, text="支持与更新", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        update_mode = "支持自动安装更新" if portable_install_root() is not None else "源码模式：自动更新不可用"
        ttk.Label(support, text=update_mode, bootstyle="primary", font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w", pady=(10, 0))
        ttk.Label(
            support,
            text="release 版支持下载新版本、替换程序本体并保留工作区、设置和登录态。",
            bootstyle="secondary",
            wraplength=360,
        ).pack(anchor="w", pady=(8, 0))
        buttons = ttk.Frame(support)
        buttons.pack(fill=X, pady=(14, 0))
        ttk.Button(buttons, text="检查更新", bootstyle="warning", command=self.check_updates_clicked).pack(side=LEFT, padx=4)
        ttk.Button(buttons, text="帮助中心", bootstyle="info", command=self.open_help_center).pack(side=LEFT, padx=4)
        ttk.Button(buttons, text="反馈建议", bootstyle="success", command=self.open_feedback_mail).pack(side=LEFT, padx=4)

        meta = self.make_surface(right, padding=18, bootstyle="light")
        meta.grid(row=1, column=0, sticky="ew")
        ttk.Label(meta, text="当前配置", font=("Microsoft YaHei UI", 15, "bold")).pack(anchor="w")
        self.make_path_block(meta, "设置文件", textvariable=self.settings_path_var, bootstyle="secondary", wraplength=360).pack(fill=X, pady=(12, 0))
        self.make_path_block(meta, "当前工作区", textvariable=self.workspace_var, bootstyle="secondary", wraplength=360).pack(fill=X, pady=(10, 0))

    def make_surface(self, parent: ttk.Frame, padding: int | tuple[int, ...] = 18, bootstyle: str = "light") -> ttk.Frame:
        return ttk.Frame(parent, padding=padding, bootstyle=bootstyle, borderwidth=1, relief="solid")

    def make_pill(self, parent: ttk.Frame, text: str, bootstyle: str = "primary") -> ttk.Label:
        return ttk.Label(parent, text=text, bootstyle=f"inverse-{bootstyle}", font=("Microsoft YaHei UI", 9, "bold"), padding=(10, 5))

    def make_step_row(
        self,
        parent: ttk.Frame,
        steps: list[tuple[str, str, str]],
        bootstyle: str = "light",
    ) -> ttk.Frame:
        row = ttk.Frame(parent)
        for column, (index, title, body) in enumerate(steps):
            row.columnconfigure(column, weight=1)
            card = self.make_surface(row, padding=14, bootstyle=bootstyle)
            card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 8, 0))
            self.make_pill(card, index, bootstyle if bootstyle != "light" else "primary").pack(anchor="w")
            ttk.Label(card, text=title, font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(10, 4))
            ttk.Label(card, text=body, bootstyle="secondary", wraplength=210, justify="left").pack(anchor="w")
        return row

    def make_feature_tile(
        self,
        parent: ttk.Frame,
        row: int,
        column: int,
        title: str,
        status_text: str,
        description: str,
        primary_label: str,
        primary_command,
        secondary_label: str,
        secondary_command,
        bootstyle: str,
    ) -> None:
        card = self.make_surface(parent, padding=18, bootstyle="light")
        card.grid(row=row, column=column, sticky="nsew", padx=(0 if column == 0 else 8, 0), pady=(0 if row == 0 else 8, 0))
        self.make_pill(card, status_text, bootstyle).pack(anchor="w")
        ttk.Label(card, text=title, font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(12, 4))
        ttk.Label(card, text=description, bootstyle="secondary", wraplength=420, justify="left").pack(anchor="w")
        buttons = ttk.Frame(card)
        buttons.pack(fill=X, pady=(14, 0))
        ttk.Button(buttons, text=primary_label, bootstyle=bootstyle, command=primary_command).pack(side=LEFT, padx=(0, 6))
        ttk.Button(buttons, text=secondary_label, bootstyle="secondary", command=secondary_command).pack(side=LEFT)

    def make_path_block(
        self,
        parent: ttk.Frame,
        title: str,
        text: str | None = None,
        textvariable=None,
        bootstyle: str = "secondary",
        wraplength: int = 720,
    ) -> ttk.Frame:
        block = self.make_surface(parent, padding=(12, 10), bootstyle="light")
        ttk.Label(block, text=title, font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")
        kwargs: dict[str, Any] = {
            "wraplength": wraplength,
            "justify": "left",
            "font": ("Microsoft YaHei UI", 10),
            "bootstyle": bootstyle,
        }
        if textvariable is not None:
            kwargs["textvariable"] = textvariable
        else:
            kwargs["text"] = text or ""
        ttk.Label(block, **kwargs).pack(anchor="w", pady=(4, 0))
        return block

    def clear_log(self) -> None:
        if self.log_widget is None:
            return
        self.log_widget.delete("1.0", END)

    def select_tab(self, tab_frame: ttk.Frame) -> None:
        self.notebook.select(tab_frame)

    def log(self, text: str) -> None:
        if self.log_widget is None:
            return
        self.log_widget.insert(END, text + "\n")
        self.log_widget.see(END)

    def set_status(self, text: str) -> None:
        self.status_text.set(text)

    def process_task_queue(self) -> None:
        try:
            while True:
                kind, payload = self.task_queue.get_nowait()
                if kind == "success":
                    title, result, callback = payload
                    self.task_running = False
                    self.set_status(f"{title}已完成")
                    self.log(f"[完成] {title}")
                    if callback:
                        callback(result)
                elif kind == "error":
                    title, error_text = payload
                    self.task_running = False
                    self.set_status(f"{title}失败")
                    self.log(f"[失败] {title}\n{error_text}")
                    messagebox.showerror("操作失败", f"{title}失败。\n\n{self.humanize_error_text(error_text)}")
        except queue.Empty:
            pass
        finally:
            self.window.after(180, self.process_task_queue)

    def start_task(self, title: str, worker: Callable[[], Any], callback: Callable[[Any], None] | None = None) -> None:
        if self.task_running:
            messagebox.showinfo("请稍等", "当前还有一个任务在执行，请稍后再试。")
            return
        self.task_running = True
        self.set_status(f"{title}进行中")
        self.log(f"[开始] {title}")

        def run() -> None:
            try:
                result = worker()
            except Exception:
                self.task_queue.put(("error", (title, traceback.format_exc())))
                return
            self.task_queue.put(("success", (title, result, callback)))

        threading.Thread(target=run, daemon=True).start()

    def refresh_paths(self) -> None:
        self.contract_input_var.set(str(self.ops.contract_input_path))
        self.contract_result_var.set(str(self.ops.contract_result_dir))
        self.clearance_input_var.set(str(self.ops.clearance_input_path))
        self.clearance_result_var.set(str(self.ops.clearance_result_root))
        self.workspace_var.set(str(self.ops.workspace_root))

    def refresh_clearance_session_hint(self) -> None:
        session_file = self.ops.clearance_site_session_dir / "req2_site_session.json"
        if session_file.exists():
            self.clearance_session_hint_var.set(f"已保存通关查询登录态：{session_file}")
        else:
            self.clearance_session_hint_var.set("还没有保存网站登录态")

    def open_path(self, path: Path) -> None:
        try:
            open_in_file_explorer(path)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def open_workspace_root(self) -> None:
        self.open_path(self.ops.workspace_root)

    def open_settings_root(self) -> None:
        self.open_path(self.config_store.config_path.parent)

    def open_help_center(self) -> None:
        self.open_help_page("index")

    def open_help_page(self, key: str) -> None:
        mapping = {
            "index": help_assets_dir() / "index.html",
            "contract": help_assets_dir() / "contract.html",
            "clearance": help_assets_dir() / "clearance.html",
            "lineup": help_assets_dir() / "lineup.html",
            "sync": help_assets_dir() / "sync.html",
        }
        target = mapping.get(key)
        if target and target.exists():
            self.open_path(target)
            return
        messagebox.showinfo("帮助页面缺失", f"没有找到帮助页面：{key}")

    def open_feedback_mail(self) -> None:
        webbrowser.open("mailto:cyh29hao@sjtu.edu.cn?subject=AMS%20Assistant%20Feedback")

    def open_latest_contract_document(self) -> None:
        path = self.ops.contract_result_dir / "00-latest-contract.docx"
        self.open_if_exists(path, "还没有找到最新合同，请先生成一次合同。")

    def open_latest_clearance_workbook(self) -> None:
        path = self.ops.clearance_updated_dir / "00-latest-clearance-workbook.xlsx"
        self.open_if_exists(path, "还没有找到最新的回填工作簿，请先执行一次整表自动回填。")

    def open_latest_clearance_site_report(self) -> None:
        path = self.ops.clearance_site_query_dir / "00-latest-site-report.txt"
        self.open_if_exists(path, "还没有找到最新的网站报告，请先执行一次查询。")

    def open_if_exists(self, path: Path, missing_message: str) -> None:
        if path.exists():
            self.open_path(path)
            return
        messagebox.showinfo("暂时没有结果", missing_message)

    def humanize_error_text(self, error_text: str) -> str:
        lowered = error_text.lower()
        if "permissionerror" in lowered:
            return "文件可能正在被 Excel、Word 或 WPS 占用。请先关闭相关文件，再重试。\n\n原始信息：\n" + error_text
        if "filenotfounderror" in lowered:
            return "需要的文件没有找到，可能被移动、重命名，或者还没有生成。\n\n原始信息：\n" + error_text
        if "no saved req2 site session" in lowered or "session" in lowered and "req2" in lowered:
            return "还没有可用的网站登录态。请先在通关查询页面点击“首次登录并保存登录态”。\n\n原始信息：\n" + error_text
        if "browser" in lowered:
            return "浏览器没有正常启动。请先确认本机可以打开 Edge 或 Chrome，再到设置里切换浏览器偏好。\n\n原始信息：\n" + error_text
        if "valueerror" in lowered:
            return "输入内容没有通过校验。请检查 Excel 字段、日期格式、区域格式或必填项。\n\n原始信息：\n" + error_text
        if "requestexception" in lowered or "site request failed" in lowered:
            return "网站请求失败。可能是网络问题、网站暂时不可用，或者登录态已经失效。\n\n原始信息：\n" + error_text
        return error_text

    def choose_workspace_clicked(self) -> None:
        selected = filedialog.askdirectory(initialdir=str(self.ops.workspace_root))
        if selected:
            self.workspace_var.set(selected)

    def save_settings_clicked(self) -> None:
        self.config = AppConfig(
            workspace_root=self.workspace_var.get().strip(),
            theme_name=self.theme_var.get().strip() or "flatly",
            auto_open_results=bool(self.auto_open_var.get()),
            check_updates_on_launch=bool(self.auto_update_check_var.get()),
            req1_input_filename=self.config.req1_input_filename,
            req2_input_filename=self.config.req2_input_filename,
            req2_browser_preference=self.req2_browser_var.get().strip() or "auto",
        )
        self.config_store.save(self.config)
        self.ops = AmsOperations(self.config)
        self.ops.ensure_workspace()
        self.window.style.theme_use(self.config.theme_name)
        self.apply_visual_styles()
        self.settings_path_var.set(str(self.config_store.config_path))
        self.refresh_paths()
        self.refresh_clearance_session_hint()
        self.set_status("设置已保存")
        self.log(f"设置已保存。工作区：{self.ops.workspace_root}")
        messagebox.showinfo("设置已保存", "新的设置已经保存并立即生效。")

    def contract_reset_template_clicked(self) -> None:
        self.start_task("重新生成合同模板", self.ops.contract_reset_template, self.on_contract_template_reset)

    def on_contract_template_reset(self, result: dict[str, Any]) -> None:
        self.refresh_paths()
        self.log(f"合同模板已生成：{result['output_path']}")
        if self.config.auto_open_results:
            self.open_path(self.ops.contract_dir)

    def contract_generate_clicked(self) -> None:
        self.start_task("生成合同", self.ops.contract_generate_from_current, self.on_contract_generated)

    def contract_pick_other_excel_clicked(self) -> None:
        path = filedialog.askopenfilename(
            title="选择一个 Excel 工作簿",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialdir=str(self.ops.contract_dir),
        )
        if not path:
            return
        self.start_task("生成合同", lambda: self.ops.contract_generate_from_file(path), self.on_contract_generated)

    def on_contract_generated(self, result: dict[str, Any]) -> None:
        self.log(f"合同已生成：{result['document_path']}")
        self.log(f"摘要已生成：{result['summary_path']}")
        self.log(f"固定入口：{result['latest_document_path']}")
        if self.config.auto_open_results:
            self.open_path(Path(result["document_path"]).parent)

    def clearance_reset_template_clicked(self) -> None:
        self.start_task("重新生成通关模板", self.ops.clearance_reset_template, self.on_clearance_template_reset)

    def on_clearance_template_reset(self, result: dict[str, Any]) -> None:
        self.refresh_paths()
        self.log(f"通关模板已生成：{result['output_path']}")
        self.open_path(self.ops.clearance_dir)

    def clearance_capture_session_clicked(self) -> None:
        self.start_task("保存网站登录态", self.ops.clearance_capture_session, self.on_clearance_capture_session_done)

    def on_clearance_capture_session_done(self, result: dict[str, Any]) -> None:
        self.refresh_clearance_session_hint()
        self.log(f"登录态已保存：{result['session_path']}")
        self.log(f"用户：{result.get('user_label', '') or '未知'}")
        self.open_path(self.ops.clearance_site_session_dir)

    def clearance_check_session_clicked(self) -> None:
        self.start_task("检查网站登录态", self.ops.clearance_check_session, self.on_clearance_check_session_done)

    def on_clearance_check_session_done(self, result: dict[str, Any]) -> None:
        self.refresh_clearance_session_hint()
        self.log(f"登录态检查结果：valid={result.get('valid')}")
        self.open_path(self.ops.clearance_site_checks_dir)

    def clearance_query_one_clicked(self) -> None:
        identifier = self.clearance_query_var.get().strip()
        if not identifier:
            messagebox.showwarning("请输入内容", "请先输入提单号、报关单号或箱号。")
            return
        iemark = self.clearance_iemark_var.get().strip() or None
        mode = self.clearance_mode_var.get().strip() or "auto"
        self.start_task(
            "单票查询网站状态",
            lambda: self.ops.clearance_query_one(identifier, mode, iemark),
            self.on_clearance_query_one_done,
        )

    def on_clearance_query_one_done(self, result: dict[str, Any]) -> None:
        lines = [
            f"识别值：{result['identifier']}",
            f"模式：{result['mode']}",
            f"I/E：{result.get('iemark') or 'auto'}",
            f"是否已放行：{result['released']}",
            f"状态：{result.get('status_text') or 'N/A'}",
            f"报关单号：{result.get('entry_no') or 'N/A'}",
            f"提单号：{result.get('bl_no') or 'N/A'}",
            f"箱号：{result.get('ctnr_no') or 'N/A'}",
            f"PCS：{result.get('pcs') or 'N/A'}",
            f"毛重 kg：{result.get('gross_weight_kg') or 'N/A'}",
            f"放行时间：{result.get('release_time') or 'N/A'}",
        ]
        self.log("\n".join(lines))
        self.log(f"查询报告：{result.get('latest_query_txt', '')}")
        messagebox.showinfo("单票查询结果", "\n".join(lines))

    def clearance_update_clicked(self) -> None:
        self.start_task("整表自动回填", self.ops.clearance_update_from_current, self.on_clearance_update_done)

    def clearance_pick_other_excel_clicked(self) -> None:
        path = filedialog.askopenfilename(
            title="选择一个 Excel 工作簿",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialdir=str(self.ops.clearance_dir),
        )
        if not path:
            return
        self.start_task("整表自动回填", lambda: self.ops.clearance_update_from_file(path), self.on_clearance_update_done)

    def on_clearance_update_done(self, result: dict[str, Any]) -> None:
        summary = result["site_summary"]
        self.log(f"更新后的工作簿：{result['updated_workbook']}")
        self.log(f"固定入口：{result['latest_updated_workbook']}")
        self.log(
            "网站查询摘要：候选={candidate_bl_count}，已放行={released_count}，未放行={pending_count}，无数据={no_data_count}".format(
                **summary
            )
        )
        if self.config.auto_open_results:
            self.open_path(self.ops.clearance_updated_dir)

    def check_updates_clicked(self) -> None:
        self.start_task("检查更新", self.ops.check_for_updates, self.on_check_updates_done)

    def on_check_updates_done(self, result: dict[str, Any]) -> None:
        latest = result.get("latest_version") or "未知"
        self.last_update_var.set(f"当前版本：{APP_VERSION} | 最新版本：{latest}")
        self.log(result["message"])
        if result.get("update_available"):
            if result.get("install_supported"):
                if messagebox.askyesno(
                    "发现新版本",
                    f"当前版本：{APP_VERSION}\n最新版本：{latest}\n\n是否现在下载并自动安装更新？",
                ):
                    self.start_task("准备自动更新", self.ops.prepare_update_install, self.on_update_ready)
            else:
                if messagebox.askyesno(
                    "发现新版本",
                    f"当前版本：{APP_VERSION}\n最新版本：{latest}\n\n当前不是 release 安装版，是否打开下载页面？",
                ):
                    webbrowser.open(result["html_url"])
        else:
            messagebox.showinfo("检查更新", result["message"])

    def on_update_ready(self, result: dict[str, Any]) -> None:
        launcher_path = result["launcher_path"]
        self.log(f"自动更新包已准备：{result['zip_path']}")
        creation_flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
        subprocess.Popen(["cmd", "/c", launcher_path], creationflags=creation_flags)
        messagebox.showinfo(
            "开始更新",
            f"新版本 {result['version']} 已准备完成。\n应用现在会关闭，自动安装后重新打开。\n\n你的工作区、设置和登录态不会被删除。",
        )
        self.on_close()

    def on_close(self) -> None:
        try:
            if self.sync_panel is not None:
                self.sync_panel.shutdown()
        finally:
            self.window.destroy()

    def run(self) -> None:
        self.window.mainloop()


def main() -> None:
    app = AmsDesktopApp()
    app.run()


if __name__ == "__main__":
    main()
