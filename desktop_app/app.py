from __future__ import annotations

import queue
import subprocess
import threading
import traceback
import webbrowser
from pathlib import Path
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
        self.window.title(f"{APP_NAME} Desktop")
        self.window.geometry("1360x920")
        self.window.minsize(1180, 820)
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

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

        self.build_ui()
        self.refresh_paths()
        self.refresh_clearance_session_hint()
        self.window.after(180, self.process_task_queue)

        if self.config.check_updates_on_launch:
            self.start_task("检查更新", self.ops.check_for_updates, self.on_check_updates_done)

    def build_ui(self) -> None:
        outer = ttk.Frame(self.window, padding=18)
        outer.pack(fill=BOTH, expand=True)

        self.build_header(outer)

        self.notebook = ttk.Notebook(outer, bootstyle="primary")
        self.notebook.pack(fill=BOTH, expand=True, pady=(12, 10))

        self.home_tab = ttk.Frame(self.notebook, padding=16)
        self.contract_tab = ttk.Frame(self.notebook, padding=16)
        self.clearance_tab = ttk.Frame(self.notebook, padding=16)
        self.lineup_tab = ttk.Frame(self.notebook, padding=16)
        self.sync_tab = ttk.Frame(self.notebook, padding=16)
        self.settings_tab = ttk.Frame(self.notebook, padding=16)

        self.notebook.add(self.home_tab, text="首页")
        self.notebook.add(self.contract_tab, text=CONTRACT_FEATURE_NAME)
        self.notebook.add(self.clearance_tab, text=CLEARANCE_FEATURE_NAME)
        self.notebook.add(self.lineup_tab, text=LINEUP_FEATURE_NAME)
        self.notebook.add(self.sync_tab, text=SYNC_FEATURE_NAME)
        self.notebook.add(self.settings_tab, text="设置")

        self.build_home_tab()
        self.build_contract_tab()
        self.build_clearance_tab()
        self.build_lineup_tab()
        self.build_sync_tab()
        self.build_settings_tab()

        log_frame = ttk.Labelframe(outer, text="运行日志", padding=10, bootstyle="secondary")
        log_frame.pack(fill=BOTH, expand=False)
        self.log_widget = ScrolledText(log_frame, height=9, font=("Microsoft YaHei UI", 10))
        self.log_widget.pack(fill=BOTH, expand=True)
        self.log("应用已启动。")

        status_bar = ttk.Label(
            outer,
            textvariable=self.status_text,
            anchor="w",
            padding=(8, 6),
            bootstyle="secondary",
        )
        status_bar.pack(fill=X, pady=(8, 0))

    def build_header(self, parent: ttk.Frame) -> None:
        header = ttk.Frame(parent)
        header.pack(fill=X)

        left = ttk.Frame(header)
        left.pack(side=LEFT, fill=X, expand=True)
        ttk.Label(
            left,
            text="AMS Assistant Desktop",
            font=("Microsoft YaHei UI", 24, "bold"),
            bootstyle="primary",
        ).pack(anchor="w")
        ttk.Label(
            left,
            text="把合同生成、通关查询、船期预留和 Excel 自动同步，收进同一个更容易上手的本地桌面工作台。",
            font=("Microsoft YaHei UI", 11),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(4, 0))

        right = ttk.Frame(header)
        right.pack(side=RIGHT)
        ttk.Button(right, text="打开工作区", bootstyle="secondary", command=self.open_workspace_root).pack(side=LEFT, padx=4)
        ttk.Button(right, text="帮助中心", bootstyle="info", command=self.open_help_center).pack(side=LEFT, padx=4)
        ttk.Button(right, text="反馈建议", bootstyle="success", command=self.open_feedback_mail).pack(side=LEFT, padx=4)
        ttk.Button(right, text="检查更新", bootstyle="warning", command=self.check_updates_clicked).pack(side=LEFT, padx=4)
        ttk.Button(right, text="GitHub", bootstyle="link", command=lambda: webbrowser.open(APP_REPO_URL)).pack(side=LEFT, padx=4)

    def build_home_tab(self) -> None:
        hero = self.make_card(
            self.home_tab,
            "欢迎使用 AMS Assistant",
            "如果你只是想做事，不想研究代码和文件结构，那就从这里开始。",
            bootstyle="primary",
        )
        hero.pack(fill=X, pady=(0, 12))
        ttk.Label(
            hero,
            text="建议使用顺序：先打开对应功能的输入文件，填好并保存，再回到应用里点击执行。结果、日志和帮助页面都尽量固定在一个地方。",
            bootstyle="inverse-primary",
            wraplength=1050,
        ).pack(anchor="w", pady=(6, 0))
        hero_buttons = ttk.Frame(hero)
        hero_buttons.pack(fill=X, pady=(10, 0))
        ttk.Button(hero_buttons, text=CONTRACT_FEATURE_NAME, bootstyle="light", command=lambda: self.select_tab(self.contract_tab)).pack(side=LEFT, padx=4)
        ttk.Button(hero_buttons, text=CLEARANCE_FEATURE_NAME, bootstyle="light", command=lambda: self.select_tab(self.clearance_tab)).pack(side=LEFT, padx=4)
        ttk.Button(hero_buttons, text=SYNC_FEATURE_NAME, bootstyle="light", command=lambda: self.select_tab(self.sync_tab)).pack(side=LEFT, padx=4)
        ttk.Button(hero_buttons, text="打开帮助中心", bootstyle="info", command=self.open_help_center).pack(side=LEFT, padx=4)

        quick_row = ttk.Frame(self.home_tab)
        quick_row.pack(fill=X, pady=(0, 12))
        quick_row.columnconfigure((0, 1, 2, 3), weight=1)
        self.make_feature_card(
            quick_row,
            0,
            CONTRACT_FEATURE_NAME,
            "填固定 Excel，一键生成 Word 合同。",
            "打开输入文件",
            lambda: self.open_path(self.ops.contract_input_path),
            "功能说明",
            lambda: self.open_help_page("contract"),
            "success",
        )
        self.make_feature_card(
            quick_row,
            1,
            CLEARANCE_FEATURE_NAME,
            "保存登录态后，查询网站并回填工作簿。",
            "打开输入文件",
            lambda: self.open_path(self.ops.clearance_input_path),
            "功能说明",
            lambda: self.open_help_page("clearance"),
            "info",
        )
        self.make_feature_card(
            quick_row,
            2,
            LINEUP_FEATURE_NAME,
            "后续接入船期表和港区矩阵生成。",
            "查看预留目录",
            lambda: self.open_path(self.ops.lineup_dir),
            "功能说明",
            lambda: self.open_help_page("lineup"),
            "warning",
        )
        self.make_feature_card(
            quick_row,
            3,
            SYNC_FEATURE_NAME,
            "复制源表到目标表，只排除少数不保留列。",
            "打开示例文件",
            lambda: self.open_path(self.ops.sync_examples_dir),
            "功能说明",
            lambda: self.open_help_page("sync"),
            "primary",
        )

        status_card = self.make_card(self.home_tab, "当前环境", "", bootstyle="secondary")
        status_card.pack(fill=X)
        ttk.Label(status_card, textvariable=self.last_update_var, bootstyle="secondary").pack(anchor="w", pady=(4, 0))
        ttk.Label(status_card, text=f"当前工作区：{self.ops.workspace_root}", bootstyle="secondary").pack(anchor="w", pady=(6, 0))
        ttk.Label(status_card, text=f"设置文件：{self.config_store.config_path}", bootstyle="secondary").pack(anchor="w", pady=(2, 0))
        ttk.Label(status_card, textvariable=self.clearance_session_hint_var, bootstyle="secondary").pack(anchor="w", pady=(2, 0))
        ttk.Label(
            status_card,
            text="如果你有改进建议，可以随时点右上角“反馈建议”，会直接打开给 cyh29hao@sjtu.edu.cn 的邮件。",
            bootstyle="secondary",
            wraplength=1020,
        ).pack(anchor="w", pady=(8, 0))

    def build_contract_tab(self) -> None:
        top = self.make_card(
            self.contract_tab,
            CONTRACT_FEATURE_NAME,
            "普通用户建议这样做：打开固定 Excel，填数据，保存并关闭，然后点击“一键生成合同”。",
            bootstyle="success",
        )
        top.pack(fill=X, pady=(0, 12))
        ttk.Label(top, text="当前输入文件", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(top, textvariable=self.contract_input_var, bootstyle="inverse-success").pack(anchor="w")
        ttk.Label(top, text="当前结果目录", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(top, textvariable=self.contract_result_var, bootstyle="inverse-success").pack(anchor="w")

        row1 = ttk.Frame(top)
        row1.pack(fill=X, pady=(10, 0))
        ttk.Button(row1, text="打开输入 Excel", bootstyle="light", command=lambda: self.open_path(self.ops.contract_input_path)).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="重新生成空白模板", bootstyle="warning", command=self.contract_reset_template_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="一键生成合同", bootstyle="success", command=self.contract_generate_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="打开结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.contract_result_dir)).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="打开最新合同", bootstyle="info", command=self.open_latest_contract_document).pack(side=LEFT, padx=4)

        row2 = ttk.Frame(top)
        row2.pack(fill=X, pady=(8, 0))
        ttk.Button(row2, text="选择别的 Excel 并生成", bootstyle="secondary", command=self.contract_pick_other_excel_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row2, text="打开功能说明", bootstyle="info", command=lambda: self.open_help_page("contract")).pack(side=LEFT, padx=4)

    def build_clearance_tab(self) -> None:
        top = self.make_card(
            self.clearance_tab,
            CLEARANCE_FEATURE_NAME,
            "这里把“保存登录态、检查登录、单票测试、整表回填”都收进一个界面里。",
            bootstyle="info",
        )
        top.pack(fill=X, pady=(0, 12))
        ttk.Label(top, text="当前输入文件", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(top, textvariable=self.clearance_input_var, bootstyle="inverse-info").pack(anchor="w")
        ttk.Label(top, text="当前结果目录", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(top, textvariable=self.clearance_result_var, bootstyle="inverse-info").pack(anchor="w")
        ttk.Label(top, textvariable=self.clearance_session_hint_var, bootstyle="inverse-info").pack(anchor="w", pady=(8, 0))

        row1 = ttk.Frame(top)
        row1.pack(fill=X, pady=(10, 0))
        ttk.Button(row1, text="打开输入 Excel", bootstyle="light", command=lambda: self.open_path(self.ops.clearance_input_path)).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="重新生成空白模板", bootstyle="warning", command=self.clearance_reset_template_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="打开结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.clearance_result_root)).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="打开功能说明", bootstyle="info", command=lambda: self.open_help_page("clearance")).pack(side=LEFT, padx=4)

        session_card = self.make_card(self.clearance_tab, "网站登录", "先保存一次登录态，后面再检查登录、做查询和回填。", bootstyle="secondary")
        session_card.pack(fill=X, pady=(0, 12))
        session_row = ttk.Frame(session_card)
        session_row.pack(fill=X, pady=(8, 0))
        ttk.Button(session_row, text="首次登录并保存登录态", bootstyle="primary", command=self.clearance_capture_session_clicked).pack(side=LEFT, padx=4)
        ttk.Button(session_row, text="检查登录态", bootstyle="info", command=self.clearance_check_session_clicked).pack(side=LEFT, padx=4)
        ttk.Button(session_row, text="打开登录态目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.clearance_site_session_dir)).pack(side=LEFT, padx=4)

        query_card = self.make_card(self.clearance_tab, "单票测试", "先查一票，确认登录和网站查询都正常。", bootstyle="secondary")
        query_card.pack(fill=X, pady=(0, 12))
        query_controls = ttk.Frame(query_card)
        query_controls.pack(fill=X, pady=(8, 0))
        ttk.Entry(query_controls, textvariable=self.clearance_query_var, width=42).pack(side=LEFT, padx=4)
        ttk.Combobox(query_controls, textvariable=self.clearance_mode_var, values=["auto", "blNo", "entryNo", "ctnrNo"], width=12, state="readonly").pack(side=LEFT, padx=4)
        ttk.Combobox(query_controls, textvariable=self.clearance_iemark_var, values=["", "E", "I"], width=8, state="readonly").pack(side=LEFT, padx=4)
        ttk.Button(query_controls, text="查询这一票", bootstyle="success", command=self.clearance_query_one_clicked).pack(side=LEFT, padx=4)
        ttk.Button(query_controls, text="打开查询结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.clearance_site_query_dir)).pack(side=LEFT, padx=4)

        update_card = self.make_card(self.clearance_tab, "整表自动回填", "填好输入 Excel 后，自动查询网站并更新工作簿。", bootstyle="secondary")
        update_card.pack(fill=X)
        update_row = ttk.Frame(update_card)
        update_row.pack(fill=X, pady=(8, 0))
        ttk.Button(update_row, text="用当前输入表自动回填", bootstyle="success", command=self.clearance_update_clicked).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="选择别的 Excel", bootstyle="secondary", command=self.clearance_pick_other_excel_clicked).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="打开最新工作簿", bootstyle="info", command=self.open_latest_clearance_workbook).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="打开最新网站报告", bootstyle="info", command=self.open_latest_clearance_site_report).pack(side=LEFT, padx=4)

    def build_lineup_tab(self) -> None:
        card = self.make_card(
            self.lineup_tab,
            LINEUP_FEATURE_NAME,
            "这个入口已经准备好，后续会接入船期表查询、港区矩阵生成和相关报告。",
            bootstyle="warning",
        )
        card.pack(fill=X)
        ttk.Label(card, text=f"当前预留目录：{self.ops.lineup_dir}", bootstyle="inverse-warning").pack(anchor="w", pady=(8, 0))
        row = ttk.Frame(card)
        row.pack(fill=X, pady=(10, 0))
        ttk.Button(row, text="打开预留目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.lineup_dir)).pack(side=LEFT, padx=4)
        ttk.Button(row, text="打开功能说明", bootstyle="info", command=lambda: self.open_help_page("lineup")).pack(side=LEFT, padx=4)
        ttk.Label(
            card,
            text="你以后验收这一块时，尽量也会从桌面应用里走，而不是回到脚本层单独折腾。",
            bootstyle="secondary",
            wraplength=980,
        ).pack(anchor="w", pady=(14, 0))

    def build_sync_tab(self) -> None:
        self.sync_panel = ExcelSyncPanel(
            parent=self.sync_tab,
            window=self.window,
            service=self.ops.build_sync_service(status_callback=lambda: None),
            open_path=self.open_path,
            open_help=self.open_help_page,
        )

    def build_settings_tab(self) -> None:
        card = self.make_card(
            self.settings_tab,
            "设置",
            "这里主要控制工作区、主题、通关查询浏览器偏好和自动更新设置。",
            bootstyle="secondary",
        )
        card.pack(fill=X)

        row1 = ttk.Frame(card)
        row1.pack(fill=X, pady=(10, 0))
        ttk.Label(row1, text="工作区根目录", width=14).pack(side=LEFT)
        ttk.Entry(row1, textvariable=self.workspace_var, width=78).pack(side=LEFT, padx=6)
        ttk.Button(row1, text="选择文件夹", bootstyle="secondary", command=self.choose_workspace_clicked).pack(side=LEFT)

        row2 = ttk.Frame(card)
        row2.pack(fill=X, pady=(10, 0))
        ttk.Label(row2, text="界面主题", width=14).pack(side=LEFT)
        ttk.Combobox(row2, textvariable=self.theme_var, values=ThemeChoices, width=18, state="readonly").pack(side=LEFT, padx=6)

        row3 = ttk.Frame(card)
        row3.pack(fill=X, pady=(10, 0))
        ttk.Label(row3, text="通关查询浏览器", width=14).pack(side=LEFT)
        ttk.Combobox(row3, textvariable=self.req2_browser_var, values=BrowserChoices, width=18, state="readonly").pack(side=LEFT, padx=6)
        ttk.Label(row3, text="推荐保持 auto，会优先尝试 Edge。", bootstyle="secondary").pack(side=LEFT, padx=6)

        row4 = ttk.Frame(card)
        row4.pack(fill=X, pady=(10, 0))
        ttk.Checkbutton(row4, text="执行后自动打开结果目录", variable=self.auto_open_var, bootstyle="round-toggle").pack(side=LEFT, padx=4)
        ttk.Checkbutton(row4, text="启动时自动检查更新", variable=self.auto_update_check_var, bootstyle="round-toggle").pack(side=LEFT, padx=16)

        row5 = ttk.Frame(card)
        row5.pack(fill=X, pady=(12, 0))
        ttk.Button(row5, text="保存设置", bootstyle="success", command=self.save_settings_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row5, text="打开工作区", bootstyle="secondary", command=self.open_workspace_root).pack(side=LEFT, padx=4)
        ttk.Button(row5, text="打开用户数据目录", bootstyle="secondary", command=self.open_settings_root).pack(side=LEFT, padx=4)
        ttk.Button(row5, text="帮助中心", bootstyle="info", command=self.open_help_center).pack(side=LEFT, padx=4)
        ttk.Button(row5, text="反馈建议", bootstyle="success", command=self.open_feedback_mail).pack(side=LEFT, padx=4)

        row6 = ttk.Frame(card)
        row6.pack(fill=X, pady=(12, 0))
        ttk.Label(row6, text="设置文件", width=14).pack(side=LEFT)
        ttk.Label(row6, textvariable=self.settings_path_var, bootstyle="secondary").pack(side=LEFT)

        row7 = ttk.Frame(card)
        row7.pack(fill=X, pady=(12, 0))
        update_mode = "支持自动安装更新" if portable_install_root() is not None else "源码模式：自动更新不可用"
        ttk.Label(row7, text=f"更新模式：{update_mode}", bootstyle="secondary").pack(side=LEFT)

    def make_card(self, parent: ttk.Frame, title: str, description: str, bootstyle: str = "light") -> ttk.Labelframe:
        card = ttk.Labelframe(parent, text=title, padding=14, bootstyle=bootstyle)
        if description:
            ttk.Label(card, text=description, bootstyle=f"inverse-{bootstyle}" if bootstyle != "light" else "secondary", wraplength=1040).pack(anchor="w")
        return card

    def make_feature_card(
        self,
        parent: ttk.Frame,
        column: int,
        title: str,
        description: str,
        primary_label: str,
        primary_command,
        secondary_label: str,
        secondary_command,
        bootstyle: str,
    ) -> None:
        card = ttk.Labelframe(parent, text=title, padding=14, bootstyle=bootstyle)
        card.grid(row=0, column=column, sticky="nsew", padx=6)
        ttk.Label(card, text=description, bootstyle=f"inverse-{bootstyle}", wraplength=210).pack(anchor="w")
        row = ttk.Frame(card)
        row.pack(fill=X, pady=(12, 0))
        ttk.Button(row, text=primary_label, bootstyle="light", command=primary_command).pack(side=LEFT, padx=2)
        ttk.Button(row, text=secondary_label, bootstyle="secondary", command=secondary_command).pack(side=LEFT, padx=2)

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
