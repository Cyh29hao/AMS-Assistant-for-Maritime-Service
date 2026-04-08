from __future__ import annotations

import queue
import threading
import traceback
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from typing import Any, Callable

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, END, LEFT, RIGHT, X, Y

from desktop_app.runtime import (
    APP_NAME,
    APP_REPO_URL,
    APP_VERSION,
    AmsOperations,
    AppConfig,
    ConfigStore,
    open_in_file_explorer,
    release_assets_dir,
)


ThemeChoices = [
    "flatly",
    "litera",
    "minty",
    "pulse",
    "sandstone",
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
        self.window.geometry("1260x860")
        self.window.minsize(1140, 760)

        self.status_text = ttk.StringVar(value="准备就绪")
        self.req1_input_var = ttk.StringVar()
        self.req1_result_var = ttk.StringVar()
        self.req2_input_var = ttk.StringVar()
        self.req2_result_var = ttk.StringVar()
        self.req2_query_var = ttk.StringVar()
        self.req2_iemark_var = ttk.StringVar(value="")
        self.req2_mode_var = ttk.StringVar(value="auto")
        self.req2_session_hint_var = ttk.StringVar(value="尚未检查 req2 登录态")
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

        self.build_ui()
        self.refresh_paths()
        self.refresh_req2_session_hint()
        self.window.after(180, self.process_task_queue)

        if self.config.check_updates_on_launch:
            self.start_task("检查更新", self.ops.check_for_updates, self.on_check_updates_done)

    def build_ui(self) -> None:
        outer = ttk.Frame(self.window, padding=18)
        outer.pack(fill=BOTH, expand=True)

        self.build_header(outer)

        notebook = ttk.Notebook(outer, bootstyle="primary")
        notebook.pack(fill=BOTH, expand=True, pady=(10, 10))

        self.home_tab = ttk.Frame(notebook, padding=16)
        self.req1_tab = ttk.Frame(notebook, padding=16)
        self.req2_tab = ttk.Frame(notebook, padding=16)
        self.req3_tab = ttk.Frame(notebook, padding=16)
        self.settings_tab = ttk.Frame(notebook, padding=16)

        notebook.add(self.home_tab, text="首页")
        notebook.add(self.req1_tab, text="Req1 出合同")
        notebook.add(self.req2_tab, text="Req2 查通关")
        notebook.add(self.req3_tab, text="Req3 入口")
        notebook.add(self.settings_tab, text="设置")

        self.build_home_tab()
        self.build_req1_tab()
        self.build_req2_tab()
        self.build_req3_tab()
        self.build_settings_tab()

        log_frame = ttk.Labelframe(outer, text="运行日志", padding=10, bootstyle="info")
        log_frame.pack(fill=BOTH, expand=False)
        self.log_widget = ScrolledText(log_frame, height=10, font=("Microsoft YaHei UI", 10))
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
            font=("Microsoft YaHei UI", 22, "bold"),
            bootstyle="primary",
        ).pack(anchor="w")
        ttk.Label(
            left,
            text="尽量让普通用户只通过这个本地应用完成 req1、req2，并为 req3 预留入口。",
            font=("Microsoft YaHei UI", 11),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(4, 0))

        right = ttk.Frame(header)
        right.pack(side=RIGHT)
        ttk.Button(right, text="打开工作区", bootstyle="secondary", command=self.open_workspace_root).pack(side=LEFT, padx=6)
        ttk.Button(right, text="打开用户数据", bootstyle="secondary", command=self.open_settings_root).pack(side=LEFT, padx=6)
        ttk.Button(right, text="打开说明", bootstyle="secondary", command=self.open_user_guide).pack(side=LEFT, padx=6)
        ttk.Button(right, text="检查更新", bootstyle="info", command=self.check_updates_clicked).pack(side=LEFT, padx=6)
        ttk.Button(right, text="GitHub", bootstyle="link", command=lambda: webbrowser.open(APP_REPO_URL)).pack(side=LEFT, padx=6)

    def build_home_tab(self) -> None:
        intro = self.make_card(self.home_tab, "欢迎", "现在你可以主要通过这个桌面界面操作 req1 和 req2。")
        intro.pack(fill=X, pady=(0, 12))
        ttk.Label(intro, text="推荐顺序：先在固定 Excel 里填数据，再回到这里点击一键执行。", bootstyle="secondary").pack(anchor="w", pady=(6, 0))
        ttk.Label(intro, text="所有用户数据都会落在工作区和用户数据目录里，更新应用后不会重置。", bootstyle="secondary").pack(anchor="w", pady=(2, 0))

        quick = self.make_card(self.home_tab, "快速入口", "")
        quick.pack(fill=X, pady=(0, 12))
        button_row = ttk.Frame(quick)
        button_row.pack(fill=X, pady=(4, 0))
        ttk.Button(button_row, text="打开 Req1 输入 Excel", bootstyle="primary", command=lambda: self.open_path(self.ops.req1_input_path)).pack(side=LEFT, padx=4)
        ttk.Button(button_row, text="打开 Req2 输入 Excel", bootstyle="primary", command=lambda: self.open_path(self.ops.req2_input_path)).pack(side=LEFT, padx=4)
        ttk.Button(button_row, text="打开 Req1 结果", bootstyle="success", command=lambda: self.open_path(self.ops.req1_result_dir)).pack(side=LEFT, padx=4)
        ttk.Button(button_row, text="打开 Req2 结果", bootstyle="success", command=lambda: self.open_path(self.ops.req2_result_root)).pack(side=LEFT, padx=4)
        ttk.Button(button_row, text="打开新手说明", bootstyle="info", command=self.open_user_guide).pack(side=LEFT, padx=4)
        button_row2 = ttk.Frame(quick)
        button_row2.pack(fill=X, pady=(8, 0))
        ttk.Button(button_row2, text="打开最新合同", bootstyle="secondary", command=self.open_latest_req1_document).pack(side=LEFT, padx=4)
        ttk.Button(button_row2, text="打开最新 Req2 工作簿", bootstyle="secondary", command=self.open_latest_req2_workbook).pack(side=LEFT, padx=4)
        ttk.Button(button_row2, text="打开最新 Req2 网站报告", bootstyle="secondary", command=self.open_latest_req2_site_report).pack(side=LEFT, padx=4)

        status_card = self.make_card(self.home_tab, "当前状态", "")
        status_card.pack(fill=X)
        ttk.Label(status_card, textvariable=self.last_update_var, bootstyle="secondary").pack(anchor="w", pady=(4, 0))
        ttk.Label(status_card, text="当前工作区：", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(10, 0))
        ttk.Label(status_card, textvariable=self.workspace_var, bootstyle="secondary").pack(anchor="w")
        ttk.Label(status_card, text="设置文件：", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(status_card, textvariable=self.settings_path_var, bootstyle="secondary").pack(anchor="w")
        ttk.Label(status_card, textvariable=self.req2_session_hint_var, bootstyle="secondary").pack(anchor="w", pady=(8, 0))

    def build_req1_tab(self) -> None:
        top = self.make_card(
            self.req1_tab,
            "Req1 系统出合同",
            "推荐普通用户这样做：打开固定 Excel，填数据，保存并关闭，然后点击“一键生成合同”。",
        )
        top.pack(fill=X, pady=(0, 12))

        ttk.Label(top, text="当前输入文件：", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(top, textvariable=self.req1_input_var, bootstyle="secondary").pack(anchor="w")
        ttk.Label(top, text="当前结果目录：", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(top, textvariable=self.req1_result_var, bootstyle="secondary").pack(anchor="w")

        row1 = ttk.Frame(top)
        row1.pack(fill=X, pady=(10, 0))
        ttk.Button(row1, text="打开输入 Excel", bootstyle="primary", command=lambda: self.open_path(self.ops.req1_input_path)).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="重置空白模板", bootstyle="warning", command=self.req1_reset_template_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="一键生成合同", bootstyle="success", command=self.req1_generate_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="打开结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.req1_result_dir)).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="打开最新合同", bootstyle="info", command=self.open_latest_req1_document).pack(side=LEFT, padx=4)

        row2 = ttk.Frame(top)
        row2.pack(fill=X, pady=(8, 0))
        ttk.Button(row2, text="选择别的 Excel 并生成", bootstyle="info", command=self.req1_pick_other_excel_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row2, text="打开 Req1 工作区", bootstyle="secondary", command=lambda: self.open_path(self.ops.req1_dir)).pack(side=LEFT, padx=4)

    def build_req2_tab(self) -> None:
        top = self.make_card(
            self.req2_tab,
            "Req2 自动查通关",
            "这里会把“登录网站、检查登录态、单票查询、整表自动回填”都收进一个界面里。",
        )
        top.pack(fill=X, pady=(0, 12))

        ttk.Label(top, text="当前 Req2 输入文件：", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(top, textvariable=self.req2_input_var, bootstyle="secondary").pack(anchor="w")
        ttk.Label(top, text="当前 Req2 结果目录：", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Label(top, textvariable=self.req2_result_var, bootstyle="secondary").pack(anchor="w")
        ttk.Label(top, textvariable=self.req2_session_hint_var, bootstyle="secondary").pack(anchor="w", pady=(8, 0))

        row1 = ttk.Frame(top)
        row1.pack(fill=X, pady=(10, 0))
        ttk.Button(row1, text="打开 Req2 输入 Excel", bootstyle="primary", command=lambda: self.open_path(self.ops.req2_input_path)).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="重置 Req2 空白模板", bootstyle="warning", command=self.req2_reset_template_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row1, text="打开 Req2 结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.req2_result_root)).pack(side=LEFT, padx=4)

        session_card = self.make_card(self.req2_tab, "网页登录层", "先保存一次登录态，后面再检查登录态、做查询和回填。")
        session_card.pack(fill=X, pady=(0, 12))
        session_row = ttk.Frame(session_card)
        session_row.pack(fill=X, pady=(8, 0))
        ttk.Button(session_row, text="首次登录并保存登录态", bootstyle="primary", command=self.req2_capture_session_clicked).pack(side=LEFT, padx=4)
        ttk.Button(session_row, text="检查登录态", bootstyle="info", command=self.req2_check_session_clicked).pack(side=LEFT, padx=4)
        ttk.Button(session_row, text="打开登录态目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.req2_site_session_dir)).pack(side=LEFT, padx=4)

        query_card = self.make_card(self.req2_tab, "单票测试", "先单独查一票，确认登录和网站查询都正常。")
        query_card.pack(fill=X, pady=(0, 12))
        query_controls = ttk.Frame(query_card)
        query_controls.pack(fill=X, pady=(8, 0))
        ttk.Entry(query_controls, textvariable=self.req2_query_var, width=42).pack(side=LEFT, padx=4)
        ttk.Combobox(query_controls, textvariable=self.req2_mode_var, values=["auto", "blNo", "entryNo", "ctnrNo"], width=12, state="readonly").pack(side=LEFT, padx=4)
        ttk.Combobox(query_controls, textvariable=self.req2_iemark_var, values=["", "E", "I"], width=8, state="readonly").pack(side=LEFT, padx=4)
        ttk.Button(query_controls, text="查询这一票", bootstyle="success", command=self.req2_query_one_clicked).pack(side=LEFT, padx=4)
        ttk.Button(query_controls, text="打开查询结果目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.req2_site_query_dir)).pack(side=LEFT, padx=4)

        update_card = self.make_card(self.req2_tab, "整表自动回填", "填好 Req2 输入 Excel 后，直接自动查询网站并更新工作簿。")
        update_card.pack(fill=X)
        update_row = ttk.Frame(update_card)
        update_row.pack(fill=X, pady=(8, 0))
        ttk.Button(update_row, text="用当前 Req2 Excel 自动回填", bootstyle="success", command=self.req2_update_clicked).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="选择别的 Req2 Excel", bootstyle="info", command=self.req2_pick_other_excel_clicked).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="打开更新后 Excel 目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.req2_updated_dir)).pack(side=LEFT, padx=4)
        ttk.Button(update_row, text="打开最新更新工作簿", bootstyle="info", command=self.open_latest_req2_workbook).pack(side=LEFT, padx=4)

    def build_req3_tab(self) -> None:
        card = self.make_card(
            self.req3_tab,
            "Req3 入口预留",
            "这里预留给后续的船期表自动化。先把入口和工作区位置固定下来，后续直接往里接功能。",
        )
        card.pack(fill=X)
        ttk.Label(card, text=f"当前 Req3 目录：{self.ops.req3_dir}", bootstyle="secondary").pack(anchor="w", pady=(8, 0))
        ttk.Button(card, text="打开 Req3 目录", bootstyle="secondary", command=lambda: self.open_path(self.ops.req3_dir)).pack(anchor="w", pady=(8, 0))
        ttk.Label(card, text="后续计划接入：读取 LINE UP、查询网站、生成港口 x 区域矩阵。", bootstyle="secondary").pack(anchor="w", pady=(10, 0))

    def build_settings_tab(self) -> None:
        card = self.make_card(
            self.settings_tab,
            "设置",
            "这里主要控制工作区、主题和一些使用偏好。工作区会保存在用户目录里，更新应用时不会被重置。",
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

        row2b = ttk.Frame(card)
        row2b.pack(fill=X, pady=(10, 0))
        ttk.Label(row2b, text="Req2 浏览器", width=14).pack(side=LEFT)
        ttk.Combobox(row2b, textvariable=self.req2_browser_var, values=BrowserChoices, width=18, state="readonly").pack(side=LEFT, padx=6)
        ttk.Label(row2b, text="`auto` 会优先尝试 Edge，再尝试 Chrome，最后尝试 Playwright 自带 Chromium。", bootstyle="secondary").pack(side=LEFT, padx=6)

        row3 = ttk.Frame(card)
        row3.pack(fill=X, pady=(10, 0))
        ttk.Checkbutton(row3, text="生成后自动打开结果目录", variable=self.auto_open_var, bootstyle="round-toggle").pack(side=LEFT, padx=4)
        ttk.Checkbutton(row3, text="启动时自动检查更新", variable=self.auto_update_check_var, bootstyle="round-toggle").pack(side=LEFT, padx=16)

        row4 = ttk.Frame(card)
        row4.pack(fill=X, pady=(12, 0))
        ttk.Button(row4, text="保存设置", bootstyle="success", command=self.save_settings_clicked).pack(side=LEFT, padx=4)
        ttk.Button(row4, text="打开当前工作区", bootstyle="secondary", command=self.open_workspace_root).pack(side=LEFT, padx=4)
        ttk.Button(row4, text="打开用户数据目录", bootstyle="secondary", command=self.open_settings_root).pack(side=LEFT, padx=4)

        row5 = ttk.Frame(card)
        row5.pack(fill=X, pady=(12, 0))
        ttk.Label(row5, text="当前设置文件：", width=14).pack(side=LEFT)
        ttk.Label(row5, textvariable=self.settings_path_var, bootstyle="secondary").pack(side=LEFT)

    def make_card(self, parent: ttk.Frame, title: str, description: str) -> ttk.Labelframe:
        card = ttk.Labelframe(parent, text=title, padding=12, bootstyle="light")
        if description:
            ttk.Label(card, text=description, bootstyle="secondary").pack(anchor="w")
        return card

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
                if kind == "log":
                    self.log(payload)
                elif kind == "success":
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
            messagebox.showinfo("请稍等", "当前还有一个任务在运行，请稍等它完成。")
            return

        self.task_running = True
        self.set_status(f"{title}进行中…")
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
        self.req1_input_var.set(str(self.ops.req1_input_path))
        self.req1_result_var.set(str(self.ops.req1_result_dir))
        self.req2_input_var.set(str(self.ops.req2_input_path))
        self.req2_result_var.set(str(self.ops.req2_result_root))
        self.workspace_var.set(str(self.ops.workspace_root))

    def refresh_req2_session_hint(self) -> None:
        session_file = self.ops.req2_site_session_dir / "req2_site_session.json"
        if session_file.exists():
            self.req2_session_hint_var.set(f"已存在 req2 登录态：{session_file}")
        else:
            self.req2_session_hint_var.set("还没有保存 req2 登录态")

    def open_path(self, path: Path) -> None:
        try:
            open_in_file_explorer(path)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def open_workspace_root(self) -> None:
        self.open_path(self.ops.workspace_root)

    def open_settings_root(self) -> None:
        self.open_path(self.config_store.config_path.parent)

    def open_user_guide(self) -> None:
        guide = release_assets_dir() / "应用使用说明.html"
        if guide.exists():
            self.open_path(guide)
            return
        messagebox.showinfo("说明文件缺失", f"没有找到说明文件：\n{guide}")

    def open_latest_req1_document(self) -> None:
        path = self.ops.req1_result_dir / "00-最新合同.docx"
        self.open_if_exists(path, "还没有找到最新合同，请先生成一次 Req1 合同。")

    def open_latest_req2_workbook(self) -> None:
        path = self.ops.req2_updated_dir / "00-最新更新工作簿.xlsx"
        self.open_if_exists(path, "还没有找到最新的 Req2 更新工作簿，请先执行一次整表自动回填。")

    def open_latest_req2_site_report(self) -> None:
        path = self.ops.req2_site_query_dir / "00-最新网站报告.txt"
        self.open_if_exists(path, "还没有找到最新的 Req2 网站报告，请先执行一次整表自动回填或单票查询。")

    def open_if_exists(self, path: Path, missing_message: str) -> None:
        if path.exists():
            self.open_path(path)
            return
        messagebox.showinfo("暂时没有结果", missing_message)

    def humanize_error_text(self, error_text: str) -> str:
        lowered = error_text.lower()
        if "permissionerror" in lowered:
            return "文件可能正在被 Excel 或 Word 占用。请先关闭相关文件，再重试。\n\n原始信息：\n" + error_text
        if "filenotfounderror" in lowered:
            return "需要的文件没有找到，可能被移动、重命名，或者还没有生成。\n\n原始信息：\n" + error_text
        if "sessionmissingerror" in lowered or "no saved req2 site session" in lowered:
            return "还没有可用的 Req2 网站登录态。请先点击“首次登录并保存登录态”。\n\n原始信息：\n" + error_text
        if "无法启动 req2 登录浏览器" in error_text or "browser" in lowered:
            return "Req2 登录浏览器没有正常启动。请先确认这台电脑能打开 Edge 或 Chrome；如果不行，再到设置里切换 Req2 浏览器。\n\n原始信息：\n" + error_text
        if "contractvalidationerror" in lowered or "clearancevalidationerror" in lowered:
            return "输入内容没有通过校验。请按照提示检查 Excel 里缺失或格式不对的字段。\n\n原始信息：\n" + error_text
        if "requestexception" in lowered or "site request failed" in lowered:
            return "网站请求失败了。可能是网络问题、网站暂时不可用，或者登录态已经失效。\n\n原始信息：\n" + error_text
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
        self.refresh_req2_session_hint()
        self.set_status("设置已保存")
        self.log(f"设置已保存。工作区：{self.ops.workspace_root}")
        messagebox.showinfo("设置已保存", "新的设置已经保存并立即生效。")

    def req1_reset_template_clicked(self) -> None:
        self.start_task("重置 Req1 模板", self.ops.req1_reset_template, self.on_req1_template_reset)

    def on_req1_template_reset(self, result: dict[str, Any]) -> None:
        self.refresh_paths()
        self.log(f"Req1 模板已生成：{result['output_path']}")
        if self.config.auto_open_results:
            self.open_path(self.ops.req1_dir)

    def req1_generate_clicked(self) -> None:
        self.start_task("生成 Req1 合同", self.ops.req1_generate_from_current, self.on_req1_generated)

    def req1_pick_other_excel_clicked(self) -> None:
        path = filedialog.askopenfilename(
            title="选择一个 Req1 Excel 工作簿",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialdir=str(self.ops.req1_dir),
        )
        if not path:
            return
        self.start_task(
            "生成 Req1 合同",
            lambda: self.ops.req1_generate_from_file(path),
            self.on_req1_generated,
        )

    def on_req1_generated(self, result: dict[str, Any]) -> None:
        self.log(f"合同已生成：{result['document_path']}")
        self.log(f"摘要已生成：{result['summary_path']}")
        self.log(f"固定入口：{result['latest_document_path']}")
        if self.config.auto_open_results:
            self.open_path(Path(result["document_path"]).parent)

    def req2_reset_template_clicked(self) -> None:
        self.start_task("重置 Req2 模板", self.ops.req2_reset_template, self.on_req2_template_reset)

    def on_req2_template_reset(self, result: dict[str, Any]) -> None:
        self.refresh_paths()
        self.log(f"Req2 模板已生成：{result['output_path']}")
        self.open_path(self.ops.req2_dir)

    def req2_capture_session_clicked(self) -> None:
        self.start_task("保存 Req2 登录态", self.ops.req2_capture_session, self.on_req2_capture_session_done)

    def on_req2_capture_session_done(self, result: dict[str, Any]) -> None:
        self.refresh_req2_session_hint()
        self.log(f"登录态已保存：{result['session_path']}")
        self.log(f"用户：{result.get('user_label', '') or '未知'}")
        self.open_path(self.ops.req2_site_session_dir)

    def req2_check_session_clicked(self) -> None:
        self.start_task("检查 Req2 登录态", self.ops.req2_check_session, self.on_req2_check_session_done)

    def on_req2_check_session_done(self, result: dict[str, Any]) -> None:
        self.refresh_req2_session_hint()
        self.log(f"Req2 登录态检查结果：valid={result.get('valid')}")
        self.log(f"检查报告：{self.ops.req2_site_checks_dir}")
        self.open_path(self.ops.req2_site_checks_dir)

    def req2_query_one_clicked(self) -> None:
        identifier = self.req2_query_var.get().strip()
        if not identifier:
            messagebox.showwarning("请输入内容", "请先输入提单号、报关单号或箱号。")
            return
        iemark = self.req2_iemark_var.get().strip() or None
        mode = self.req2_mode_var.get().strip() or "auto"
        self.start_task(
            "单票查询 Req2 网站",
            lambda: self.ops.req2_query_one(identifier, mode, iemark),
            self.on_req2_query_one_done,
        )

    def on_req2_query_one_done(self, result: dict[str, Any]) -> None:
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

    def req2_update_clicked(self) -> None:
        self.start_task("Req2 整表自动回填", self.ops.req2_update_from_current, self.on_req2_update_done)

    def req2_pick_other_excel_clicked(self) -> None:
        path = filedialog.askopenfilename(
            title="选择一个 Req2 Excel 工作簿",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialdir=str(self.ops.req2_dir),
        )
        if not path:
            return
        self.start_task(
            "Req2 整表自动回填",
            lambda: self.ops.req2_update_from_file(path),
            self.on_req2_update_done,
        )

    def on_req2_update_done(self, result: dict[str, Any]) -> None:
        summary = result["site_summary"]
        self.log(f"更新后工作簿：{result['updated_workbook']}")
        self.log(f"固定入口：{result['latest_updated_workbook']}")
        self.log(
            "网站查询摘要：candidates={candidate_bl_count}, released={released_count}, pending={pending_count}, no_data={no_data_count}".format(
                **summary
            )
        )
        if self.config.auto_open_results:
            self.open_path(self.ops.req2_updated_dir)

    def check_updates_clicked(self) -> None:
        self.start_task("检查更新", self.ops.check_for_updates, self.on_check_updates_done)

    def on_check_updates_done(self, result: dict[str, Any]) -> None:
        latest = result.get("latest_version") or "未知"
        self.last_update_var.set(f"当前版本：{APP_VERSION} | 线上版本：{latest}")
        self.log(result["message"])
        if result.get("update_available"):
            if messagebox.askyesno(
                "发现新版本",
                f"当前版本：{APP_VERSION}\n线上版本：{latest}\n\n是否打开下载页面？",
            ):
                webbrowser.open(result["html_url"])
        else:
            messagebox.showinfo("检查更新", result["message"])

    def run(self) -> None:
        self.window.mainloop()


def main() -> None:
    app = AmsDesktopApp()
    app.run()


if __name__ == "__main__":
    main()
