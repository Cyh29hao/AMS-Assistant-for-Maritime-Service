from __future__ import annotations

import argparse
import queue
import sys
import threading
import traceback
from pathlib import Path
from tkinter import messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, LEFT, X

THIS_DIR = Path(__file__).resolve().parent
ROOT = THIS_DIR
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from desktop_app.runtime import APP_NAME, AmsOperations, AppConfig


class UpdateInstallerWindow:
    def __init__(self, manifest_path: Path) -> None:
        self.manifest_path = manifest_path
        self.ops = AmsOperations(AppConfig.default())
        self.queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.running = True
        self.debug_log_path = self.manifest_path.parent / "updater-debug.log"
        self.write_debug(f"START manifest={self.manifest_path}")

        self.window = ttk.Window(themename="flatly")
        self.window.title(f"{APP_NAME} Updater")
        self.window.geometry("560x280")
        self.window.minsize(560, 280)
        self.window.maxsize(700, 420)
        self.window.protocol("WM_DELETE_WINDOW", self.on_close_requested)

        self.progress_var = ttk.DoubleVar(value=0)
        self.percent_var = ttk.StringVar(value="0%")
        self.status_var = ttk.StringVar(value="正在准备更新…")
        self.detail_var = ttk.StringVar(value=str(self.manifest_path.parent))

        self.build_ui()
        self.window.after(120, self.process_queue)
        self.start_worker()

    def build_ui(self) -> None:
        outer = ttk.Frame(self.window, padding=(22, 20, 22, 20))
        outer.pack(fill=BOTH, expand=True)

        card = ttk.Frame(outer, padding=22, bootstyle="light", borderwidth=1, relief="solid")
        card.pack(fill=BOTH, expand=True)

        ttk.Label(
            card,
            text="AUTO UPDATE",
            bootstyle="inverse-primary",
            font=("Microsoft YaHei UI", 9, "bold"),
            padding=(10, 5),
        ).pack(anchor="w")
        ttk.Label(card, text=f"{APP_NAME} 正在更新", font=("Microsoft YaHei UI", 18, "bold"), bootstyle="primary").pack(anchor="w", pady=(12, 4))
        ttk.Label(
            card,
            text="这一步会等待旧版本退出、安装新版本，并自动重新打开应用。你的工作区、设置和登录态不会丢失。",
            bootstyle="secondary",
            wraplength=500,
            justify="left",
        ).pack(anchor="w")

        progress_row = ttk.Frame(card)
        progress_row.pack(fill=X, pady=(18, 8))
        ttk.Progressbar(
            progress_row,
            variable=self.progress_var,
            maximum=100,
            bootstyle="success-striped",
        ).pack(side=LEFT, fill=X, expand=True)
        ttk.Label(progress_row, textvariable=self.percent_var, font=("Microsoft YaHei UI", 10, "bold"), bootstyle="primary").pack(side=LEFT, padx=(12, 0))

        ttk.Label(card, textvariable=self.status_var, font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(10, 4))
        ttk.Label(card, textvariable=self.detail_var, bootstyle="secondary", wraplength=500, justify="left").pack(anchor="w")
        ttk.Label(
            card,
            text="请不要在这一步手动删除程序文件，也不用重复点击检查更新。",
            bootstyle="secondary",
            wraplength=500,
            justify="left",
        ).pack(anchor="w", pady=(18, 0))

    def start_worker(self) -> None:
        def worker() -> None:
            try:
                result = self.ops.apply_prepared_update(
                    self.manifest_path,
                    progress_callback=lambda payload: self.queue.put(("progress", payload)),
                )
            except Exception:
                self.queue.put(("error", traceback.format_exc()))
                return
            self.queue.put(("success", result))

        threading.Thread(target=worker, daemon=True).start()

    def process_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "progress":
                    progress = payload if isinstance(payload, dict) else {}
                    value = float(progress.get("value", 0))
                    self.progress_var.set(value)
                    self.percent_var.set(f"{value:.0f}%")
                    self.status_var.set(str(progress.get("message", "正在更新…")))
                    self.detail_var.set(str(progress.get("detail", "")))
                    self.write_debug(
                        "PROGRESS "
                        f"value={value:.1f} message={self.status_var.get()} detail={self.detail_var.get()}"
                    )
                elif kind == "success":
                    result = payload if isinstance(payload, dict) else {}
                    self.running = False
                    self.progress_var.set(100)
                    self.percent_var.set("100%")
                    self.status_var.set("更新完成，正在打开新版本。")
                    self.detail_var.set(str(result.get("restarted_executable", "")))
                    self.write_debug(f"SUCCESS result={result}")
                    self.window.after(1400, self.window.destroy)
                elif kind == "error":
                    self.running = False
                    error_text = str(payload)
                    self.status_var.set("更新失败")
                    self.detail_var.set("可以重试一次检查更新；如果仍失败，请保留当前文件夹并反馈日志。")
                    self.write_debug(f"ERROR {error_text}")
                    messagebox.showerror("自动更新失败", error_text)
        except queue.Empty:
            pass
        finally:
            if self.window.winfo_exists():
                self.window.after(120, self.process_queue)

    def on_close_requested(self) -> None:
        if self.running:
            self.write_debug("CLOSE_BLOCKED running=true")
            messagebox.showinfo("正在更新", "自动更新正在进行，请稍等片刻。")
            return
        self.write_debug("CLOSE_ALLOWED")
        self.window.destroy()

    def run(self) -> int:
        self.window.mainloop()
        self.write_debug("EXIT")
        return 0

    def write_debug(self, message: str) -> None:
        self.debug_log_path.parent.mkdir(parents=True, exist_ok=True)
        with self.debug_log_path.open("a", encoding="utf-8") as fh:
            fh.write(message + "\n")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--apply-update-manifest", required=True)
    return parser


def write_update_bootstrap_log(manifest_arg: str, stage: str) -> None:
    try:
        manifest_path = Path(manifest_arg).resolve()
        log_path = manifest_path.parent / "updater-bootstrap.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as fh:
            fh.write(f"{stage} argv={sys.argv}\n")
    except Exception:
        pass


def main() -> int:
    parser = build_parser()
    args, _unknown = parser.parse_known_args()
    write_update_bootstrap_log(args.apply_update_manifest, "ENTER_UPDATE_HELPER")
    return UpdateInstallerWindow(Path(args.apply_update_manifest).resolve()).run()


if __name__ == "__main__":
    raise SystemExit(main())
