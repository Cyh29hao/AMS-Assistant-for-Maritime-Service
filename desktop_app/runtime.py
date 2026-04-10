from __future__ import annotations

import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Callable

import requests

from desktop_app.excel_sync_engine import (
    DEFAULT_SETTINGS as DEFAULT_SYNC_SETTINGS,
    ExcelSyncPaths,
    SyncService,
    create_demo_workbooks,
)


APP_NAME = "AMS Assistant"
APP_VERSION = "0.4.2"
APP_REPO_URL = "https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service"
APP_RELEASE_API = "https://api.github.com/repos/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases/latest"
RELEASE_PACKAGE_PREFIX = "AMS-Assistant-Desktop-v"
UPDATER_EXE_NAME = "AMS-Assistant-Updater.exe"
UPDATE_TOP_LEVEL_FILES = [
    "README.html",
    "00-从这里开始.html",
    "00-先看这里.html",
    "VERSION.txt",
    "Start AMS Assistant.bat",
    "Open Workspace.bat",
    "Open User Data.bat",
    "Open Guide.bat",
    "Run Desktop Self Test.bat",
    "1-启动AMS桌面应用.bat",
    "2-运行桌面版自检.bat",
    "3-打开工作区.bat",
    "4-打开说明.bat",
    "5-打开用户数据目录.bat",
]

ENV_SETTINGS_DIR = "AMS_ASSISTANT_SETTINGS_DIR"
ENV_DEFAULT_WORKSPACE = "AMS_ASSISTANT_DEFAULT_WORKSPACE"
ENV_REQ2_BROWSER = "AMS_REQ2_BROWSER"

CONTRACT_FEATURE_NAME = "合同生成中心"
CLEARANCE_FEATURE_NAME = "通关查询中心"
LINEUP_FEATURE_NAME = "船期与港区矩阵"
SYNC_FEATURE_NAME = "表格自动同步"

LEGACY_CONTRACT_DIR = "req1-系统出合同"
LEGACY_CLEARANCE_DIR = "req2-自动查通关"
LEGACY_LINEUP_DIR = "req3-船期表"

CONTRACT_DIR_NAME = "合同生成"
CLEARANCE_DIR_NAME = "通关查询"
LINEUP_DIR_NAME = "船期与港区矩阵"
SYNC_DIR_NAME = "表格自动同步"


def user_home() -> Path:
    return Path.home().resolve()


def default_documents_dir() -> Path:
    candidate = user_home() / "Documents"
    return candidate if candidate.exists() else user_home()


def settings_root() -> Path:
    override = os.environ.get(ENV_SETTINGS_DIR, "").strip()
    if override:
        return Path(override).expanduser().resolve()
    appdata = os.environ.get("APPDATA")
    root = Path(appdata).resolve() if appdata else (user_home() / "AppData" / "Roaming")
    return root / APP_NAME


def settings_path() -> Path:
    return settings_root() / "settings.json"


def default_workspace_root() -> Path:
    override = os.environ.get(ENV_DEFAULT_WORKSPACE, "").strip()
    if override:
        return Path(override).expanduser().resolve()
    return default_documents_dir() / "AMS Assistant Workspace"


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS).resolve()
    return Path(__file__).resolve().parents[1]


def skill_root() -> Path:
    return runtime_base_dir() / "maritime-service"


def release_assets_dir() -> Path:
    return runtime_base_dir() / "desktop_app" / "release_assets"


def help_assets_dir() -> Path:
    return release_assets_dir() / "help"


def scripts_dir() -> Path:
    return skill_root() / "scripts"


def ensure_scripts_dir_on_path() -> None:
    path = str(scripts_dir())
    if path not in sys.path:
        sys.path.insert(0, path)


def load_source_module(module_name: str, file_path: Path):
    ensure_scripts_dir_on_path()
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load module: {module_name}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def open_in_file_explorer(target: Path) -> None:
    target = target.resolve()
    if sys.platform.startswith("win"):
        os.startfile(str(target))
        return
    raise RuntimeError("当前桌面版仅支持 Windows。")


def parse_version(version: str) -> tuple[int, ...]:
    normalized = re.sub(r"^[^0-9]+", "", version.strip())
    parts = [int(part) for part in normalized.split(".") if part.isdigit()]
    return tuple(parts) if parts else (0,)


def portable_install_root() -> Path | None:
    if not getattr(sys, "frozen", False):
        return None
    executable = Path(sys.executable).resolve()
    if executable.parent.name != "AMS-Assistant-Desktop":
        return None
    if executable.parent.parent.name != "app":
        return None
    return executable.parent.parent.parent


@dataclass
class ReleaseInfo:
    version: str
    html_url: str
    published_at: str
    asset_name: str
    asset_download_url: str
    asset_size: int = 0
    body: str = ""


@dataclass
class AppConfig:
    workspace_root: str
    theme_name: str = "flatly"
    auto_open_results: bool = True
    check_updates_on_launch: bool = True
    req1_input_filename: str = "contract-input.xlsx"
    req2_input_filename: str = "clearance-input.xlsx"
    req2_browser_preference: str = "auto"

    @classmethod
    def default(cls) -> "AppConfig":
        return cls(workspace_root=str(default_workspace_root()))


class ConfigStore:
    def __init__(self) -> None:
        self.path = settings_path()

    @property
    def config_path(self) -> Path:
        return self.path

    def load(self) -> AppConfig:
        if not self.path.exists():
            config = AppConfig.default()
            self.save(config)
            return config
        payload = json.loads(self.path.read_text(encoding="utf-8"))
        default = AppConfig.default()
        merged = asdict(default)
        merged.update({k: v for k, v in payload.items() if k in merged})
        config = AppConfig(**merged)
        self.save(config)
        return config

    def save(self, config: AppConfig) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps(asdict(config), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )


class AmsOperations:
    def __init__(self, config: AppConfig) -> None:
        self.config = config

    @property
    def workspace_root(self) -> Path:
        return Path(self.config.workspace_root).expanduser().resolve()

    @property
    def contract_dir(self) -> Path:
        return self.workspace_root / CONTRACT_DIR_NAME

    @property
    def contract_input_path(self) -> Path:
        return self.contract_dir / self.config.req1_input_filename

    @property
    def contract_result_dir(self) -> Path:
        return self.contract_dir / "result"

    @property
    def contract_summary_dir(self) -> Path:
        return self.contract_result_dir / "summary"

    @property
    def contract_backup_dir(self) -> Path:
        return self.contract_dir / "backup"

    @property
    def clearance_dir(self) -> Path:
        return self.workspace_root / CLEARANCE_DIR_NAME

    @property
    def clearance_input_path(self) -> Path:
        return self.clearance_dir / self.config.req2_input_filename

    @property
    def clearance_result_root(self) -> Path:
        return self.clearance_dir / "result"

    @property
    def clearance_root(self) -> Path:
        return self.clearance_result_root / "clearance"

    @property
    def clearance_updated_dir(self) -> Path:
        return self.clearance_root / "updated"

    @property
    def clearance_site_session_dir(self) -> Path:
        return self.clearance_root / "site_session"

    @property
    def clearance_site_checks_dir(self) -> Path:
        return self.clearance_root / "site_checks"

    @property
    def clearance_site_query_dir(self) -> Path:
        return self.clearance_root / "site_query_results"

    @property
    def lineup_dir(self) -> Path:
        return self.workspace_root / LINEUP_DIR_NAME

    @property
    def sync_dir(self) -> Path:
        return self.workspace_root / SYNC_DIR_NAME

    @property
    def sync_examples_dir(self) -> Path:
        return self.sync_dir / "示例文件"

    @property
    def sync_runtime_dir(self) -> Path:
        return self.sync_dir / "运行日志"

    @property
    def sync_settings_dir(self) -> Path:
        return settings_root() / "feature-data" / "excel-sync"

    @property
    def sync_tasks_path(self) -> Path:
        return self.sync_settings_dir / "sync-tasks.json"

    @property
    def sync_template_path(self) -> Path:
        return self.sync_settings_dir / "sync-tasks.template.json"

    def sync_paths(self) -> ExcelSyncPaths:
        return ExcelSyncPaths(
            data_path=self.sync_tasks_path,
            log_dir=self.sync_runtime_dir,
            template_data_path=self.sync_template_path,
        )

    def set_runtime_env(self) -> None:
        os.environ["AMS_DATA_ROOT"] = str(self.clearance_result_root)
        os.environ[ENV_REQ2_BROWSER] = self.config.req2_browser_preference

    def load_contract_module(self):
        return load_source_module("contract_workflow_runtime", scripts_dir() / "contract_workflow.py")

    def load_clearance_module(self):
        return load_source_module("clearance_workflow_runtime", scripts_dir() / "clearance_workflow.py")

    def load_site_module(self):
        self.set_runtime_env()
        if "clearance_workflow_runtime" in sys.modules:
            del sys.modules["clearance_workflow_runtime"]
        return load_source_module("clearance_site_workflow_runtime", scripts_dir() / "clearance_site_workflow.py")

    def build_sync_service(self, status_callback=None) -> SyncService:
        self.ensure_sync_storage()
        return SyncService(self.sync_paths(), status_callback=status_callback)

    def ensure_sync_storage(self) -> None:
        self.sync_settings_dir.mkdir(parents=True, exist_ok=True)
        if not self.sync_template_path.exists():
            self.sync_template_path.write_text(
                json.dumps({"settings": DEFAULT_SYNC_SETTINGS, "tasks": []}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

    def ensure_workspace(self) -> dict[str, Any]:
        self.workspace_root.mkdir(parents=True, exist_ok=True)
        self._migrate_legacy_dir(LEGACY_CONTRACT_DIR, self.contract_dir)
        self._migrate_legacy_dir(LEGACY_CLEARANCE_DIR, self.clearance_dir)
        self._migrate_legacy_dir(LEGACY_LINEUP_DIR, self.lineup_dir)

        self.contract_dir.mkdir(parents=True, exist_ok=True)
        self.contract_result_dir.mkdir(parents=True, exist_ok=True)
        self.contract_summary_dir.mkdir(parents=True, exist_ok=True)
        self.contract_backup_dir.mkdir(parents=True, exist_ok=True)

        self.clearance_dir.mkdir(parents=True, exist_ok=True)
        self.clearance_result_root.mkdir(parents=True, exist_ok=True)
        self.clearance_root.mkdir(parents=True, exist_ok=True)
        self.clearance_updated_dir.mkdir(parents=True, exist_ok=True)
        self.clearance_site_session_dir.mkdir(parents=True, exist_ok=True)
        self.clearance_site_checks_dir.mkdir(parents=True, exist_ok=True)
        self.clearance_site_query_dir.mkdir(parents=True, exist_ok=True)

        self.lineup_dir.mkdir(parents=True, exist_ok=True)
        self.sync_dir.mkdir(parents=True, exist_ok=True)
        self.sync_runtime_dir.mkdir(parents=True, exist_ok=True)
        self.ensure_sync_storage()

        self._write_workspace_notes()
        create_demo_workbooks(self.sync_examples_dir)

        if not self.contract_input_path.exists():
            self.contract_reset_template(backup_existing=False)
        if not self.clearance_input_path.exists():
            self.clearance_reset_template()

        return {
            "workspace_root": str(self.workspace_root),
            "contract_input_path": str(self.contract_input_path),
            "clearance_input_path": str(self.clearance_input_path),
            "sync_tasks_path": str(self.sync_tasks_path),
        }

    def contract_reset_template(self, backup_existing: bool = True) -> dict[str, Any]:
        contract = self.load_contract_module()
        self.contract_dir.mkdir(parents=True, exist_ok=True)
        self.contract_backup_dir.mkdir(parents=True, exist_ok=True)
        if backup_existing and self.contract_input_path.exists():
            backup_path = self.contract_backup_dir / f"contract-input-{self._timestamp()}.xlsx"
            shutil.copy2(self.contract_input_path, backup_path)
        output_path = contract.create_workbook_template(self.contract_input_path)
        return {"output_path": str(output_path)}

    def contract_generate_from_current(self) -> dict[str, Any]:
        return self.contract_generate_from_file(self.contract_input_path)

    def contract_generate_from_file(self, source_path: Path | str) -> dict[str, Any]:
        contract = self.load_contract_module()
        source_path = Path(source_path).expanduser().resolve()
        if not source_path.exists():
            raise FileNotFoundError(f"Excel file not found: {source_path}")
        request = contract.load_workbook_request(source_path)
        normalized = contract.validate_request(request, contract.load_registry())
        result = contract.render_contract(
            normalized,
            self.contract_result_dir,
            self.contract_summary_dir,
        )
        latest_document = self.contract_result_dir / "00-latest-contract.docx"
        latest_summary = self.contract_summary_dir / "00-latest-summary.json"
        self.copy_as_latest(result.document_path, latest_document)
        self.copy_as_latest(result.summary_path, latest_summary)
        return {
            "document_path": str(result.document_path),
            "summary_path": str(result.summary_path),
            "latest_document_path": str(latest_document),
            "latest_summary_path": str(latest_summary),
            "contract_no": result.summary["contract_no"],
        }

    def clearance_reset_template(self) -> dict[str, Any]:
        clearance = self.load_clearance_module()
        self.clearance_dir.mkdir(parents=True, exist_ok=True)
        output_path = clearance.create_workbook_template(self.clearance_input_path, None, None)
        return {"output_path": str(output_path)}

    def clearance_capture_session(self) -> dict[str, Any]:
        site = self.load_site_module()
        self.clearance_result_root.mkdir(parents=True, exist_ok=True)
        session = site.capture_session()
        return {
            "session_path": str(session.session_path),
            "storage_state_path": str(session.storage_state_path),
            "captured_at": session.captured_at,
            "user_label": site.extract_user_label(session.user_info) or "",
        }

    def clearance_check_session(self) -> dict[str, Any]:
        site = self.load_site_module()
        return site.validate_saved_session()

    def clearance_query_one(self, identifier: str, mode: str = "auto", iemark: str | None = None) -> dict[str, Any]:
        site = self.load_site_module()
        selection = site.query_site_identifier(identifier=identifier, mode=mode, iemark=iemark)
        row = selection.selected_row or {}
        result = {
            "identifier": selection.identifier,
            "mode": selection.mode,
            "iemark": selection.iemark or "",
            "released": selection.released,
            "status_text": selection.status_text,
            "result_count": selection.result_count,
            "entry_no": row.get("entryNo", ""),
            "bl_no": row.get("blNo", ""),
            "ctnr_no": row.get("ctnrNo", ""),
            "pcs": row.get("pkgNbr", ""),
            "gross_weight_kg": row.get("cargoGrossWeight", ""),
            "release_time": row.get("cusLetpasTime", ""),
        }
        latest_query_json = self.clearance_site_query_dir / "00-latest-single-query.json"
        latest_query_txt = self.clearance_site_query_dir / "00-latest-single-query.txt"
        latest_site_report_json = self.clearance_site_query_dir / "00-latest-site-report.json"
        latest_site_report_txt = self.clearance_site_query_dir / "00-latest-site-report.txt"
        latest_query_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        latest_query_txt.write_text(self.render_clearance_query_text(result), encoding="utf-8")
        self.copy_as_latest(latest_query_json, latest_site_report_json)
        self.copy_as_latest(latest_query_txt, latest_site_report_txt)
        result["latest_query_json"] = str(latest_query_json)
        result["latest_query_txt"] = str(latest_query_txt)
        result["latest_site_report_json"] = str(latest_site_report_json)
        result["latest_site_report_txt"] = str(latest_site_report_txt)
        return result

    def clearance_update_from_current(self) -> dict[str, Any]:
        return self.clearance_update_from_file(self.clearance_input_path)

    def clearance_update_from_file(self, source_path: Path | str) -> dict[str, Any]:
        site = self.load_site_module()
        source_path = Path(source_path).expanduser().resolve()
        if not source_path.exists():
            raise FileNotFoundError(f"Workbook not found: {source_path}")
        result = site.query_and_update_workbook(
            workbook_path=source_path,
            output_dir=self.clearance_updated_dir,
        )
        latest_workbook = self.clearance_updated_dir / "00-latest-clearance-workbook.xlsx"
        latest_clearance_report_json = self.clearance_root / "reports" / "00-latest-clearance-report.json"
        latest_clearance_report_txt = self.clearance_root / "reports" / "00-latest-clearance-report.txt"
        latest_site_report_json = self.clearance_site_query_dir / "00-latest-site-report.json"
        latest_site_report_txt = self.clearance_site_query_dir / "00-latest-site-report.txt"
        self.copy_as_latest(result.clearance_result.workbook_path, latest_workbook)
        self.copy_as_latest(result.clearance_result.report_json_path, latest_clearance_report_json)
        self.copy_as_latest(result.clearance_result.report_txt_path, latest_clearance_report_txt)
        self.copy_as_latest(result.site_report_json_path, latest_site_report_json)
        self.copy_as_latest(result.site_report_txt_path, latest_site_report_txt)
        return {
            "updated_workbook": str(result.clearance_result.workbook_path),
            "clearance_report_json": str(result.clearance_result.report_json_path),
            "clearance_report_txt": str(result.clearance_result.report_txt_path),
            "site_report_json": str(result.site_report_json_path),
            "site_report_txt": str(result.site_report_txt_path),
            "latest_updated_workbook": str(latest_workbook),
            "latest_clearance_report_json": str(latest_clearance_report_json),
            "latest_clearance_report_txt": str(latest_clearance_report_txt),
            "latest_site_report_json": str(latest_site_report_json),
            "latest_site_report_txt": str(latest_site_report_txt),
            "site_summary": result.site_summary,
        }

    def fetch_latest_release(self) -> ReleaseInfo:
        response = requests.get(APP_RELEASE_API, timeout=20, headers={"User-Agent": APP_NAME})
        if response.status_code == 404:
            raise FileNotFoundError("GitHub 上还没有可用的 release。")
        response.raise_for_status()
        payload = response.json()
        version = str(payload.get("tag_name", "")).strip()
        assets = payload.get("assets", []) or []
        package_asset = None
        for asset in assets:
            name = str(asset.get("name", ""))
            if name.startswith(RELEASE_PACKAGE_PREFIX) and name.endswith(".zip"):
                package_asset = asset
                break
        if package_asset is None:
            for asset in assets:
                name = str(asset.get("name", ""))
                if name.endswith(".zip"):
                    package_asset = asset
                    break
        if package_asset is None:
            raise RuntimeError("最新 release 里没有找到桌面版 zip 安装包。")
        return ReleaseInfo(
            version=version,
            html_url=payload.get("html_url") or APP_REPO_URL,
            published_at=payload.get("published_at", ""),
            asset_name=package_asset.get("name", ""),
            asset_download_url=package_asset.get("browser_download_url", ""),
            asset_size=int(package_asset.get("size") or 0),
            body=payload.get("body", "") or "",
        )

    def check_for_updates(self) -> dict[str, Any]:
        try:
            release = self.fetch_latest_release()
        except FileNotFoundError:
            return {
                "update_available": False,
                "message": "GitHub 上还没有可用的桌面版 release。",
                "current_version": APP_VERSION,
                "latest_version": "",
                "html_url": APP_REPO_URL,
                "install_supported": False,
            }
        latest_tuple = parse_version(release.version)
        current_tuple = parse_version(APP_VERSION)
        update_available = latest_tuple > current_tuple
        install_supported = portable_install_root() is not None and bool(release.asset_download_url)
        return {
            "update_available": update_available,
            "message": "发现新版本。" if update_available else "当前已经是最新版本。",
            "current_version": APP_VERSION,
            "latest_version": release.version,
            "html_url": release.html_url,
            "published_at": release.published_at,
            "release_name": release.asset_name,
            "install_supported": install_supported,
            "asset_download_url": release.asset_download_url,
        }

    def prepare_update_install(
        self,
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
    ) -> dict[str, Any]:
        install_root = portable_install_root()
        if install_root is None:
            raise RuntimeError("自动更新只支持 release 版桌面应用。源码模式请直接更新仓库或重新下载 release。")
        self._emit_update_progress(progress_callback, value=4, message="正在读取最新版本信息…")
        release = self.fetch_latest_release()
        if parse_version(release.version) <= parse_version(APP_VERSION):
            raise RuntimeError("当前已经是最新版本，不需要更新。")

        update_root = settings_root() / "updates" / release.version
        update_root.mkdir(parents=True, exist_ok=True)
        zip_path = update_root / release.asset_name

        self._download_file(
            release.asset_download_url,
            zip_path,
            expected_size=release.asset_size,
            progress_callback=progress_callback,
            progress_start=10,
            progress_end=78,
        )

        updater_root = update_root / "updater-app"
        updater_exe_path = self._prepare_updater_bundle(
            updater_root,
            progress_callback=progress_callback,
            progress_start=82,
            progress_end=94,
        )

        self._emit_update_progress(progress_callback, value=96, message="正在写入更新信息…")
        manifest_path = update_root / "update-manifest.json"
        manifest_payload = {
            "version": release.version,
            "current_pid": os.getpid(),
            "zip_path": str(zip_path),
            "install_root": str(install_root),
            "bundle_dir_name": Path(sys.executable).resolve().parent.name,
            "exe_name": Path(sys.executable).resolve().name,
            "top_level_files": UPDATE_TOP_LEVEL_FILES,
        }
        manifest_path.write_text(json.dumps(manifest_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self._emit_update_progress(progress_callback, value=100, message="更新包准备完成，即将切换到安装器…")

        return {
            "version": release.version,
            "zip_path": str(zip_path),
            "manifest_path": str(manifest_path),
            "updater_exe_path": str(updater_exe_path),
        }

    def launch_prepared_update(self, prepared_update: dict[str, Any]) -> None:
        updater_exe = Path(prepared_update["updater_exe_path"]).resolve()
        manifest_path = Path(prepared_update["manifest_path"]).resolve()
        if not updater_exe.exists():
            raise FileNotFoundError(f"Updater executable not found: {updater_exe}")
        if not manifest_path.exists():
            raise FileNotFoundError(f"Update manifest not found: {manifest_path}")

        subprocess.Popen(
            [str(updater_exe), "--apply-update-manifest", str(manifest_path)],
            cwd=str(updater_exe.parent),
        )

    def apply_prepared_update(
        self,
        manifest_path: Path | str,
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
    ) -> dict[str, Any]:
        manifest_path = Path(manifest_path).expanduser().resolve()
        payload = json.loads(manifest_path.read_text(encoding="utf-8-sig"))

        install_root = Path(payload["install_root"]).resolve()
        zip_path = Path(payload["zip_path"]).resolve()
        current_pid = int(payload["current_pid"])
        bundle_dir_name = str(payload.get("bundle_dir_name") or "AMS-Assistant-Desktop")
        exe_name = str(payload.get("exe_name") or "AMS-Assistant-Desktop.exe")
        top_level_files = list(payload.get("top_level_files") or UPDATE_TOP_LEVEL_FILES)

        update_root = manifest_path.parent
        extract_root = update_root / "extracted"
        app_root = install_root / "app"
        target_app_dir = app_root / bundle_dir_name
        stage_app_dir = app_root / f"{bundle_dir_name}.incoming"
        backup_app_dir = app_root / f"{bundle_dir_name}.backup-{self._timestamp()}"
        help_dir = install_root / "help"
        stage_help_dir = install_root / "help.incoming"
        user_data_dir = install_root / "user-data"

        self._wait_for_process_exit(current_pid, progress_callback=progress_callback, progress_start=0, progress_end=12)
        payload_root = self._extract_release_archive(
            zip_path,
            extract_root,
            progress_callback=progress_callback,
            progress_start=12,
            progress_end=40,
        )

        source_app_dir = payload_root / "app" / bundle_dir_name
        if not source_app_dir.exists():
            raise RuntimeError("更新包中缺少 app 目录，无法继续安装。")

        app_root.mkdir(parents=True, exist_ok=True)
        if stage_app_dir.exists():
            shutil.rmtree(stage_app_dir)
        if source_app_dir.drive.lower() == stage_app_dir.drive.lower():
            self._emit_update_progress(progress_callback, value=40, message="正在快速搬运新版本程序文件…")
            source_app_dir.rename(stage_app_dir)
            self._emit_update_progress(progress_callback, value=72, message="新版本程序文件已就位。")
        else:
            self._copy_tree_with_progress(
                source_app_dir,
                stage_app_dir,
                progress_callback=progress_callback,
                progress_start=40,
                progress_end=72,
                message="正在复制新版本程序文件…",
            )

        self._emit_update_progress(progress_callback, value=74, message="正在切换到新版本程序文件…")
        if target_app_dir.exists():
            if backup_app_dir.exists():
                shutil.rmtree(backup_app_dir)
            target_app_dir.rename(backup_app_dir)
        try:
            stage_app_dir.rename(target_app_dir)
        except Exception:
            if target_app_dir.exists():
                shutil.rmtree(target_app_dir)
            if backup_app_dir.exists() and not target_app_dir.exists():
                backup_app_dir.rename(target_app_dir)
            raise

        source_help_dir = payload_root / "help"
        if stage_help_dir.exists():
            shutil.rmtree(stage_help_dir)
        if source_help_dir.exists():
            if source_help_dir.drive.lower() == stage_help_dir.drive.lower():
                self._emit_update_progress(progress_callback, value=72, message="正在快速搬运帮助与说明文件…")
                source_help_dir.rename(stage_help_dir)
                self._emit_update_progress(progress_callback, value=84, message="帮助与说明文件已就位。")
            else:
                self._copy_tree_with_progress(
                    source_help_dir,
                    stage_help_dir,
                    progress_callback=progress_callback,
                    progress_start=72,
                    progress_end=84,
                    message="正在更新帮助与说明文件…",
                )
            help_backup_dir = install_root / f"help.backup-{self._timestamp()}"
            if help_backup_dir.exists():
                shutil.rmtree(help_backup_dir)
            try:
                if help_dir.exists():
                    help_dir.rename(help_backup_dir)
                stage_help_dir.rename(help_dir)
                if help_backup_dir.exists():
                    shutil.rmtree(help_backup_dir)
            except Exception:
                if help_dir.exists():
                    shutil.rmtree(help_dir)
                if help_backup_dir.exists() and not help_dir.exists():
                    help_backup_dir.rename(help_dir)
                raise

        self._copy_top_level_files(
            payload_root,
            install_root,
            top_level_files,
            progress_callback=progress_callback,
            progress_start=84,
            progress_end=92,
        )

        payload_user_data_dir = payload_root / "user-data"
        if not user_data_dir.exists() and payload_user_data_dir.exists():
            self._copy_tree_with_progress(
                payload_user_data_dir,
                user_data_dir,
                progress_callback=progress_callback,
                progress_start=92,
                progress_end=95,
                message="正在准备用户数据目录…",
            )

        self._emit_update_progress(progress_callback, value=96, message="正在清理安装缓存…")
        if extract_root.exists():
            shutil.rmtree(extract_root)

        self._emit_update_progress(progress_callback, value=98, message="正在启动新版本…")
        restarted = self._restart_installed_app(
            install_root=install_root,
            bundle_dir_name=bundle_dir_name,
            exe_name=exe_name,
        )
        if backup_app_dir.exists():
            shutil.rmtree(backup_app_dir)
        self._emit_update_progress(progress_callback, value=100, message="更新完成，已尝试启动新版本。")
        return {
            "version": payload.get("version", ""),
            "install_root": str(install_root),
            "restarted_executable": str(restarted),
        }

    def _download_file(
        self,
        url: str,
        destination: Path,
        expected_size: int = 0,
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
        progress_start: float = 0,
        progress_end: float = 100,
    ) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        if destination.exists() and expected_size and destination.stat().st_size == expected_size:
            self._emit_update_progress(
                progress_callback,
                value=progress_end,
                message="已复用已下载的安装包。",
                detail=self._format_size(expected_size),
            )
            return

        partial_path = destination.with_suffix(destination.suffix + ".part")
        if partial_path.exists():
            partial_path.unlink()

        with requests.get(url, timeout=60, stream=True, headers={"User-Agent": APP_NAME}) as response:
            response.raise_for_status()
            total_bytes = int(response.headers.get("content-length") or 0) or expected_size
            downloaded = 0
            with partial_path.open("wb") as fh:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if not chunk:
                        continue
                    fh.write(chunk)
                    downloaded += len(chunk)
                    if total_bytes > 0:
                        ratio = min(downloaded / total_bytes, 1.0)
                        value = progress_start + ratio * (progress_end - progress_start)
                        detail = f"{self._format_size(downloaded)} / {self._format_size(total_bytes)}"
                    else:
                        value = progress_start
                        detail = f"已下载 {self._format_size(downloaded)}"
                    self._emit_update_progress(
                        progress_callback,
                        value=value,
                        message="正在下载更新包…",
                        detail=detail,
                    )
        partial_path.replace(destination)

    def _prepare_updater_bundle(
        self,
        destination_root: Path,
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
        progress_start: float = 82,
        progress_end: float = 94,
    ) -> Path:
        current_bundle_dir = Path(sys.executable).resolve().parent
        helper_source = current_bundle_dir / UPDATER_EXE_NAME
        if destination_root.exists():
            shutil.rmtree(destination_root)
        destination_root.mkdir(parents=True, exist_ok=True)

        if helper_source.exists():
            helper_target = destination_root / UPDATER_EXE_NAME
            self._emit_update_progress(progress_callback, value=progress_start, message="正在准备独立更新器…", detail=UPDATER_EXE_NAME)
            shutil.copy2(helper_source, helper_target)
            self._emit_update_progress(progress_callback, value=progress_end, message="独立更新器已准备完成。", detail=str(helper_target.name))
            return helper_target

        destination_bundle_dir = destination_root / current_bundle_dir.name
        self._copy_tree_with_progress(
            current_bundle_dir,
            destination_bundle_dir,
            progress_callback=progress_callback,
            progress_start=progress_start,
            progress_end=progress_end,
            message="正在准备独立更新器…",
        )
        return destination_bundle_dir / Path(sys.executable).resolve().name

    def _wait_for_process_exit(
        self,
        current_pid: int,
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
        progress_start: float = 0,
        progress_end: float = 100,
        timeout_seconds: int = 120,
    ) -> None:
        for tick in range(timeout_seconds):
            if not self._is_pid_running(current_pid):
                self._emit_update_progress(progress_callback, value=progress_end, message="当前版本已退出，准备安装…")
                return
            ratio = (tick + 1) / max(timeout_seconds, 1)
            value = progress_start + ratio * (progress_end - progress_start)
            self._emit_update_progress(
                progress_callback,
                value=value,
                message="正在等待当前版本安全退出…",
                detail=f"最多再等待 {timeout_seconds - tick - 1} 秒",
            )
            time.sleep(1)
        raise TimeoutError("等待当前版本退出超时，请先关闭仍在运行的 AMS Assistant 后重试。")

    def _is_pid_running(self, pid: int) -> bool:
        if pid <= 0:
            return False
        if sys.platform.startswith("win"):
            result = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
            )
            output = (result.stdout or "") + (result.stderr or "")
            return str(pid) in output and "No tasks are running" not in output
        try:
            os.kill(pid, 0)
        except OSError:
            return False
        return True

    def _extract_release_archive(
        self,
        zip_path: Path,
        destination_root: Path,
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
        progress_start: float = 0,
        progress_end: float = 100,
    ) -> Path:
        if destination_root.exists():
            shutil.rmtree(destination_root)
        destination_root.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(zip_path, "r") as archive:
            members = archive.infolist()
            total_members = max(len(members), 1)
            for index, member in enumerate(members, start=1):
                archive.extract(member, destination_root)
                value = progress_start + (index / total_members) * (progress_end - progress_start)
                self._emit_update_progress(
                    progress_callback,
                    value=value,
                    message="正在解压更新包…",
                    detail=member.filename,
                )

        payload_root = destination_root
        if not (payload_root / "app").exists():
            subdirs = [item for item in destination_root.iterdir() if item.is_dir()]
            if len(subdirs) == 1 and (subdirs[0] / "app").exists():
                payload_root = subdirs[0]
            else:
                raise RuntimeError("无法识别 release 压缩包结构。")
        return payload_root

    def _copy_tree_with_progress(
        self,
        source_dir: Path,
        destination_dir: Path,
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
        progress_start: float = 0,
        progress_end: float = 100,
        message: str = "正在复制文件…",
    ) -> None:
        files = [path for path in source_dir.rglob("*") if path.is_file()]
        destination_dir.mkdir(parents=True, exist_ok=True)
        if not files:
            self._emit_update_progress(progress_callback, value=progress_end, message=message)
            return

        total_files = len(files)
        for index, source_path in enumerate(files, start=1):
            relative_path = source_path.relative_to(source_dir)
            target_path = destination_dir / relative_path
            target_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, target_path)
            value = progress_start + (index / total_files) * (progress_end - progress_start)
            self._emit_update_progress(
                progress_callback,
                value=value,
                message=message,
                detail=str(relative_path),
            )

    def _copy_top_level_files(
        self,
        source_root: Path,
        destination_root: Path,
        filenames: list[str],
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
        progress_start: float = 0,
        progress_end: float = 100,
    ) -> None:
        if not filenames:
            return
        total = len(filenames)
        for index, filename in enumerate(filenames, start=1):
            source_path = source_root / filename
            if source_path.exists():
                shutil.copy2(source_path, destination_root / filename)
            value = progress_start + (index / total) * (progress_end - progress_start)
            self._emit_update_progress(
                progress_callback,
                value=value,
                message="正在更新启动器与说明文件…",
                detail=filename,
            )

    def _restart_installed_app(self, install_root: Path, bundle_dir_name: str, exe_name: str) -> Path:
        executable_path = install_root / "app" / bundle_dir_name / exe_name
        if not executable_path.exists():
            raise FileNotFoundError(f"Updated executable not found: {executable_path}")

        user_data_root = install_root / "user-data"
        settings_dir = user_data_root / "settings"
        workspace_dir = user_data_root / "workspace"
        settings_dir.mkdir(parents=True, exist_ok=True)
        workspace_dir.mkdir(parents=True, exist_ok=True)

        env = os.environ.copy()
        env[ENV_SETTINGS_DIR] = str(settings_dir)
        env[ENV_DEFAULT_WORKSPACE] = str(workspace_dir)

        subprocess.Popen(
            [str(executable_path)],
            cwd=str(executable_path.parent),
            env=env,
        )
        return executable_path

    def _emit_update_progress(
        self,
        progress_callback: Callable[[dict[str, Any]], None] | None,
        *,
        value: float,
        message: str,
        detail: str = "",
    ) -> None:
        if progress_callback is None:
            return
        progress_callback(
            {
                "value": max(0.0, min(100.0, float(value))),
                "message": message,
                "detail": detail,
            }
        )

    def _format_size(self, size_bytes: int) -> str:
        units = ["B", "KB", "MB", "GB"]
        size = float(max(size_bytes, 0))
        for unit in units:
            if size < 1024 or unit == units[-1]:
                if unit == "B":
                    return f"{int(size)} {unit}"
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{int(size_bytes)} B"

    def render_clearance_query_text(self, result: dict[str, Any]) -> str:
        lines = [
            "通关查询结果",
            f"identifier: {result.get('identifier', '')}",
            f"mode: {result.get('mode', '')}",
            f"iemark: {result.get('iemark', '') or 'auto'}",
            f"released: {result.get('released', False)}",
            f"status_text: {result.get('status_text', '') or 'N/A'}",
            f"entry_no: {result.get('entry_no', '') or 'N/A'}",
            f"bl_no: {result.get('bl_no', '') or 'N/A'}",
            f"ctnr_no: {result.get('ctnr_no', '') or 'N/A'}",
            f"pcs: {result.get('pcs', '') or 'N/A'}",
            f"gross_weight_kg: {result.get('gross_weight_kg', '') or 'N/A'}",
            f"release_time: {result.get('release_time', '') or 'N/A'}",
        ]
        return "\n".join(lines) + "\n"

    def copy_as_latest(self, source_path: Path | str, target_path: Path | str) -> None:
        source = Path(source_path).resolve()
        target = Path(target_path).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)

    def _timestamp(self) -> str:
        from datetime import datetime

        return datetime.now().strftime("%Y%m%d-%H%M%S")

    def _migrate_legacy_dir(self, legacy_name: str, new_path: Path) -> None:
        legacy_path = self.workspace_root / legacy_name
        if legacy_path.exists() and not new_path.exists():
            shutil.move(str(legacy_path), str(new_path))

    def _write_workspace_notes(self) -> None:
        workspace_note = self.workspace_root / "00-工作区说明.html"
        if not workspace_note.exists():
            workspace_note.write_text(
                """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>AMS 工作区说明</title></head>
<body style="font-family:'Microsoft YaHei UI',sans-serif;padding:24px;line-height:1.8;">
<h1>AMS 工作区说明</h1>
<p>这里是 AMS Assistant 给普通用户准备的固定工作区。</p>
<ul>
<li><strong>合同生成</strong>：合同输入 Excel、合同结果、摘要 JSON、备份文件。</li>
<li><strong>通关查询</strong>：通关输入 Excel、登录态、查询报告、更新后的工作簿。</li>
<li><strong>船期与港区矩阵</strong>：后续功能预留入口。</li>
<li><strong>表格自动同步</strong>：同步功能的示例文件和运行日志。</li>
</ul>
<p>建议优先在 AMS Assistant Desktop 内操作，而不是手工改动这些目录结构。</p>
</body></html>
""",
                encoding="utf-8",
            )

        lineup_note = self.lineup_dir / "README.txt"
        if not lineup_note.exists():
            lineup_note.write_text(
                "这里是“船期与港区矩阵”功能的预留目录。\n后续会在这里接入船期表查询、港区矩阵生成和报告输出。\n",
                encoding="utf-8",
            )
