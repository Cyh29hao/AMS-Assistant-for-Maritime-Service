from __future__ import annotations

import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

import requests

from desktop_app.excel_sync_engine import (
    DEFAULT_SETTINGS as DEFAULT_SYNC_SETTINGS,
    ExcelSyncPaths,
    SyncService,
    create_demo_workbooks,
)


APP_NAME = "AMS Assistant"
APP_VERSION = "0.3.1"
APP_REPO_URL = "https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service"
APP_RELEASE_API = "https://api.github.com/repos/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases/latest"
RELEASE_PACKAGE_PREFIX = "AMS-Assistant-Desktop-v"

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

    def prepare_update_install(self) -> dict[str, Any]:
        install_root = portable_install_root()
        if install_root is None:
            raise RuntimeError("自动更新只支持 release 版桌面应用。源码模式请直接更新仓库或重新下载 release。")
        release = self.fetch_latest_release()
        if parse_version(release.version) <= parse_version(APP_VERSION):
            raise RuntimeError("当前已经是最新版本，不需要更新。")

        update_root = settings_root() / "updates" / release.version
        update_root.mkdir(parents=True, exist_ok=True)
        zip_path = update_root / release.asset_name
        if not zip_path.exists():
            self._download_file(release.asset_download_url, zip_path)

        script_path = update_root / "install-update.ps1"
        launcher_path = update_root / "install-update.cmd"
        script_path.write_text(
            self._build_update_script(
                current_pid=os.getpid(),
                zip_path=zip_path,
                install_root=install_root,
            ),
            encoding="utf-8-sig",
        )
        launcher_path.write_text(
            (
                "@echo off\r\n"
                "setlocal\r\n"
                f"powershell -ExecutionPolicy Bypass -File \"{script_path}\"\r\n"
                "exit /b %errorlevel%\r\n"
            ),
            encoding="utf-8",
        )
        return {
            "version": release.version,
            "zip_path": str(zip_path),
            "script_path": str(script_path),
            "launcher_path": str(launcher_path),
        }

    def _download_file(self, url: str, destination: Path) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        with requests.get(url, timeout=60, stream=True, headers={"User-Agent": APP_NAME}) as response:
            response.raise_for_status()
            with destination.open("wb") as fh:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        fh.write(chunk)

    def _build_update_script(self, current_pid: int, zip_path: Path, install_root: Path) -> str:
        top_level_files = [
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
        files_literal = "@(" + ",".join(f"'{name}'" for name in top_level_files) + ")"
        return f"""$ErrorActionPreference = 'Stop'
$currentPid = {current_pid}
$zipPath = '{zip_path}'
$installRoot = '{install_root}'
$extractRoot = Join-Path (Split-Path -Parent $zipPath) 'extracted'
$backupApp = Join-Path $installRoot ('app-backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
$targetApp = Join-Path $installRoot 'app'
$targetHelp = Join-Path $installRoot 'help'
$userDataDst = Join-Path $installRoot 'user-data'

try {{
  Wait-Process -Id $currentPid -Timeout 120 -ErrorAction SilentlyContinue
}} catch {{
}}

if (Test-Path $extractRoot) {{
  Remove-Item -LiteralPath $extractRoot -Recurse -Force
}}

Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force
$payloadRoot = $extractRoot
if (!(Test-Path (Join-Path $payloadRoot 'app'))) {{
  $dirs = Get-ChildItem -LiteralPath $extractRoot -Directory
  if ($dirs.Count -eq 1 -and (Test-Path (Join-Path $dirs[0].FullName 'app'))) {{
    $payloadRoot = $dirs[0].FullName
  }} else {{
    throw '无法识别 release 压缩包结构。'
  }}
}}

$sourceApp = Join-Path $payloadRoot 'app'
if (!(Test-Path $sourceApp)) {{
  throw '压缩包中缺少 app 目录。'
}}

if (Test-Path $targetApp) {{
  Move-Item -LiteralPath $targetApp -Destination $backupApp
}}
Copy-Item -LiteralPath $sourceApp -Destination $targetApp -Recurse -Force

$sourceHelp = Join-Path $payloadRoot 'help'
if (Test-Path $targetHelp) {{
  Remove-Item -LiteralPath $targetHelp -Recurse -Force
}}
if (Test-Path $sourceHelp) {{
  Copy-Item -LiteralPath $sourceHelp -Destination $targetHelp -Recurse -Force
}}

foreach ($name in {files_literal}) {{
  $src = Join-Path $payloadRoot $name
  if (Test-Path $src) {{
    Copy-Item -LiteralPath $src -Destination (Join-Path $installRoot $name) -Force
  }}
}}

$userDataSrc = Join-Path $payloadRoot 'user-data'
if (!(Test-Path $userDataDst) -and (Test-Path $userDataSrc)) {{
  Copy-Item -LiteralPath $userDataSrc -Destination $userDataDst -Recurse -Force
}}

if (Test-Path $backupApp) {{
  Remove-Item -LiteralPath $backupApp -Recurse -Force
}}

Start-Process -FilePath (Join-Path $installRoot '1-启动AMS桌面应用.bat')
"""

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
