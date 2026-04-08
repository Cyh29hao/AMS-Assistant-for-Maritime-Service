from __future__ import annotations

import importlib.util
import json
import os
import shutil
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

import requests


APP_NAME = "AMS Assistant"
APP_VERSION = "0.1.0"
APP_REPO_URL = "https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service"
APP_RELEASE_API = "https://api.github.com/repos/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases/latest"
ENV_SETTINGS_DIR = "AMS_ASSISTANT_SETTINGS_DIR"
ENV_DEFAULT_WORKSPACE = "AMS_ASSISTANT_DEFAULT_WORKSPACE"
ENV_REQ2_BROWSER = "AMS_REQ2_BROWSER"


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
    raise RuntimeError("Only Windows is supported in this desktop build right now.")


@dataclass
class AppConfig:
    workspace_root: str
    theme_name: str = "flatly"
    auto_open_results: bool = True
    check_updates_on_launch: bool = False
    req1_input_filename: str = "contract-input.xlsx"
    req2_input_filename: str = "req2-input.xlsx"
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
    def req1_dir(self) -> Path:
        return self.workspace_root / "req1-系统出合同"

    @property
    def req1_input_path(self) -> Path:
        return self.req1_dir / self.config.req1_input_filename

    @property
    def req1_result_dir(self) -> Path:
        return self.req1_dir / "result"

    @property
    def req1_summary_dir(self) -> Path:
        return self.req1_result_dir / "summary"

    @property
    def req1_backup_dir(self) -> Path:
        return self.req1_dir / "backup"

    @property
    def req2_dir(self) -> Path:
        return self.workspace_root / "req2-自动查通关"

    @property
    def req2_input_path(self) -> Path:
        return self.req2_dir / self.config.req2_input_filename

    @property
    def req2_result_root(self) -> Path:
        return self.req2_dir / "result"

    @property
    def req2_clearance_root(self) -> Path:
        return self.req2_result_root / "clearance"

    @property
    def req2_updated_dir(self) -> Path:
        return self.req2_clearance_root / "updated"

    @property
    def req2_site_session_dir(self) -> Path:
        return self.req2_clearance_root / "site_session"

    @property
    def req2_site_checks_dir(self) -> Path:
        return self.req2_clearance_root / "site_checks"

    @property
    def req2_site_query_dir(self) -> Path:
        return self.req2_clearance_root / "site_query_results"

    @property
    def req3_dir(self) -> Path:
        return self.workspace_root / "req3-船期表"

    def set_runtime_env(self) -> None:
        os.environ["AMS_DATA_ROOT"] = str(self.req2_result_root)
        os.environ[ENV_REQ2_BROWSER] = self.config.req2_browser_preference

    def load_contract_module(self):
        return load_source_module("contract_workflow_runtime", scripts_dir() / "contract_workflow.py")

    def load_clearance_module(self):
        return load_source_module("clearance_workflow", scripts_dir() / "clearance_workflow.py")

    def load_site_module(self):
        self.set_runtime_env()
        if "clearance_workflow" in sys.modules:
            del sys.modules["clearance_workflow"]
        return load_source_module("clearance_site_workflow_runtime", scripts_dir() / "clearance_site_workflow.py")

    def ensure_workspace(self) -> dict[str, Any]:
        self.workspace_root.mkdir(parents=True, exist_ok=True)
        self.req1_dir.mkdir(parents=True, exist_ok=True)
        self.req1_result_dir.mkdir(parents=True, exist_ok=True)
        self.req1_summary_dir.mkdir(parents=True, exist_ok=True)
        self.req1_backup_dir.mkdir(parents=True, exist_ok=True)

        self.req2_dir.mkdir(parents=True, exist_ok=True)
        self.req2_result_root.mkdir(parents=True, exist_ok=True)
        self.req2_clearance_root.mkdir(parents=True, exist_ok=True)
        self.req2_updated_dir.mkdir(parents=True, exist_ok=True)
        self.req2_site_session_dir.mkdir(parents=True, exist_ok=True)
        self.req2_site_checks_dir.mkdir(parents=True, exist_ok=True)
        self.req2_site_query_dir.mkdir(parents=True, exist_ok=True)

        self.req3_dir.mkdir(parents=True, exist_ok=True)
        workspace_note = self.workspace_root / "00-工作区说明.html"
        if not workspace_note.exists():
            workspace_note.write_text(
                """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>AMS 工作区说明</title></head>
<body style="font-family:'Microsoft YaHei UI',sans-serif;padding:24px;line-height:1.8;">
<h1>AMS 工作区说明</h1>
<p>这里是 AMS Assistant 为普通用户准备的工作区。</p>
<ul>
<li><strong>req1-系统出合同</strong>：合同输入 Excel、合同结果、摘要 JSON、备份文件</li>
<li><strong>req2-自动查通关</strong>：通关输入 Excel、登录态、查询报告、更新后工作簿</li>
<li><strong>req3-船期表</strong>：后续预留入口</li>
</ul>
<p>建议优先回到 AMS Assistant Desktop 里操作，而不是直接手工改这些文件夹结构。</p>
</body></html>
""",
                encoding="utf-8",
            )
        req3_note = self.req3_dir / "README.txt"
        if not req3_note.exists():
            req3_note.write_text(
                "req3 入口已预留。\n后续会在这里接入船期表查询与生成。\n",
                encoding="utf-8",
            )

        if not self.req1_input_path.exists():
            self.req1_reset_template(backup_existing=False)
        if not self.req2_input_path.exists():
            self.req2_reset_template()

        return {
            "workspace_root": str(self.workspace_root),
            "req1_input_path": str(self.req1_input_path),
            "req2_input_path": str(self.req2_input_path),
        }

    def req1_reset_template(self, backup_existing: bool = True) -> dict[str, Any]:
        contract = self.load_contract_module()
        self.req1_dir.mkdir(parents=True, exist_ok=True)
        self.req1_backup_dir.mkdir(parents=True, exist_ok=True)
        if backup_existing and self.req1_input_path.exists():
            timestamp = self._timestamp()
            backup_path = self.req1_backup_dir / f"contract-input-{timestamp}.xlsx"
            shutil.copy2(self.req1_input_path, backup_path)
        output_path = contract.create_workbook_template(self.req1_input_path)
        return {"output_path": str(output_path)}

    def req1_generate_from_current(self) -> dict[str, Any]:
        return self.req1_generate_from_file(self.req1_input_path)

    def req1_generate_from_file(self, source_path: Path | str) -> dict[str, Any]:
        contract = self.load_contract_module()
        source_path = Path(source_path).expanduser().resolve()
        if not source_path.exists():
            raise FileNotFoundError(f"Excel file not found: {source_path}")
        request = contract.load_workbook_request(source_path)
        normalized = contract.validate_request(request, contract.load_registry())
        result = contract.render_contract(
            normalized,
            self.req1_result_dir,
            self.req1_summary_dir,
        )
        latest_document = self.req1_result_dir / "00-最新合同.docx"
        latest_summary = self.req1_summary_dir / "00-最新摘要.json"
        self.copy_as_latest(result.document_path, latest_document)
        self.copy_as_latest(result.summary_path, latest_summary)
        return {
            "document_path": str(result.document_path),
            "summary_path": str(result.summary_path),
            "latest_document_path": str(latest_document),
            "latest_summary_path": str(latest_summary),
            "contract_no": result.summary["contract_no"],
        }

    def req2_reset_template(self) -> dict[str, Any]:
        clearance = self.load_clearance_module()
        self.req2_dir.mkdir(parents=True, exist_ok=True)
        output_path = clearance.create_workbook_template(self.req2_input_path, None, None)
        return {"output_path": str(output_path)}

    def req2_capture_session(self) -> dict[str, Any]:
        site = self.load_site_module()
        self.req2_result_root.mkdir(parents=True, exist_ok=True)
        session = site.capture_session()
        return {
            "session_path": str(session.session_path),
            "storage_state_path": str(session.storage_state_path),
            "captured_at": session.captured_at,
            "user_label": site.extract_user_label(session.user_info) or "",
        }

    def req2_check_session(self) -> dict[str, Any]:
        site = self.load_site_module()
        report = site.validate_saved_session()
        return report

    def req2_query_one(self, identifier: str, mode: str = "auto", iemark: str | None = None) -> dict[str, Any]:
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
        latest_query_json = self.req2_site_query_dir / "00-最新单票查询.json"
        latest_query_txt = self.req2_site_query_dir / "00-最新单票查询.txt"
        latest_site_report_json = self.req2_site_query_dir / "00-最新网站报告.json"
        latest_site_report_txt = self.req2_site_query_dir / "00-最新网站报告.txt"
        latest_query_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        latest_query_txt.write_text(self.render_req2_query_one_text(result), encoding="utf-8")
        self.copy_as_latest(latest_query_json, latest_site_report_json)
        self.copy_as_latest(latest_query_txt, latest_site_report_txt)
        result["latest_query_json"] = str(latest_query_json)
        result["latest_query_txt"] = str(latest_query_txt)
        result["latest_site_report_json"] = str(latest_site_report_json)
        result["latest_site_report_txt"] = str(latest_site_report_txt)
        return result

    def req2_update_from_current(self) -> dict[str, Any]:
        return self.req2_update_from_file(self.req2_input_path)

    def req2_update_from_file(self, source_path: Path | str) -> dict[str, Any]:
        site = self.load_site_module()
        source_path = Path(source_path).expanduser().resolve()
        if not source_path.exists():
            raise FileNotFoundError(f"Req2 workbook not found: {source_path}")
        result = site.query_and_update_workbook(
            workbook_path=source_path,
            output_dir=self.req2_updated_dir,
        )
        latest_workbook = self.req2_updated_dir / "00-最新更新工作簿.xlsx"
        latest_clearance_report_json = self.req2_clearance_root / "reports" / "00-最新清关报告.json"
        latest_clearance_report_txt = self.req2_clearance_root / "reports" / "00-最新清关报告.txt"
        latest_site_report_json = self.req2_site_query_dir / "00-最新网站报告.json"
        latest_site_report_txt = self.req2_site_query_dir / "00-最新网站报告.txt"
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

    def check_for_updates(self) -> dict[str, Any]:
        response = requests.get(APP_RELEASE_API, timeout=20, headers={"User-Agent": APP_NAME})
        if response.status_code == 404:
            return {
                "update_available": False,
                "message": "GitHub 上还没有可用的 release。",
                "current_version": APP_VERSION,
                "latest_version": "",
                "html_url": APP_REPO_URL,
            }
        response.raise_for_status()
        payload = response.json()
        latest_version = str(payload.get("tag_name", "")).strip()
        html_url = payload.get("html_url") or APP_REPO_URL
        update_available = bool(latest_version) and latest_version != APP_VERSION
        return {
            "update_available": update_available,
            "message": "发现新版本。" if update_available else "当前已经是最新版本，或线上版本号尚未更新。",
            "current_version": APP_VERSION,
            "latest_version": latest_version,
            "html_url": html_url,
            "release_name": payload.get("name", ""),
            "published_at": payload.get("published_at", ""),
        }

    def _timestamp(self) -> str:
        from datetime import datetime

        return datetime.now().strftime("%Y%m%d-%H%M%S")

    def copy_as_latest(self, source_path: Path | str, target_path: Path | str) -> None:
        source = Path(source_path).resolve()
        target = Path(target_path).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)

    def render_req2_query_one_text(self, result: dict[str, Any]) -> str:
        lines = [
            "req2 单票查询结果",
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
