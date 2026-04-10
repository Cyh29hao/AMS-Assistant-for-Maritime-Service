#!/usr/bin/env python3
"""Req2 site session capture and real-site query workflow."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
from openpyxl import load_workbook
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

import clearance_workflow as cw


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
DATA_ROOT = Path(os.environ.get("AMS_DATA_ROOT", str(SKILL_ROOT / "output"))).resolve()
OUTPUT_ROOT = DATA_ROOT / "clearance"
SITE_SESSION_DIR = OUTPUT_ROOT / "site_session"
SITE_QUERY_DIR = OUTPUT_ROOT / "site_query_results"
SITE_CHECK_DIR = OUTPUT_ROOT / "site_checks"
CONFIG_PATH = Path(
    os.environ.get("AMS_CLEARANCE_SITE_CONFIG_PATH", str(SCRIPT_DIR / "clearance_site_config.json"))
).resolve()
ENV_REQ2_BROWSER = "AMS_REQ2_BROWSER"


class SiteWorkflowError(Exception):
    """Base error for the req2 site workflow."""


class SessionMissingError(SiteWorkflowError):
    """Raised when a saved site session is missing or invalid."""


@dataclass
class SavedSession:
    token: str
    refresh_token: str
    local_storage: dict[str, str]
    user_info: dict[str, Any] | None
    captured_at: str
    session_path: Path
    storage_state_path: Path


@dataclass
class QuerySelection:
    identifier: str
    mode: str
    iemark: str | None
    released: bool
    status_text: str
    selected_row: dict[str, Any] | None
    result_count: int
    attempts: list[dict[str, Any]]


@dataclass
class WorkbookSiteResult:
    clearance_result: cw.ClearanceResult
    site_report_json_path: Path
    site_report_txt_path: Path
    site_summary: dict[str, Any]


def now_iso() -> str:
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_config() -> dict[str, Any]:
    return read_json(CONFIG_PATH)


def session_json_path() -> Path:
    return SITE_SESSION_DIR / "req2_site_session.json"


def storage_state_path() -> Path:
    return SITE_SESSION_DIR / "req2_site_storage_state.json"


def screenshot_path() -> Path:
    return SITE_SESSION_DIR / "req2_site_login_ok.png"


def check_report_json_path() -> Path:
    return SITE_CHECK_DIR / "req2_site_check.json"


def check_report_txt_path() -> Path:
    return SITE_CHECK_DIR / "req2_site_check.txt"


def compact(value: Any) -> str:
    return cw.compact_spaces(value)


def detect_identifier_mode(identifier: str) -> str:
    text = compact(identifier)
    if re.fullmatch(r"[A-Za-z]{4}\d{7}", text):
        return "ctnrNo"
    if re.fullmatch(r"\d{18}", text):
        return "entryNo"
    return "blNo"


def is_released_status(status_text: Any) -> bool:
    text = compact(status_text).lower()
    if not text:
        return False
    keywords = ("放行", "已放行", "released", "release", "letpas", "通关", "放货")
    return any(keyword in text for keyword in keywords)


def parse_json_maybe(text: str) -> dict[str, Any] | None:
    text = text.strip()
    if not text:
        return None
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        return None
    return payload if isinstance(payload, dict) else None


def extract_local_storage(page) -> dict[str, str]:
    return page.evaluate(
        """() => {
            const out = {};
            for (let i = 0; i < window.localStorage.length; i += 1) {
                const key = window.localStorage.key(i);
                out[key] = window.localStorage.getItem(key);
            }
            return out;
        }"""
    )


def extract_user_label(user_info: dict[str, Any] | None) -> str:
    if not user_info:
        return ""
    attrs = user_info.get("attrs") if isinstance(user_info.get("attrs"), dict) else {}
    for key in ("UserName", "UserloginId", "Userid", "CompanyName", "companyName"):
        value = attrs.get(key) or user_info.get(key)
        if compact(value):
            return compact(value)
    return ""


def build_headers(config: dict[str, Any], token: str, referer: str | None = None) -> dict[str, str]:
    return {
        "User-Agent": config["user_agent"],
        "Origin": config["base_origin"],
        "Referer": referer or config["default_referer"],
        "language": "languageChina",
        "Gw-Authorization-Rt": token,
    }


def auth_candidates(session: SavedSession) -> list[tuple[str, str]]:
    candidates: list[tuple[str, str]] = []
    token = compact(session.token)
    refresh_token = compact(session.refresh_token)
    pairs = [
        (token, refresh_token or token),
        (token, token),
        (refresh_token, refresh_token),
        (refresh_token, token or refresh_token),
    ]
    seen: set[tuple[str, str]] = set()
    for header_token, body_refresh_token in pairs:
        if not header_token or not body_refresh_token:
            continue
        pair = (header_token, body_refresh_token)
        if pair in seen:
            continue
        seen.add(pair)
        candidates.append(pair)
    return candidates


def save_session_payload(
    *,
    config: dict[str, Any],
    local_storage: dict[str, str],
    user_info: dict[str, Any] | None,
    captured_at: str,
) -> SavedSession:
    SITE_SESSION_DIR.mkdir(parents=True, exist_ok=True)
    session_path = session_json_path()
    storage_path = storage_state_path()
    token = compact(local_storage.get("TOKEN", ""))
    refresh_token = compact(local_storage.get("refresh_token", ""))
    payload = {
        "site_name": config["site_name"],
        "base_origin": config["base_origin"],
        "captured_at": captured_at,
        "token": token,
        "refresh_token": refresh_token,
        "user_label": extract_user_label(user_info),
        "user_info": user_info,
        "local_storage": local_storage,
        "storage_state_path": str(storage_path),
    }
    write_json(session_path, payload)
    return SavedSession(
        token=token,
        refresh_token=refresh_token,
        local_storage=local_storage,
        user_info=user_info,
        captured_at=captured_at,
        session_path=session_path,
        storage_state_path=storage_path,
    )


def load_saved_session() -> SavedSession:
    path = session_json_path()
    if not path.exists():
        raise SessionMissingError(
            "No saved req2 site session was found. Run capture-session first."
        )
    payload = read_json(path)
    return SavedSession(
        token=compact(payload.get("token", "")),
        refresh_token=compact(payload.get("refresh_token", "")),
        local_storage=payload.get("local_storage", {}) if isinstance(payload.get("local_storage"), dict) else {},
        user_info=payload.get("user_info") if isinstance(payload.get("user_info"), dict) else None,
        captured_at=compact(payload.get("captured_at", "")),
        session_path=path,
        storage_state_path=Path(payload.get("storage_state_path", storage_state_path())),
    )


def browser_candidates() -> list[str]:
    preferred = compact(os.environ.get(ENV_REQ2_BROWSER, "")).lower()
    if preferred in {"msedge", "chrome", "playwright"}:
        return [preferred]
    return ["msedge", "chrome", "playwright"]


def browser_label(candidate: str) -> str:
    mapping = {
        "msedge": "Microsoft Edge",
        "chrome": "Google Chrome",
        "playwright": "Playwright Chromium",
    }
    return mapping.get(candidate, candidate)


def launch_login_browser(playwright, headless: bool = False):
    errors: list[str] = []
    for candidate in browser_candidates():
        try:
            if candidate == "playwright":
                browser = playwright.chromium.launch(headless=headless)
            else:
                browser = playwright.chromium.launch(channel=candidate, headless=headless)
            return browser, browser_label(candidate)
        except PlaywrightError as exc:
            errors.append(f"{browser_label(candidate)}: {compact(exc)}")
    raise SiteWorkflowError(
        "无法启动 req2 登录浏览器。请确认这台电脑上至少装有 Edge 或 Chrome，"
        "或者当前环境已经具备 Playwright Chromium。\n"
        + "\n".join(errors)
    )


def capture_session(timeout_minutes: int | None = None) -> SavedSession:
    config = load_config()
    timeout_seconds = max(60, int((timeout_minutes or config["login_wait_minutes"]) * 60))
    SITE_SESSION_DIR.mkdir(parents=True, exist_ok=True)
    captured_at = now_iso()

    with sync_playwright() as playwright:
        browser, browser_name = launch_login_browser(playwright, headless=False)
        context = browser.new_context(locale="zh-CN", viewport={"width": 1440, "height": 900})
        page = context.new_page()
        page.goto(config["login_url"], wait_until="domcontentloaded")
        print("[INFO] Browser opened.")
        print(f"[INFO] Browser engine: {browser_name}")
        print("[INFO] Please log in on the opened website window.")
        print("[INFO] This script will save the session automatically once TOKEN is detected.")

        started = time.time()
        last_notice = 0.0
        while True:
            if page.is_closed():
                browser.close()
                raise SessionMissingError("The browser window was closed before a login session was captured.")
            elapsed = time.time() - started
            if elapsed > timeout_seconds:
                browser.close()
                raise SessionMissingError(
                    f"No usable login session was detected within {timeout_seconds // 60} minutes."
                )

            try:
                local_storage = extract_local_storage(page)
            except PlaywrightError:
                time.sleep(2)
                continue

            token = compact(local_storage.get("TOKEN", ""))
            refresh_token = compact(local_storage.get("refresh_token", ""))
            user_info = parse_json_maybe(local_storage.get("USERINFO", ""))
            if token or refresh_token:
                context.storage_state(path=str(storage_state_path()))
                try:
                    page.goto(config["query_page_url"], wait_until="networkidle", timeout=30000)
                    page.screenshot(path=str(screenshot_path()), full_page=True)
                except PlaywrightError:
                    pass
                browser.close()
                return save_session_payload(
                    config=config,
                    local_storage=local_storage,
                    user_info=user_info,
                    captured_at=captured_at,
                )

            if elapsed - last_notice >= 15:
                print(f"[INFO] Waiting for login... {int(elapsed)}s elapsed")
                last_notice = elapsed
            time.sleep(2)


def request_json(
    *,
    method: str,
    url: str,
    headers: dict[str, str],
    timeout_seconds: int,
    params: dict[str, Any] | None = None,
    data: dict[str, Any] | None = None,
) -> tuple[requests.Response, dict[str, Any]]:
    response = requests.request(
        method=method,
        url=url,
        headers=headers,
        params=params,
        data=data,
        timeout=timeout_seconds,
    )
    try:
        payload = response.json()
    except ValueError as exc:
        raise SiteWorkflowError(f"Site API returned non-JSON content from {url}") from exc
    return response, payload if isinstance(payload, dict) else {"raw": payload}


def validate_saved_session() -> dict[str, Any]:
    config = load_config()
    session = load_saved_session()
    check_url = config["api_base"] + config["check_endpoint"]

    attempts: list[dict[str, Any]] = []
    for header_token, _body_refresh_token in auth_candidates(session):
        headers = build_headers(config, header_token, config["login_url"])
        response, payload = request_json(
            method="get",
            url=check_url,
            headers=headers,
            timeout_seconds=int(config["request_timeout_seconds"]),
        )
        attempts.append(
            {
                "http_status": response.status_code,
                "code": payload.get("code"),
                "message": payload.get("message") or payload.get("meta", {}).get("message"),
                "used_header_token_suffix": header_token[-8:] if header_token else "",
            }
        )
        code = payload.get("code")
        if response.status_code == 200 and code not in (40101, 401, 403):
            result = {
                "checked_at": now_iso(),
                "valid": True,
                "user_label": extract_user_label(session.user_info),
                "captured_at": session.captured_at,
                "session_path": str(session.session_path),
                "storage_state_path": str(session.storage_state_path),
                "attempts": attempts,
            }
            write_json(check_report_json_path(), result)
            check_report_txt_path().write_text(render_check_report(result), encoding="utf-8")
            return result

    result = {
        "checked_at": now_iso(),
        "valid": False,
        "user_label": extract_user_label(session.user_info),
        "captured_at": session.captured_at,
        "session_path": str(session.session_path),
        "storage_state_path": str(session.storage_state_path),
        "attempts": attempts,
    }
    write_json(check_report_json_path(), result)
    check_report_txt_path().write_text(render_check_report(result), encoding="utf-8")
    return result


def build_query_payload(
    *,
    identifier: str,
    mode: str,
    iemark: str | None,
    refresh_token: str,
) -> dict[str, Any]:
    payload = {
        "autoIdentifyData": identifier,
        "blNo": "",
        "entryNo": "",
        "ctnrNo": "",
        "iEMark": iemark or "",
        "refreshToken": refresh_token,
        "page": "1",
        "rows": "10",
        "operationSource": "WEB",
    }
    payload[mode] = identifier
    return payload


def parse_result_list(payload: dict[str, Any]) -> list[dict[str, Any]]:
    data_block = payload.get("data")
    if not isinstance(data_block, dict):
        return []
    req_data = data_block.get("reqData")
    if not isinstance(req_data, dict):
        return []
    data_list = req_data.get("dataList")
    if not isinstance(data_list, list):
        return []
    return [item for item in data_list if isinstance(item, dict)]


def score_site_row(row: dict[str, Any]) -> tuple[int, int, int, int]:
    return (
        1 if is_released_status(row.get("cusLetpasStatus")) else 0,
        1 if compact(row.get("cusLetpasTime")) else 0,
        1 if compact(row.get("pkgNbr")) else 0,
        1 if compact(row.get("cargoGrossWeight")) else 0,
    )


def select_best_site_row(rows: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not rows:
        return None
    return sorted(rows, key=score_site_row, reverse=True)[0]


def query_site_identifier(
    *,
    identifier: str,
    mode: str = "auto",
    iemark: str | None = None,
) -> QuerySelection:
    config = load_config()
    session = load_saved_session()
    normalized_identifier = compact(identifier)
    if not normalized_identifier:
        raise SiteWorkflowError("The identifier is empty.")

    selected_mode = detect_identifier_mode(normalized_identifier) if mode == "auto" else mode
    if selected_mode not in {"blNo", "entryNo", "ctnrNo"}:
        raise SiteWorkflowError(f"Unsupported query mode: {selected_mode}")

    if selected_mode == "ctnrNo":
        candidate_iemarks: list[str | None] = [iemark or None]
    elif iemark:
        candidate_iemarks = [iemark]
    else:
        candidate_iemarks = [None, "E", "I"]

    query_url = config["api_base"] + config["query_endpoint"]
    attempts: list[dict[str, Any]] = []

    for header_token, body_refresh_token in auth_candidates(session):
        for try_iemark in candidate_iemarks:
            payload = build_query_payload(
                identifier=normalized_identifier,
                mode=selected_mode,
                iemark=try_iemark,
                refresh_token=body_refresh_token,
            )
            headers = build_headers(config, header_token, config["query_page_url"])
            response, body = request_json(
                method="post",
                url=query_url,
                headers=headers,
                timeout_seconds=int(config["request_timeout_seconds"]),
                data=payload,
            )
            result_rows = parse_result_list(body)
            attempts.append(
                {
                    "mode": selected_mode,
                    "iemark": try_iemark or "",
                    "http_status": response.status_code,
                    "result_count": len(result_rows),
                    "meta_code": body.get("meta", {}).get("code"),
                    "meta_message": body.get("meta", {}).get("message"),
                    "gateway_code": body.get("code"),
                    "gateway_message": body.get("message"),
                    "used_header_token_suffix": header_token[-8:] if header_token else "",
                    "used_refresh_token_suffix": body_refresh_token[-8:] if body_refresh_token else "",
                }
            )

            invalid_codes = {40101, 401, 403}
            if body.get("code") in invalid_codes or body.get("meta", {}).get("code") in invalid_codes:
                continue

            best_row = select_best_site_row(result_rows)
            if best_row is not None:
                status_text = compact(best_row.get("cusLetpasStatus", ""))
                return QuerySelection(
                    identifier=normalized_identifier,
                    mode=selected_mode,
                    iemark=try_iemark,
                    released=is_released_status(status_text),
                    status_text=status_text,
                    selected_row=best_row,
                    result_count=len(result_rows),
                    attempts=attempts,
                )

    return QuerySelection(
        identifier=normalized_identifier,
        mode=selected_mode,
        iemark=iemark,
        released=False,
        status_text="",
        selected_row=None,
        result_count=0,
        attempts=attempts,
    )


def to_number_or_blank(value: Any) -> int | float | str:
    text = compact(value)
    if not text:
        return ""
    try:
        number = float(text)
    except ValueError:
        return text
    if number.is_integer():
        return int(number)
    return round(number, 3)


def kg_to_mt(value: Any) -> int | float | str:
    number = to_number_or_blank(value)
    if number == "" or not isinstance(number, (int, float)):
        return ""
    mt = round(float(number) / 1000, 3)
    if mt.is_integer():
        return int(mt)
    return mt


def selection_to_query_row(selection: QuerySelection) -> dict[str, Any]:
    row = selection.selected_row or {}
    output = {header: "" for header in cw.QUERY_HEADERS}
    output[cw.QUERY_HEADERS[0]] = selection.identifier
    output[cw.QUERY_HEADERS[1]] = selection.status_text or "No result"
    output[cw.QUERY_HEADERS[2]] = "yes" if selection.released else "no"
    output[cw.QUERY_HEADERS[3]] = to_number_or_blank(row.get("pkgNbr", ""))
    output[cw.QUERY_HEADERS[4]] = kg_to_mt(row.get("cargoGrossWeight", ""))
    output[cw.QUERY_HEADERS[5]] = ""
    output[cw.QUERY_HEADERS[6]] = ""
    output[cw.QUERY_HEADERS[7]] = now_iso()
    note_parts = [
        f"mode={selection.mode}",
        f"iemark={selection.iemark or 'auto'}",
        f"entryNo={compact(row.get('entryNo', '')) or 'N/A'}",
        f"ctnrNo={compact(row.get('ctnrNo', '')) or 'N/A'}",
        f"pkgNbr={compact(row.get('pkgNbr', '')) or 'N/A'}",
        f"cargoGrossWeightKg={compact(row.get('cargoGrossWeight', '')) or 'N/A'}",
    ]
    output[cw.QUERY_HEADERS[8]] = "; ".join(note_parts)
    return output


def guess_iemark_from_business_row(row: dict[str, Any]) -> str | None:
    mapping = {
        "I": {"I", "IMPORT", "进口"},
        "E": {"E", "EXPORT", "出口"},
    }
    candidate_keys = {"I/E", "IE", "IEMARK", "IMPORTEXPORT", "进出口", "进出口标记"}
    for key, value in row.items():
        if compact(key).upper() not in candidate_keys:
            continue
        normalized = compact(value).upper()
        for target, values in mapping.items():
            if normalized in values:
                return target
    return None


def collect_pending_bls(workbook_path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(workbook_path)
    business_sheet = workbook[cw.BUSINESS_SHEET]
    rows = cw.read_business_rows(business_sheet)
    pending: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in rows:
        bl_no = compact(row.get("B/L NO.", ""))
        normalized_bl = cw.normalize_bl(bl_no)
        if not normalized_bl:
            continue
        if not cw.needs_lookup(row.get("备货", row.get("澶囪揣", ""))):
            continue
        if normalized_bl in seen:
            continue
        seen.add(normalized_bl)
        pending.append(
            {
                "bl_no": bl_no,
                "normalized_bl": normalized_bl,
                "iemark": guess_iemark_from_business_row(row),
            }
        )
    return pending


def site_report_paths(source_name: str) -> tuple[Path, Path]:
    stem = Path(source_name).stem
    return (
        SITE_QUERY_DIR / f"{stem}-site-report.json",
        SITE_QUERY_DIR / f"{stem}-site-report.txt",
    )


def render_check_report(report: dict[str, Any]) -> str:
    lines = [
        f"Checked at: {report['checked_at']}",
        f"Valid: {report['valid']}",
        f"User: {report.get('user_label', '') or 'Unknown'}",
        f"Captured at: {report.get('captured_at', '') or 'Unknown'}",
        f"Session file: {report.get('session_path', '')}",
        f"Storage state: {report.get('storage_state_path', '')}",
        "",
        "Attempts:",
    ]
    for attempt in report.get("attempts", []):
        lines.append(
            " - status={http_status}, code={code}, message={message}, token=*{used_header_token_suffix}".format(
                **attempt
            )
        )
    return "\n".join(lines) + "\n"


def render_site_report(report: dict[str, Any]) -> str:
    lines = [
        f"Run at: {report['run_at']}",
        f"Workbook: {report['source_workbook']}",
        f"Candidate BL count: {report['candidate_bl_count']}",
        f"Released count: {report['released_count']}",
        f"Pending count: {report['pending_count']}",
        f"No data count: {report['no_data_count']}",
        "",
        "Items:",
    ]
    for item in report.get("items", []):
        lines.append(
            " - BL={bl_no}; released={released}; status={status}; mode={mode}; iemark={iemark}; pcs={pcs}; mt={mt}".format(
                bl_no=item.get("bl_no", ""),
                released=item.get("released", False),
                status=item.get("status", "") or "N/A",
                mode=item.get("mode", ""),
                iemark=item.get("iemark", "") or "auto",
                pcs=item.get("pcs", "") if item.get("pcs", "") != "" else "N/A",
                mt=item.get("mt", "") if item.get("mt", "") != "" else "N/A",
            )
        )
    return "\n".join(lines) + "\n"


def query_and_update_workbook(workbook_path: Path, output_dir: Path) -> WorkbookSiteResult:
    SITE_QUERY_DIR.mkdir(parents=True, exist_ok=True)
    pending_bls = collect_pending_bls(workbook_path)
    items: list[dict[str, Any]] = []
    query_rows: list[dict[str, Any]] = []
    released_count = 0
    pending_count = 0
    no_data_count = 0

    for entry in pending_bls:
        selection = query_site_identifier(identifier=entry["bl_no"], mode="blNo", iemark=entry["iemark"])
        query_rows.append(selection_to_query_row(selection))

        selected_row = selection.selected_row or {}
        items.append(
            {
                "bl_no": entry["bl_no"],
                "released": selection.released,
                "status": selection.status_text,
                "mode": selection.mode,
                "iemark": selection.iemark or "",
                "result_count": selection.result_count,
                "attempts": selection.attempts,
                "selected_row": selected_row,
                "pcs": to_number_or_blank(selected_row.get("pkgNbr", "")),
                "kg": to_number_or_blank(selected_row.get("cargoGrossWeight", "")),
                "mt": kg_to_mt(selected_row.get("cargoGrossWeight", "")),
            }
        )

        if selection.selected_row is None:
            no_data_count += 1
        elif selection.released:
            released_count += 1
        else:
            pending_count += 1

    site_summary = {
        "run_at": now_iso(),
        "source_workbook": workbook_path.name,
        "candidate_bl_count": len(pending_bls),
        "released_count": released_count,
        "pending_count": pending_count,
        "no_data_count": no_data_count,
        "items": items,
    }
    site_json_path, site_txt_path = site_report_paths(workbook_path.name)
    write_json(site_json_path, site_summary)
    site_txt_path.write_text(render_site_report(site_summary), encoding="utf-8")

    clearance_result = cw.update_clearance_workbook(
        workbook_path=workbook_path,
        output_dir=output_dir,
        query_rows_override=query_rows,
    )
    return WorkbookSiteResult(
        clearance_result=clearance_result,
        site_report_json_path=site_json_path.resolve(),
        site_report_txt_path=site_txt_path.resolve(),
        site_summary=site_summary,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    parser_capture = subparsers.add_parser("capture-session", help="Open the login page and save the req2 site session.")
    parser_capture.add_argument("--timeout-minutes", type=int, help="How long to wait for a manual login.")

    subparsers.add_parser("check-session", help="Check whether the saved req2 site session is still valid.")

    parser_query = subparsers.add_parser("query-one", help="Query one BL / entry / container number from the real site.")
    parser_query.add_argument("--identifier", required=True, help="BL number, customs declaration number, or container number.")
    parser_query.add_argument("--mode", choices=["auto", "blNo", "entryNo", "ctnrNo"], default="auto")
    parser_query.add_argument("--iemark", choices=["I", "E"], help="Optional import/export flag.")

    parser_workbook = subparsers.add_parser(
        "from-workbook",
        help="Query the real site for pending BLs in a req2 workbook, then update the workbook.",
    )
    parser_workbook.add_argument("--input", required=True, help="Path to the req2 workbook.")
    parser_workbook.add_argument("--output-dir", default=str(cw.UPDATED_DIR), help="Where to write updated workbooks.")

    return parser


def print_session_summary(session: SavedSession) -> None:
    print("[OK] Session saved.")
    print(f"[INFO] Session JSON: {session.session_path}")
    print(f"[INFO] Storage state: {session.storage_state_path}")
    print(f"[INFO] Screenshot: {screenshot_path()}")
    print(f"[INFO] Captured at: {session.captured_at}")
    print(f"[INFO] User: {extract_user_label(session.user_info) or 'Unknown'}")


def print_query_one(selection: QuerySelection) -> None:
    print(f"[INFO] Identifier: {selection.identifier}")
    print(f"[INFO] Mode: {selection.mode}")
    print(f"[INFO] I/E: {selection.iemark or 'auto'}")
    print(f"[INFO] Result count: {selection.result_count}")
    if selection.selected_row is None:
        print("[WARN] No site result was found for this identifier.")
    else:
        row = selection.selected_row
        print(f"[INFO] Release status: {selection.status_text or 'N/A'}")
        print(f"[INFO] Released: {selection.released}")
        print(f"[INFO] Entry No: {compact(row.get('entryNo', '')) or 'N/A'}")
        print(f"[INFO] BL No: {compact(row.get('blNo', '')) or 'N/A'}")
        print(f"[INFO] Container No: {compact(row.get('ctnrNo', '')) or 'N/A'}")
        print(f"[INFO] PCS: {to_number_or_blank(row.get('pkgNbr', '')) or 'N/A'}")
        print(f"[INFO] Gross Weight (kg): {to_number_or_blank(row.get('cargoGrossWeight', '')) or 'N/A'}")
        print(f"[INFO] MT (converted): {kg_to_mt(row.get('cargoGrossWeight', '')) or 'N/A'}")
        print(f"[INFO] Release time: {compact(row.get('cusLetpasTime', '')) or 'N/A'}")


def print_workbook_result(result: WorkbookSiteResult) -> None:
    print("[OK] Req2 workbook updated from the real site.")
    print(f"[INFO] Updated workbook: {result.clearance_result.workbook_path}")
    print(f"[INFO] Clearance report JSON: {result.clearance_result.report_json_path}")
    print(f"[INFO] Clearance report TXT: {result.clearance_result.report_txt_path}")
    print(f"[INFO] Site report JSON: {result.site_report_json_path}")
    print(f"[INFO] Site report TXT: {result.site_report_txt_path}")
    print(
        "[INFO] Site summary: candidates={candidate_bl_count}, released={released_count}, pending={pending_count}, no_data={no_data_count}".format(
            **result.site_summary
        )
    )


def main(argv: list[str] | None = None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.command == "capture-session":
            session = capture_session(timeout_minutes=args.timeout_minutes)
            print_session_summary(session)
            return 0

        if args.command == "check-session":
            report = validate_saved_session()
            if report["valid"]:
                print("[OK] Saved req2 site session is valid.")
            else:
                print("[ERROR] Saved req2 site session is not valid anymore.")
                print("[INFO] Please run capture-session again.")
                return 1
            print(f"[INFO] Check report JSON: {check_report_json_path()}")
            print(f"[INFO] Check report TXT: {check_report_txt_path()}")
            return 0

        if args.command == "query-one":
            selection = query_site_identifier(
                identifier=args.identifier,
                mode=args.mode,
                iemark=args.iemark,
            )
            print_query_one(selection)
            return 0 if selection.selected_row is not None else 1

        if args.command == "from-workbook":
            result = query_and_update_workbook(
                workbook_path=Path(args.input),
                output_dir=Path(args.output_dir),
            )
            print_workbook_result(result)
            return 0

        raise SiteWorkflowError(f"Unsupported command: {args.command}")
    except SessionMissingError as exc:
        print(f"[ERROR] {exc}")
        return 1
    except requests.RequestException as exc:
        print(f"[ERROR] Site request failed: {exc}")
        return 1
    except SiteWorkflowError as exc:
        print(f"[ERROR] {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
