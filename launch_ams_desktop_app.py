from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from desktop_app.app import main as gui_main
from desktop_app.runtime import AmsOperations, AppConfig, release_assets_dir, skill_root


def run_self_test(output_path: Path, workspace_root: Path | None = None) -> int:
    config = AppConfig.default() if workspace_root is None else AppConfig(workspace_root=str(workspace_root.resolve()))
    ops = AmsOperations(config)
    info = ops.ensure_workspace()

    example_dir = skill_root() / "examples" / "workbooks"
    example_workbook = next(example_dir.glob("domestic-forwarder-*.xlsx"))
    req1_result = ops.req1_generate_from_file(example_workbook)

    report = {
        "success": True,
        "frozen": bool(getattr(sys, "frozen", False)),
        "workspace_root": info["workspace_root"],
        "req1_input_path": info["req1_input_path"],
        "req2_input_path": info["req2_input_path"],
        "req1_document_path": req1_result["document_path"],
        "req1_summary_path": req1_result["summary_path"],
        "req1_latest_document_path": req1_result["latest_document_path"],
        "req1_latest_summary_path": req1_result["latest_summary_path"],
        "req1_latest_document_exists": Path(req1_result["latest_document_path"]).exists(),
        "req1_latest_summary_exists": Path(req1_result["latest_summary_path"]).exists(),
        "guide_exists": (release_assets_dir() / "应用使用说明.html").exists(),
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--self-test-output")
    parser.add_argument("--self-test-workspace")
    return parser


def main() -> int:
    parser = build_parser()
    args, _unknown = parser.parse_known_args()
    if args.self_test:
        output_path = Path(args.self_test_output or "ams-desktop-self-test.json").resolve()
        workspace_root = Path(args.self_test_workspace).resolve() if args.self_test_workspace else None
        return run_self_test(output_path, workspace_root)

    gui_main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
