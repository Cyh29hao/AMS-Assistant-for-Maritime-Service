#!/usr/bin/env python3
"""Validate a maritime document bundle against starter service-specific rules."""

from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
REQUIREMENTS_PATH = SCRIPT_DIR / "document_requirements.json"


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    return date.fromisoformat(value)


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def index_documents(documents: list[dict]) -> dict[str, dict]:
    indexed: dict[str, dict] = {}
    for document in documents:
        document_type = str(document.get("type", "")).strip()
        if document_type:
            indexed[document_type] = document
    return indexed


def validate_presence(requirements: dict, documents_by_type: dict[str, dict]) -> list[str]:
    missing_documents: list[str] = []
    for required in requirements["required_documents"]:
        document = documents_by_type.get(required["type"])
        if not document or not document.get("provided"):
            missing_documents.append(required["type"])
    return missing_documents


def validate_required_fields(requirements: dict, documents_by_type: dict[str, dict]) -> list[dict]:
    missing_fields: list[dict] = []
    for required in requirements["required_documents"]:
        document = documents_by_type.get(required["type"])
        if not document or not document.get("provided"):
            continue

        fields = document.get("fields", {})
        for field_name in required["required_fields"]:
            field_value = fields.get(field_name)
            if field_value in ("", None, [], {}):
                missing_fields.append({"document_type": required["type"], "field": field_name})
    return missing_fields


def validate_expiry(requirements: dict, documents_by_type: dict[str, dict], deadline: date | None) -> list[dict]:
    expired_documents: list[dict] = []
    if deadline is None:
        return expired_documents

    for required in requirements["required_documents"]:
        document = documents_by_type.get(required["type"])
        if not document or not document.get("provided"):
            continue

        expiry = parse_date(document.get("expiry_date"))
        if expiry is not None and expiry < deadline:
            expired_documents.append(
                {
                    "document_type": required["type"],
                    "expiry_date": expiry.isoformat(),
                    "deadline": deadline.isoformat(),
                }
            )
    return expired_documents


def validate_consistency(task: dict, documents_by_type: dict[str, dict]) -> list[dict]:
    vessel = task.get("vessel", {})
    vessel_name = vessel.get("name")
    imo_number = vessel.get("imo_number")
    consistency_issues: list[dict] = []

    particulars = documents_by_type.get("vessel-particulars")
    if particulars and particulars.get("provided"):
        fields = particulars.get("fields", {})
        if vessel_name and fields.get("vessel_name") and fields["vessel_name"] != vessel_name:
            consistency_issues.append(
                {
                    "document_type": "vessel-particulars",
                    "field": "vessel_name",
                    "expected": vessel_name,
                    "actual": fields["vessel_name"],
                }
            )
        if imo_number and fields.get("imo_number") and fields["imo_number"] != imo_number:
            consistency_issues.append(
                {
                    "document_type": "vessel-particulars",
                    "field": "imo_number",
                    "expected": imo_number,
                    "actual": fields["imo_number"],
                }
            )

    return consistency_issues


def build_follow_up_actions(result: dict) -> list[str]:
    actions: list[str] = []

    if result["missing_documents"]:
        actions.append("Request the missing documents from the operator or customer.")
    if result["missing_fields"]:
        actions.append("Ask for the missing document fields before submission.")
    if result["expired_documents"]:
        actions.append("Replace or renew documents that expire before the deadline.")
    if result["consistency_issues"]:
        actions.append("Resolve vessel-identity inconsistencies across the bundle.")
    if not actions:
        actions.append("Bundle is ready for operator review and submission preparation.")

    return actions


def evaluate(task: dict, requirements_map: dict) -> dict:
    service_type = task.get("service_type")
    if service_type not in requirements_map:
        raise ValueError(f"Unsupported service_type: {service_type!r}")

    requirements = requirements_map[service_type]
    documents = task.get("documents", [])
    documents_by_type = index_documents(documents)
    deadline = parse_date(task.get("deadline"))

    result = {
        "request_id": task.get("request_id"),
        "service_type": service_type,
        "deadline": task.get("deadline"),
        "ready": False,
        "status": "not_ready",
        "missing_documents": validate_presence(requirements, documents_by_type),
        "missing_fields": validate_required_fields(requirements, documents_by_type),
        "expired_documents": validate_expiry(requirements, documents_by_type, deadline),
        "consistency_issues": validate_consistency(task, documents_by_type),
    }

    blockers = (
        len(result["missing_documents"])
        + len(result["missing_fields"])
        + len(result["expired_documents"])
        + len(result["consistency_issues"])
    )

    if blockers == 0:
        result["ready"] = True
        result["status"] = "ready"
    elif result["missing_documents"] or result["expired_documents"]:
        result["status"] = "not_ready"
    else:
        result["status"] = "needs_review"

    result["follow_up_actions"] = build_follow_up_actions(result)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--task-file", type=Path, required=True, help="Path to the JSON task payload")
    args = parser.parse_args()

    task = load_json(args.task_file)
    requirements_map = load_json(REQUIREMENTS_PATH)
    result = evaluate(task, requirements_map)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
