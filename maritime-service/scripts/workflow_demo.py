#!/usr/bin/env python3
"""Build a sample maritime-service execution plan from a task payload."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


WORKFLOWS = {
    "service_request_triage": {
        "keywords": ["request", "inquiry", "quote", "new service"],
        "specificity": 0,
        "steps": [
            "Normalize the request into a structured ticket",
            "Identify missing fields and service urgency",
            "Assign owner and recommend next action",
        ],
        "approvals": [],
        "suggested_tools": ["read_inbox_message", "lookup_customer_profile"],
    },
    "port_call_support": {
        "keywords": ["port", "eta", "etd", "berth", "arrival", "departure"],
        "specificity": 1,
        "steps": [
            "Collect vessel timing and port constraints",
            "Check readiness of vendors and dependencies",
            "Prepare execution checklist and escalation notes",
        ],
        "approvals": ["Any booking, cancellation, or operational commitment"],
        "suggested_tools": ["lookup_vessel_schedule", "lookup_vendor_contacts"],
    },
    "document_readiness_check": {
        "keywords": [
            "document",
            "document bundle",
            "certificate",
            "manifest",
            "filing",
            "paperwork",
            "ready",
            "readiness",
        ],
        "specificity": 1,
        "steps": [
            "List required documents for the service",
            "Compare provided bundle against required fields",
            "Draft exception summary and follow-up request",
        ],
        "approvals": ["Any formal submission to an authority or external party"],
        "suggested_tools": ["validate_document_bundle", "draft_follow_up_email"],
    },
    "vendor_coordination": {
        "keywords": ["vendor", "supplier", "confirm", "booking", "agent"],
        "specificity": 1,
        "steps": [
            "Determine vendor category and service window",
            "Prepare request or follow-up message",
            "Track response status and escalation timing",
        ],
        "approvals": ["Sending a commitment that changes scope, price, or timing"],
        "suggested_tools": ["lookup_vendor_contacts", "draft_vendor_message"],
    },
    "daily_ops_reporting": {
        "keywords": ["report", "daily", "ops", "status", "summary", "exception"],
        "specificity": 1,
        "steps": [
            "Gather open tasks and current exceptions",
            "Summarize important movements and blockers",
            "Produce a concise leadership-ready report",
        ],
        "approvals": [],
        "suggested_tools": ["query_ops_dashboard", "generate_ops_report"],
    },
}


def load_task(path: Path | None, request_text: str | None, task_type: str | None) -> dict:
    if path is not None:
        return json.loads(path.read_text(encoding="utf-8"))

    return {
        "request_text": request_text or "",
        "task_type": task_type or "",
        "priority": "normal",
        "requires_external_write": False,
    }


def classify(task: dict) -> str:
    explicit_type = task.get("task_type", "").strip().lower()
    if explicit_type in WORKFLOWS:
        return explicit_type

    haystack = " ".join(
        [
            str(task.get("request_text", "")),
            str(task.get("service_type", "")),
            str(task.get("notes", "")),
        ]
    ).lower()

    best_key = "service_request_triage"
    best_score = (-1, -1)

    for workflow_key, config in WORKFLOWS.items():
        keyword_score = 0
        for keyword in config["keywords"]:
            if keyword in haystack:
                keyword_score += 2 if " " in keyword else 1

        score = (keyword_score, config["specificity"])
        if score > best_score:
            best_key = workflow_key
            best_score = score

    return best_key


def build_plan(task: dict) -> dict:
    workflow_key = classify(task)
    workflow = WORKFLOWS[workflow_key]

    approvals = list(workflow["approvals"])
    if task.get("requires_external_write"):
        approvals.append("Human approval before any external write action")

    return {
        "workflow": workflow_key,
        "priority": task.get("priority", "normal"),
        "task_summary": task.get("request_text", "").strip() or "No request text provided",
        "steps": workflow["steps"],
        "suggested_tools": workflow["suggested_tools"],
        "approval_gates": approvals,
        "next_actions": [
            "Confirm the workflow family is correct",
            "Replace demo tools with real system integrations",
            "Add validation rules for the chosen workflow",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--task-file", type=Path, help="Path to a JSON task payload")
    parser.add_argument("--request-text", help="Short free-form request description")
    parser.add_argument("--task-type", help="Explicit workflow key to force classification")
    args = parser.parse_args()

    task = load_task(args.task_file, args.request_text, args.task_type)
    plan = build_plan(task)
    print(json.dumps(plan, indent=2))


if __name__ == "__main__":
    main()
