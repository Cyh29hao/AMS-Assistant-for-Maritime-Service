﻿---
name: maritime-service
description: Use when handling maritime-service operations, planning shipping-service workflows, triaging vessel or document tasks, coordinating vendors, or designing automations for recurring service requests.
---

# Maritime Service

## Overview

Use this skill to turn a maritime-service request into a controlled workflow that an agent can plan, execute with tools, and escalate for approval when needed. The skill is intentionally broad in its first version so it can support discovery, semi-automation, and later full automation.

The first production-leaning workflow in this skill is `req1 系统出合同`, implemented as a Word-contract generation workflow with both JSON and Excel entrypoints.

The second production-leaning workflow is `req2 自动查通关并回填表格`, implemented as an Excel update workflow that reads business rows, evaluates customs-clearance query results, updates eligible rows to `已通关`, and emits structured reports.

The latest extension of `req2` is a website session layer for `dp.eptrade.cn`, which can capture a reusable login session, validate it, query the real `pss026` maritime-release endpoint by B/L number, and feed those results back into the accepted req2 workbook-update flow.

## When To Use

Use this skill when the task involves any of the following:

- triaging incoming service requests from customers, agents, or internal teams
- coordinating a vessel, port, vendor, or service milestone
- checking the readiness of documents, data, or operational prerequisites
- producing daily, weekly, or exception-based operational reports
- deciding whether a workflow should be manual, semi-automatic, or fully automated
- designing or extending tools for a maritime operation

## Core Operating Model

Treat the agent as an operations coordinator, not just a chat assistant.

For each request:

1. Identify the request type, urgency, owner, and expected deliverable.
2. Gather the minimum context required from tools or provided files.
3. Choose the correct workflow from `references/workflow-catalog.md`.
4. Execute only the steps allowed by the available tool permissions.
5. Stop for approval before irreversible, financial, regulatory, or customer-facing actions.
6. Record the result, blockers, and next action in a structured handoff.

## Automation Modes

Choose the lightest mode that is still useful:

- Copilot mode: the agent drafts plans, emails, reports, and checklists for a human operator.
- Semi-automatic mode: the agent reads systems, prepares updates, and executes low-risk write actions after rule checks.
- Full automation mode: the agent runs end-to-end for narrow, well-instrumented workflows with explicit approval gates.

Prefer starting in Copilot mode and only moving a workflow upward after it is stable.

## Tool Strategy

Prefer direct system access over GUI imitation:

- First choice: API, database, email API, file store, or internal scripts
- Second choice: browser automation or computer-use tooling for systems with no API
- Last choice: ad hoc manual copying across screens

Tool contracts, safety rules, and naming conventions live in `references/tool-contracts.md`.

## Workflow Template

Apply this template to every new automation:

1. Intake: normalize the request into a structured task.
2. Classification: map the task to a workflow family.
3. Context collection: fetch vessel, customer, vendor, document, and timing data.
4. Decision: determine whether the task can proceed automatically.
5. Action: call tools, draft outputs, or update systems.
6. Approval: require human confirmation for guarded steps.
7. Completion: produce a concise audit trail and next-step summary.

## Reference Map

Read only what you need:

- `references/operating-model.md` for architecture, platform choices, and execution patterns
- `references/workflow-catalog.md` for example workflow families and sample deliverables
- `references/document-readiness-check.md` for the first implemented real subprocess
- `references/contract-generation.md` for the implemented contract-generation workflow
- `references/clearance-status-update.md` for the implemented req2 clearance-update workflow
- `docs/11-req2网页登录使用教程.md` for the real-site login and query workflow
- `docs/12-req2网页登录如何验收.md` for the req2 website-layer acceptance steps
- `references/tool-contracts.md` for future API and tool design rules
- `references/build-roadmap.md` for how to grow this skill from prototype to production

## Scripts

- `scripts/workflow_demo.py` builds a sample execution plan from a task payload
- `scripts/validate_document_bundle.py` runs the first real document-readiness check
- `scripts/contract_workflow.py` runs the real contract-generation workflow
- `scripts/clearance_workflow.py` runs the req2 customs-clearance update workflow
- `scripts/example-task.json` is a starter input for local testing
- `scripts/example-document-task.json` is a starter document-bundle payload
- `scripts/example-document-task-ready.json` is a passing payload for local verification

Use the demo script to pressure-test the workflow design before wiring in real maritime systems.

## First Real Subprocess

The first implemented subprocess is `document_readiness_check`.

Use it when the agent needs to answer:

- Is the document bundle complete for this service type?
- Are any required fields missing?
- Are any documents expired before the operational deadline?
- Which follow-up items should be sent to the operator or customer?

Current local entrypoint:

```bash
python scripts/validate_document_bundle.py --task-file scripts/example-document-task.json
```

Treat this subprocess as the baseline pattern for future workflows:

1. define a narrow business objective
2. define explicit input schema
3. implement deterministic checks
4. return structured output plus next actions
5. add approval gates before any external submission

## Real Workflow: Req1 System Contract Generation

Use this workflow when the operator needs to:

- choose one of several contract templates
- fill a bilingual fixture note from business data
- preserve the Word layout instead of generating plain text
- produce a deliverable that can be checked and sent quickly

Current local entrypoints:

```bash
python scripts/contract_workflow.py build-examples
python scripts/contract_workflow.py from-json --input examples/contract_requests/domestic-forwarder-mu-chuang-362.json
python scripts/contract_workflow.py from-workbook --input examples/workbooks/blank-contract-template.xlsx
```

This workflow is designed for two audiences:

- operators who prefer filling an Excel workbook
- developers or agents who prefer structured JSON

See `docs/00-如何验收这套skill.md` and `docs/01-小学生级别使用教程.md` for human-facing instructions.

## Real Workflow: Req2 Customs Clearance Update

Use this workflow when the operator needs to:

- identify uncleared bills of lading in an operational workbook
- compare them against customs-clearance query results
- update `备货` to `已通关` only when the result is actually released
- backfill `PCS` and `MT`
- preserve an auditable workbook and report trail

Current local entrypoints:

```bash
python scripts/clearance_workflow.py build-examples
python scripts/clearance_workflow.py verify-examples
python scripts/clearance_workflow.py from-workbook --input examples/clearance_workbooks/blank-clearance-template.xlsx
```

This workflow currently focuses on the stable half of req2:

- workbook normalization
- deterministic update rules
- output reporting

The live website-query layer is intentionally left as the next integration step, because its selectors, login stability, and anti-bot constraints still need real-environment verification.

See `docs/09-req2-自动查通关使用教程.md` and `docs/10-req2-如何验收.md` for operator-facing instructions.
