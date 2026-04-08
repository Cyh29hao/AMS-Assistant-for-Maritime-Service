# Workflow Catalog

This file lists broad workflow families for the first version of the skill. Replace or expand them with your real business processes over time.

## 1. Service Request Triage

Objective: turn an incoming request into a structured operational ticket.

Typical inputs:

- customer email or chat message
- requested service type
- vessel or voyage reference
- target port and timing

Expected outputs:

- normalized task record
- urgency level
- missing-information checklist
- recommended owner

Typical tools:

- email reader
- CRM or ticket system
- customer master data lookup

## 2. Port Call Support

Objective: coordinate readiness for a port-related service milestone.

Typical inputs:

- vessel identifier
- ETA or ETD
- required service package
- local port constraints

Expected outputs:

- readiness summary
- dependency checklist
- vendor coordination tasks
- operator escalation if prerequisites are missing

Typical tools:

- vessel schedule API
- vendor directory
- internal task board

## 3. Document Readiness Check

Objective: validate whether required operational documents are complete, current, and internally consistent.

Status: first implemented subprocess in this skill. See `references/document-readiness-check.md`.

Typical inputs:

- document bundle
- vessel or voyage reference
- service type
- filing deadline

Expected outputs:

- pass or fail status
- exception list
- missing fields
- draft follow-up message

Typical tools:

- file parser
- document schema checker
- template library
- `scripts/validate_document_bundle.py` for local deterministic validation

## 4. Vendor Coordination

Objective: prepare or execute communications and follow-ups with third-party providers.

Typical inputs:

- vendor type
- service request
- service window
- pricing or scope context

Expected outputs:

- shortlist of vendors
- draft request for confirmation
- follow-up schedule
- decision record

Typical tools:

- vendor database
- email or messaging connector
- pricing reference data

## 5. Daily Operations Reporting

Objective: summarize operational state for leaders or dispatchers.

Typical inputs:

- open tasks
- active vessels or voyages
- exceptions, delays, and blockers
- turnaround metrics

Expected outputs:

- concise daily report
- exception highlights
- owner-by-owner follow-up list

Typical tools:

- database query
- spreadsheet connector
- reporting script

## Selection Rule

When a request arrives, first assign it to one workflow family. Only create a new family when the request cannot be handled by extending an existing one.

## 6. Contract Generation

Objective: generate a bilingual Word contract from structured business inputs and a selected Word template.

Status: implemented and usable as the current main workflow. See `references/contract-generation.md`.

Typical inputs:

- template key
- contract header fields
- laycan and pricing fields
- one or more cargo-detail rows

Expected outputs:

- generated `.docx` contract
- generation summary JSON
- optional operator workbook for future reuse

Typical tools:

- `python-docx`
- `openpyxl`
- `scripts/contract_workflow.py`
