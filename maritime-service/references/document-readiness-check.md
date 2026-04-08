# Document Readiness Check

This is the first real subprocess implemented in the skill.

## Purpose

Determine whether a service-specific document bundle is ready for operational use before submission, filing, or execution.

## Supported Service Types

Current starter service types:

- `port-clearance`
- `crew-change`
- `stores-delivery`

Add more service types by extending `scripts/document_requirements.json`.

## Input Contract

The validator expects a JSON payload with fields like:

- `request_id`
- `service_type`
- `deadline`
- `vessel.name`
- `vessel.imo_number`
- `documents[]`

Each document should contain:

- `type`
- `provided`
- `fields`
- optional `expiry_date`

## Checks Performed

The current subprocess performs four categories of checks:

1. Presence check
   Ensures every required document type exists and is marked as provided.
2. Field completeness check
   Ensures mandatory fields exist inside each document payload.
3. Expiry check
   Flags documents that expire before the service deadline.
4. Cross-document consistency check
   Confirms selected vessel identifiers match the top-level task.

## Output Contract

The validator returns structured JSON with:

- `ready`: overall pass or fail
- `status`: `ready`, `needs_review`, or `not_ready`
- `missing_documents`
- `missing_fields`
- `expired_documents`
- `consistency_issues`
- `follow_up_actions`

## Approval Rule

Passing validation does not mean the agent may submit documents externally.

The agent may:

- summarize the result
- prepare a follow-up email
- prepare a submission package

The agent must still stop for approval before:

- filing to an authority
- sending a formal external commitment
- updating a source-of-truth record outside a preapproved workflow

## How To Extend

When you move from example logic to real business logic, extend in this order:

1. replace starter service types with your actual service catalog
2. refine required fields per document type
3. add date and format validation rules
4. add source-system lookups for vessel, voyage, and customer consistency
5. connect the output to draft emails, tickets, or approval tasks
