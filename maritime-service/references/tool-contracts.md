# Tool Contracts

Future maritime-service tools should follow these rules.

## Design Rules

- One tool should do one clear business action.
- Inputs should be explicit JSON fields, not free-form guesses.
- Outputs should be structured and machine-readable first, human-readable second.
- Prefer idempotent actions so retries do not create duplicates.
- Every write tool should support a dry-run or preview mode when practical.

## Naming Pattern

Use verb-led tool names such as:

- `lookup_vessel_schedule`
- `validate_document_bundle`
- `draft_vendor_message`
- `create_service_order`
- `post_ops_report`

## Minimum Fields

Every write-oriented tool should expose:

- `request_id`
- `actor`
- `reason`
- `dry_run`
- `source_reference`

Every tool result should include:

- `status`
- `summary`
- `artifacts`
- `warnings`
- `next_actions`

## Approval Levels

- Level 0: read-only actions and draft generation
- Level 1: internal low-risk updates
- Level 2: customer-facing or vendor-facing actions
- Level 3: financial, regulatory, or source-of-truth updates

By default, only Level 0 should run unattended in a new workflow.

## Failure Handling

If a tool fails:

1. capture the input and error summary
2. classify whether the failure is retryable
3. avoid partial duplicate writes
4. produce a human handoff with a proposed next step

## Logging

Keep a lightweight audit trail for each run:

- task identifier
- workflow family
- tools called
- approval decisions
- final outcome
