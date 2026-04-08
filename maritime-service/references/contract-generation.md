# Contract Generation

This reference describes the first major business workflow implemented in the skill: `req1 系统出合同`.

## Purpose

Generate bilingual Word fixture-note contracts from structured business data while preserving the source template layout.

## Implemented Entry Modes

- JSON request mode
- Excel workbook mode
- example build mode for acceptance and regression checks

## Files

- `scripts/contract_workflow.py` -> main CLI
- `scripts/contract_template_registry.json` -> template registry and paragraph coordinates
- `assets/contract_templates/` -> source `.docx` templates
- `examples/contract_requests/` -> sample JSON requests
- `examples/workbooks/` -> generated workbook examples and blank workbook template
- `output/contracts/` -> generated `.docx` outputs
- `output/reports/` -> generation summaries

## Supported Template Keys

- `domestic_forwarder`
- `domestic_logistics`
- `overseas_astar`
- `overseas_rta`

## How It Works

1. Validate the structured request.
2. Load the matching template from the registry.
3. Compute derived fields such as contract number, English date, laycan text, demurrage text, and freight text.
4. Replace the targeted Word paragraphs using fixed template coordinates.
5. Rebuild the cargo table for domestic templates.
6. Save the `.docx` and a JSON summary side by side.

## Why This Design

The provided source templates are not placeholder templates such as `{{vessel_name}}`. They are visual Word templates whose variable content lives in fixed paragraphs or tables. For that reason, the current implementation uses template coordinates rather than a generic mail-merge engine.

This is the right tradeoff for the current materials because it:

- works with the provided templates immediately
- avoids asking the business side to redesign templates before using the tool
- keeps output layout close to the original samples

## Current Boundaries

- The workflow assumes each template family keeps a stable paragraph structure.
- If the business team later edits a template heavily, the registry coordinates may need to be updated.
- The workflow does not yet parse a raw business spreadsheet automatically; it expects either a prepared JSON request or the provided workbook format.

## Recommended Future Upgrade

When the business is ready, migrate from coordinate-based templates to placeholder-based templates. That will make template maintenance easier and reduce the need to adjust paragraph indexes after Word edits.
