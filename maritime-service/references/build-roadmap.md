# Build Roadmap

Use this roadmap to turn the current framework into a real operational system.

## Phase 1: Domain Definition

Goal: describe the business clearly before automating it.

Add:

- glossary of maritime terms used by your team
- list of service types
- entity map for vessel, voyage, customer, vendor, port, and document
- example requests and expected outcomes

## Phase 2: Assisted Operations

Goal: make the agent useful before it can write to systems.

Add:

- intake templates
- report templates
- exception checklists
- read-only connectors to your existing systems

Success measure:

- operators save time on triage, drafting, and reporting

## Phase 3: Semi-Automation

Goal: let the agent perform low-risk actions with review.

Add:

- write tools for internal updates
- approval prompts for guarded actions
- deterministic validation scripts
- basic audit logs

Success measure:

- one or two workflows can run with only a quick human approval step

## Phase 4: Controlled Full Automation

Goal: automate a narrow set of repetitive workflows end to end.

Add:

- event triggers
- retry logic
- idempotency keys
- service-level alerts
- operational dashboards

Success measure:

- selected workflows run automatically with low exception rates and clean audit trails

## Practical Next Step

Pick one workflow family from `workflow-catalog.md` and answer four questions:

1. What starts the workflow?
2. What facts must be gathered?
3. What output marks success?
4. Which step is the first safe automation target?

Do not automate five workflows halfway. Automate one workflow properly.
