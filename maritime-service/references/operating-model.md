# Operating Model

## What This Skill Is

This skill is the operational playbook for a maritime-service agent. It tells the agent:

- what kinds of work belong in scope
- how to break work into safe steps
- when to use tools
- when to stop for human approval

The skill is not the automation engine by itself. Real automation appears when this skill is combined with:

- an agent runtime such as Codex or an API-based agent
- tools such as APIs, shell scripts, email connectors, or browser automation
- a trigger such as a user request, inbox item, schedule, or webhook

## Recommended Architecture

Use this layered model:

1. Trigger layer
   Examples: a dispatcher email, a CRM event, a spreadsheet update, a daily schedule, or a manual request from an operator.
2. Agent layer
   Interprets the request, loads this skill, chooses the workflow, and decides what to do next.
3. Tool layer
   Reads or writes data through APIs, scripts, databases, documents, or browsers.
4. Guardrail layer
   Applies validations, approval gates, retry rules, and audit logging.
5. Output layer
   Produces a status update, drafted email, report, system update, or escalation note.

## Platform Choices

### Codex

Codex is a strong fit when you want an agent that can:

- use local files and scripts
- integrate with APIs and structured tools
- iteratively improve the workflow in the same workspace
- evolve from prototype to real automation

### ChatGPT

ChatGPT is a good fit for front-end interaction when you want:

- a conversational assistant for operators
- fast drafting, analysis, and decision support
- lightweight integrations through actions or external tools

It is usually not the first choice for deep local workflow engineering unless paired with back-end tooling.

### GUI automation or computer-use tools

Use browser or desktop automation only when the target system has no stable API or structured integration path. If an API exists, prefer the API.

## Safety Model

Treat these as guarded actions:

- sending a customer-facing commitment
- creating or changing a financial record
- submitting a regulatory filing
- booking, cancelling, or changing a service order
- updating a source-of-truth operational system

For guarded actions, the agent should prepare the action, show the reasoning, and wait for approval unless the workflow has already been certified for automatic execution.

## How Work Should Progress

Start with a narrow loop:

1. Choose one workflow family.
2. Define the input schema.
3. Define the success output.
4. Add read-only tools first.
5. Add one low-risk write action.
6. Observe failures and add rules.
7. Expand to the next workflow family.

This keeps the skill useful early while making later automation safer.
