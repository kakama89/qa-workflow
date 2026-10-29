# Characterization Test

A characterization test records what an existing system **does today**, not
what it should do. Use it as a regression baseline when requirements are
missing, outdated, or untrusted: before refactoring, migrating, upgrading a
dependency, replacing a legacy module, or taking over a system nobody
documented.

The Expected Result of a characterization case is the **observed current
behavior**, backed by evidence. Observed behavior is not proof of correct
behavior: a baseline can capture bugs, and that is intended as long as they
are labeled.

## When to use this mode

Switch to it when the tester asks for characterization / baseline /
"chụp hành vi hiện tại" / "test hồi quy trước khi refactor/migrate", or when
Step 2 shows there is no usable requirement for an existing feature and the
tester agrees to baseline it. Say which mode is active. Requirement-based and
characterization cases can coexist in one run; keep them in separate sections.

## Step 1 — Intake

Ask for (only what is missing):

- System under test: feature/module, entry points (screens, API endpoints,
  batch jobs, reports), and what is out of scope.
- Build/version, environment (Dev/Test/UAT…), and date of observation. Never
  characterize against Production with real customer data; if Production is
  the only option, stop and ask the tester to get approval and use read-only
  observations.
- Reason for the baseline (refactor, migration, upgrade…) — it decides which
  behaviors matter most.
- Available evidence: screenshots or screen recordings, API requests and
  responses (HAR, Postman/curl output), logs, DB query results, exported
  files/reports, existing test cases, and source code if the tester shares it.
- Test data: accounts/roles and data states that can be reused to reproduce
  each observation (placeholders for secrets).

## Step 2 — Behavior inventory (replaces Requirement analysis)

Produce a **Behavior Inventory** instead of a Requirement Analysis Report:

1. Source inventory — each evidence item with ID (`EV-01`…), type, build,
   environment, date, and who captured it.
2. Scope and exclusions.
3. Observed behaviors, one row each: ID (`BEH-01`…), entry point, input /
   data state, observed output (UI text, response code and key fields, DB
   change, file content, log/event), evidence IDs, and stability
   (`Stable`, `Varies`, `Not reproduced`).
4. Non-deterministic parts that must be masked or ignored when comparing:
   timestamps, generated IDs, ordering without a sort key, random values,
   locale/timezone formats.
5. Behaviors that look wrong or risky, labeled `Suspected defect — confirm`
   (security, data loss, wrong calculation, inconsistent with another screen
   or a known rule). Do not fix or reinterpret them; record them as observed.
6. Gaps: inputs, roles, states, or branches with no evidence yet, and what
   evidence would cover them.
7. Decision questions for BA/PO/dev.

Rules:

- Every observed behavior needs at least one evidence item. Never fill in a
  behavior from guessing, from the UI label, from the code's intent, or from
  how "it usually works". If evidence is missing, list it as a gap.
- Observations come only from evidence the tester provided or from actions
  the tester performed and reported. Do not run requests, scripts, or UI
  actions against the system yourself unless the tester explicitly asks for
  that specific action on a non-production environment, and never run
  actions that change data without that explicit request.
- A behavior seen only once is `Stable` only after it is reproduced (same
  input, same output apart from masked fields); otherwise mark it `Varies` or
  `Not reproduced` and say so.
- When existing documentation disagrees with the observation, record both
  and flag a finding of type `Contradiction`; the baseline still records the
  observation.

## Step 3 — Characterization Test Case Review

Use the normal Test Case Review table with these differences:

- Test Type: `Characterization`.
- Requirement/AC column: the `BEH-xx` ID (and a requirement ID if one exists).
- Expected Result: the observed behavior, precise enough to compare —
  exact messages, status codes, key field values, record counts, file
  columns. Mark masked parts, for example `createdAt = <any ISO timestamp>`.
- Test Data / Preconditions: the exact data state and build/environment the
  observation came from.
- Issue/Assumption: evidence IDs, `Suspected defect — confirm` when it
  applies, and `Varies` notes.
- Prioritize behaviors on the path being changed (refactor/migration scope),
  then high-traffic flows, money/data calculations, permissions, and
  integrations.

## Step 4 — Review and decisions

In addition to the usual three sections, show a **Behavior Decisions** list:
for each `Suspected defect — confirm` and each `Varies` behavior, ask the
tester (or BA/PO) to choose:

- `Keep as baseline` — the current behavior is the expected behavior.
- `Defect — baseline as-is, log bug` — keep the observed behavior in the case
  so the regression run still detects changes, and note the bug ID.
- `Exclude` — not part of the baseline (with reason).

Record each decision in the case's Issue/Assumption (Notes) field.

## Step 5 — Export

Export with the normal scripts. In the workbook:

- Put `Characterization` in the type/notes column available in the template,
  and evidence IDs plus build/environment/date in Notes (or a column approved
  in the Template Change Review).
- Leave Test Results blank or `Untested`. Capturing a baseline is not an
  execution result; the first regression run after the change produces the
  results.
- If the team wants evidence kept with the workbook, propose an `Evidence`
  sheet (ID, type, build, environment, date, location/link) in the Template
  Change Review. Store links or file names, never secrets or real personal
  data.

## Handoff

Add to the usual handoff: baseline build/environment/date, number of
behaviors baselined vs. gaps, suspected defects and their decisions, and the
masked fields the regression run must ignore.
