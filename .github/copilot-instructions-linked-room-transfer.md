# Linked Room Parameter Transfer Hardening Protocol

Use this protocol when editing `MEINHARDT.tab/MEP Data.panel/Data Manage.pulldown/Linked Room Parameter.pushbutton/`.

## Core Safety Rules

- Never claim a successful write unless the parameter value is verified after transaction commit.
- Distinguish outcomes clearly:
  - `SUCCESS`: value changed and persisted.
  - `SKIP`: no write needed, empty value skipped, duplicate type write skipped, or policy skip.
  - `FAIL`: attempted write did not persist or raised an unrecoverable write error.
- Keep transactions short; never hold a transaction open during user selection prompts.
- Wrap UI handlers and transfer operations in defensive `try/except` and present actionable user alerts.
- Preserve Safe Mode guardrails and do not silently disable them.

## Transfer Truthfulness Rules

- Summary numbers must be internally consistent and non-overlapping:
  - `Mappings evaluated`
  - `Parameter writes (verified)`
  - `Write calls returned success (pre-verify)`
  - `Failed writes`
  - `Skipped`
- If `verified writes == 0`, show an explicit outcome line that explains why (for example already up-to-date).
- If writes are blocked by policy (for example type writes in auto-room mode), count and report them separately.

## Value-Matching Rules

- Treat unchanged values as `SKIP`, not `FAIL` and not `SUCCESS`.
- Compare values by storage type:
  - String: trimmed exact text.
  - Integer: integer equality.
  - Double: numeric tolerance.
  - ElementId: integer id equality.
- Preserve `AsValueString()` fallback behavior for linked-room parameter reads.

## Logging Rules

- Always write a human-readable transfer log and a detailed CSV.
- CSV must include at least:
  - `element_id`
  - `room_param`
  - `target_param`
  - `target_source` (instance/type)
  - `storage`
  - `before_value`
  - `attempted_value`
  - `after_value`
  - `result`
  - `message`
- For debugging, include runtime breadcrumbs in summary output.

## Runtime UX Rules

- Before transfer, provide estimated runtime in milliseconds.
- If estimated runtime is large, prompt for confirmation before continuing.
- Keep UI state responsive under maximize/resize:
  - Adjust major list/panel heights dynamically.
  - Keep content reachable with scrolling.

## Auto-Match Rules

- Support keyword exclusions from auto-match (for example `uniclass,class`).
- Provide a target-parameter checkbox filter panel in Step 3.
- Auto-match must honor both:
  - keyword exclusions
  - checkbox target filters
- Report exclusion counts in auto-match feedback.

## Source Detection Rules

- Respect selected source mode and selected link scope consistently.
- In per-level mode, use selected level from UI (not only active view level fallback).
- Keep host/linked source behavior explicit and deterministic.

## Validation Checklist Before Finalizing Changes

- `get_errors` returns no errors for modified files.
- Transfer summary language matches actual logic and counters.
- No UI control/event name mismatches between XAML and `script.py`.
- No new crash-prone paths introduced in selection, mapping, or transfer handlers.
