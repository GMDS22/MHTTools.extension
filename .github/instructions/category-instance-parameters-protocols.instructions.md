---
description: "Safety, sizing, timing, write-verification, and value-translation protocols for Category Instance Parameters."
applyTo: "MEINHARDT.tab/MEP Data.panel/Data Manage.pulldown/Category Instance Parameters.pushbutton/**"
---
# Category Instance Parameters Protocols

Use these rules whenever editing, reviewing, or extending the Category Instance Parameters tool.

## Crash Safety
- Treat window launch, collection, preview, and write actions as crash-prone until verified.
- Keep XAML control names and Python handlers in sync at all times.
- Never remove defensive `try/except` wrappers around window startup, long-running actions, or UI callbacks.
- Keep a busy/re-entry guard so only one expensive operation runs at a time.
- Skip invalid, missing, or stale Revit elements instead of dereferencing them.
- Keep transactions short, deterministic, and rollback-safe.
- Do not hold a transaction open while prompting the user.

## Window And Layout
- The window must fit the screen work area and may launch maximized.
- UI panels must auto-adjust to the available height.
- Category, parameter, and preview lists must scroll internally.
- Avoid layouts that force the full dialog past the monitor height.
- Keep dropdowns and list controls bounded so they do not expand indefinitely.

## Long-Run Warnings
- Before any large collection or write operation, show a warning that includes an estimated duration.
- Provide the estimate in both seconds and milliseconds when possible.
- Require explicit confirmation when the estimate or element count suggests a long run.
- If the action can touch many elements, explain that it may take time and may affect model responsiveness.

## Write Verification
- Never count a write as successful until the parameter is read back and compared with the expected final value.
- Success means the actual stored value matches the planned value, not merely that `Set()` returned true.
- For string values, verify the final user-facing text, not only the raw storage form.
- For numeric values, verify the read-back storage value with a tolerance for doubles.
- For ElementId values, verify the resolved id or readable reference value.
- Report verification failures explicitly and do not include them in success counts.

## Value Display And Preview
- Show user-facing values in preview and write summaries whenever possible.
- Distinguish `<missing>`, `<empty>`, and valid values.
- If a parameter has a readable display string, prefer it over raw numeric storage for UI output.
- When showing sample writes, display the actual post-write value, not the attempted value.

## Default Translation And Manipulation
- Preserve the built-in dictionary translation mode as the default path for code-to-label mapping.
- If no custom dictionary is provided, fall back to the default translation table.
- Keep dictionary translation case-insensitive.
- Support multiline `key=value` or `key=>value` entries in the expression field.
- Keep template/token transforms and safe formula transforms opt-in and explicitly labeled.
- Do not allow arbitrary Python execution for manipulation logic.
- Any formula mode must be sandboxed to a narrow allowlist of simple operators and functions.

## User Feedback
- Show a clear summary after each run with counts for written, skipped, and failed items.
- Include sample written values and sample failures when available.
- If a long run was estimated, include the estimate in the confirmation prompt.
- If a verification mismatch occurs, report the mismatch directly and do not claim success.

## Validation Expectations
- Validate changed Python files for syntax after edits.
- Validate XAML bindings and control names after UI changes.
- Prefer targeted checks over broad repo-wide operations.
- If Revit is frozen or unstable, pause and reduce operation scope before retrying.
