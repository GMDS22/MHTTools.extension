---
description: "Use when capturing, updating, or auditing tool screenshots so UI documentation assets stay standardized and complete."
applyTo: "MEINHARDT.tab/Documentation.panel/Docs.stack/ToolsDescription.pushbutton/**,scripts/capture-tool-ui.ps1,docs/MHTTools-UI-Screenshot-Audit.md,MeinhardtTabTools.html"
---
# MHTTools Screenshot Orchestration Rules

Always load these references:
- docs/MHTTools-Tool-Description-Standard.md
- docs/MHTTools-UI-Screenshot-Audit.md
- .github/instructions/meinhardt-tool-description-assets.instructions.md
- scripts/capture-tool-ui.ps1

## Capture Rules
- Prefer real UI screenshots for user-facing tools.
- Use standard filename patterns:
  - ui-<tool>.png
  - ui-<tool>-LargeUI.png
  - ui-<tool>-step1/2/3.png
  - panel-<tool>.png
- Keep naming lowercase kebab-case.

## Placement Rules
- Store assets only in Documentation.panel/Docs.stack/ToolsDescription.pushbutton unless explicitly required elsewhere.
- Update MeinhardtTabTools.html references in the same change when asset names change.
- Keep screenshot type aligned to UI complexity (compact, LargeUI, or multi-step).

## Validation
- Confirm each new image is referenced by docs where needed.
- Confirm replaced images do not leave stale references.
- Confirm the audit document reflects newly covered tools.
