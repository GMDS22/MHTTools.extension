---
applyTo: "MEINHARDT.tab/Documentation.panel/Docs.stack/ToolsDescription.pushbutton/**"
description: "Use when adding, renaming, or replacing screenshot assets for Meinhardt tool descriptions so naming, placement, and screenshot type remain standardized."
---
Store tool-description screenshots in this folder and follow the standard naming rules from docs/MHTTools-Tool-Description-Standard.md.

Asset rules:
- Use PNG for UI screenshots.
- Prefer these filename patterns:
  - ui-<tool>.png
  - ui-<tool>-LargeUI.png
  - ui-<tool>-step1.png, ui-<tool>-step2.png, ui-<tool>-step3.png
  - panel-<tool>.png
- Use lowercase kebab-case slugs for new files.
- Do not introduce new dot-separated LargeUI variants; keep new files on the hyphenated pattern.
- Replace placeholders with real screenshots whenever a tool has a user-facing interface.
- Keep filenames stable once referenced in MeinhardtTabTools.html unless the HTML is updated in the same change.
