# MHTTools Tool Description Standard

## Purpose

This document standardizes how Meinhardt tab tool descriptions are written, laid out, themed, and illustrated.
It is based on the current production format in MeinhardtTabTools.html and should be treated as the documentation contract for future updates.

## Current Format Analysis

The current description page already has a strong structure. Its main patterns should be preserved and reused:

- Page shell:
  - `brand-head`
  - `hero`
  - `toc`
  - `section.panel`
- Tool card shell:
  - `article.tool`
  - `tool-head`
  - short summary paragraph
  - flat usage bullets
- Screenshot wrappers:
  - `ui-placeholder`
  - `shot-ui`
  - `shot-panel-ribbon`
  - `panel-shot`
- Specialized content layouts:
  - `tool-split` for copy + screenshot side-by-side
  - `steps-layout` for multi-step UI workflows
  - `panel-shot-full` for wide ribbon/panel imagery

This means the standard should evolve from the current page, not replace it with a new documentation style.

## Required Tool Entry Structure

Every documented tool entry should include these minimum parts:

1. Tool name in `tool-head`
2. One concise summary paragraph
3. A flat list of action-oriented usage bullets
4. A screenshot when the tool has a user-facing UI
5. A caption that explains what the screenshot demonstrates

## Screenshot Requirement

If a tool has a visible UI, the document must include an actual UI screenshot.

Only omit screenshots when one of these is true:

- the tool has no UI and runs directly from the ribbon
- the workflow is purely background or API-driven with no meaningful interface
- a screenshot is temporarily impossible, and the gap is explicitly tracked for follow-up

Placeholders are acceptable only as temporary gaps, not as the final state for UI tools.

## Screenshot Type Standard

Use the existing HTML sizing logic as the standard.

### 1. Compact Dialog or Form

Use for typical WPF or modal tool windows.

- HTML wrapper: `ui-placeholder shot-ui`
- Filename pattern: `ui-<tool>.png`
- Display behavior: constrained height, centered image

### 2. Large or Dense Interface

Use for wide, complex, or multi-pane windows.

- HTML wrapper: `ui-placeholder shot-ui`
- Filename pattern: `ui-<tool>-LargeUI.png`
- Display behavior: the `LargeUI` token triggers full-width rendering in the current CSS

Note:

- Existing legacy files like `ui-family-renamer.LargeUI.png` can remain for backward compatibility.
- New files should use the hyphenated form: `-LargeUI.png`.

### 3. Multi-Step Workflow

Use when one screenshot is not enough to explain the tool.

- Layout: `steps-layout`
- Filename pattern:
  - `ui-<tool>-step1.png`
  - `ui-<tool>-step2.png`
  - `ui-<tool>-step3.png`
- Caption pattern: explain the purpose of each step, not just the visual.

### 4. Ribbon or Panel Context

Use when the panel location or ribbon layout matters.

- Wrapper: `panel-shot` or `shot-panel-ribbon`
- Filename pattern: `panel-<tool>.png`
- Use for contextual images, not as a substitute for the real tool UI when a UI exists.

## Caption Standard

Captions must be factual and useful.

Good caption qualities:

- describes what the user is seeing
- points out the stage or action the screenshot supports
- distinguishes between step screenshots when there are several

Avoid:

- repeating the tool name without added value
- vague captions like "Tool screenshot"
- decorative captions that do not help the reader

## Layout Selection Rules

Choose the card layout based on the tool type.

- Simple tool with one interface:
  - standard `article.tool` with one `shot-ui`
- Tool with balanced copy and media:
  - `tool-split`
- Multi-step wizard or staged workflow:
  - `steps-layout`
- Panel overview or ribbon location:
  - `panel-shot` or `shot-panel-ribbon`

## Theme Coordination

The description system should stay coordinated with the Meinhardt visual theme.

Current page direction:

- blue-led base theme
- non-monochrome accent support
- warm, gold, rose, and ice accents used sparingly for hierarchy and energy
- dark panel depth with readable contrast

Theme rules for tool descriptions:

- Do not flatten the documentation into one blue tone.
- Preserve accent variety when improving the page.
- If a change affects shared tokens, palette roles, or global documentation CSS, coordinate with `.github/agents/meinhardt-ui-theme-unifier.agent.md`.
- Documentation styling should remain visually related to the tool UIs and the Meinhardt tab brand.

## Screenshot Asset Location

Store tool-description screenshots here:

- `MEINHARDT.tab/Documentation.panel/Docs.stack/ToolsDescription.pushbutton/`

## Screenshot Naming Rules

Use lowercase kebab-case slugs for new assets.

Preferred patterns:

- `ui-<tool>.png`
- `ui-<tool>-LargeUI.png`
- `ui-<tool>-step1.png`
- `ui-<tool>-step2.png`
- `ui-<tool>-step3.png`
- `panel-<tool>.png`

## Capture Workflow

Use the helper script:

- `scripts/capture-tool-ui.ps1`

Recommended workflow:

1. Open the target tool UI in Revit.
2. Capture the active window or a named window title using the script.
3. Save the screenshot directly into the ToolsDescription asset folder.
4. Reference the new image in `MeinhardtTabTools.html` using the correct wrapper and caption.

Example commands:

```powershell
./scripts/capture-tool-ui.ps1 -ToolSlug format-v3 -Variant LargeUI -WindowTitleContains "FORMAT"
./scripts/capture-tool-ui.ps1 -ToolSlug unified-create -Variant Step1 -WindowTitleContains "Unified Create"
./scripts/capture-tool-ui.ps1 -ToolSlug search-tool -Variant LargeUI -WindowTitleContains "Search Tool"
```

## Agent and Instruction Usage

Use these workspace customizations for future work:

- Custom agent:
  - `.github/agents/mhttools-tool-description-curator.agent.md`
- File instructions:
  - `.github/instructions/meinhardt-tool-description-html.instructions.md`
  - `.github/instructions/meinhardt-tool-description-assets.instructions.md`

## Quality Checklist

Before finishing a tool-description update, confirm:

- the tool card matches an existing approved layout pattern
- UI tools have screenshots
- screenshot filenames match the naming convention
- captions are specific and useful
- the page still fits the Meinhardt theme direction
- any new pattern is documented here before reuse
