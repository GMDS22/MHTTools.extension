# Meinhardt UI Theme System

This document is the single source of truth for UI theme tokens and usage rules across tools under MEINHARDT.tab.

## Goals

- Keep all tool interfaces visually consistent.
- Improve visual quality and readability.
- Allow fast global theme updates by changing tokens in one place.

## Core 3-Role Palette

### 1. Primary (Blue)

- Purpose: identity, primary actions, key emphasis.
- Base token: `--theme-primary: #0E639C`
- Hover token: `--theme-primary-hover: #1177BB`

### 2. Secondary (Teal)

- Purpose: supportive highlights, secondary action states, contextual chips.
- Base token: `--theme-secondary: #1FA6A6`
- Hover token: `--theme-secondary-hover: #22B8B8`

### 3. Accent (Amber)

- Purpose: attention and semantic highlight (warnings, notable states, selected key indicators).
- Base token: `--theme-accent: #F2B134`
- Strong token: `--theme-accent-strong: #D99A20`

## Neutral Support Tokens

- `--surface-0: #1A1A1C`
- `--surface-1: #232327`
- `--surface-2: #2A2A2E`
- `--border: #3A3A40`
- `--text-main: #E8E8E8`
- `--text-muted: #8A8A9A`
- `--success: #4CAF7D`
- `--danger: #CC4B4B`

## Usage Rules

1. Use Primary Blue for primary buttons, step badges, and major anchors.
2. Use Secondary Teal for supporting emphasis, alternative actions, and non-critical highlights.
3. Use Accent Amber for alerts, callouts, and high-attention visual markers.
4. Keep core identity to the 3-role palette; do not introduce additional identity colors unless explicitly approved.
5. Preserve contrast and readability on dark surfaces.
6. Theme-aware controls must switch foreground and background tokens together; never allow light text on light fields or dark text on dark fields.
7. In WPF dialogs, controls that commonly ignore simple brush setters, especially ComboBox, must use explicit templates for closed state, popup state, and selected-item state when the dialog is themed.

## Component Guidance

- Buttons:
  - Primary action: blue background, white text.
  - Secondary action: neutral surface with teal border or subtle teal emphasis.
  - Critical warning action: amber accent treatment.
- Combo boxes:
  - Closed-state background must use a dark surface token such as `surface-1` or `surface-2`, never a light or white fill inside dark-theme tools.
  - Selected text and dropdown-item text must use `text-main` or another clearly contrasting foreground.
  - Dropdown popup backgrounds and item rows must use dark surface tokens distinct from the foreground text so the selected value remains readable.
  - If a tool intentionally uses a light surface, swap the text and highlight colors to dark-on-light equivalents in the same template so OS or default-control fallback cannot produce white-on-white or light-on-light states.
  - Do not rely on default OS ComboBox rendering in themed dialogs; template the closed control, popup container, and item states explicitly.
- Cards/Sections:
  - Use consistent surface layering (`surface-1` and `surface-2`) and border token.
- Typography:
  - Keep body text in `text-main`; supporting hints in `text-muted`.
- Lists/Tables:
  - Keep alternating row contrast subtle and consistent.

## Migration Checklist

1. Update tokens here first.
2. Propagate token usage to XAML resource dictionaries/styles.
3. Remove one-off hardcoded colors where practical.
4. Verify labels/tooltips remain clear and consistent.
5. Run visual QA pass for consistency and contrast, including ComboBox closed state, popup state, and selected-item readability.

## Change Log

- 2026-05-20: Initial single-source theme system created (Blue primary, Teal secondary, Amber accent).
- 2026-05-20: First rollout pass applied to MEP Data.panel XAML dialogs for consistent dark surfaces and blue/teal action styling.
- 2026-08-01: Added theme protocol requiring explicit ComboBox templating and matched foreground/background adaptation to prevent white-field fallback and unreadable text in MHT tools.
