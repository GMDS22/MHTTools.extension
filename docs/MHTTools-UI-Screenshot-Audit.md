# MHTTools UI Screenshot Audit

Updated: 2026-07-02

This audit tracks the UI screenshot coverage required by the tool-description standard.
It compares the current MeinhardtTabTools.html entries against confirmed XAML-backed tool interfaces and the screenshot assets currently available in MEINHARDT.tab/Documentation.panel/Docs.stack/ToolsDescription.pushbutton/.

## Completed In This Pass

- Fixed the broken Family Renamer screenshot reference in MeinhardtTabTools.html.
- Added the existing Views Links Manager screenshot to the Views Links Manager card.

## Confirmed Documented Tools Still Needing Actual UI Screenshots

- MEP Data Panel: NWB Data Validator
  - UI evidence: MEINHARDT.tab/MEP Data.panel/Data Manage.pulldown/NWB Data Validator.pushbutton/ValidatorOptionsWindow.xaml
  - Gap: the read-only validation setup window was added without a captured Revit screenshot. Capture `ui-nwb-data-validator.png` after the next in-Revit validation run, then add the image and factual caption to its documentation card.

- Project Panel: 06 Legend Creator
  - UI evidence: MEINHARDT.tab/Project.panel/06 Legend Creator.pushbutton/LegendCreatorUI.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- Views Links Panel: Plan Link Setup
  - UI evidence: MEINHARDT.tab/Views Links.panel/Plan Link Setup.pushbutton/PlanLinkSetup.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- MHT SHEETS Panel: RENUMBER_RENAME
  - UI evidence: MEINHARDT.tab/MHT SHEETS.panel/RENUMBER_RENAME.pushbutton/Script.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- MHT SHEETS Panel: AlignViewports
  - UI evidence: MEINHARDT.tab/MHT SHEETS.panel/AlignViewports.pushbutton/Script.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- IMPORT Panel: Excel
  - UI evidence: MEINHARDT.tab/IMPORT.panel/Excel.pushbutton/WPFWindow.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- Mechanical Services Panel: SMART CHECKER
  - UI evidence: MEINHARDT.tab/Mechanical Services.panel/SMART CHECKER.pushbutton/SmartChecker.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- MEP Create Panel: BatchCreateSystems
  - UI evidence: MEINHARDT.tab/MEP Create.panel/create1.stack/BatchCreation.pulldown/BatchCreateSystems.pushbutton/WPFWindow.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- MEP Create Panel: BatchWorksetCreation
  - UI evidence: MEINHARDT.tab/MEP Create.panel/create1.stack/BatchCreation.pulldown/BatchWorksetCreation.pushbutton/WPFWindow.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- MEP Data Panel: ElementChangeLevel
  - UI evidence: MEINHARDT.tab/MEP Data.panel/ElementChangeLevel.pushbutton/ReferenceLevelSelection.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- MEP Data Panel: ElevationUnder
  - UI evidence: MEINHARDT.tab/MEP Data.panel/ElevationUnder.pushbutton/WPFWindow.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- MEP Data Panel: RoomToSpace
  - UI evidence: MEINHARDT.tab/MEP Data.panel/RoomToSpace.pushbutton/WPFWindow.xaml
  - Gap: tool card is documented, but no UI screenshot asset is currently referenced.
- MEP Data Panel: Linked Room Parameter
  - UI evidence: MEINHARDT.tab/MEP Data.panel/Linked Room Parameter.pushbutton/WPFWindow.xaml
  - Gap: tool card is now documented, but no UI screenshot asset is currently referenced.
- MEP Data Panel: Linked Elements Parameter
  - UI evidence: MEINHARDT.tab/MEP Data.panel/Linked Elements Parameter.pushbutton/WPFWindow.xaml
  - Gap: tool card is now documented, but no UI screenshot asset is currently referenced.
- MEP Data Panel: Parameter Transfer
  - UI evidence: MEINHARDT.tab/MEP Data.panel/ParameterTransfer.pushbutton/ParameterTransfer.xaml
  - Gap: tool card is now documented, but no UI screenshot asset is currently referenced.
- Audit and Exchange Panel: Copy From Link
  - UI evidence: pyRevit command dialogs in MEINHARDT.tab/Audit and Exchange.panel/Audit.stack/Copy From Link.pushbutton/script.py
  - Gap: tool card is now documented, but no UI screenshot asset is currently referenced.

## Recommended Capture Order

1. Plan Link Setup
2. SMART CHECKER
3. RENUMBER_RENAME
4. AlignViewports
5. Excel
6. Legend Creator
7. BatchCreateSystems
8. BatchWorksetCreation
9. ElementChangeLevel
10. ElevationUnder
11. RoomToSpace
12. Linked Room Parameter
13. Linked Elements Parameter
14. Parameter Transfer
15. Copy From Link

## Capture Notes

Use scripts/capture-tool-ui.ps1 and save new assets into MEINHARDT.tab/Documentation.panel/Docs.stack/ToolsDescription.pushbutton/ with the naming rules from docs/MHTTools-Tool-Description-Standard.md.
