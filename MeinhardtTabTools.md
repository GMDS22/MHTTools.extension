# Meinhardt Tab Tools Manual

This document describes all tools currently under the MEINHARDT tab, excluding MEP Lab tools.

## Scope
- Included: all pushbutton tools under MEINHARDT tab panels and pulldowns.
- Excluded: all tools under MEP Lab panel.
- Removed from tab: Project Health Monitor panel and its tools.

## Documentation.panel

### FamilyNamingConvention
Opens the family naming convention document.
Basic operation:
1. Click the tool.
2. Read the opened markdown document.

### Search Tool
Searches visible pyRevit ribbon tools and runs the selected one.
Basic operation:
1. Open Search Tool.
2. Type a keyword.
3. Select a result to execute.

### ToolsDescription
Opens this tools documentation file.
Basic operation:
1. Click the tool.
2. Read this manual.

## IMPORT.panel

### Excel
Imports Excel data into a drafting-view table and supports update from saved link info.
Basic operation:
1. Browse to Excel file and select a sheet.
2. Set target view name.
3. Click Import.
4. Use Update later to refresh from saved source.

## MEP Check.panel

### SpaceVsRoom
Checks mismatches between rooms and spaces (number/name comparison) and reports issues.
Basic operation:
1. Select room source document when prompted.
2. Run comparison.
3. Review report output.

## MEP Create.panel

### BatchCreateSystems
Batch creates/duplicates MEP system types from a template workbook.
Basic operation:
1. Open/import the template data.
2. Fill required system entries.
3. Run import in the tool UI.

### BatchDependentViewCreation
Creates dependent views from selected parent views using selected scope boxes.
Basic operation:
1. Select scope boxes.
2. Select parent views.
3. Run creation.

### BatchWorksetCreation
Creates many worksets from typed list or imported text file.
Basic operation:
1. Enter or import workset names.
2. Confirm list.
3. Click create.

### PipeTypeFromCSV
Creates pipe schedule/segment sizes from CSV and applies selected material.
Basic operation:
1. Select CSV file.
2. Enter schedule/type name.
3. Select or create material.

### CreateSection
Creates section views along selected linear elements with offsets and naming.
Basic operation:
1. Select section type/options.
2. Select linear elements.
3. Generate sections.

### QuickDimension
Creates one dimension line to multiple selected references.
Basic operation:
1. Select elements.
2. Pick dimension line location.
3. Tool places dimension.

### Transition
Creates transition fitting between two selected open connectors.
Basic operation:
1. Pick first connector-side element.
2. Pick second connector-side element.
3. Tool creates transition.

## MEP Data.panel

### ElementChangeLevel
Changes level/reference level for selected MEP elements/spaces while preserving elevation logic.
Basic operation:
1. Select elements.
2. Choose target level.
3. Apply.

### ElevationUnder
Computes vertical distance to nearest floor/roof above and writes result to parameter.
Basic operation:
1. Select target elements.
2. Select settings/view and output parameter.
3. Run calculation.

### ExportMaterialsGraphics
Exports material graphics data to JSON for material database workflows.
Basic operation:
1. Run tool.
2. JSON is generated in expected update location.

### UpdateMaterials
Updates/regenerates material database files from source cache/XML and applies graphics/assets.
Basic operation:
1. Run tool.
2. Wait for rebuild/update.
3. Review output folder/logs.

### RoomToSpace
Copies mapped room parameter values into spaces from selected source document.
Basic operation:
1. Select source and target docs.
2. Define room-to-space parameter mapping.
3. Run update.

## MEP Design.panel

### About
Opens author/project GitHub page.
Basic operation:
1. Click tool.
2. Browser opens.

### RemoveDuplicates
Finds and removes duplicate elements based on selected categories/options.
Basic operation:
1. Select categories/options.
2. Confirm.
3. Tool deletes duplicates.

### SplitPipes
Splits all eligible pipes by configured distance/type logic.
Basic operation:
1. Configure split settings.
2. Run split.

### SplitSelectedPipes
Splits only selected pipes by configured distance.
Basic operation:
1. Select pipes.
2. Configure split distance/options.
3. Run.

## MEP Export.panel

### MultiIFC
Batch exports selected 3D views to IFC using chosen IFC configuration.
Basic operation:
1. Select IFC configuration.
2. Select 3D views.
3. Choose output folder and export.

## MEP Favorites.panel

### ConnectTo
Moves selected element and connects nearest compatible connectors.
Basic operation:
1. Pick element to move.
2. Pick target connector element.
3. Tool aligns/connects.

### DisConnect
Disconnects connectors on selected elements.
Basic operation:
1. Select elements.
2. Run tool.

### MakeParallel
Rotates target elements to match reference element XY direction.
Basic operation:
1. Pick reference element.
2. Pick target element(s).
3. Apply rotation.

### RemoveDuplicates
Same duplicate cleanup workflow as Design panel.
Basic operation:
1. Select categories/options.
2. Run cleanup.

### RoomToSpace
Same room-to-space parameter mapping workflow as Data panel.
Basic operation:
1. Choose source/target docs.
2. Define mappings.
3. Run.

### SplitPipes
Same batch split workflow as Design panel.
Basic operation:
1. Configure split settings.
2. Run.

### SplitSelectedPipes
Same selected-pipe split workflow as Design panel.
Basic operation:
1. Select pipes.
2. Set split distance/options.
3. Run.

### Transition
Same two-connector transition creation workflow as Create panel.
Basic operation:
1. Pick first connector side.
2. Pick second connector side.
3. Repeat as needed.

## MEP Manage.panel

### Callout Riser Renamer
Renames selected callout views using riser-related parameter values.
Basic operation:
1. Select callout views.
2. Run tool.
3. Review renamed count.

### CopyPipeType
Copies selected pipe type from source open document into active document.
Basic operation:
1. Select source document.
2. Select pipe type.
3. Copy.

### CopyProjectUnits
Copies project units from source document to target document.
Basic operation:
1. Select source doc.
2. Select target doc.
3. Apply.

### CopyViewRange
Copies view range settings from source plan view(s) to target plan view(s).
Basic operation:
1. Select source view(s).
2. Select target view(s).
3. Apply.

### CopyViewType
Copies view family types from source document and updates template defaults in target doc.
Basic operation:
1. Select source and target docs.
2. Run copy.

### FamilyReLoad
Reloads selected families into selected target document with overwrite options.
Basic operation:
1. Select family instances.
2. Select target doc and overwrite behavior.
3. Run reload.

### FluidCreate
Creates/updates fluid types and temperature-property tables.
Basic operation:
1. Choose fluid and property settings.
2. Set temperature range/steps.
3. Create/update.

### ReplaceFluid
Replaces fluid type/temperature settings in piping systems.
Basic operation:
1. Select source fluid conditions.
2. Select target fluid conditions.
3. Apply replacement.

### ManageDocumentParameters
Opens document parameter manager for family/project parameter operations.
Basic operation:
1. Run tool.
2. Use dialog to add/edit/manage parameters.

### ManageSharedParameter
Opens shared parameter manager for definitions/groups.
Basic operation:
1. Run tool.
2. Manage shared parameter groups/definitions.

## MEP Modify.panel

### GreyOutElements
Applies gray graphic overrides to selected/target elements in active view.
Basic operation:
1. Select elements (or follow tool targeting behavior).
2. Run tool.

### GreyOutElements_reset
Resets gray graphic overrides applied by GreyOutElements.
Basic operation:
1. Run reset in target view.

### ConnectTo
Connector move-and-connect workflow (same core behavior as Favorites).
Basic operation:
1. Pick moving element.
2. Pick target element.
3. Apply connection.

### DisConnect
Disconnects connectors on selected elements.
Basic operation:
1. Select elements.
2. Run tool.

### Element3DRotation
Rotates selected element(s) around chosen axis using modeless UI/external event.
Basic operation:
1. Select element and axis reference.
2. Input rotation values.
3. Apply.

### MakeParallel
Aligns target elements parallel to reference element in XY.
Basic operation:
1. Pick reference.
2. Pick targets.
3. Apply.

### FamilyDelete
Deletes selected families from project.
Basic operation:
1. Select families.
2. Confirm delete.

### FamilyTypeDelete
Deletes selected family types.
Basic operation:
1. Select family types.
2. Confirm delete.

### ParameterDelete
Deletes selected project parameters.
Basic operation:
1. Select parameters.
2. Confirm delete.

### SystemDelete
Deletes selected system objects.
Basic operation:
1. Select systems/elements.
2. Run delete.

### FlexFlatten
Flattens selected flex elements.
Basic operation:
1. Select flex objects.
2. Run tool.

### MoveLabelToOrigin
Moves selected labels/annotations to origin.
Basic operation:
1. Select labels.
2. Run tool.

### MoveSpaceToRoom
Moves spaces to corresponding room positions by room number.
Basic operation:
1. Select room-source document.
2. Run tool.

### MoveTitleBlockToOrigin
Moves title block annotation to origin.
Basic operation:
1. Select title block.
2. Run tool.

## MEP Samples.panel

### DataContextSample
Sample tool demonstrating WPF DataContext usage in pyRevit.
Basic operation:
1. Run sample.
2. Interact with demo form.

### FormExternalEventHandler
Sample modeless form + external event workflow.
Basic operation:
1. Open sample form.
2. Trigger sample external event action.

### SimpleExternalEventHandler
Minimal external event sample tool.
Basic operation:
1. Run sample.
2. Trigger event.

## MEP Tests.panel

### unittest_parameter
Runs unit tests for parameter-related behavior.
Basic operation:
1. Run tool.
2. Review test output.

## Project.panel

### 03 Unified Create (Rooms-Areas-Spaces-Zones)
Creates coordinated rooms/spaces/areas/zones and related boundaries/tags from host/link context.
Basic operation:
1. Open valid plan view.
2. Configure scope/options.
3. Run creation.

### 04 Color by Value (Export-Import)
Builds/updates parameter-based filters with color assignment and CSV import/export.
Basic operation:
1. Select category/views/parameter.
2. Assign/import colors.
3. Apply filters.

### 05 Color Scheme Editor
Edits color fill scheme entries and supports backup/import/export.
Basic operation:
1. Select scheme and category.
2. Modify/import/export entries.
3. Apply changes.

### 06 Legend Creator
Creates legends (family text, color scheme, spaces) in legend/drafting views.
Basic operation:
1. Select legend mode/options.
2. Set placement/view settings.
3. Generate.

### 07 Family Renamer
Applies classification and naming rules to rename families/types.
Basic operation:
1. Load/select naming rules.
2. Preview changes.
3. Execute rename.

### 08 Selection Filter
Filters current selection by parameter values with keep/remove mode.
Basic operation:
1. Choose scope and parameter.
2. Select values.
3. Apply filter.

### BOUNDS / HOST
Host-view boundary workflow for room/space/area generation and cleanup.
Basic operation:
1. Open valid plan view.
2. Configure scheme/options.
3. Run.

### BOUNDS / LINK
Boundary workflow using linked model context.
Basic operation:
1. Select link/options.
2. Run boundary-based creation.

### BOUNDS / S2A
Converts spaces to areas with area boundaries/tags.
Basic operation:
1. Open target area plan.
2. Choose area scheme/plan.
3. Run conversion.

### GM Filters by Value
Creates/updates value-based filters (including linked contexts) and applies graphics overrides.
Basic operation:
1. Select target views and settings.
2. Configure filter/color behavior.
3. Apply.

### LINKS
Applies linked-view visibility settings to selected Revit links in active host view.
Basic operation:
1. Select link instances.
2. Select linked view/options.
3. Apply.

### LINKVIS
Batch applies link visibility actions across selected views/templates and link types.
Basic operation:
1. Select views/templates.
2. Select link types and action.
3. Apply.

## Search.panel

### Search Tool
Searchable launcher for loaded pyRevit commands.
Basic operation:
1. Open tool.
2. Type filter text.
3. Double-click or press Enter to run selected command.

## Refresh.panel

### Refresh
Reloads pyRevit tools/session.
Basic operation:
1. Click Refresh.
2. Wait for reload to finish.

## Notes
- Documentation-only panel: Documentation.
- Utility-only panels: Search and Refresh.
- Development/support panels: MEP Samples and MEP Tests.