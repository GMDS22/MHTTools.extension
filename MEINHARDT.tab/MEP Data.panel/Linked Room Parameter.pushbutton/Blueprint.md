# Linked Room Parameter Transfer - Blueprint

## 1. Tool Identity
- Tool name: Linked Room Parameter Transfer
- Platform: pyRevit (Python + Revit API + WPF/XAML)
- Ribbon location: MEINHARDT.tab > MEP Data.panel > Linked Room Parameter Transfer
- Goal: transfer selected linked-room parameter values into selected host model elements for practical MEP coordination workflows.

## 2. Purpose
This tool lets users manually pick a Room from a loaded linked architectural model, inspect room parameters, map them to writable host-element parameters, and push values to selected MEP/General elements with transaction safety and clear reporting.

Design priorities:
- usability
- reliability
- flexibility
- professional UI/UX
- real-world BIM workflow practicality

## 3. Scope
### 3.1 In Scope
- Linked model detection.
- Manual linked room picking.
- Room parameter extraction (name, value, storage type, has-value state).
- Category-driven element targeting.
- Common writable parameter discovery (default behavior).
- Optional non-common parameter mode.
- Smart parameter matching suggestions.
- Transfer with duplicate handling options.
- Summary report of transfer results.

### 3.2 Out of Scope (Current Release)
- Automatic room containment / nearest-room logic.
- Batch processing across multiple rooms.
- Schedule/Excel-driven mapping import-export (future).

## 4. Workflow
1. User launches tool.
2. Tool detects all loaded Revit links.
3. User chooses linked model from dropdown.
4. User clicks Select Linked Room and picks a Room from selected link in view.
5. Tool extracts room parameters (empty values hidden by default).
6. User chooses target categories.
7. User chooses target elements by:
   - Use Current Selection, or
   - Select Elements From Model (category-filtered pick).
8. Tool scans target elements and discovers writable parameters.
9. Tool suggests mappings via smart matching.
10. User executes Transfer Parameters.
11. Tool writes values inside a Revit transaction.
12. Tool shows transfer summary (updated, failed, skipped).

## 5. UI Specification (WPF)
Window title: Linked Room Parameter Transfer

### 5.1 Top Section
- Linked model dropdown.
- Select Linked Room button.
- Selected Room Info panel:
  - Room Name
  - Room Number
  - Level
  - Link Name

### 5.2 Category Section
- Target Categories panel.
- Search box for category names.
- Multi-select checklist.
- Select All / Deselect All buttons.

### 5.3 Target Element Section
- Use Current Selection button.
- Select Elements From Model button.
- Element count summary by category.
- Toggle: Show Non-Common Parameters.

### 5.4 Mapping Section
- Room parameters list (with optional Show Empty Parameters).
- Writable target parameter list.
- Smart Mapping Preview list.
- Toggle: Skip Empty Values.
- Duplicate handling selector:
  - Overwrite
  - Skip
  - Append

### 5.5 Bottom Section
- Transfer Parameters
- Refresh
- Cancel

## 6. Category Support Matrix
### 6.1 HVAC / Plumbing
- OST_DuctCurves
- OST_DuctFitting
- OST_DuctAccessory
- OST_FlexDuctCurves
- OST_DuctTerminal
- OST_MechanicalEquipment
- OST_PipeCurves
- OST_PipeFitting
- OST_PipeAccessory
- OST_PlumbingFixtures
- OST_Sprinklers
- OST_FlexPipeCurves
- OST_DuctInsulations / OST_PipeInsulations

### 6.2 Electrical
- OST_LightingFixtures
- OST_LightingDevices
- OST_ElectricalEquipment
- OST_ElectricalFixtures
- OST_DataDevices
- OST_CommunicationDevices
- OST_FireAlarmDevices
- OST_NurseCallDevices
- OST_SecurityDevices
- OST_TelephoneDevices
- OST_CableTray
- OST_Conduit

### 6.3 General
- OST_GenericModel
- OST_SpecialityEquipment

## 7. Data and Matching Rules
### 7.1 Linked Room Parameter Extraction
For each room parameter capture:
- parameter name
- parameter value
- storage type
- has-value state

Default behavior:
- ignore empty/null values

Optional behavior:
- Show Empty Parameters toggle

### 7.2 Target Parameter Eligibility
Only include parameters that are:
- writable
- non-read-only
- storage compatible or safely convertible

### 7.3 Smart Matching
Matching strategy order:
1. exact match (case-insensitive)
2. normalized match (ignore underscores/spaces/symbols)
3. fuzzy containment match

Examples:
- Room Name -> ROOM_NAME
- Room Number -> ROOM_NO
- Department -> Department

User can override mapping.

## 8. Transfer Logic
### 8.1 Transaction
- Execute inside Revit Transaction.
- On fatal error: rollback transaction.

### 8.2 Value Write Rules
- String target: convert to string.
- Integer target: cast safely to int.
- Double target: cast safely to float/double.
- ElementId target: skip unless explicit resolver exists.

### 8.3 Duplicate Handling
When target parameter already has value:
- Overwrite
- Skip
- Append (string parameters)

### 8.4 Multi-Category Safety
If a mapped target parameter does not exist on an element:
- skip safely
- log as skipped

## 9. Error Handling Requirements
Handle and report clearly:
- no linked models loaded
- no room selected
- invalid linked selection
- unloaded link selected
- no categories selected
- no target elements selected
- read-only parameter attempts
- storage mismatch conversion failures
- missing target parameters on some elements

## 10. Reporting Requirements
After transfer show summary with:
- room used
- elements updated
- total parameter writes
- failed writes
- skipped writes
- sample failure messages

Optional future enhancement:
- Export log text file

## 11. Code Architecture
### 11.1 Modules (Target Design)
- ui/: WPF views + interaction events
- services/link_room_service.py: link/room selection + extraction
- services/element_service.py: category filtering + target parameter discovery
- services/mapping_service.py: smart matching logic
- services/transfer_service.py: write rules + transaction execution
- models/: DTOs for RoomParameter, TargetParameter, MappingRow, TransferResult
- utils/: conversion helpers and storage compatibility helpers

### 11.2 Key APIs
- FilteredElementCollector
- RevitLinkInstance.GetLinkDocument
- UIDocument.Selection.PickObject(ObjectType.LinkedElement)
- LookupParameter / GetParameters
- Parameter.StorageType / IsReadOnly / HasValue
- Transaction

## 12. UI/UX Design Direction
- dark-theme compatible palette
- compact dense rows for BIM productivity
- alternating visual groups to reduce scan fatigue
- responsive resizing behavior
- avoid modal spam
- keep controls discoverable and grouped by workflow order

## 13. Current Scaffold Status (This Commit)
Implemented in this pushbutton scaffold:
- linked model detection
- linked room manual selection
- category checklist with search
- element selection from current selection or model pick
- writable target parameter extraction
- common/non-common mode
- smart mapping preview
- transfer execution with duplicate mode and summary report

Planned next increment:
- full mapping table with per-row target dropdown override
- searchable per-row target dropdown
- mapping profile presets (save/load)
- category-grouped mapping mode
- log export

## 14. Icon Theme Convention
For this tool use both files:
- icon.png (light ribbon icon)
- icon.dark.png (dark ribbon icon)

MHT convention alignment:
- keep geometric/simple glyph style
- blue/cyan brand accent
- maintain visual clarity at small ribbon sizes

## 15. Future Feature Backlog
- automatic room detection / nearest room assignment
- batch room processing
- schedule integration
- Dynamo interoperability
- Excel import/export mapping templates
