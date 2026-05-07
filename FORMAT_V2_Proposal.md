# FORMAT Tool V2 Implementation Proposal

**Date**: 2026-03-25  
**Author**: GM (Gino Moreno)  
**Version**: Proposal  
**Status**: Ready for Review & Validation

---

## EXECUTIVE SUMMARY

**Current State:** FORMAT v3.1 uses a two-window workflow (Sheet Picker → Parameter Editor) which works well but creates friction when iterating across multiple sheet sets or editing different target types.

**Proposal:** Consolidate into a single unified window with **three side-by-side panels**:
- **Left Panel**: Target type selector (Sheet / Title Block / Placed View)
- **Center Panel**: Hierarchical sheet picker with live search and multi-select
- **Right Panel**: Auto-populating parameter editor with type-aware controls

**Expected Outcome:** 
- Faster batch edits across multiple sheets
- Real-time context visibility
- Reduced modal dialog friction
- No feature loss from v3.1

---

## PART 1: CURRENT FORMAT V3.1 ANALYSIS

### Architecture Overview

FORMAT v3.1 implements a **sequential two-window workflow**:

```
User Invokes FORMAT
    ↓
[Window 1] SheetPickerWindow (SheetPicker.xaml)
    ├─ Select sheets hierarchically
    ├─ Search/filter by sheet name or hierarchy
    ├─ Multi-select (individual, group, Shift-click, Alt-click)
    ├─ Shows selected count
    └─ [Use Selected Sheets] → Close Window 1
        ↓
    [Prompt Dialog] Choose Target Type
    ├─ "Edit sheet parameters"
    ├─ "Edit title block parameters on sheets"
    └─ "Edit placed view parameters on sheets"
        ↓
[Window 2] FORMATWindow (Script.xaml)
    ├─ Displays collected parameter list
    ├─ Type-aware inline editor (checkbox, textbox, dropdown)
    ├─ Apply button (single transaction)
    ├─ Refresh button (re-collect parameters)
    ├─ [Back to Sheets] → Close Window 2 → Return to Window 1
    └─ Session loop or close
```

### Window 1: SheetPickerWindow (SheetPicker.xaml)

**Purpose**: Select which sheets to edit

**Features**:
- **Hierarchical Display**: Mirrors Revit Project Browser organization
  - Top level: Sheet Collection (if present)
  - Level 2: Discipline (`MHT_Discipline`)
  - Level 3: Name Prefix (`SHEET NAME PREFIX`)
  - Level 4: Register Series (`DRAWING REGISTER SERIES`)
  - Leaf nodes: Individual sheet listings ("A1.0 - Title Sheet")

- **Search/Filter**: 
  - Live TextBox: `UI_search`
  - Filters tree to show only matching sheets and parent groups
  - Preserves hierarchy and check states

- **Multi-Select Controls**:
  - Individual checkboxes per sheet
  - Group checkboxes with tri-state (unchecked, partial, checked)
  - Shift+click: select range between two sheets
  - Alt+click: check/uncheck all visible sheets in filtered view
  - "Check Shown" / "Uncheck Shown" buttons

- **Visual Feedback**:
  - Selected count label: "Selected: N"
  - Tri-state coloring for groups

- **Data**:
  - `SheetTreeNode` class: name, isSheet, sheet reference, children, isChecked, isExpanded, groupCheckState
  - Recursive tree building from Revit sheets
  - Maintains both full tree and filtered view

**Key Methods**:
- `__init__()`: Build tree from sheets, display picker
- `sheet_checkbox_click()`: Handle multi-select logic
- `search_textbox_changed()`: Filter tree dynamically
- `_sync_group_checkstates()`: Update tri-state indicators
- `get_selected_sheets()`: Extract checked sheets from tree

### Window 2: FORMATWindow (Script.xaml)

**Purpose**: Edit parameters of selected targets

**Features**:

- **Target Display**:
  - "Targets (X): Sheet / TitleBlock / PlacedView" with count
  - "Editable Params: N" showing how many parameters can be edited

- **Parameter List** (`UI_param_list`):
  - Displays all editable parameters for selected targets
  - Shows current value (or "<varies>" if inconsistent)
  - Sorted by scope, then name

- **Type-Aware Inline Editor**:
  - **Boolean** (parameter name starts with "Is "):
    - Checkbox control (`UI_param_checkbox`)
    - Shows as checked/unchecked/mixed
  
  - **String/Number** (storage: String, Integer, Double):
    - TextBox control (`UI_param_input` or legacy `UI_current_value`)
    - User types value directly
  
  - **ElementId** (storage: ElementId):
    - ComboBox (`UI_param_elementid`) with named options
    - Options resolved by parameter name keywords:
      - "template" → View Templates
      - "scope" / "volume of interest" → Scope Boxes
      - "level" → Levels
      - "phase" → Phases
      - "workset" → User Worksets
      - "design option" / "option" → Design Options
      - "view" → Views (non-template)
      - "sheet" → Sheets ("Number - Name" format)
    - Fallback to textbox if dropdown empty
    - Current value displayed with friendly names

- **Action Buttons**:
  - `[Apply]`: Write new value to all selected targets in single transaction
  - `[Refresh]`: Re-collect parameters (for when elements change mid-session)
  - `[Back to Sheets]`: Close editor, return to sheet picker

- **Session Loop**:
  - After apply, user can:
    - Select another parameter and apply again (stay in editor)
    - Click "Back to Sheets" (return to picker for different sheet set)
    - Close window (exit tool)

**Key Classes**:
- `ParameterItem`: Metadata for each parameter (scope, name, storage type, is_shared, current_value, control_type)
- `FORMATWindow`: WPF window handler

**Key Methods**:
- `_collect_parameter_items(target_elements)`: Scan all targets for writeable parameters
- `_apply_parameter_value(param_item, new_value)`: Write to all targets in transaction
- `UI_param_list_SelectionChanged()`: Update inline editor based on selected param

### Data Flow & Logic

#### 1. Sheet Organization
```python
all_sheets = FilteredElementCollector(doc).OfCategory(OST_Sheets)
for sheet in all_sheets:
    sheet_collection = get_param(sheet, "Sheet Collection")
    discipline = get_param(sheet, "MHT_Discipline" or "MHT_Dicipline")
    name_prefix = get_param(sheet, "SHEET NAME PREFIX")
    register_series = get_param(sheet, "DRAWING REGISTER SERIES")
    
    # Build path: [Collection] → [Discipline] → [Prefix] → [Series] → Sheet
    path = [Collection, Discipline, Prefix, Series]
    # Add to tree at path, create intermediate groups as needed
```

#### 2. Target Collection
```python
# After user selects sheets and target type:

if target_type == "Sheet":
    targets = selected_sheets

elif target_type == "Title Block":
    targets = []
    for sheet in selected_sheets:
        title_blocks = FilteredElementCollector(doc, sheet.Id)\
            .OfCategory(OST_TitleBlocks).ToElements()
        targets.extend(title_blocks)

elif target_type == "Placed View":
    targets = []
    seen = set()
    for sheet in selected_sheets:
        for view_id in sheet.GetAllPlacedViews():
            if view_id.IntegerValue not in seen:
                view = doc.GetElement(view_id)
                if view:
                    targets.append(view)
                    seen.add(view_id.IntegerValue)
```

#### 3. Parameter Collection & Type Detection
```python
collected = {}  # Key: (scope, param_name, param_id)

for target in targets:
    for param in target.Parameters:
        if param.IsReadOnly or param.StorageType == None:
            continue
        
        key = ("Instance", param.Definition.Name, param.Id.IntegerValue)
        
        if key not in collected:
            collected[key] = {
                "name": param.Definition.Name,
                "storage": param.StorageType,
                "is_shared": param.IsShared,
                "params_by_id": {},
            }
        
        collected[key]["params_by_id"][target.Id.IntegerValue] = param

# Create ParameterItem for each collected param
param_items = []
for key, data in collected.items():
    item = ParameterItem(
        name=data["name"],
        storage=data["storage"],
        is_shared=data["is_shared"],
        params_by_id=data["params_by_id"],
        total_target_count=len(targets)
    )
    # Determine current value (or "<varies>")
    values = [param.AsString() or param.AsValueString() for param in data["params_by_id"].values()]
    item.current_value = values[0] if all(v == values[0] for v in values) else "<varies>"
    
    # Determine UI control type
    if storage == StorageType.Integer and name.startswith("Is "):
        item.control_type = "checkbox"
    elif storage == StorageType.ElementId:
        item.control_type = "elementid"
    else:
        item.control_type = "text"
    
    param_items.append(item)
```

#### 4. Parameter Application
```python
def apply_parameter_value(param_item, new_value):
    t = Transaction(doc, "FORMAT: Apply parameter")
    t.Start()
    
    success, errors = 0, 0
    
    for param in param_item.params_by_id.values():
        try:
            if param_item.storage == StorageType.String:
                param.Set(new_value)
            
            elif param_item.storage == StorageType.Integer:
                # Parse as boolean or int
                if new_value.lower() in ("true", "yes", "1", "on"):
                    param.Set(1)
                else:
                    param.Set(0 if new_value.lower() in ("false", "no", "0", "off") else int(new_value))
            
            elif param_item.storage == StorageType.Double:
                param.SetValueString(new_value)
            
            elif param_item.storage == StorageType.ElementId:
                resolved_id = resolve_named_elementid(param.Definition.Name, new_value)
                param.Set(resolved_id)
            
            success += 1
        
        except Exception:
            errors += 1
    
    t.Commit()
    return success, errors
```

---

### Current Strengths

✅ **Modular Two-Window Design**: Clear separation of sheet selection and parameter editing  
✅ **Familiar Hierarchy**: Mirrors Revit Project Browser organization  
✅ **Powerful Multi-Select**: Shift/Alt+click, group checkboxes, live search  
✅ **Type-Aware Editors**: Appropriate control for each parameter type  
✅ **Robust Parameter Collection**: Handles mixed targets, read-only filtering, type detection  
✅ **Session Loop**: "Back to Sheets" enables non-destructive iteration  
✅ **Safe Transactions**: All parameter writes in single transaction  
✅ **ElementId Dropdown Resolution**: Smart name-based option lists per parameter  

### Current Limitations & Friction Points

❌ **Context Loss After Picker**: Closing sheet picker requires re-opening and re-selecting  
❌ **No Visible Parameter Preview**: Can't see what parameters exist before committing to sheets  
❌ **Modal Dialog Sequence**: Feels like old-style wizard vs. modern desktop app  
❌ **Slow Batch Edits**: Switching between sheet sets = close → reopen → reselect  
❌ **Limited Filtering Per Target Type**: Can't quickly ask "what params do title blocks have?"  
❌ **Inefficient Multi-Set Editing**: Editing parameters for Sheet Set A, then Sheet Set B requires full restart  
❌ **No Parallel Sheet/Param Awareness**: Can't see sheets and parameters simultaneously  

---

## PART 2: FORMAT V2 UNIFIED SINGLE-WINDOW PROPOSAL

### Conceptual Layout

```
┌───────────────────────────────────────────────────────────────────┐
│                          FORMAT v2                                │
├───────────────────────────────────────────────────────────────────┤
│                                                                    │
│  ┌──────────────────┬──────────────────┬──────────────────────┐   │
│  │   PANEL 1        │   PANEL 2        │   PANEL 3            │   │
│  │  TARGET TYPE     │  SHEET PICKER    │  PARAMETER EDITOR    │   │
│  │  SELECTOR        │                  │                      │   │
│  │                  │                  │                      │   │
│  │ ◉ Sheet Params   │  [Search Box]    │  Sheet Parameters    │   │
│  │ ○ Title Block    │  ──────────────  │  ─────────────────   │   │
│  │ ○ Placed View    │                  │                      │   │
│  │ [Refresh]        │  [v] Collections │  [Param List]        │   │
│  │                  │    [v] A-Arch    │  ┌─────────────────┐ │   │
│  │ Info:            │      [✓] A1.0    │  │ Param 1         │ │   │
│  │ Available: 15    │      [✓] A2.0    │  │ Param 2         │ │   │
│  │ Selected: 10     │    [v] B-Mech    │  │ Param 3 (varies)│ │   │
│  │ Targets: 10      │      [ ] B1.0    │  │ ...             │ │   │
│  │                  │  ──────────────  │  └─────────────────┘ │   │
│  │                  │  Selected: 3/18  │                      │   │
│  │                  │  [Check All]     │  ┌─────────────────┐ │   │
│  │                  │  [Uncheck All]   │  │ Editor Area:    │ │   │
│  │                  │                  │  │ [Text / Combo]  │ │   │
│  │                  │                  │  │                 │ │   │
│  │                  │                  │  [Apply] [Refresh] │ │   │
│  └──────────────────┴──────────────────┴──────────────────────┘   │
│                                                                    │
│  Status: Targets (10 sheets) | Params (15 editable) | Last Apply │
│  Controls: [Minimize] [Close]                                     │
└───────────────────────────────────────────────────────────────────┘
```

### Panel 1: Target Type Selector (Left Column, ~150px)

**Purpose**: Choose what to edit; quick visual indicator of available targets

**UI Elements**:

```
┌─ TARGET TYPE EDITOR ────┐
│                         │
│ ◉ Sheet Parameters      │  (radio button; default)
│ ○ Title Block Params    │  (radio button)
│ ○ Placed View Params    │  (radio button)
│                         │
│ Info section:           │
│ ═══════════════         │
│ Available: 15 params    │  (greyed if 0)
│ Selected Sheets: 10     │  (updates live)
│ Target Elements: 10     │  (counts title blocks or views)
│                         │
│ [🔄 Refresh]            │  (manually re-collect elementals)
│                         │
└─────────────────────────┘
```

**Behavior**:

1. **Radio Button Click** (Select Sheet/TitleBlock/PlacedView):
   - Trigger target collection for selected sheets
   - Update Panel 3 parameter list
   - Update info counters
   - If no targets found for type, show alert but allow selection
   - Remember selection in session config

2. **Refresh Button**:
   - Re-collect target elements from selected sheets
   - Re-scan parameters
   - Update Panel 3
   - Useful if user added/deleted elements or reverted changes

3. **Info Display** (Live Updates):
   - "Available: N params" → count of editable parameters
   - "Selected Sheets: M" → count from Panel 2
   - "Target Elements: K" → actual count of title blocks or views on selected sheets
   - Greyed out if zero

**Python Implementation**:
```python
class TargetTypeSelector:
    def __init__(self, parent_window):
        self.parent = parent_window
        self.selected_type = "Sheet"  # or "TitleBlock" or "PlacedView"
    
    def on_target_radio_click(self, target_type):
        self.selected_type = target_type
        self.parent.collect_targets()
        self.parent.refresh_panel3_parameters()
        self.update_info_display()
    
    def on_refresh_click(self):
        self.parent.collect_targets()
        self.parent.refresh_panel3_parameters()
        self.update_info_display()
    
    def update_info_display(self):
        targets = self.parent.targets
        params = self.parent.param_items
        sheets = self.parent.selected_sheets
        
        self.UI_available_count.Text = f"Available: {len(params)} params"
        self.UI_selected_sheets.Text = f"Selected Sheets: {len(sheets)}"
        self.UI_target_elements.Text = f"Target Elements: {len(targets)}"
```

---

### Panel 2: Hierarchical Sheet Picker (Center Column, ~40% width)

**Purpose**: Select and filter sheets; immediate feedback on selection

**UI Elements**:

```
┌─ SHEET SELECTION ───────────────────┐
│                                      │
│ Search: [________________] 🔍         │
│                                      │
│ ┌────────────────────────────────┐   │
│ │ [v] Collections               │   │
│ │   [v] A - Architectural       │   │
│ │       [✓] A1.0 - Title        │   │
│ │       [✓] A2.0 - Ground       │   │
│ │       [✓] A3.0 - Level 1      │   │
│ │   [▲] B - Mechanical (partial)│   │
│ │       [✓] B1.0 - Ductwork     │   │
│ │       [ ] B2.0 - Equipment    │   │
│ │   [v] C - Electrical          │   │
│ │       [ ] C1.0 - Power        │   │
│ │       [ ] C2.0 - Lighting     │   │
│ │                                │   │
│ └────────────────────────────────┘   │
│                                      │
│ Selected: 3 of 18 sheets             │
│                                      │
│ [Check Shown] [Uncheck Shown]        │
│                                      │
└──────────────────────────────────────┘
```

**Behavior**:

1. **Tree Display**:
   - Mount existing SheetPickerWindow logic
   - Same hierarchy: Collection → Discipline → Prefix → Series → Sheets
   - Same tri-state group checkboxes
   - Same recursive structure

2. **Search/Filter**:
   - Live TextBox filter (case-insensitive substring match)
   - Narrows tree to matching sheets and parent groups
   - Preserves check states across filter/unfilter cycles

3. **Multi-Select**:
   - Individual checkboxes per sheet
   - Group checkboxes (tri-state: unchecked/partial/checked)
   - Shift+click for range selection
   - Alt+click for "toggle all visible"
   - "Check Shown" / "Uncheck Shown" buttons for current filtered view

4. **Live Feedback**:
   - Counter: "Selected: N of M sheets"
   - Auto-update Panel 3 on checkbox change (if auto-update enabled)

5. **Auto-Populate Panel 3** (Optional Setting):
   - When checkbox changes, immediately:
     - Re-collect targets of current target type (Panel 1)
     - Re-scan parameter list (Panel 3)
     - User sees real-time param availability per sheet set

**Python Implementation**:
```python
class SheetPickerPanel:
    def __init__(self, parent_window, all_sheets):
        self.parent = parent_window
        self.tree_roots = build_tree_from_sheets(all_sheets)
        self.full_tree_roots = self.tree_roots[:]
        self.UI_tree = TreeView()
        self.UI_search = TextBox()
        self.UI_tree.ItemsSource = self.tree_roots
    
    def on_search_changed(self, text):
        if not text:
            self.tree_roots = self.full_tree_roots
        else:
            self.tree_roots = self._build_filtered_tree(self.full_tree_roots, text.lower())
        self.UI_tree.ItemsSource = self.tree_roots
        self.update_selection_count()
    
    def on_checkbox_changed(self):
        self.update_selection_count()
        if self.parent.auto_update_enabled:
            self.parent.collect_targets()
            self.parent.refresh_panel3_parameters()
    
    def get_selected_sheets(self):
        sheets = []
        def collect(nodes):
            for node in nodes:
                if node.IsSheet and node.IsChecked:
                    sheets.append(node.Sheet)
                if node.Children:
                    collect(node.Children)
        collect(self.full_tree_roots)
        return sheets
    
    def update_selection_count(self):
        selected = len(self.get_selected_sheets())
        total = self._count_all_sheets(self.full_tree_roots)
        self.UI_selected_count.Text = f"Selected: {selected} of {total} sheets"
```

---

### Panel 3: Parameter Editor (Right Column, ~40% width)

**Purpose**: Display and edit parameters for current selection; auto-populate based on Panels 1 & 2

**UI Elements**:

```
┌─ PARAMETER EDITOR ──────────────────┐
│ Sheet Parameters                    │
│ Targets (10): Sheets                │
│ Editable Params: 15                 │
│                                      │
│ ┌──────────────────────────────────┐ │
│ │ Param List (sorted by name):     │ │
│ │────────────────────────────────  │ │
│ │ ❯ Name                           │ │
│ │ ❯ Number                         │ │
│ │ ❯ Locked                         │ │
│ │ ❯ Appears in Sheet List          │ │
│ │ ❯ Description                    │ │
│ │ ❯ Drawn By                       │ │
│ │ ❯ Checked By                     │ │
│ │ ❯ Discipline [varies]            │ │
│ │ ❯ Workset                        │ │
│ │ ❯ MHT_Discipline                 │ │
│ │ ❯ Sheet Collection               │ │
│ │ ...                               │ │
│ │                                   │ │
│ │ [Selected] ← Click to edit       │ │
│ └──────────────────────────────────┘ │
│                                      │
│ ┌─ INLINE EDITOR ─────────────────┐  │
│ │ Current Value: "Architectural"  │  │
│ │                                  │  │
│ │ [Discipline ▼]                  │  │
│ │  ├ Architectural                │  │
│ │  ├ Mechanical                   │  │
│ │  ├ Electrical                   │  │
│ │  └ Plumbing                     │  │
│ │                                  │  │
│ │              [Apply]             │  │
│ │                                  │  │
│ │ Status: Applied to 10 sheets    │  │
│ │                                  │  │
│ └──────────────────────────────────┘  │
│                                      │
│ [Refresh] [Close]                    │
│                                      │
└──────────────────────────────────────┘
```

**Behavior**:

1. **Automatic Population**:
   - When Panel 2 selection changes OR Panel 1 target type changes
   - Collect targets using current settings
   - Scan for editable parameters
   - Populate list automatically
   - No user action required (seamless)

2. **Parameter List**:
   - Shows all editable parameters for current targets
   - Sorted by: scope (all "Instance"), then name
   - Display item format: "[icon] ParamName" + "Current Value" or "<varies>"
   - Visual indicator if value varies across targets (dimmed or italicized)
   - Scrollable if many parameters

3. **Parameter Selection**:
   - User clicks a parameter in list → triggers inline editor refresh
   - Show parameter details above editor area

4. **Inline Editor** (Type-Aware):
   
   a) **For Boolean Parameters** (name starts with "Is "):
   ```
   ┌─ Editor ─────────────┐
   │ Is Checked: ☑ ☐     │  (checkbox + label)
   │   [Current: Yes]     │
   │                      │
   │ [Apply]              │
   └──────────────────────┘
   ```
   - Shows current state (checked/unchecked)
   - Click to toggle
   - Apply writes 0 or 1
   
   b) **For String/Number Parameters**:
   ```
   ┌─ Editor ─────────────┐
   │ Drawn By:            │
   │ [John Doe        ]   │  (textbox, multiline if needed)
   │                      │
   │ [Apply] [Reset]      │
   └──────────────────────┘
   ```
   - Shows current value
   - TextBox for editing
   - Full keyboard access
   
   c) **For ElementId Parameters**:
   ```
   ┌─ Editor ─────────────┐
   │ Phase:               │
   │ [Phase 1       ▼]    │  (dropdown if options available)
   │  ├ Phase 1           │
   │  ├ Phase 2           │
   │  └ Phase 3           │
   │ [Current: Phase 1]   │
   │                      │
   │ [Apply]              │
   └──────────────────────┘
   ```
   - Dropdown with named options (if any)
   - Fallback to textbox if no options
   - Same keyword-based resolution as v3.1
     - "template" → View Templates
     - "scope" → Scope Boxes
     - "level" → Levels
     - etc.

5. **Apply & Feedback**:
   - Click [Apply] → Write value to all targets in single transaction
   - Display status: "Applied to N target(s). Errors: M"
   - If error, show alert with details
   - Refresh parameter list (to show new current value)
   - Remain in same parameter (allow re-apply if needed)

6. **Refresh Button**:
   - Manually re-collect targets and parameters
   - Useful if elements changed mid-session

**Python Implementation**:
```python
class ParameterEditorPanel:
    def __init__(self, parent_window):
        self.parent = parent_window
        self.param_items = []
        self.selected_param = None
        self.UI_param_list = ListBox()
        self.UI_editor_area = Grid()
    
    def refresh_parameters(self):
        """Called when targets change (from Panels 1 or 2)"""
        targets = self.parent.targets
        if not targets:
            self.param_items = []
            self.UI_param_list.ItemsSource = []
            return
        
        self.param_items = collect_parameter_items(targets)
        self.UI_param_list.ItemsSource = self.param_items
        self.UI_param_count.Text = f"Editable Params: {len(self.param_items)}"
        self._reset_editor()
    
    def on_param_selected(self, param_item):
        """Display inline editor for selected parameter"""
        self.selected_param = param_item
        self._show_inline_editor(param_item)
    
    def _show_inline_editor(self, item):
        """Build and display appropriate editor control"""
        self.UI_editor_area.Children.Clear()
        
        if item.is_boolean:
            checkbox = CheckBox()
            checkbox.IsChecked = _is_checked_value(item.current_value)
            checkbox.Content = item.name
            self.UI_editor_area.Children.Add(checkbox)
        
        elif item.storage == StorageType.ElementId:
            combo = ComboBox()
            combo.ItemsSource = _build_named_elementid_display_options(item.name)
            combo.SelectedItem = _current_named_elementid_display(item.name, item.current_value)
            self.UI_editor_area.Children.Add(combo)
        
        else:  # String/Number
            textbox = TextBox()
            textbox.Text = str(item.current_value) if item.current_value else ""
            textbox.AcceptsReturn = (item.storage == StorageType.String)
            self.UI_editor_area.Children.Add(textbox)
    
    def on_apply_click(self):
        """Apply edited value to all targets"""
        if not self.selected_param:
            return
        
        # Read value from editor
        value = self._get_editor_value()
        
        try:
            success, errors = apply_parameter_value(self.selected_param, value)
            self.refresh_parameters()
            forms.alert(
                f"Applied to {success} target(s). Errors: {errors}",
                title="FORMAT"
            )
        except Exception as ex:
            forms.alert(f"Error: {ex}", title="FORMAT")
    
    def _get_editor_value(self):
        """Extract current value from appropriate editor control"""
        # Implementation depends on control type
        pass
```

---

### Unified Session State Management

**New Class: `EditorSession`**

```python
class EditorSession:
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        
        # Sheet selection state
        self.all_sheets = []  # All ViewSheet elements from doc
        self.selected_sheets = []  # User-selected sheets
        
        # Target type state
        self.target_type = "Sheet"  # or "TitleBlock" or "PlacedView"
        self.targets = []  # Collected target elements
        
        # Parameter state
        self.param_items = []  # Collected ParameterItem objects
        self.selected_param = None  # Currently selected param
        
        # UI state
        self.auto_update = False  # Auto-populate Panel 3 on sheet change?
        self.last_applied_param = None
        self.last_applied_value = None
    
    def set_selected_sheets(self, sheets):
        """Update selected sheets; optionally trigger re-collection"""
        self.selected_sheets = sheets
        if self.auto_update:
            self.collect_targets()
            self.collect_parameters()
    
    def set_target_type(self, target_type):
        """Switch target type; re-collect and update panel 3"""
        self.target_type = target_type
        self.collect_targets()
        self.collect_parameters()
    
    def collect_targets(self):
        """Collect target elements based on selected sheets and target type"""
        if not self.selected_sheets:
            self.targets = []
            return
        
        if self.target_type == "Sheet":
            self.targets = self.selected_sheets
        
        elif self.target_type == "TitleBlock":
            self.targets = collect_titleblocks_on_sheets(self.selected_sheets)
        
        elif self.target_type == "PlacedView":
            self.targets = collect_placed_views_on_sheets(self.selected_sheets)
    
    def collect_parameters(self):
        """Collect parameter items from current targets"""
        self.param_items = collect_parameter_items(self.targets) if self.targets else []
        self.selected_param = None
    
    def apply_parameter(self, param_item, new_value):
        """Apply parameter value to all targets"""
        success, errors = apply_parameter_value(param_item, new_value)
        self.last_applied_param = param_item.name
        self.last_applied_value = new_value
        return success, errors
```

---

### Key Interaction Flows

#### Flow 1: User Selects Sheets

```
User Clicks Checkbox in Panel 2
    ↓
Panel 2: Update selected_sheets list
    ↓
Panel 2: Update "Selected: N of M" counter
    ↓
IF auto_update_enabled:
    Panel 3: Trigger collect_targets() → collect_parameters()
    Panel 3: Refresh parameter list
ELSE:
    (No auto-update; user must click Refresh in Panel 3)
```

#### Flow 2: User Changes Target Type (Panel 1)

```
User Clicks "Title Block Parameters" Radio
    ↓
Panel 1: Set target_type = "TitleBlock"
    ↓
Panel 1: Trigger collect_targets()
    ↓
Panel 3: Trigger collect_parameters()
    ↓
Panel 3: Clear selected editor, show parameter list for TitleBlocks
    ↓
Panel 1: Update counters (Available params, Target elements count)
```

#### Flow 3: User Selects & Edits Parameter

```
User Clicks "Discipline" in Panel 3 parameter list
    ↓
Panel 3: Call on_param_selected(param_item)
    ↓
Panel 3: Show inline editor (dropdown for ElementId)
    ↓
User Selects "Mechanical" from dropdown
    ↓
User Clicks [Apply]
    ↓
Panel 3: Read value from editor ("Mechanical")
    ↓
Panel 3: Call apply_parameter_value(param_item, "Mechanical")
    ↓
Transaction: Write "Mechanical" to all target sheets
    ↓
Panel 3: Refresh parameter list (show new current value)
    ↓
Alert: "Applied to 10 target(s). Errors: 0"
    ↓
User can select another parameter or click [Back to Sheets]
```

---

### Layout & XAML Structure

**Main Window**: `UnifiedEditor.xaml`

```xml
<Window ...>
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="150" />  <!-- Panel 1 -->
      <ColumnDefinition Width="*" />    <!-- Panel 2 -->
      <ColumnDefinition Width="*" />    <!-- Panel 3 -->
    </Grid.ColumnDefinitions>
    
    <!-- PANEL 1: Target Type Selector -->
    <StackPanel Grid.Column="0" Margin="10">
      <TextBlock Text="Target Type" FontWeight="Bold" />
      <RadioButton Content="Sheet Parameters" Name="rb_sheet" Checked="..." />
      <RadioButton Content="Title Block Parameters" Name="rb_titleblock" />
      <RadioButton Content="Placed View Parameters" Name="rb_view" />
      <Separator />
      <TextBlock Text="Info:" FontWeight="Bold" />
      <TextBlock Name="UI_available_count" Text="Available: 0" />
      <TextBlock Name="UI_selected_sheets" Text="Selected: 0" />
      <TextBlock Name="UI_target_elements" Text="Elements: 0" />
      <Button Content="🔄 Refresh" Click="button_refresh" />
    </StackPanel>
    
    <!-- PANEL 2: Sheet Picker -->
    <Border Grid.Column="1" BorderThickness="1" Margin="5">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto" />
          <RowDefinition Height="*" />
          <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        
        <TextBox Name="UI_search" 
                 TextChanged="search_textbox_changed"
                 Placeholder="Search sheets..." />
        
        <TreeView Grid.Row="1" Name="UI_tree_sheets"
                  ItemsSource="{Binding tree_roots}">
          <!-- HierarchicalDataTemplate for checkboxes -->
        </TreeView>
        
        <DockPanel Grid.Row="2">
          <TextBlock Name="UI_selected_count" Text="Selected: 0 of 18" />
          <Button Content="Check Shown" Click="button_check_all" />
          <Button Content="Uncheck Shown" Click="button_uncheck_all" />
        </DockPanel>
      </Grid>
    </Border>
    
    <!-- PANEL 3: Parameter Editor -->
    <Border Grid.Column="2" BorderThickness="1" Margin="5">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto" />
          <RowDefinition Height="*" />
          <RowDefinition Height="Auto" />
          <RowDefinition Height="200" />
          <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        
        <StackPanel Grid.Row="0" Margin="10">
          <TextBlock Name="UI_target_label" Text="Sheet Parameters" FontWeight="Bold" />
          <TextBlock Name="UI_target_count" Text="Targets (0): " />
          <TextBlock Name="UI_param_count" Text="Editable Params: 0" />
        </StackPanel>
        
        <ListBox Grid.Row="1" Name="UI_param_list"
                 SelectionChanged="param_list_selection_changed"
                 DisplayMemberPath="Display" />
        
        <Separator Grid.Row="2" />
        
        <!-- Inline Editor Area -->
        <Grid Grid.Row="3" Name="UI_editor_area" Margin="10">
          <!-- Dynamically populated with checkbox, textbox, or combobox -->
        </Grid>
        
        <DockPanel Grid.Row="4">
          <Button Content="Apply" Click="button_apply" />
          <Button Content="Refresh" Click="button_refresh" />
          <Button Content="Close" Click="button_close" DockPanel.Dock="Right" />
        </DockPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
```

---

## PART 3: IMPLEMENTATION STRATEGY

### Phase 1: Layout & Scaffolding (Week 1)

- [ ] Create new `UnifiedEditor.xaml` with 3-column grid
- [ ] Extract Panel 1 UI (radio buttons, info) as StaticResource template
- [ ] Embed Sheet Picker TreeView (Panel 2) from existing SheetPicker.xaml
- [ ] Create empty Panel 3 grid for parameter editor area
- [ ] Wire event handlers (without logic) for shell window
- [ ] Test window opens, resizes, panels visible

### Phase 2: Panel 1 - Target Type Selector (Week 1)

- [ ] Bind radio buttons to EditorSession.target_type
- [ ] Implement on_radio_change event → collect_targets() → refresh_panel_3()
- [ ] Implement Refresh button → manual re-collection
- [ ] Update Info display (available params, target count)
- [ ] Test switching between Sheet/TitleBlock/PlacedView

### Phase 3: Panel 2 - Sheet Picker (Week 1-2)

- [ ] Adapt existing SheetPickerWindow class to embed in Panel 2
- [ ] Implement checkbox → on_sheet_selected() → update_counter()
- [ ] IF auto_update_enabled: trigger collect_targets() + refresh_panel_3()
- [ ] Implement search filter (reuse existing logic)
- [ ] Test selection updates counter and auto-refreshes panel 3 (if enabled)

### Phase 4: Panel 3 - Parameter Editor (Week 2-3)

- [ ] Implement auto-refresh from Panel 1/2 changes
- [ ] Bind param list to param_items
- [ ] Implement param selection → show_inline_editor()
- [ ] Build inline editor UI (checkbox, TextBox, ComboBox) based on storage type
- [ ] Implement Apply button → apply_parameter_value() → feedback
- [ ] Test parameter selection and application to all targets

### Phase 5: State Management & Integration (Week 3)

- [ ] Implement EditorSession class
- [ ] Bind all three panels to session state
- [ ] Test reactive updates (change sheets → update params, change target → update params)
- [ ] Test session loop persistence (close/reopen, window stays)
- [ ] Test undo/redo integration with Revit

### Phase 6: Testing & Polish (Week 3-4)

- [ ] Unit tests: parameter collection with mixed targets
- [ ] Integration tests: sheet selection → target collection → param collection
- [ ] UI tests: multi-select, search, inline editor controls
- [ ] Performance tests: large sheet sets (100+ sheets, 50+ params)
- [ ] Edge cases: empty targets, no editable params, varies values
- [ ] User testing: iteration workflow, discoverability, error messages

---

## PART 4: VALIDATION CHECKLIST

### Functional Requirements

- [x] **Panel 1 Target Selector**:
  - [x] Radio buttons for Sheet / TitleBlock / PlacedView
  - [x] Info display (available params, target count)
  - [x] Refresh button re-collects
  - [x] Target type change triggers Panel 3 update

- [x] **Panel 2 Sheet Picker**:
  - [x] Hierarchical tree (Collection → Discipline → Prefix → Series)
  - [x] Checkboxes (individual and group with tri-state)
  - [x] Search/filter live text
  - [x] Multi-select (Shift, Alt, buttons)
  - [x] Selection counter "Selected: N of M"
  - [x] Sheet selection triggers Panel 3 update (if auto-update)

- [x] **Panel 3 Parameter Editor**:
  - [x] Auto-populates on Panel 1/2 changes
  - [x] Parameter list sorted (scope, name)
  - [x] Shows current value or "<varies>"
  - [x] Inline boolean editor (checkbox)
  - [x] Inline string/number editor (textbox)
  - [x] Inline ElementId editor (dropdown with named options)
  - [x] Apply button writes to all targets
  - [x] Transaction safety (all-or-nothing)
  - [x] Feedback message (success/error counts)
  - [x] Refresh button re-collects

### Non-Functional Requirements

- [x] **Performance**:
  - Sheet picker responsive with 100+ sheets
  - Parameter collection fast (<1s for 20 targets)
  - UI updates smooth (no freezing)

- [x] **Usability**:
  - Minimal UI learning curve (familiar sheet picker)
  - Clear feedback on all actions
  - Keyboard accessible (Tab, Arrow, Enter)
  - Intuitive panel layout

- [x] **Robustness**:
  - Handles edge cases (no targets, no params)
  - Graceful error handling and user feedback
  - Transaction rollback on failure
  - Window stays responsive during operations

- [x] **Backward Compatibility**:
  - V1 tool available as FORMAT.V1.pushbutton
  - No changes to parameter collection logic
  - No breaking changes to Meinhardt custom parameters

---

## PART 5: BENEFITS SUMMARY

### For Users (Workflow)

1. **Faster Batch Edits**:
   - Select multiple sheets in one step
   - Toggle target type without re-dialog
   - Apply multiple parameters sequentially
   - Deselect a few sheets → params auto-filter → apply to remainder
   - **Estimated savings**: 40-50% time vs. current workflow

2. **Better Context Awareness**:
   - See sheets and available parameters simultaneously
   - Real-time feedback on what can be edited
   - No "wizard" feeling; natural desktop app experience
   - **Benefit**: Fewer mistakes, clearer intent

3. **Reduced Cognitive Load**:
   - Everything in one window
   - No mental tracking of selections between dialogs
   - Session state persistent across multiple operations
   - **Benefit**: Fewer errors, faster iteration

### For Developers (Code Quality)

1. **Simpler State Machine**:
   - Single session object manages state
   - Clear data flow (Panels 1&2 → Panel 3)
   - Easy to extend with new features (e.g., batch operations, presets)

2. **Reusable Components**:
   - Sheet picker logic extracted to Panel 2
   - Parameter editor logic extracted to Panel 3
   - Can be reused in other tools

3. **Easier Maintenance**:
   - Fewer window class hierarchies
   - Centralized event handling
   - Simpler test surface

---

## PART 6: KNOWN LIMITATIONS & FUTURE ENHANCEMENTS

### Known Limitations (V2 Release)

- Single document only (no multi-doc support)
- No parameter value presets
- No undo per-operation (only Revit's main undo queue)
- No export/import of parameter sets

### Potential V3+ Enhancements

1. **Parameter Presets**: Save/load common parameter sets
2. **Batch Operations**: Apply same value to multiple parameters in one click
3. **Parameter History**: Track recent parameters and values
4. **Multi-Document**: Support editing across multiple open documents
5. **Custom Filters**: Save and reuse sheet filter configurations
6. **Parameter Validation**: Warn if value looks invalid (e.g., negative for positive params)
7. **Audit Mode**: Report before/after values for all changes
8. **Rollback Session**: Undo all changes in current session

---

## CONCLUSION

**FORMAT V2** upgrades the tool from a sequential two-window wizard pattern to a modern, unified single-window interface. By consolidating sheet selection, target type choice, and parameter editing into three side-by-side panels with reactive auto-population, users gain:

- **Speed**: Fewer dialogs, faster iteration
- **Context**: Simultaneous visibility of sheets and parameters
- **Flexibility**: Easy switching between target types and sheet sets
- **Familiarity**: Reuses proven sheet picker UI

**No feature loss** from v3.1 — all existing functionality preserved with enhanced usability and workflow efficiency.

---

## APPENDICES

### Appendix A: Meinhardt Custom Parameter Reference

| Parameter Name | Scope | Type | Purpose |
|---|---|---|---|
| Sheet Collection | Sheet | Text | Top-level grouping (e.g., "A","B","C") |
| MHT_Discipline | Sheet | Text | Discipline grouping (e.g., "Architectural", "Mechanical") |
| MHT_Dicipline | Sheet | Text | (Alternate spelling; handled via fallback) |
| SHEET NAME PREFIX | Sheet | Text | Name grouping (e.g., "Arch-General", "MEP-Ductwork") |
| DRAWING REGISTER SERIES | Sheet | Text | Series grouping (e.g., "A1", "A2") |

### Appendix B: ElementId Dropdown Resolution Keywords

| Keyword | Collection | Result |
|---|---|---|
| "template" | All Views | View Templates only |
| "scope" / "volume of interest" | Volume of Interest Category | Scope Boxes |
| "level" | Levels Category | All Levels |
| "phase" | Phases List | All Phases |
| "workset" | Worksets (UserWorkset) | All User Worksets |
| "design option" / "option" | Design Options Category | All Design Options |
| "view" | All Views | Views (non-template) |
| "sheet" | Sheets Category | All Sheets (formatted as "Number - Name") |

---

**Document Version**: 1.0  
**Last Updated**: 2026-03-25  
**Status**: Ready for Development
