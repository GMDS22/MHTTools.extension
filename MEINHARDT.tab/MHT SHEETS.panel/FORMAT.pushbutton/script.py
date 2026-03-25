# -*- coding: utf-8 -*-
__title__ = "FORMAT"
__author__ = "Gino Moreno (GM)"
__doc__ = """Version = 4.0
Date    = 2026-03-25
Author  = GM
_____________________________________________________________________
Description:

FORMAT V3 - Meinhardt sheet formatting tool by GM.

All controls are visible simultaneously in three side-by-side panels:

  Panel 1 (Left)   — Target type selector: Sheet / Title Block / Placed View
  Panel 2 (Center) — Hierarchical sheet picker with live search & multi-select
  Panel 3 (Right)  — Auto-populating parameter editor with type-aware controls

Features:
  - Hierarchical sheet picker mirroring the Revit Project Browser grouping.
  - Search/filter field to quickly narrow down sheets.
  - Group checkboxes with tri-state (partial) indication.
  - Shift-click for range selection; Alt-click to check/uncheck all visible.
  - Type-aware parameter editor:
      String / Number  ->  text field
      Boolean          ->  checkbox
      ElementId        ->  named dropdown (Scope Box, View Template, Level,
                           Phase, Workset, Design Option, View, Sheet)
  - Target type toggle (Sheet / TitleBlock / PlacedView) without re-invoking.
  - Live info counters update as sheets are selected/deselected.
    - Single-window session: modal for stable event handling.
    - Smart Auto Fill from sheet number/name tokens (no Excel required).
    - Smart Align Viewports centered to titleblock bounds per sheet.
        - Keyplan visibility auto-toggle on titleblocks by zone + view scale.
        - Placed view Scope Box auto-update during Smart Auto Fill.
_____________________________________________________________________
How-to:

1. Run FORMAT.
2. In the CENTER PANEL, tick sheets to edit (Search to filter, group
   checkboxes for bulk select, Shift-click for range, Alt-click for all).
3. In the LEFT PANEL, choose what to edit:
      • Sheet Parameters
      • Title Block Parameters
      • Placed View Parameters
   The right panel auto-populates with available parameters.
4. In the RIGHT PANEL, click a parameter, set the value, click Apply.
5. Switch sheets or target type at any time and repeat.
_____________________________________________________________________
"""

from Autodesk.Revit.DB import (
        FilteredElementCollector,
        FilteredWorksetCollector,
        BuiltInCategory,
    Revision,
    RevisionNumberType,
        ViewSheet,
        ViewType,
        View,
        ElementId,
        XYZ,
        StorageType,
        Transaction,
        WorksetKind,
)
from System.Collections.Generic import List
import clr
clr.AddReference("System")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
import System
from System.Windows.Forms import Control, Keys
from System.Diagnostics.Process import Start
from pyrevit import forms, script
import re
import datetime


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None


# ─────────────────────────────────────────────────────────────────────────────
#  SHARED TREE NODE
# ─────────────────────────────────────────────────────────────────────────────

class SheetTreeNode(object):
    """Tree node used by the sheet picker (both groups and leaf sheets)."""
    def __init__(self, name, is_sheet=False, sheet=None):
        self.Name = name
        self.NameLower = (name or "").lower()
        self.IsSheet = is_sheet
        self.Sheet = sheet
        self.SheetSearchText = self.NameLower
        self.Children = []
        self.IsChecked = False
        self.IsExpanded = True
        self.GroupCheckState = False


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET TREE BUILDER  (shared helper; no UI references)
# ─────────────────────────────────────────────────────────────────────────────

def _clean_group_value(value):
    txt = (value or "").strip()
    if not txt or txt.lower() in ("<none>", "none"):
        return ""
    return txt


def _get_sheet_param_text(sheet, candidate_names):
    wanted = set(n.replace("_", " ").strip().lower() for n in candidate_names)
    try:
        for p in sheet.Parameters:
            try:
                defn = getattr(p, "Definition", None)
                pname = defn.Name if defn else ""
                norm = pname.replace("_", " ").strip().lower() if pname else ""
                if norm not in wanted:
                    continue
                val = ""
                try:
                    val = p.AsString() or ""
                except Exception:
                    pass
                if not val:
                    try:
                        val = p.AsValueString() or ""
                    except Exception:
                        pass
                return _clean_group_value(val)
            except Exception:
                continue
    except Exception:
        pass
    return ""


def _build_sheet_tree(all_sheets, preselected_ids=None):
    """Return (roots, root_index) from a collection of ViewSheet elements."""
    preselected_ids = preselected_ids or set()

    has_any_sheet_collection = any(
        _get_sheet_param_text(s, ["Sheet Collection"])
        for s in all_sheets
    )

    def _browser_path(sheet):
        sheet_collection = _get_sheet_param_text(sheet, ["Sheet Collection"])
        discipline = _get_sheet_param_text(sheet, ["MHT_Dicipline", "MHT_Discipline"])
        name_prefix = _get_sheet_param_text(sheet, ["SHEET NAME PREFIX", "Sheet Name Prefix"])
        register_series = _get_sheet_param_text(sheet, ["DRAWING REGISTER SERIES", "Drawing Register Series"])
        path = []
        if has_any_sheet_collection:
            path.append(sheet_collection if sheet_collection else "<No Sheet Collection>")
        if discipline:
            path.append(discipline)
        if name_prefix:
            path.append(name_prefix)
        if register_series:
            path.append(register_series)
        return path if path else ["Ungrouped"]

    roots = []
    root_index = {}

    def _get_or_create_root(name):
        key = (None, name.lower())
        node = root_index.get(key)
        if node is None:
            node = SheetTreeNode(name, is_sheet=False)
            roots.append(node)
            root_index[key] = node
        return node

    def _get_or_create_child(parent, child_name):
        key = (id(parent), child_name.lower())
        child = root_index.get(key)
        if child is None:
            child = SheetTreeNode(child_name, is_sheet=False)
            parent.Children.append(child)
            root_index[key] = child
        return child

    for sheet in all_sheets:
        group_path = _browser_path(sheet)
        current = _get_or_create_root(group_path[0])
        for folder_name in group_path[1:]:
            current = _get_or_create_child(current, folder_name)
        label = "{} - {}".format(sheet.SheetNumber, sheet.Name)
        sheet_node = SheetTreeNode(label, is_sheet=True, sheet=sheet)
        sheet_node.SheetSearchText = label.lower()
        sheet_node.IsChecked = sheet.Id.IntegerValue in preselected_ids
        current.Children.append(sheet_node)

    def _sort_nodes(nodes):
        groups = [n for n in nodes if not n.IsSheet]
        sheets = [n for n in nodes if n.IsSheet]
        groups.sort(key=lambda n: n.Name.lower())
        sheets.sort(key=lambda n: n.Name.lower())
        for g in groups:
            _sort_nodes(g.Children)
        nodes[:] = groups + sheets

    _sort_nodes(roots)
    return roots, root_index


# ─────────────────────────────────────────────────────────────────────────────
#  PARAMETER HELPERS  (all reused from v3.1 logic; no changes)
# ─────────────────────────────────────────────────────────────────────────────

_ELEMENTID_OPTIONS_CACHE = {}

def _storage_name(storage_type):
    names = {
        StorageType.String: "String",
        StorageType.Integer: "Integer",
        StorageType.Double: "Double",
        StorageType.ElementId: "ElementId",
    }
    return names.get(storage_type, "Unknown")


def _collect_named_elementid_options(param_name):
    pname = (param_name or "").strip().lower()
    if pname in _ELEMENTID_OPTIONS_CACHE:
        return _ELEMENTID_OPTIONS_CACHE[pname]

    options = []

    if "template" in pname:
        for v in FilteredElementCollector(doc).OfClass(View).ToElements():
            try:
                if v.IsTemplate:
                    options.append((str(v.Name), v.Id))
            except Exception:
                continue
        _ELEMENTID_OPTIONS_CACHE[pname] = options
        return options

    if "scope" in pname or "volume of interest" in pname:
        for sb in FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(sb.Name), sb.Id))
            except Exception:
                continue
        _ELEMENTID_OPTIONS_CACHE[pname] = options
        return options

    if "level" in pname:
        for lv in FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(lv.Name), lv.Id))
            except Exception:
                continue
        _ELEMENTID_OPTIONS_CACHE[pname] = options
        return options

    if "phase" in pname:
        try:
            for ph in doc.Phases:
                options.append((str(ph.Name), ph.Id))
        except Exception:
            pass
        _ELEMENTID_OPTIONS_CACHE[pname] = options
        return options

    if "workset" in pname:
        try:
            for ws in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets():
                options.append((str(ws.Name), ElementId(ws.Id.IntegerValue)))
        except Exception:
            pass
        _ELEMENTID_OPTIONS_CACHE[pname] = options
        return options

    if "design option" in pname or "option" in pname:
        for do in FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_DesignOptions).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(do.Name), do.Id))
            except Exception:
                continue
        _ELEMENTID_OPTIONS_CACHE[pname] = options
        return options

    if "view" in pname:
        for v in FilteredElementCollector(doc).OfClass(View).ToElements():
            try:
                if not v.IsTemplate:
                    options.append((str(v.Name), v.Id))
            except Exception:
                continue
        _ELEMENTID_OPTIONS_CACHE[pname] = options
        return options

    if "sheet" in pname:
        for s in FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements():
            try:
                label = "{} - {}".format(s.SheetNumber, s.Name)
                options.append((label, s.Id))
            except Exception:
                continue
        _ELEMENTID_OPTIONS_CACHE[pname] = options
        return options

    _ELEMENTID_OPTIONS_CACHE[pname] = options
    return options


def _build_named_elementid_display_options(param_name):
    display_names = ["<None>"]
    seen = set()
    for display_name, _ in _collect_named_elementid_options(param_name):
        low = display_name.lower()
        if low in seen:
            continue
        seen.add(low)
        display_names.append(display_name)
    return sorted(display_names[:1]) + sorted(display_names[1:], key=lambda x: x.lower())


def _resolve_named_elementid_from_text(param_name, text_value):
    raw = (text_value or "").strip()
    if not raw or raw.lower() in ("none", "<none>", "-"):
        return ElementId(-1)
    try:
        return ElementId(int(raw))
    except Exception:
        pass
    for option_name, option_id in _collect_named_elementid_options(param_name):
        try:
            if option_name.strip().lower() == raw.lower():
                return option_id
        except Exception:
            continue
    return None


def _current_named_elementid_display(param_name, current_value, prefer_options=False):
    if current_value in (None, "", "<varies>"):
        return "<None>"
    try:
        current_id = int(str(current_value).strip())
    except Exception:
        return str(current_value)
    if current_id < 0:
        return "<None>"

    # Fast path: direct element lookup avoids expensive model-wide collectors
    # when rendering the full parameter list for many targets.
    try:
        elem = doc.GetElement(ElementId(current_id))
        name = getattr(elem, "Name", None) if elem else None
        if name:
            return str(name)
    except Exception:
        pass

    if prefer_options:
        for option_name, option_id in _collect_named_elementid_options(param_name):
            try:
                if option_id.IntegerValue == current_id:
                    return option_name
            except Exception:
                continue

    return str(current_id)


def _parameter_value_text(param):
    st = param.StorageType
    if st == StorageType.String:
        val = param.AsString()
        return "" if val is None else val
    if st == StorageType.Integer:
        try:
            if param.Definition.Name.lower().startswith("is "):
                return "True" if param.AsInteger() == 1 else "False"
        except Exception:
            pass
        return str(param.AsInteger())
    if st == StorageType.Double:
        vs = param.AsValueString()
        return vs if vs else str(param.AsDouble())
    if st == StorageType.ElementId:
        return str(param.AsElementId().IntegerValue)
    return ""


# ─────────────────────────────────────────────────────────────────────────────
#  PARAMETER ITEM
# ─────────────────────────────────────────────────────────────────────────────

class ParameterItem(object):
    def __init__(self, key, scope, name, storage, is_shared, params_by_id, total_target_count):
        self.key = key
        self.scope = scope
        self.name = name
        self.storage = storage
        self.is_shared = is_shared
        self.params_by_id = params_by_id
        self.total_target_count = total_target_count

        self.current_value = self._compute_current_value(params_by_id)

        self.is_boolean = (storage == StorageType.Integer
                           and name.lower().startswith("is "))
        if self.is_boolean:
            self.control_type = "checkbox"
        elif storage == StorageType.ElementId:
            self.control_type = "elementid"
        else:
            self.control_type = "text"

        # Properties-style row fields (shown side by side in the UI list).
        self.ParamName = name
        self.ValueDisplay = self._build_value_display()
        self.Display = name

    def _compute_current_value(self, params_by_id):
        first = None
        has_value = False
        for p in params_by_id.values():
            v = _parameter_value_text(p)
            if not has_value:
                first = v
                has_value = True
            elif v != first:
                return "<varies>"
        return first if has_value else ""

    def _build_value_display(self):
        if self.current_value == "<varies>":
            return "<varies>"

        if self.is_boolean:
            raw = str(self.current_value).strip().lower()
            return "Yes" if raw in ("1", "true", "yes", "on") else "No"

        if self.storage == StorageType.ElementId:
            return _current_named_elementid_display(self.name, self.current_value, prefer_options=False)

        if self.current_value in (None, ""):
            return "<empty>"

        return str(self.current_value)


def _collect_parameter_items(target_elements):
    collected = {}
    total_count = len(target_elements)

    for element in target_elements:
        for param in element.Parameters:
            try:
                if param.IsReadOnly:
                    continue
                if str(param.StorageType) == "None":
                    continue
                p_name = param.Definition.Name
                p_id = param.Id.IntegerValue
                key = ("Instance", p_name, p_id)
                if key not in collected:
                    is_shared = False
                    try:
                        is_shared = param.IsShared
                    except Exception:
                        pass
                    collected[key] = {
                        "scope": "Instance",
                        "name": p_name,
                        "storage": param.StorageType,
                        "is_shared": is_shared,
                        "params_by_id": {},
                    }
                collected[key]["params_by_id"][element.Id.IntegerValue] = param
            except Exception:
                continue

    items = []
    for key, data in collected.items():
        items.append(ParameterItem(
            key=key,
            scope=data["scope"],
            name=data["name"],
            storage=data["storage"],
            is_shared=data["is_shared"],
            params_by_id=data["params_by_id"],
            total_target_count=total_count,
        ))
    return sorted(items, key=lambda i: (i.scope, i.name.lower()))


def _apply_parameter_value(param_item, new_value, allowed_element_ids=None):
    def _resolve_eid(param_obj, text_value):
        pname = ""
        try:
            pname = param_obj.Definition.Name
        except Exception:
            pass
        resolved = _resolve_named_elementid_from_text(pname, text_value)
        if resolved is not None:
            return resolved
        raw = (text_value or "").strip()
        raise ValueError("Could not resolve ElementId '{}' for '{}'".format(raw, pname))

    t = Transaction(doc, "FORMAT: Apply parameter")
    t.Start()
    success = errors = 0
    try:
        for elem_id, p in param_item.params_by_id.items():
            if allowed_element_ids is not None and elem_id not in allowed_element_ids:
                continue
            try:
                if param_item.storage == StorageType.String:
                    p.Set(new_value)
                elif param_item.storage == StorageType.Integer:
                    txt = new_value.strip().lower()
                    if txt in ("true", "yes", "1", "on"):
                        p.Set(1)
                    elif txt in ("false", "no", "0", "off"):
                        p.Set(0)
                    else:
                        p.Set(int(new_value))
                elif param_item.storage == StorageType.Double:
                    ok = p.SetValueString(new_value)
                    if not ok:
                        p.Set(float(new_value))
                elif param_item.storage == StorageType.ElementId:
                    p.Set(_resolve_eid(p, new_value))
                else:
                    errors += 1
                    continue
                success += 1
            except Exception:
                errors += 1
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return success, errors


# ─────────────────────────────────────────────────────────────────────────────
#  TARGET COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

def _collect_titleblocks_on_sheets(sheets):
    owner_ids = set(s.Id.IntegerValue for s in sheets)
    if not owner_ids:
        return []

    out = []
    for tb in (FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_TitleBlocks)
               .WhereElementIsNotElementType()
               .ToElements()):
        try:
            if tb.OwnerViewId and tb.OwnerViewId.IntegerValue in owner_ids:
                out.append(tb)
        except Exception:
            continue
    return out


def _collect_placed_views_on_sheets(sheets):
    view_elements = []
    seen = set()
    for sheet in sheets:
        try:
            placed = list(sheet.GetAllPlacedViews())
        except Exception:
            placed = []
        for vid in placed:
            if vid is None:
                continue
            key = vid.IntegerValue
            if key in seen:
                continue
            view = doc.GetElement(vid)
            if view:
                seen.add(key)
                view_elements.append(view)
    return view_elements


def _collect_titleblock_type_map():
    """Return dict[label] = ElementId for available title block types."""
    mapping = {}
    types = (FilteredElementCollector(doc)
             .OfCategory(BuiltInCategory.OST_TitleBlocks)
             .WhereElementIsElementType()
             .ToElements())

    for tb_type in types:
        try:
            type_name = str(getattr(tb_type, "Name", "") or "")
            family_name = ""
            try:
                family_name = str(getattr(tb_type, "FamilyName", "") or "")
            except Exception:
                family_name = ""

            base_label = "{} : {}".format(family_name, type_name) if family_name else type_name
            label = base_label
            idx = 2
            while label in mapping:
                label = "{} ({})".format(base_label, idx)
                idx += 1

            mapping[label] = tb_type.Id
        except Exception:
            continue

    return dict(sorted(mapping.items(), key=lambda kv: kv[0].lower()))


def _set_selection(element_ids):
    uidoc.Selection.SetElementIds(List[ElementId](element_ids))


# ─────────────────────────────────────────────────────────────────────────────
#  UNIFIED FORMAT WINDOW  (v2 single-window, 3-panel layout)
# ─────────────────────────────────────────────────────────────────────────────

class FORMATWindowV2(forms.WPFWindow):
    """
    Three-panel unified editor:
      Panel 1 (left)   – Target type selector + info counters
      Panel 2 (center) – Hierarchical sheet picker with search
      Panel 3 (right)  – Auto-populating parameter editor
    """

    _TARGET_SHEET      = "Sheet"
    _TARGET_TITLEBLOCK = "TitleBlock"
    _TARGET_VIEW       = "PlacedView"

    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)

        self._all_sheets               = []
        self._full_tree_roots          = []
        self._tree_roots               = []
        self._root_index               = {}
        self._last_clicked_sheet_node  = None
        self._target_type              = self._TARGET_SHEET
        self._targets                  = []
        self._param_items              = []
        self._building                 = False
        self._search_timer             = None
        self._selection_timer          = None
        self._pending_search_text      = ""
        self._pending_selection_refresh = False
        self._last_applied_search_text = None
        self._tb_type_map              = {}
        self._panel_ratio              = (1.1, 1.8, 1.1)
        self._is_custom_maximized      = False
        self._restore_bounds           = None

        self._apply_screen_constraints()

        self._load_all_sheets()
        self._build_tree()
        self._init_search_debounce()
        self._init_selection_debounce()
        self._apply_tree_to_ui()
        self._refresh_targets_and_params()
        self._update_window_buttons()
        self.Topmost = False
        self._init_revision_ui()

    def _get_current_screen_work_area(self):
        try:
            cursor = Control.MousePosition
            screen = System.Windows.Forms.Screen.FromPoint(cursor)
            if screen is not None:
                return screen.WorkingArea
        except Exception:
            pass

        try:
            return System.Windows.SystemParameters.WorkArea
        except Exception:
            return None

    def _apply_screen_constraints(self):
        wa = self._get_current_screen_work_area()
        if wa is None:
            return
        try:
            self.MaxHeight = wa.Height
            self.MaxWidth = wa.Width
        except Exception:
            pass

    def _store_restore_bounds(self):
        if self._is_custom_maximized:
            return
        try:
            self._restore_bounds = {
                "Left": self.Left,
                "Top": self.Top,
                "Width": self.Width,
                "Height": self.Height,
            }
        except Exception:
            self._restore_bounds = None

    def _restore_window_bounds(self):
        if not self._restore_bounds:
            return
        try:
            self.Left = self._restore_bounds["Left"]
            self.Top = self._restore_bounds["Top"]
            self.Width = self._restore_bounds["Width"]
            self.Height = self._restore_bounds["Height"]
        except Exception:
            pass

    def _maximize_to_current_monitor(self):
        self._store_restore_bounds()
        wa = self._get_current_screen_work_area()
        if wa is None:
            return
        try:
            self.WindowState = System.Windows.WindowState.Normal
        except Exception:
            pass
        try:
            self.Left = wa.Left
            self.Top = wa.Top
            self.Width = wa.Width
            self.Height = wa.Height
            self._is_custom_maximized = True
        except Exception:
            pass

    def _restore_from_custom_maximize(self):
        self._is_custom_maximized = False
        self._restore_window_bounds()

    def _init_revision_ui(self):
        try:
            if getattr(self, 'UI_rev_mode', None):
                self.UI_rev_mode.ItemsSource = [
                    'Next Interim Revision',
                    'Next Formal Submission',
                ]
                self.UI_rev_mode.SelectedIndex = 0
            if getattr(self, 'UI_rev_date', None) and not self.UI_rev_date.Text:
                self.UI_rev_date.Text = datetime.datetime.now().strftime('%d.%m.%y')
        except Exception:
            pass

    def _init_search_debounce(self):
        """Delay filter execution while user is typing to avoid UI freezes."""
        self._search_timer = System.Windows.Threading.DispatcherTimer()
        self._search_timer.Interval = System.TimeSpan.FromMilliseconds(250)
        self._search_timer.Tick += self._on_search_timer_tick

    def _init_selection_debounce(self):
        """Debounce heavy target/parameter refreshes after checkbox bursts."""
        self._selection_timer = System.Windows.Threading.DispatcherTimer()
        self._selection_timer.Interval = System.TimeSpan.FromMilliseconds(180)
        self._selection_timer.Tick += self._on_selection_timer_tick

    def _schedule_selection_refresh(self):
        self._pending_selection_refresh = True
        try:
            self._selection_timer.Stop()
            self._selection_timer.Start()
        except Exception:
            self._refresh_targets_and_params()

    def _on_selection_timer_tick(self, sender, e):
        try:
            self._selection_timer.Stop()
        except Exception:
            pass
        if self._pending_selection_refresh:
            self._pending_selection_refresh = False
            self._refresh_targets_and_params()

    def _on_search_timer_tick(self, sender, e):
        try:
            self._search_timer.Stop()
        except Exception:
            pass
        self._apply_search_filter(self._pending_search_text)

    def _apply_search_filter(self, raw_text):
        search_text = (raw_text or "").lower().strip()
        if search_text == self._last_applied_search_text:
            return

        if not search_text:
            self._tree_roots = self._full_tree_roots
        elif len(search_text) < 2:
            # Avoid expensive broad scans when user just started typing.
            self._tree_roots = self._full_tree_roots
        else:
            self._tree_roots = self._build_filtered_tree(self._full_tree_roots, search_text)

        self._sync_group_checkstates(self._tree_roots)
        self.UI_tree_sheets.ItemsSource = self._tree_roots
        self._refresh_sheet_counter()
        self.UI_tree_sheets.Items.Refresh()
        self._last_applied_search_text = search_text

        if search_text and len(search_text) < 2:
            self._set_status("Type at least 2 characters to filter sheets.")
        elif search_text:
            self._set_status("Sheet filter applied: '{}'".format(search_text))
        else:
            self._set_status("Search filter cleared.")

    # ── Sheet loading & tree ───────────────────────────────────────────────

    def _load_all_sheets(self):
        self._all_sheets = list(
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Sheets)
            .WhereElementIsNotElementType()
            .ToElements()
        )

    def _build_tree(self):
        self._full_tree_roots, self._root_index = _build_sheet_tree(self._all_sheets)
        self._tree_roots = self._full_tree_roots

    def _apply_tree_to_ui(self):
        self._sync_group_checkstates(self._full_tree_roots)
        self.UI_tree_sheets.ItemsSource = self._tree_roots
        self._refresh_sheet_counter()

    # ── Sheet counter ──────────────────────────────────────────────────────

    def _refresh_sheet_counter(self):
        selected = self._count_selected_sheets(self._full_tree_roots)
        total    = self._count_all_sheets(self._full_tree_roots)
        self.UI_selected_count.Text = "Selected: {}".format(selected)
        self.UI_sheet_count.Text    = "({} of {})".format(selected, total)

    def _count_selected_sheets(self, nodes):
        total = 0
        for node in nodes:
            if node.IsSheet:
                if node.IsChecked:
                    total += 1
            elif node.Children:
                total += self._count_selected_sheets(node.Children)
        return total

    def _count_all_sheets(self, nodes):
        total = 0
        for node in nodes:
            if node.IsSheet:
                total += 1
            elif node.Children:
                total += self._count_all_sheets(node.Children)
        return total

    # ── Selected sheet list ────────────────────────────────────────────────

    def _get_selected_sheets(self):
        selected = []
        def collect(nodes):
            for node in nodes:
                if node.IsSheet and node.Sheet and node.IsChecked:
                    selected.append(node.Sheet)
                if node.Children:
                    collect(node.Children)
        collect(self._full_tree_roots)
        return selected

    # ── Target collection & parameter refresh ─────────────────────────────

    def _refresh_targets_and_params(self):
        if self._building:
            return
        self._building = True
        try:
            sheets = self._get_selected_sheets()

            if self._target_type == self._TARGET_SHEET:
                self._targets = list(sheets)
            elif self._target_type == self._TARGET_TITLEBLOCK:
                self._targets = _collect_titleblocks_on_sheets(sheets)
            else:
                self._targets = _collect_placed_views_on_sheets(sheets)

            if self._targets:
                self._param_items = _collect_parameter_items(self._targets)
            else:
                self._param_items = []

            self._bind_param_list()
            self._update_info_counters(len(sheets))
            self._update_target_label()
            self._update_header_mode(len(sheets))
            self._update_titleblock_replace_ui()
            self._reset_inline_editor()
        finally:
            self._building = False

    def _update_header_mode(self, sheet_count):
        labels = {
            self._TARGET_SHEET: "Mode: Sheet Parameters",
            self._TARGET_TITLEBLOCK: "Mode: Title Block Parameters",
            self._TARGET_VIEW: "Mode: Placed View Parameters",
        }
        base = labels.get(self._target_type, "Mode")
        self.UI_header_mode.Text = "{} | Sheets: {} | Targets: {}".format(
            base, sheet_count, len(self._targets))

    def _set_status(self, text):
        self.UI_status_bar.Text = text

    def _capture_panel_ratio(self):
        try:
            widths = [
                float(self.UI_col_panel_1.ActualWidth),
                float(self.UI_col_panel_2.ActualWidth),
                float(self.UI_col_panel_3.ActualWidth),
            ]
            total = sum(w for w in widths if w > 0)
            if total > 0:
                self._panel_ratio = tuple(w / total for w in widths)
        except Exception:
            pass

    def _apply_panel_ratio(self):
        try:
            r1, r2, r3 = self._panel_ratio
            total = r1 + r2 + r3
            if total <= 0:
                return
            self.UI_col_panel_1.Width = System.Windows.GridLength(r1 / total, System.Windows.GridUnitType.Star)
            self.UI_col_panel_2.Width = System.Windows.GridLength(r2 / total, System.Windows.GridUnitType.Star)
            self.UI_col_panel_3.Width = System.Windows.GridLength(r3 / total, System.Windows.GridUnitType.Star)
        except Exception:
            pass

    def _update_titleblock_replace_ui(self):
        combo = getattr(self, "UI_tb_type_combo", None)
        button = getattr(self, "UI_btn_replace_tb", None)
        if combo is None or button is None:
            return

        is_tb_mode = (self._target_type == self._TARGET_TITLEBLOCK)
        has_targets = bool(self._targets)

        combo.IsEnabled = is_tb_mode and has_targets
        button.IsEnabled = is_tb_mode and has_targets

        if not self._tb_type_map:
            self._tb_type_map = _collect_titleblock_type_map()

        names = list(self._tb_type_map.keys())
        combo.ItemsSource = names

        if not names:
            combo.SelectedIndex = -1
            return

        selected_label = None
        if has_targets:
            type_ids = set()
            for tb in self._targets:
                try:
                    type_ids.add(tb.GetTypeId().IntegerValue)
                except Exception:
                    continue
            if len(type_ids) == 1:
                current_id = next(iter(type_ids))
                for label, eid in self._tb_type_map.items():
                    try:
                        if eid.IntegerValue == current_id:
                            selected_label = label
                            break
                    except Exception:
                        continue

        if selected_label and selected_label in names:
            combo.SelectedItem = selected_label
        elif combo.SelectedItem not in names:
            combo.SelectedIndex = 0

    # ── Smart automation helpers (V3) ─────────────────────────────────────

    def _get_single_titleblock_on_sheet(self, sheet):
        elems = list(
            FilteredElementCollector(doc, sheet.Id)
            .OfCategory(BuiltInCategory.OST_TitleBlocks)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        return elems[0] if len(elems) == 1 else None

    def _extract_sheet_tokens(self, sheet):
        number = (sheet.SheetNumber or "").strip()
        name = (sheet.Name or "").strip()
        up_name = name.upper()

        # Extract basic tokens from naming conventions.
        segments = [seg.strip() for seg in number.split("-") if seg and seg.strip()]

        series = ""
        for seg in segments:
            if re.match(r"^\d{3}$", seg):
                series = seg
                break

        discipline = ""
        for seg in segments:
            if re.match(r"^[A-Z]{2,4}$", seg):
                discipline = seg

        zone = ""
        m_zone = re.search(r"ZONE\s*([0-9]{2,4})", up_name)
        if m_zone:
            zone = m_zone.group(1).zfill(3)
        elif segments and re.match(r"^\d{3,4}$", segments[-1]):
            zone = segments[-1].zfill(3)

        level_code = ""
        m_lvl_name = re.search(r"LEVEL\s*([0-9]+)", up_name)
        if m_lvl_name:
            level_code = m_lvl_name.group(1).zfill(2)
        elif "LOWER GROUND" in up_name:
            level_code = "B1"
        elif "GROUND LEVEL" in up_name or "GROUND" in up_name:
            level_code = "GR"
        else:
            for seg in segments:
                if re.match(r"^\d{2}$", seg):
                    level_code = seg
                    break
                if seg.upper() in ("GR", "B1", "B2"):
                    level_code = seg.upper()
                    break

        name_prefix = name.split("-")[0].strip() if "-" in name else name

        return {
            "sheet_number": number,
            "sheet_name": name,
            "series": series,
            "discipline": discipline,
            "zone": zone,
            "level_code": level_code,
            "name_prefix": name_prefix,
        }

    def _param_has_nonempty_value(self, param):
        try:
            st = param.StorageType
            if st == StorageType.String:
                return bool((param.AsString() or "").strip())
            if st == StorageType.Integer:
                return param.AsInteger() != 0
            if st == StorageType.Double:
                return abs(param.AsDouble()) > 1e-9
            if st == StorageType.ElementId:
                eid = param.AsElementId()
                return eid is not None and eid.IntegerValue > 0
        except Exception:
            pass
        return False

    def _set_parameter_value_safe(self, param, value_text):
        try:
            st = param.StorageType
            raw = "" if value_text is None else str(value_text).strip()

            if st == StorageType.String:
                return bool(param.Set(raw))

            if st == StorageType.Integer:
                low = raw.lower()
                if low in ("true", "yes", "on"):
                    return bool(param.Set(1))
                if low in ("false", "no", "off"):
                    return bool(param.Set(0))
                return bool(param.Set(int(raw) if raw else 0))

            if st == StorageType.Double:
                if not raw:
                    return bool(param.Set(0.0))
                ok = param.SetValueString(raw)
                if not ok:
                    return bool(param.Set(float(raw)))
                return True

            if st == StorageType.ElementId:
                pname = ""
                try:
                    pname = param.Definition.Name
                except Exception:
                    pass
                resolved = _resolve_named_elementid_from_text(pname, raw)
                if resolved is None:
                    return False
                return bool(param.Set(resolved))
        except Exception:
            return False
        return False

    def _set_param_on_element(self, element, candidate_names, value_text, fill_blanks_only=True):
        if element is None or value_text in (None, ""):
            return False

        for pname in candidate_names:
            param = None
            try:
                param = element.LookupParameter(pname)
            except Exception:
                param = None
            if param is None or param.IsReadOnly:
                continue

            if fill_blanks_only and self._param_has_nonempty_value(param):
                continue

            if self._set_parameter_value_safe(param, value_text):
                return True

        return False

    def _get_primary_view_scale_on_sheet(self, sheet):
        try:
            for vpid in list(sheet.GetAllViewports()):
                vp = doc.GetElement(vpid)
                if vp is None:
                    continue
                view = doc.GetElement(vp.ViewId)
                if view is None:
                    continue
                scale_val = getattr(view, "Scale", None)
                if scale_val:
                    try:
                        return int(scale_val)
                    except Exception:
                        pass
        except Exception:
            pass
        return None

    def _auto_toggle_keyplan_for_titleblock(self, titleblock, zone_code, scale_int):
        if titleblock is None:
            return 0, False

        zone = (zone_code or "").strip()
        zone_tokens = []
        if zone:
            zone_tokens.append(zone)
            try:
                zone_tokens.append(str(int(zone)))
            except Exception:
                pass
            if zone.isdigit() and len(zone) < 3:
                zone_tokens.append(zone.zfill(3))

        scale_tokens = []
        if scale_int:
            s = str(scale_int)
            scale_tokens = [s, "1:{}".format(s), "1/{}".format(s), "1-{}".format(s)]

        def _extract_zone_from_name(name_low):
            m = re.search(r"zone\s*([0-9]{2,4})", name_low)
            if m:
                z = m.group(1)
                return z.zfill(3) if z.isdigit() else z
            return ""

        def _extract_scale_from_name(name_low):
            m = re.search(r"scale\s*1\s*[-:/]\s*([0-9]{2,4})", name_low)
            if not m:
                m = re.search(r"\b1\s*[-:/]\s*([0-9]{2,4})\b", name_low)
            if m:
                try:
                    return int(m.group(1))
                except Exception:
                    return None
            return None

        candidates = []
        try:
            for p in titleblock.Parameters:
                try:
                    if p.IsReadOnly or p.StorageType != StorageType.Integer:
                        continue
                    pname = p.Definition.Name if p.Definition else ""
                    low = (pname or "").strip().lower()
                    if not low:
                        continue
                    has_keyplan_word = ("keyplan" in low or "key plan" in low or "kp" in low)
                    has_scale_zone_words = ("scale" in low and "zone" in low)
                    if not has_keyplan_word and not has_scale_zone_words:
                        continue
                    candidates.append((p, low))
                except Exception:
                    continue
        except Exception:
            return 0

        if not candidates:
            return 0, False

        scored = []
        for p, low in candidates:
            parsed_zone = _extract_zone_from_name(low)
            parsed_scale = _extract_scale_from_name(low)

            zone_hit = False
            for z in zone_tokens:
                zlow = z.lower()
                if zlow and re.search(r"(^|[^0-9]){}([^0-9]|$)".format(re.escape(zlow)), low):
                    zone_hit = True
                    break

            scale_hit = False
            for st in scale_tokens:
                if st.lower() in low:
                    scale_hit = True
                    break

            score = 0
            if "keyplan" in low or "key plan" in low:
                score += 1

            if zone_tokens and parsed_zone and parsed_zone in [z.zfill(3) if z.isdigit() else z for z in zone_tokens if z]:
                score += 6
            if zone_hit:
                score += 3

            if scale_int and parsed_scale == int(scale_int):
                score += 5
            if scale_hit:
                score += 2

            scored.append((score, p, low))

        best_score = max([s for s, _, _ in scored]) if scored else 0
        if best_score <= 0:
            return 0, False

        writes = 0
        for score, p, _ in scored:
            target_val = 1 if score == best_score else 0
            try:
                if p.AsInteger() != target_val:
                    p.Set(target_val)
                    writes += 1
            except Exception:
                continue

        return writes, True

    def _set_titleblock_keyplan_visibility(self, titleblock, is_visible):
        val = "1" if is_visible else "0"
        return self._set_param_on_element(
            titleblock,
            [
                "SHOW TITLEBLOCK KEY PLAN",
                "SHOW TITLEBLOCK KEYPLAN",
                "SHOW KEY PLAN",
                "SHOW KEYPLAN",
            ],
            val,
            fill_blanks_only=False,
        )

    def _get_titleblock_anchor_center(self, sheet, titleblock):
        tb_box = titleblock.get_BoundingBox(sheet)
        if tb_box is None:
            return None

        def _param_double_contains(tokens):
            wanted = [t.lower() for t in tokens]
            try:
                for p in titleblock.Parameters:
                    try:
                        if p.IsReadOnly or p.StorageType != StorageType.Double:
                            continue
                        pname = (p.Definition.Name if p.Definition else "") or ""
                        low = pname.strip().lower()
                        if all(tok in low for tok in wanted):
                            return abs(p.AsDouble())
                    except Exception:
                        continue
            except Exception:
                pass
            return 0.0

        # Prefer explicit drawing-field geometry inferred from family parameters.
        perimeter = max(
            _param_double_contains(["perimeter", "boundary"]),
            _param_double_contains(["titleblock", "perimeter", "boundary"]),
        )
        ribbon_width = max(
            _param_double_contains(["ribbon", "width"]),
            _param_double_contains(["titleblock", "ribbon", "width"]),
        )
        ribbon_offset = max(
            _param_double_contains(["ribbon", "line", "offset"]),
            _param_double_contains(["titleblock", "ribbon", "line", "offset"]),
        )

        usable_min_x = tb_box.Min.X + perimeter
        usable_max_x = tb_box.Max.X - perimeter
        usable_min_y = tb_box.Min.Y + perimeter
        usable_max_y = tb_box.Max.Y - perimeter

        if ribbon_width > 0:
            usable_max_x -= ribbon_width
        if ribbon_offset > 0:
            usable_max_x -= ribbon_offset

        if usable_max_x <= usable_min_x or usable_max_y <= usable_min_y:
            return XYZ((tb_box.Min.X + tb_box.Max.X) * 0.5,
                       (tb_box.Min.Y + tb_box.Max.Y) * 0.5,
                       0)

        return XYZ((usable_min_x + usable_max_x) * 0.5,
                   (usable_min_y + usable_max_y) * 0.5,
                   0)

    def _match_scopebox_name_for_zone(self, zone_code):
        if not zone_code:
            return None

        options = _collect_named_elementid_options("scope box")
        if not options:
            return None

        # Strong match: exact zone token boundary.
        pattern = r"(^|[^0-9]){}([^0-9]|$)".format(re.escape(zone_code))
        for name, _ in options:
            if re.search(pattern, name):
                return name

        # Fallback: simple containment.
        for name, _ in options:
            if zone_code in name:
                return name

        return None

    def _parse_revision_code(self, text):
        raw = (text or '').strip().upper()
        match = re.match(r'^([A-Z])\.(\d{2})$', raw)
        if not match:
            return None
        return match.group(1), int(match.group(2))

    def _format_revision_code(self, letter_code, number_code):
        return '{}.{:02d}'.format(letter_code, int(number_code))

    def _next_revision_code(self, current_code, mode_name):
        parsed = self._parse_revision_code(current_code)
        if not parsed:
            if 'formal' in (mode_name or '').lower():
                return 'A.00'
            return 'A.01'

        letter_code, number_code = parsed
        if 'formal' in (mode_name or '').lower():
            next_letter = chr(min(ord('Z'), ord(letter_code) + 1))
            return self._format_revision_code(next_letter, 0)
        return self._format_revision_code(letter_code, number_code + 1)

    def _get_sheet_revision_ids(self, sheet):
        rev_ids = []
        seen = set()
        try:
            current_id = sheet.GetCurrentRevision()
            if current_id and current_id.IntegerValue > 0:
                seen.add(current_id.IntegerValue)
                rev_ids.append(current_id)
        except Exception:
            pass
        try:
            for rid in list(sheet.GetAdditionalRevisionIds()):
                if rid and rid.IntegerValue > 0 and rid.IntegerValue not in seen:
                    seen.add(rid.IntegerValue)
                    rev_ids.append(rid)
        except Exception:
            pass
        return rev_ids

    def _get_latest_revision_code_on_sheet(self, sheet):
        best_seq = -1
        best_code = ''
        for rid in self._get_sheet_revision_ids(sheet):
            rev = doc.GetElement(rid)
            if rev is None:
                continue
            try:
                seq = int(rev.SequenceNumber)
            except Exception:
                seq = -1
            try:
                code = sheet.GetRevisionNumberOnSheet(rid) or ''
            except Exception:
                code = ''
            if code and seq >= best_seq:
                best_seq = seq
                best_code = code
        return best_code

    def _find_revision_by_code(self, revision_code):
        for rid in Revision.GetAllRevisionIds(doc):
            rev = doc.GetElement(rid)
            if rev is None:
                continue
            try:
                if (rev.RevisionNumber or '').strip().upper() == revision_code.upper():
                    return rev
            except Exception:
                continue
        return None

    def _get_or_create_revision(self, revision_code, description, revision_date):
        rev = self._find_revision_by_code(revision_code)
        created = False
        if rev is None:
            rev = Revision.Create(doc)
            created = True

        try:
            rev.NumberType = RevisionNumberType.Alphanumeric
        except Exception:
            pass

        try:
            rev.RevisionNumber = revision_code
        except Exception:
            pass

        try:
            if created or not rev.Issued:
                rev.Description = description
                rev.RevisionDate = revision_date
        except Exception:
            pass

        return rev, created

    def _add_revision_to_sheet_if_missing(self, sheet, revision_id):
        try:
            current = sheet.GetAdditionalRevisionIds()
            for rid in current:
                if rid.IntegerValue == revision_id.IntegerValue:
                    return False
            current.Add(revision_id)
            sheet.SetAdditionalRevisionIds(current)
            return True
        except Exception:
            return False

    def _smart_autofill_selected_sheets(self):
        sheets = self._get_selected_sheets()
        if not sheets:
            forms.alert("Select at least one sheet first.", title=__title__)
            return

        fill_blanks_only = bool(getattr(self, "UI_chk_fill_blanks", None) and self.UI_chk_fill_blanks.IsChecked)
        apply_views = bool(getattr(self, "UI_chk_apply_view_scope", None) and self.UI_chk_apply_view_scope.IsChecked)
        apply_keyplan = bool(getattr(self, "UI_chk_keyplan_auto", None) and self.UI_chk_keyplan_auto.IsChecked)

        writes = 0
        processed = 0

        tx = Transaction(doc, "FORMAT V3: Smart Auto Fill")
        tx.Start()
        try:
            for sheet in sheets:
                processed += 1
                tokens = self._extract_sheet_tokens(sheet)

                writes += 1 if self._set_param_on_element(sheet, ["SHEET NAME PREFIX", "Sheet Name Prefix"], tokens.get("name_prefix"), fill_blanks_only) else 0
                writes += 1 if self._set_param_on_element(sheet, ["DRAWING REGISTER SERIES", "Sheet Series", "SHEET SERIES"], tokens.get("series"), fill_blanks_only) else 0
                writes += 1 if self._set_param_on_element(sheet, ["MHT_Discipline", "MHT_Dicipline", "Discipline"], tokens.get("discipline"), fill_blanks_only) else 0
                writes += 1 if self._set_param_on_element(sheet, ["Zone", "ZONE", "Vic_Zone"], tokens.get("zone"), fill_blanks_only) else 0
                writes += 1 if self._set_param_on_element(sheet, ["LEVEL", "Level"], tokens.get("level_code"), fill_blanks_only) else 0

                titleblock = self._get_single_titleblock_on_sheet(sheet)
                if titleblock is not None:
                    writes += 1 if self._set_param_on_element(titleblock, ["LEVEL", "Level"], tokens.get("level_code"), fill_blanks_only) else 0
                    writes += 1 if self._set_param_on_element(titleblock, ["Zone", "ZONE", "Vic_Zone"], tokens.get("zone"), fill_blanks_only) else 0
                    writes += 1 if self._set_param_on_element(titleblock, ["DRAWING REGISTER SERIES", "SHEET SERIES", "Sheet Series"], tokens.get("series"), fill_blanks_only) else 0
                    if apply_keyplan:
                        kp_writes, kp_found = self._auto_toggle_keyplan_for_titleblock(
                            titleblock,
                            tokens.get("zone"),
                            self._get_primary_view_scale_on_sheet(sheet),
                        )
                        writes += kp_writes
                        if kp_found:
                            writes += 1 if self._set_titleblock_keyplan_visibility(titleblock, True) else 0

                if apply_views:
                    scope_name = self._match_scopebox_name_for_zone(tokens.get("zone"))
                    for vpid in list(sheet.GetAllViewports()):
                        vp = doc.GetElement(vpid)
                        if vp is None:
                            continue
                        view = doc.GetElement(vp.ViewId)
                        if view is None:
                            continue

                        if scope_name:
                            writes += 1 if self._set_param_on_element(
                                view,
                                ["Scope Box", "Volume of Interest", "ScopeBox"],
                                scope_name,
                                False,
                            ) else 0

                        writes += 1 if self._set_param_on_element(view, ["LEVEL", "Level"], tokens.get("level_code"), fill_blanks_only) else 0
                        writes += 1 if self._set_param_on_element(vp, ["Zone", "ZONE", "Vic_Zone"], tokens.get("zone"), fill_blanks_only) else 0

            tx.Commit()
        except Exception as ex:
            tx.RollBack()
            forms.alert("Smart Auto Fill failed: {}".format(ex), title=__title__)
            return

        msg = "Smart Auto Fill complete. Sheets: {} | Parameter writes: {}".format(processed, writes)
        self.UI_feedback.Text = msg
        self._set_status(msg)

    def _smart_align_viewports_selected_sheets(self):
        sheets = self._get_selected_sheets()
        if not sheets:
            forms.alert("Select at least one sheet first.", title=__title__)
            return

        moved = 0
        skipped = 0

        tx = Transaction(doc, "FORMAT V3: Smart Align Viewports")
        tx.Start()
        try:
            for sheet in sheets:
                tb = self._get_single_titleblock_on_sheet(sheet)
                if tb is None:
                    skipped += 1
                    continue

                target_center = self._get_titleblock_anchor_center(sheet, tb)
                if target_center is None:
                    skipped += 1
                    continue

                vp_ids = list(sheet.GetAllViewports())
                if not vp_ids:
                    skipped += 1
                    continue

                viewports = [doc.GetElement(vpid) for vpid in vp_ids if doc.GetElement(vpid)]
                if not viewports:
                    skipped += 1
                    continue

                # Use the largest viewport as the alignment anchor and move the
                # rest by the same delta. This is more stable when auxiliary
                # views use different scope boxes or crop extents.
                primary = None
                primary_center = None
                primary_area = -1.0
                for vp in viewports:
                    try:
                        outline = vp.GetBoxOutline()
                        min_pt = outline.MinimumPoint
                        max_pt = outline.MaximumPoint
                        area = abs(max_pt.X - min_pt.X) * abs(max_pt.Y - min_pt.Y)
                        if area > primary_area:
                            primary_area = area
                            primary = vp
                            primary_center = vp.GetBoxCenter()
                    except Exception:
                        continue

                if primary is None or primary_center is None:
                    skipped += 1
                    continue

                delta = XYZ(target_center.X - primary_center.X, target_center.Y - primary_center.Y, 0)

                for vp in viewports:
                    try:
                        new_center = vp.GetBoxCenter().Add(delta)
                        vp.SetBoxCenter(new_center)
                        moved += 1
                    except Exception:
                        pass

            tx.Commit()
        except Exception as ex:
            tx.RollBack()
            forms.alert("Smart Align Viewports failed: {}".format(ex), title=__title__)
            return

        msg = "Smart Align complete. Viewports moved: {} | Sheets skipped: {}".format(moved, skipped)
        self.UI_feedback.Text = msg
        self._set_status(msg)

    def _update_window_buttons(self):
        try:
            if self._is_custom_maximized or self.WindowState == System.Windows.WindowState.Maximized:
                self.UI_btn_maximize.Content = "❐"
                self.UI_btn_maximize.ToolTip = "Restore"
            else:
                self.UI_btn_maximize.Content = "□"
                self.UI_btn_maximize.ToolTip = "Maximize"
        except Exception:
            pass

    def _update_info_counters(self, sheet_count):
        self.UI_info_sheets.Text  = "Sheets: {} selected".format(sheet_count)
        self.UI_info_targets.Text = "Elements: {}".format(len(self._targets))
        self.UI_info_params.Text  = "Params: {} editable".format(len(self._param_items))
        self.UI_tb_count.Text     = "Targets ({}): {}".format(
            self._target_type, len(self._targets))
        self.UI_param_count.Text  = "Editable Params: {}".format(len(self._param_items))

    def _update_target_label(self):
        labels = {
            self._TARGET_SHEET:      "— Sheet Parameters",
            self._TARGET_TITLEBLOCK: "— Title Block Parameters",
            self._TARGET_VIEW:       "— Placed View Parameters",
        }
        self.UI_target_label.Text = labels.get(self._target_type, "")

    def _bind_param_list(self):
        self.UI_param_list.ItemsSource = self._param_items

    # ── Inline editor ──────────────────────────────────────────────────────

    def _reset_inline_editor(self, label="Select a parameter..."):
        self.UI_param_name.Text = label
        self.UI_param_checkbox.Visibility    = System.Windows.Visibility.Collapsed
        self.UI_param_checkbox.IsChecked     = False
        self.UI_param_input.Visibility       = System.Windows.Visibility.Collapsed
        self.UI_param_input.Text             = ""
        self.UI_param_elementid.Visibility   = System.Windows.Visibility.Collapsed
        self.UI_param_elementid.ItemsSource  = []
        self.UI_param_elementid.SelectedIndex = -1

    def _show_inline_editor(self, item):
        self.UI_param_name.Text = item.name
        self.UI_param_checkbox.Visibility   = System.Windows.Visibility.Collapsed
        self.UI_param_input.Visibility      = System.Windows.Visibility.Collapsed
        self.UI_param_elementid.Visibility  = System.Windows.Visibility.Collapsed

        if item.is_boolean:
            self.UI_param_checkbox.Visibility = System.Windows.Visibility.Visible
            self.UI_param_checkbox.IsChecked  = self._is_checked_value(item.current_value)

        elif item.storage == StorageType.ElementId:
            options = _build_named_elementid_display_options(item.name)
            current_display = _current_named_elementid_display(item.name, item.current_value, prefer_options=True)
            if len(options) > 1:
                self.UI_param_elementid.ItemsSource = options
                self.UI_param_elementid.Visibility  = System.Windows.Visibility.Visible
                try:
                    self.UI_param_elementid.SelectedItem = (
                        current_display if current_display in options else "<None>"
                    )
                except Exception:
                    self.UI_param_elementid.SelectedIndex = 0
            else:
                self.UI_param_input.Visibility = System.Windows.Visibility.Visible
                self.UI_param_input.Text = current_display
        else:
            self.UI_param_input.Visibility = System.Windows.Visibility.Visible
            self.UI_param_input.Text = str(item.current_value) if item.current_value else ""

    def _is_checked_value(self, value_text):
        text = "" if value_text is None else str(value_text).strip().lower()
        if text in ("1", "true", "yes", "on"):
            return True
        try:
            return int(text) != 0
        except Exception:
            return False

    def _read_editor_value(self, item):
        if item.is_boolean:
            return "1" if self.UI_param_checkbox.IsChecked else "0"
        if (item.storage == StorageType.ElementId
                and self.UI_param_elementid.Visibility == System.Windows.Visibility.Visible):
            selected = self.UI_param_elementid.SelectedItem
            if selected is None:
                raise ValueError("Select a value for '{}' first.".format(item.name))
            return str(selected)
        return self.UI_param_input.Text

    # ── Tree group-state helpers ───────────────────────────────────────────

    def _sync_group_checkstates(self, nodes):
        for node in nodes:
            if node.IsSheet:
                continue
            if node.Children:
                self._sync_group_checkstates(node.Children)
            descendant_sheets = self._collect_descendant_sheets(node)
            if not descendant_sheets:
                node.GroupCheckState = False
                continue
            checked = sum(1 for s in descendant_sheets if s.IsChecked)
            if checked == 0:
                node.GroupCheckState = False
            elif checked == len(descendant_sheets):
                node.GroupCheckState = True
            else:
                node.GroupCheckState = None

    def _collect_descendant_sheets(self, node):
        sheets = []
        def collect(children):
            for child in children:
                if child.IsSheet:
                    sheets.append(child)
                elif child.Children:
                    collect(child.Children)
        collect(node.Children)
        return sheets

    def _get_visible_sheet_nodes(self):
        visible = []
        def collect(nodes):
            for n in nodes:
                if n.IsSheet:
                    visible.append(n)
                elif n.Children:
                    collect(n.Children)
        collect(self._tree_roots)
        return visible

    def _set_all_checked(self, nodes, value):
        for node in nodes:
            if node.IsSheet:
                node.IsChecked = value
            if node.Children:
                self._set_all_checked(node.Children, value)

    def _set_expanded_for_groups(self, nodes, is_expanded):
        for node in nodes:
            if not node.IsSheet:
                node.IsExpanded = is_expanded
                if node.Children:
                    self._set_expanded_for_groups(node.Children, is_expanded)

    def _set_checked_for_group(self, group_node, value):
        if not group_node or group_node.IsSheet:
            return

        target_ids = set()

        def collect_ids(node):
            for child in node.Children:
                if child.IsSheet and child.Sheet:
                    try:
                        target_ids.add(child.Sheet.Id.IntegerValue)
                    except Exception:
                        continue
                elif child.Children:
                    collect_ids(child)

        collect_ids(group_node)
        if not target_ids:
            return

        def apply_to_nodes(nodes):
            for n in nodes:
                if n.IsSheet and n.Sheet:
                    try:
                        if n.Sheet.Id.IntegerValue in target_ids:
                            n.IsChecked = value
                    except Exception:
                        continue
                elif n.Children:
                    apply_to_nodes(n.Children)

        apply_to_nodes(self._full_tree_roots)

    def _set_checked_for_range(self, start_node, end_node, value):
        visible = self._get_visible_sheet_nodes()
        if not visible:
            return
        try:
            si = visible.index(start_node)
            ei = visible.index(end_node)
        except ValueError:
            end_node.IsChecked = value
            return
        if si > ei:
            si, ei = ei, si
        for idx in range(si, ei + 1):
            visible[idx].IsChecked = value

    def _normalize_checkbox_value(self, value):
        return True if value is None else bool(value)

    # ── Filtered tree helpers ──────────────────────────────────────────────

    def _clone_group_subtree(self, group_node):
        clone = SheetTreeNode(group_node.Name, is_sheet=False)
        clone.IsExpanded = True
        for child in group_node.Children:
            if child.IsSheet:
                clone.Children.append(child)
            else:
                clone.Children.append(self._clone_group_subtree(child))
        return clone

    def _build_filtered_tree(self, nodes, search_text):
        filtered = []
        for node in nodes:
            if node.IsSheet:
                label = getattr(node, "SheetSearchText", node.NameLower)
                if search_text in label:
                    filtered.append(node)
            else:
                if search_text in node.NameLower:
                    # Reuse existing node when the whole group matches to reduce cloning.
                    filtered.append(node)
                else:
                    child_matches = self._build_filtered_tree(node.Children, search_text)
                    if child_matches:
                        g = SheetTreeNode(node.Name, is_sheet=False)
                        g.IsExpanded = True
                        g.Children = child_matches
                        filtered.append(g)
        return filtered

    # ── EVENT HANDLERS ─────────────────────────────────────────────────────

    def header_drag(self, sender, e):
        try:
            original = getattr(e, 'OriginalSource', None)
            if isinstance(original, System.Windows.Controls.Button):
                return
        except Exception:
            pass

        try:
            click_count = getattr(e, 'ClickCount', 1)
        except Exception:
            click_count = 1

        if click_count >= 2:
            self.button_maximize(sender, e)
            return

        try:
            if self._is_custom_maximized:
                mouse = Control.MousePosition
                restore = self._restore_bounds or {
                    "Width": max(self.MinWidth, self.Width * 0.75),
                    "Height": max(self.MinHeight, self.Height * 0.75),
                    "Left": self.Left,
                    "Top": self.Top,
                }
                new_width = restore.get("Width", self.Width)
                new_height = restore.get("Height", self.Height)
                self._restore_from_custom_maximize()
                self.Left = mouse.X - (new_width * 0.5)
                self.Top = max(0, mouse.Y - 12)
                self.Width = new_width
                self.Height = new_height
        except Exception:
            pass

        try:
            self.DragMove()
        except Exception:
            pass

    def button_minimize(self, sender, e):
        self.WindowState = System.Windows.WindowState.Minimized

    def button_maximize(self, sender, e):
        self._capture_panel_ratio()
        if self._is_custom_maximized or self.WindowState == System.Windows.WindowState.Maximized:
            self.WindowState = System.Windows.WindowState.Normal
            self._restore_from_custom_maximize()
        else:
            self._maximize_to_current_monitor()
        try:
            self.UpdateLayout()
        except Exception:
            pass
        self._apply_panel_ratio()
        self._update_window_buttons()

    def button_close(self, sender, e):
        self.Close()

    def Hyperlink_RequestNavigate(self, sender, e):
        Start(e.Uri.AbsoluteUri)

    # Panel 1
    def target_radio_changed(self, sender, e):
        if self.UI_rb_sheet.IsChecked:
            self._target_type = self._TARGET_SHEET
        elif self.UI_rb_titleblock.IsChecked:
            self._target_type = self._TARGET_TITLEBLOCK
        else:
            self._target_type = self._TARGET_VIEW
        self._refresh_targets_and_params()

    def button_refresh_targets(self, sender, e):
        self._refresh_targets_and_params()
        self._set_status("Targets and parameters refreshed.")

    def button_smart_autofill(self, sender, e):
        self._smart_autofill_selected_sheets()

    def button_smart_align_viewports(self, sender, e):
        self._smart_align_viewports_selected_sheets()

    def button_replace_titleblocks(self, sender, e):
        if self._target_type != self._TARGET_TITLEBLOCK:
            forms.alert("Switch target mode to Title Block Params first.", title=__title__)
            return

        sheets = self._get_selected_sheets()
        if not sheets:
            forms.alert("Select at least one sheet first.", title=__title__)
            return

        selected_label = None
        try:
            selected_label = self.UI_tb_type_combo.SelectedItem
        except Exception:
            selected_label = None

        if not selected_label:
            forms.alert("Select a title block type from the dropdown first.", title=__title__)
            return

        new_type_id = self._tb_type_map.get(str(selected_label))
        if new_type_id is None:
            forms.alert("Selected title block type could not be resolved.", title=__title__)
            return

        targets = _collect_titleblocks_on_sheets(sheets)
        if not targets:
            forms.alert("No title blocks found on selected sheets.", title=__title__)
            return

        changed = 0
        failed = 0

        tx = Transaction(doc, "FORMAT: Replace title blocks")
        tx.Start()
        try:
            for tb in targets:
                try:
                    cur = tb.GetTypeId()
                    if cur and cur.IntegerValue == new_type_id.IntegerValue:
                        continue
                    tb.ChangeTypeId(new_type_id)
                    changed += 1
                except Exception:
                    failed += 1
            tx.Commit()
        except Exception as ex:
            tx.RollBack()
            forms.alert("Replace title blocks failed: {}".format(ex), title=__title__)
            return

        self._refresh_targets_and_params()
        msg = "Title block replace complete. Changed: {} | Failed: {}".format(changed, failed)
        self.UI_feedback.Text = msg
        self._set_status(msg)

    def button_auto_revision_update(self, sender, e):
        sheets = self._get_selected_sheets()
        if not sheets:
            forms.alert('Select at least one sheet first.', title=__title__)
            return

        description = ''
        revision_date = ''
        mode_name = 'Next Interim Revision'
        try:
            description = (self.UI_rev_description.Text or '').strip()
        except Exception:
            description = ''
        try:
            revision_date = (self.UI_rev_date.Text or '').strip()
        except Exception:
            revision_date = ''
        try:
            mode_name = str(self.UI_rev_mode.SelectedItem or mode_name)
        except Exception:
            pass

        if not description:
            forms.alert('Enter a revision description first.', title=__title__)
            return
        if not revision_date:
            forms.alert('Enter a revision date first.', title=__title__)
            return

        created_count = 0
        added_count = 0
        skipped_count = 0

        tx = Transaction(doc, 'FORMAT: Auto Revision Update')
        tx.Start()
        try:
            for sheet in sheets:
                current_code = self._get_latest_revision_code_on_sheet(sheet)
                next_code = self._next_revision_code(current_code, mode_name)
                revision, created = self._get_or_create_revision(next_code, description, revision_date)
                if created:
                    created_count += 1
                if self._add_revision_to_sheet_if_missing(sheet, revision.Id):
                    added_count += 1
                else:
                    skipped_count += 1
            tx.Commit()
        except Exception as ex:
            tx.RollBack()
            forms.alert('Auto Revision Update failed: {}'.format(ex), title=__title__)
            return

        msg = 'Auto Revision Update complete. Sheets updated: {} | Revisions created: {} | Already assigned: {}'.format(
            added_count, created_count, skipped_count)
        self.UI_feedback.Text = msg
        self._set_status(msg)

    def button_reset_layout(self, sender, e):
        self._panel_ratio = (1.1, 1.8, 1.1)
        self._apply_panel_ratio()
        self._set_status("Panel layout reset: side panels equal, center panel prioritized.")

    # Panel 2
    def search_textbox_changed(self, sender, e):
        self._pending_search_text = self.UI_search.Text if self.UI_search else ""
        try:
            self._search_timer.Stop()
            self._search_timer.Start()
        except Exception:
            # Fallback if timer fails: apply immediately.
            self._apply_search_filter(self._pending_search_text)

    def button_clear_search(self, sender, e):
        self.UI_search.Text = ""
        self.UI_search.Focus()
        self._pending_search_text = ""
        try:
            self._search_timer.Stop()
        except Exception:
            pass
        self._apply_search_filter("")

    def sheet_checkbox_click(self, sender, e):
        node = getattr(sender, "Tag", None)
        if not node:
            return
        target_value = self._normalize_checkbox_value(
            getattr(sender, "IsChecked", False))
        modifiers = Control.ModifierKeys
        is_shift  = (modifiers & Keys.Shift) == Keys.Shift
        is_alt    = (modifiers & Keys.Alt)   == Keys.Alt

        if node.IsSheet:
            if is_alt:
                for n in self._get_visible_sheet_nodes():
                    n.IsChecked = target_value
            elif is_shift and self._last_clicked_sheet_node:
                self._set_checked_for_range(
                    self._last_clicked_sheet_node, node, target_value)
            else:
                node.IsChecked = target_value
            self._last_clicked_sheet_node = node
        else:
            if is_alt:
                for n in self._get_visible_sheet_nodes():
                    n.IsChecked = target_value
            else:
                self._set_checked_for_group(node, target_value)

        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self._refresh_sheet_counter()
        self.UI_tree_sheets.Items.Refresh()
        self._schedule_selection_refresh()
        self._set_status("Sheet selection updated.")

    def button_check_all(self, sender, e):
        self._set_all_checked(self._tree_roots, True)
        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self._refresh_sheet_counter()
        self.UI_tree_sheets.Items.Refresh()
        self._schedule_selection_refresh()
        self._set_status("Checked all visible sheets.")

    def button_uncheck_all(self, sender, e):
        self._set_all_checked(self._tree_roots, False)
        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self._refresh_sheet_counter()
        self.UI_tree_sheets.Items.Refresh()
        self._schedule_selection_refresh()
        self._set_status("Unchecked all visible sheets.")

    def button_expand_all(self, sender, e):
        self._set_expanded_for_groups(self._full_tree_roots, True)
        self._set_expanded_for_groups(self._tree_roots, True)
        self.UI_tree_sheets.Items.Refresh()
        self._set_status("Expanded all sheet groups.")

    def button_collapse_all(self, sender, e):
        self._set_expanded_for_groups(self._full_tree_roots, False)
        self._set_expanded_for_groups(self._tree_roots, False)
        self.UI_tree_sheets.Items.Refresh()
        self._set_status("Collapsed all sheet groups.")

    # Panel 3
    def UI_param_list_SelectionChanged(self, sender, e):
        item = self.UI_param_list.SelectedItem
        if not item:
            self._reset_inline_editor("None selected")
            return
        self._show_inline_editor(item)

    def button_refresh_params(self, sender, e):
        self._refresh_targets_and_params()
        self._set_status("Parameter list refreshed.")

    def button_apply(self, sender, e):
        item = self.UI_param_list.SelectedItem
        if not item:
            forms.alert("Select a parameter first.", title=__title__)
            return
        try:
            new_value = self._read_editor_value(item)
        except ValueError as ex:
            forms.alert(str(ex), title=__title__)
            return
        if not new_value and item.storage != StorageType.String:
            forms.alert("A value is required for this parameter.", title=__title__)
            return
        try:
            allowed_ids = set(elem.Id.IntegerValue for elem in self._targets)
            success, errors = _apply_parameter_value(item, new_value, allowed_element_ids=allowed_ids)
            _set_selection([elem.Id for elem in self._targets])
            self._refresh_targets_and_params()
            for p in self._param_items:
                if p.key == item.key:
                    self.UI_param_list.SelectedItem = p
                    break
            msg = "Applied '{}' to {} element(s).".format(item.name, success)
            if errors:
                msg += " ({} error(s))".format(errors)
            self.UI_feedback.Text    = msg
            self.UI_status_bar.Text  = msg
        except Exception as ex:
            forms.alert("Failed to apply: {}".format(ex), title=__title__)


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if not doc:
        forms.alert("Open a Revit document first.", exitscript=True)

    form = FORMATWindowV2("UnifiedEditor.xaml")
    form.ShowDialog()
