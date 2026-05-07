# -*- coding: utf-8 -*-
__title__ = "FORMAT"
__author__ = "Gino Moreno (GM)"
__doc__ = """Version = 4.0
Date    = 2026-03-25
Author  = GM
_____________________________________________________________________
Description:

FORMAT v2 — Unified single-window batch parameter editor.

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
  - Single-window session: no modal dialogs after launch.
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
        ViewSheet,
        ViewType,
        View,
        ElementId,
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
    options = []

    if "template" in pname:
        for v in FilteredElementCollector(doc).OfClass(View).ToElements():
            try:
                if v.IsTemplate:
                    options.append((str(v.Name), v.Id))
            except Exception:
                continue
        return options

    if "scope" in pname or "volume of interest" in pname:
        for sb in FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(sb.Name), sb.Id))
            except Exception:
                continue
        return options

    if "level" in pname:
        for lv in FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(lv.Name), lv.Id))
            except Exception:
                continue
        return options

    if "phase" in pname:
        try:
            for ph in doc.Phases:
                options.append((str(ph.Name), ph.Id))
        except Exception:
            pass
        return options

    if "workset" in pname:
        try:
            for ws in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets():
                options.append((str(ws.Name), ElementId(ws.Id.IntegerValue)))
        except Exception:
            pass
        return options

    if "design option" in pname or "option" in pname:
        for do in FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_DesignOptions).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(do.Name), do.Id))
            except Exception:
                continue
        return options

    if "view" in pname:
        for v in FilteredElementCollector(doc).OfClass(View).ToElements():
            try:
                if not v.IsTemplate:
                    options.append((str(v.Name), v.Id))
            except Exception:
                continue
        return options

    if "sheet" in pname:
        for s in FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements():
            try:
                label = "{} - {}".format(s.SheetNumber, s.Name)
                options.append((label, s.Id))
            except Exception:
                continue
        return options

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


def _current_named_elementid_display(param_name, current_value):
    if current_value in (None, "", "<varies>"):
        return "<None>"
    try:
        current_id = int(str(current_value).strip())
    except Exception:
        return str(current_value)
    if current_id < 0:
        return "<None>"
    for option_name, option_id in _collect_named_elementid_options(param_name):
        try:
            if option_id.IntegerValue == current_id:
                return option_name
        except Exception:
            continue
    try:
        elem = doc.GetElement(ElementId(current_id))
        name = getattr(elem, "Name", None) if elem else None
        if name:
            return str(name)
    except Exception:
        pass
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

        values = [_parameter_value_text(p) for p in params_by_id.values()]
        uniq = sorted(set(values))
        self.current_value = uniq[0] if len(uniq) == 1 else "<varies>"

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

    def _build_value_display(self):
        if self.current_value == "<varies>":
            return "<varies>"

        if self.is_boolean:
            raw = str(self.current_value).strip().lower()
            return "Yes" if raw in ("1", "true", "yes", "on") else "No"

        if self.storage == StorageType.ElementId:
            return _current_named_elementid_display(self.name, self.current_value)

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


def _apply_parameter_value(param_item, new_value):
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
        for p in param_item.params_by_id.values():
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
    tb_ids = []
    for sheet in sheets:
        ids = (FilteredElementCollector(doc, sheet.Id)
               .OfCategory(BuiltInCategory.OST_TitleBlocks)
               .WhereElementIsNotElementType()
               .ToElementIds())
        if ids and ids.Count:
            tb_ids.extend(list(ids))
    return [doc.GetElement(eid) for eid in tb_ids if doc.GetElement(eid)]


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
        self._pending_search_text      = ""
        self._last_applied_search_text = None

        # Keep maximize constrained to working area for a borderless window.
        try:
            wa = System.Windows.SystemParameters.WorkArea
            self.MaxHeight = wa.Height + 12
            self.MaxWidth = wa.Width + 12
        except Exception:
            pass

        self._load_all_sheets()
        self._build_tree()
        self._init_search_debounce()
        self._apply_tree_to_ui()
        self._refresh_targets_and_params()
        self._update_window_buttons()

        self.ShowDialog()

    def _init_search_debounce(self):
        """Delay filter execution while user is typing to avoid UI freezes."""
        self._search_timer = System.Windows.Threading.DispatcherTimer()
        self._search_timer.Interval = System.TimeSpan.FromMilliseconds(250)
        self._search_timer.Tick += self._on_search_timer_tick

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

    def _update_window_buttons(self):
        try:
            ws = self.WindowState
            if ws == System.Windows.WindowState.Maximized:
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
            current_display = _current_named_elementid_display(item.name, item.current_value)
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
        for child in group_node.Children:
            if child.IsSheet:
                child.IsChecked = value
            else:
                self._set_checked_for_group(child, value)

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
        if str(e.LeftButton) == "Pressed":
            self.DragMove()

    def button_minimize(self, sender, e):
        self.WindowState = System.Windows.WindowState.Minimized

    def button_maximize(self, sender, e):
        if self.WindowState == System.Windows.WindowState.Maximized:
            self.WindowState = System.Windows.WindowState.Normal
        else:
            self.WindowState = System.Windows.WindowState.Maximized
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

    def button_reset_layout(self, sender, e):
        star = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        self.UI_col_panel_1.Width = star
        self.UI_col_panel_2.Width = star
        self.UI_col_panel_3.Width = star
        self._set_status("Panel widths reset to equal layout.")

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
        self._refresh_targets_and_params()
        self._set_status("Sheet selection updated.")

    def button_check_all(self, sender, e):
        self._set_all_checked(self._tree_roots, True)
        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self._refresh_sheet_counter()
        self.UI_tree_sheets.Items.Refresh()
        self._refresh_targets_and_params()
        self._set_status("Checked all visible sheets.")

    def button_uncheck_all(self, sender, e):
        self._set_all_checked(self._tree_roots, False)
        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self._refresh_sheet_counter()
        self.UI_tree_sheets.Items.Refresh()
        self._refresh_targets_and_params()
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
            success, errors = _apply_parameter_value(item, new_value)
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

    FORMATWindowV2("UnifiedEditor.xaml")
