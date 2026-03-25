# -*- coding: utf-8 -*-
__title__ = "FORMAT"
__author__ = "Gino Moreno (GM)"
__doc__ = """Version = 3.1
Date    = 2026-03-21
Author  = GM
_____________________________________________________________________
Description:

Batch-edit parameters across multiple sheets in one session.

Targets (choose per run):
  - Sheet parameters
  - Title block parameters on sheets
  - Placed view parameters on sheets

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
  - "Back to Sheets" button to re-select sheets without rerunning the tool.
  - Session loop: closing the editor returns to sheet selection automatically.
_____________________________________________________________________
How-to:

1. Run FORMAT.
2. Check sheets in the picker (use Search to filter, group checkboxes for
   bulk selection, Shift-click for range, Alt-click for all visible).
3. Click "Use Selected Sheets".
4. Choose what to edit: Sheet / Title Block / Placed View parameters.
5. Pick a parameter from the list.
6. Set the value using the appropriate control (text, checkbox, or dropdown).
7. Click Apply.
8. Click "Back to Sheets" or close the editor to pick a new set of sheets.
_____________________________________________________________________
"""

from Autodesk.Revit.DB import (
        FilteredElementCollector,
        FilteredWorksetCollector,
        BuiltInCategory,
        BrowserOrganization,
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
clr.AddReference("System.Windows.Forms")
import System
from System.Windows.Forms import Control, Keys
from System.Diagnostics.Process import Start
from pyrevit import forms, script


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None
cfg = script.get_config()


class SheetTreeNode(object):
    """Tree node used by the sheet picker window."""
    def __init__(self, name, is_sheet=False, sheet=None):
        self.Name = name
        self.IsSheet = is_sheet
        self.Sheet = sheet
        self.Children = []
        self.IsChecked = False
        self.IsExpanded = True
        self.GroupCheckState = False


class SheetPickerWindow(forms.WPFWindow):
    def __init__(self, xaml_name, tree_roots):
        forms.WPFWindow.__init__(self, xaml_name)
        self.full_tree_roots = tree_roots or []
        self.tree_roots = self.full_tree_roots
        self.was_accepted = False
        self._last_clicked_sheet_node = None
        self._sync_group_checkstates(self.full_tree_roots)
        self.UI_tree_sheets.ItemsSource = self.tree_roots
        self._refresh_selected_count_label()
        self.ShowDialog()

    def button_select(self, sender, e):
        self.was_accepted = True
        self.DialogResult = True
        self.Close()

    def button_cancel(self, sender, e):
        self.was_accepted = False
        self.DialogResult = False
        self.Close()

    def button_check_all(self, sender, e):
        self._set_all_checked(self.tree_roots, True)
        self._sync_group_checkstates(self.full_tree_roots)
        self._sync_group_checkstates(self.tree_roots)
        self._refresh_selected_count_label()
        self.UI_tree_sheets.Items.Refresh()

    def button_uncheck_all(self, sender, e):
        self._set_all_checked(self.tree_roots, False)
        self._sync_group_checkstates(self.full_tree_roots)
        self._sync_group_checkstates(self.tree_roots)
        self._refresh_selected_count_label()
        self.UI_tree_sheets.Items.Refresh()

    def sheet_checkbox_click(self, sender, e):
        node = getattr(sender, "Tag", None)
        if not node:
            return

        target_value = self._normalize_checkbox_value(getattr(sender, "IsChecked", False))
        modifiers = Control.ModifierKeys
        is_shift = (modifiers & Keys.Shift) == Keys.Shift
        is_alt = (modifiers & Keys.Alt) == Keys.Alt

        if node.IsSheet:
            if is_alt:
                self._set_checked_for_nodes(self._get_visible_sheet_nodes(), target_value)
            elif is_shift and self._last_clicked_sheet_node:
                self._set_checked_for_range(self._last_clicked_sheet_node, node, target_value)
            else:
                node.IsChecked = target_value

            self._last_clicked_sheet_node = node
        else:
            if is_alt:
                self._set_checked_for_nodes(self._get_visible_sheet_nodes(), target_value)
            else:
                self._set_checked_for_group(node, target_value)

        self._sync_group_checkstates(self.full_tree_roots)
        self._sync_group_checkstates(self.tree_roots)
        self._refresh_selected_count_label()
        self.UI_tree_sheets.Items.Refresh()

    def _refresh_selected_count_label(self):
        if not hasattr(self, "UI_selected_count"):
            return
        selected_count = self._count_selected_sheets(self.full_tree_roots)
        self.UI_selected_count.Text = "Selected: {}".format(selected_count)

    def _count_selected_sheets(self, nodes):
        total = 0
        for node in nodes:
            if node.IsSheet:
                if node.IsChecked:
                    total += 1
            elif node.Children:
                total += self._count_selected_sheets(node.Children)
        return total

    def _normalize_checkbox_value(self, value):
        # Indeterminate click is treated as checked for quick bulk actions.
        return True if value is None else bool(value)

    def _set_checked_for_nodes(self, sheet_nodes, value):
        for s_node in sheet_nodes:
            s_node.IsChecked = value

    def _get_visible_sheet_nodes(self):
        visible = []

        def collect(nodes):
            for n in nodes:
                if n.IsSheet:
                    visible.append(n)
                elif n.Children:
                    collect(n.Children)

        collect(self.tree_roots)
        return visible

    def _set_checked_for_range(self, start_node, end_node, value):
        visible = self._get_visible_sheet_nodes()
        if not visible:
            return

        try:
            start_index = visible.index(start_node)
            end_index = visible.index(end_node)
        except ValueError:
            end_node.IsChecked = value
            return

        if start_index > end_index:
            start_index, end_index = end_index, start_index

        for idx in range(start_index, end_index + 1):
            visible[idx].IsChecked = value

    def _set_checked_for_group(self, group_node, value):
        if not group_node or group_node.IsSheet:
            return
        for child in group_node.Children:
            if child.IsSheet:
                child.IsChecked = value
            else:
                self._set_checked_for_group(child, value)

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

            checked_count = sum(1 for s_node in descendant_sheets if s_node.IsChecked)
            if checked_count == 0:
                node.GroupCheckState = False
            elif checked_count == len(descendant_sheets):
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

    def _set_all_checked(self, nodes, value):
        for node in nodes:
            if node.IsSheet:
                node.IsChecked = value
            if node.Children:
                self._set_all_checked(node.Children, value)

    def get_selected_sheets(self):
        selected = []

        def collect(nodes):
            for node in nodes:
                if node.IsSheet and node.Sheet and node.IsChecked:
                    selected.append(node.Sheet)
                if node.Children:
                    collect(node.Children)

        collect(self.full_tree_roots)
        return selected

    def search_textbox_changed(self, sender, e):
        """Rebuild visible tree to show only matching branches like Project Browser search."""
        search_text = self.UI_search.Text.lower().strip() if self.UI_search else ""

        if not search_text:
            self.tree_roots = self.full_tree_roots
        else:
            self.tree_roots = self._build_filtered_tree(self.full_tree_roots, search_text)

        self._sync_group_checkstates(self.tree_roots)
        self.UI_tree_sheets.ItemsSource = self.tree_roots
        self._refresh_selected_count_label()
        self.UI_tree_sheets.Items.Refresh()

    def _clone_group_subtree(self, group_node):
        """Clone a group branch while preserving sheet node references and check states."""
        clone = SheetTreeNode(group_node.Name, is_sheet=False)
        clone.IsExpanded = True
        for child in group_node.Children:
            if child.IsSheet:
                clone.Children.append(child)
            else:
                clone.Children.append(self._clone_group_subtree(child))
        return clone

    def _build_filtered_tree(self, nodes, search_text):
        """Return a pruned tree containing only matching sheets and required parent groups."""
        filtered = []
        for node in nodes:
            if node.IsSheet:
                if node.Sheet:
                    sheet_label = "{} - {}".format(node.Sheet.SheetNumber, node.Sheet.Name).lower()
                else:
                    sheet_label = node.Name.lower()

                if search_text in sheet_label:
                    filtered.append(node)
            else:
                group_matches = search_text in node.Name.lower()
                if group_matches:
                    filtered.append(self._clone_group_subtree(node))
                else:
                    child_matches = self._build_filtered_tree(node.Children, search_text)
                    if child_matches:
                        group_clone = SheetTreeNode(node.Name, is_sheet=False)
                        group_clone.IsExpanded = True
                        group_clone.Children = child_matches
                        filtered.append(group_clone)

        return filtered


def _selected_sheets():
    ids = list(uidoc.Selection.GetElementIds())
    sheets = []
    for eid in ids:
        e = doc.GetElement(eid)
        if not e:
            continue

        is_sheet = False
        if isinstance(e, ViewSheet):
            is_sheet = True
        else:
            try:
                is_sheet = (e.ViewType == ViewType.DrawingSheet)
            except Exception:
                is_sheet = False

        if is_sheet:
            sheets.append(e)
    return sheets


def _prompt_select_sheets():
    all_sheets = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Sheets
    ).WhereElementIsNotElementType().ToElements()

    if not all_sheets:
        return []

    def _clean_group_value(value):
        txt = (value or "").strip()
        if not txt:
            return ""
        if txt.lower() in ("<none>", "none"):
            return ""
        return txt

    def _get_sheet_param_text(sheet, candidate_names):
        wanted = set([n.replace("_", " ").strip().lower() for n in candidate_names])
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

    # If at least one sheet has a Sheet Collection value, keep it as the
    # top grouping level for all sheets; blanks become <No Sheet Collection>.
    has_any_sheet_collection = False
    for s in all_sheets:
        if _get_sheet_param_text(s, ["Sheet Collection"]):
            has_any_sheet_collection = True
            break

    def _browser_path_for_sheet(sheet):
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

        # Keep the hierarchy resilient even when some grouping params are blank.
        return path if path else ["Ungrouped"]

    # Start with a clean picker state so only explicit user checks are applied.
    selected_ids = set()

    roots = []
    root_index = {}

    def get_or_create_child(parent_node, child_name):
        key = (id(parent_node), child_name.lower())
        child = root_index.get(key)
        if child is not None:
            return child
        child = SheetTreeNode(child_name, is_sheet=False)
        parent_node.Children.append(child)
        root_index[key] = child
        return child

    def get_or_create_root(name):
        key = (None, name.lower())
        node = root_index.get(key)
        if node is not None:
            return node
        node = SheetTreeNode(name, is_sheet=False)
        roots.append(node)
        root_index[key] = node
        return node

    for sheet in all_sheets:
        group_path = _browser_path_for_sheet(sheet)
        if not group_path:
            group_path = ["Ungrouped"]

        current_parent = get_or_create_root(group_path[0])
        for folder_name in group_path[1:]:
            current_parent = get_or_create_child(current_parent, folder_name)

        label = "{} - {}".format(sheet.SheetNumber, sheet.Name)
        sheet_node = SheetTreeNode(label, is_sheet=True, sheet=sheet)
        sheet_node.IsChecked = sheet.Id.IntegerValue in selected_ids
        current_parent.Children.append(sheet_node)

    def sort_nodes(nodes):
        group_nodes = [n for n in nodes if not n.IsSheet]
        sheet_nodes = [n for n in nodes if n.IsSheet]

        group_nodes.sort(key=lambda n: n.Name.lower())
        sheet_nodes.sort(key=lambda n: n.Name.lower())

        for g in group_nodes:
            sort_nodes(g.Children)

        nodes[:] = group_nodes + sheet_nodes

    sort_nodes(roots)

    picker = SheetPickerWindow("SheetPicker.xaml", roots)
    if not picker.was_accepted:
        return []

    return picker.get_selected_sheets()


def _selected_titleblocks():
    ids = list(uidoc.Selection.GetElementIds())
    tbs = []
    for eid in ids:
        e = doc.GetElement(eid)
        if not e:
            continue
        cat = getattr(e, "Category", None)
        if cat and cat.Id.IntegerValue == int(BuiltInCategory.OST_TitleBlocks):
            tbs.append(e)
    return tbs


def _collect_titleblocks_on_sheets(sheets):
    """Fast path: scoped collector by sheet view id."""
    tb_ids = []
    for sheet in sheets:
        ids = FilteredElementCollector(doc, sheet.Id) \
            .OfCategory(BuiltInCategory.OST_TitleBlocks) \
            .WhereElementIsNotElementType() \
            .ToElementIds()
        if ids and ids.Count:
            tb_ids.extend(list(ids))
    return tb_ids


def _collect_placed_views_on_sheets(sheets):
    view_ids = []
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
            if not view:
                continue
            seen.add(key)
            view_ids.append(vid)
    return view_ids


def _set_selection(ids):
    uidoc.Selection.SetElementIds(List[ElementId](ids))


def _save_last(ids):
    cfg.last_titleblock_ids = [str(i.IntegerValue) for i in ids]
    script.save_config()


def _load_last_existing_ids():
    raw = getattr(cfg, "last_titleblock_ids", None)
    if not raw:
        return []

    valid = []
    for sid in raw:
        try:
            eid = ElementId(int(sid))
            if doc.GetElement(eid):
                valid.append(eid)
        except Exception:
            continue
    return valid


def _storage_name(storage_type):
    if storage_type == StorageType.String:
        return "String"
    if storage_type == StorageType.Integer:
        return "Integer"
    if storage_type == StorageType.Double:
        return "Double"
    if storage_type == StorageType.ElementId:
        return "ElementId"
    return "Unknown"


def _collect_named_elementid_options(param_name):
    pname = (param_name or "").strip().lower()
    options = []

    if "template" in pname:
        for view in FilteredElementCollector(doc).OfClass(View).ToElements():
            try:
                if view.IsTemplate:
                    options.append((str(view.Name), view.Id))
            except Exception:
                continue
        return options

    if "scope" in pname or "volume of interest" in pname:
        for scope_box in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(scope_box.Name), scope_box.Id))
            except Exception:
                continue
        return options

    if "level" in pname:
        for level in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(level.Name), level.Id))
            except Exception:
                continue
        return options

    if "phase" in pname:
        try:
            for phase in doc.Phases:
                options.append((str(phase.Name), phase.Id))
        except Exception:
            pass
        return options

    if "workset" in pname:
        try:
            for workset in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets():
                options.append((str(workset.Name), ElementId(workset.Id.IntegerValue)))
        except Exception:
            pass
        return options

    if "design option" in pname or "option" in pname:
        for design_option in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DesignOptions).WhereElementIsNotElementType().ToElements():
            try:
                options.append((str(design_option.Name), design_option.Id))
            except Exception:
                continue
        return options

    if "view" in pname:
        for view in FilteredElementCollector(doc).OfClass(View).ToElements():
            try:
                if not view.IsTemplate:
                    options.append((str(view.Name), view.Id))
            except Exception:
                continue
        return options

    if "sheet" in pname:
        for sheet in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements():
            try:
                label = "{} - {}".format(sheet.SheetNumber, sheet.Name)
                options.append((label, sheet.Id))
            except Exception:
                continue
        return options

    return options


def _build_named_elementid_display_options(param_name):
    display_names = ["<None>"]
    seen = set()
    for display_name, _ in _collect_named_elementid_options(param_name):
        lowered = display_name.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
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
        if vs:
            return vs
        return str(param.AsDouble())
    if st == StorageType.ElementId:
        return str(param.AsElementId().IntegerValue)
    return ""


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
        if len(uniq) == 1:
            self.current_value = uniq[0]
        else:
            self.current_value = "<varies>"

        # Determine UI control type based on storage and name hints.
        self.is_boolean = (storage == StorageType.Integer and name.lower().startswith("is "))
        self.control_type = "checkbox" if self.is_boolean else "text"
        if storage == StorageType.ElementId:
            self.control_type = "elementid"
        
        # Show exactly the same parameter text users see in Revit Properties.
        self.Display = name


def _collect_parameter_items(target_elements):
    collected = {}
    total_count = len(target_elements)

    def collect_from_target(scope, element_id, target):
        for param in target.Parameters:
            try:
                if param.IsReadOnly:
                    continue
                if str(param.StorageType) == "None":
                    continue

                p_name = param.Definition.Name
                p_id = param.Id.IntegerValue
                key = (scope, p_name, p_id)

                if key not in collected:
                    is_shared = False
                    try:
                        is_shared = param.IsShared
                    except Exception:
                        pass

                    collected[key] = {
                        "scope": scope,
                        "name": p_name,
                        "storage": param.StorageType,
                        "is_shared": is_shared,
                        "params_by_id": {},
                    }

                collected[key]["params_by_id"][element_id.IntegerValue] = param
            except Exception:
                continue

    for element in target_elements:
        collect_from_target("Instance", element.Id, element)

    items = []
    for key, data in collected.items():
        items.append(
            ParameterItem(
                key=key,
                scope=data["scope"],
                name=data["name"],
                storage=data["storage"],
                is_shared=data["is_shared"],
                params_by_id=data["params_by_id"],
                total_target_count=total_count,
            )
        )

    return sorted(items, key=lambda i: (i.scope, i.name.lower()))


def _apply_parameter_value(param_item, new_value):
    if not new_value and param_item.storage != StorageType.String:
        raise ValueError("Value is required for non-string parameters.")

    def _resolve_elementid_from_text(param_obj, text_value):
        pname = ""
        try:
            pname = param_obj.Definition.Name
        except Exception:
            pname = ""

        resolved = _resolve_named_elementid_from_text(pname, text_value)
        if resolved is not None:
            return resolved

        raw = (text_value or "").strip()
        raise ValueError("Could not resolve ElementId value '{}' for parameter '{}'".format(raw, param_obj.Definition.Name))

    t = Transaction(doc, "FORMAT: Apply parameter")
    t.Start()
    success = 0
    errors = 0
    try:
        for p in param_item.params_by_id.values():
            try:
                if param_item.storage == StorageType.String:
                    p.Set(new_value)
                elif param_item.storage == StorageType.Integer:
                    text = new_value.strip().lower()
                    if text in ("true", "yes", "1", "on"):
                        p.Set(1)
                    elif text in ("false", "no", "0", "off"):
                        p.Set(0)
                    else:
                        p.Set(int(new_value))
                elif param_item.storage == StorageType.Double:
                    ok = p.SetValueString(new_value)
                    if not ok:
                        p.Set(float(new_value))
                elif param_item.storage == StorageType.ElementId:
                    resolved_id = _resolve_elementid_from_text(p, new_value)
                    p.Set(resolved_id)
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


class FORMATWindow(forms.WPFWindow):
    def __init__(self, xaml_name, targets, target_label):
        forms.WPFWindow.__init__(self, xaml_name)
        self.targets = targets
        self.target_ids = [e.Id for e in targets]
        self.target_label = target_label
        self.param_items = []
        self._bind_editor_controls()

        self.main_title.Text = __title__
        self._refresh_parameters()
        self.ShowDialog()

    def _bind_editor_controls(self):
        """Support both legacy and current XAML control names."""
        text_input = getattr(self, "UI_param_input", None)
        legacy_text_input = getattr(self, "UI_current_value", None)

        if text_input is None:
            text_input = legacy_text_input
        if legacy_text_input is None:
            legacy_text_input = text_input

        self.UI_param_input = text_input
        self.UI_current_value = legacy_text_input
        self.UI_param_checkbox = getattr(self, "UI_param_checkbox", None)
        self.UI_param_elementid = getattr(self, "UI_param_elementid", None)

    @property
    def text_input_control(self):
        return self.UI_param_input or self.UI_current_value

    def _reset_inline_editor(self, label_text):
        self.UI_param_name.Text = label_text

        if self.UI_param_checkbox:
            self.UI_param_checkbox.Visibility = System.Windows.Visibility.Collapsed
            self.UI_param_checkbox.IsChecked = False

        text_input = self.text_input_control
        if text_input:
            text_input.Visibility = System.Windows.Visibility.Collapsed
            text_input.Text = ""

        if self.UI_param_elementid:
            self.UI_param_elementid.Visibility = System.Windows.Visibility.Collapsed
            self.UI_param_elementid.ItemsSource = []
            self.UI_param_elementid.SelectedIndex = -1

    def _build_elementid_options(self, item):
        """Return user-friendly names for common ElementId parameters."""
        return _build_named_elementid_display_options(item.name)

    def _current_elementid_display(self, item):
        return _current_named_elementid_display(item.name, item.current_value)

    def _is_checked_value(self, value_text):
        text = "" if value_text is None else str(value_text).strip().lower()
        if text in ("1", "true", "yes", "on"):
            return True
        if text in ("0", "false", "no", "off", "", "<varies>"):
            return False

        try:
            return int(text) != 0
        except Exception:
            return False

    def _refresh_parameters(self):
        self.param_items = _collect_parameter_items(self.targets)
        self.UI_param_list.DisplayMemberPath = "Display"
        self.UI_param_list.ItemsSource = self.param_items

        self.UI_tb_count.Text = "Targets ({}): {}".format(self.target_label, len(self.targets))
        self.UI_param_count.Text = "Editable Params: {}".format(len(self.param_items))
        self._reset_inline_editor("Select a parameter...")

    @property
    def selected_param_item(self):
        return self.UI_param_list.SelectedItem

    def button_close(self, sender, e):
        self.Close()

    def button_back_to_sheets(self, sender, e):
        self.Close()

    def header_drag(self, sender, e):
        if str(e.LeftButton) == "Pressed":
            self.DragMove()

    def Hyperlink_RequestNavigate(self, sender, e):
        Start(e.Uri.AbsoluteUri)

    def UI_param_list_SelectionChanged(self, sender, e):
        """Update visible inline editor based on selected parameter type."""
        item = self.selected_param_item
        if not item:
            self._reset_inline_editor("None selected")
            return

        # Update header with parameter name
        self.UI_param_name.Text = item.name
        text_input = self.text_input_control
        if self.UI_param_elementid:
            self.UI_param_elementid.Visibility = System.Windows.Visibility.Collapsed
            self.UI_param_elementid.ItemsSource = []
            self.UI_param_elementid.SelectedIndex = -1
        
        # Show/hide controls based on parameter storage type
        if item.is_boolean:
            # Show checkbox for boolean ("Is " prefix) parameters
            if self.UI_param_checkbox:
                self.UI_param_checkbox.Visibility = System.Windows.Visibility.Visible
                self.UI_param_checkbox.IsChecked = self._is_checked_value(item.current_value)
            if text_input:
                text_input.Visibility = System.Windows.Visibility.Collapsed
                text_input.Text = ""
            if not self.UI_param_checkbox:
                forms.alert("Boolean editor control is not available in this FORMAT UI.", title=__title__)
                return
        elif item.storage == StorageType.ElementId:
            if self.UI_param_checkbox:
                self.UI_param_checkbox.Visibility = System.Windows.Visibility.Collapsed
                self.UI_param_checkbox.IsChecked = False
            if text_input:
                text_input.Visibility = System.Windows.Visibility.Collapsed
                text_input.Text = ""

            options = self._build_elementid_options(item)
            current_display = self._current_elementid_display(item)

            if self.UI_param_elementid and len(options) > 1:
                self.UI_param_elementid.ItemsSource = options
                self.UI_param_elementid.Visibility = System.Windows.Visibility.Visible
                try:
                    self.UI_param_elementid.SelectedItem = current_display if current_display in options else "<None>"
                except Exception:
                    self.UI_param_elementid.SelectedIndex = 0
            elif text_input:
                # Fallback for legacy XAML or unknown ElementId parameters.
                text_input.Visibility = System.Windows.Visibility.Visible
                text_input.Text = current_display
        else:
            # Show textbox for string/number parameters
            if self.UI_param_checkbox:
                self.UI_param_checkbox.Visibility = System.Windows.Visibility.Collapsed
                self.UI_param_checkbox.IsChecked = False
            if not text_input:
                forms.alert("Text editor control is not available in this FORMAT UI.", title=__title__)
                return
            text_input.Visibility = System.Windows.Visibility.Visible
            text_input.Text = str(item.current_value) if item.current_value else ""

    def button_refresh(self, sender, e):
        refreshed = [doc.GetElement(eid) for eid in self.target_ids if doc.GetElement(eid)]
        self.targets = refreshed
        self.target_ids = [e.Id for e in refreshed]
        _set_selection(self.target_ids)
        self._refresh_parameters()

    def button_apply(self, sender, e):
        """Apply the edited value from the inline editor to all targeted elements."""
        item = self.selected_param_item
        if not item:
            forms.alert("Select a parameter first.")
            return

        # Read value from the appropriate control
        if item.is_boolean:
            # Convert checkbox state to 0/1
            if not self.UI_param_checkbox:
                forms.alert("Boolean editor control is not available in this FORMAT UI.", title=__title__)
                return
            new_value = "1" if self.UI_param_checkbox.IsChecked else "0"
        elif item.storage == StorageType.ElementId:
            if self.UI_param_elementid and self.UI_param_elementid.Visibility == System.Windows.Visibility.Visible:
                selected = self.UI_param_elementid.SelectedItem
                if selected is None:
                    forms.alert("Select a value for '{}' first.".format(item.name), title=__title__)
                    return
                new_value = str(selected)
            else:
                text_input = self.text_input_control
                if not text_input:
                    forms.alert("ElementId editor control is not available in this FORMAT UI.", title=__title__)
                    return
                new_value = text_input.Text
        else:
            # Read from textbox
            text_input = self.text_input_control
            if not text_input:
                forms.alert("Text editor control is not available in this FORMAT UI.", title=__title__)
                return
            new_value = text_input.Text
        
        try:
            success, errors = _apply_parameter_value(item, new_value)
            _set_selection(self.target_ids)
            if self.target_label == "TitleBlocks":
                _save_last(self.target_ids)
            self._refresh_parameters()
            forms.alert(
                "Applied parameter '{}' to {} target element(s). Errors: {}".format(item.name, success, errors),
                title=__title__,
                warn_icon=False,
            )
        except Exception as ex:
            forms.alert("Failed to apply value: {}".format(ex), title=__title__)


def _resolve_targets():
    sheets = _prompt_select_sheets()
    if not sheets:
        return None, None, False

    edit_choice = forms.CommandSwitchWindow.show(
        [
            "Edit sheet parameters",
            "Edit title block parameters on sheets",
            "Edit placed view parameters on sheets",
        ],
        message="Choose what to edit",
        title=__title__,
    )
    if not edit_choice:
        return None, None, False

    if edit_choice == "Edit sheet parameters":
        ids = [s.Id for s in sheets]
        _set_selection(ids)
        return sheets, "Sheets", True

    if edit_choice == "Edit title block parameters on sheets":
        ids = _collect_titleblocks_on_sheets(sheets)
        if not ids:
            forms.alert("No title blocks found on selected sheets.", title=__title__)
            return None, None, True
        _set_selection(ids)
        _save_last(ids)
        return [doc.GetElement(eid) for eid in ids if doc.GetElement(eid)], "TitleBlocks", True

    view_ids = _collect_placed_views_on_sheets(sheets)
    if not view_ids:
        forms.alert("No placed views found on selected sheets.", title=__title__)
        return None, None, True
    _set_selection(view_ids)
    return [doc.GetElement(eid) for eid in view_ids if doc.GetElement(eid)], "PlacedViews", True


if __name__ == '__main__':
    if not doc:
        forms.alert("Open a Revit document first.", exitscript=True)

    while True:
        targets, target_label, keep_running = _resolve_targets()
        if not keep_running:
            break
        if not targets:
            continue

        FORMATWindow("Script.xaml", targets, target_label)
