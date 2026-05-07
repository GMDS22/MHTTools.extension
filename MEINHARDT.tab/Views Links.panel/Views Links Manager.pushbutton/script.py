# -*- coding: utf-8 -*-
__title__ = "Views Links Manager"
__doc__ = """Version = 1.0
Date: 2026-03-25
Author: GM
Description:
Manage Revit link display settings across placed views on selected sheets.
How-to:
1. Select sheets in the center panel.
2. Confirm which placed views should be edited.
3. Select the links to manage.
4. Choose display, visibility, and grid actions.
5. Apply the batch update.
"""

import clr
clr.AddReference("System")
clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Windows.Forms")

import System
from System.Windows.Forms import Control, Keys
from System.Diagnostics.Process import Start
from System.Collections.Generic import List

from pyrevit import DB, forms, revit, script


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None


class SheetTreeNode(object):
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


class ViewRow(object):
    def __init__(self, sheet, view, is_checked=True):
        self.Sheet = sheet
        self.View = view
        self.IsChecked = is_checked
        self.SheetNumber = getattr(sheet, "SheetNumber", "") or ""
        self.SheetName = getattr(sheet, "Name", "") or ""
        self.ViewName = getattr(view, "Name", "") or ""
        self.ViewType = self._safe_view_type_name(view)
        self.LevelName = self._safe_level_name(view)
        self.TemplateName = self._safe_template_name(view)
        self.Display = "{} | {} | {} | {}".format(
            self.SheetNumber,
            self.ViewName,
            self.LevelName,
            self.ViewType,
        )
        self.SearchText = "{} {} {} {} {} {}".format(
            self.SheetNumber,
            self.SheetName,
            self.ViewName,
            self.LevelName,
            self.ViewType,
            self.TemplateName,
        ).lower()

    def _safe_level_name(self, view):
        try:
            level = _safe_get_view_level(view)
            if level:
                return level.Name
        except Exception:
            pass
        return "No Level"

    def _safe_view_type_name(self, view):
        try:
            return str(view.ViewType)
        except Exception:
            return "Unknown"

    def _safe_template_name(self, view):
        try:
            template_id = view.ViewTemplateId
            if template_id and template_id != DB.ElementId.InvalidElementId:
                template = doc.GetElement(template_id)
                if template:
                    return template.Name or ""
        except Exception:
            pass
        return ""


class LinkRow(object):
    def __init__(self, link_type, loaded, instance_count, is_checked=False):
        self.LinkType = link_type
        self.IsChecked = is_checked
        self.IsLoaded = loaded
        self.InstanceCount = instance_count
        self.Name = _safe_link_type_name(link_type)
        self.IsArchitectural = _looks_architectural(self.Name)
        tags = []
        tags.append("loaded" if loaded else "unloaded")
        tags.append("{} inst".format(instance_count))
        if self.IsArchitectural:
            tags.append("ARCH")
        self.Display = "{} [{}]".format(self.Name, ", ".join(tags))
        self.SearchText = (self.Name or "").lower()


class LinkedViewChoice(object):
    def __init__(self, display, mode, view_name=None, view_id=None):
        self.Display = display
        self.Mode = mode
        self.ViewName = view_name or ""
        self.ViewId = view_id

    def __str__(self):
        return self.Display


class ReferenceLinkChoice(object):
    def __init__(self, display, link_row=None, mode="link"):
        self.Display = display
        self.LinkRow = link_row
        self.Mode = mode

    def __str__(self):
        return self.Display


class SettingChoice(object):
    def __init__(self, display, value=None, mode="value"):
        self.Display = display
        self.Value = display if value is None else value
        self.Mode = mode

    def __str__(self):
        return self.Display


def _clean_group_value(value):
    text = (value or "").strip()
    if not text or text.lower() in ("<none>", "none"):
        return ""
    return text


def _get_sheet_param_text(sheet, candidate_names):
    wanted = set(name.replace("_", " ").strip().lower() for name in candidate_names)
    try:
        for param in sheet.Parameters:
            try:
                definition = getattr(param, "Definition", None)
                param_name = definition.Name if definition else ""
                normalized = param_name.replace("_", " ").strip().lower() if param_name else ""
                if normalized not in wanted:
                    continue
                value = ""
                try:
                    value = param.AsString() or ""
                except Exception:
                    value = ""
                if not value:
                    try:
                        value = param.AsValueString() or ""
                    except Exception:
                        value = ""
                return _clean_group_value(value)
            except Exception:
                continue
    except Exception:
        pass
    return ""


def _build_sheet_tree(all_sheets, preselected_ids=None):
    preselected_ids = preselected_ids or set()
    has_any_sheet_collection = any(
        _get_sheet_param_text(sheet, ["Sheet Collection"])
        for sheet in all_sheets
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
        groups = [node for node in nodes if not node.IsSheet]
        sheets = [node for node in nodes if node.IsSheet]
        groups.sort(key=lambda item: item.Name.lower())
        sheets.sort(key=lambda item: item.Name.lower())
        for group in groups:
            _sort_nodes(group.Children)
        nodes[:] = groups + sheets

    _sort_nodes(roots)
    return roots, root_index


def _safe_get_view_level(view):
    try:
        if hasattr(view, "GenLevel"):
            return view.GenLevel
    except Exception:
        return None
    return None


def _safe_link_type_name(link_type):
    try:
        return link_type.Name or "<Unnamed Link>"
    except Exception:
        try:
            parameter = link_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if parameter:
                return parameter.AsString() or "<Unnamed Link>"
        except Exception:
            pass
    return "<Unnamed Link>"


def _looks_architectural(text):
    normalized = (text or "").lower()
    tokens = ("arch", "architect", "arc", "a-model", "a_")
    return any(token in normalized for token in tokens)


def _find_link_doc_for_type(host_doc, link_type_id):
    try:
        instances = DB.FilteredElementCollector(host_doc).OfClass(DB.RevitLinkInstance).ToElements()
        for instance in instances:
            try:
                if instance.GetTypeId() == link_type_id:
                    link_doc = instance.GetLinkDocument()
                    if link_doc:
                        return link_doc
            except Exception:
                continue
    except Exception:
        pass
    return None


def _safe_element_name(element):
    try:
        name = getattr(element, "Name", None)
        if name:
            return name
    except Exception:
        pass
    try:
        return DB.Element.Name.GetValue(element) or ""
    except Exception:
        return ""


def _safe_view_choice_display(view):
    view_name = _safe_element_name(view) or "<Unnamed View>"
    try:
        view_type = str(view.ViewType)
    except Exception:
        view_type = "View"
    level_name = ""
    try:
        level = _safe_get_view_level(view)
        if level:
            level_name = _safe_element_name(level)
    except Exception:
        level_name = ""
    if level_name:
        return "{} | {} | {}".format(view_name, level_name, view_type)
    return "{} | {}".format(view_name, view_type)


def _collect_link_views(link_doc):
    views = []
    try:
        for view in DB.FilteredElementCollector(link_doc).OfClass(DB.View).ToElements():
            try:
                if view.IsTemplate:
                    continue
                view_name = _safe_element_name(view)
                if not view_name:
                    continue
                views.append(view)
            except Exception:
                continue
    except Exception:
        return []
    return views


def _find_link_view_by_name(link_doc, view_name):
    if not link_doc or not view_name:
        return None
    wanted = (view_name or "").strip().lower()
    for view in _collect_link_views(link_doc):
        try:
            if (_safe_element_name(view) or "").strip().lower() == wanted:
                return view
        except Exception:
            continue
    return None


def _find_link_view_by_id(link_doc, view_id):
    if not link_doc or view_id in (None, DB.ElementId.InvalidElementId):
        return None
    try:
        return link_doc.GetElement(view_id)
    except Exception:
        return None


def _find_link_phase_by_name(link_doc, phase_name):
    if not link_doc or not phase_name:
        return None
    wanted = (phase_name or "").strip().lower()
    try:
        for phase in list(link_doc.Phases):
            try:
                if (phase.Name or "").strip().lower() == wanted:
                    return phase
            except Exception:
                continue
    except Exception:
        pass
    return None


def _find_link_phase_filter_by_name(link_doc, filter_name):
    if not link_doc or not filter_name:
        return None
    wanted = (filter_name or "").strip().lower()
    try:
        filters = DB.FilteredElementCollector(link_doc).OfClass(DB.PhaseFilter).ToElements()
    except Exception:
        filters = []
    for phase_filter in filters:
        try:
            if (phase_filter.Name or "").strip().lower() == wanted:
                return phase_filter
        except Exception:
            continue
    return None


def _placeholder_choice(text):
    return SettingChoice(text, value=None, mode="placeholder")


def _get_enum_member(enum_type, names):
    if enum_type is None:
        return None
    for name in names:
        try:
            value = getattr(enum_type, name, None)
            if value is not None:
                return value
        except Exception:
            continue
    return None


def _get_link_visibility_member(mode_label):
    link_visibility_enum = getattr(DB, "LinkVisibility", None)
    if link_visibility_enum is None:
        return None

    mapping = {
        "By host view": ("ByHostView", "ByHost"),
        "By linked view": ("ByLinkedView", "ByLinkView"),
        "Custom / None": ("Custom", "None"),
    }
    return _get_enum_member(link_visibility_enum, mapping.get(mode_label, ()))


def _get_view_discipline_member(mode_label):
    view_discipline_enum = getattr(DB, "ViewDiscipline", None)
    if view_discipline_enum is None:
        return None
    mapping = {
        "Architectural": ("Architectural",),
        "Coordination": ("Coordination",),
        "Mechanical": ("Mechanical",),
        "Electrical": ("Electrical",),
        "Plumbing": ("Plumbing",),
        "Structural": ("Structural",),
    }
    return _get_enum_member(view_discipline_enum, mapping.get(mode_label, ()))


def _get_view_detail_level_member(mode_label):
    view_detail_level_enum = getattr(DB, "ViewDetailLevel", None)
    if view_detail_level_enum is None:
        return None
    mapping = {
        "Coarse": ("Coarse",),
        "Medium": ("Medium",),
        "Fine": ("Fine",),
        "Undefined": ("Undefined",),
    }
    return _get_enum_member(view_detail_level_enum, mapping.get(mode_label, ()))


def _safe_invoke_method(target, method_name, argument_sets):
    method = getattr(target, method_name, None)
    if method is None:
        return False, "{} is not available".format(method_name)

    last_error = None
    for args in argument_sets:
        try:
            method(*args)
            return True, None
        except Exception as exc:
            last_error = exc
    if last_error is None:
        return False, "{} did not receive a compatible argument set".format(method_name)
    return False, "{} failed: {}".format(method_name, last_error)


def _current_setting_text(control, default="No change"):
    try:
        return str(control.SelectedItem or default)
    except Exception:
        return default


CUSTOM_LINK_VISIBILITY_OPTIONS = [
    "No change",
    "By host view",
    "By linked view",
    "Custom / None",
]

LINK_VISIBILITY_OPTIONS = [
    "No change",
    "By host view",
    "By linked view",
]

UNDERLAY_OPTIONS = [
    "No change",
    "By host view",
    "By linked view",
    "Custom / None",
]

DISCIPLINE_OPTIONS = [
    "No change",
    "By host view",
    "By linked view",
    "Architectural",
    "Coordination",
    "Mechanical",
    "Electrical",
    "Plumbing",
    "Structural",
]

DETAIL_LEVEL_OPTIONS = [
    "No change",
    "By host view",
    "By linked view",
    "Coarse",
    "Medium",
    "Fine",
]

PHASE_MODE_OPTIONS = [
    "No change",
    "By host view",
    "By linked view",
    "Custom phase",
]

PHASE_FILTER_MODE_OPTIONS = [
    "No change",
    "By host view",
    "By linked view",
    "Custom phase filter",
]


def _best_match_link_level(link_doc, host_level):
    if not link_doc or not host_level:
        return None
    try:
        link_levels = DB.FilteredElementCollector(link_doc).OfClass(DB.Level).ToElements()
    except Exception:
        return None
    if not link_levels:
        return None

    host_name = None
    host_elevation = None
    try:
        host_name = host_level.Name
    except Exception:
        host_name = None
    try:
        host_elevation = host_level.Elevation
    except Exception:
        host_elevation = None

    if host_name:
        for level in link_levels:
            try:
                if level.Name == host_name:
                    return level
            except Exception:
                continue

    if host_elevation is not None:
        best = None
        best_delta = None
        for level in link_levels:
            try:
                delta = abs(level.Elevation - host_elevation)
                if best_delta is None or delta < best_delta:
                    best_delta = delta
                    best = level
            except Exception:
                continue
        return best

    return None


def _choose_linked_view_for_level(link_doc, src_link_view, dst_host_level, fallback_viewtype=None):
    if not link_doc:
        return None
    dst_link_level = _best_match_link_level(link_doc, dst_host_level)
    if not dst_link_level:
        return None

    src_viewtype = fallback_viewtype
    src_template_id = None
    src_name = None
    src_level_name = None
    if src_link_view:
        try:
            src_viewtype = src_link_view.ViewType
        except Exception:
            pass
        try:
            src_template_id = src_link_view.ViewTemplateId
        except Exception:
            src_template_id = None
        try:
            src_name = src_link_view.Name
        except Exception:
            src_name = None
        try:
            if hasattr(src_link_view, "GenLevel") and src_link_view.GenLevel:
                src_level_name = src_link_view.GenLevel.Name
        except Exception:
            src_level_name = None

    desired_name = None
    try:
        if src_name and src_level_name and dst_link_level.Name and src_level_name in src_name:
            desired_name = src_name.replace(src_level_name, dst_link_level.Name)
    except Exception:
        desired_name = None

    candidates = []
    try:
        for view in DB.FilteredElementCollector(link_doc).OfClass(DB.View).ToElements():
            try:
                if view.IsTemplate:
                    continue
                if src_viewtype is not None and view.ViewType != src_viewtype:
                    continue
                if not hasattr(view, "GenLevel") or not view.GenLevel:
                    continue
                if view.GenLevel.Id != dst_link_level.Id:
                    continue
                candidates.append(view)
            except Exception:
                continue
    except Exception:
        return None

    if not candidates:
        return None

    if desired_name:
        for view in candidates:
            try:
                if view.Name == desired_name:
                    return view
            except Exception:
                continue

    if src_template_id and src_template_id != DB.ElementId.InvalidElementId:
        for view in candidates:
            try:
                if view.ViewTemplateId == src_template_id:
                    return view
            except Exception:
                continue

    return candidates[0]


def _collect_selected_sheet_views(sheets):
    view_rows = []
    seen_ids = set()
    for sheet in sheets:
        try:
            placed_views = list(sheet.GetAllPlacedViews())
        except Exception:
            placed_views = []
        for view_id in placed_views:
            try:
                if view_id.IntegerValue in seen_ids:
                    continue
                view = doc.GetElement(view_id)
                if view is None:
                    continue
                seen_ids.add(view_id.IntegerValue)
                view_rows.append(ViewRow(sheet, view, is_checked=True))
            except Exception:
                continue
    view_rows.sort(key=lambda row: (row.SheetNumber, row.ViewName.lower()))
    return view_rows


def _collect_link_rows(previous_selection=None):
    previous_selection = previous_selection or {}
    instance_counts = {}
    loaded_ids = set()
    try:
        instances = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()
        for instance in instances:
            try:
                type_id = instance.GetTypeId().IntegerValue
                instance_counts[type_id] = instance_counts.get(type_id, 0) + 1
                if instance.GetLinkDocument():
                    loaded_ids.add(type_id)
            except Exception:
                continue
    except Exception:
        pass

    rows = []
    try:
        link_types = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkType).ToElements()
    except Exception:
        link_types = []

    for link_type in link_types:
        try:
            key = link_type.Id.IntegerValue
            rows.append(LinkRow(
                link_type=link_type,
                loaded=(key in loaded_ids),
                instance_count=instance_counts.get(key, 0),
                is_checked=previous_selection.get(key, False),
            ))
        except Exception:
            continue

    rows.sort(key=lambda row: row.Name.lower())
    return rows


def _get_view_template_id(view):
    try:
        return view.ViewTemplateId
    except Exception:
        return DB.ElementId.InvalidElementId


def _get_view_template(view):
    template_id = _get_view_template_id(view)
    if not template_id or template_id == DB.ElementId.InvalidElementId:
        return None
    try:
        return doc.GetElement(template_id)
    except Exception:
        return None


def _template_controls_revit_links(view):
    template = _get_view_template(view)
    if template is None:
        return False

    try:
        non_controlled_params = template.GetNonControlledTemplateParameterIds()
        links_param_id = DB.ElementId(int(DB.BuiltInParameter.VIS_GRAPHICS_RVT_LINKS))
        return links_param_id not in non_controlled_params
    except Exception:
        return True


def _release_revit_links_template_control(view_rows):
    changed_templates = 0
    skipped_templates = []
    processed_ids = set()
    links_param_id = DB.ElementId(int(DB.BuiltInParameter.VIS_GRAPHICS_RVT_LINKS))

    for row in view_rows:
        view = getattr(row, "View", None)
        if view is None:
            continue
        template = _get_view_template(view)
        if template is None:
            continue
        try:
            template_int = template.Id.IntegerValue
        except Exception:
            continue
        if template_int in processed_ids:
            continue
        processed_ids.add(template_int)

        try:
            non_controlled = list(template.GetNonControlledTemplateParameterIds())
            if any(existing == links_param_id for existing in non_controlled):
                continue
            updated = List[DB.ElementId]()
            for existing in non_controlled:
                updated.Add(existing)
            updated.Add(links_param_id)
            template.SetNonControlledTemplateParameterIds(updated)
            changed_templates += 1
        except Exception as exc:
            try:
                skipped_templates.append("{} | {}".format(template.Name, exc))
            except Exception:
                skipped_templates.append(str(exc))

    return changed_templates, skipped_templates


def _resolve_target_views(view_rows, template_policy):
    targets = []
    template_ids = set()
    view_ids = set()

    for row in view_rows:
        view = row.View
        if view is None:
            continue
        template_id = _get_view_template_id(view)

        if template_policy == "Edit view templates" and template_id and template_id != DB.ElementId.InvalidElementId:
            try:
                if template_id.IntegerValue not in template_ids:
                    template = doc.GetElement(template_id)
                    if template:
                        template_ids.add(template_id.IntegerValue)
                        targets.append(template)
            except Exception:
                continue
            continue

        if template_policy == "Edit views unless template controls RVT Links":
            if template_id and template_id != DB.ElementId.InvalidElementId and _template_controls_revit_links(view):
                continue

        if template_policy == "Skip templated views" and template_id and template_id != DB.ElementId.InvalidElementId:
            continue

        try:
            if view.Id.IntegerValue not in view_ids:
                view_ids.add(view.Id.IntegerValue)
                targets.append(view)
        except Exception:
            continue

    return targets


def _analyze_target_resolution(view_rows, template_policy):
    total_checked = 0
    templated_views = 0
    untemplated_views = 0
    unique_template_ids = set()
    link_uncontrolled_template_views = 0
    link_controlled_template_views = 0

    for row in view_rows:
        view = getattr(row, "View", None)
        if view is None:
            continue
        total_checked += 1
        template_id = _get_view_template_id(view)

        has_template = bool(template_id and template_id != DB.ElementId.InvalidElementId)
        if has_template:
            templated_views += 1
            try:
                unique_template_ids.add(template_id.IntegerValue)
            except Exception:
                pass
            if _template_controls_revit_links(view):
                link_controlled_template_views += 1
            else:
                link_uncontrolled_template_views += 1
        else:
            untemplated_views += 1

    if total_checked == 0:
        reason = "No placed views are checked."
    elif template_policy == "Edit views unless template controls RVT Links" and (
        templated_views > 0 and link_controlled_template_views == templated_views and untemplated_views == 0
    ):
        reason = (
            "All checked views are templated, and every assigned template still controls RVT Links. "
            "Use Edit view templates, or uncheck RVT Links in the template controls first."
        )
    elif template_policy == "Skip templated views" and templated_views == total_checked:
        reason = (
            "All checked views are controlled by view templates, and the current Template Policy is Skip templated views. "
            "Use Edit view templates if you want to change those plans."
        )
    elif template_policy == "Edit view templates" and len(unique_template_ids) == 0:
        reason = (
            "None of the checked views has a view template assigned, so there are no templates to edit. "
            "Use Edit selected views instead."
        )
    else:
        reason = "No editable targets were derived from the current selection and policy."

    return {
        "total_checked": total_checked,
        "templated_views": templated_views,
        "untemplated_views": untemplated_views,
        "template_count": len(unique_template_ids),
        "link_uncontrolled_template_views": link_uncontrolled_template_views,
        "link_controlled_template_views": link_controlled_template_views,
        "reason": reason,
    }


def _get_grid_category():
    try:
        return doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Grids)
    except Exception:
        return None


def _try_invoke_reflection_method(target, methods, link_type, grid_category, hide):
    if target is None:
        return False, None

    for method in methods:
        try:
            name = method.Name
            parameters = list(method.GetParameters())
        except Exception:
            continue

        bool_primary = hide if ("hide" in name.lower() or "hidden" in name.lower()) else (not hide)

        value_sets = []
        for parameter in parameters:
            type_name = ""
            param_name = ""
            try:
                type_name = parameter.ParameterType.Name
            except Exception:
                type_name = ""
            try:
                param_name = parameter.Name.lower()
            except Exception:
                param_name = ""

            candidates = []
            if "Boolean" in type_name:
                candidates = [bool_primary, not bool_primary]
            elif "ElementId" in type_name:
                if "link" in param_name:
                    candidates = [link_type.Id]
                elif "cat" in param_name or "grid" in param_name:
                    candidates = [grid_category.Id]
                else:
                    candidates = [link_type.Id, grid_category.Id]
            elif "BuiltInCategory" in type_name:
                candidates = [DB.BuiltInCategory.OST_Grids]
            elif "Category" in type_name:
                candidates = [grid_category]
            elif "RevitLinkType" in type_name:
                candidates = [link_type]
            elif "Int32" in type_name:
                if "link" in param_name:
                    candidates = [link_type.Id.IntegerValue]
                elif "cat" in param_name or "grid" in param_name:
                    candidates = [grid_category.Id.IntegerValue]
                else:
                    candidates = [link_type.Id.IntegerValue, grid_category.Id.IntegerValue]
            else:
                if "link" in param_name:
                    candidates = [link_type.Id, link_type]
                elif "cat" in param_name or "grid" in param_name:
                    candidates = [grid_category.Id, grid_category, DB.BuiltInCategory.OST_Grids]
                else:
                    candidates = [grid_category.Id, link_type.Id, bool_primary]
            value_sets.append(candidates)

        def _attempt(index, current_args):
            if index >= len(value_sets):
                try:
                    method.Invoke(target, System.Array[System.Object](current_args))
                    return True
                except Exception:
                    return False
            for value in value_sets[index]:
                if _attempt(index + 1, current_args + [value]):
                    return True
            return False

        if _attempt(0, []):
            return True, name

    return False, None


def _apply_link_grid_visibility(view, link_type, hide):
    grid_category = _get_grid_category()
    if grid_category is None:
        return False, "Grid category not available"

    methods_to_try = []
    settings = None
    try:
        settings = view.GetLinkOverrides(link_type.Id)
    except Exception:
        settings = None

    if settings is not None:
        try:
            methods_to_try = [
                method for method in settings.GetType().GetMethods()
                if ("category" in method.Name.lower() or "grid" in method.Name.lower())
                and any(token in method.Name.lower() for token in ("hide", "hidden", "visible", "visibility"))
            ]
        except Exception:
            methods_to_try = []
        success, method_name = _try_invoke_reflection_method(settings, methods_to_try, link_type, grid_category, hide)
        if success:
            try:
                view.SetLinkOverrides(link_type.Id, settings)
                return True, "settings.{}".format(method_name)
            except Exception:
                pass

    view_methods = []
    try:
        view_methods = [
            method for method in view.GetType().GetMethods()
            if "link" in method.Name.lower()
            and ("category" in method.Name.lower() or "grid" in method.Name.lower())
            and any(token in method.Name.lower() for token in ("hide", "hidden", "visible", "visibility"))
        ]
    except Exception:
        view_methods = []

    success, method_name = _try_invoke_reflection_method(view, view_methods, link_type, grid_category, hide)
    if success:
        return True, "view.{}".format(method_name)

    return False, "Linked grid category API not exposed in this Revit build"


class ViewsLinksManager(forms.WPFWindow):
    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)

        self._full_tree_roots = []
        self._tree_roots = []
        self._root_index = {}
        self._all_sheets = []
        self._all_view_rows = []
        self._view_rows = []
        self._all_link_rows = []
        self._link_rows = []
        self._linked_view_choices = []
        self._phase_choices = []
        self._phase_filter_choices = []
        self._last_clicked_sheet_node = None
        self._panel_ratio = (1.05, 1.55, 1.2)
        self._is_custom_maximized = False
        self._restore_bounds = None

        self._apply_screen_constraints()
        self._init_static_controls()
        self._load_context()
        self._update_window_buttons()

    def _init_static_controls(self):
        self.UI_template_policy.ItemsSource = [
            "Edit selected views",
            "Edit views unless template controls RVT Links",
            "Edit view templates",
            "Skip templated views",
        ]
        self.UI_template_policy.SelectedIndex = 0

        self.UI_display_mode.ItemsSource = [
            "No change",
            "By host view",
            "By linked view",
            "Custom",
        ]
        self.UI_display_mode.SelectedIndex = 0

        self.UI_reference_link.ItemsSource = [
            ReferenceLinkChoice("<Select checked loaded links first>", mode="placeholder"),
        ]
        self.UI_reference_link.DisplayMemberPath = "Display"
        self.UI_reference_link.SelectedIndex = 0

        self.UI_linked_view.ItemsSource = [
            LinkedViewChoice("<Select checked links first>", "placeholder"),
        ]
        self.UI_linked_view.DisplayMemberPath = "Display"
        self.UI_linked_view.SelectedIndex = 0

        self.UI_halftone_mode.ItemsSource = [
            "No change",
            "On",
            "Off",
        ]
        self.UI_halftone_mode.SelectedIndex = 0

        self.UI_link_visibility_mode.ItemsSource = [
            "No change",
            "Show",
            "Hide",
        ]
        self.UI_link_visibility_mode.SelectedIndex = 0

        self.UI_grid_action.ItemsSource = [
            "No change",
            "Hide grids on selected links",
            "Show grids on selected links",
            "Hide grids on all links",
            "Show grids on all links",
        ]
        self.UI_grid_action.SelectedIndex = 0

        self.UI_object_styles_mode.ItemsSource = CUSTOM_LINK_VISIBILITY_OPTIONS
        self.UI_object_styles_mode.SelectedIndex = 0

        self.UI_view_filters_mode.ItemsSource = CUSTOM_LINK_VISIBILITY_OPTIONS
        self.UI_view_filters_mode.SelectedIndex = 0

        self.UI_view_range_mode.ItemsSource = LINK_VISIBILITY_OPTIONS
        self.UI_view_range_mode.SelectedIndex = 0

        self.UI_nested_links_mode.ItemsSource = LINK_VISIBILITY_OPTIONS
        self.UI_nested_links_mode.SelectedIndex = 0

        self.UI_color_fill_mode.ItemsSource = LINK_VISIBILITY_OPTIONS
        self.UI_color_fill_mode.SelectedIndex = 0

        self.UI_underlay_mode.ItemsSource = UNDERLAY_OPTIONS
        self.UI_underlay_mode.SelectedIndex = 0

        self.UI_worksets_mode.ItemsSource = CUSTOM_LINK_VISIBILITY_OPTIONS
        self.UI_worksets_mode.SelectedIndex = 0

        self.UI_discipline_mode.ItemsSource = DISCIPLINE_OPTIONS
        self.UI_discipline_mode.SelectedIndex = 0

        self.UI_detail_level_mode.ItemsSource = DETAIL_LEVEL_OPTIONS
        self.UI_detail_level_mode.SelectedIndex = 0

        self.UI_phase_mode.ItemsSource = PHASE_MODE_OPTIONS
        self.UI_phase_mode.SelectedIndex = 0

        self.UI_phase_choice.ItemsSource = [_placeholder_choice("<Select a reference link first>")]
        self.UI_phase_choice.DisplayMemberPath = "Display"
        self.UI_phase_choice.SelectedIndex = 0

        self.UI_phase_filter_mode.ItemsSource = PHASE_FILTER_MODE_OPTIONS
        self.UI_phase_filter_mode.SelectedIndex = 0

        self.UI_phase_filter_choice.ItemsSource = [_placeholder_choice("<Select a reference link first>")]
        self.UI_phase_filter_choice.DisplayMemberPath = "Display"
        self.UI_phase_filter_choice.SelectedIndex = 0

        self._update_display_controls()

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
        work_area = self._get_current_screen_work_area()
        if work_area is None:
            return
        try:
            self.MaxHeight = work_area.Height
            self.MaxWidth = work_area.Width
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
        work_area = self._get_current_screen_work_area()
        if work_area is None:
            return
        try:
            self.WindowState = System.Windows.WindowState.Normal
            self.Left = work_area.Left
            self.Top = work_area.Top
            self.Width = work_area.Width
            self.Height = work_area.Height
            self._is_custom_maximized = True
        except Exception:
            pass

    def _restore_from_custom_maximize(self):
        self._is_custom_maximized = False
        self._restore_window_bounds()

    def _capture_panel_ratio(self):
        try:
            widths = [
                float(self.UI_col_panel_1.ActualWidth),
                float(self.UI_col_panel_2.ActualWidth),
                float(self.UI_col_panel_3.ActualWidth),
            ]
            total = sum(width for width in widths if width > 0)
            if total > 0:
                self._panel_ratio = tuple(width / total for width in widths)
        except Exception:
            pass

    def _apply_panel_ratio(self):
        try:
            ratio1, ratio2, ratio3 = self._panel_ratio
            total = ratio1 + ratio2 + ratio3
            if total <= 0:
                return
            self.UI_col_panel_1.Width = System.Windows.GridLength(ratio1 / total, System.Windows.GridUnitType.Star)
            self.UI_col_panel_2.Width = System.Windows.GridLength(ratio2 / total, System.Windows.GridUnitType.Star)
            self.UI_col_panel_3.Width = System.Windows.GridLength(ratio3 / total, System.Windows.GridUnitType.Star)
        except Exception:
            pass

    def _update_window_buttons(self):
        try:
            self.UI_btn_maximize.Content = "o" if self._is_custom_maximized else "[]"
        except Exception:
            pass

    def _walk_visual_tree(self, root):
        if root is None:
            return
        yield root
        try:
            child_count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(root)
        except Exception:
            child_count = 0
        for index in range(child_count):
            try:
                child = System.Windows.Media.VisualTreeHelper.GetChild(root, index)
            except Exception:
                continue
            for descendant in self._walk_visual_tree(child):
                yield descendant

    def _apply_combo_theme(self):
        combo_boxes = [
            self.UI_template_policy,
            self.UI_display_mode,
            self.UI_reference_link,
            self.UI_linked_view,
            self.UI_halftone_mode,
            self.UI_link_visibility_mode,
            self.UI_grid_action,
            self.UI_object_styles_mode,
            self.UI_view_filters_mode,
            self.UI_view_range_mode,
            self.UI_nested_links_mode,
            self.UI_color_fill_mode,
            self.UI_underlay_mode,
            self.UI_worksets_mode,
            self.UI_discipline_mode,
            self.UI_detail_level_mode,
            self.UI_phase_mode,
            self.UI_phase_choice,
            self.UI_phase_filter_mode,
            self.UI_phase_filter_choice,
        ]
        for combo in combo_boxes:
            if combo is None:
                continue
            try:
                combo.ApplyTemplate()
                combo.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(238, 242, 214))
                combo.Foreground = System.Windows.Media.Brushes.Black
                combo.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 94, 168))
            except Exception:
                pass
            try:
                editable = combo.Template.FindName("PART_EditableTextBox", combo)
                if editable is not None:
                    editable.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(238, 242, 214))
                    editable.Foreground = System.Windows.Media.Brushes.Black
                    editable.CaretBrush = System.Windows.Media.Brushes.Black
                    editable.BorderThickness = System.Windows.Thickness(0)
                    editable.IsReadOnly = True
            except Exception:
                pass
            for child in self._walk_visual_tree(combo):
                try:
                    if isinstance(child, System.Windows.Controls.TextBox):
                        child.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(238, 242, 214))
                        child.Foreground = System.Windows.Media.Brushes.Black
                        child.CaretBrush = System.Windows.Media.Brushes.Black
                        child.BorderThickness = System.Windows.Thickness(0)
                        child.IsReadOnly = True
                    elif isinstance(child, System.Windows.Controls.ContentPresenter):
                        try:
                            child.SetValue(System.Windows.Controls.TextElement.ForegroundProperty, System.Windows.Media.Brushes.Black)
                        except Exception:
                            pass
                except Exception:
                    continue

    def _is_source_from_button(self, source):
        current = source
        while current is not None:
            try:
                if isinstance(current, System.Windows.Controls.Button):
                    return True
            except Exception:
                pass
            try:
                current = System.Windows.Media.VisualTreeHelper.GetParent(current)
            except Exception:
                break
        return False

    def _get_selected_loaded_link_rows(self):
        return [row for row in self._all_link_rows if row.IsChecked and row.IsLoaded]

    def _refresh_reference_link_choices(self):
        previous_choice = None
        try:
            selected = self.UI_reference_link.SelectedItem
            if selected is not None:
                previous_choice = str(selected)
        except Exception:
            previous_choice = None

        selected_link_rows = self._get_selected_loaded_link_rows()
        if not selected_link_rows:
            choices = [ReferenceLinkChoice("<Select checked loaded links first>", mode="placeholder")]
            self.UI_reference_link.ItemsSource = choices
            self.UI_reference_link.SelectedIndex = 0
            return

        ordered_rows = sorted(
            selected_link_rows,
            key=lambda row: (0 if row.IsArchitectural else 1, row.Name.lower())
        )
        choices = [ReferenceLinkChoice(row.Display, link_row=row) for row in ordered_rows]
        self.UI_reference_link.ItemsSource = choices

        restored = False
        if previous_choice:
            for choice in choices:
                if choice.Display == previous_choice:
                    self.UI_reference_link.SelectedItem = choice
                    restored = True
                    break
        if not restored and choices:
            self.UI_reference_link.SelectedIndex = 0

    def _collect_linked_view_choices(self, reference_link_choice):
        if not reference_link_choice or getattr(reference_link_choice, "Mode", "") == "placeholder":
            return [LinkedViewChoice("<Select a reference link first>", "placeholder")]

        link_row = reference_link_choice.LinkRow
        if not link_row:
            return [LinkedViewChoice("<Select a reference link first>", "placeholder")]

        link_doc = _find_link_doc_for_type(doc, link_row.LinkType.Id)
        if not link_doc:
            return [LinkedViewChoice("<Reference link is not loaded>", "placeholder")]

        views = []
        for view in _collect_link_views(link_doc):
            try:
                display = _safe_view_choice_display(view)
                view_name = _safe_element_name(view) or ""
                if display:
                    views.append(LinkedViewChoice(display, "named_view", view_name, view.Id))
            except Exception:
                continue

        if not views:
            return [LinkedViewChoice("<No linked views found>", "placeholder")]

        views.sort(key=lambda choice: choice.Display.lower())
        return views

    def _collect_phase_choices(self, reference_link_choice):
        if not reference_link_choice or getattr(reference_link_choice, "Mode", "") == "placeholder":
            return [_placeholder_choice("<Select a reference link first>")]

        link_row = reference_link_choice.LinkRow
        link_doc = _find_link_doc_for_type(doc, link_row.LinkType.Id) if link_row else None
        if not link_doc:
            return [_placeholder_choice("<Reference link is not loaded>")]

        items = []
        try:
            for phase in list(link_doc.Phases):
                try:
                    items.append(SettingChoice(phase.Name or "<Unnamed Phase>", phase.Name or ""))
                except Exception:
                    continue
        except Exception:
            items = []

        if not items:
            return [_placeholder_choice("<No phases found in reference link>")]

        items.sort(key=lambda choice: choice.Display.lower())
        return items

    def _collect_phase_filter_choices(self, reference_link_choice):
        if not reference_link_choice or getattr(reference_link_choice, "Mode", "") == "placeholder":
            return [_placeholder_choice("<Select a reference link first>")]

        link_row = reference_link_choice.LinkRow
        link_doc = _find_link_doc_for_type(doc, link_row.LinkType.Id) if link_row else None
        if not link_doc:
            return [_placeholder_choice("<Reference link is not loaded>")]

        items = []
        try:
            filters = DB.FilteredElementCollector(link_doc).OfClass(DB.PhaseFilter).ToElements()
        except Exception:
            filters = []
        for phase_filter in filters:
            try:
                items.append(SettingChoice(phase_filter.Name or "<Unnamed Phase Filter>", phase_filter.Name or ""))
            except Exception:
                continue

        if not items:
            return [_placeholder_choice("<No phase filters found in reference link>")]

        items.sort(key=lambda choice: choice.Display.lower())
        return items

    def _restore_choice_by_display(self, control, choices, previous_display):
        if previous_display:
            for choice in choices:
                try:
                    if choice.Display == previous_display:
                        control.SelectedItem = choice
                        return True
                except Exception:
                    continue
        if choices:
            control.SelectedIndex = 0
        return False

    def _refresh_phase_choices(self):
        previous_choice = None
        try:
            selected = self.UI_phase_choice.SelectedItem
            if selected is not None:
                previous_choice = str(selected)
        except Exception:
            previous_choice = None

        choices = self._collect_phase_choices(self.UI_reference_link.SelectedItem)
        self.UI_phase_choice.ItemsSource = choices
        self._phase_choices = choices
        self._restore_choice_by_display(self.UI_phase_choice, choices, previous_choice)

    def _refresh_phase_filter_choices(self):
        previous_choice = None
        try:
            selected = self.UI_phase_filter_choice.SelectedItem
            if selected is not None:
                previous_choice = str(selected)
        except Exception:
            previous_choice = None

        choices = self._collect_phase_filter_choices(self.UI_reference_link.SelectedItem)
        self.UI_phase_filter_choice.ItemsSource = choices
        self._phase_filter_choices = choices
        self._restore_choice_by_display(self.UI_phase_filter_choice, choices, previous_choice)

    def _refresh_linked_view_choices(self):
        previous_choice = None
        try:
            selected = self.UI_linked_view.SelectedItem
            if selected is not None:
                previous_choice = str(selected)
        except Exception:
            previous_choice = None

        reference_link_choice = self.UI_reference_link.SelectedItem
        choices = self._collect_linked_view_choices(reference_link_choice)
        self.UI_linked_view.ItemsSource = choices
        self._linked_view_choices = choices

        restored = False
        if previous_choice:
            for choice in choices:
                try:
                    if choice.Display == previous_choice:
                        self.UI_linked_view.SelectedItem = choice
                        restored = True
                        break
                except Exception:
                    continue
        if not restored and choices:
            self.UI_linked_view.SelectedIndex = 0

        self._update_display_controls()

    def _custom_controls_need_linked_view(self):
        return any([
            _current_setting_text(self.UI_object_styles_mode) == "By linked view",
            _current_setting_text(self.UI_view_filters_mode) == "By linked view",
            _current_setting_text(self.UI_view_range_mode) == "By linked view",
            _current_setting_text(self.UI_nested_links_mode) == "By linked view",
            _current_setting_text(self.UI_color_fill_mode) == "By linked view",
            _current_setting_text(self.UI_underlay_mode) == "By linked view",
            _current_setting_text(self.UI_worksets_mode) == "By linked view",
            _current_setting_text(self.UI_discipline_mode) == "By linked view",
            _current_setting_text(self.UI_detail_level_mode) == "By linked view",
            _current_setting_text(self.UI_phase_mode) == "By linked view",
            _current_setting_text(self.UI_phase_filter_mode) == "By linked view",
        ])

    def _custom_controls_have_changes(self):
        return any([
            _current_setting_text(self.UI_object_styles_mode) != "No change",
            _current_setting_text(self.UI_view_filters_mode) != "No change",
            _current_setting_text(self.UI_view_range_mode) != "No change",
            _current_setting_text(self.UI_nested_links_mode) != "No change",
            _current_setting_text(self.UI_color_fill_mode) != "No change",
            _current_setting_text(self.UI_underlay_mode) != "No change",
            _current_setting_text(self.UI_worksets_mode) != "No change",
            _current_setting_text(self.UI_discipline_mode) != "No change",
            _current_setting_text(self.UI_detail_level_mode) != "No change",
            _current_setting_text(self.UI_phase_mode) != "No change",
            _current_setting_text(self.UI_phase_filter_mode) != "No change",
        ])

    def _apply_link_visibility_property(self, settings, property_name, selection):
        if selection == "No change":
            return False, None
        if not hasattr(settings, property_name):
            return False, "{} is not available in this Revit build".format(property_name)

        target_value = _get_link_visibility_member(selection)
        if target_value is None:
            return False, "{} enum value is not available".format(selection)

        if selection == "By linked view":
            try:
                linked_view_id = settings.LinkedViewId
                if not linked_view_id or linked_view_id == DB.ElementId.InvalidElementId:
                    return False, "{} requires a linked view".format(property_name)
            except Exception:
                return False, "{} requires a linked view".format(property_name)

        try:
            setattr(settings, property_name, target_value)
            return True, None
        except Exception as exc:
            return False, "{} failed: {}".format(property_name, exc)

    def _apply_discipline_setting(self, settings, selection):
        if selection == "No change":
            return False, None
        if not hasattr(settings, "SetDiscipline"):
            return False, "Discipline is not available in this Revit build"

        explicit = _get_view_discipline_member(selection)
        if selection in ("By host view", "By linked view"):
            discipline_type = _get_link_visibility_member(selection)
            try:
                discipline = settings.GetDiscipline()
            except Exception:
                discipline = None
            if discipline is None:
                discipline = _get_view_discipline_member("Coordination")
        else:
            discipline_type = _get_link_visibility_member("Custom / None")
            discipline = explicit

        if discipline_type is None or discipline is None:
            return False, "Discipline enum values are not available"
        if selection == "By linked view":
            try:
                if settings.LinkedViewId == DB.ElementId.InvalidElementId:
                    return False, "Discipline requires a linked view"
            except Exception:
                return False, "Discipline requires a linked view"

        return _safe_invoke_method(settings, "SetDiscipline", [
            (discipline, discipline_type),
            (discipline_type, discipline),
        ])

    def _apply_detail_level_setting(self, settings, selection):
        if selection == "No change":
            return False, None
        if not hasattr(settings, "SetViewDetailLevel"):
            return False, "Detail Level is not available in this Revit build"

        if selection in ("By host view", "By linked view"):
            detail_type = _get_link_visibility_member(selection)
            try:
                detail_level = settings.GetViewDetailLevel()
            except Exception:
                detail_level = None
            if detail_level is None:
                detail_level = _get_view_detail_level_member("Medium")
        else:
            detail_type = _get_link_visibility_member("Custom / None")
            detail_level = _get_view_detail_level_member(selection)

        if detail_type is None or detail_level is None:
            return False, "Detail Level enum values are not available"
        if selection == "By linked view":
            try:
                if settings.LinkedViewId == DB.ElementId.InvalidElementId:
                    return False, "Detail Level requires a linked view"
            except Exception:
                return False, "Detail Level requires a linked view"

        return _safe_invoke_method(settings, "SetViewDetailLevel", [
            (detail_level, detail_type),
            (detail_type, detail_level),
        ])

    def _apply_phase_setting(self, settings, link_doc, mode_label, phase_choice):
        if mode_label == "No change":
            return False, None
        if not hasattr(settings, "SetPhase"):
            return False, "Phase is not available in this Revit build"

        if mode_label in ("By host view", "By linked view"):
            phase_type = _get_link_visibility_member(mode_label)
            try:
                phase_id = settings.GetPhaseId()
            except Exception:
                phase_id = DB.ElementId.InvalidElementId
        else:
            phase_type = _get_link_visibility_member("Custom / None")
            if phase_choice is None or getattr(phase_choice, "Mode", "") == "placeholder":
                return False, "Custom phase requires a phase selection"
            phase = _find_link_phase_by_name(link_doc, phase_choice.Value)
            if phase is None:
                return False, "Phase '{}' was not found in target link".format(phase_choice.Value)
            phase_id = phase.Id

        if phase_type is None:
            return False, "Phase type enum value is not available"
        if mode_label == "By linked view":
            try:
                if settings.LinkedViewId == DB.ElementId.InvalidElementId:
                    return False, "Phase requires a linked view"
            except Exception:
                return False, "Phase requires a linked view"

        return _safe_invoke_method(settings, "SetPhase", [
            (phase_id, phase_type),
            (phase_type, phase_id),
        ])

    def _apply_phase_filter_setting(self, settings, link_doc, mode_label, phase_filter_choice):
        if mode_label == "No change":
            return False, None
        if not hasattr(settings, "SetPhaseFilter"):
            return False, "Phase Filter is not available in this Revit build"

        if mode_label in ("By host view", "By linked view"):
            phase_filter_type = _get_link_visibility_member(mode_label)
            try:
                phase_filter_id = settings.GetPhaseFilterId()
            except Exception:
                phase_filter_id = DB.ElementId.InvalidElementId
        else:
            phase_filter_type = _get_link_visibility_member("Custom / None")
            if phase_filter_choice is None or getattr(phase_filter_choice, "Mode", "") == "placeholder":
                return False, "Custom phase filter requires a phase filter selection"
            phase_filter = _find_link_phase_filter_by_name(link_doc, phase_filter_choice.Value)
            if phase_filter is None:
                return False, "Phase Filter '{}' was not found in target link".format(phase_filter_choice.Value)
            phase_filter_id = phase_filter.Id

        if phase_filter_type is None:
            return False, "Phase Filter type enum value is not available"
        if mode_label == "By linked view":
            try:
                if settings.LinkedViewId == DB.ElementId.InvalidElementId:
                    return False, "Phase Filter requires a linked view"
            except Exception:
                return False, "Phase Filter requires a linked view"

        return _safe_invoke_method(settings, "SetPhaseFilter", [
            (phase_filter_id, phase_filter_type),
            (phase_filter_type, phase_filter_id),
        ])

    def _apply_custom_settings(self, settings, link_doc, phase_choice, phase_filter_choice):
        skip_messages = []
        applied_count = 0

        property_map = [
            ("ObjectStyles", _current_setting_text(self.UI_object_styles_mode)),
            ("ViewFilterType", _current_setting_text(self.UI_view_filters_mode)),
            ("ViewRange", _current_setting_text(self.UI_view_range_mode)),
            ("NestedLinks", _current_setting_text(self.UI_nested_links_mode)),
            ("ColorFill", _current_setting_text(self.UI_color_fill_mode)),
            ("Underlay", _current_setting_text(self.UI_underlay_mode)),
            ("Worksets", _current_setting_text(self.UI_worksets_mode)),
        ]

        for property_name, selection in property_map:
            success, message = self._apply_link_visibility_property(settings, property_name, selection)
            if success:
                applied_count += 1
            elif message:
                skip_messages.append(message)

        success, message = self._apply_discipline_setting(settings, _current_setting_text(self.UI_discipline_mode))
        if success:
            applied_count += 1
        elif message:
            skip_messages.append(message)

        success, message = self._apply_detail_level_setting(settings, _current_setting_text(self.UI_detail_level_mode))
        if success:
            applied_count += 1
        elif message:
            skip_messages.append(message)

        success, message = self._apply_phase_setting(settings, link_doc, _current_setting_text(self.UI_phase_mode), phase_choice)
        if success:
            applied_count += 1
        elif message:
            skip_messages.append(message)

        success, message = self._apply_phase_filter_setting(
            settings,
            link_doc,
            _current_setting_text(self.UI_phase_filter_mode),
            phase_filter_choice,
        )
        if success:
            applied_count += 1
        elif message:
            skip_messages.append(message)

        return applied_count, skip_messages

    def _update_display_controls(self):
        display_mode = str(self.UI_display_mode.SelectedItem or "No change")
        is_by_linked_view = (display_mode == "By linked view")
        uses_reference_link = display_mode in ("By linked view", "Custom")
        is_custom = (display_mode == "Custom")
        phase_custom = (_current_setting_text(self.UI_phase_mode) == "Custom phase")
        phase_filter_custom = (_current_setting_text(self.UI_phase_filter_mode) == "Custom phase filter")
        try:
            self.UI_reference_link.IsEnabled = uses_reference_link
            self.UI_reference_link_label.Opacity = 1.0 if uses_reference_link else 0.55
            self.UI_reference_link.Opacity = 1.0 if uses_reference_link else 0.55
            self.UI_linked_view.IsEnabled = uses_reference_link
            self.UI_linked_view_label.Opacity = 1.0 if uses_reference_link else 0.55
            self.UI_linked_view.Opacity = 1.0 if uses_reference_link else 0.55
            self.UI_adapt_level_views.IsEnabled = uses_reference_link
            self.UI_adapt_level_views.Opacity = 1.0 if uses_reference_link else 0.55
            self.UI_freeze_to_custom.IsEnabled = is_by_linked_view
            self.UI_freeze_to_custom.Opacity = 1.0 if is_by_linked_view else 0.55

            custom_controls = [
                self.UI_object_styles_mode,
                self.UI_view_filters_mode,
                self.UI_view_range_mode,
                self.UI_nested_links_mode,
                self.UI_color_fill_mode,
                self.UI_underlay_mode,
                self.UI_worksets_mode,
                self.UI_discipline_mode,
                self.UI_detail_level_mode,
                self.UI_phase_mode,
                self.UI_phase_filter_mode,
            ]
            for control in custom_controls:
                control.IsEnabled = is_custom
                control.Opacity = 1.0 if is_custom else 0.55

            self.UI_phase_choice.IsEnabled = is_custom and phase_custom
            self.UI_phase_choice.Opacity = 1.0 if (is_custom and phase_custom) else 0.55
            self.UI_phase_filter_choice.IsEnabled = is_custom and phase_filter_custom
            self.UI_phase_filter_choice.Opacity = 1.0 if (is_custom and phase_filter_custom) else 0.55
            self.UI_custom_settings_note.Opacity = 1.0 if is_custom else 0.55
        except Exception:
            pass

    def _load_context(self):
        self._load_sheets()
        self._apply_sheet_filter("")
        self._refresh_views()
        self._refresh_links()
        self._refresh_reference_link_choices()
        self._refresh_linked_view_choices()
        self._refresh_phase_choices()
        self._refresh_phase_filter_choices()
        self._update_counts()
        self._set_status("Context loaded.")

    def _load_sheets(self):
        previous_selection = set()
        for sheet in self._get_selected_sheets():
            try:
                previous_selection.add(sheet.Id.IntegerValue)
            except Exception:
                continue
        self._all_sheets = list(
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_Sheets)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        self._full_tree_roots, self._root_index = _build_sheet_tree(self._all_sheets, previous_selection)
        self._tree_roots = self._full_tree_roots
        self._sync_group_checkstates(self._full_tree_roots)
        self.UI_tree_sheets.ItemsSource = self._tree_roots

    def _get_selected_sheets(self):
        selected = []

        def _collect(nodes):
            for node in nodes:
                if node.IsSheet and node.Sheet and node.IsChecked:
                    selected.append(node.Sheet)
                if node.Children:
                    _collect(node.Children)

        _collect(self._full_tree_roots)
        return selected

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
            checked = sum(1 for sheet_node in descendant_sheets if sheet_node.IsChecked)
            if checked == 0:
                node.GroupCheckState = False
            elif checked == len(descendant_sheets):
                node.GroupCheckState = True
            else:
                node.GroupCheckState = None

    def _collect_descendant_sheets(self, node):
        sheet_nodes = []

        def _collect(children):
            for child in children:
                if child.IsSheet:
                    sheet_nodes.append(child)
                elif child.Children:
                    _collect(child.Children)

        _collect(node.Children)
        return sheet_nodes

    def _get_visible_sheet_nodes(self):
        visible = []

        def _collect(nodes):
            for node in nodes:
                if node.IsSheet:
                    visible.append(node)
                elif node.Children:
                    _collect(node.Children)

        _collect(self._tree_roots)
        return visible

    def _set_all_checked(self, nodes, value):
        for node in nodes:
            if node.IsSheet:
                node.IsChecked = value
            if node.Children:
                self._set_all_checked(node.Children, value)

    def _invert_checked(self, nodes):
        for node in nodes:
            if node.IsSheet:
                node.IsChecked = not bool(node.IsChecked)
            if node.Children:
                self._invert_checked(node.Children)

    def _set_expanded_for_groups(self, nodes, is_expanded):
        for node in nodes:
            if not node.IsSheet:
                node.IsExpanded = is_expanded
                if node.Children:
                    self._set_expanded_for_groups(node.Children, is_expanded)

    def _set_combo_to_text(self, combo, wanted_text):
        if combo is None:
            return False
        try:
            items = list(combo.ItemsSource) if combo.ItemsSource is not None else []
        except Exception:
            items = []
        for item in items:
            try:
                if str(item) == wanted_text:
                    combo.SelectedItem = item
                    return True
            except Exception:
                continue
        return False

    def _reset_custom_setting_controls(self):
        for control in [
            self.UI_object_styles_mode,
            self.UI_view_filters_mode,
            self.UI_view_range_mode,
            self.UI_nested_links_mode,
            self.UI_color_fill_mode,
            self.UI_underlay_mode,
            self.UI_worksets_mode,
            self.UI_discipline_mode,
            self.UI_detail_level_mode,
            self.UI_phase_mode,
            self.UI_phase_filter_mode,
        ]:
            self._set_combo_to_text(control, "No change")

    def _apply_typical_format_preset(self):
        self._set_combo_to_text(self.UI_template_policy, "Edit views unless template controls RVT Links")
        self._set_combo_to_text(self.UI_display_mode, "By linked view")
        self._set_combo_to_text(self.UI_halftone_mode, "No change")
        self._set_combo_to_text(self.UI_link_visibility_mode, "No change")
        self._set_combo_to_text(self.UI_grid_action, "No change")
        self._reset_custom_setting_controls()
        try:
            self.UI_adapt_level_views.IsChecked = True
            self.UI_freeze_to_custom.IsChecked = True
            self.UI_release_rvt_links_from_templates.IsChecked = True
        except Exception:
            pass
        self._update_display_controls()

    def _set_checked_for_group(self, group_node, value):
        if not group_node or group_node.IsSheet:
            return
        target_ids = set()

        def _collect_ids(node):
            for child in node.Children:
                if child.IsSheet and child.Sheet:
                    try:
                        target_ids.add(child.Sheet.Id.IntegerValue)
                    except Exception:
                        continue
                elif child.Children:
                    _collect_ids(child)

        _collect_ids(group_node)

        def _apply(nodes):
            for node in nodes:
                if node.IsSheet and node.Sheet:
                    try:
                        if node.Sheet.Id.IntegerValue in target_ids:
                            node.IsChecked = value
                    except Exception:
                        continue
                elif node.Children:
                    _apply(node.Children)

        _apply(self._full_tree_roots)

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
        for index in range(start_index, end_index + 1):
            visible[index].IsChecked = value

    def _normalize_checkbox_value(self, value):
        return True if value is None else bool(value)

    def _build_filtered_tree(self, nodes, search_text):
        filtered = []
        for node in nodes:
            if node.IsSheet:
                label = getattr(node, "SheetSearchText", node.NameLower)
                if search_text in label:
                    filtered.append(node)
            else:
                if search_text in node.NameLower:
                    filtered.append(node)
                else:
                    child_matches = self._build_filtered_tree(node.Children, search_text)
                    if child_matches:
                        clone = SheetTreeNode(node.Name, is_sheet=False)
                        clone.IsExpanded = True
                        clone.Children = child_matches
                        filtered.append(clone)
        return filtered

    def _apply_sheet_filter(self, raw_text):
        search_text = (raw_text or "").strip().lower()
        if not search_text or len(search_text) < 2:
            self._tree_roots = self._full_tree_roots
        else:
            self._tree_roots = self._build_filtered_tree(self._full_tree_roots, search_text)
        self._last_clicked_sheet_node = None
        self._sync_group_checkstates(self._tree_roots)
        self.UI_tree_sheets.ItemsSource = self._tree_roots
        self.UI_tree_sheets.Items.Refresh()
        self._update_counts()

    def _refresh_views(self):
        previous = {}
        for row in self._all_view_rows:
            try:
                previous[row.View.Id.IntegerValue] = row.IsChecked
            except Exception:
                continue
        fresh_rows = _collect_selected_sheet_views(self._get_selected_sheets())
        for row in fresh_rows:
            try:
                row.IsChecked = previous.get(row.View.Id.IntegerValue, True)
            except Exception:
                row.IsChecked = True
        self._all_view_rows = fresh_rows
        self._apply_view_filter(self.UI_view_filter.Text if self.UI_view_filter else "")

    def _apply_view_filter(self, raw_text):
        search_text = (raw_text or "").strip().lower()
        if not search_text:
            self._view_rows = self._all_view_rows
        else:
            self._view_rows = [row for row in self._all_view_rows if search_text in row.SearchText]
        self.UI_view_list.ItemsSource = self._view_rows
        self.UI_view_list.Items.Refresh()
        self._update_counts()

    def _refresh_links(self):
        previous = {}
        for row in self._all_link_rows:
            try:
                previous[row.LinkType.Id.IntegerValue] = row.IsChecked
            except Exception:
                continue
        self._all_link_rows = _collect_link_rows(previous)
        self._apply_link_filter(self.UI_link_filter.Text if self.UI_link_filter else "")
        self._refresh_reference_link_choices()
        self._refresh_linked_view_choices()
        self._refresh_phase_choices()
        self._refresh_phase_filter_choices()

    def _apply_link_filter(self, raw_text):
        search_text = (raw_text or "").strip().lower()
        if not search_text:
            self._link_rows = self._all_link_rows
        else:
            self._link_rows = [row for row in self._all_link_rows if search_text in row.SearchText]
        self.UI_link_list.ItemsSource = self._link_rows
        self.UI_link_list.Items.Refresh()
        self._update_counts()

    def _get_selected_view_rows(self):
        return [row for row in self._all_view_rows if row.IsChecked]

    def _get_selected_link_rows(self):
        return [row for row in self._all_link_rows if row.IsChecked]

    def _update_counts(self):
        sheet_total = len(self._get_selected_sheets())
        view_total = len(self._all_view_rows)
        view_selected = len(self._get_selected_view_rows())
        link_total = len(self._all_link_rows)
        link_selected = len(self._get_selected_link_rows())

        self.UI_sheet_count.Text = "({} selected)".format(sheet_total)
        self.UI_view_count.Text = "({} of {})".format(view_selected, view_total)
        self.UI_link_count.Text = "({} of {})".format(link_selected, link_total)
        self.UI_header_mode.Text = "Sheets: {} | Views: {} | Links: {}".format(
            sheet_total,
            view_selected,
            link_selected,
        )
        self.UI_info_summary.Text = (
            "Sheets: {} selected\n"
            "Views: {} checked\n"
            "Links: {} checked"
        ).format(sheet_total, view_selected, link_selected)

    def _append_log(self, text):
        try:
            current = self.UI_log.Text or ""
            self.UI_log.Text = "{}\n{}".format(current, text).strip()
            self.UI_log.ScrollToEnd()
        except Exception:
            pass

    def _set_status(self, text):
        self.UI_status_bar.Text = text

    def _resolve_linked_view(self, view, link_type, settings, reference_link_choice, linked_view_choice, adapt_level_views):
        link_doc = _find_link_doc_for_type(doc, link_type.Id)
        if not link_doc:
            return None, "Link document is not loaded"

        if not linked_view_choice or getattr(linked_view_choice, "Mode", "") == "placeholder":
            return None, "No linked view choice selected"

        if not reference_link_choice or getattr(reference_link_choice, "Mode", "") == "placeholder":
            return None, "No reference link selected"

        reference_row = reference_link_choice.LinkRow
        if not reference_row:
            return None, "Reference link row not available"

        reference_doc = _find_link_doc_for_type(doc, reference_row.LinkType.Id)
        if not reference_doc:
            return None, "Reference link document is not loaded"

        if linked_view_choice.Mode == "named_view":
            source_view = _find_link_view_by_id(reference_doc, getattr(linked_view_choice, "ViewId", None))
            if source_view is None:
                source_view = _find_link_view_by_name(reference_doc, linked_view_choice.ViewName)
            if source_view is None:
                return None, "Selected linked view name not found in reference link"

            if adapt_level_views:
                dst_level = _safe_get_view_level(view)
                fallback_viewtype = None
                try:
                    fallback_viewtype = view.ViewType
                except Exception:
                    fallback_viewtype = None
                linked_view = _choose_linked_view_for_level(link_doc, source_view, dst_level, fallback_viewtype)
                if linked_view:
                    return linked_view, None

            linked_view = _find_link_view_by_name(link_doc, linked_view_choice.ViewName)
            return linked_view, None if linked_view else "Selected linked view name not found in target link"

        src_link_view = None
        try:
            if settings.LinkedViewId and settings.LinkedViewId != DB.ElementId.InvalidElementId:
                src_link_view = link_doc.GetElement(settings.LinkedViewId)
        except Exception:
            src_link_view = None

        dst_level = _safe_get_view_level(view)
        fallback_viewtype = None
        try:
            fallback_viewtype = view.ViewType
        except Exception:
            fallback_viewtype = None

        linked_view = _choose_linked_view_for_level(link_doc, src_link_view, dst_level, fallback_viewtype)
        if linked_view:
            return linked_view, None
        return None, "No matching linked view found for target level"

    def _apply_to_targets(self):
        selected_view_rows = self._get_selected_view_rows()
        if not selected_view_rows:
            forms.alert("Check at least one placed view first.", title=__title__)
            return

        display_mode = str(self.UI_display_mode.SelectedItem or "No change")
        reference_link_choice = self.UI_reference_link.SelectedItem
        linked_view_choice = self.UI_linked_view.SelectedItem
        phase_choice = self.UI_phase_choice.SelectedItem
        phase_filter_choice = self.UI_phase_filter_choice.SelectedItem
        template_policy = str(self.UI_template_policy.SelectedItem or "Edit selected views")
        halftone_mode = str(self.UI_halftone_mode.SelectedItem or "No change")
        link_visibility_mode = str(self.UI_link_visibility_mode.SelectedItem or "No change")
        grid_action = str(self.UI_grid_action.SelectedItem or "No change")
        freeze_to_custom = bool(getattr(self.UI_freeze_to_custom, "IsChecked", False))
        adapt_level_views = bool(getattr(self.UI_adapt_level_views, "IsChecked", False))
        release_template_link_control = bool(getattr(self.UI_release_rvt_links_from_templates, "IsChecked", False))

        if display_mode == "By linked view" and (
            reference_link_choice is None or getattr(reference_link_choice, "Mode", "") == "placeholder"
        ):
            forms.alert("Choose a reference link first.", title=__title__)
            return

        if (display_mode == "By linked view" or (display_mode == "Custom" and self._custom_controls_need_linked_view())) and (
            linked_view_choice is None or getattr(linked_view_choice, "Mode", "") == "placeholder"
        ):
            forms.alert("Choose a linked view option from the Linked view list first.", title=__title__)
            return

        if (display_mode == "By linked view" or (display_mode == "Custom" and self._custom_controls_need_linked_view())) and (
            reference_link_choice is None or getattr(reference_link_choice, "Mode", "") == "placeholder"
        ):
            forms.alert("Choose a reference link first.", title=__title__)
            return

        if display_mode == "Custom" and _current_setting_text(self.UI_phase_mode) == "Custom phase" and (
            phase_choice is None or getattr(phase_choice, "Mode", "") == "placeholder"
        ):
            forms.alert("Choose a phase from the reference link first.", title=__title__)
            return

        if display_mode == "Custom" and _current_setting_text(self.UI_phase_filter_mode) == "Custom phase filter" and (
            phase_filter_choice is None or getattr(phase_filter_choice, "Mode", "") == "placeholder"
        ):
            forms.alert("Choose a phase filter from the reference link first.", title=__title__)
            return

        selected_link_rows = self._get_selected_link_rows()
        if not selected_link_rows and not ("all links" in grid_action.lower()):
            forms.alert("Select at least one link first.", title=__title__)
            return

        template_release_count = 0
        template_release_skips = []

        self.UI_log.Text = "Applying batch update..."

        if release_template_link_control:
            with revit.Transaction("Release RVT Links template control"):
                template_release_count, template_release_skips = _release_revit_links_template_control(selected_view_rows)

        targets = _resolve_target_views(selected_view_rows, template_policy)
        if not targets:
            diagnostic = _analyze_target_resolution(selected_view_rows, template_policy)
            message = (
                "No editable target views were resolved.\n\n"
                "Template Policy: {}\n"
                "Checked views: {}\n"
                "Templated views: {}\n"
                "Non-templated views: {}\n"
                "Templated views with RVT Links free: {}\n"
                "Templated views with RVT Links controlled: {}\n"
                "Unique templates found: {}\n\n"
                "Reason: {}"
            ).format(
                template_policy,
                diagnostic["total_checked"],
                diagnostic["templated_views"],
                diagnostic["untemplated_views"],
                diagnostic["link_uncontrolled_template_views"],
                diagnostic["link_controlled_template_views"],
                diagnostic["template_count"],
                diagnostic["reason"],
            )
            self.UI_feedback.Text = diagnostic["reason"]
            self._set_status(diagnostic["reason"])
            self._append_log(message)
            forms.alert(message, title=__title__)
            return

        if "all links" in grid_action.lower():
            link_rows_for_grid = list(self._all_link_rows)
        else:
            link_rows_for_grid = list(selected_link_rows)

        changed_overrides = 0
        changed_visibility = 0
        changed_grids = 0
        grid_skipped = 0
        failed = 0
        frozen_to_custom = 0
        missing_link_views = []
        custom_applied = 0
        custom_skips = set()
        grid_support_details = set()
        failure_samples = []

        def _record_failure(operation_name, view_name, link_name, exc):
            if len(failure_samples) >= 18:
                return
            try:
                message = str(exc)
            except Exception:
                message = "Unknown error"
            failure_samples.append("{} | {} | {} | {}".format(operation_name, view_name, link_name, message))

        with revit.Transaction("Views Links Manager"):
            for target_view in targets:
                for link_row in selected_link_rows:
                    link_type = link_row.LinkType
                    link_doc = _find_link_doc_for_type(doc, link_type.Id)
                    try:
                        settings = target_view.GetLinkOverrides(link_type.Id)
                    except Exception:
                        settings = None

                    try:
                        if settings and display_mode != "No change":
                            link_visibility_enum = getattr(DB, "LinkVisibility", None)
                            if link_visibility_enum:
                                if display_mode == "By host view":
                                    settings.LinkVisibilityType = link_visibility_enum.ByHostView
                                elif display_mode == "By linked view":
                                    settings.LinkVisibilityType = link_visibility_enum.ByLinkedView
                                    linked_view, error_message = self._resolve_linked_view(
                                        target_view,
                                        link_type,
                                        settings,
                                        reference_link_choice,
                                        linked_view_choice,
                                        adapt_level_views,
                                    )
                                    if linked_view:
                                        settings.LinkedViewId = linked_view.Id
                                    else:
                                        missing_link_views.append(
                                            "{} -> {} ({})".format(
                                                target_view.Name,
                                                _safe_link_type_name(link_type),
                                                error_message,
                                            )
                                        )
                                elif display_mode == "Custom":
                                    settings.LinkVisibilityType = link_visibility_enum.Custom

                        if settings and display_mode == "Custom" and self._custom_controls_have_changes():
                            applied_count, skip_messages = self._apply_custom_settings(
                                settings,
                                link_doc,
                                phase_choice,
                                phase_filter_choice,
                            )
                            custom_applied += applied_count
                            for message in skip_messages:
                                if message:
                                    custom_skips.add(message)

                        if settings and halftone_mode != "No change" and hasattr(settings, "Halftone"):
                            settings.Halftone = (halftone_mode == "On")

                        if settings and (
                            display_mode != "No change"
                            or halftone_mode != "No change"
                            or (display_mode == "Custom" and self._custom_controls_have_changes())
                        ):
                            target_view.SetLinkOverrides(link_type.Id, settings)
                            changed_overrides += 1

                            if display_mode == "By linked view" and freeze_to_custom and link_visibility_enum:
                                try:
                                    frozen_settings = target_view.GetLinkOverrides(link_type.Id)
                                except Exception:
                                    frozen_settings = None
                                if frozen_settings:
                                    try:
                                        frozen_settings.LinkVisibilityType = link_visibility_enum.Custom
                                        target_view.SetLinkOverrides(link_type.Id, frozen_settings)
                                        frozen_to_custom += 1
                                    except Exception as exc:
                                        _record_failure("freeze_to_custom", target_view.Name, _safe_link_type_name(link_type), exc)
                                        failed += 1
                    except Exception as exc:
                        _record_failure("link_overrides", target_view.Name, _safe_link_type_name(link_type), exc)
                        failed += 1

                    try:
                        if link_visibility_mode == "Hide" and link_type.CanBeHidden(target_view):
                            if not link_type.IsHidden(target_view):
                                target_view.HideElements(List[DB.ElementId]([link_type.Id]))
                                changed_visibility += 1
                        elif link_visibility_mode == "Show" and link_type.CanBeHidden(target_view):
                            if link_type.IsHidden(target_view):
                                target_view.UnhideElements(List[DB.ElementId]([link_type.Id]))
                                changed_visibility += 1
                    except Exception as exc:
                        _record_failure("link_visibility", target_view.Name, _safe_link_type_name(link_type), exc)
                        failed += 1

                if grid_action != "No change":
                    hide_grids = grid_action.startswith("Hide")
                    for link_row in link_rows_for_grid:
                        try:
                            success, detail = _apply_link_grid_visibility(target_view, link_row.LinkType, hide_grids)
                            if success:
                                changed_grids += 1
                                if detail:
                                    grid_support_details.add(detail)
                            else:
                                grid_skipped += 1
                        except Exception as exc:
                            _record_failure("grid_action", target_view.Name, _safe_link_type_name(link_row.LinkType), exc)
                            failed += 1

        summary = (
            "Updated {} target views/templates. Templates released for RVT Links: {}. Link overrides: {}. Custom subsettings: {}. Frozen to Custom: {}. Link hide/unhide: {}. "
            "Grid actions: {}. Grid skipped: {}. Failures: {}."
        ).format(len(targets), template_release_count, changed_overrides, custom_applied, frozen_to_custom, changed_visibility, changed_grids, grid_skipped, failed)
        self.UI_feedback.Text = summary
        self._set_status(summary)
        self._append_log(summary)

        if template_release_skips:
            self._append_log("Template release notes:\n{}".format("\n".join(sorted(set(template_release_skips)))))
        if grid_support_details:
            self._append_log("Grid API route: {}".format(", ".join(sorted(grid_support_details))))
        if custom_skips:
            self._append_log("Custom setting notes:\n{}".format("\n".join(sorted(custom_skips))))
        if missing_link_views:
            self._append_log("Missing linked views:\n{}".format("\n".join(sorted(set(missing_link_views)))))
        if failure_samples:
            self._append_log("Failure samples:\n{}".format("\n".join(failure_samples)))

    def header_drag(self, sender, event_args):
        try:
            original = getattr(event_args, "OriginalSource", None)
            if self._is_source_from_button(original):
                return
        except Exception:
            pass

        try:
            click_count = getattr(event_args, "ClickCount", 1)
        except Exception:
            click_count = 1

        if click_count >= 2:
            self.button_maximize(sender, event_args)
            return

        try:
            self.DragMove()
        except Exception:
            pass

    def window_loaded(self, sender, event_args):
        self._apply_panel_ratio()
        self._apply_combo_theme()
        self._update_display_controls()

    def button_minimize(self, sender, event_args):
        try:
            self.WindowState = System.Windows.WindowState.Minimized
        except Exception:
            pass

    def button_maximize(self, sender, event_args):
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

    def button_close(self, sender, event_args):
        try:
            self.Close()
        except Exception:
            pass

    def panel_splitter_drag_completed(self, sender, event_args):
        self._capture_panel_ratio()

    def display_mode_changed(self, sender, event_args):
        self._update_display_controls()

    def reference_link_changed(self, sender, event_args):
        self._refresh_linked_view_choices()
        self._refresh_phase_choices()
        self._refresh_phase_filter_choices()

    def custom_setting_changed(self, sender, event_args):
        self._update_display_controls()

    def button_reset_layout(self, sender, event_args):
        self._panel_ratio = (1.05, 1.55, 1.2)
        self._apply_panel_ratio()
        self._set_status("Panel layout reset.")

    def button_refresh_context(self, sender, event_args):
        self._load_context()

    def search_textbox_changed(self, sender, event_args):
        self._apply_sheet_filter(self.UI_search.Text if self.UI_search else "")

    def button_clear_search(self, sender, event_args):
        self.UI_search.Text = ""
        self._apply_sheet_filter("")

    def sheet_checkbox_click(self, sender, event_args):
        node = getattr(sender, "Tag", None)
        if not node:
            return
        target_value = self._normalize_checkbox_value(getattr(sender, "IsChecked", False))
        modifiers = Control.ModifierKeys
        is_shift = (modifiers & Keys.Shift) == Keys.Shift
        is_alt = (modifiers & Keys.Alt) == Keys.Alt

        if node.IsSheet:
            if is_alt:
                for visible_node in self._get_visible_sheet_nodes():
                    visible_node.IsChecked = target_value
            elif is_shift and self._last_clicked_sheet_node:
                self._set_checked_for_range(self._last_clicked_sheet_node, node, target_value)
            else:
                node.IsChecked = target_value
            self._last_clicked_sheet_node = node
        else:
            if is_alt:
                for visible_node in self._get_visible_sheet_nodes():
                    visible_node.IsChecked = target_value
            else:
                self._set_checked_for_group(node, target_value)

        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self.UI_tree_sheets.Items.Refresh()
        self._refresh_views()
        self._update_counts()
        self._set_status("Sheet selection updated.")

    def button_check_all_sheets(self, sender, event_args):
        self._set_all_checked(self._tree_roots, True)
        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self.UI_tree_sheets.Items.Refresh()
        self._refresh_views()
        self._update_counts()

    def button_uncheck_all_sheets(self, sender, event_args):
        self._set_all_checked(self._tree_roots, False)
        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self.UI_tree_sheets.Items.Refresh()
        self._refresh_views()
        self._update_counts()

    def button_invert_shown_sheets(self, sender, event_args):
        self._invert_checked(self._tree_roots)
        self._sync_group_checkstates(self._full_tree_roots)
        self._sync_group_checkstates(self._tree_roots)
        self.UI_tree_sheets.Items.Refresh()
        self._refresh_views()
        self._update_counts()
        self._set_status("Shown sheets inverted.")

    def button_expand_all(self, sender, event_args):
        self._set_expanded_for_groups(self._full_tree_roots, True)
        self._set_expanded_for_groups(self._tree_roots, True)
        self.UI_tree_sheets.Items.Refresh()

    def button_collapse_all(self, sender, event_args):
        self._set_expanded_for_groups(self._full_tree_roots, False)
        self._set_expanded_for_groups(self._tree_roots, False)
        self.UI_tree_sheets.Items.Refresh()

    def view_filter_text_changed(self, sender, event_args):
        self._apply_view_filter(self.UI_view_filter.Text if self.UI_view_filter else "")

    def view_checkbox_click(self, sender, event_args):
        row = getattr(sender, "Tag", None)
        if not row:
            return
        row.IsChecked = self._normalize_checkbox_value(getattr(sender, "IsChecked", False))
        self._update_counts()

    def button_check_all_views(self, sender, event_args):
        for row in self._view_rows:
            row.IsChecked = True
        self.UI_view_list.Items.Refresh()
        self._update_counts()

    def button_uncheck_all_views(self, sender, event_args):
        for row in self._view_rows:
            row.IsChecked = False
        self.UI_view_list.Items.Refresh()
        self._update_counts()

    def button_invert_shown_views(self, sender, event_args):
        for row in self._view_rows:
            row.IsChecked = not bool(row.IsChecked)
        self.UI_view_list.Items.Refresh()
        self._update_counts()
        self._set_status("Shown views inverted.")

    def link_filter_text_changed(self, sender, event_args):
        self._apply_link_filter(self.UI_link_filter.Text if self.UI_link_filter else "")

    def link_checkbox_click(self, sender, event_args):
        row = getattr(sender, "Tag", None)
        if not row:
            return
        row.IsChecked = self._normalize_checkbox_value(getattr(sender, "IsChecked", False))
        self._refresh_reference_link_choices()
        self._refresh_linked_view_choices()
        self._update_counts()

    def button_select_arch_links(self, sender, event_args):
        for row in self._all_link_rows:
            row.IsChecked = bool(row.IsArchitectural)
        self.UI_link_list.Items.Refresh()
        self._refresh_reference_link_choices()
        self._refresh_linked_view_choices()
        self._update_counts()
        self._set_status("Architectural links selected.")

    def button_select_all_links(self, sender, event_args):
        for row in self._link_rows:
            row.IsChecked = True
        self.UI_link_list.Items.Refresh()
        self._refresh_reference_link_choices()
        self._refresh_linked_view_choices()
        self._update_counts()

    def button_clear_link_selection(self, sender, event_args):
        for row in self._all_link_rows:
            row.IsChecked = False
        self.UI_link_list.Items.Refresh()
        self._refresh_reference_link_choices()
        self._refresh_linked_view_choices()
        self._update_counts()

    def button_apply(self, sender, event_args):
        self._apply_to_targets()

    def button_apply_typical_format(self, sender, event_args):
        self._apply_typical_format_preset()
        self._set_status("Typical sheet format preset loaded. Applying to current selection...")
        self._apply_to_targets()


if __name__ == "__main__":
    if not doc:
        forms.alert("Open a Revit model first.", title=__title__)
        script.exit()

    window = ViewsLinksManager("ViewsLinksManager.xaml")
    window.ShowDialog()