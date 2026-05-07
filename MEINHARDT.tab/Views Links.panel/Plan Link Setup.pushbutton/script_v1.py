# -*- coding: utf-8 -*-
__title__ = "Plan Link Setup"
__doc__ = """Version = 2.0
Date: 2026-03-26
Author: GM
Description:
Batch setup for linked room display across selected placed views.
How-to:
1. Select sheets in the center panel.
2. Confirm the placed views to update.
3. Pick the source link and linked view.
4. Apply the setup.
5. The tool can finish the selected link as Custom and hide non-room linked annotation.
"""

import clr
clr.AddReference("System")
clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Windows.Forms")

import System
from System.Collections.Generic import List
from System.Windows.Forms import Clipboard, Control, Keys

from pyrevit import DB, forms, revit


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
        self.ViewType = _safe_view_type_name(view)
        self.LevelName = _safe_level_name(view)
        self.TemplateName = _safe_template_name(view)
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


class LinkRow(object):
    def __init__(self, link_type, loaded, instance_count):
        self.LinkType = link_type
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

    def __str__(self):
        return self.Display


class LinkedViewOption(object):
    def __init__(self, view):
        self.View = view
        self.Name = getattr(view, "Name", "<Unnamed View>") or "<Unnamed View>"
        self.LevelName = _safe_level_name(view)
        self.ViewTypeName = _safe_view_type_name(view)
        self.Display = "{} | {} | {}".format(self.Name, self.LevelName, self.ViewTypeName)

    def __str__(self):
        return self.Display


def _safe_view_type_name(view):
    try:
        return str(view.ViewType)
    except Exception:
        return "Unknown"


def _safe_template_name(view):
    try:
        template_id = view.ViewTemplateId
        if template_id and template_id != DB.ElementId.InvalidElementId:
            template = doc.GetElement(template_id)
            if template:
                return template.Name or ""
    except Exception:
        pass
    return ""


def _normalize_label(text):
    return "".join(ch for ch in (text or "").lower() if ch.isalnum())


def _clean_group_value(value):
    text = (value or "").strip()
    if not text or text.lower() in ("<none>", "none"):
        return ""
    return text


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
    return any(token in normalized for token in ("arch", "architect", "arc", "a-model", "a_"))


def _safe_get_view_level(view):
    try:
        if hasattr(view, "GenLevel"):
            return view.GenLevel
    except Exception:
        return None
    return None


def _safe_level_name(view):
    try:
        level = _safe_get_view_level(view)
        if level:
            return level.Name
    except Exception:
        pass
    return "No Level"


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
    has_any_sheet_collection = any(_get_sheet_param_text(sheet, ["Sheet Collection"]) for sheet in all_sheets)

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
        try:
            sheet_node.IsChecked = sheet.Id.IntegerValue in preselected_ids
        except Exception:
            sheet_node.IsChecked = False
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
    return roots


def _find_sheets_for_view(view_id):
    matched = set()
    try:
        sheets = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
    except Exception:
        sheets = []
    for sheet in sheets:
        try:
            for placed_id in list(sheet.GetAllPlacedViews()):
                if placed_id == view_id:
                    matched.add(sheet.Id.IntegerValue)
                    break
        except Exception:
            continue
    return matched


def _get_default_sheet_ids():
    selected = set()
    active_view = revit.active_view
    if active_view is None:
        return selected
    try:
        if active_view.ViewType == DB.ViewType.DrawingSheet:
            selected.add(active_view.Id.IntegerValue)
            return selected
    except Exception:
        pass
    try:
        selected.update(_find_sheets_for_view(active_view.Id))
    except Exception:
        pass
    return selected


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


def _collect_link_views(link_doc):
    views = []
    try:
        for view in DB.FilteredElementCollector(link_doc).OfClass(DB.View).ToElements():
            try:
                if view.IsTemplate:
                    continue
                if hasattr(view, "CanBePrinted") and not view.CanBePrinted:
                    continue
                views.append(view)
            except Exception:
                continue
    except Exception:
        return []
    return views


def _best_match_link_level(link_doc, host_level):
    if not link_doc or not host_level:
        return None
    try:
        link_levels = DB.FilteredElementCollector(link_doc).OfClass(DB.Level).ToElements()
    except Exception:
        link_levels = []
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


def _collect_visible_link_options(view):
    link_data = {}
    try:
        instances = DB.FilteredElementCollector(doc, view.Id).OfClass(DB.RevitLinkInstance).ToElements()
    except Exception:
        instances = []

    for instance in instances:
        try:
            link_type_id = instance.GetTypeId()
            link_type = doc.GetElement(link_type_id)
            if link_type is None:
                continue
            data = link_data.setdefault(link_type_id.IntegerValue, {"type": link_type, "count": 0, "loaded": False})
            data["count"] += 1
            try:
                if instance.GetLinkDocument() is not None:
                    data["loaded"] = True
            except Exception:
                pass
        except Exception:
            continue

    options = [LinkRow(item["type"], item["loaded"], item["count"]) for item in link_data.values()]
    options.sort(key=lambda item: (0 if item.IsArchitectural else 1, item.Name.lower()))
    return options


def _collect_link_rows_for_views(view_rows):
    collected = {}
    source_views = [row.View for row in view_rows if getattr(row, "View", None) is not None]
    if not source_views and revit.active_view is not None:
        source_views = [revit.active_view]

    for view in source_views:
        for row in _collect_visible_link_options(view):
            try:
                key = row.LinkType.Id.IntegerValue
            except Exception:
                continue
            existing = collected.get(key)
            if existing is None:
                collected[key] = LinkRow(row.LinkType, row.IsLoaded, row.InstanceCount)
            else:
                existing.InstanceCount = max(existing.InstanceCount, row.InstanceCount)
                existing.IsLoaded = existing.IsLoaded or row.IsLoaded

    rows = list(collected.values())
    rows.sort(key=lambda item: (0 if item.IsArchitectural else 1, item.Name.lower()))
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
    elif template_policy == "Edit views unless template controls RVT Links" and templated_views > 0 and link_controlled_template_views == templated_views and untemplated_views == 0:
        reason = (
            "All checked views are templated, and every assigned template still controls RVT Links. "
            "Use Edit view templates, or release RVT Links from templates first."
        )
    elif template_policy == "Skip templated views" and templated_views == total_checked:
        reason = "All checked views are controlled by view templates, and the current Template Policy skips templated views."
    elif template_policy == "Edit view templates" and len(unique_template_ids) == 0:
        reason = "None of the checked views has a view template assigned."
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
    return _get_enum_member(view_discipline_enum, ("Coordination",))


def _get_view_detail_level_member(mode_label):
    view_detail_level_enum = getattr(DB, "ViewDetailLevel", None)
    if view_detail_level_enum is None:
        return None
    return _get_enum_member(view_detail_level_enum, ("Medium",))


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


def _apply_link_visibility_property(settings, property_name, selection):
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


def _apply_discipline_setting(settings, selection):
    if not hasattr(settings, "SetDiscipline"):
        return False, "Discipline is not available in this Revit build"
    discipline_type = _get_link_visibility_member(selection)
    if selection == "By linked view":
        try:
            if settings.LinkedViewId == DB.ElementId.InvalidElementId:
                return False, "Discipline requires a linked view"
        except Exception:
            return False, "Discipline requires a linked view"
    try:
        discipline = settings.GetDiscipline()
    except Exception:
        discipline = _get_view_discipline_member("Coordination")
    if discipline_type is None or discipline is None:
        return False, "Discipline enum values are not available"
    return _safe_invoke_method(settings, "SetDiscipline", [(discipline, discipline_type), (discipline_type, discipline)])


def _apply_detail_level_setting(settings, selection):
    if not hasattr(settings, "SetViewDetailLevel"):
        return False, "Detail Level is not available in this Revit build"
    detail_type = _get_link_visibility_member(selection)
    if selection == "By linked view":
        try:
            if settings.LinkedViewId == DB.ElementId.InvalidElementId:
                return False, "Detail Level requires a linked view"
        except Exception:
            return False, "Detail Level requires a linked view"
    try:
        detail_level = settings.GetViewDetailLevel()
    except Exception:
        detail_level = _get_view_detail_level_member("Medium")
    if detail_type is None or detail_level is None:
        return False, "Detail Level enum values are not available"
    return _safe_invoke_method(settings, "SetViewDetailLevel", [(detail_level, detail_type), (detail_type, detail_level)])


def _apply_phase_setting(settings, selection):
    if not hasattr(settings, "SetPhase"):
        return False, "Phase is not available in this Revit build"
    phase_type = _get_link_visibility_member(selection)
    if phase_type is None:
        return False, "Phase type enum value is not available"
    if selection == "By linked view":
        try:
            if settings.LinkedViewId == DB.ElementId.InvalidElementId:
                return False, "Phase requires a linked view"
        except Exception:
            return False, "Phase requires a linked view"
    try:
        phase_id = settings.GetPhaseId()
    except Exception:
        phase_id = DB.ElementId.InvalidElementId
    return _safe_invoke_method(settings, "SetPhase", [(phase_id, phase_type), (phase_type, phase_id)])


def _apply_phase_filter_setting(settings, selection):
    if not hasattr(settings, "SetPhaseFilter"):
        return False, "Phase Filter is not available in this Revit build"
    phase_filter_type = _get_link_visibility_member(selection)
    if phase_filter_type is None:
        return False, "Phase Filter type enum value is not available"
    if selection == "By linked view":
        try:
            if settings.LinkedViewId == DB.ElementId.InvalidElementId:
                return False, "Phase Filter requires a linked view"
        except Exception:
            return False, "Phase Filter requires a linked view"
    try:
        phase_filter_id = settings.GetPhaseFilterId()
    except Exception:
        phase_filter_id = DB.ElementId.InvalidElementId
    return _safe_invoke_method(settings, "SetPhaseFilter", [(phase_filter_id, phase_filter_type), (phase_filter_type, phase_filter_id)])


def _get_grid_category():
    try:
        return doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Grids)
    except Exception:
        return None


def _get_category_from_builtin(source_doc, built_in_category):
    try:
        return source_doc.Settings.Categories.get_Item(built_in_category)
    except Exception:
        return None


def _collect_named_subcategories(parent_category, wanted_names):
    found = {}
    if parent_category is None:
        return found
    wanted_lookup = dict((_normalize_label(name), name) for name in wanted_names)
    try:
        subcategories = parent_category.SubCategories
    except Exception:
        subcategories = None
    if subcategories is None:
        return found
    try:
        for subcategory in subcategories:
            try:
                actual_name = subcategory.Name
            except Exception:
                continue
            wanted_name = wanted_lookup.get(_normalize_label(actual_name))
            if wanted_name:
                found[wanted_name] = subcategory
    except Exception:
        pass
    return found


def _iter_annotation_categories(source_doc):
    categories = []
    try:
        for category in source_doc.Settings.Categories:
            try:
                if category.CategoryType == DB.CategoryType.Annotation:
                    categories.append(category)
            except Exception:
                continue
    except Exception:
        return []
    return categories


def _try_invoke_link_category_method(target, methods, link_type, category, hide):
    if target is None or category is None:
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
                pass
            try:
                param_name = parameter.Name.lower()
            except Exception:
                pass

            if "Boolean" in type_name:
                candidates = [bool_primary, not bool_primary]
            elif "ElementId" in type_name:
                if "link" in param_name:
                    candidates = [link_type.Id]
                elif "cat" in param_name or "category" in param_name or "sub" in param_name:
                    candidates = [category.Id]
                else:
                    candidates = [link_type.Id, category.Id]
            elif "Category" in type_name:
                candidates = [category]
            elif "RevitLinkType" in type_name:
                candidates = [link_type]
            elif "Int32" in type_name:
                if "link" in param_name:
                    candidates = [link_type.Id.IntegerValue]
                elif "cat" in param_name or "category" in param_name or "sub" in param_name:
                    candidates = [category.Id.IntegerValue]
                else:
                    candidates = [link_type.Id.IntegerValue, category.Id.IntegerValue]
            else:
                if "link" in param_name:
                    candidates = [link_type.Id, link_type]
                elif "cat" in param_name or "category" in param_name or "sub" in param_name:
                    candidates = [category.Id, category]
                else:
                    candidates = [category.Id, link_type.Id, bool_primary]
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


def _apply_link_category_visibility(view, link_type, category, hide):
    category_name = getattr(category, "Name", "<Unnamed Category>")
    settings = None
    try:
        settings = view.GetLinkOverrides(link_type.Id)
    except Exception:
        settings = None

    if settings is not None:
        try:
            methods_to_try = [
                method for method in settings.GetType().GetMethods()
                if any(token in method.Name.lower() for token in ("category", "annotation", "model", "sub"))
                and any(token in method.Name.lower() for token in ("hide", "hidden", "visible", "visibility"))
            ]
        except Exception:
            methods_to_try = []
        success, method_name = _try_invoke_link_category_method(settings, methods_to_try, link_type, category, hide)
        if success:
            try:
                view.SetLinkOverrides(link_type.Id, settings)
                return True, "settings.{} | {}".format(method_name, category_name)
            except Exception as exc:
                return False, "{} failed: {}".format(category_name, exc)

    try:
        view_methods = [
            method for method in view.GetType().GetMethods()
            if "link" in method.Name.lower()
            and any(token in method.Name.lower() for token in ("category", "annotation", "model", "sub"))
            and any(token in method.Name.lower() for token in ("hide", "hidden", "visible", "visibility"))
        ]
    except Exception:
        view_methods = []

    success, method_name = _try_invoke_link_category_method(view, view_methods, link_type, category, hide)
    if success:
        return True, "view.{} | {}".format(method_name, category_name)
    return False, "Linked category API not exposed for {}".format(category_name)


def _apply_link_grid_visibility(view, link_type, hide):
    grid_category = _get_grid_category()
    if grid_category is None:
        return False, "Grid category not available"
    return _apply_link_category_visibility(view, link_type, grid_category, hide)


def _copy_text_to_clipboard(text):
    if not text:
        return False, "Nothing to copy"
    try:
        Clipboard.SetText(text)
        return True, None
    except Exception as exc:
        return False, str(exc)


class PlanLinkSetupWindow(forms.WPFWindow):
    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)
        self._full_tree_roots = []
        self._tree_roots = []
        self._all_sheets = []
        self._all_view_rows = []
        self._view_rows = []
        self._all_link_rows = []
        self._link_rows = []
        self._linked_view_choices = []
        self._last_clicked_sheet_node = None

        self._init_static_controls()
        self._load_context(initial_sheet_ids=_get_default_sheet_ids())

    def _init_static_controls(self):
        self.UI_template_policy.ItemsSource = [
            "Edit selected views",
            "Edit views unless template controls RVT Links",
            "Edit view templates",
            "Skip templated views",
        ]
        self.UI_template_policy.SelectedItem = "Edit views unless template controls RVT Links"

        self.UI_release_rvt_links_from_templates.IsChecked = True
        self.UI_hide_grids_all_links.IsChecked = True
        self.UI_adapt_level_views.IsChecked = True
        self.UI_finish_custom.IsChecked = True
        self.UI_view_filters_by_host.IsChecked = True

        self.UI_linked_view.ItemsSource = ["<Select a source link first>"]
        self.UI_linked_view.SelectedIndex = 0

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

    def _load_sheets(self, initial_sheet_ids=None):
        previous_selection = set(initial_sheet_ids or [])
        if self._full_tree_roots:
            for sheet in self._get_selected_sheets():
                try:
                    previous_selection.add(sheet.Id.IntegerValue)
                except Exception:
                    continue
        try:
            self._all_sheets = list(
                DB.FilteredElementCollector(doc)
                .OfCategory(DB.BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception:
            self._all_sheets = []
        self._full_tree_roots = _build_sheet_tree(self._all_sheets, previous_selection)
        self._tree_roots = self._full_tree_roots
        self._sync_group_checkstates(self._full_tree_roots)
        self.UI_tree_sheets.ItemsSource = self._tree_roots

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
        self._refresh_links()
        self._update_counts()

    def _get_selected_view_rows(self):
        return [row for row in self._all_view_rows if row.IsChecked]

    def _apply_link_filter(self, raw_text):
        search_text = (raw_text or "").strip().lower()
        if not search_text:
            self._link_rows = self._all_link_rows
        else:
            self._link_rows = [row for row in self._all_link_rows if search_text in row.SearchText]
        self.UI_link_list.ItemsSource = self._link_rows
        self.UI_link_list.Items.Refresh()
        self._update_counts()

    def _refresh_links(self):
        previous_id = None
        try:
            selected = self.UI_link_list.SelectedItem
            if selected is not None:
                previous_id = selected.LinkType.Id.IntegerValue
        except Exception:
            previous_id = None

        self._all_link_rows = _collect_link_rows_for_views(self._get_selected_view_rows())
        self._apply_link_filter(self.UI_link_filter.Text if self.UI_link_filter else "")

        restored = False
        if previous_id is not None:
            for row in self._link_rows:
                try:
                    if row.LinkType.Id.IntegerValue == previous_id:
                        self.UI_link_list.SelectedItem = row
                        restored = True
                        break
                except Exception:
                    continue
        if not restored and self._link_rows:
            self.UI_link_list.SelectedIndex = 0
        self._refresh_linked_view_choices()

    def _refresh_linked_view_choices(self):
        previous_display = None
        try:
            selected = self.UI_linked_view.SelectedItem
            if selected is not None:
                previous_display = str(selected)
        except Exception:
            previous_display = None

        source_link = self.UI_link_list.SelectedItem
        if source_link is None:
            self._linked_view_choices = ["<Select a source link first>"]
            self.UI_linked_view.ItemsSource = self._linked_view_choices
            self.UI_linked_view.SelectedIndex = 0
            self._update_counts()
            return

        link_doc = _find_link_doc_for_type(doc, source_link.LinkType.Id)
        if not link_doc:
            self._linked_view_choices = ["<Selected link is not loaded>"]
            self.UI_linked_view.ItemsSource = self._linked_view_choices
            self.UI_linked_view.SelectedIndex = 0
            self._update_counts()
            return

        choices = [LinkedViewOption(view) for view in _collect_link_views(link_doc)]
        choices.sort(key=lambda item: item.Display.lower())
        if not choices:
            self._linked_view_choices = ["<No printable linked views found>"]
            self.UI_linked_view.ItemsSource = self._linked_view_choices
            self.UI_linked_view.SelectedIndex = 0
            self._update_counts()
            return

        self._linked_view_choices = choices
        self.UI_linked_view.ItemsSource = choices
        restored = False
        if previous_display:
            for choice in choices:
                if choice.Display == previous_display:
                    self.UI_linked_view.SelectedItem = choice
                    restored = True
                    break
        if not restored:
            self.UI_linked_view.SelectedIndex = 0
        self._update_counts()

    def _update_counts(self):
        sheet_total = len(self._get_selected_sheets())
        view_total = len(self._all_view_rows)
        view_selected = len(self._get_selected_view_rows())
        loaded_links = sum(1 for row in self._all_link_rows if row.IsLoaded)
        selected_link = self.UI_link_list.SelectedItem
        selected_link_name = selected_link.Name if selected_link else "None"

        self.UI_sheet_count.Text = "({} selected)".format(sheet_total)
        self.UI_view_count.Text = "({} of {})".format(view_selected, view_total)
        self.UI_link_count.Text = "({} loaded | selected: {})".format(loaded_links, selected_link_name)
        self.UI_info_summary.Text = (
            "Sheets: {} selected\n"
            "Views: {} checked\n"
            "Loaded links in checked views: {}"
        ).format(sheet_total, view_selected, loaded_links)

    def _append_log(self, text):
        try:
            current = self.UI_log.Text or ""
            self.UI_log.Text = "{}\n{}".format(current, text).strip()
            self.UI_log.ScrollToEnd()
        except Exception:
            pass

    def _set_status(self, text):
        try:
            self.UI_status_bar.Text = text
        except Exception:
            pass

    def _resolve_linked_view_for_target(self, source_link, selected_linked_view, target_view, adapt_level_views):
        link_doc = _find_link_doc_for_type(doc, source_link.LinkType.Id)
        if not link_doc:
            return None, "Selected source link is not loaded"
        linked_view = getattr(selected_linked_view, "View", None)
        if linked_view is None:
            return None, "No linked view selected"
        if not adapt_level_views:
            return linked_view, None
        target_level = _safe_get_view_level(target_view)
        fallback_viewtype = None
        try:
            fallback_viewtype = target_view.ViewType
        except Exception:
            pass
        matched = _choose_linked_view_for_level(link_doc, linked_view, target_level, fallback_viewtype)
        if matched is not None:
            return matched, None
        return linked_view, "No level-matched linked view found; used selected linked view directly"

    def _apply_fixed_custom_settings(self, settings, use_view_filters_by_host):
        applied = []
        skipped = []

        success, message = _apply_link_visibility_property(settings, "ObjectStyles", "By linked view")
        if success:
            applied.append("Object styles set to By Linked Model")
        elif message:
            skipped.append(message)

        if use_view_filters_by_host:
            success, message = _apply_link_visibility_property(settings, "ViewFilterType", "By host view")
            if success:
                applied.append("View filters set to By Host View")
            elif message:
                skipped.append(message)

        for property_name, label in [
            ("ViewRange", "View range set to By Linked View"),
            ("NestedLinks", "Nested links set to By Linked View"),
            ("ColorFill", "Color fill set to By Linked View"),
        ]:
            success, message = _apply_link_visibility_property(settings, property_name, "By linked view")
            if success:
                applied.append(label)
            elif message:
                skipped.append(message)

        success, message = _apply_discipline_setting(settings, "By linked view")
        if success:
            applied.append("Discipline set to By Linked View")
        elif message:
            skipped.append(message)

        success, message = _apply_detail_level_setting(settings, "By linked view")
        if success:
            applied.append("Detail level set to By Linked View")
        elif message:
            skipped.append(message)

        success, message = _apply_phase_setting(settings, "By linked view")
        if success:
            applied.append("Phase set to By Linked View")
        elif message:
            skipped.append(message)

        success, message = _apply_phase_filter_setting(settings, "By linked view")
        if success:
            applied.append("Phase filter set to By Linked View")
        elif message:
            skipped.append(message)

        return applied, skipped

    def _apply_room_model_cleanup(self, target_view, link_type):
        applied = []
        skipped = []
        room_category = _get_category_from_builtin(doc, DB.BuiltInCategory.OST_Rooms)
        subcategories = _collect_named_subcategories(room_category, ["Color Fill", "Interior Fill", "Reference"])
        for subcategory_name in ["Color Fill", "Interior Fill", "Reference"]:
            category = subcategories.get(subcategory_name)
            if category is None:
                skipped.append("Rooms > {} category not found".format(subcategory_name))
                continue
            success, message = _apply_link_category_visibility(target_view, link_type, category, True)
            if success:
                applied.append("Rooms > {} hidden".format(subcategory_name))
            elif message:
                skipped.append(message)
        return applied, skipped

    def _apply_annotation_cleanup(self, target_view, link_type):
        applied = []
        skipped = []
        hidden_count = 0
        room_tags_category = _get_category_from_builtin(doc, DB.BuiltInCategory.OST_RoomTags)
        text_notes_category = _get_category_from_builtin(doc, DB.BuiltInCategory.OST_TextNotes)

        seen = set()
        for category in _iter_annotation_categories(doc):
            try:
                category_id = category.Id.IntegerValue
            except Exception:
                continue
            if category_id in seen:
                continue
            seen.add(category_id)
            if room_tags_category is not None and category_id == room_tags_category.Id.IntegerValue:
                continue
            success, message = _apply_link_category_visibility(target_view, link_type, category, True)
            if success:
                hidden_count += 1
            elif message:
                skipped.append(message)

        if hidden_count:
            applied.append("Annotation categories set to Custom; hid {} non-room-tag categories".format(hidden_count))

        if room_tags_category is not None:
            success, message = _apply_link_category_visibility(target_view, link_type, room_tags_category, False)
            if success:
                applied.append("Room Tags kept visible")
            elif message:
                skipped.append(message)

        if text_notes_category is not None:
            success, message = _apply_link_category_visibility(target_view, link_type, text_notes_category, True)
            if success:
                applied.append("Text Notes hidden")
            elif message:
                skipped.append(message)

        return applied, skipped

    def _apply_to_targets(self):
        selected_view_rows = self._get_selected_view_rows()
        if not selected_view_rows:
            forms.alert("Check at least one placed view first.", title=__title__)
            return

        source_link = self.UI_link_list.SelectedItem
        if source_link is None:
            forms.alert("Select the source link first.", title=__title__)
            return

        selected_linked_view = self.UI_linked_view.SelectedItem
        if not isinstance(selected_linked_view, LinkedViewOption):
            forms.alert("Choose a linked view from the selected link first.", title=__title__)
            return

        template_policy = str(self.UI_template_policy.SelectedItem or "Edit views unless template controls RVT Links")
        release_template_link_control = bool(getattr(self.UI_release_rvt_links_from_templates, "IsChecked", False))
        hide_grids_all_links = bool(getattr(self.UI_hide_grids_all_links, "IsChecked", False))
        adapt_level_views = bool(getattr(self.UI_adapt_level_views, "IsChecked", False))
        finish_custom = bool(getattr(self.UI_finish_custom, "IsChecked", False))
        view_filters_by_host = bool(getattr(self.UI_view_filters_by_host, "IsChecked", False))

        if not source_link.IsLoaded:
            forms.alert("The selected source link is not loaded.", title=__title__)
            return

        template_release_count = 0
        template_release_skips = []
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

        confirmation = (
            "Target views/templates: {}\n"
            "Source link: {}\n"
            "Linked view: {}\n"
            "Hide grids on visible links: {}\n"
            "Finish source link as Custom: {}\n\n"
            "Proceed with the batch setup?"
        ).format(
            len(targets),
            source_link.Name,
            selected_linked_view.Display,
            "Yes" if hide_grids_all_links else "No",
            "Yes" if finish_custom else "No",
        )
        if not forms.alert(confirmation, title=__title__, yes=True, no=True):
            return

        applied_counts = {}
        skipped_notes = []
        grid_support_details = set()

        def _count(label):
            applied_counts[label] = applied_counts.get(label, 0) + 1

        with revit.Transaction("Plan Link Setup"):
            for target_view in targets:
                linked_view, linked_view_note = self._resolve_linked_view_for_target(
                    source_link,
                    selected_linked_view,
                    target_view,
                    adapt_level_views,
                )
                if linked_view is None:
                    skipped_notes.append("{} | {}".format(target_view.Name, linked_view_note or "Could not resolve linked view"))
                    continue
                if linked_view_note:
                    skipped_notes.append("{} | {}".format(target_view.Name, linked_view_note))

                if hide_grids_all_links:
                    for link_row in _collect_visible_link_options(target_view):
                        success, detail = _apply_link_grid_visibility(target_view, link_row.LinkType, True)
                        if success:
                            _count("Hidden linked grids")
                            if detail:
                                grid_support_details.add(detail)
                        elif detail:
                            skipped_notes.append("{} | {} | {}".format(target_view.Name, link_row.Name, detail))

                try:
                    settings = target_view.GetLinkOverrides(source_link.LinkType.Id)
                except Exception:
                    settings = None
                if settings is None:
                    skipped_notes.append("{} | Could not read link overrides for '{}'".format(target_view.Name, source_link.Name))
                    continue

                by_linked_view_member = _get_link_visibility_member("By linked view")
                if by_linked_view_member is None:
                    skipped_notes.append("{} | By Linked View mode is not available in this Revit build".format(target_view.Name))
                    continue

                try:
                    settings.LinkVisibilityType = by_linked_view_member
                    settings.LinkedViewId = linked_view.Id
                    target_view.SetLinkOverrides(source_link.LinkType.Id, settings)
                    _count("Set source link to By Linked View")
                except Exception as exc:
                    skipped_notes.append("{} | By Linked View failed: {}".format(target_view.Name, exc))
                    continue

                if not finish_custom:
                    continue

                try:
                    frozen_settings = target_view.GetLinkOverrides(source_link.LinkType.Id)
                except Exception:
                    frozen_settings = None
                if frozen_settings is None:
                    skipped_notes.append("{} | Could not reload link overrides for Custom finishing".format(target_view.Name))
                    continue

                custom_member = _get_link_visibility_member("Custom / None")
                if custom_member is None:
                    skipped_notes.append("{} | Custom link display mode is not available".format(target_view.Name))
                    continue

                try:
                    frozen_settings.LinkVisibilityType = custom_member
                except Exception as exc:
                    skipped_notes.append("{} | Could not switch link to Custom: {}".format(target_view.Name, exc))
                    continue

                custom_applied, custom_skipped = self._apply_fixed_custom_settings(frozen_settings, view_filters_by_host)
                try:
                    target_view.SetLinkOverrides(source_link.LinkType.Id, frozen_settings)
                    _count("Switched source link to Custom")
                except Exception as exc:
                    skipped_notes.append("{} | Applying Custom settings failed: {}".format(target_view.Name, exc))
                    continue

                for label in custom_applied:
                    _count(label)
                for note in custom_skipped:
                    skipped_notes.append("{} | {}".format(target_view.Name, note))

                room_applied, room_skipped = self._apply_room_model_cleanup(target_view, source_link.LinkType)
                for label in room_applied:
                    _count(label)
                for note in room_skipped:
                    skipped_notes.append("{} | {}".format(target_view.Name, note))

                annotation_applied, annotation_skipped = self._apply_annotation_cleanup(target_view, source_link.LinkType)
                for label in annotation_applied:
                    _count(label)
                for note in annotation_skipped:
                    skipped_notes.append("{} | {}".format(target_view.Name, note))

        summary_lines = []
        summary_lines.append("Plan Link Setup")
        summary_lines.append("Targets: {}".format(len(targets)))
        summary_lines.append("Source link: {}".format(source_link.Name))
        summary_lines.append("Linked view: {}".format(selected_linked_view.Display))
        summary_lines.append("")
        summary_lines.append("Applied actions:")

        if template_release_count:
            summary_lines.append("- Released RVT Links control on {} template(s)".format(template_release_count))

        ordered = [
            "Hidden linked grids",
            "Set source link to By Linked View",
            "Switched source link to Custom",
            "View filters set to By Host View",
            "Object styles set to By Linked Model",
            "View range set to By Linked View",
            "Nested links set to By Linked View",
            "Color fill set to By Linked View",
            "Discipline set to By Linked View",
            "Detail level set to By Linked View",
            "Phase set to By Linked View",
            "Phase filter set to By Linked View",
            "Rooms > Color Fill hidden",
            "Rooms > Interior Fill hidden",
            "Rooms > Reference hidden",
            "Room Tags kept visible",
            "Text Notes hidden",
        ]
        for label in ordered:
            count = applied_counts.get(label, 0)
            if count:
                summary_lines.append("- {} on {} target(s)".format(label, count))

        if len(summary_lines) == 6:
            summary_lines.append("- No settings were changed")

        if grid_support_details:
            summary_lines.append("")
            summary_lines.append("Execution details:")
            summary_lines.append("- Grid API route: {}".format(", ".join(sorted(grid_support_details))))

        unique_skips = sorted(set(skipped_notes + template_release_skips))
        if unique_skips:
            summary_lines.append("")
            summary_lines.append("Not applied:")
            summary_lines.extend("- {}".format(line) for line in unique_skips)

        summary_text = "\n".join(summary_lines)
        self.UI_feedback.Text = "Applied to {} target(s).".format(len(targets))
        self._set_status("Applied to {} target(s).".format(len(targets)))
        self.UI_log.Text = summary_text

        clipboard_ok, clipboard_error = _copy_text_to_clipboard(summary_text)
        if clipboard_ok:
            summary_text = "{}\n\nAction list copied to clipboard.".format(summary_text)
        elif clipboard_error:
            summary_text = "{}\n\nCould not copy to clipboard: {}".format(summary_text, clipboard_error)
        forms.alert(summary_text, title=__title__)

    def window_loaded(self, sender, event_args):
        self._set_status("Context loaded.")

    def button_refresh_context(self, sender, event_args):
        self._load_context(initial_sheet_ids=None)

    def _load_context(self, initial_sheet_ids=None):
        self._load_sheets(initial_sheet_ids=initial_sheet_ids)
        self._apply_sheet_filter(self.UI_search.Text if self.UI_search else "")
        self._refresh_views()
        self._refresh_links()
        self._update_counts()
        self._set_status("Context loaded.")

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
        self._refresh_links()
        self._update_counts()

    def button_check_all_views(self, sender, event_args):
        for row in self._view_rows:
            row.IsChecked = True
        self.UI_view_list.Items.Refresh()
        self._refresh_links()
        self._update_counts()

    def button_uncheck_all_views(self, sender, event_args):
        for row in self._view_rows:
            row.IsChecked = False
        self.UI_view_list.Items.Refresh()
        self._refresh_links()
        self._update_counts()

    def button_invert_shown_views(self, sender, event_args):
        for row in self._view_rows:
            row.IsChecked = not bool(row.IsChecked)
        self.UI_view_list.Items.Refresh()
        self._refresh_links()
        self._update_counts()

    def link_filter_text_changed(self, sender, event_args):
        self._apply_link_filter(self.UI_link_filter.Text if self.UI_link_filter else "")

    def link_selection_changed(self, sender, event_args):
        self._refresh_linked_view_choices()

    def button_select_arch_link(self, sender, event_args):
        for row in self._link_rows:
            if row.IsArchitectural and row.IsLoaded:
                self.UI_link_list.SelectedItem = row
                break
        self._refresh_linked_view_choices()

    def button_select_first_link(self, sender, event_args):
        if self._link_rows:
            self.UI_link_list.SelectedIndex = 0
        self._refresh_linked_view_choices()

    def button_clear_link_selection(self, sender, event_args):
        self.UI_link_list.SelectedIndex = -1
        self._refresh_linked_view_choices()

    def button_apply(self, sender, event_args):
        self._apply_to_targets()


if __name__ == "__main__":
    if not doc:
        forms.alert("Open a Revit model first.", title=__title__)
    else:
        PlanLinkSetupWindow("PlanLinkSetup.xaml").ShowDialog()
