# -*- coding: utf-8 -*-
__title__ = "Auto Room Names + Hide Grids"
__doc__ = """Version = 1.0
Date: 2026-03-27
Author: GM
Description:
Combines Auto Room Names and Hide Grids functionality.
Automatically finds the best architectural linked view for room names,
applies room-specific visibility settings, and hides grids on all visible links.
How-to:
1. Open the target model view.
2. Click this tool.
3. The tool scans visible architectural links and scores linked views.
4. Applies the best match for room names.
5. Hides grids and applies cleanup to all visible links.
"""

import clr
clr.AddReference("System")

import System
from System.Collections.Generic import List

from pyrevit import DB, forms, revit


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None


def _safe_view_name(view):
    try:
        return view.Name or "<Unnamed View>"
    except Exception:
        return "<Unnamed View>"


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


def _safe_get_view_level(view):
    try:
        if hasattr(view, "GenLevel"):
            return view.GenLevel
    except Exception:
        return None
    return None


def _normalize_label(text):
    return "".join(ch for ch in (text or "").lower() if ch.isalnum())


def _looks_architectural(text):
    normalized = (text or "").lower()
    return any(token in normalized for token in ("arch", "architect", "arc", "a-model", "a_"))


def _find_link_doc_for_type(host_doc, link_type_id):
    try:
        instances = DB.FilteredElementCollector(host_doc).OfClass(DB.RevitLinkInstance).ToElements()
    except Exception:
        return None
    for instance in instances:
        try:
            if instance.GetTypeId() == link_type_id:
                link_doc = instance.GetLinkDocument()
                if link_doc:
                    return link_doc
        except Exception:
            continue
    return None


def _collect_visible_link_types(view):
    visible = {}
    try:
        instances = DB.FilteredElementCollector(doc, view.Id).OfClass(DB.RevitLinkInstance).ToElements()
    except Exception:
        instances = []
    for instance in instances:
        try:
            # Check if the link instance is actually visible in this view
            if not instance.IsHidden(view):
                link_type_id = instance.GetTypeId()
                if link_type_id == DB.ElementId.InvalidElementId:
                    continue
                link_type = doc.GetElement(link_type_id)
                if link_type is not None and hasattr(link_type, 'IsLoaded') and link_type.IsLoaded:
                    visible[link_type_id.IntegerValue] = link_type
        except Exception:
            continue
    rows = list(visible.values())
    rows.sort(key=lambda item: (0 if _looks_architectural(_safe_link_type_name(item)) else 1, _safe_link_type_name(item).lower()))
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


def _release_revit_links_template_control(view):
    template = _get_view_template(view)
    if template is None:
        return False, None
    links_param_id = DB.ElementId(int(DB.BuiltInParameter.VIS_GRAPHICS_RVT_LINKS))
    try:
        non_controlled = list(template.GetNonControlledTemplateParameterIds())
        if any(existing == links_param_id for existing in non_controlled):
            return False, None
        updated = List[DB.ElementId]()
        for existing in non_controlled:
            updated.Add(existing)
        updated.Add(links_param_id)
        template.SetNonControlledTemplateParameterIds(updated)
        return True, template.Name
    except Exception as exc:
        return False, str(exc)


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


def _get_view_discipline_member():
    enum_type = getattr(DB, "ViewDiscipline", None)
    if enum_type is None:
        return None
    return _get_enum_member(enum_type, ("Coordination",))


def _get_view_detail_level_member():
    enum_type = getattr(DB, "ViewDetailLevel", None)
    if enum_type is None:
        return None
    return _get_enum_member(enum_type, ("Medium",))


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


def _choose_linked_view_for_level(link_doc, host_view):
    if not link_doc:
        return None

    host_level = _safe_get_view_level(host_view)
    link_level = _best_match_link_level(link_doc, host_level)
    if link_level is None:
        return None

    fallback_viewtype = None
    try:
        fallback_viewtype = host_view.ViewType
    except Exception:
        fallback_viewtype = None

    candidates = []
    try:
        views = DB.FilteredElementCollector(link_doc).OfClass(DB.View).ToElements()
    except Exception:
        views = []

    for view in views:
        try:
            if view.IsTemplate:
                continue
            if fallback_viewtype is not None and view.ViewType != fallback_viewtype:
                continue
            if not hasattr(view, "GenLevel") or not view.GenLevel:
                continue
            if view.GenLevel.Id != link_level.Id:
                continue
            if hasattr(view, "CanBePrinted") and not view.CanBePrinted:
                continue
            candidates.append(view)
        except Exception:
            continue

    if not candidates:
        return None
    return candidates[0]


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


def _name_match_score(host_view_name, linked_view_name, host_level_name, linked_level_name):
    score = 0
    host_view_lower = (host_view_name or "").lower()
    linked_view_lower = (linked_view_name or "").lower()
    host_level_lower = (host_level_name or "").lower()
    linked_level_lower = (linked_level_name or "").lower()

    if host_view_lower == linked_view_lower:
        score += 90
    elif host_view_lower in linked_view_lower or linked_view_lower in host_view_lower:
        score += 45

    if host_level_lower and linked_level_lower:
        if host_level_lower == linked_level_lower:
            score += 40
        elif host_level_lower in linked_level_lower or linked_level_lower in host_level_lower:
            score += 20

    return score


def _score_linked_view(host_view, linked_view, link_type):
    try:
        host_view_name = _safe_view_name(host_view)
        linked_view_name = _safe_view_name(linked_view)
        host_level = _safe_get_view_level(host_view)
        linked_level = _safe_get_view_level(linked_view)
        host_level_name = _safe_element_name(host_level) if host_level else ""
        linked_level_name = _safe_element_name(linked_level) if linked_level else ""

        score = 0
        score += _name_match_score(host_view_name, linked_view_name, host_level_name, linked_level_name)

        room_count = 0
        room_tag_count = 0
        text_note_count = 0

        try:
            collector = DB.FilteredElementCollector(linked_view.Document, linked_view.Id)
            elements = collector.WhereElementIsNotElementType().ToElements()
        except Exception:
            elements = []

        for element in elements:
            try:
                category = element.Category
                if category:
                    if category.Id == DB.ElementId(DB.BuiltInCategory.OST_Rooms):
                        room_count += 1
                    elif category.Id == DB.ElementId(DB.BuiltInCategory.OST_RoomTags):
                        room_tag_count += 1
                    elif category.Id == DB.ElementId(DB.BuiltInCategory.OST_TextNotes):
                        text_note_count += 1
            except Exception:
                continue

        score += min(room_tag_count, 12) * 12
        score += min(room_count, 20) * 2
        score += _name_match_score(host_view_name, linked_view_name, host_level_name, linked_level_name)
        score += min(text_note_count, 10)

        return {
            "score": score,
            "room_count": room_count,
            "room_tag_count": room_tag_count,
            "text_note_count": text_note_count,
            "linked_view": linked_view,
            "link_type": link_type,
        }
    except Exception:
        return None


def _find_best_arch_room_view(host_view):
    visible_links = _collect_visible_link_types(host_view)
    if not visible_links:
        return None, []

    candidates = []
    for link_type in visible_links:
        if not _looks_architectural(_safe_link_type_name(link_type)):
            continue

        link_doc = _find_link_doc_for_type(doc, link_type.Id)
        if not link_doc:
            continue

        try:
            views = DB.FilteredElementCollector(link_doc).OfClass(DB.View).ToElements()
        except Exception:
            continue

        for linked_view in views:
            try:
                if linked_view.IsTemplate or not hasattr(linked_view, "CanBePrinted") or not linked_view.CanBePrinted:
                    continue
                if linked_view.ViewType != DB.ViewType.FloorPlan and linked_view.ViewType != DB.ViewType.CeilingPlan:
                    continue

                scored = _score_linked_view(host_view, linked_view, link_type)
                if scored:
                    candidates.append(scored)
            except Exception:
                continue

    if not candidates:
        return None, []

    candidates.sort(key=lambda item: -item["score"])
    return candidates[0], candidates


def _apply_fixed_custom_settings(settings):
    applied = []
    skipped = []

    # Apply discipline
    success, message = _apply_discipline_setting(settings, "By linked view")
    if success:
        applied.append("Discipline set to Coordination")
    elif message:
        skipped.append(message)

    # Apply detail level
    success, message = _apply_detail_level_setting(settings, "By linked view")
    if success:
        applied.append("Detail Level set to Medium")
    elif message:
        skipped.append(message)

    # Apply phase
    success, message = _apply_phase_setting(settings, "By linked view")
    if success:
        applied.append("Phase matched to linked view")
    elif message:
        skipped.append(message)

    # Apply phase filter
    success, message = _apply_phase_filter_setting(settings, "By linked view")
    if success:
        applied.append("Phase Filter matched to linked view")
    elif message:
        skipped.append(message)

    return applied, skipped


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
        discipline = _get_view_discipline_member()
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
        detail_level = _get_view_detail_level_member()
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


def _get_category_from_builtin(source_doc, built_in_category):
    try:
        return source_doc.Settings.Categories.get_Item(built_in_category)
    except Exception:
        return None


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
                return True, "settings.{}".format(method_name)
            except Exception:
                pass

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
        return True, "view.{}".format(method_name)

    return False, "Category visibility API not exposed for {} in this Revit build".format(category_name)


def _apply_room_model_cleanup(target_view, link_type):
    applied = []
    skipped = []

    # Hide model categories except rooms
    model_categories = []
    try:
        for category in doc.Settings.Categories:
            try:
                if category.CategoryType == DB.CategoryType.Model and category.Id != DB.ElementId(DB.BuiltInCategory.OST_Rooms):
                    model_categories.append(category)
            except Exception:
                continue
    except Exception:
        model_categories = []

    for category in model_categories:
        success, message = _apply_link_category_visibility(target_view, link_type, category, True)
        if success:
            applied.append("Model category '{}' hidden".format(_safe_element_name(category)))
        elif message:
            skipped.append("Model category '{}': {}".format(_safe_element_name(category), message))

    return applied, skipped


def _apply_annotation_cleanup(target_view, link_type):
    applied = []
    skipped = []

    # Get annotation categories
    annotation_categories = _iter_annotation_categories(doc)
    room_tags_category = None
    text_notes_category = None

    for category in annotation_categories:
        try:
            if category.Id == DB.ElementId(DB.BuiltInCategory.OST_RoomTags):
                room_tags_category = category
            elif category.Id == DB.ElementId(DB.BuiltInCategory.OST_TextNotes):
                text_notes_category = category
        except Exception:
            continue

    # Hide non-room-tag annotations
    hidden_count = 0
    for category in annotation_categories:
        if category == room_tags_category:
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


def _apply_hide_grids_to_all_links(target_view, visible_links):
    """Apply Hide Grids logic to all visible links."""
    grid_count = 0
    grid_notes = []
    
    for link_type in visible_links:
        link_name = _safe_link_type_name(link_type)
        
        # Get or create settings
        try:
            settings = target_view.GetLinkOverrides(link_type.Id)
        except Exception:
            settings = None
        
        if settings is not None:
            # Set to Custom
            custom_member = _get_link_visibility_member("Custom / None")
            if custom_member is not None:
                try:
                    settings.LinkVisibilityType = custom_member
                except Exception as exc:
                    grid_notes.append("{} | Could not set Custom visibility: {}".format(link_name, exc))
            
            # Enable halftone
            if hasattr(settings, "Halftone"):
                try:
                    settings.Halftone = True
                except Exception as exc:
                    grid_notes.append("{} | Could not set Halftone: {}".format(link_name, exc))
            
            # Set underlay
            link_doc = _find_link_doc_for_type(doc, link_type.Id)
            matched_linked_view = _choose_linked_view_for_level(link_doc, target_view)
            if matched_linked_view is not None:
                try:
                    settings.LinkedViewId = matched_linked_view.Id
                    success, message = _apply_link_visibility_property(settings, "Underlay", "By linked view")
                    if not success and message:
                        grid_notes.append("{} | {}".format(link_name, message))
                except Exception as exc:
                    grid_notes.append("{} | Could not assign linked view: {}".format(link_name, exc))
            else:
                success, message = _apply_link_visibility_property(settings, "Underlay", "By host view")
                if not success and message:
                    grid_notes.append("{} | {}".format(link_name, message))
            
            # Apply settings
            try:
                target_view.SetLinkOverrides(link_type.Id, settings)
            except Exception as exc:
                grid_notes.append("{} | Could not apply link overrides: {}".format(link_name, exc))
        
        # Hide grids
        success, detail = _apply_link_grid_visibility(target_view, link_type, True)
        if success:
            grid_count += 1
        else:
            grid_notes.append("{} | {}".format(link_name, detail or "Grid hide failed"))
    
    return grid_count, grid_notes


def main():
    if uidoc is None or doc is None:
        forms.alert("Open a Revit model before running this tool.", exitscript=True)

    target_view = revit.active_view
    if target_view is None:
        forms.alert("No active view is available.", exitscript=True)

    try:
        if target_view.ViewType == DB.ViewType.DrawingSheet:
            forms.alert("This tool works on model views, not sheets.", exitscript=True)
    except Exception:
        pass

    visible_links = _collect_visible_link_types(target_view)
    if not visible_links:
        forms.alert("No visible Revit links were found in the current view.", exitscript=True)

    # Find best architectural view for room names
    best_match, ranked_matches = _find_best_arch_room_view(target_view)
    if best_match is None:
        forms.alert(
            "No visible architectural link with a matching linked view could be resolved for this view.",
            title=__title__,
            exitscript=True,
        )

    selected_link = best_match["link_type"]
    linked_view = best_match["linked_view"]

    applied_counts = {}
    skipped_notes = []
    template_notes = []

    def _count(label):
        applied_counts[label] = applied_counts.get(label, 0) + 1

    with revit.Transaction(__title__):
        # Release template control
        if _template_controls_revit_links(target_view):
            released, detail = _release_revit_links_template_control(target_view)
            if released:
                template_notes.append("Released RVT Links from template: {}".format(detail or "<Unnamed Template>"))
                _count("Released RVT Links template control")
            elif detail:
                skipped_notes.append("Template release failed: {}".format(detail))

        if _template_controls_revit_links(target_view):
            forms.alert(
                "RVT Links are still controlled by the active view template. Release the template control first and run the tool again.",
                title=__title__,
                exitscript=True,
            )

        # Apply room names setup to selected link
        try:
            settings = target_view.GetLinkOverrides(selected_link.Id)
        except Exception:
            settings = None
        if settings is None:
            forms.alert("Could not read link overrides for the selected source link.", title=__title__, exitscript=True)

        # Set to By Linked View
        by_linked_view_member = _get_link_visibility_member("By linked view")
        if by_linked_view_member is None:
            forms.alert("By Linked View mode is not available in this Revit build.", title=__title__, exitscript=True)

        try:
            settings.LinkVisibilityType = by_linked_view_member
            settings.LinkedViewId = linked_view.Id
            target_view.SetLinkOverrides(selected_link.Id, settings)
            _count("Set source link to By Linked View")
        except Exception as exc:
            forms.alert("Applying By Linked View failed: {}".format(exc), title=__title__, exitscript=True)

        # Switch to Custom and apply room settings
        try:
            frozen_settings = target_view.GetLinkOverrides(selected_link.Id)
        except Exception:
            frozen_settings = None
        if frozen_settings is None:
            forms.alert("Could not reload link overrides for Custom finishing.", title=__title__, exitscript=True)

        custom_member = _get_link_visibility_member("Custom / None")
        if custom_member is None:
            forms.alert("Custom link display mode is not available.", title=__title__, exitscript=True)

        try:
            frozen_settings.LinkVisibilityType = custom_member
        except Exception as exc:
            forms.alert("Could not switch link to Custom: {}".format(exc), title=__title__, exitscript=True)

        # Apply room-specific settings
        custom_applied, custom_skipped = _apply_fixed_custom_settings(frozen_settings)
        room_applied, room_skipped = _apply_room_model_cleanup(target_view, selected_link)
        annotation_applied, annotation_skipped = _apply_annotation_cleanup(target_view, selected_link)
        
        try:
            target_view.SetLinkOverrides(selected_link.Id, frozen_settings)
            _count("Switched source link to Custom")
        except Exception as exc:
            forms.alert("Applying Custom settings failed: {}".format(exc), title=__title__, exitscript=True)

        # Count applied settings
        for label in custom_applied + room_applied + annotation_applied:
            _count(label)
        for note in custom_skipped + room_skipped + annotation_skipped:
            skipped_notes.append(note)

        # Apply Hide Grids to ALL visible links
        grid_count, grid_notes = _apply_hide_grids_to_all_links(target_view, visible_links)
        applied_counts["Hidden linked grids"] = grid_count
        skipped_notes.extend(grid_notes)

    # Summary
    summary_lines = [
        "View: {}".format(_safe_view_name(target_view)),
        "Source link: {}".format(_safe_link_type_name(selected_link)),
        "Linked view: {}".format(_safe_view_name(linked_view)),
        "Detected room tags: {}".format(best_match["room_tag_count"]),
        "Links processed: {}".format(len(visible_links)),
        "Hidden linked grids: {}".format(grid_count),
    ]
    
    note_count = len(template_notes) + len(skipped_notes)
    if note_count:
        summary_lines.append("Notes: {} item(s) were skipped or used fallback behavior.".format(note_count))
    
    forms.alert("\n".join(summary_lines), title=__title__, warn_icon=False)


main()