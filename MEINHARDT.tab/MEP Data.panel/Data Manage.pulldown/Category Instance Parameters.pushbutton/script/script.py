# coding: utf8
from __future__ import print_function

from collections import defaultdict
import ast
import csv
import datetime
import os
import re

from Autodesk.Revit.DB import (
    BuiltInParameter,
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    RevitLinkInstance,
    StorageType,
    Transaction,
    TransactionGroup,
)
from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow
from System.Windows.Controls import CheckBox
import System
from Microsoft.Win32 import SaveFileDialog


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
config = script.get_config()

__title__ = "Category\nInstance Params"
__doc__ = "Batch update instance parameter values for multiple categories at once."


MAX_COLLECT_LIMIT = 50000
MAX_APPLY_LIMIT = 50000

try:
    _TEXT_TYPE = unicode
except NameError:
    _TEXT_TYPE = str


def _sanitize_filename_part(value, max_len=60):
    text = "" if value is None else str(value)
    text = text.strip()
    if not text:
        return "NA"
    text = re.sub(r"[^A-Za-z0-9._-]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-._")
    if not text:
        text = "NA"
    if len(text) > max_len:
        text = text[:max_len].rstrip("-._")
    return text or "NA"


def _csv_cell(value):
    if value is None:
        return ""
    try:
        text = value if isinstance(value, _TEXT_TYPE) else _TEXT_TYPE(value)
    except Exception:
        text = str(value)
    try:
        return text.encode("utf-8")
    except Exception:
        return str(text)


def _normalize_match_key(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    return text


DEFAULT_TRANSLATION_ITEMS = [
    ("ALL", "Allied Health"),
    ("AMB", "Ambulatory Care Clinics"),
    ("BIR", "Labour & Birth Suite"),
    ("BOH", "Loading Dock, Logistics Services & Back of House"),
    ("BVM", "Breastfeeding Centre & Visiting Midwifery Service"),
    ("CES", "Corporate & Executive Services"),
    ("CSD", "Central Sterilising Services Department"),
    ("ENG", "Engineering"),
    ("EOT", "End of Trip Facilities"),
    ("ETC", "Education & Training Centre"),
    ("EXT", "External Areas"),
    ("FBC", "Family Birth Centre"),
    ("FOH", "Front of House"),
    ("HDU", "High Dependency Unit"),
    ("HIS", "Health Information & Administration Service"),
    ("IP1", "Inpatient Obstetric 1"),
    ("IP2", "Inpatient Obstetric 2"),
    ("IP3", "Inpatient Obstetric 3"),
    ("IP4", "Inpatient Obstetric 4"),
    ("IPG", "Inpatient Gynaecology"),
    ("KIT", "Kitchen & Catering Service"),
    ("LNK", "Link"),
    ("MFA", "Maternal & Fetal Assessment Unit"),
    ("MFM", "Maternal Fetal Medicine"),
    ("MHS", "Women & Newborn Mental Health Service"),
    ("MIM", "Medical Imaging"),
    ("MOR", "Mortuary"),
    ("NEO", "Neonatology Service"),
    ("NET", "Neonatal Emergency Transfer Service"),
    ("PAT", "Pathology"),
    ("PHA", "Pharmacy Service"),
    ("POS", "Perioperative Services"),
    ("PPA", "Parent & Patient Accommodation"),
    ("PPT", "Perinatal Pathology"),
    ("RIC", "Research & Innovation Centre (Cold Shell)"),
    ("TRA", "Travel"),
    ("BK2", "Back of House & Kitchen Stage 2"),
    ("BKS", "Back of House & Kitchen Staff Hub"),
    ("BOH", "Back of House & Staff Amenity"),
    ("FAR", "First Aid Room"),
    ("LMP", "Labour, Maternal Assessment, & Periop Entry, Wait, & Recept"),
    ("LMS", "Labour & Birth / Maternal Assessment Staff Facilities"),
    ("LPS", "Labour & Birth / Perioperative Staff Change"),
    ("MAU", "Maternal Assessment Unit"),
    ("MHS", "WNHS - Mental Health Service"),
    ("ZOO", "Not applicable"),
]


TARGET_CATEGORIES = [
    ("Spaces", BuiltInCategory.OST_MEPSpaces),
    ("Rooms", BuiltInCategory.OST_Rooms),
    ("HVAC Zones", BuiltInCategory.OST_HVAC_Zones),
    ("Ducts", BuiltInCategory.OST_DuctCurves),
    ("Duct Fittings", BuiltInCategory.OST_DuctFitting),
    ("Duct Accessories", BuiltInCategory.OST_DuctAccessory),
    ("Duct Insulations", BuiltInCategory.OST_DuctInsulations),
    ("Flex Ducts", BuiltInCategory.OST_FlexDuctCurves),
    ("Air Terminals", BuiltInCategory.OST_DuctTerminal),
    ("Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment),
    ("Pipes", BuiltInCategory.OST_PipeCurves),
    ("Pipe Fittings", BuiltInCategory.OST_PipeFitting),
    ("Pipe Accessories", BuiltInCategory.OST_PipeAccessory),
    ("Pipe Insulations", BuiltInCategory.OST_PipeInsulations),
    ("Flex Pipes", BuiltInCategory.OST_FlexPipeCurves),
    ("Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures),
    ("Sprinklers", BuiltInCategory.OST_Sprinklers),
    ("Cable Trays", BuiltInCategory.OST_CableTray),
    ("Cable Tray Fittings", BuiltInCategory.OST_CableTrayFitting),
    ("Conduits", BuiltInCategory.OST_Conduit),
    ("Conduit Fittings", BuiltInCategory.OST_ConduitFitting),
    ("Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment),
    ("Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures),
    ("Lighting Fixtures", BuiltInCategory.OST_LightingFixtures),
    ("Lighting Devices", BuiltInCategory.OST_LightingDevices),
    ("Data Devices", BuiltInCategory.OST_DataDevices),
    ("Communication Devices", BuiltInCategory.OST_CommunicationDevices),
    ("Fire Alarm Devices", BuiltInCategory.OST_FireAlarmDevices),
    ("Nurse Call Devices", BuiltInCategory.OST_NurseCallDevices),
    ("Security Devices", BuiltInCategory.OST_SecurityDevices),
    ("Telephone Devices", BuiltInCategory.OST_TelephoneDevices),
    ("Generic Models", BuiltInCategory.OST_GenericModel),
    ("Speciality Equipment", BuiltInCategory.OST_SpecialityEquipment),
    ("Casework", BuiltInCategory.OST_Casework),
    ("Furniture", BuiltInCategory.OST_Furniture),
    ("Furniture Systems", BuiltInCategory.OST_FurnitureSystems),
    ("Medical Equipment", BuiltInCategory.OST_MedicalEquipment),
]


def _storage_label(storage_type):
    if storage_type == StorageType.String:
        return "String"
    if storage_type == StorageType.Integer:
        return "Integer"
    if storage_type == StorageType.Double:
        return "Double"
    if storage_type == StorageType.ElementId:
        return "ElementId"
    return str(storage_type)


def _is_valid_element(element):
    if element is None:
        return False
    try:
        return bool(getattr(element, "IsValidObject", True))
    except Exception:
        return False


def _safe_element_id(element):
    if not _is_valid_element(element):
        return "<invalid>"
    try:
        return str(element.Id.IntegerValue)
    except Exception:
        return "<unknown>"


def _display_parameter_value(param, source_doc=None):
    if param is None:
        return "<missing>"

    src_doc = source_doc or doc

    try:
        st = param.StorageType
    except Exception:
        st = None

    try:
        if st == StorageType.String:
            raw = param.AsString()
            if raw is None or str(raw).strip() == "":
                return "<empty>"
            return str(raw)
        if st == StorageType.Integer:
            raw = param.AsInteger()
            display = param.AsValueString()
            if display is not None and str(display).strip() != "":
                return str(display)
            return str(raw)
        if st == StorageType.Double:
            display = param.AsValueString()
            if display is not None and str(display).strip() != "":
                return str(display)
            return str(param.AsDouble())
        if st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None or eid == ElementId.InvalidElementId:
                return "<empty>"
            try:
                ref_el = src_doc.GetElement(eid)
            except Exception:
                ref_el = None
            if ref_el is not None:
                try:
                    ref_name = getattr(ref_el, "Name", None)
                    if ref_name:
                        return str(ref_name)
                except Exception:
                    pass
            try:
                return str(eid.IntegerValue)
            except Exception:
                return "<unknown>"
    except Exception:
        pass

    try:
        value_string = param.AsValueString()
        if value_string is not None and str(value_string).strip() != "":
            return str(value_string)
    except Exception:
        pass

    try:
        text_value = param.AsString()
        if text_value is not None and str(text_value).strip() != "":
            return str(text_value)
    except Exception:
        pass

    return "<empty>"


def _values_match(left_value, right_value, storage_type):
    if storage_type == StorageType.String:
        left_text = "" if left_value is None else str(left_value)
        right_text = "" if right_value is None else str(right_value)
        return left_text == right_text

    if storage_type == StorageType.Integer:
        try:
            return int(left_value) == int(right_value)
        except Exception:
            return False

    if storage_type == StorageType.Double:
        try:
            return abs(float(left_value) - float(right_value)) <= 1e-6
        except Exception:
            return False

    if storage_type == StorageType.ElementId:
        try:
            return int(left_value) == int(right_value)
        except Exception:
            return False

    return str(left_value) == str(right_value)


def _coerce_readback_value(param, storage_type):
    if param is None:
        return None

    try:
        if storage_type == StorageType.String:
            raw = param.AsString()
            if raw is not None:
                return raw
            display = param.AsValueString()
            if display is not None:
                return display
            return ""
        if storage_type == StorageType.Integer:
            return param.AsInteger()
        if storage_type == StorageType.Double:
            return param.AsDouble()
        if storage_type == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None:
                return None
            return eid.IntegerValue
    except Exception:
        pass

    return _read_parameter_value(param)


def _estimate_seconds_for_collect(selected_category_count, active_view_only, selected_level_mode, entire_model):
    base = 0.5
    if entire_model:
        base += 4.0
    elif selected_level_mode:
        base += 1.5
    elif active_view_only:
        base += 0.8
    base += max(0, selected_category_count) * 0.35
    return max(0.5, base)


def _estimate_seconds_for_apply(element_count, source_mode):
    if element_count <= 0:
        return 0.5
    per_item = 0.016 if not source_mode else 0.024
    base = 0.5 + (element_count * per_item)
    return max(0.5, base)


def _format_estimate(seconds_value):
    try:
        seconds_float = float(seconds_value)
    except Exception:
        seconds_float = 0.0
    milliseconds = int(round(seconds_float * 1000.0))
    return "~{0:.1f}s ({1} ms)".format(seconds_float, milliseconds)


def _plan_string_value(existing_value, coerced_value, duplicate_mode):
    existing_text = "" if existing_value is None else str(existing_value)
    coerced_text = "" if coerced_value is None else str(coerced_value)

    if existing_text and duplicate_mode == "Skip Existing":
        return existing_text
    if existing_text and duplicate_mode == "Append (String only)":
        return existing_text + "; " + coerced_text
    return coerced_text


def _plan_expected_value(param, coerced_value, storage_type, duplicate_mode):
    existing_value = _read_parameter_value(param)
    if storage_type == StorageType.String:
        return _plan_string_value(existing_value, coerced_value, duplicate_mode)
    return coerced_value


def _confirm_with_estimate(window_owner, title, message, seconds_value):
    estimate_text = _format_estimate(seconds_value)
    full_message = "{0}\n\nEstimated run time: {1}\n\nContinue?".format(message, estimate_text)
    return window_owner._confirm(full_message, title=title)


def _load_settings():
    return {
        "selected_categories": getattr(config, "selected_categories", []),
        "collect_scope": getattr(config, "collect_scope", "Active View Only"),
        "selected_level_id": getattr(config, "selected_level_id", -1),
        "duplicate_mode": getattr(config, "duplicate_mode", "Overwrite"),
        "operation_mode": getattr(config, "operation_mode", "Direct Set Value (Current Mode)"),
        "value_transform": getattr(config, "value_transform", "Keep Original"),
        "pad_width": getattr(config, "pad_width", "2"),
        "prefix": getattr(config, "prefix", ""),
        "suffix": getattr(config, "suffix", ""),
        "replace_from": getattr(config, "replace_from", ""),
        "replace_to": getattr(config, "replace_to", ""),
        "transform_expression": getattr(config, "transform_expression", ""),
        "source_parameter": getattr(config, "source_parameter", ""),
        "existing_value_policy": getattr(config, "existing_value_policy", "Override Existing Values"),
        "export_mapped_preview_csv": getattr(config, "export_mapped_preview_csv", False),
        "mode3_source_location": getattr(config, "mode3_source_location", "Current Model"),
        "mode3_link_instance": getattr(config, "mode3_link_instance", ""),
        "mode3_target_match_parameter": getattr(config, "mode3_target_match_parameter", ""),
        "mode3_source_match_parameter": getattr(config, "mode3_source_match_parameter", ""),
    }


def _save_settings(settings):
    try:
        config.selected_categories = settings.get("selected_categories", [])
        config.collect_scope = settings.get("collect_scope", "Active View Only")
        config.selected_level_id = settings.get("selected_level_id", -1)
        config.duplicate_mode = settings.get("duplicate_mode", "Overwrite")
        config.operation_mode = settings.get("operation_mode", "Direct Set Value (Current Mode)")
        config.value_transform = settings.get("value_transform", "Keep Original")
        config.pad_width = settings.get("pad_width", "2")
        config.prefix = settings.get("prefix", "")
        config.suffix = settings.get("suffix", "")
        config.replace_from = settings.get("replace_from", "")
        config.replace_to = settings.get("replace_to", "")
        config.transform_expression = settings.get("transform_expression", "")
        config.source_parameter = settings.get("source_parameter", "")
        config.existing_value_policy = settings.get("existing_value_policy", "Override Existing Values")
        config.export_mapped_preview_csv = bool(settings.get("export_mapped_preview_csv", False))
        config.mode3_source_location = settings.get("mode3_source_location", "Current Model")
        config.mode3_link_instance = settings.get("mode3_link_instance", "")
        config.mode3_target_match_parameter = settings.get("mode3_target_match_parameter", "")
        config.mode3_source_match_parameter = settings.get("mode3_source_match_parameter", "")
        script.save_config()
    except Exception:
        pass


def _read_parameter_value(param):
    if param is None:
        return None

    try:
        has_value = bool(param.HasValue)
    except Exception:
        has_value = False

    try:
        if param.StorageType == StorageType.String:
            raw = param.AsString()
            if raw not in (None, ""):
                return raw
        if param.StorageType == StorageType.Integer:
            raw = param.AsInteger()
            if has_value or raw != 0:
                return raw
        if param.StorageType == StorageType.Double:
            raw = param.AsDouble()
            if has_value or abs(float(raw)) > 1e-12:
                return raw
        if param.StorageType == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None:
                return None
            eid_int = eid.IntegerValue
            if eid_int != ElementId.InvalidElementId.IntegerValue:
                return eid_int
    except Exception:
        pass

    try:
        value_string = param.AsValueString()
        if value_string is not None:
            value_string = str(value_string).strip()
            if value_string:
                return value_string
    except Exception:
        pass

    try:
        text_value = param.AsString()
        if text_value is not None:
            text_value = str(text_value).strip()
            if text_value:
                return text_value
    except Exception:
        pass

    return None


def _to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None


def _safe_numeric_eval(formula_text, variables):
    expr = "" if formula_text is None else str(formula_text).strip()
    if not expr:
        raise ValueError("Formula is empty.")
    if expr.startswith("="):
        expr = expr[1:].strip()
    if not expr:
        raise ValueError("Formula is empty.")

    allowed_nodes = (
        ast.Expression,
        ast.BinOp,
        ast.UnaryOp,
        ast.Num,
        ast.Constant,
        ast.Name,
        ast.Load,
        ast.Add,
        ast.Sub,
        ast.Mult,
        ast.Div,
        ast.Mod,
        ast.Pow,
        ast.FloorDiv,
        ast.USub,
        ast.UAdd,
        ast.Call,
    )
    allowed_func_names = set(["abs", "round", "min", "max"])

    tree = ast.parse(expr, mode="eval")
    for node in ast.walk(tree):
        if not isinstance(node, allowed_nodes):
            raise ValueError("Formula contains unsupported syntax.")
        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name):
                raise ValueError("Only simple function calls are allowed.")
            if node.func.id not in allowed_func_names:
                raise ValueError("Function '{0}' is not allowed.".format(node.func.id))
        if isinstance(node, ast.Name):
            if node.id not in variables and node.id not in allowed_func_names:
                raise ValueError("Unknown variable '{0}'.".format(node.id))

    safe_globals = {"__builtins__": {}}
    safe_locals = {
        "abs": abs,
        "round": round,
        "min": min,
        "max": max,
    }
    safe_locals.update(variables)
    return eval(compile(tree, "<formula>", "eval"), safe_globals, safe_locals)


def _default_dictionary_text():
    lines = []
    for key, value in DEFAULT_TRANSLATION_ITEMS:
        lines.append("{0}={1}".format(key, value))
    return "\n".join(lines)


def _parse_dictionary_entries(text_value):
    raw = "" if text_value is None else str(text_value).strip()
    entries = []

    if not raw:
        return entries

    for line in raw.splitlines():
        item = line.strip()
        if not item:
            continue
        if item.startswith("#"):
            continue

        separator = None
        if "=>" in item:
            separator = "=>"
        elif "=" in item:
            separator = "="

        if separator is None:
            continue

        key, value = item.split(separator, 1)
        key = key.strip()
        value = value.strip()
        if key:
            entries.append((key, value))

    return entries


def _translate_dictionary_value(raw_value, dictionary_text=None):
    key_text = "" if raw_value is None else str(raw_value).strip()
    if not key_text:
        return raw_value

    entries = _parse_dictionary_entries(dictionary_text)
    if not entries:
        entries = list(DEFAULT_TRANSLATION_ITEMS)

    key_lower = key_text.lower()
    for entry_key, entry_value in entries:
        if key_lower == str(entry_key).strip().lower():
            return entry_value

    return raw_value


def _coerce_value(value_text, storage_type):
    text = "" if value_text is None else str(value_text).strip()

    if storage_type == StorageType.String:
        return text

    if storage_type == StorageType.Integer:
        if not text:
            raise ValueError("Integer value is required.")
        try:
            return int(text)
        except Exception:
            return int(float(text))

    if storage_type == StorageType.Double:
        if not text:
            raise ValueError("Double value is required.")
        return float(text)

    if storage_type == StorageType.ElementId:
        if not text:
            raise ValueError("ElementId integer value is required.")
        return ElementId(int(text))

    raise ValueError("Unsupported storage type.")


def _clear_parameter(param, storage_type):
    if storage_type == StorageType.String:
        param.Set("")
        return True
    if storage_type == StorageType.Integer:
        param.Set(0)
        return True
    if storage_type == StorageType.Double:
        param.Set(0.0)
        return True
    if storage_type == StorageType.ElementId:
        param.Set(ElementId.InvalidElementId)
        return True
    return False


def _set_parameter_value(param, coerced_value, storage_type, duplicate_mode):
    existing = _read_parameter_value(param)

    if storage_type == StorageType.String:
        if existing not in (None, ""):
            if duplicate_mode == "Skip Existing":
                return False, "skipped existing"
            if duplicate_mode == "Append (String only)":
                old_text = str(existing)
                new_text = str(coerced_value)
                success = param.Set(old_text + "; " + new_text)
                return bool(success), "appended" if success else "failed"

        success = param.Set(str(coerced_value))
        return bool(success), "written" if success else "failed"

    if storage_type == StorageType.Integer:
        success = param.Set(int(coerced_value))
        return bool(success), "written" if success else "failed"

    if storage_type == StorageType.Double:
        success = param.Set(float(coerced_value))
        return bool(success), "written" if success else "failed"

    if storage_type == StorageType.ElementId:
        success = param.Set(coerced_value)
        return bool(success), "written" if success else "failed"

    return False, "unsupported storage"


def _has_effective_parameter_value(param, storage_type):
    if param is None:
        return False

    existing = _read_parameter_value(param)

    if storage_type == StorageType.String:
        return existing not in (None, "")

    if storage_type == StorageType.ElementId:
        try:
            return int(existing) != ElementId.InvalidElementId.IntegerValue
        except Exception:
            return False

    return existing is not None


def _find_writable_instance_parameter(element, parameter_name, storage_type):
    if not _is_valid_element(element) or not parameter_name:
        return None

    target_name = parameter_name.strip()
    storage_none = getattr(StorageType, "None")

    try:
        for param in element.Parameters:
            if param is None or param.Definition is None:
                continue
            if param.Definition.Name != target_name:
                continue
            if param.IsReadOnly:
                continue
            if param.StorageType == storage_none:
                continue
            if param.StorageType != storage_type:
                continue
            return param
    except Exception:
        return None

    return None


def _find_readable_instance_parameter(element, parameter_name, storage_type=None):
    if not _is_valid_element(element) or not parameter_name:
        return None

    target_name = parameter_name.strip()
    storage_none = getattr(StorageType, "None")

    try:
        for param in element.Parameters:
            if param is None or param.Definition is None:
                continue
            if param.Definition.Name != target_name:
                continue
            if param.StorageType == storage_none:
                continue
            if storage_type is not None and param.StorageType != storage_type:
                continue
            return param
    except Exception:
        return None

    return None


def _collect_instance_parameter_map(element):
    parameters = {}
    storage_none = getattr(StorageType, "None")

    if not _is_valid_element(element):
        return parameters

    try:
        for param in element.Parameters:
            if param is None or param.Definition is None:
                continue
            if param.IsReadOnly:
                continue
            if param.StorageType == storage_none:
                continue

            name = param.Definition.Name
            if not name:
                continue

            if name not in parameters:
                parameters[name] = param.StorageType
    except Exception:
        pass

    return parameters


def _collect_readable_parameter_map(element):
    parameters = {}
    storage_none = getattr(StorageType, "None")

    if not _is_valid_element(element):
        return parameters

    try:
        for param in element.Parameters:
            if param is None or param.Definition is None:
                continue
            if param.StorageType == storage_none:
                continue

            name = param.Definition.Name
            if not name:
                continue

            if name not in parameters:
                parameters[name] = param.StorageType
    except Exception:
        pass

    return parameters


def _get_element_level_id(element):
    if not _is_valid_element(element):
        return None

    invalid_id = ElementId.InvalidElementId

    try:
        lvl_id = element.LevelId
        if lvl_id is not None and lvl_id != invalid_id and lvl_id.IntegerValue > 0:
            return lvl_id.IntegerValue
    except Exception:
        pass

    param_ids = [
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.RBS_START_LEVEL_PARAM,
    ]
    for bip in param_ids:
        try:
            p = element.get_Parameter(bip)
            if p is None:
                continue
            lvl_id = p.AsElementId()
            if lvl_id is not None and lvl_id != invalid_id and lvl_id.IntegerValue > 0:
                return lvl_id.IntegerValue
        except Exception:
            continue

    return None


def _is_element_visible_in_view(element, view):
    if not _is_valid_element(element) or view is None:
        return False

    try:
        category = element.Category
    except Exception:
        category = None

    if category is None:
        return False

    try:
        if view.CanCategoryBeHidden(category.Id) and view.GetCategoryHidden(category.Id):
            return False
    except Exception:
        pass

    try:
        if element.IsHidden(view):
            return False
    except Exception:
        pass

    try:
        if getattr(element, "ViewSpecific", False):
            owner_view_id = getattr(element, "OwnerViewId", ElementId.InvalidElementId)
            if owner_view_id is not None and owner_view_id != ElementId.InvalidElementId:
                return owner_view_id.IntegerValue == view.Id.IntegerValue
    except Exception:
        pass

    return True


def _iter_category_elements(category_ids, active_view_only=True, hard_limit=50000, level_id=None):
    count = 0
    seen_ids = set()
    active_view = doc.ActiveView if active_view_only else None

    for cat_id in category_ids:
        try:
            if active_view_only:
                collector = FilteredElementCollector(doc, doc.ActiveView.Id)
            else:
                collector = FilteredElementCollector(doc)

            elems = collector.OfCategoryId(ElementId(cat_id)).WhereElementIsNotElementType()
        except Exception:
            continue

        for element in elems:
            if not _is_valid_element(element):
                continue

            if active_view_only and not _is_element_visible_in_view(element, active_view):
                continue

            if level_id is not None and _get_element_level_id(element) != int(level_id):
                continue

            eid = element.Id.IntegerValue
            if eid in seen_ids:
                continue
            seen_ids.add(eid)
            yield element

            count += 1
            if count >= hard_limit:
                return


def _iter_category_elements_in_document(source_doc, category_ids, hard_limit=50000):
    if source_doc is None:
        return

    count = 0
    seen_ids = set()

    for cat_id in category_ids:
        try:
            collector = FilteredElementCollector(source_doc)
            elems = collector.OfCategoryId(ElementId(cat_id)).WhereElementIsNotElementType()
        except Exception:
            continue

        for element in elems:
            if not _is_valid_element(element):
                continue
            try:
                eid = element.Id.IntegerValue
            except Exception:
                continue
            if eid in seen_ids:
                continue
            seen_ids.add(eid)
            yield element

            count += 1
            if count >= hard_limit:
                return


def _iter_all_instance_elements_in_document(source_doc, hard_limit=50000):
    if source_doc is None:
        return

    count = 0
    seen_ids = set()
    try:
        collector = FilteredElementCollector(source_doc).WhereElementIsNotElementType()
    except Exception:
        return

    for element in collector:
        if not _is_valid_element(element):
            continue
        try:
            cat = element.Category
            if cat is None:
                continue
            eid = element.Id.IntegerValue
        except Exception:
            continue
        if eid in seen_ids:
            continue
        seen_ids.add(eid)
        yield element
        count += 1
        if count >= hard_limit:
            return


class CategoryInstanceParameterWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        self._fit_to_screen()

        self._ui_ready = False
        self._is_busy = False
        self._busy_action = ""
        self.category_items = []
        self.level_items = []
        self.selected_elements = []
        self.common_param_map = {}
        self.source_param_map = {}
        self.mode3_link_map = {}
        self.mode3_source_elements = []
        self.mode3_source_index = {}
        self._suppress_mode3_events = False
        self._settings = _load_settings()

        self._build_level_list()
        self._build_category_list()
        if not hasattr(self, "txtExpression"):
            self.txtExpression = None
        self._apply_saved_settings()
        self._update_scope_dependent_ui()
        self._update_mode_dependent_ui()
        self._update_transform_hint()
        self._refresh_element_summary()
        self._refresh_parameter_controls()
        self._maximize_to_work_area()
        self._ui_ready = True

    def _maximize_to_work_area(self):
        try:
            work_area = System.Windows.SystemParameters.WorkArea
            self.Left = float(work_area.Left)
            self.Top = float(work_area.Top)
            self.Width = float(work_area.Width)
            self.Height = float(work_area.Height)
            self.WindowState = System.Windows.WindowState.Maximized
        except Exception:
            try:
                self.WindowState = System.Windows.WindowState.Maximized
            except Exception:
                pass

    def _begin_busy(self, action_name):
        if self._is_busy:
            forms.alert(
                "Please wait. '{0}' is still running.".format(self._busy_action or "Operation"),
                title="Category Instance Parameters"
            )
            return False
        self._is_busy = True
        self._busy_action = action_name or "Operation"
        return True

    def _end_busy(self):
        self._is_busy = False
        self._busy_action = ""

    def _fit_to_screen(self):
        try:
            work_area = System.Windows.SystemParameters.WorkArea
            max_height = max(560.0, float(work_area.Height) - 20.0)
            max_width = max(920.0, float(work_area.Width) - 20.0)

            self.MaxHeight = max_height
            self.MaxWidth = max_width

            if float(self.Height) > max_height:
                self.Height = max_height
            if float(self.Width) > max_width:
                self.Width = max_width
        except Exception:
            pass

    def window_loaded(self, sender, e):
        try:
            self._maximize_to_work_area()
            self._apply_responsive_panel_heights()
        except Exception:
            pass

    def window_size_changed(self, sender, e):
        try:
            self._apply_responsive_panel_heights()
        except Exception:
            pass

    def _apply_responsive_panel_heights(self):
        try:
            h = float(getattr(self, "ActualHeight", 0.0) or 0.0)
        except Exception:
            h = 0.0
        if h <= 0.0:
            return

        def _clamp(v, lo, hi):
            try:
                return max(float(lo), min(float(hi), float(v)))
            except Exception:
                return float(lo)

        list_h = _clamp((h * 0.26), 220.0, 520.0)
        preview_h = _clamp((h * 0.20), 120.0, 360.0)

        for name, val in [
            ("lstCategories", list_h),
            ("lstCommonParams", list_h),
            ("lstSourcePreview", preview_h),
        ]:
            try:
                ctrl = getattr(self, name, None)
                if ctrl is not None:
                    ctrl.Height = val
            except Exception:
                pass

    def _confirm(self, message, title="Confirm"):
        try:
            result = forms.alert(message, title=title, yes=True, no=True)
            if isinstance(result, bool):
                return result
            text = (str(result) if result is not None else "").strip().lower()
            return text in ("yes", "y", "true", "ok", "1")
        except Exception:
            return False

    def _show_ui_error(self, action_name, ex):
        try:
            logger.exception("Category Instance Parameters UI error during %s", action_name)
        except Exception:
            pass
        forms.alert("Action failed: {0}\n\n{1}".format(action_name, ex), title="Category Instance Parameters")

    def _build_category_list(self):
        selected = set(self._settings.get("selected_categories", []))
        self.category_items = []

        for name, bic in TARGET_CATEGORIES:
            cb = CheckBox()
            cb.Content = name
            cb.IsChecked = True if not selected else (int(bic) in selected)

            self.category_items.append({
                "name": name,
                "bic": int(bic),
                "cb": cb,
            })

        self._refresh_category_list("")

    def _apply_saved_settings(self):
        scope_name = self._settings.get("collect_scope", "Active View Only")
        if "Selected Level" in scope_name:
            self.cmbCollectScope.SelectedIndex = 1
        elif "Entire" in scope_name:
            self.cmbCollectScope.SelectedIndex = 2
        else:
            self.cmbCollectScope.SelectedIndex = 0

        saved_level_id = int(self._settings.get("selected_level_id", -1) or -1)
        if self.level_items:
            selected_idx = -1
            for idx, item in enumerate(self.level_items):
                if item["id"] == saved_level_id:
                    selected_idx = idx
                    break
            if selected_idx < 0:
                selected_idx = 0
            self.cmbSelectLevel.SelectedIndex = selected_idx

        mode_name = self._settings.get("duplicate_mode", "Overwrite")
        names = ["Overwrite", "Skip Existing", "Append (String only)"]
        if mode_name in names:
            self.cmbDuplicateMode.SelectedIndex = names.index(mode_name)
        else:
            self.cmbDuplicateMode.SelectedIndex = 0

        operation_mode = self._settings.get("operation_mode", "Direct Set Value (Current Mode)")
        if "Matched Element" in operation_mode:
            self.cmbOperationMode.SelectedIndex = 2
        elif "Copy/Process" in operation_mode:
            self.cmbOperationMode.SelectedIndex = 1
        else:
            self.cmbOperationMode.SelectedIndex = 0

        mode3_source_location = self._settings.get("mode3_source_location", "Current Model")
        mode3_locations = ["Current Model", "Linked Model"]
        if mode3_source_location in mode3_locations:
            self.cmbMode3SourceLocation.SelectedIndex = mode3_locations.index(mode3_source_location)
        else:
            self.cmbMode3SourceLocation.SelectedIndex = 0

        transform_mode = self._settings.get("value_transform", "Keep Original")
        transform_names = [
            "Keep Original",
            "Convert To String",
            "Digits Only",
            "Pad Number (Width)",
            "Digits + Pad Number",
            "Level Number (2-digit)",
            "Level Number (3-digit)",
            "Uppercase",
            "Lowercase",
            "Title Case",
            "Trim Spaces",
            "Remove Spaces",
            "Left 3 Characters",
            "Dictionary (Translate)",
            "Template (Tokens)",
            "Formula (Safe Numeric)",
        ]
        if transform_mode in transform_names:
            self.cmbValueTransform.SelectedIndex = transform_names.index(transform_mode)
        else:
            self.cmbValueTransform.SelectedIndex = 0

        self.txtPadWidth.Text = str(self._settings.get("pad_width", "2"))
        self.txtPrefix.Text = str(self._settings.get("prefix", ""))
        self.txtSuffix.Text = str(self._settings.get("suffix", ""))
        self.txtReplaceFrom.Text = str(self._settings.get("replace_from", ""))
        self.txtReplaceTo.Text = str(self._settings.get("replace_to", ""))
        self.txtExpression.Text = str(self._settings.get("transform_expression", ""))
        if self._selected_transform_mode() == "Dictionary (Translate)" and not (self.txtExpression.Text or "").strip():
            self.txtExpression.Text = _default_dictionary_text()

        try:
            self.chkExportMappedPreviewCsv.IsChecked = bool(self._settings.get("export_mapped_preview_csv", False))
        except Exception:
            pass

        existing_policy = self._settings.get("existing_value_policy", "Override Existing Values")
        policy_names = [
            "Override Existing Values",
            "Skip Elements With Existing Value",
        ]
        if existing_policy in policy_names:
            self.cmbExistingValueHandling.SelectedIndex = policy_names.index(existing_policy)
        else:
            self.cmbExistingValueHandling.SelectedIndex = 0

    def _save_current_settings(self):
        selected_source = ""
        try:
            if self.cmbSourceParameter is not None and self.cmbSourceParameter.SelectedItem is not None:
                selected_source = str(self.cmbSourceParameter.SelectedItem)
        except Exception:
            selected_source = ""

        settings = {
            "selected_categories": self._selected_category_ids(),
            "collect_scope": self._selected_scope(),
            "selected_level_id": self._selected_level_id(),
            "duplicate_mode": self._selected_duplicate_mode(),
            "operation_mode": self._selected_operation_mode(),
            "value_transform": self._selected_transform_mode(),
            "pad_width": self.txtPadWidth.Text if self.txtPadWidth is not None else "2",
            "prefix": self.txtPrefix.Text if self.txtPrefix is not None else "",
            "suffix": self.txtSuffix.Text if self.txtSuffix is not None else "",
            "replace_from": self.txtReplaceFrom.Text if self.txtReplaceFrom is not None else "",
            "replace_to": self.txtReplaceTo.Text if self.txtReplaceTo is not None else "",
            "transform_expression": self.txtExpression.Text if self.txtExpression is not None else "",
            "source_parameter": selected_source,
            "existing_value_policy": self._selected_existing_value_policy(),
            "export_mapped_preview_csv": bool(self.chkExportMappedPreviewCsv.IsChecked) if hasattr(self, "chkExportMappedPreviewCsv") else False,
            "mode3_source_location": self._selected_mode3_source_location(),
            "mode3_link_instance": self._selected_combo_item_text(self.cmbMode3LinkInstance),
            "mode3_target_match_parameter": self._selected_combo_item_text(self.cmbMode3TargetMatchParameter),
            "mode3_source_match_parameter": self._selected_combo_item_text(self.cmbMode3SourceMatchParameter),
        }
        _save_settings(settings)
        try:
            self._settings.update(settings)
        except Exception:
            pass

    def _build_level_list(self):
        self.level_items = []
        self.cmbSelectLevel.Items.Clear()

        try:
            levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        except Exception:
            levels = []

        sortable = []
        for lvl in levels:
            try:
                sortable.append((float(lvl.Elevation), lvl))
            except Exception:
                sortable.append((0.0, lvl))

        for _, lvl in sorted(sortable, key=lambda x: (x[0], x[1].Name.lower() if hasattr(x[1], "Name") else "")):
            try:
                item = {"id": lvl.Id.IntegerValue, "name": lvl.Name}
            except Exception:
                continue
            self.level_items.append(item)
            self.cmbSelectLevel.Items.Add(item["name"])

    def _selected_level_id(self):
        idx = getattr(self.cmbSelectLevel, "SelectedIndex", -1)
        if idx is None or idx < 0 or idx >= len(self.level_items):
            return -1
        return int(self.level_items[idx]["id"])

    def _selected_scope(self):
        try:
            selected_item = self.cmbCollectScope.SelectedItem
            if selected_item is None:
                return "Active View Only"
            text = str(selected_item.Content) if hasattr(selected_item, "Content") else str(selected_item)
            return text
        except Exception:
            return "Active View Only"

    def _selected_duplicate_mode(self):
        try:
            selected_item = self.cmbDuplicateMode.SelectedItem
            if selected_item is None:
                return "Overwrite"
            text = str(selected_item.Content) if hasattr(selected_item, "Content") else str(selected_item)
            return text
        except Exception:
            return "Overwrite"

    def _selected_operation_mode(self):
        try:
            selected_item = self.cmbOperationMode.SelectedItem
            if selected_item is None:
                return "Direct Set Value (Current Mode)"
            text = str(selected_item.Content) if hasattr(selected_item, "Content") else str(selected_item)
            return text
        except Exception:
            return "Direct Set Value (Current Mode)"

    def _selected_combo_item_text(self, combo):
        try:
            selected = combo.SelectedItem
            if selected is None:
                return ""
            return str(selected.Content) if hasattr(selected, "Content") else str(selected)
        except Exception:
            return ""

    def _select_combo_item_by_text(self, combo, text_value):
        if combo is None:
            return
        target = "" if text_value is None else str(text_value)
        if not target:
            if combo.Items.Count > 0:
                combo.SelectedIndex = 0
            return
        for idx in range(combo.Items.Count):
            item = combo.Items[idx]
            current = str(item.Content) if hasattr(item, "Content") else str(item)
            if current == target:
                combo.SelectedIndex = idx
                return
        if combo.Items.Count > 0:
            combo.SelectedIndex = 0

    def _selected_mode3_source_location(self):
        text = self._selected_combo_item_text(self.cmbMode3SourceLocation)
        return text or "Current Model"

    def _is_mode3_source_mode(self):
        return "Matched Element" in self._selected_operation_mode()

    def _is_mode2_source_mode(self):
        mode_name = self._selected_operation_mode()
        return "Copy/Process" in mode_name and "Matched Element" not in mode_name

    def _selected_existing_value_policy(self):
        try:
            selected_item = self.cmbExistingValueHandling.SelectedItem
            if selected_item is None:
                return "Override Existing Values"
            text = str(selected_item.Content) if hasattr(selected_item, "Content") else str(selected_item)
            return text
        except Exception:
            return "Override Existing Values"

    def _is_source_mode(self):
        return self._is_mode2_source_mode() or self._is_mode3_source_mode()

    def _selected_transform_mode(self):
        try:
            selected_item = self.cmbValueTransform.SelectedItem
            if selected_item is None:
                return "Keep Original"
            text = str(selected_item.Content) if hasattr(selected_item, "Content") else str(selected_item)
            return text
        except Exception:
            return "Keep Original"

    def _get_apply_chunk_size(self, source_mode=False):
        # Conservative chunking keeps transactions short and reduces UI freeze risk.
        return 220 if source_mode else 380

    def _confirm_target_coverage(self, parameter_name, storage_type):
        if not parameter_name or not self.selected_elements:
            return True

        writable = 0
        missing = 0
        readonly_or_blocked = 0

        for element in self.selected_elements:
            if not _is_valid_element(element):
                continue
            p = _find_writable_instance_parameter(element, parameter_name, storage_type)
            if p is not None:
                writable += 1
                continue

            p_read = _find_readable_instance_parameter(element, parameter_name, storage_type)
            if p_read is None:
                missing += 1
            else:
                readonly_or_blocked += 1

        if missing == 0 and readonly_or_blocked == 0:
            return True

        lines = [
            "Target parameter coverage warning:",
            "Some selected elements do not expose the target as writable instance parameter.",
            "Those elements will be skipped.",
            "",
            "- Target: {0}".format(parameter_name),
            "- Writable: {0}".format(writable),
            "- Missing: {0}".format(missing),
            "- Read-only/blocked: {0}".format(readonly_or_blocked),
            "",
            "Continue anyway?",
        ]
        return self._confirm("\n".join(lines), title="Target Coverage Warning")

    def _update_scope_dependent_ui(self):
        panel = getattr(self, "pnlSelectedLevel", None)
        if panel is None:
            return
        scope_name = self._selected_scope()
        show_level = "Selected Level" in scope_name
        panel.Visibility = System.Windows.Visibility.Visible if show_level else System.Windows.Visibility.Collapsed

    def _update_mode_dependent_ui(self):
        show_source_mode = self._is_source_mode()
        self.pnlDirectValueMode.Visibility = System.Windows.Visibility.Collapsed if show_source_mode else System.Windows.Visibility.Visible
        self.pnlSourceProcessMode.Visibility = System.Windows.Visibility.Visible if show_source_mode else System.Windows.Visibility.Collapsed

        is_mode3 = self._is_mode3_source_mode()
        self.pnlMode3SourceOptions.Visibility = System.Windows.Visibility.Visible if is_mode3 else System.Windows.Visibility.Collapsed
        self.pnlMode3LinkSelection.Visibility = System.Windows.Visibility.Visible if (is_mode3 and self._selected_mode3_source_location() == "Linked Model") else System.Windows.Visibility.Collapsed

        if show_source_mode:
            self._refresh_source_parameter_controls()

    def _update_transform_hint(self):
        transform = self._selected_transform_mode()
        self.txtPadWidth.IsEnabled = (
            "Pad Number" in transform
            or "Level Number (2-digit)" in transform
            or "Level Number (3-digit)" in transform
        )
        self.txtExpression.IsEnabled = (
            transform == "Dictionary (Translate)"
            or transform == "Template (Tokens)"
            or transform == "Formula (Safe Numeric)"
        )
        self.txtTransformHint.Text = (
            "Step 3: Click Apply Update to write processed source values into the selected target parameter. "
            "Template tokens: {value}, {id}, {category}, {level_name}, {level_number}. "
            "Dictionary entries: code=value per line. Formula vars: value or v (supports + - * / and abs/round/min/max)."
        )
        if transform == "Dictionary (Translate)":
            current_text = (self.txtExpression.Text or "").strip() if self.txtExpression is not None else ""
            if not current_text:
                self.txtExpression.Text = _default_dictionary_text()

    def _refresh_category_list(self, search_text):
        self.lstCategories.Items.Clear()

        query = (search_text or "").strip().lower()
        for item in self.category_items:
            if query and query not in item["name"].lower():
                continue
            self.lstCategories.Items.Add(item["cb"])

    def _selected_category_ids(self):
        ids = []
        for item in self.category_items:
            try:
                if item["cb"].IsChecked:
                    ids.append(item["bic"])
            except Exception:
                continue
        return ids

    def _refresh_element_summary(self):
        if not self.selected_elements:
            self.txtElementSummary.Text = "No elements collected yet."
            return

        by_cat = defaultdict(int)
        for element in self.selected_elements:
            if not _is_valid_element(element):
                continue
            try:
                cat_name = element.Category.Name if element.Category else "<No Category>"
            except Exception:
                cat_name = "<No Category>"
            by_cat[cat_name] += 1

        lines = ["Collected elements: {0}".format(len(self.selected_elements))]
        for cat_name in sorted(by_cat.keys()):
            lines.append("- {0}: {1}".format(cat_name, by_cat[cat_name]))

        self.txtElementSummary.Text = "\n".join(lines)

    def _refresh_parameter_controls(self):
        self.lstCommonParams.Items.Clear()
        self.cmbParameter.Items.Clear()
        self.common_param_map = {}

        self.lstSourcePreview.Items.Clear()

        if not self.selected_elements:
            self.txtSelectedParameterInfo.Text = "No parameter selected."
            if self._is_source_mode():
                self._refresh_source_parameter_controls()
            else:
                self.cmbSourceParameter.Items.Clear()
            return

        param_maps = [_collect_instance_parameter_map(el) for el in self.selected_elements]
        if not param_maps:
            self.txtSelectedParameterInfo.Text = "No parameter selected."
            if self._is_source_mode():
                self._refresh_source_parameter_controls()
            else:
                self.cmbSourceParameter.Items.Clear()
            return

        common = dict(param_maps[0])
        for pmap in param_maps[1:]:
            next_common = {}
            for name, storage in common.items():
                if name in pmap and pmap[name] == storage:
                    next_common[name] = storage
            common = next_common
            if not common:
                break

        self.common_param_map = common

        for param_name in sorted(common.keys()):
            storage_label = _storage_label(common[param_name])
            self.lstCommonParams.Items.Add("{0} | {1}".format(param_name, storage_label))
            self.cmbParameter.Items.Add(param_name)

        if self.cmbParameter.Items.Count > 0:
            self.cmbParameter.SelectedIndex = 0
        else:
            self.txtSelectedParameterInfo.Text = "No common writable instance parameters found for this element set."

        if self._is_source_mode():
            self._refresh_source_parameter_controls()
        else:
            self.cmbSourceParameter.Items.Clear()

    def _refresh_source_parameter_controls(self):
        self.cmbSourceParameter.Items.Clear()
        self.source_param_map = {}

        if self._is_mode3_source_mode():
            self._refresh_mode3_source_controls()
            return

        self.mode3_source_elements = []
        self.mode3_source_index = {}
        self.mode3_link_map = {}
        self.cmbMode3LinkInstance.Items.Clear()
        self.cmbMode3TargetMatchParameter.Items.Clear()
        self.cmbMode3SourceMatchParameter.Items.Clear()
        self.txtMode3Status.Text = "Mode 3 not active."

        # Avoid extra heavy scans unless source processing mode is active.
        if not self._is_mode2_source_mode():
            return

        if not self.selected_elements:
            return

        readable_maps = [_collect_readable_parameter_map(el) for el in self.selected_elements]
        if not readable_maps:
            return

        common = dict(readable_maps[0])
        for pmap in readable_maps[1:]:
            next_common = {}
            for name, storage in common.items():
                if name in pmap and pmap[name] == storage:
                    next_common[name] = storage
            common = next_common
            if not common:
                break

        for name in sorted(common.keys()):
            storage = common[name]
            display = "{0} | {_storage}".format(name, _storage=_storage_label(storage))
            self.source_param_map[display] = {
                "kind": "parameter",
                "name": name,
                "storage": storage,
            }
            self.cmbSourceParameter.Items.Add(display)

        virtuals = [
            ("Virtual | Level Name", {"kind": "virtual", "id": "level_name", "storage": StorageType.String}),
            ("Virtual | Level Number", {"kind": "virtual", "id": "level_number", "storage": StorageType.String}),
            ("Virtual | Level Elevation", {"kind": "virtual", "id": "level_elevation", "storage": StorageType.Double}),
        ]
        for display, meta in virtuals:
            self.source_param_map[display] = meta
            self.cmbSourceParameter.Items.Add(display)

        saved_source = self._settings.get("source_parameter", "")
        idx = -1
        for i in range(self.cmbSourceParameter.Items.Count):
            if str(self.cmbSourceParameter.Items[i]) == str(saved_source):
                idx = i
                break
        if idx < 0 and self.cmbSourceParameter.Items.Count > 0:
            idx = 0
        if idx >= 0:
            self.cmbSourceParameter.SelectedIndex = idx

    def _build_common_readable_map(self, elements):
        readable_maps = [_collect_readable_parameter_map(el) for el in elements if _is_valid_element(el)]
        if not readable_maps:
            return {}

        common = dict(readable_maps[0])
        for pmap in readable_maps[1:]:
            next_common = {}
            for name, storage in common.items():
                if name in pmap and pmap[name] == storage:
                    next_common[name] = storage
            common = next_common
            if not common:
                break
        return common

    def _build_readable_union_map(self, elements, scan_limit=1500):
        union_map = {}
        if not elements:
            return union_map

        count = 0
        for element in elements:
            if not _is_valid_element(element):
                continue
            readable = _collect_readable_parameter_map(element)
            for name, storage in readable.items():
                if name not in union_map:
                    union_map[name] = storage
            count += 1
            if count >= scan_limit:
                break

        return union_map

    def _refresh_mode3_link_instances(self, preferred_link_name=""):
        self.cmbMode3LinkInstance.Items.Clear()
        self.mode3_link_map = {}

        try:
            link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
        except Exception:
            link_instances = []

        for link in link_instances:
            if not _is_valid_element(link):
                continue
            try:
                link_doc = link.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc is None:
                continue
            try:
                link_name = "{0} | {1}".format(link.Name, link_doc.Title)
            except Exception:
                link_name = "Link-{0}".format(_safe_element_id(link))
            if link_name in self.mode3_link_map:
                continue
            self.mode3_link_map[link_name] = link
            self.cmbMode3LinkInstance.Items.Add(link_name)

        link_name = preferred_link_name or self._selected_combo_item_text(self.cmbMode3LinkInstance) or self._settings.get("mode3_link_instance", "")
        self._select_combo_item_by_text(self.cmbMode3LinkInstance, link_name)

    def _collect_mode3_source_elements(self):
        location = self._selected_mode3_source_location()
        category_ids = self._selected_category_ids()
        if not category_ids:
            category_ids = []

        if location == "Linked Model":
            selected_link_name = self._selected_combo_item_text(self.cmbMode3LinkInstance)
            selected_link = self.mode3_link_map.get(selected_link_name)
            if selected_link is None:
                return []
            try:
                linked_doc = selected_link.GetLinkDocument()
            except Exception:
                linked_doc = None
            if linked_doc is None:
                return []
            scoped = []
            if category_ids:
                scoped = list(_iter_category_elements_in_document(linked_doc, category_ids, hard_limit=MAX_COLLECT_LIMIT))
            if scoped:
                return scoped
            # Fallback for heterogeneous linked models where selected categories have no candidates.
            return list(_iter_all_instance_elements_in_document(linked_doc, hard_limit=min(MAX_COLLECT_LIMIT, 20000)))

        scope_name = self._selected_scope()
        active_view_only = "Active" in scope_name
        selected_level_mode = "Selected Level" in scope_name
        selected_level_id = self._selected_level_id() if selected_level_mode else None

        source_elements = []
        if category_ids:
            source_elements = list(
                _iter_category_elements(
                    category_ids,
                    active_view_only=active_view_only,
                    hard_limit=MAX_COLLECT_LIMIT,
                    level_id=selected_level_id,
                )
            )

        if not source_elements:
            # Fallback to broader host scan to keep Mode 3 Step 1 usable when category filters are too narrow.
            try:
                if active_view_only:
                    collector = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType()
                else:
                    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
            except Exception:
                collector = []

            broad = []
            for element in collector:
                if not _is_valid_element(element):
                    continue
                try:
                    if element.Category is None:
                        continue
                except Exception:
                    continue
                if selected_level_id is not None and _get_element_level_id(element) != int(selected_level_id):
                    continue
                broad.append(element)
                if len(broad) >= min(MAX_COLLECT_LIMIT, 20000):
                    break
            source_elements = broad

        target_ids = set()
        for target in self.selected_elements:
            try:
                target_ids.add(target.Id.IntegerValue)
            except Exception:
                continue
        filtered = []
        for element in source_elements:
            try:
                if element.Id.IntegerValue in target_ids:
                    continue
            except Exception:
                pass
            filtered.append(element)
        return filtered

    def _refresh_mode3_source_controls(self):
        if self._suppress_mode3_events:
            return
        self._suppress_mode3_events = True
        try:
            current_link_name = self._selected_combo_item_text(self.cmbMode3LinkInstance)
            current_target_match = self._selected_combo_item_text(self.cmbMode3TargetMatchParameter)
            current_source_match = self._selected_combo_item_text(self.cmbMode3SourceMatchParameter)
            current_source_param = self._selected_combo_item_text(self.cmbSourceParameter)

            self._refresh_mode3_link_instances(preferred_link_name=current_link_name)
            self.pnlMode3LinkSelection.Visibility = System.Windows.Visibility.Visible if self._selected_mode3_source_location() == "Linked Model" else System.Windows.Visibility.Collapsed

            self.cmbMode3TargetMatchParameter.Items.Clear()
            self.cmbMode3SourceMatchParameter.Items.Clear()

            if not self.selected_elements:
                self.txtMode3Status.Text = "Collect target elements first to configure Mode 3."
                return

            target_union_readable = self._build_readable_union_map(self.selected_elements)
            for name in sorted(target_union_readable.keys()):
                self.cmbMode3TargetMatchParameter.Items.Add(name)
            self._select_combo_item_by_text(
                self.cmbMode3TargetMatchParameter,
                current_target_match or self._settings.get("mode3_target_match_parameter", ""),
            )

            self.mode3_source_elements = self._collect_mode3_source_elements()
            if not self.mode3_source_elements:
                self.txtMode3Status.Text = "No Mode 3 source elements found. Check source model/link and selected categories."
                return

            source_union_readable = self._build_readable_union_map(self.mode3_source_elements)
            self.cmbSourceParameter.Items.Clear()
            self.source_param_map = {}
            for name in sorted(source_union_readable.keys()):
                storage = source_union_readable[name]
                display = "{0} | {_storage}".format(name, _storage=_storage_label(storage))
                self.source_param_map[display] = {
                    "kind": "parameter",
                    "name": name,
                    "storage": storage,
                }
                self.cmbSourceParameter.Items.Add(display)

            virtuals = [
                ("Virtual | Level Name", {"kind": "virtual", "id": "level_name", "storage": StorageType.String}),
                ("Virtual | Level Number", {"kind": "virtual", "id": "level_number", "storage": StorageType.String}),
                ("Virtual | Level Elevation", {"kind": "virtual", "id": "level_elevation", "storage": StorageType.Double}),
            ]
            for display, meta in virtuals:
                self.source_param_map[display] = meta
                self.cmbSourceParameter.Items.Add(display)

            self._select_combo_item_by_text(
                self.cmbSourceParameter,
                current_source_param or self._settings.get("source_parameter", ""),
            )

            for name in sorted(source_union_readable.keys()):
                self.cmbMode3SourceMatchParameter.Items.Add(name)
            self._select_combo_item_by_text(
                self.cmbMode3SourceMatchParameter,
                current_source_match or self._settings.get("mode3_source_match_parameter", ""),
            )

            if self.cmbMode3TargetMatchParameter.Items.Count == 0:
                self.txtMode3Status.Text = "No readable target match parameters found."
                return
            if self.cmbMode3SourceMatchParameter.Items.Count == 0:
                self.txtMode3Status.Text = "No readable source match parameters found."
                return

            self._rebuild_mode3_source_index()
        finally:
            self._suppress_mode3_events = False

    def _rebuild_mode3_source_index(self):
        self.mode3_source_index = {}

        source_match_name = self._selected_combo_item_text(self.cmbMode3SourceMatchParameter)
        if not source_match_name:
            self.txtMode3Status.Text = "Choose a source match key parameter."
            return

        source_storage = None
        for source in self.mode3_source_elements:
            param = _find_readable_instance_parameter(source, source_match_name)
            if param is not None:
                source_storage = param.StorageType
                break

        indexed = 0
        duplicates = 0
        for source in self.mode3_source_elements:
            source_param = _find_readable_instance_parameter(source, source_match_name, source_storage)
            if source_param is None:
                continue
            raw_key = _read_parameter_value(source_param)
            normalized = _normalize_match_key(raw_key)
            if not normalized:
                continue
            if normalized in self.mode3_source_index:
                duplicates += 1
                continue
            self.mode3_source_index[normalized] = source
            indexed += 1

        self.txtMode3Status.Text = "Mode 3 source candidates: {0} | Indexed keys: {1} | Duplicates ignored: {2}".format(
            len(self.mode3_source_elements),
            indexed,
            duplicates,
        )

    def _ensure_mode3_ready(self):
        if not self._is_mode3_source_mode():
            return True

        if not self.mode3_source_elements:
            self._refresh_mode3_source_controls()

        target_match = self._selected_combo_item_text(self.cmbMode3TargetMatchParameter)
        source_match = self._selected_combo_item_text(self.cmbMode3SourceMatchParameter)
        if not target_match or not source_match:
            forms.alert("Mode 3 requires both target match key and source match key parameters.", title="Category Instance Parameters")
            return False

        self._rebuild_mode3_source_index()
        if not self.mode3_source_index:
            forms.alert("Mode 3 could not build a source match index. Check match keys and source model selection.", title="Category Instance Parameters")
            return False

        return True

    def _get_source_meta(self):
        selected = self.cmbSourceParameter.SelectedItem
        if selected is None:
            return None
        return self.source_param_map.get(str(selected))

    def _get_source_display_name(self):
        source_meta = self._get_source_meta()
        if source_meta is None:
            return "Parameter"

        if source_meta.get("kind") == "parameter":
            return source_meta.get("name", "Parameter")

        virtual_id = source_meta.get("id")
        if virtual_id == "level_name":
            return "Level Name"
        if virtual_id == "level_number":
            return "Level Number"
        if virtual_id == "level_elevation":
            return "Level Elevation"
        return "Parameter"

    def _extract_level_number_text(self, level_name):
        text = "" if level_name is None else str(level_name)
        match = re.search(r"-?\d+", text)
        if not match:
            return ""
        return match.group(0)

    def _read_source_value(self, element, source_meta):
        if element is None or source_meta is None:
            return None

        if source_meta.get("kind") == "parameter":
            source_param = _find_readable_instance_parameter(
                element,
                source_meta.get("name", ""),
                source_meta.get("storage"),
            )
            if source_param is None:
                source_param = _find_readable_instance_parameter(
                    element,
                    source_meta.get("name", ""),
                    None,
                )
            if source_param is None:
                return None
            return _read_parameter_value(source_param)

        virtual_id = source_meta.get("id")
        lvl_id = _get_element_level_id(element)
        if lvl_id is None:
            return None

        lvl = doc.GetElement(ElementId(int(lvl_id)))
        if lvl is None:
            return None

        if virtual_id == "level_name":
            try:
                return lvl.Name
            except Exception:
                return None

        if virtual_id == "level_number":
            try:
                return self._extract_level_number_text(lvl.Name)
            except Exception:
                return None

        if virtual_id == "level_elevation":
            try:
                return float(lvl.Elevation)
            except Exception:
                return None

        return None

    def _read_mode3_source_value(self, target_element, source_meta):
        if target_element is None or source_meta is None:
            return None, None, "Missing target element or source parameter."

        target_match_name = self._selected_combo_item_text(self.cmbMode3TargetMatchParameter)
        if not target_match_name:
            return None, None, "Choose a target match key parameter."

        target_match_param = _find_readable_instance_parameter(target_element, target_match_name)
        if target_match_param is None:
            return None, None, "Target match key parameter not found on element."

        target_key_raw = _read_parameter_value(target_match_param)
        target_key = _normalize_match_key(target_key_raw)
        if not target_key:
            return None, None, "Target match key is empty."

        source_element = self.mode3_source_index.get(target_key)
        if source_element is None:
            return None, None, "No source element matched key '{0}'.".format(target_key_raw)

        source_raw = self._read_source_value(source_element, source_meta)
        return source_raw, source_element, None

    def _get_pad_width(self):
        try:
            width = int((self.txtPadWidth.Text or "2").strip())
        except Exception:
            width = 2
        return max(1, min(8, width))

    def _build_transform_context(self, element, raw_value):
        category_name = ""
        try:
            if element is not None and element.Category is not None and element.Category.Name:
                category_name = str(element.Category.Name)
        except Exception:
            category_name = ""

        level_name = ""
        level_number = ""
        lvl_id = _get_element_level_id(element)
        if lvl_id is not None:
            try:
                lvl = doc.GetElement(ElementId(int(lvl_id)))
            except Exception:
                lvl = None
            if lvl is not None:
                try:
                    level_name = str(lvl.Name or "")
                except Exception:
                    level_name = ""
                level_number = self._extract_level_number_text(level_name)

        element_id_text = ""
        try:
            if element is not None and element.Id is not None:
                element_id_text = str(element.Id.IntegerValue)
        except Exception:
            element_id_text = ""

        raw_text = "" if raw_value is None else str(raw_value)
        return {
            "value": raw_text,
            "id": element_id_text,
            "category": category_name,
            "level_name": level_name,
            "level_number": level_number,
        }

    def _apply_template_expression(self, template_text, context):
        template = "" if template_text is None else str(template_text)
        result = template
        for key, value in context.items():
            token = "{" + str(key) + "}"
            result = result.replace(token, "" if value is None else str(value))
        return result

    def _format_preview_value(self, value):
        if value is None:
            return "<missing>"
        text = str(value)
        if text == "":
            return "<empty>"
        return text

    def _selected_category_names(self):
        names = []
        for item in self.category_items:
            try:
                if item["cb"].IsChecked:
                    names.append(str(item.get("name", "")))
            except Exception:
                continue
        return names

    def _build_mapped_preview_filename(self):
        scope_name = self._selected_scope()
        scope_tag = _sanitize_filename_part(scope_name, max_len=28)

        if "Active View" in scope_name:
            try:
                view_name = doc.ActiveView.Name if doc is not None and doc.ActiveView is not None else "ActiveView"
            except Exception:
                view_name = "ActiveView"
            scope_tag = "ActiveView-{0}".format(_sanitize_filename_part(view_name, max_len=36))

        selected_categories = self._selected_category_names()
        if selected_categories:
            cat_tokens = [_sanitize_filename_part(name, max_len=16) for name in selected_categories]
            categories_tag = "-".join(cat_tokens[:5])
            if len(selected_categories) > 5:
                categories_tag = "{0}-plus{1}".format(categories_tag, len(selected_categories) - 5)
        else:
            categories_tag = "NoCategory"

        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        return "CategoryInstanceMappedPreview_{0}_{1}_{2}.csv".format(scope_tag, categories_tag, timestamp)

    def _prompt_save_mapped_preview_path(self, default_filename):
        dialog = SaveFileDialog()
        dialog.Title = "Save Mapped Preview CSV"
        dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        dialog.FileName = default_filename

        try:
            last_dir = getattr(config, "last_preview_export_dir", "")
        except Exception:
            last_dir = ""
        if last_dir and os.path.isdir(last_dir):
            dialog.InitialDirectory = last_dir

        result = dialog.ShowDialog()
        if result:
            return dialog.FileName
        return None

    def _export_mapped_preview_csv(self, export_path, rows, source_name):
        header = [
            "Row",
            "ElementId",
            "MatchedSourceElementId",
            "Source",
            "RawValue",
            "MappedValue",
            "Transform",
            "TargetParameter",
            "ExistingValueHandling",
            "Scope",
            "Categories",
            "ExportedAt",
        ]

        target_parameter = self.cmbParameter.SelectedItem if self.cmbParameter is not None else None
        target_text = "" if target_parameter is None else str(target_parameter)
        scope_text = self._selected_scope()
        categories_text = ", ".join(self._selected_category_names())
        export_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(export_path, "wb") as stream:
            writer = csv.writer(stream)
            writer.writerow([_csv_cell(col) for col in header])
            for row in rows:
                writer.writerow([
                    _csv_cell(row.get("row", "")),
                    _csv_cell(row.get("element_id", "")),
                    _csv_cell(row.get("source_element_id", "")),
                    _csv_cell(source_name),
                    _csv_cell(row.get("raw", "")),
                    _csv_cell(row.get("mapped", "")),
                    _csv_cell(self._selected_transform_mode()),
                    _csv_cell(target_text),
                    _csv_cell(self._selected_existing_value_policy()),
                    _csv_cell(scope_text),
                    _csv_cell(categories_text),
                    _csv_cell(export_time),
                ])

    def _maybe_export_mapped_preview(self, preview_rows, source_name):
        try:
            export_enabled = bool(self.chkExportMappedPreviewCsv.IsChecked)
        except Exception:
            export_enabled = False

        if not export_enabled:
            return

        if not preview_rows:
            forms.alert("No preview rows available to export.", title="Category Instance Parameters")
            return

        default_name = self._build_mapped_preview_filename()
        export_path = self._prompt_save_mapped_preview_path(default_name)
        if not export_path:
            return

        self._export_mapped_preview_csv(export_path, preview_rows, source_name)

        try:
            export_dir = os.path.dirname(export_path)
            if export_dir:
                config.last_preview_export_dir = export_dir
                script.save_config()
        except Exception:
            pass

        forms.alert("Mapped preview exported to:\n{0}".format(export_path), title="Category Instance Parameters")

    def _transform_source_value(self, raw_value, element=None):
        transform = self._selected_transform_mode()
        prefix = self.txtPrefix.Text or ""
        suffix = self.txtSuffix.Text or ""
        replace_from = self.txtReplaceFrom.Text if self.txtReplaceFrom is not None else ""
        replace_to = self.txtReplaceTo.Text if self.txtReplaceTo is not None else ""
        expression_text = self.txtExpression.Text if self.txtExpression is not None else ""
        context = self._build_transform_context(element, raw_value)

        if transform == "Keep Original":
            result = raw_value
        else:
            text = "" if raw_value is None else str(raw_value)
            if transform == "Convert To String":
                result = text
            elif transform == "Digits Only":
                result = re.sub(r"\D", "", text)
            elif transform == "Pad Number (Width)":
                digits = re.sub(r"\D", "", text)
                if digits:
                    result = digits.zfill(self._get_pad_width())
                else:
                    result = ""
            elif transform == "Digits + Pad Number":
                digits = re.sub(r"\D", "", text)
                result = digits.zfill(self._get_pad_width()) if digits else ""
            elif transform == "Level Number (2-digit)":
                digits = re.sub(r"\D", "", text)
                result = digits.zfill(2) if digits else ""
            elif transform == "Level Number (3-digit)":
                digits = re.sub(r"\D", "", text)
                result = digits.zfill(3) if digits else ""
            elif transform == "Uppercase":
                result = text.upper()
            elif transform == "Lowercase":
                result = text.lower()
            elif transform == "Title Case":
                result = text.title()
            elif transform == "Trim Spaces":
                result = text.strip()
            elif transform == "Remove Spaces":
                result = "".join(text.split())
            elif transform == "Left 3 Characters":
                result = text[:3]
            elif transform == "Dictionary (Translate)":
                result = _translate_dictionary_value(raw_value, expression_text)
            elif transform == "Template (Tokens)":
                result = self._apply_template_expression(expression_text, context)
            elif transform == "Formula (Safe Numeric)":
                base_value = _to_number(raw_value)
                if base_value is None:
                    raise ValueError("Formula mode requires a numeric source value.")
                level_num = _to_number(context.get("level_number", ""))
                formula_vars = {
                    "value": base_value,
                    "v": base_value,
                    "level": 0.0 if level_num is None else float(level_num),
                }
                result = _safe_numeric_eval(expression_text, formula_vars)
            else:
                result = text

        if replace_from:
            result = ("" if result is None else str(result)).replace(replace_from, replace_to)

        if prefix or suffix:
            result = "{0}{1}{2}".format(prefix, "" if result is None else str(result), suffix)

        return result

    def category_search_changed(self, sender, e):
        try:
            self._refresh_category_list(self.txtCategorySearch.Text)
        except Exception as ex:
            self._show_ui_error("Category search", ex)

    def select_all_categories_click(self, sender, e):
        try:
            for item in self.category_items:
                item["cb"].IsChecked = True
        except Exception as ex:
            self._show_ui_error("Select all categories", ex)

    def deselect_all_categories_click(self, sender, e):
        try:
            for item in self.category_items:
                item["cb"].IsChecked = False
        except Exception as ex:
            self._show_ui_error("Deselect all categories", ex)

    def collect_scope_changed(self, sender, e):
        try:
            if not getattr(self, "_ui_ready", False):
                return
            self._update_scope_dependent_ui()
            self._save_current_settings()
        except Exception as ex:
            self._show_ui_error("Collect scope changed", ex)

    def operation_mode_changed(self, sender, e):
        try:
            if not getattr(self, "_ui_ready", False):
                return
            self._update_mode_dependent_ui()
            self._save_current_settings()
        except Exception as ex:
            self._show_ui_error("Operation mode changed", ex)

    def mode3_source_options_changed(self, sender, e):
        try:
            if getattr(self, "_suppress_mode3_events", False):
                return
            if not getattr(self, "_ui_ready", False):
                return
            self._update_mode_dependent_ui()
            self._save_current_settings()
        except Exception as ex:
            self._show_ui_error("Mode 3 source options changed", ex)

    def source_parameter_changed(self, sender, e):
        try:
            if not getattr(self, "_ui_ready", False):
                return
            self._save_current_settings()
        except Exception as ex:
            self._show_ui_error("Source parameter changed", ex)

    def value_transform_changed(self, sender, e):
        try:
            if not getattr(self, "_ui_ready", False):
                return
            self._update_transform_hint()
            self._save_current_settings()
        except Exception as ex:
            self._show_ui_error("Value transform changed", ex)

    def persist_inputs_changed(self, sender, e):
        try:
            if not getattr(self, "_ui_ready", False):
                return
            self._save_current_settings()
        except Exception:
            pass

    def preview_source_values_click(self, sender, e):
        if not self._begin_busy("Preview Source Values"):
            return
        try:
            self.lstSourcePreview.Items.Clear()
            preview_rows = []

            if not self.selected_elements:
                self.lstSourcePreview.Items.Add("Collect elements first.")
                return

            source_meta = self._get_source_meta()
            if source_meta is None:
                self.lstSourcePreview.Items.Add("Select a source parameter first.")
                return

            if not self._ensure_mode3_ready():
                return

            param_name = self._get_source_display_name()

            limit = min(len(self.selected_elements), 25)
            for i in range(limit):
                element = self.selected_elements[i]
                element_id = _safe_element_id(element)
                if not _is_valid_element(element):
                    self.lstSourcePreview.Items.Add("{0}. {1} | Raw: <invalid element> -> New: <invalid element>".format(i + 1, param_name))
                    preview_rows.append({
                        "row": i + 1,
                        "element_id": element_id,
                        "raw": "<invalid element>",
                        "mapped": "<invalid element>",
                    })
                    continue

                mode3_note = None
                matched_source = None
                if self._is_mode3_source_mode():
                    raw, matched_source, mode3_note = self._read_mode3_source_value(element, source_meta)
                else:
                    raw = self._read_source_value(element, source_meta)

                try:
                    transform_element = matched_source if matched_source is not None else element
                    transformed = self._transform_source_value(raw, transform_element)
                except Exception as row_ex:
                    transformed = "<error: {0}>".format(row_ex)
                raw_shown = self._format_preview_value(raw)
                transformed_shown = self._format_preview_value(transformed)
                source_suffix = ""
                if matched_source is not None:
                    source_suffix = " | Source Element: {0}".format(_safe_element_id(matched_source))
                elif mode3_note:
                    source_suffix = " | Source: {0}".format(mode3_note)
                self.lstSourcePreview.Items.Add(
                    "{0}. {1} | Raw: {2} -> New: {3}{4}".format(
                        i + 1,
                        param_name,
                        raw_shown,
                        transformed_shown,
                        source_suffix,
                    )
                )
                preview_rows.append({
                    "row": i + 1,
                    "element_id": element_id,
                    "raw": raw_shown,
                    "mapped": transformed_shown,
                    "source_element_id": _safe_element_id(matched_source) if matched_source is not None else "",
                })

            self._maybe_export_mapped_preview(preview_rows, param_name)
        except Exception as ex:
            self._show_ui_error("Preview source values", ex)
        finally:
            self._end_busy()

    def collect_elements_click(self, sender, e):
        if not self._begin_busy("Collect Elements"):
            return
        try:
            category_ids = self._selected_category_ids()
            if not category_ids:
                forms.alert("Select at least one category first.")
                return

            scope_name = self._selected_scope()
            active_view_only = "Active" in scope_name
            selected_level_mode = "Selected Level" in scope_name
            selected_level_id = None

            if selected_level_mode:
                selected_level_id = self._selected_level_id()
                if selected_level_id < 0:
                    forms.alert("Select a level first.")
                    return

            if "Entire" in scope_name:
                proceed = self._confirm(
                    "Collecting from Entire Model can take longer on large projects. Continue?",
                    title="Large Collection Warning"
                )
                if not proceed:
                    return

            estimate_seconds = _estimate_seconds_for_collect(
                len(category_ids),
                active_view_only,
                selected_level_mode,
                "Entire" in scope_name,
            )
            proceed = _confirm_with_estimate(
                self,
                "Confirm Collection",
                "This will scan the selected categories before building the parameter list.",
                estimate_seconds,
            )
            if not proceed:
                return

            elements = list(
                _iter_category_elements(
                    category_ids,
                    active_view_only=active_view_only,
                    hard_limit=MAX_COLLECT_LIMIT,
                    level_id=selected_level_id,
                )
            )
            if len(elements) >= MAX_COLLECT_LIMIT:
                forms.alert(
                    "Element collection reached safety limit ({0:,}). Narrow categories or use Active View scope.".format(MAX_COLLECT_LIMIT)
                )

            self.selected_elements = elements
            self._refresh_element_summary()
            self._refresh_parameter_controls()
            self._save_current_settings()
        except Exception as ex:
            self._show_ui_error("Collect elements", ex)
        finally:
            self._end_busy()

    def use_current_selection_click(self, sender, e):
        if not self._begin_busy("Use Current Selection"):
            return
        try:
            category_ids = set(self._selected_category_ids())
            if not category_ids:
                forms.alert("Select at least one category first.")
                return

            selected_ids = uidoc.Selection.GetElementIds()
            if not selected_ids:
                forms.alert("No elements are selected in Revit.")
                return

            elements = []
            for eid in selected_ids:
                element = doc.GetElement(eid)
                if not _is_valid_element(element) or element.Category is None:
                    continue
                if element.Category.Id.IntegerValue not in category_ids:
                    continue
                elements.append(element)

            self.selected_elements = elements
            self._refresh_element_summary()
            self._refresh_parameter_controls()
            self._save_current_settings()
        except Exception as ex:
            self._show_ui_error("Use current selection", ex)
        finally:
            self._end_busy()

    def parameter_selection_changed(self, sender, e):
        try:
            parameter_name = self.cmbParameter.SelectedItem
            if parameter_name is None:
                self.txtSelectedParameterInfo.Text = "No parameter selected."
                return

            storage_type = self.common_param_map.get(parameter_name)
            if storage_type is None:
                self.txtSelectedParameterInfo.Text = "No parameter selected."
                return

            self.txtSelectedParameterInfo.Text = (
                "Selected parameter: {0}\n"
                "Storage type: {1}\n"
                "Note: this tool writes instance parameters only."
            ).format(parameter_name, _storage_label(storage_type))
        except Exception as ex:
            self._show_ui_error("Target parameter changed", ex)

    def clear_value_changed(self, sender, e):
        try:
            clear_mode = bool(self.chkClearValue.IsChecked)
            self.txtValue.IsEnabled = not clear_mode
        except Exception as ex:
            self._show_ui_error("Clear value toggle", ex)

    def _apply_direct_update(self):
        parameter_name = self.cmbParameter.SelectedItem
        if parameter_name is None:
            forms.alert("Choose a target parameter first.")
            return

        storage_type = self.common_param_map.get(parameter_name)
        if storage_type is None:
            forms.alert("Selected parameter is not available.")
            return

        clear_mode = bool(self.chkClearValue.IsChecked)
        value_text = self.txtValue.Text if self.txtValue is not None else ""
        duplicate_mode = self._selected_duplicate_mode()
        existing_value_policy = self._selected_existing_value_policy()

        effective_duplicate_mode = duplicate_mode
        if existing_value_policy == "Skip Elements With Existing Value":
            effective_duplicate_mode = "Skip Existing"
        elif duplicate_mode == "Skip Existing":
            effective_duplicate_mode = "Overwrite"

        coerced_value = None
        if not clear_mode:
            try:
                coerced_value = _coerce_value(value_text, storage_type)
            except Exception as ex:
                forms.alert("Invalid value for {0}: {1}".format(_storage_label(storage_type), ex))
                return

        element_count = len(self.selected_elements)
        if element_count > MAX_APPLY_LIMIT:
            forms.alert(
                "Safety limit exceeded: {0:,} elements selected. Limit is {1:,}. Narrow your selection and retry.".format(
                    element_count,
                    MAX_APPLY_LIMIT,
                ),
                title="Category Instance Parameters"
            )
            return

        if not self._confirm_target_coverage(parameter_name, storage_type):
            return

        estimate_seconds = _estimate_seconds_for_apply(element_count, False)
        confirm_message = (
            "This will write parameter values to the selected elements.\n"
            "Existing value handling: {0}".format(existing_value_policy)
        )
        proceed = _confirm_with_estimate(
            self,
            "Confirm Update",
            confirm_message,
            estimate_seconds,
        )
        if not proceed:
            return

        written = 0
        skipped = 0
        failed = 0
        skipped_invalid_element = 0
        skipped_missing_target_param = 0
        skipped_existing_value = 0
        fail_samples = []
        skip_samples = []
        write_samples = []

        chunk_size = self._get_apply_chunk_size(source_mode=False)
        processed = 0
        cancelled_by_user = False

        tg = TransactionGroup(doc, "Batch Update Category Instance Parameters")
        tg.Start()
        try:
            with forms.ProgressBar(title="Applying updates {value}/{max_value}", cancellable=True) as pb_apply:
                i = 0
                while i < element_count:
                    if pb_apply.cancelled:
                        cancelled_by_user = True
                        break

                    chunk = self.selected_elements[i:i + chunk_size]
                    tx = Transaction(doc, "Batch Update Category Instance Parameters")
                    tx.Start()
                    try:
                        for element in chunk:
                            if pb_apply.cancelled:
                                cancelled_by_user = True
                                break

                            if not _is_valid_element(element):
                                skipped += 1
                                skipped_invalid_element += 1
                                processed += 1
                                if (processed % 10) == 0 or processed == element_count:
                                    pb_apply.update_progress(processed, element_count)
                                continue

                            param = _find_writable_instance_parameter(element, parameter_name, storage_type)
                            if param is None:
                                skipped += 1
                                skipped_missing_target_param += 1
                                processed += 1
                                if (processed % 10) == 0 or processed == element_count:
                                    pb_apply.update_progress(processed, element_count)
                                continue

                            if (not clear_mode) and existing_value_policy == "Skip Elements With Existing Value":
                                if _has_effective_parameter_value(param, storage_type):
                                    skipped += 1
                                    skipped_existing_value += 1
                                    if len(skip_samples) < 10:
                                        skip_samples.append("Element {0}: target already has a value".format(_safe_element_id(element)))
                                    processed += 1
                                    if (processed % 10) == 0 or processed == element_count:
                                        pb_apply.update_progress(processed, element_count)
                                    continue

                            try:
                                if clear_mode:
                                    ok = _clear_parameter(param, storage_type)
                                    if ok:
                                        verified_value = _coerce_readback_value(param, storage_type)
                                        if _values_match(verified_value, ElementId.InvalidElementId.IntegerValue if storage_type == StorageType.ElementId else ("" if storage_type == StorageType.String else 0 if storage_type == StorageType.Integer else 0.0), storage_type):
                                            written += 1
                                            if len(write_samples) < 10:
                                                try:
                                                    write_samples.append("Element {0}: {1} cleared -> {2}".format(_safe_element_id(element), parameter_name, _display_parameter_value(param)))
                                                except Exception:
                                                    write_samples.append("Element {0}: {1} cleared".format(_safe_element_id(element), parameter_name))
                                        else:
                                            failed += 1
                                            if len(fail_samples) < 10:
                                                fail_samples.append("Element {0}: clear verification failed (read back {1})".format(_safe_element_id(element), _display_parameter_value(param)))
                                    else:
                                        failed += 1
                                        if len(fail_samples) < 10:
                                            fail_samples.append("Element {0}: clear failed".format(_safe_element_id(element)))
                                else:
                                    expected_value = _plan_expected_value(param, coerced_value, storage_type, effective_duplicate_mode)
                                    ok, message = _set_parameter_value(param, coerced_value, storage_type, effective_duplicate_mode)
                                    if ok:
                                        verified_value = _coerce_readback_value(param, storage_type)
                                        if _values_match(verified_value, expected_value, storage_type):
                                            written += 1
                                            if len(write_samples) < 10:
                                                try:
                                                    write_samples.append("Element {0}: {1} -> {2}".format(_safe_element_id(element), parameter_name, _display_parameter_value(param)))
                                                except Exception:
                                                    write_samples.append("Element {0}: {1} written".format(_safe_element_id(element), parameter_name))
                                        else:
                                            failed += 1
                                            if len(fail_samples) < 10:
                                                fail_samples.append(
                                                    "Element {0}: verification failed, expected {1}, read {2}".format(
                                                        _safe_element_id(element),
                                                        expected_value,
                                                        _display_parameter_value(param),
                                                    )
                                                )
                                    elif message == "skipped existing":
                                        skipped += 1
                                        skipped_existing_value += 1
                                    else:
                                        failed += 1
                                        if len(fail_samples) < 10:
                                            fail_samples.append("Element {0}: {1}".format(_safe_element_id(element), message))
                            except Exception as ex:
                                failed += 1
                                if len(fail_samples) < 10:
                                    fail_samples.append("Element {0}: {1}".format(_safe_element_id(element), ex))

                            processed += 1
                            if (processed % 10) == 0 or processed == element_count:
                                pb_apply.update_progress(processed, element_count)

                        if cancelled_by_user:
                            tx.RollBack()
                            break

                        tx.Commit()
                    except Exception:
                        tx.RollBack()
                        raise

                    i += chunk_size

            if cancelled_by_user:
                tg.RollBack()
                forms.alert("Update cancelled by user. No changes were committed.", title="Category Instance Parameters")
                return

            tg.Assimilate()
        except Exception as ex:
            try:
                tg.RollBack()
            except Exception:
                pass
            forms.alert("Update failed and was rolled back.\n\n{0}".format(ex))
            return

        summary_lines = [
            "Batch update complete.",
            "Parameter: {0}".format(parameter_name),
            "Estimated duration: {0}".format(_format_estimate(estimate_seconds)),
            "Existing value handling: {0}".format(existing_value_policy),
            "Write mode: {0}".format(effective_duplicate_mode),
            "Elements processed: {0}/{1}".format(processed, element_count),
            "Transaction chunk size: {0}".format(chunk_size),
            "Written: {0}".format(written),
            "Skipped: {0}".format(skipped),
            "Failed: {0}".format(failed),
        ]

        if skipped_invalid_element:
            summary_lines.append("Skipped invalid elements: {0}".format(skipped_invalid_element))
        if skipped_missing_target_param:
            summary_lines.append("Skipped missing/non-writable target parameter: {0}".format(skipped_missing_target_param))
        if skipped_existing_value:
            summary_lines.append("Skipped due to existing value policy: {0}".format(skipped_existing_value))

        if write_samples:
            summary_lines.append("")
            summary_lines.append("Sample written values:")
            summary_lines.extend(["- {0}".format(item) for item in write_samples])

        if fail_samples:
            summary_lines.append("")
            summary_lines.append("Sample failures:")
            summary_lines.extend(["- {0}".format(item) for item in fail_samples])

        if skip_samples:
            summary_lines.append("")
            summary_lines.append("Sample skipped items:")
            summary_lines.extend(["- {0}".format(item) for item in skip_samples])

        forms.alert("\n".join(summary_lines), title="Category Instance Parameters")

    def _apply_source_process_update(self):
        target_param_name = self.cmbParameter.SelectedItem
        if target_param_name is None:
            forms.alert("Choose a target parameter first.")
            return

        target_storage = self.common_param_map.get(target_param_name)
        if target_storage is None:
            forms.alert("Selected target parameter is not available.")
            return

        source_meta = self._get_source_meta()
        if source_meta is None:
            forms.alert("Choose a source parameter first.")
            return

        if not self._ensure_mode3_ready():
            return

        existing_value_policy = self._selected_existing_value_policy()
        effective_duplicate_mode = "Overwrite"
        if existing_value_policy == "Skip Elements With Existing Value":
            effective_duplicate_mode = "Skip Existing"
        element_count = len(self.selected_elements)

        if element_count > MAX_APPLY_LIMIT:
            forms.alert(
                "Safety limit exceeded: {0:,} elements selected. Limit is {1:,}. Narrow your selection and retry.".format(
                    element_count,
                    MAX_APPLY_LIMIT,
                ),
                title="Category Instance Parameters"
            )
            return

        if not self._confirm_target_coverage(target_param_name, target_storage):
            return

        estimate_seconds = _estimate_seconds_for_apply(element_count, True)
        mode3_source_location = self._selected_mode3_source_location() if self._is_mode3_source_mode() else "Same Element"
        confirm_message = "This will read source values, transform them, and write to the target parameter.\nSource mode: {0}\nExisting value handling: {1}".format(mode3_source_location, existing_value_policy)
        proceed = _confirm_with_estimate(
            self,
            "Confirm Update",
            confirm_message,
            estimate_seconds,
        )
        if not proceed:
            return

        written = 0
        skipped = 0
        failed = 0
        skipped_invalid_element = 0
        skipped_missing_target_param = 0
        skipped_existing_value = 0
        skipped_missing_source = 0
        skipped_empty_processed = 0
        fail_samples = []
        skip_samples = []
        write_samples = []

        chunk_size = self._get_apply_chunk_size(source_mode=True)
        processed = 0
        cancelled_by_user = False

        tg = TransactionGroup(doc, "Process and Copy Category Instance Parameters")
        tg.Start()
        try:
            with forms.ProgressBar(title="Processing and applying values {value}/{max_value}", cancellable=True) as pb_apply:
                i = 0
                while i < element_count:
                    if pb_apply.cancelled:
                        cancelled_by_user = True
                        break

                    chunk = self.selected_elements[i:i + chunk_size]
                    tx = Transaction(doc, "Process and Copy Category Instance Parameters")
                    tx.Start()
                    try:
                        for element in chunk:
                            if pb_apply.cancelled:
                                cancelled_by_user = True
                                break

                            if not _is_valid_element(element):
                                skipped += 1
                                skipped_invalid_element += 1
                                processed += 1
                                if (processed % 10) == 0 or processed == element_count:
                                    pb_apply.update_progress(processed, element_count)
                                continue

                            target_param = _find_writable_instance_parameter(element, target_param_name, target_storage)
                            if target_param is None:
                                skipped += 1
                                skipped_missing_target_param += 1
                                processed += 1
                                if (processed % 10) == 0 or processed == element_count:
                                    pb_apply.update_progress(processed, element_count)
                                continue

                            if existing_value_policy == "Skip Elements With Existing Value":
                                if _has_effective_parameter_value(target_param, target_storage):
                                    skipped += 1
                                    skipped_existing_value += 1
                                    if len(skip_samples) < 10:
                                        skip_samples.append("Element {0}: target already has a value".format(_safe_element_id(element)))
                                    processed += 1
                                    if (processed % 10) == 0 or processed == element_count:
                                        pb_apply.update_progress(processed, element_count)
                                    continue

                            try:
                                matched_source = None
                                mode3_note = None
                                if self._is_mode3_source_mode():
                                    raw_value, matched_source, mode3_note = self._read_mode3_source_value(element, source_meta)
                                    if mode3_note and raw_value is None:
                                        skipped += 1
                                        skipped_missing_source += 1
                                        if len(skip_samples) < 10:
                                            skip_samples.append("Element {0}: {1}".format(_safe_element_id(element), mode3_note))
                                        processed += 1
                                        if (processed % 10) == 0 or processed == element_count:
                                            pb_apply.update_progress(processed, element_count)
                                        continue
                                else:
                                    raw_value = self._read_source_value(element, source_meta)

                                transform_element = matched_source if matched_source is not None else element
                                processed_value = self._transform_source_value(raw_value, transform_element)

                                if target_storage == StorageType.String:
                                    coerced = "" if processed_value is None else str(processed_value)
                                else:
                                    if processed_value in (None, ""):
                                        skipped += 1
                                        skipped_empty_processed += 1
                                        processed += 1
                                        if (processed % 10) == 0 or processed == element_count:
                                            pb_apply.update_progress(processed, element_count)
                                        continue
                                    coerced = _coerce_value(processed_value, target_storage)

                                expected_value = _plan_expected_value(target_param, coerced, target_storage, effective_duplicate_mode)
                                ok, message = _set_parameter_value(target_param, coerced, target_storage, effective_duplicate_mode)
                                if ok:
                                    verified_value = _coerce_readback_value(target_param, target_storage)
                                    if _values_match(verified_value, expected_value, target_storage):
                                        written += 1
                                        if len(write_samples) < 10:
                                            try:
                                                source_name = self._get_source_display_name()
                                                write_samples.append(
                                                    "Element {0}: {1} -> {2} = {3}{4}".format(
                                                        _safe_element_id(element),
                                                        source_name,
                                                        target_param_name,
                                                        _display_parameter_value(target_param),
                                                        " (Source {0})".format(_safe_element_id(matched_source)) if matched_source is not None else "",
                                                    )
                                                )
                                            except Exception:
                                                write_samples.append("Element {0}: {1} written".format(_safe_element_id(element), target_param_name))
                                    else:
                                        failed += 1
                                        if len(fail_samples) < 10:
                                            fail_samples.append(
                                                "Element {0}: verification failed, expected {1}, read {2}".format(
                                                    _safe_element_id(element),
                                                    expected_value,
                                                    _display_parameter_value(target_param),
                                                )
                                            )
                                elif message == "skipped existing":
                                    skipped += 1
                                    skipped_existing_value += 1
                                else:
                                    failed += 1
                                    if len(fail_samples) < 10:
                                        fail_samples.append("Element {0}: {1}".format(_safe_element_id(element), message))
                            except Exception as ex:
                                failed += 1
                                if len(fail_samples) < 10:
                                    fail_samples.append("Element {0}: {1}".format(_safe_element_id(element), ex))

                            processed += 1
                            if (processed % 10) == 0 or processed == element_count:
                                pb_apply.update_progress(processed, element_count)

                        if cancelled_by_user:
                            tx.RollBack()
                            break

                        tx.Commit()
                    except Exception:
                        tx.RollBack()
                        raise

                    i += chunk_size

            if cancelled_by_user:
                tg.RollBack()
                forms.alert("Update cancelled by user. No changes were committed.", title="Category Instance Parameters")
                return

            tg.Assimilate()
        except Exception as ex:
            try:
                tg.RollBack()
            except Exception:
                pass
            forms.alert("Update failed and was rolled back.\n\n{0}".format(ex))
            return

        summary_lines = [
            "Source processing update complete.",
            "Operation mode: {0}".format(self._selected_operation_mode()),
            "Source mode: {0}".format(mode3_source_location),
            "Target parameter: {0}".format(target_param_name),
            "Transform: {0}".format(self._selected_transform_mode()),
            "Estimated duration: {0}".format(_format_estimate(estimate_seconds)),
            "Existing value handling: {0}".format(existing_value_policy),
            "Write mode: {0}".format(effective_duplicate_mode),
            "Elements processed: {0}/{1}".format(processed, element_count),
            "Transaction chunk size: {0}".format(chunk_size),
            "Written: {0}".format(written),
            "Skipped: {0}".format(skipped),
            "Failed: {0}".format(failed),
        ]

        if skipped_invalid_element:
            summary_lines.append("Skipped invalid elements: {0}".format(skipped_invalid_element))
        if skipped_missing_target_param:
            summary_lines.append("Skipped missing/non-writable target parameter: {0}".format(skipped_missing_target_param))
        if skipped_existing_value:
            summary_lines.append("Skipped due to existing value policy: {0}".format(skipped_existing_value))
        if skipped_missing_source:
            summary_lines.append("Skipped due to missing source match/value: {0}".format(skipped_missing_source))
        if skipped_empty_processed:
            summary_lines.append("Skipped due to empty processed value for numeric target: {0}".format(skipped_empty_processed))

        if write_samples:
            summary_lines.append("")
            summary_lines.append("Sample written values:")
            summary_lines.extend(["- {0}".format(item) for item in write_samples])

        if fail_samples:
            summary_lines.append("")
            summary_lines.append("Sample failures:")
            summary_lines.extend(["- {0}".format(item) for item in fail_samples])

        if skip_samples:
            summary_lines.append("")
            summary_lines.append("Sample skipped items:")
            summary_lines.extend(["- {0}".format(item) for item in skip_samples])

        forms.alert("\n".join(summary_lines), title="Category Instance Parameters")

    def apply_update_click(self, sender, e):
        if not self._begin_busy("Apply Update"):
            return
        try:
            if not self.selected_elements:
                forms.alert("Collect or select elements before applying updates.")
                return

            if self._is_source_mode():
                self._apply_source_process_update()
            else:
                self._apply_direct_update()

            self._save_current_settings()
        except Exception as ex:
            self._show_ui_error("Apply update", ex)
        finally:
            self._end_busy()

    def close_click(self, sender, e):
        if self._is_busy:
            forms.alert(
                "Please wait for the current operation to finish before closing.",
                title="Category Instance Parameters"
            )
            return
        self._save_current_settings()
        self.Close()


try:
    window = CategoryInstanceParameterWindow("WPFWindow.xaml")
    window.ShowDialog()
except Exception as ex:
    forms.alert("Unable to launch Category Instance Parameters tool.\n\n{0}".format(ex))
