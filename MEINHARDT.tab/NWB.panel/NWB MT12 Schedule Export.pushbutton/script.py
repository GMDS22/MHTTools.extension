# coding: utf8
from __future__ import print_function

import codecs
import os
import time

import clr

from System import Activator, Type
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal

from Autodesk.Revit.DB import FilteredElementCollector
from pyrevit import forms, revit


doc = revit.doc

TOOL_TITLE = "NWB Maintainability Schedule Export"
AUTOFILL_FOLDER = "NWB_PARAMETERS AutoFill.pushbutton"
XLSX_FILE_FORMAT = 51
DEFAULT_MAINTAINABILITY_TYPES = ["1", "2"]
COMMON_MAINTAINABILITY_TYPES = ["1", "2", "3", "Not required"]


def _safe_str(value):
    try:
        if value is None:
            return ""
        return str(value)
    except Exception:
        return ""


def _tool_folder():
    return os.path.dirname(__file__)


def _logs_folder():
    path = os.path.join(_tool_folder(), "logs")
    if not os.path.isdir(path):
        os.makedirs(path)
    return path


def _autofill_script_path():
    return os.path.normpath(os.path.join(_tool_folder(), "..", AUTOFILL_FOLDER, "script.py"))


def _load_mapping_module():
    path = _autofill_script_path()
    if not os.path.isfile(path):
        raise RuntimeError("NWB_PARAMETERS AutoFill script was not found: {0}".format(path))

    namespace = {
        "__name__": "nwb_mt12_schedule_export",
        "__file__": path,
    }

    with open(path, "r") as source_file:
        source = source_file.read()
    exec(compile(source, path, "exec"), namespace)
    return namespace


def _timestamp():
    return time.strftime("%Y%m%d_%H%M%S")


def _sanitize_filename_part(value, max_len=80):
    text = _safe_str(value).strip()
    if not text:
        return "NA"
    cleaned = []
    for char in text:
        if char.isalnum() or char in "._-":
            cleaned.append(char)
        else:
            cleaned.append("-")
    text = "".join(cleaned)
    while "--" in text:
        text = text.replace("--", "-")
    text = text.strip("-._")
    if not text:
        text = "NA"
    if len(text) > max_len:
        text = text[:max_len].rstrip("-._")
    return text or "NA"


def _ordered_unique_text(values):
    ordered = []
    seen = set()
    for value in values:
        text = _safe_str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        ordered.append(text)
    return ordered


def _maintainability_sort_key(value):
    text = _safe_str(value).strip()
    try:
        return (0, int(text))
    except Exception:
        pass

    if text.lower() == "not required":
        return (2, text.lower())
    return (1, text.lower())


def _available_maintainability_types(workbook_data):
    values = []
    for asset_row in getattr(workbook_data, "asset_rows", []) or []:
        if _safe_str(asset_row.get("Hierarchy - L1", "")).strip() != "Building Services":
            continue
        values.append(asset_row.get("NWB_MaintainabilityType", ""))

    values.extend(COMMON_MAINTAINABILITY_TYPES)
    return sorted(_ordered_unique_text(values), key=_maintainability_sort_key)


def _select_maintainability_types(workbook_data):
    options = _available_maintainability_types(workbook_data)
    if not options:
        options = COMMON_MAINTAINABILITY_TYPES[:]

    selected = forms.SelectFromList.show(
        options,
        multiselect=True,
        title=TOOL_TITLE,
        button_name="Export Schedule",
        width=520,
        height=480,
    )
    if not selected:
        return []
    if isinstance(selected, str):
        selected = [selected]
    return sorted(_ordered_unique_text(selected), key=_maintainability_sort_key)


def _sheet_name_for_types(selected_types):
    if not selected_types:
        return "NWB Schedule"

    compact = []
    for value in selected_types:
        text = _safe_str(value).strip()
        if text.lower() == "not required":
            compact.append("NR")
        else:
            compact.append(text)

    name = "NWB MT {0}".format("-".join(compact))
    if len(name) > 31:
        name = name[:31].rstrip()
    return name or "NWB Schedule"


def _types_slug(selected_types):
    if not selected_types:
        return "none"
    return _sanitize_filename_part("-".join(selected_types), max_len=40)


def _csv_escape(value):
    text = _safe_str(value).replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    return '"{0}"'.format(text.replace('"', '""'))


def _write_csv(path, headers, rows):
    stream = codecs.open(path, "w", "utf-8-sig")
    try:
        stream.write(",".join([_csv_escape(header) for header in headers]))
        stream.write("\n")
        for row in rows:
            stream.write(",".join([_csv_escape(row.get(header, "")) for header in headers]))
            stream.write("\n")
    finally:
        stream.close()


def _com_release(obj):
    try:
        if obj is not None:
            Marshal.ReleaseComObject(obj)
    except Exception:
        pass


def _com_set(obj, member_name, value):
    try:
        setattr(obj, member_name, value)
        return
    except Exception:
        pass

    try:
        obj.GetType().InvokeMember(
            member_name,
            BindingFlags.SetProperty,
            None,
            obj,
            (value,),
        )
    except Exception:
        pass


def _com_get(obj, member_name):
    try:
        return getattr(obj, member_name)
    except Exception:
        return obj.GetType().InvokeMember(
            member_name,
            BindingFlags.GetProperty,
            None,
            obj,
            None,
        )


def _com_call(obj, member_name, *args):
    try:
        member = getattr(obj, member_name)
        return member(*args)
    except Exception:
        return obj.GetType().InvokeMember(
            member_name,
            BindingFlags.InvokeMethod,
            None,
            obj,
            args,
        )


def _com_item(collection_obj, index):
    try:
        return collection_obj[index]
    except Exception:
        pass

    try:
        return collection_obj.Item[index]
    except Exception:
        pass

    try:
        return collection_obj.Item(index)
    except Exception:
        pass

    return collection_obj.GetType().InvokeMember(
        "Item",
        BindingFlags.GetProperty,
        None,
        collection_obj,
        (index,),
    )


def _com_item2(collection_obj, index1, index2):
    try:
        return collection_obj[index1, index2]
    except Exception:
        pass

    try:
        return collection_obj.Item[index1, index2]
    except Exception:
        pass

    try:
        return collection_obj.Item(index1, index2)
    except Exception:
        pass

    return collection_obj.GetType().InvokeMember(
        "Item",
        BindingFlags.GetProperty,
        None,
        collection_obj,
        (index1, index2),
    )


def _write_xlsx(path, headers, rows, sheet_name):
    excel_type = Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise RuntimeError("Microsoft Excel is not available on this machine.")

    excel = None
    workbooks = None
    workbook = None
    worksheets = None
    worksheet = None
    cells = None
    used_range = None
    used_columns = None
    active_window = None

    try:
        excel = Activator.CreateInstance(excel_type)
        _com_set(excel, "Visible", False)
        _com_set(excel, "DisplayAlerts", False)

        workbooks = _com_get(excel, "Workbooks")
        workbook = _com_call(workbooks, "Add")
        worksheets = _com_get(workbook, "Worksheets")
        worksheet = _com_item(worksheets, 1)
        _com_set(worksheet, "Name", _safe_str(sheet_name) or "NWB Schedule")
        cells = _com_get(worksheet, "Cells")

        for col_index, header in enumerate(headers, 1):
            cell = None
            font = None
            interior = None
            try:
                cell = _com_item2(cells, 1, col_index)
                _com_set(cell, "Value2", _safe_str(header))
                font = _com_get(cell, "Font")
                interior = _com_get(cell, "Interior")
                _com_set(font, "Bold", True)
                _com_set(font, "Color", 0xFFFFFF)
                _com_set(interior, "Color", 0x9C630E)
            finally:
                _com_release(interior)
                _com_release(font)
                _com_release(cell)

        for row_index, row in enumerate(rows, 2):
            for col_index, header in enumerate(headers, 1):
                cell = None
                try:
                    cell = _com_item2(cells, row_index, col_index)
                    _com_set(cell, "Value2", _safe_str(row.get(header, "")))
                finally:
                    _com_release(cell)

        used_range = _com_get(worksheet, "UsedRange")
        try:
            _com_call(used_range, "AutoFilter")
        except Exception:
            pass

        try:
            used_columns = _com_get(used_range, "Columns")
            _com_call(used_columns, "AutoFit")
        except Exception:
            pass

        try:
            active_window = _com_get(excel, "ActiveWindow")
            _com_set(active_window, "SplitRow", 1)
            _com_set(active_window, "FreezePanes", True)
        except Exception:
            pass

        _com_call(workbook, "SaveAs", path, XLSX_FILE_FORMAT)
    finally:
        try:
            if workbook is not None:
                _com_call(workbook, "Close", False)
        except Exception:
            pass

        try:
            if excel is not None:
                _com_call(excel, "Quit")
        except Exception:
            pass

        _com_release(active_window)
        _com_release(used_columns)
        _com_release(used_range)
        _com_release(cells)
        _com_release(worksheet)
        _com_release(worksheets)
        _com_release(workbook)
        _com_release(workbooks)
        _com_release(excel)


def _collect_all_mep_elements(mapping):
    elements = []
    seen = set()

    for bic in mapping.get("MEP_CATEGORIES", []):
        try:
            collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        except Exception:
            continue

        try:
            for element in collector:
                if element is None:
                    continue
                try:
                    element_id_value = int(element.Id.IntegerValue)
                except Exception:
                    continue
                if element_id_value in seen:
                    continue
                seen.add(element_id_value)
                elements.append(element)
        except Exception:
            continue

    return elements


def _get_type_element(element):
    if element is None:
        return None
    try:
        type_id = element.GetTypeId()
    except Exception:
        type_id = None
    if type_id is None:
        return None
    try:
        if type_id.IntegerValue <= 0:
            return None
    except Exception:
        pass
    try:
        return doc.GetElement(type_id)
    except Exception:
        return None


def _get_element_name(element):
    if element is None:
        return ""
    try:
        return _safe_str(element.Name).strip()
    except Exception:
        pass
    return ""


def _get_family_name(element):
    try:
        symbol = getattr(element, "Symbol", None)
        if symbol is not None:
            family = getattr(symbol, "Family", None)
            if family is not None:
                return _safe_str(getattr(family, "Name", "")).strip()
    except Exception:
        pass

    type_element = _get_type_element(element)
    if type_element is not None:
        try:
            return _safe_str(getattr(type_element, "FamilyName", "")).strip()
        except Exception:
            pass

    return ""


def _get_type_name(element):
    type_element = _get_type_element(element)
    if type_element is not None:
        try:
            return _safe_str(getattr(type_element, "Name", "")).strip()
        except Exception:
            pass
    return ""


def _get_level_name(element):
    try:
        level_id = getattr(element, "LevelId", None)
        if level_id is not None and level_id.IntegerValue > 0:
            level = doc.GetElement(level_id)
            if level is not None:
                return _safe_str(getattr(level, "Name", "")).strip()
    except Exception:
        pass
    return ""


def _read_param_text_by_name(owner, names):
    if owner is None:
        return ""
    for name in names:
        try:
            param = owner.LookupParameter(name)
        except Exception:
            param = None
        if param is None:
            try:
                candidates = owner.GetParameters(name)
                if candidates:
                    param = candidates[0]
            except Exception:
                param = None
        if param is None:
            continue
        try:
            value = param.AsValueString() or param.AsString() or ""
        except Exception:
            try:
                value = param.AsString() or ""
            except Exception:
                value = ""
        value = _safe_str(value).strip()
        if value:
            return value
    return ""


def _get_system_name(element):
    try:
        mep_system = getattr(element, "MEPSystem", None)
        if mep_system is not None:
            text = _safe_str(getattr(mep_system, "Name", "")).strip()
            if text:
                return text
    except Exception:
        pass

    for owner in [element, _get_type_element(element)]:
        text = _read_param_text_by_name(
            owner,
            [
                "System Type",
                "System Name",
                "RBS_SYSTEM_NAME_PARAM",
                "RBS_DUCT_SYSTEM_TYPE_PARAM",
                "RBS_PIPING_SYSTEM_TYPE_PARAM",
            ],
        )
        if text:
            return text

    return ""


def _build_match_facts(mapping, element):
    category = _safe_str(mapping["_get_category_name"](element)).strip()
    family = _get_family_name(element)
    type_name = _get_type_name(element)
    element_name = _get_element_name(element)
    system_name = _get_system_name(element)

    facts = {
        "category": category,
        "family": family,
        "type": type_name,
        "name": element_name,
        "system": system_name,
        "category_norm": mapping["_normalize_text"](category),
        "family_norm": mapping["_normalize_text"](family),
        "type_norm": mapping["_normalize_text"](type_name),
        "name_norm": mapping["_normalize_text"](element_name),
        "system_norm": mapping["_normalize_text"](system_name),
    }
    facts["search_blob"] = " ".join(
        [
            facts["category_norm"],
            facts["family_norm"],
            facts["type_norm"],
            facts["name_norm"],
            facts["system_norm"],
        ]
    ).strip()
    facts["discipline_hint_norm"] = mapping["_normalize_text"](mapping["_derive_discipline_hint"](facts))
    facts["asset_category_candidates"] = mapping["_asset_category_candidates"](facts)
    return facts


def _facts_cache_key(facts):
    return (
        facts.get("category_norm", ""),
        facts.get("family_norm", ""),
        facts.get("type_norm", ""),
        facts.get("name_norm", ""),
        facts.get("system_norm", ""),
    )


def _selected_asset_row(mapping, workbook_data, facts, match_cache):
    key = _facts_cache_key(facts)
    if key not in match_cache:
        match_cache[key] = mapping["_select_asset_row"](workbook_data.asset_rows, facts)
    return match_cache[key]


def _is_allowed_maintainability_row(asset_row, selected_types):
    if not asset_row:
        return False
    l1 = _safe_str(asset_row.get("Hierarchy - L1", "")).strip()
    maintainability = _safe_str(asset_row.get("NWB_MaintainabilityType", "")).strip()
    return l1 == "Building Services" and maintainability in selected_types


def _build_shared_parameter_map(mapping, element, target_parameters):
    selected = {}
    values = {}
    target_set = set(target_parameters)

    for param_name in target_parameters:
        values[param_name] = ""

    for scope_name, owner in mapping["_iter_param_owners"](element):
        try:
            parameters = owner.Parameters
        except Exception:
            parameters = []

        for param in parameters:
            if param is None:
                continue
            try:
                definition = param.Definition
            except Exception:
                definition = None
            if definition is None:
                continue

            name = _safe_str(getattr(definition, "Name", "")).strip()
            if name not in target_set:
                continue
            try:
                is_shared = bool(mapping["_param_is_shared"](param))
            except Exception:
                is_shared = False
            if not is_shared:
                continue
            if name in selected:
                continue

            selected[name] = scope_name
            values[name] = _safe_str(mapping["_parameter_text"](param)).strip()

    return values


def _build_export_headers(mapping):
    headers = [
        "ElementId",
        "UniqueId",
        "RevitCategory",
        "Family",
        "Type",
        "ElementName",
        "Level",
        "System",
        "NWB_MaintainabilityType",
    ]
    for target in mapping.get("TARGET_PARAMETERS", []):
        headers.append(target)
    return headers


def _build_export_row(mapping, element, asset_row, shared_values):
    row = {
        "ElementId": _safe_str(getattr(getattr(element, "Id", None), "IntegerValue", "")).strip(),
        "UniqueId": _safe_str(getattr(element, "UniqueId", "")).strip(),
        "RevitCategory": _safe_str(mapping["_get_category_name"](element)).strip(),
        "Family": _get_family_name(element),
        "Type": _get_type_name(element),
        "ElementName": _get_element_name(element),
        "Level": _get_level_name(element),
        "System": _get_system_name(element),
        "NWB_MaintainabilityType": _safe_str(asset_row.get("NWB_MaintainabilityType", "")).strip(),
    }

    for target in mapping.get("TARGET_PARAMETERS", []):
        row[target] = _safe_str(shared_values.get(target, "")).strip()

    return row


def _sort_export_rows(rows):
    def _sort_key(row):
        element_id_text = _safe_str(row.get("ElementId", "")).strip()
        try:
            element_id_value = int(element_id_text)
        except Exception:
            element_id_value = 0
        return (
            _safe_str(row.get("NWB_Category", "")).strip().lower(),
            _safe_str(row.get("Type", "")).strip().lower(),
            _safe_str(row.get("Family", "")).strip().lower(),
            element_id_value,
        )

    return sorted(rows, key=_sort_key)


def _open_file(path):
    try:
        if path and os.path.isfile(path):
            os.startfile(path)
            return True
    except Exception:
        pass
    return False


def _selected_types_label(selected_types):
    if not selected_types:
        return "None"
    return ", ".join([_safe_str(value).strip() for value in selected_types if _safe_str(value).strip()])


def run():
    mapping = _load_mapping_module()
    mapping["_set_run_options"](
        {
            "arch_room_link_id": "",
            "site_override": "",
            "building_override": "",
            "design_pkg_override": "",
            "building_permit_override": "",
        }
    )

    workbook_path = mapping["_build_workbook_path"]()
    if not mapping["_validate_workbook_path"](workbook_path):
        forms.alert("Workbook not found:\n{0}".format(workbook_path), title=TOOL_TITLE)
        return

    try:
        workbook_data = mapping["_load_workbook_data"](workbook_path)
    except Exception as ex:
        forms.alert("Could not read the NWB workbook.\n\n{0}".format(_safe_str(ex)), title=TOOL_TITLE)
        return

    selected_types = _select_maintainability_types(workbook_data)
    if not selected_types:
        return

    elements = _collect_all_mep_elements(mapping)
    if not elements:
        forms.alert("No eligible MEP elements were found in the model.", title=TOOL_TITLE)
        return

    export_headers = _build_export_headers(mapping)
    export_rows = []
    match_cache = {}
    counters = {
        "total": len(elements),
        "matched": 0,
        "exported": 0,
        "no_match": 0,
        "filtered_out": 0,
    }

    try:
        with forms.ProgressBar(title="NWB maintainability export {value}/{max_value}", cancellable=True) as progress:
            progress.update_progress(0, counters["total"])

            for index, element in enumerate(elements, 1):
                if progress.cancelled:
                    break

                facts = _build_match_facts(mapping, element)
                asset_row, asset_score = _selected_asset_row(mapping, workbook_data, facts, match_cache)
                if not asset_row or int(asset_score or 0) <= 0:
                    counters["no_match"] += 1
                elif not _is_allowed_maintainability_row(asset_row, selected_types):
                    counters["matched"] += 1
                    counters["filtered_out"] += 1
                else:
                    counters["matched"] += 1
                    shared_values = _build_shared_parameter_map(mapping, element, mapping.get("TARGET_PARAMETERS", []))
                    export_rows.append(_build_export_row(mapping, element, asset_row, shared_values))
                    counters["exported"] += 1

                if index <= 10 or index % 25 == 0 or index == counters["total"]:
                    progress.update_progress(index, counters["total"])
                    try:
                        progress.title = "NWB maintainability export {0}/{1}".format(index, counters["total"])
                    except Exception:
                        pass

                if index % 25 == 0:
                    try:
                        mapping["_pump_ui"]()
                    except Exception:
                        pass
    except Exception as ex:
        forms.alert("Export failed.\n\n{0}".format(_safe_str(ex)), title=TOOL_TITLE)
        return

    if not export_rows:
        forms.alert(
            "No Building Services MEP elements matched the selected maintainability types.\n\n"
            "Selected types: {0}\n"
            "Scanned: {1}\nMatched asset rows: {2}\nNo asset match: {3}\nFiltered out by maintainability: {4}".format(
                _selected_types_label(selected_types),
                counters["total"],
                counters["matched"],
                counters["no_match"],
                counters["filtered_out"],
            ),
            title=TOOL_TITLE,
        )
        return

    export_rows = _sort_export_rows(export_rows)
    stamp = _timestamp()
    base_name = "NWB_Maintainability_{0}_Schedule_{1}".format(_types_slug(selected_types), stamp)
    csv_path = os.path.join(_logs_folder(), base_name + ".csv")
    xlsx_path = os.path.join(_logs_folder(), base_name + ".xlsx")
    sheet_name = _sheet_name_for_types(selected_types)

    _write_csv(csv_path, export_headers, export_rows)

    output_path = csv_path
    export_notice = ""
    try:
        _write_xlsx(xlsx_path, export_headers, export_rows, sheet_name)
        output_path = xlsx_path
    except Exception as ex:
        export_notice = "\nExcel export failed, so CSV was kept instead.\n{0}".format(_safe_str(ex))

    _open_file(output_path)

    forms.alert(
        "{0} finished.\n\n"
        "Workbook: {1}\n"
        "Selected maintainability types: {2}\n"
        "Scanned MEP elements: {3}\n"
        "Matched asset rows: {4}\n"
        "Exported elements: {5}\n"
        "No asset match: {6}\n"
        "Filtered out by maintainability: {7}\n"
        "Output: {8}{9}".format(
            TOOL_TITLE,
            workbook_path,
            _selected_types_label(selected_types),
            counters["total"],
            counters["matched"],
            counters["exported"],
            counters["no_match"],
            counters["filtered_out"],
            output_path,
            export_notice,
        ),
        title=TOOL_TITLE,
    )


if __name__ == "__main__":
    try:
        run()
    except Exception as ex:
        forms.alert("{0} failed.\n\n{1}".format(TOOL_TITLE, _safe_str(ex)), title=TOOL_TITLE)