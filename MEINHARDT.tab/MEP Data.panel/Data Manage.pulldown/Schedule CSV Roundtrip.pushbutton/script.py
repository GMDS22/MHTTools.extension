# coding: utf8
from __future__ import print_function

import csv
import json
import os
import sys
import time
import traceback
from datetime import datetime

import clr

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Application as WinFormsApplication
from System.Windows.Forms import DialogResult, OpenFileDialog, SaveFileDialog

from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    SectionType,
    StorageType,
    Transaction,
    TransactionGroup,
    ViewSchedule,
)
from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow

if sys.version_info[0] >= 3:
    unicode = str


doc = revit.doc
logger = script.get_logger()
config = script.get_config()

__title__ = "Schedule\nManager"
__doc__ = "Export schedule-scoped instance parameter values to CSV/Excel, preview changes, and import with verified batched writes."

PARAMETER_TARGET_MODES = [
    "SharedOnly",
    "FamilyOnly",
    "Both",
    "FirstMatch",
]

DEFAULT_PARAMETER_TARGET_MODE = "SharedOnly"

META_ELEMENT_ID = "__MHT_ElementId"
META_UNIQUE_ID = "__MHT_UniqueId"
META_SCHEDULE_NAME = "__MHT_ScheduleName"
META_EXPORTED_AT = "__MHT_ExportedAt"

META_COLUMNS = (META_ELEMENT_ID, META_UNIQUE_ID, META_SCHEDULE_NAME, META_EXPORTED_AT)

MAX_EXPORT_ROWS_CONFIRM = 5000
MAX_IMPORT_ROWS_CONFIRM = 5000
MAX_PREVIEW_TABLE_ROWS = 1200
PREVIEW_LOG_SAMPLE_LIMIT = 18
IMPORT_BATCH_SIZE = 750
DOUBLE_TOLERANCE = 1e-6

EXPORT_TEMPLATE_FILE_NAME = "export_template.json"
DEFAULT_EXPORT_TEMPLATE = {
    "template_version": 1,
    "header_prefix": [
        META_ELEMENT_ID,
        META_UNIQUE_ID,
        META_SCHEDULE_NAME,
        META_EXPORTED_AT,
    ],
    "excel": {
        "sheet_name": "Schedule Manager Export",
        "header_bold": True,
        "header_fill_color": "#0E639C",
        "header_font_color": "#FFFFFF",
        "freeze_header_row": True,
        "auto_filter": True,
        "auto_fit_columns": True,
    },
}


class ImportPreviewRow(object):
    def __init__(self, row_number, element_id, column_name, current_value, incoming_value, status):
        self.Row = _safe_str(row_number)
        self.ElementId = _safe_str(element_id)
        self.Column = _safe_str(column_name)
        self.Current = _safe_str(current_value)
        self.Incoming = _safe_str(incoming_value)
        self.Status = _safe_str(status)


class OperationMetrics(object):
    def __init__(self, operation_name):
        self.operation_name = operation_name
        self.started = time.time()
        self.ended = None

    def stop(self):
        self.ended = time.time()

    @property
    def elapsed(self):
        end = self.ended if self.ended is not None else time.time()
        return max(0.0, end - self.started)


class RuntimeCache(object):
    def __init__(self):
        self.element_cache = {}
        self.param_map_cache = {}
        self.param_cache = {}
        self.param_display_cache = {}

    def _element_id_value(self, element):
        try:
            return int(element.Id.IntegerValue)
        except Exception:
            return None

    def get_element_from_row(self, row):
        element_id_text = _safe_str(row.get(META_ELEMENT_ID, "")).strip()
        unique_id_text = _safe_str(row.get(META_UNIQUE_ID, "")).strip()
        cache_key = element_id_text if element_id_text else "uid:{0}".format(unique_id_text)

        if cache_key in self.element_cache:
            return self.element_cache[cache_key]

        element = None
        if element_id_text:
            try:
                element = doc.GetElement(ElementId(int(element_id_text)))
            except Exception:
                element = None
        if element is None and unique_id_text:
            try:
                element = doc.GetElement(unique_id_text)
            except Exception:
                element = None

        if not _is_valid_element(element):
            element = None

        self.element_cache[cache_key] = element
        return element

    def get_parameter_map(self, element):
        element_id = self._element_id_value(element)
        if element_id is None:
            return {}

        cached = self.param_map_cache.get(element_id)
        if cached is not None:
            return cached

        param_map = {}
        for scope_name, owner in _iter_param_owners(element):
            if owner is None:
                continue
            try:
                for parameter in owner.Parameters:
                    try:
                        definition = parameter.Definition
                        if definition is None:
                            continue
                        name_key = _safe_str(definition.Name).strip().lower()
                        if not name_key:
                            continue
                        param_map.setdefault(name_key, []).append({
                            "param": parameter,
                            "scope": scope_name,
                            "is_shared": _param_is_shared(parameter),
                        })
                    except Exception:
                        continue
            except Exception:
                continue

        for name_key, candidates in list(param_map.items()):
            param_map[name_key] = sorted(candidates, key=_param_candidate_sort_key)

        self.param_map_cache[element_id] = param_map
        return param_map

    def get_parameter(self, element, parameter_name, target_mode=DEFAULT_PARAMETER_TARGET_MODE):
        element_id = self._element_id_value(element)
        if element_id is None:
            return None

        param_key = _safe_str(parameter_name).strip().lower()
        if not param_key:
            return None

        cache_key = (element_id, param_key, target_mode)
        if cache_key in self.param_cache:
            return self.param_cache[cache_key]

        param_map = self.get_parameter_map(element)
        chosen, _ = _resolve_param_candidate(param_map.get(param_key, []), target_mode)
        parameter = chosen["param"] if chosen is not None else None
        self.param_cache[cache_key] = parameter
        return parameter

    def get_display_value(self, element, parameter_name, target_mode=DEFAULT_PARAMETER_TARGET_MODE):
        element_id = self._element_id_value(element)
        if element_id is None:
            return ""

        param_key = _safe_str(parameter_name).strip().lower()
        cache_key = (element_id, param_key, target_mode)
        if cache_key in self.param_display_cache:
            return self.param_display_cache[cache_key]

        parameter = self.get_parameter(element, parameter_name, target_mode)
        value = _display_parameter_value(parameter)
        self.param_display_cache[cache_key] = value
        return value

    def invalidate_parameter_display(self, element, parameter_name, target_mode=None):
        element_id = self._element_id_value(element)
        if element_id is None:
            return
        param_key = _safe_str(parameter_name).strip().lower()
        if target_mode:
            cache_key = (element_id, param_key, target_mode)
            if cache_key in self.param_display_cache:
                del self.param_display_cache[cache_key]
            return

        for cache_key in list(self.param_display_cache.keys()):
            if cache_key[0] == element_id and cache_key[1] == param_key:
                del self.param_display_cache[cache_key]


class PreviewResult(object):
    def __init__(self):
        self.total_rows = 0
        self.schedule_name = ""
        self.schedule_view = None
        self.headers = []
        self.mapped_columns = []
        self.unsupported_columns = []
        self.preview_rows = []
        self.summary_lines = []
        self.missing_elements = 0
        self.changed_cells = 0
        self.no_change_cells = 0
        self.read_only_cells = 0
        self.missing_parameter_cells = 0
        self.blank_skipped = 0
        self.unsupported_type_cells = 0
        self.cancelled = False


class ImportResult(object):
    def __init__(self):
        self.total_rows = 0
        self.processed_rows = 0
        self.total_cells_seen = 0
        self.written_cells = 0
        self.no_change_cells = 0
        self.read_only_cells = 0
        self.missing_parameter_cells = 0
        self.missing_elements = 0
        self.unsupported_columns = 0
        self.unsupported_type_cells = 0
        self.blank_skipped = 0
        self.failed_cells = 0
        self.verification_failures = 0
        self.cancelled = False
        self.stopped_on_failure = False
        self.sample_failures = []


def _safe_str(value):
    try:
        if value is None:
            return ""
        return str(value)
    except Exception:
        try:
            return unicode(value)
        except Exception:
            return ""


def _pump_ui():
    try:
        WinFormsApplication.DoEvents()
    except Exception:
        pass


def _sanitize_filename_token(value):
    text = _safe_str(value).strip().replace(" ", "_")
    if not text:
        return "Schedule"

    invalid = set('\\/:*?"<>|')
    clean_chars = []
    for char in text:
        clean_chars.append("_" if char in invalid else char)

    cleaned = "".join(clean_chars).strip("._")
    return cleaned or "Schedule"


def _build_default_export_filename(schedule_name):
    schedule_part = _sanitize_filename_token(schedule_name)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return "Schedule Manager -_{0}_{1}.csv".format(schedule_part, stamp)


def _template_file_path():
    return os.path.join(os.path.dirname(__file__), EXPORT_TEMPLATE_FILE_NAME)


def _ensure_export_template_file():
    path = _template_file_path()
    if os.path.exists(path):
        return path

    try:
        with open(path, "w") as stream:
            json.dump(DEFAULT_EXPORT_TEMPLATE, stream, indent=2, sort_keys=True)
    except Exception:
        pass
    return path


def _load_export_template():
    path = _ensure_export_template_file()
    template = dict(DEFAULT_EXPORT_TEMPLATE)

    try:
        with open(path, "r") as stream:
            loaded = json.load(stream)
            if isinstance(loaded, dict):
                template.update(loaded)
    except Exception:
        pass

    excel_defaults = dict(DEFAULT_EXPORT_TEMPLATE.get("excel", {}))
    excel_raw = template.get("excel") if isinstance(template.get("excel"), dict) else {}
    excel_defaults.update(excel_raw)
    template["excel"] = excel_defaults

    if not isinstance(template.get("header_prefix"), list):
        template["header_prefix"] = list(DEFAULT_EXPORT_TEMPLATE["header_prefix"])

    return template


def _ordered_headers_by_template(headers, template):
    if not headers:
        return []

    header_prefix = template.get("header_prefix", []) if isinstance(template, dict) else []
    if not isinstance(header_prefix, list) or not header_prefix:
        return list(headers)

    lower_map = {}
    for header in headers:
        key = _safe_str(header).strip().lower()
        if key and key not in lower_map:
            lower_map[key] = header

    ordered = []
    used = set()
    for desired in header_prefix:
        key = _safe_str(desired).strip().lower()
        if not key:
            continue
        actual = lower_map.get(key)
        if actual is None or actual in used:
            continue
        ordered.append(actual)
        used.add(actual)

    for header in headers:
        if header not in used:
            ordered.append(header)
    return ordered


def _hex_to_excel_color(hex_color):
    text = _safe_str(hex_color).strip().lstrip("#")
    if len(text) != 6:
        return None
    try:
        red = int(text[0:2], 16)
        green = int(text[2:4], 16)
        blue = int(text[4:6], 16)
        return red + (green * 256) + (blue * 65536)
    except Exception:
        return None


def _is_valid_element(element):
    if element is None:
        return False
    try:
        return bool(getattr(element, "IsValidObject", True))
    except Exception:
        return False


def _iter_element_params(element):
    try:
        for parameter in element.Parameters:
            yield parameter
    except Exception:
        return


def _iter_param_owners(element):
    if element is None:
        return

    yield ("instance", element)

    try:
        symbol = getattr(element, "Symbol", None)
        if symbol is not None:
            yield ("type", symbol)
            return
    except Exception:
        pass

    try:
        type_id = getattr(element, "GetTypeId", lambda: None)()
        if type_id is not None and type_id != ElementId.InvalidElementId:
            type_el = doc.GetElement(type_id)
            if type_el is not None:
                yield ("type", type_el)
    except Exception:
        pass


def _param_is_shared(parameter):
    try:
        return bool(getattr(parameter, "IsShared", False))
    except Exception:
        return False


def _param_candidate_sort_key(candidate):
    scope_priority = 0 if candidate.get("scope") == "instance" else 1
    shared_priority = 0 if candidate.get("is_shared") else 1
    return (
        scope_priority,
        shared_priority,
        _safe_str(getattr(candidate.get("param"), "StorageType", "")),
        _safe_str(getattr(getattr(candidate.get("param"), "Definition", None), "Name", "")),
    )


def _resolve_param_candidate(candidates, target_mode):
    if not candidates:
        return None, "missing"

    writable = []
    for candidate in candidates:
        try:
            if not candidate["param"].IsReadOnly:
                writable.append(candidate)
        except Exception:
            continue

    if not writable:
        return None, "read-only"

    shared = [candidate for candidate in writable if candidate.get("is_shared")]
    nonshared = [candidate for candidate in writable if not candidate.get("is_shared")]

    mode = target_mode if target_mode in PARAMETER_TARGET_MODES else DEFAULT_PARAMETER_TARGET_MODE
    if mode == "FamilyOnly":
        if nonshared:
            return nonshared[0], "family-only"
        return None, "no-family"

    if mode == "Both":
        if shared:
            return shared[0], "both"
        return writable[0], "both"

    if mode == "FirstMatch":
        return writable[0], "first-match"

    if shared:
        return shared[0], "shared-only"
    return None, "no-shared"


def _display_parameter_value(parameter):
    if parameter is None:
        return "<missing>"

    try:
        storage_type = parameter.StorageType
    except Exception:
        storage_type = None

    try:
        if storage_type == StorageType.String:
            raw = parameter.AsString()
            return _safe_str(raw) if raw not in (None, "") else ""
        if storage_type == StorageType.Integer:
            display = parameter.AsValueString()
            if display not in (None, ""):
                return _safe_str(display)
            return _safe_str(parameter.AsInteger())
        if storage_type == StorageType.Double:
            display = parameter.AsValueString()
            if display not in (None, ""):
                return _safe_str(display)
            return _safe_str(parameter.AsDouble())
        if storage_type == StorageType.ElementId:
            display = parameter.AsValueString()
            if display not in (None, ""):
                return _safe_str(display)
            element_id = parameter.AsElementId()
            if element_id is None or element_id == ElementId.InvalidElementId:
                return ""
            return _safe_str(element_id.IntegerValue)
    except Exception:
        pass

    try:
        display = parameter.AsValueString()
        if display not in (None, ""):
            return _safe_str(display)
    except Exception:
        pass

    try:
        raw = parameter.AsString()
        if raw not in (None, ""):
            return _safe_str(raw)
    except Exception:
        pass

    return ""


def _read_parameter_value(parameter):
    if parameter is None:
        return None

    try:
        storage_type = parameter.StorageType
        if storage_type == StorageType.String:
            return parameter.AsString() or ""
        if storage_type == StorageType.Integer:
            return parameter.AsInteger()
        if storage_type == StorageType.Double:
            return parameter.AsDouble()
        if storage_type == StorageType.ElementId:
            element_id = parameter.AsElementId()
            if element_id is None or element_id == ElementId.InvalidElementId:
                return None
            return int(element_id.IntegerValue)
    except Exception:
        pass
    return None


def _values_equal(left_value, right_value, storage_type):
    if storage_type == StorageType.String:
        return _safe_str(left_value) == _safe_str(right_value)
    if storage_type == StorageType.Integer:
        try:
            return int(left_value) == int(right_value)
        except Exception:
            return False
    if storage_type == StorageType.Double:
        try:
            return abs(float(left_value) - float(right_value)) <= DOUBLE_TOLERANCE
        except Exception:
            return False
    if storage_type == StorageType.ElementId:
        try:
            return int(left_value) == int(right_value)
        except Exception:
            return False
    return _safe_str(left_value) == _safe_str(right_value)


def _parse_numeric_text(raw_value):
    text = _safe_str(raw_value).strip()
    if not text:
        raise ValueError("blank")
    return float(text.replace(",", "."))


def _parameter_value_matches_incoming(parameter, incoming_text):
    if parameter is None:
        return False

    incoming = _safe_str(incoming_text)

    try:
        storage_type = parameter.StorageType
    except Exception:
        storage_type = None

    if storage_type == StorageType.String:
        current = parameter.AsString() or ""
        return _safe_str(current) == incoming

    if storage_type == StorageType.Integer:
        try:
            return int(parameter.AsInteger()) == int(_parse_numeric_text(incoming))
        except Exception:
            try:
                return _safe_str(parameter.AsValueString()) == incoming
            except Exception:
                return False

    if storage_type == StorageType.Double:
        try:
            return abs(float(parameter.AsDouble()) - float(_parse_numeric_text(incoming))) <= DOUBLE_TOLERANCE
        except Exception:
            try:
                return _safe_str(parameter.AsValueString()) == incoming
            except Exception:
                return False

    if storage_type == StorageType.ElementId:
        try:
            current_id = parameter.AsElementId()
            current_int = int(current_id.IntegerValue) if current_id is not None else None
            return current_int == int(_parse_numeric_text(incoming))
        except Exception:
            try:
                return _safe_str(parameter.AsValueString()) == incoming
            except Exception:
                return False

    try:
        return _safe_str(parameter.AsValueString()) == incoming
    except Exception:
        return False


def _set_parameter_value(parameter, incoming_value):
    if parameter is None:
        return False, "missing parameter"
    if parameter.IsReadOnly:
        return False, "read-only parameter"

    value_text = _safe_str(incoming_value)

    try:
        storage_type = parameter.StorageType
    except Exception:
        return False, "unknown storage type"

    current_value = _read_parameter_value(parameter)
    target_value = value_text
    success = False

    try:
        if storage_type == StorageType.String:
            if _values_equal(current_value, value_text, storage_type):
                return False, "no change"
            success = parameter.Set(value_text)
            target_value = value_text
        elif storage_type == StorageType.Integer:
            if value_text.strip() == "":
                return False, "blank integer value"
            if hasattr(parameter, "SetValueString"):
                try:
                    success = parameter.SetValueString(value_text)
                    target_value = _read_parameter_value(parameter)
                except Exception:
                    success = False
            if not success:
                try:
                    target_value = int(float(value_text.strip()))
                    if _values_equal(current_value, target_value, storage_type):
                        return False, "no change"
                    success = parameter.Set(target_value)
                except Exception:
                    return False, "invalid integer value"
        elif storage_type == StorageType.Double:
            if value_text.strip() == "":
                return False, "blank numeric value"
            if hasattr(parameter, "SetValueString"):
                try:
                    success = parameter.SetValueString(value_text)
                    target_value = _read_parameter_value(parameter)
                except Exception:
                    success = False
            if not success:
                try:
                    target_value = float(value_text.strip().replace(",", "."))
                    if _values_equal(current_value, target_value, storage_type):
                        return False, "no change"
                    success = parameter.Set(target_value)
                except Exception:
                    return False, "invalid double value"
        else:
            return False, "unsupported storage type"

        if not success:
            return False, "set returned False"

        verified = _read_parameter_value(parameter)
        if target_value is None:
            target_value = _read_parameter_value(parameter)
        if not _values_equal(verified, target_value, storage_type):
            return False, "write verification failed"

        return True, "written"
    except Exception as ex:
        return False, _safe_str(ex)


def _confirm(message, title="Confirm"):
    try:
        result = forms.alert(message, title=title, yes=True, no=True)
        if isinstance(result, bool):
            return result
        return _safe_str(result).strip().lower() in ("yes", "y", "true", "ok", "1")
    except Exception:
        return False


def _schedule_is_supported(schedule_view):
    if schedule_view is None:
        return False
    try:
        if getattr(schedule_view, "IsTemplate", False):
            return False
    except Exception:
        pass
    try:
        if schedule_view.Definition.IsKeySchedule:
            return False
    except Exception:
        pass
    try:
        if getattr(schedule_view, "IsTitleblockRevisionSchedule", False):
            return False
    except Exception:
        pass
    return True


def _get_supported_schedules():
    schedules = []
    try:
        for schedule_view in FilteredElementCollector(doc).OfClass(ViewSchedule):
            if _schedule_is_supported(schedule_view):
                schedules.append(schedule_view)
    except Exception:
        return []

    schedules.sort(key=lambda x: _safe_str(x.Name).lower())
    return schedules


def _get_schedule_field_parameter_name(schedule_field):
    try:
        parameter_id = schedule_field.ParameterId
        if parameter_id and parameter_id != ElementId.InvalidElementId:
            parameter_element = doc.GetElement(parameter_id)
            if parameter_element is not None:
                return _safe_str(parameter_element.Name)
    except Exception:
        pass

    try:
        return _safe_str(schedule_field.GetName())
    except Exception:
        return ""


def _make_unique_heading(raw_heading, taken):
    base = _safe_str(raw_heading).strip() or "Column"
    candidate = base
    suffix = 2
    lower_taken = taken
    while candidate.lower() in lower_taken:
        candidate = "{0} ({1})".format(base, suffix)
        suffix += 1
    lower_taken.add(candidate.lower())
    return candidate


def _get_schedule_field_infos(schedule_view, include_hidden=False):
    field_infos = []
    definition = schedule_view.Definition
    used = set()
    visible_col_offset = 0

    try:
        body_section = schedule_view.GetTableData().GetSectionData(SectionType.Body)
        first_body_col = int(body_section.FirstColumnNumber)
    except Exception:
        first_body_col = 0

    for field_id in definition.GetFieldOrder():
        try:
            field = definition.GetField(field_id)
        except Exception:
            continue

        try:
            is_hidden = bool(field.IsHidden)
        except Exception:
            is_hidden = False

        if is_hidden and not include_hidden:
            continue

        heading = _make_unique_heading(getattr(field, "ColumnHeading", None) or field.GetName(), used)
        parameter_name = _get_schedule_field_parameter_name(field)
        importable = bool(parameter_name)

        body_col = None
        if not is_hidden:
            body_col = first_body_col + visible_col_offset
            visible_col_offset += 1

        field_infos.append({
            "heading": heading,
            "parameter_name": parameter_name,
            "importable": importable,
            "is_hidden": is_hidden,
            "body_col": body_col,
        })

    return field_infos


def _collect_schedule_elements(schedule_view):
    try:
        elements = list(
            FilteredElementCollector(doc, schedule_view.Id)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        elements = []

    return [el for el in elements if _is_valid_element(el)]


def _get_schedule_body_row_numbers(schedule_view):
    try:
        body = schedule_view.GetTableData().GetSectionData(SectionType.Body)
        start = int(body.FirstRowNumber)
        end = int(body.LastRowNumber)
        if end < start:
            return []
        return list(range(start, end + 1))
    except Exception:
        return []


def _get_schedule_cell_text(schedule_view, row_number, col_number):
    try:
        return _safe_str(schedule_view.GetCellText(SectionType.Body, int(row_number), int(col_number)))
    except Exception:
        return ""


def _get_schedule_export_definition(schedule_view, include_hidden, cache):
    elements = _collect_schedule_elements(schedule_view)
    body_rows = _get_schedule_body_row_numbers(schedule_view)
    field_infos = _get_schedule_field_infos(schedule_view, include_hidden=include_hidden)
    importable_count = 0
    for field in field_infos:
        if field.get("importable"):
            importable_count += 1
    return elements, body_rows, field_infos, importable_count


def _compose_export_headers(field_infos, template):
    headers = list(META_COLUMNS)
    headers.extend(field["heading"] for field in field_infos)
    return _ordered_headers_by_template(headers, template)


def _yield_export_row_dicts(schedule_view, elements, body_rows, field_infos, exported_at, cache, parameter_target_mode=DEFAULT_PARAMETER_TARGET_MODE):
    schedule_name = _safe_str(schedule_view.Name)

    for row_index, row_number in enumerate(body_rows):
        element = elements[row_index] if row_index < len(elements) else None
        row = {
            META_ELEMENT_ID: _safe_str(element.Id.IntegerValue) if element is not None else "",
            META_UNIQUE_ID: _safe_str(getattr(element, "UniqueId", "")) if element is not None else "",
            META_SCHEDULE_NAME: schedule_name,
            META_EXPORTED_AT: exported_at,
        }

        for field_info in field_infos:
            heading = field_info["heading"]
            body_col = field_info.get("body_col")

            if body_col is not None:
                row[heading] = _get_schedule_cell_text(schedule_view, row_number, body_col)
                continue

            parameter_name = field_info.get("parameter_name")
            if element is not None and parameter_name:
                row[heading] = cache.get_display_value(element, parameter_name, parameter_target_mode)
            else:
                row[heading] = ""

        yield row


def _write_csv_stream(path, headers, row_iterable):
    if sys.version_info[0] >= 3:
        with open(path, "w", newline="", encoding="utf-8-sig") as stream:
            writer = csv.writer(stream)
            writer.writerow(headers)
            for row in row_iterable:
                writer.writerow([_safe_str(row.get(header, "")) for header in headers])
        return

    with open(path, "wb") as stream:
        writer = csv.writer(stream)
        writer.writerow([_safe_str(header).encode("utf-8") for header in headers])
        for row in row_iterable:
            out = []
            for header in headers:
                out.append(_safe_str(row.get(header, "")).encode("utf-8"))
            writer.writerow(out)


def _write_excel_stream(path, headers, row_iterable, template):
    Excel = None
    app = None
    load_errors = []

    # Try version-agnostic Excel interop first.
    try:
        clr.AddReference("Microsoft.Office.Interop.Excel")
        from Microsoft.Office.Interop import Excel as Excel  # pylint: disable=import-error
    except Exception as ex:
        load_errors.append(_safe_str(ex))

    # Fallback for environments where only a specific PIA version is available.
    if Excel is None:
        interop_versions = (16, 15, 14, 12, 11)
        for major in interop_versions:
            try:
                clr.AddReferenceByName(
                    "Microsoft.Office.Interop.Excel, Version={0}.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c".format(major)
                )
                from Microsoft.Office.Interop import Excel as Excel  # pylint: disable=import-error
                break
            except Exception as ex:
                load_errors.append(_safe_str(ex))

    # Final fallback: create Excel COM instance directly.
    if Excel is not None:
        try:
            app = Excel.ApplicationClass()
        except Exception as ex:
            load_errors.append(_safe_str(ex))
            app = None

    if app is None:
        try:
            from System import Activator, Type

            excel_type = Type.GetTypeFromProgID("Excel.Application")
            if excel_type is not None:
                app = Activator.CreateInstance(excel_type)
        except Exception as ex:
            load_errors.append(_safe_str(ex))

    if app is None:
        detail = ""
        if load_errors:
            detail = "\n\nDetails: {0}".format(" | ".join([e for e in load_errors if e][:3]))
        raise RuntimeError(
            "Excel export requires Microsoft Excel desktop to be installed and registered on this machine.{0}".format(detail)
        )

    excel_cfg = template.get("excel", {}) if isinstance(template, dict) else {}
    sheet_name = _safe_str(excel_cfg.get("sheet_name", "Schedule Manager Export")) or "Schedule Manager Export"
    header_fill = _hex_to_excel_color(excel_cfg.get("header_fill_color", "#0E639C"))
    header_font = _hex_to_excel_color(excel_cfg.get("header_font_color", "#FFFFFF"))
    header_bold = bool(excel_cfg.get("header_bold", True))
    freeze_header_row = bool(excel_cfg.get("freeze_header_row", True))
    auto_filter = bool(excel_cfg.get("auto_filter", True))
    auto_fit_columns = bool(excel_cfg.get("auto_fit_columns", True))

    workbook = None

    try:
        app.Visible = False
        app.DisplayAlerts = False

        workbook = app.Workbooks.Add()
        worksheet = workbook.Worksheets[1]
        try:
            worksheet.Name = sheet_name[:31]
        except Exception:
            pass

        for col_index, header in enumerate(headers, 1):
            cell = worksheet.Cells[1, col_index]
            cell.Value2 = _safe_str(header)
            try:
                cell.Font.Bold = header_bold
            except Exception:
                pass
            if header_fill is not None:
                try:
                    cell.Interior.Color = header_fill
                except Exception:
                    pass
            if header_font is not None:
                try:
                    cell.Font.Color = header_font
                except Exception:
                    pass

        row_number = 2
        for row in row_iterable:
            for col_index, header in enumerate(headers, 1):
                worksheet.Cells[row_number, col_index].Value2 = _safe_str(row.get(header, ""))
            row_number += 1

        last_row = row_number - 1
        last_col = len(headers)
        if last_row >= 1 and last_col > 0:
            top_left = worksheet.Cells[1, 1]
            bottom_right = worksheet.Cells[last_row, last_col]
            used_range = worksheet.Range[top_left, bottom_right]

            if auto_filter:
                try:
                    used_range.AutoFilter()
                except Exception:
                    pass

            if auto_fit_columns:
                try:
                    used_range.Columns.AutoFit()
                except Exception:
                    pass

            if freeze_header_row:
                try:
                    app.ActiveWindow.SplitRow = 1
                    app.ActiveWindow.FreezePanes = True
                except Exception:
                    pass

        workbook.SaveAs(path, 51)
    except Exception as ex:
        raise RuntimeError("Excel export failed: {0}".format(_safe_str(ex)))
    finally:
        try:
            if workbook is not None:
                workbook.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass


def _count_csv_data_rows(path):
    total = 0
    if sys.version_info[0] >= 3:
        with open(path, "r", newline="", encoding="utf-8-sig") as stream:
            for _ in stream:
                total += 1
    else:
        with open(path, "rb") as stream:
            for _ in stream:
                total += 1
    return max(0, total - 1)


def _iter_csv_dict_rows(path):
    if sys.version_info[0] >= 3:
        stream = open(path, "r", newline="", encoding="utf-8-sig")
        try:
            reader = csv.DictReader(stream)
            headers = list(reader.fieldnames or [])
            yield headers, None
            for row in reader:
                yield None, row or {}
        finally:
            stream.close()
        return

    stream = open(path, "rb")
    try:
        reader = csv.reader(stream)
        try:
            first = next(reader)
        except StopIteration:
            first = []

        headers = []
        for item in first:
            text = _safe_str(item)
            if text.startswith("\ufeff"):
                text = text.lstrip("\ufeff")
            headers.append(text)

        yield headers, None

        for raw in reader:
            row = {}
            for i, header in enumerate(headers):
                row[header] = _safe_str(raw[i] if i < len(raw) else "")
            yield None, row
    finally:
        stream.close()


def _normalize_column_key(value):
    text = _safe_str(value).strip().lower()
    if not text:
        return ""

    if text.endswith(")"):
        pivot = text.rfind(" (")
        if pivot > 0:
            suffix = text[pivot + 2 : -1]
            if suffix.isdigit():
                return text[:pivot].strip()
    return text


def _build_import_column_map(schedule_view, headers):
    mapped_columns = []
    unsupported_columns = []

    by_heading = {}
    by_heading_norm = {}
    by_param = {}
    by_param_norm = {}

    if schedule_view is not None:
        for info in _get_schedule_field_infos(schedule_view, include_hidden=False):
            if not info.get("importable"):
                continue

            heading = _safe_str(info.get("heading", "")).strip()
            param_name = _safe_str(info.get("parameter_name", "")).strip()
            if not param_name:
                continue

            h_key = heading.lower()
            p_key = param_name.lower()
            h_norm = _normalize_column_key(heading)
            p_norm = _normalize_column_key(param_name)

            if h_key and h_key not in by_heading:
                by_heading[h_key] = param_name
            if p_key and p_key not in by_param:
                by_param[p_key] = param_name
            if h_norm and h_norm not in by_heading_norm:
                by_heading_norm[h_norm] = param_name
            if p_norm and p_norm not in by_param_norm:
                by_param_norm[p_norm] = param_name

    for header in headers:
        if header in META_COLUMNS:
            continue

        raw = _safe_str(header).strip()
        if not raw:
            continue

        parameter_name = None
        if schedule_view is not None:
            key = raw.lower()
            norm = _normalize_column_key(raw)
            parameter_name = (
                by_heading.get(key)
                or by_param.get(key)
                or by_heading_norm.get(norm)
                or by_param_norm.get(norm)
            )
            if not parameter_name:
                unsupported_columns.append(raw)
                continue
        else:
            parameter_name = raw

        mapped_columns.append({"column": raw, "parameter_name": parameter_name})

    return mapped_columns, unsupported_columns


def _find_schedule_by_name(schedule_name):
    target = _safe_str(schedule_name).strip().lower()
    if not target:
        return None
    for schedule_view in _get_supported_schedules():
        try:
            if _safe_str(schedule_view.Name).strip().lower() == target:
                return schedule_view
        except Exception:
            continue
    return None


def _estimate_import_seconds(row_count, candidate_cells):
    return max(0.5, 0.45 + (float(row_count) * 0.01) + (float(candidate_cells) * 0.005))


def _format_metrics_line(metrics, rows_processed, cells_written):
    elapsed = metrics.elapsed
    rps = (float(rows_processed) / elapsed) if elapsed > 0 else 0.0
    return (
        "Elapsed: {0:.2f}s | Rows processed: {1} | Rows/sec: {2:.2f} | Cells written: {3}".format(
            elapsed,
            rows_processed,
            rps,
            cells_written,
        )
    )


def _try_update_progress(pb, current, maximum):
    if pb is None:
        return
    try:
        pb.update_progress(current, maximum)
    except Exception:
        pass


class ScheduleCsvRoundtripWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        self._is_busy = False
        self._all_schedules = []
        self._schedules = []
        self._schedule_lookup = {}
        self._last_preview = None
        self._last_preview_signature = None
        self._export_template = _load_export_template()

    def window_loaded(self, sender, e):
        self._load_schedules()
        self._seed_current_schedule()
        self._load_last_import_path()
        self._load_parameter_target_mode()
        self._refresh_export_button_text()

    def _set_status(self, message):
        self.txtStatus.Text = _safe_str(message)

    def _append_log(self, message):
        stamp = datetime.now().strftime("%H:%M:%S")
        existing = _safe_str(getattr(self.txtLog, "Text", ""))
        line = "[{0}] {1}".format(stamp, _safe_str(message))
        self.txtLog.Text = (existing + "\r\n" + line).strip() if existing else line
        try:
            self.txtLog.ScrollToEnd()
        except Exception:
            pass

    def _selected_schedule(self):
        item = self.cmbSchedule.SelectedItem
        if item is None:
            return None
        return self._schedule_lookup.get(_safe_str(item))

    def _selected_export_format(self):
        try:
            item = self.cmbExportFormat.SelectedItem
            if item is not None:
                content = getattr(item, "Content", None)
                if content is not None:
                    text = _safe_str(content).strip().lower()
                    if text in ("csv", "excel"):
                        return text
        except Exception:
            pass
        return "csv"

    def _refresh_export_button_text(self):
        fmt = self._selected_export_format()
        btn_main = getattr(self, "btnExportCsv", None)
        btn_head = getattr(self, "btnHeaderExport", None)

        if fmt == "excel":
            if btn_main is not None:
                btn_main.Content = "Export Excel"
            if btn_head is not None:
                btn_head.Content = "Export Excel"
        else:
            if btn_main is not None:
                btn_main.Content = "Export CSV"
            if btn_head is not None:
                btn_head.Content = "Export CSV"

    def _selected_parameter_target_mode(self):
        combo = getattr(self, "cmbParameterTargetMode", None)
        if combo is not None:
            try:
                item = combo.SelectedItem
                if item is not None:
                    raw = getattr(item, "Tag", None)
                    if raw in (None, ""):
                        raw = getattr(item, "Content", None)
                    value = _safe_str(raw).strip()
                    if value in PARAMETER_TARGET_MODES:
                        return value
            except Exception:
                pass
        return DEFAULT_PARAMETER_TARGET_MODE

    def _load_parameter_target_mode(self):
        saved_value = _safe_str(getattr(config, "schedule_csv_roundtrip_parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE)).strip()
        desired = saved_value if saved_value in PARAMETER_TARGET_MODES else DEFAULT_PARAMETER_TARGET_MODE
        combo = getattr(self, "cmbParameterTargetMode", None)
        if combo is None:
            return

        try:
            for idx in range(combo.Items.Count):
                item = combo.Items[idx]
                raw = getattr(item, "Tag", None)
                if raw in (None, ""):
                    raw = getattr(item, "Content", None)
                if _safe_str(raw).strip() == desired:
                    combo.SelectedIndex = idx
                    return
        except Exception:
            pass

        try:
            combo.SelectedIndex = 0
        except Exception:
            pass

    def _save_parameter_target_mode(self):
        try:
            config.schedule_csv_roundtrip_parameter_target_mode = self._selected_parameter_target_mode()
            script.save_config()
        except Exception:
            pass

    def parameter_target_mode_changed(self, sender, e):
        self._save_parameter_target_mode()
        self._last_preview = None
        self._last_preview_signature = None
        self._set_status("Parameter target mode updated to {0}.".format(self._selected_parameter_target_mode()))

    def _schedule_search_text(self):
        box = getattr(self, "txtScheduleSearch", None)
        if box is None:
            return ""
        return _safe_str(getattr(box, "Text", "")).strip().lower()

    def _apply_schedule_filter(self, preserve_selection_name=None):
        query = self._schedule_search_text()
        selected_name = _safe_str(preserve_selection_name or self.cmbSchedule.SelectedItem)

        self.cmbSchedule.Items.Clear()
        self._schedules = []
        self._schedule_lookup = {}

        query_parts = [part for part in query.split() if part]
        for schedule_view in self._all_schedules:
            name = _safe_str(schedule_view.Name)
            lowered = name.lower()
            if query_parts:
                matched = True
                for part in query_parts:
                    if part not in lowered:
                        matched = False
                        break
                if not matched:
                    continue

            self._schedules.append(schedule_view)
            self.cmbSchedule.Items.Add(name)
            self._schedule_lookup[name] = schedule_view

        if self.cmbSchedule.Items.Count > 0:
            if selected_name and selected_name in self._schedule_lookup:
                self.cmbSchedule.SelectedItem = selected_name
            elif self.cmbSchedule.SelectedIndex < 0:
                self.cmbSchedule.SelectedIndex = 0

        self._refresh_schedule_summary()

    def _load_schedules(self):
        current = _safe_str(self.cmbSchedule.SelectedItem)
        self._all_schedules = _get_supported_schedules()
        self._apply_schedule_filter(preserve_selection_name=current)

    def _seed_current_schedule(self):
        try:
            active_view = doc.ActiveView
            if isinstance(active_view, ViewSchedule) and _schedule_is_supported(active_view):
                active_name = _safe_str(active_view.Name)
                if active_name in self._schedule_lookup:
                    self.cmbSchedule.SelectedItem = active_name
                    self._refresh_schedule_summary()
        except Exception:
            pass

    def _load_last_import_path(self):
        last_path = _safe_str(getattr(config, "schedule_csv_roundtrip_last_import", ""))
        if last_path and os.path.exists(last_path):
            self.txtImportPath.Text = last_path

    def _save_last_import_path(self, path):
        try:
            config.schedule_csv_roundtrip_last_import = _safe_str(path)
            script.save_config()
        except Exception:
            pass

    def _refresh_schedule_summary(self):
        schedule_view = self._selected_schedule()
        self.lstFieldSummary.Items.Clear()
        if schedule_view is None:
            search_text = self._schedule_search_text()
            if search_text:
                self.txtScheduleSummary.Text = "No schedule matched your search."
                self._set_status("No schedules matched filter: '{0}'".format(search_text))
            else:
                self.txtScheduleSummary.Text = "No supported schedule selected."
                self._set_status("No supported schedule is currently selected.")
            return

        try:
            cache = RuntimeCache()
            elements, body_rows, field_infos, importable_count = _get_schedule_export_definition(
                schedule_view,
                bool(self.chkExportWritableOnly.IsChecked),
                cache,
            )
            try:
                is_itemized = bool(schedule_view.Definition.IsItemized)
            except Exception:
                is_itemized = False

            hidden_count = 0
            for item in field_infos:
                if item.get("is_hidden"):
                    hidden_count += 1

            self.txtScheduleSummary.Text = (
                "Name: {0}\n"
                "Visible schedule rows: {1}\n"
                "Elements in schedule scope: {2}\n"
                "Exported columns: {3}\n"
                "Writable import columns: {4}\n"
                "Hidden columns included: {5}\n"
                "Itemized: {6}"
            ).format(
                _safe_str(schedule_view.Name),
                len(body_rows),
                len(elements),
                len(field_infos),
                importable_count,
                hidden_count,
                "Yes" if is_itemized else "No",
            )

            for field_info in field_infos[:250]:
                heading = field_info.get("heading", "")
                parameter_name = field_info.get("parameter_name") or "<display-only>"
                if field_info.get("importable"):
                    self.lstFieldSummary.Items.Add("{0} -> {1}".format(heading, parameter_name))
                else:
                    self.lstFieldSummary.Items.Add("{0} -> <display-only>".format(heading))

            if not is_itemized:
                self._set_status("Selected schedule is not itemized. Export mirrors displayed rows; metadata mapping can be partial on grouped rows.")
            else:
                self._set_status("Ready to export all visible schedule columns and displayed values.")
        except Exception as ex:
            self.txtScheduleSummary.Text = "Unable to inspect schedule.\n\n{0}".format(_safe_str(ex))
            self._set_status("Schedule inspection failed.")

    def schedule_changed(self, sender, e):
        if self._is_busy:
            return
        self._refresh_schedule_summary()

    def schedule_search_changed(self, sender, e):
        if self._is_busy:
            return
        self._apply_schedule_filter(preserve_selection_name=self.cmbSchedule.SelectedItem)

    def export_options_changed(self, sender, e):
        if self._is_busy:
            return
        self._refresh_schedule_summary()

    def export_format_changed(self, sender, e):
        self._refresh_export_button_text()

    def header_export_click(self, sender, e):
        self.export_csv_click(sender, e)

    def header_preview_click(self, sender, e):
        self.preview_import_click(sender, e)

    def header_import_click(self, sender, e):
        self.apply_import_click(sender, e)

    def refresh_schedules_click(self, sender, e):
        if self._is_busy:
            forms.alert("A run is already in progress.")
            return
        self._load_schedules()
        self._append_log("Reloaded supported schedules.")

    def export_csv_click(self, sender, e):
        if self._is_busy:
            forms.alert("A run is already in progress.")
            return

        schedule_view = self._selected_schedule()
        if schedule_view is None:
            forms.alert("Choose a supported schedule first.")
            return

        self._is_busy = True
        metrics = OperationMetrics("export")

        try:
            self._export_template = _load_export_template()
            cache = RuntimeCache()
            elements, body_rows, field_infos, importable_count = _get_schedule_export_definition(
                schedule_view,
                bool(self.chkExportWritableOnly.IsChecked),
                cache,
            )

            if not body_rows:
                forms.alert("No visible rows were found in the selected schedule body.", title="Schedule Manager")
                return

            if not field_infos:
                forms.alert("No visible columns were found in the selected schedule.", title="Schedule Manager")
                return

            if len(body_rows) >= MAX_EXPORT_ROWS_CONFIRM:
                estimate = max(0.5, 0.25 + (len(body_rows) * 0.004))
                if not _confirm(
                    "This export will write {0} rows. Estimated duration: {1:.1f} s ({2:.0f} ms). Continue?".format(
                        len(body_rows), estimate, estimate * 1000.0
                    ),
                    title="Large Export",
                ):
                    return

            dialog = SaveFileDialog()
            export_format = self._selected_export_format()
            if export_format == "excel":
                dialog.Title = "Save Schedule Excel"
                dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*"
                dialog.FileName = _build_default_export_filename(schedule_view.Name).replace(".csv", ".xlsx")
            else:
                dialog.Title = "Save Schedule CSV"
                dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
                dialog.FileName = _build_default_export_filename(schedule_view.Name)

            if dialog.ShowDialog() != DialogResult.OK:
                return

            export_path = dialog.FileName
            if export_format == "excel" and not export_path.lower().endswith(".xlsx"):
                export_path += ".xlsx"
            if export_format == "csv" and not export_path.lower().endswith(".csv"):
                export_path += ".csv"

            exported_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            headers = _compose_export_headers(field_infos, self._export_template)
            total_rows = len(body_rows)
            cancelled = False
            progress_state = {"count": 0}
            parameter_target_mode = self._selected_parameter_target_mode()

            def row_generator(pb):
                count = 0
                for row in _yield_export_row_dicts(schedule_view, elements, body_rows, field_infos, exported_at, cache, parameter_target_mode):
                    count += 1
                    if pb.cancelled:
                        break
                    if count == 1 or (count % 50 == 0) or (count == total_rows):
                        _try_update_progress(pb, count, total_rows)
                        _pump_ui()
                    progress_state["count"] = count
                    yield row
                return

            with forms.ProgressBar(title="Export {value}/{max_value}", cancellable=True) as pb_export:
                if export_format == "excel":
                    _write_excel_stream(export_path, headers, row_generator(pb_export), self._export_template)
                else:
                    _write_csv_stream(export_path, headers, row_generator(pb_export))
                cancelled = bool(pb_export.cancelled)

            if cancelled:
                self._append_log("Export cancelled by user. Partial file may exist: {0}".format(export_path))
                self._set_status("Export cancelled.")
                forms.alert("Export cancelled by user. Partial file may exist:\n{0}".format(export_path), title="Schedule Manager")
                return

            metrics.stop()
            self._append_log(
                "Exported {0} rows and {1} visible columns from '{2}' to {3} ({4}, target mode: {5}).".format(
                    total_rows,
                    len(field_infos),
                    _safe_str(schedule_view.Name),
                    export_path,
                    export_format.upper(),
                    parameter_target_mode,
                )
            )
            self._append_log("Columns mapped as writable for import: {0}".format(importable_count))
            self._append_log(_format_metrics_line(metrics, progress_state["count"], progress_state["count"] * len(field_infos)))
            self._set_status("Export completed.")
            forms.alert("Exported file:\n{0}".format(export_path), title="Schedule Manager")
        except Exception as ex:
            logger.warning(traceback.format_exc())
            forms.alert("Export failed.\n\n{0}".format(_safe_str(ex)), title="Schedule Manager")
            self._append_log("Export failed: {0}".format(_safe_str(ex)))
            self._set_status("Export failed.")
        finally:
            self._is_busy = False

    def browse_import_click(self, sender, e):
        if self._is_busy:
            forms.alert("A run is already in progress.")
            return

        return self._prompt_import_path()

    def _prompt_import_path(self):
        current_path = _safe_str(getattr(self.txtImportPath, "Text", "")).strip()
        if current_path and os.path.exists(current_path):
            return current_path

        dialog = OpenFileDialog()
        dialog.Title = "Select Schedule CSV"
        dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if dialog.ShowDialog() == DialogResult.OK:
            self.txtImportPath.Text = dialog.FileName
            self._save_last_import_path(dialog.FileName)
            self._append_log("Selected import file: {0}".format(dialog.FileName))
            self._last_preview = None
            self._last_preview_signature = None
            return dialog.FileName

        return ""

    def _preview_signature(self, path):
        target_mode = self._selected_parameter_target_mode()
        try:
            stat = os.stat(path)
            return (
                path,
                float(stat.st_mtime),
                int(stat.st_size),
                bool(self.chkSkipBlankIncoming.IsChecked),
                target_mode,
            )
        except Exception:
            return (path, bool(self.chkSkipBlankIncoming.IsChecked), target_mode)

    def _run_preview(self):
        import_path = _safe_str(getattr(self.txtImportPath, "Text", "")).strip()
        if not import_path:
            import_path = _prompt_path = self._prompt_import_path()
            if not _safe_str(_prompt_path).strip():
                raise RuntimeError("Import cancelled. No CSV file was selected.")
        if not os.path.exists(import_path):
            raise RuntimeError("The selected CSV file does not exist.")

        signature = self._preview_signature(import_path)
        if self._last_preview is not None and signature == self._last_preview_signature:
            preview = self._last_preview
        else:
            preview = self._preview_import_stream(import_path, bool(self.chkSkipBlankIncoming.IsChecked))
            self._last_preview = preview
            self._last_preview_signature = signature

        self._save_last_import_path(import_path)

        self.txtImportSummary.Text = (
            "CSV rows: {0}\n"
            "Detected schedule: {1}\n"
            "Schedule found in model: {2}\n"
            "Target mode: {3}\n"
            "Mapped import columns: {4}\n"
            "Changed writable cells: {5}\n"
            "No-change cells: {6}\n"
            "Read-only cells: {7}\n"
            "Missing-parameter cells: {8}\n"
            "Blank values skipped: {9}\n"
            "Rows with missing elements: {10}\n"
            "Unsupported columns: {11}"
        ).format(
            preview.total_rows,
            preview.schedule_name or "<none>",
            "Yes" if preview.schedule_view is not None else "No",
            self._selected_parameter_target_mode(),
            len(preview.mapped_columns),
            preview.changed_cells,
            preview.no_change_cells,
            preview.read_only_cells,
            preview.missing_parameter_cells,
            preview.blank_skipped,
            preview.missing_elements,
            len(preview.unsupported_columns),
        )

        self.lvwImportPreview.Items.Clear()
        for row in preview.preview_rows:
            self.lvwImportPreview.Items.Add(row)

        self._append_log(
            "Previewed import. Rows: {0}, pending writes: {1}, missing elements: {2}, unsupported columns: {3}.".format(
                preview.total_rows,
                preview.changed_cells,
                preview.missing_elements,
                len(preview.unsupported_columns),
            )
        )

        if preview.cancelled:
            self._set_status("Import preview cancelled.")
        else:
            self._set_status("Import preview ready.")

        return preview

    def _preview_import_stream(self, path, skip_blank_values):
        metrics = OperationMetrics("preview")
        preview = PreviewResult()
        preview.total_rows = _count_csv_data_rows(path)
        cache = RuntimeCache()
        parameter_target_mode = self._selected_parameter_target_mode()

        if preview.total_rows <= 0:
            raise RuntimeError("The selected CSV is empty.")

        with forms.ProgressBar(title="Preview {value}/{max_value}", cancellable=True) as pb_preview:
            mapped_columns = None
            row_index = 0
            headers = []

            for headers_or_none, row in _iter_csv_dict_rows(path):
                if headers_or_none is not None:
                    headers = headers_or_none
                    preview.headers = headers
                    if META_ELEMENT_ID not in headers and META_UNIQUE_ID not in headers:
                        raise RuntimeError("The CSV is missing the required metadata columns for safe reimport.")
                    continue

                row_index += 1
                if pb_preview.cancelled:
                    preview.cancelled = True
                    break

                if mapped_columns is None:
                    preview.schedule_name = _safe_str(row.get(META_SCHEDULE_NAME, "")).strip()
                    preview.schedule_view = _find_schedule_by_name(preview.schedule_name) if preview.schedule_name else None
                    mapped_columns, unsupported_columns = _build_import_column_map(preview.schedule_view, headers)
                    preview.mapped_columns = mapped_columns
                    preview.unsupported_columns = unsupported_columns

                    if unsupported_columns and len(preview.preview_rows) < MAX_PREVIEW_TABLE_ROWS:
                        preview.preview_rows.append(
                            ImportPreviewRow("-", "-", ", ".join(unsupported_columns[:3]), "", "", "Unsupported columns skipped")
                        )

                element = cache.get_element_from_row(row)
                if element is None:
                    preview.missing_elements += 1
                    if len(preview.preview_rows) < MAX_PREVIEW_TABLE_ROWS:
                        preview.preview_rows.append(ImportPreviewRow(row_index, "", "", "", "", "Missing element"))
                    if len(preview.summary_lines) < PREVIEW_LOG_SAMPLE_LIMIT:
                        preview.summary_lines.append("Row {0}: element not found in current model.".format(row_index))

                    _try_update_progress(pb_preview, row_index, preview.total_rows)
                    if row_index % 50 == 0:
                        _pump_ui()
                    continue

                for mapping in preview.mapped_columns:
                    preview.total_rows = max(preview.total_rows, row_index)
                    incoming = _safe_str(row.get(mapping["column"], ""))
                    if incoming == "" and skip_blank_values:
                        preview.blank_skipped += 1
                        continue

                    parameter = cache.get_parameter(element, mapping["parameter_name"], parameter_target_mode)
                    if parameter is None:
                        preview.missing_parameter_cells += 1
                        continue

                    try:
                        if parameter.IsReadOnly:
                            preview.read_only_cells += 1
                            continue
                        if parameter.StorageType not in (StorageType.String, StorageType.Integer, StorageType.Double):
                            preview.unsupported_type_cells += 1
                            continue
                    except Exception:
                        preview.unsupported_type_cells += 1
                        continue

                    current_display = cache.get_display_value(element, mapping["parameter_name"], parameter_target_mode)
                    if _parameter_value_matches_incoming(parameter, incoming):
                        preview.no_change_cells += 1
                        continue

                    preview.changed_cells += 1
                    if len(preview.preview_rows) < MAX_PREVIEW_TABLE_ROWS:
                        preview.preview_rows.append(
                            ImportPreviewRow(
                                row_index,
                                int(element.Id.IntegerValue),
                                mapping["column"],
                                current_display,
                                incoming,
                                "Will write",
                            )
                        )

                    if len(preview.summary_lines) < PREVIEW_LOG_SAMPLE_LIMIT:
                        preview.summary_lines.append(
                            "Row {0} | Element {1} | {2}: '{3}' -> '{4}'".format(
                                row_index,
                                int(element.Id.IntegerValue),
                                mapping["column"],
                                current_display,
                                incoming,
                            )
                        )

                _try_update_progress(pb_preview, row_index, preview.total_rows)
                if row_index % 50 == 0:
                    _pump_ui()

        if not preview.summary_lines:
            preview.summary_lines.append("No writable changes were detected in the CSV.")
        if preview.unsupported_columns:
            preview.summary_lines.append(
                "Unsupported columns skipped: {0}".format(", ".join(sorted(set(preview.unsupported_columns))[:8]))
            )
        if preview.changed_cells > len(preview.preview_rows):
            preview.summary_lines.append(
                "Preview table shows first {0} rows for responsiveness.".format(len(preview.preview_rows))
            )

        metrics.stop()
        self._append_log(_format_metrics_line(metrics, row_index, preview.changed_cells))
        return preview

    def preview_import_click(self, sender, e):
        if self._is_busy:
            forms.alert("A run is already in progress.")
            return
        self._is_busy = True
        try:
            self._run_preview()
        except Exception as ex:
            forms.alert("Preview failed.\n\n{0}".format(_safe_str(ex)), title="Schedule Manager")
            self._append_log("Preview failed: {0}".format(_safe_str(ex)))
            self._set_status("Import preview failed.")
        finally:
            self._is_busy = False

    def apply_import_click(self, sender, e):
        if self._is_busy:
            forms.alert("A run is already in progress.")
            return

        self._is_busy = True
        metrics = OperationMetrics("import")
        result = ImportResult()

        try:
            preview = self._run_preview()
            if preview.changed_cells <= 0:
                forms.alert("No writable changes were found in the CSV.", title="Schedule Manager")
                return

            result.total_rows = preview.total_rows
            result.unsupported_columns = len(preview.unsupported_columns)

            estimate = _estimate_import_seconds(preview.total_rows, preview.changed_cells)
            parameter_target_mode = self._selected_parameter_target_mode()
            confirm_text = (
                "This import will evaluate {0} rows with {1} mapped columns.\n"
                "Target mode: {2}\n"
                "Potential writes (preview): {3}\n\n"
                "Estimated duration: {4:.1f} s ({5:.0f} ms).\n"
                "Batch size: {6}\n"
                "Stop on first verification failure: {7}\n\n"
                "Continue?"
            ).format(
                preview.total_rows,
                len(preview.mapped_columns),
                parameter_target_mode,
                preview.changed_cells,
                estimate,
                estimate * 1000.0,
                IMPORT_BATCH_SIZE,
                "Yes" if bool(self.chkStopOnFailure.IsChecked) else "No",
            )

            large_run = preview.total_rows >= MAX_IMPORT_ROWS_CONFIRM
            title = "Large Import" if large_run else "Confirm Import"
            if not _confirm(confirm_text, title=title):
                return

            import_path = _safe_str(self.txtImportPath.Text).strip()
            if not import_path or not os.path.exists(import_path):
                raise RuntimeError("The selected CSV file does not exist.")

            stop_on_failure = bool(self.chkStopOnFailure.IsChecked)
            cache = RuntimeCache()
            mapped_columns = list(preview.mapped_columns)
            tg = TransactionGroup(doc, "Schedule Manager Import")
            tg.Start()

            active_tx = None
            active_batch_writes = 0
            stop_processing = False

            try:
                with forms.ProgressBar(title="Import {value}/{max_value}", cancellable=True) as pb_import:
                    row_index = 0
                    headers_seen = None

                    for headers_or_none, row in _iter_csv_dict_rows(import_path):
                        if headers_or_none is not None:
                            headers_seen = headers_or_none
                            if META_ELEMENT_ID not in headers_seen and META_UNIQUE_ID not in headers_seen:
                                raise RuntimeError("The CSV is missing required metadata columns for safe import.")
                            continue

                        row_index += 1
                        result.processed_rows = row_index

                        if pb_import.cancelled:
                            result.cancelled = True
                            stop_processing = True
                            break

                        element = cache.get_element_from_row(row)
                        if element is None:
                            result.missing_elements += 1
                            _try_update_progress(pb_import, row_index, result.total_rows)
                            if row_index % 40 == 0:
                                _pump_ui()
                            continue

                        for mapping in mapped_columns:
                            result.total_cells_seen += 1
                            incoming = _safe_str(row.get(mapping["column"], ""))

                            if incoming == "" and bool(self.chkSkipBlankIncoming.IsChecked):
                                result.blank_skipped += 1
                                continue

                            parameter = cache.get_parameter(element, mapping["parameter_name"], parameter_target_mode)
                            if parameter is None:
                                result.missing_parameter_cells += 1
                                continue

                            try:
                                if parameter.IsReadOnly:
                                    result.read_only_cells += 1
                                    continue
                                if parameter.StorageType not in (StorageType.String, StorageType.Integer, StorageType.Double):
                                    result.unsupported_type_cells += 1
                                    continue
                            except Exception:
                                result.unsupported_type_cells += 1
                                continue

                            if _parameter_value_matches_incoming(parameter, incoming):
                                result.no_change_cells += 1
                                continue

                            if active_tx is None:
                                active_tx = Transaction(doc, "Schedule Manager Import Batch")
                                active_tx.Start()
                                active_batch_writes = 0

                            success, reason = _set_parameter_value(parameter, incoming)
                            if success:
                                result.written_cells += 1
                                active_batch_writes += 1
                                cache.invalidate_parameter_display(element, mapping["parameter_name"], parameter_target_mode)
                            else:
                                if reason == "no change":
                                    result.no_change_cells += 1
                                elif reason == "read-only parameter":
                                    result.read_only_cells += 1
                                elif reason == "missing parameter":
                                    result.missing_parameter_cells += 1
                                elif reason == "write verification failed":
                                    result.verification_failures += 1
                                    result.failed_cells += 1
                                else:
                                    result.failed_cells += 1

                                if reason != "no change" and len(result.sample_failures) < 12:
                                    result.sample_failures.append(
                                        "Row {0} | Element {1} | {2}: {3}".format(
                                            row_index,
                                            int(element.Id.IntegerValue),
                                            mapping["column"],
                                            reason,
                                        )
                                    )

                                if stop_on_failure and reason not in ("no change", "read-only parameter", "missing parameter"):
                                    result.stopped_on_failure = True
                                    stop_processing = True
                                    break

                            if active_batch_writes >= IMPORT_BATCH_SIZE:
                                try:
                                    active_tx.Commit()
                                finally:
                                    active_tx = None
                                    active_batch_writes = 0

                        _try_update_progress(pb_import, row_index, result.total_rows)
                        if row_index % 40 == 0:
                            _pump_ui()

                        if stop_processing:
                            break

                    if active_tx is not None:
                        if result.stopped_on_failure:
                            active_tx.RollBack()
                        else:
                            active_tx.Commit()
                        active_tx = None

                tg.Assimilate()
            except Exception:
                try:
                    if active_tx is not None:
                        active_tx.RollBack()
                except Exception:
                    pass
                try:
                    tg.RollBack()
                except Exception:
                    pass
                raise

            metrics.stop()

            summary_lines = [
                "Import completed." if not result.cancelled else "Import cancelled by user.",
                "Target mode: {0}".format(parameter_target_mode),
                "Successful writes: {0}".format(result.written_cells),
                "No changes: {0}".format(result.no_change_cells),
                "Read-only parameters: {0}".format(result.read_only_cells),
                "Missing parameters: {0}".format(result.missing_parameter_cells),
                "Missing elements: {0}".format(result.missing_elements),
                "Unsupported columns: {0}".format(result.unsupported_columns),
                "Unsupported type cells: {0}".format(result.unsupported_type_cells),
                "Blank values skipped: {0}".format(result.blank_skipped),
                "Verification failures: {0}".format(result.verification_failures),
                "Failed cells: {0}".format(result.failed_cells),
                "Rows processed: {0}/{1}".format(result.processed_rows, result.total_rows),
                _format_metrics_line(metrics, result.processed_rows, result.written_cells),
            ]

            if result.stopped_on_failure:
                summary_lines.append("Stopped on first verification failure as requested. Completed batches were retained.")

            if result.sample_failures:
                summary_lines.append("")
                summary_lines.append("Sample failures:")
                summary_lines.extend(result.sample_failures[:10])

            self._append_log(" | ".join(summary_lines[:5]))
            self._append_log(_format_metrics_line(metrics, result.processed_rows, result.written_cells))
            if result.sample_failures:
                self._append_log("Sample failure: {0}".format(result.sample_failures[0]))

            if result.cancelled:
                self._set_status("Import cancelled. Completed batches were retained.")
            elif result.stopped_on_failure:
                self._set_status("Import stopped on first failure. Completed batches were retained.")
            else:
                self._set_status("Import completed.")

            forms.alert("\n".join(summary_lines), title="Schedule Manager")
        except Exception as ex:
            logger.warning(traceback.format_exc())
            forms.alert("Import failed.\n\n{0}".format(_safe_str(ex)), title="Schedule Manager")
            self._append_log("Import failed: {0}".format(_safe_str(ex)))
            self._set_status("Import failed.")
        finally:
            self._is_busy = False


def _launch_window():
    xaml_file = script.get_bundle_file("WPFWindow.xaml")
    try:
        window = ScheduleCsvRoundtripWindow(xaml_file)
        window.ShowDialog()
    except Exception as ex:
        logger.warning(traceback.format_exc())
        forms.alert("Unable to launch Schedule Manager.\n\n{0}".format(_safe_str(ex)))


_launch_window()
