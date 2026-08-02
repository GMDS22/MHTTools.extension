# coding: utf8
from __future__ import print_function

import os
import shutil
import time

from System import Activator, Type
from System.Runtime.InteropServices import Marshal

from pyrevit import forms
from pyrevit.forms import WPFWindow


__title__ = "NWB Data Validator"

TOOL_TITLE = "NWB Data Validator"
AUTOFILL_FOLDER = "NWB_PARAMETERS AutoFill.pushbutton"
TEMPLATE_FOLDER_NAME = "templates"
CSV_TEMPLATE_NAME = "NWB_Data_Validator_Template.csv"
XLSX_TEMPLATE_NAME = "NWB_Data_Validator_Template.xlsx"
XLSX_FILE_FORMAT = 51
REPORT_COLUMNS = [
    "run_stamp",
    "workbook_path",
    "view_name",
    "scope",
    "element_id",
    "unique_id",
    "category",
    "family",
    "type",
    "level",
    "system",
    "asset_score",
    "asset_match_status",
    "asset_match_reason",
    "asset_type_code",
    "asset_type",
    "asset_category",
    "target_parameter",
    "expected_value",
    "actual_value",
    "validation_status",
    "reason",
    "target_policy",
    "candidate_count",
    "selected_target_count",
    "target_details",
    "linked_room_source",
    "linked_room_number",
    "linked_room_department",
    "derived_wbs_code",
    "suggested_action",
]
OUTPUT_FORMAT_CHOICES = [
    ("CSV report (*.csv)", "csv"),
    ("Excel workbook (*.xlsx)", "xlsx"),
]


def _safe_str(value):
    try:
        return str(value)
    except Exception:
        return ""


def _tool_folder():
    return os.path.dirname(__file__)


def _templates_folder(tool_folder):
    path = os.path.join(tool_folder, TEMPLATE_FOLDER_NAME)
    if not os.path.isdir(path):
        os.makedirs(path)
    return path


def _csv_template_path(tool_folder):
    return os.path.join(_templates_folder(tool_folder), CSV_TEMPLATE_NAME)


def _xlsx_template_path(tool_folder):
    return os.path.join(_templates_folder(tool_folder), XLSX_TEMPLATE_NAME)


def _autofill_script_path():
    return os.path.normpath(
        os.path.join(_tool_folder(), "..", AUTOFILL_FOLDER, "script.py")
    )


def _load_mapping_module():
    path = _autofill_script_path()
    if not os.path.isfile(path):
        raise RuntimeError("NWB_PARAMETERS AutoFill script was not found: {0}".format(path))

    namespace = {
        "__name__": "nwb_parameter_mapping",
        "__file__": path,
    }
    with open(path, "r") as source_file:
        source = source_file.read()
    exec(compile(source, path, "exec"), namespace)
    return namespace


def _csv_escape(value):
    text = _safe_str(value).replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    return '"{0}"'.format(text.replace('"', '""'))


def _open_export_file(path):
    try:
        if path and os.path.isfile(path):
            os.startfile(path)
            return True
    except Exception:
        pass
    return False


def _parse_template_columns(header_line):
    if not header_line:
        return []
    columns = []
    for token in header_line.split(","):
        value = _safe_str(token).strip().strip('"').strip()
        if value:
            columns.append(value)
    return columns


def _write_default_csv_template(path):
    stream = open(path, "w")
    try:
        stream.write(",".join([_csv_escape(column) for column in REPORT_COLUMNS]))
        stream.write("\n")
    finally:
        stream.close()


def _read_csv_template_columns(path):
    if not os.path.isfile(path):
        return []
    stream = open(path, "r")
    try:
        header_line = stream.readline()
    finally:
        stream.close()
    return _parse_template_columns(header_line)


def _create_excel_app():
    excel_type = Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise RuntimeError("Excel is not available on this machine.")
    app = Activator.CreateInstance(excel_type)
    app.Visible = False
    app.DisplayAlerts = False
    return app


def _close_excel(app):
    if app is None:
        return
    try:
        app.Quit()
    except Exception:
        pass
    try:
        Marshal.ReleaseComObject(app)
    except Exception:
        pass


def _ensure_default_xlsx_template(path):
    if os.path.isfile(path):
        return
    app = None
    workbook = None
    try:
        app = _create_excel_app()
        workbook = app.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        for index, column in enumerate(REPORT_COLUMNS, 1):
            sheet.Cells(1, index).Value2 = column
        workbook.SaveAs(path, XLSX_FILE_FORMAT)
    finally:
        if workbook is not None:
            try:
                workbook.Close(True)
            except Exception:
                pass
        _close_excel(app)


def _read_xlsx_template_columns(path):
    app = None
    workbook = None
    columns = []
    try:
        app = _create_excel_app()
        workbook = app.Workbooks.Open(path)
        sheet = workbook.Worksheets(1)
        for index in range(1, 512):
            value = _safe_str(sheet.Cells(1, index).Value2).strip()
            if not value:
                break
            columns.append(value)
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        _close_excel(app)
    return columns


class CsvReport(object):
    def __init__(self, tool_folder, stamp):
        logs_dir = os.path.join(tool_folder, "logs")
        if not os.path.isdir(logs_dir):
            os.makedirs(logs_dir)
        self.path = os.path.join(logs_dir, "NWB_Data_Validator_{0}.csv".format(stamp))
        self._stream = open(self.path, "w")
        self._stream.write(",".join([_csv_escape(column) for column in REPORT_COLUMNS]))
        self._stream.write("\n")

    def write_row(self, row):
        self._stream.write(",".join([_csv_escape((row or {}).get(column, "")) for column in REPORT_COLUMNS]))
        self._stream.write("\n")

    def close(self):
        try:
            self._stream.close()
        except Exception:
            pass


class CsvTemplateReport(object):
    def __init__(self, tool_folder, stamp):
        logs_dir = os.path.join(tool_folder, "logs")
        if not os.path.isdir(logs_dir):
            os.makedirs(logs_dir)

        template_path = _csv_template_path(tool_folder)
        if not os.path.isfile(template_path):
            _write_default_csv_template(template_path)
        columns = _read_csv_template_columns(template_path)
        if not columns:
            columns = list(REPORT_COLUMNS)

        self.columns = columns
        self.template_path = template_path
        self.path = os.path.join(logs_dir, "NWB_Data_Validator_{0}.csv".format(stamp))
        self._stream = open(self.path, "w")
        self._stream.write(",".join([_csv_escape(column) for column in self.columns]))
        self._stream.write("\n")

    def write_row(self, row):
        self._stream.write(",".join([_csv_escape((row or {}).get(column, "")) for column in self.columns]))
        self._stream.write("\n")

    def close(self):
        try:
            self._stream.close()
        except Exception:
            pass


class XlsxTemplateReport(object):
    def __init__(self, tool_folder, stamp):
        logs_dir = os.path.join(tool_folder, "logs")
        if not os.path.isdir(logs_dir):
            os.makedirs(logs_dir)

        self.template_path = _xlsx_template_path(tool_folder)
        _ensure_default_xlsx_template(self.template_path)

        columns = _read_xlsx_template_columns(self.template_path)
        if not columns:
            columns = list(REPORT_COLUMNS)
        self.columns = columns
        self.path = os.path.join(logs_dir, "NWB_Data_Validator_{0}.xlsx".format(stamp))
        self._rows = []

    def write_row(self, row):
        self._rows.append(dict(row or {}))

    def close(self):
        shutil.copyfile(self.template_path, self.path)

        app = None
        workbook = None
        try:
            app = _create_excel_app()
            workbook = app.Workbooks.Open(self.path)
            sheet = workbook.Worksheets(1)

            for row_index, row in enumerate(self._rows, 2):
                for col_index, column in enumerate(self.columns, 1):
                    sheet.Cells(row_index, col_index).Value2 = _safe_str(row.get(column, ""))

            workbook.Save()
        finally:
            if workbook is not None:
                try:
                    workbook.Close(True)
                except Exception:
                    pass
            _close_excel(app)


def _create_report_writer(tool_folder, stamp, output_format):
    notices = []
    if output_format == "xlsx":
        try:
            return XlsxTemplateReport(tool_folder, stamp), notices
        except Exception as ex:
            notices.append("Excel output was unavailable; exported CSV instead: {0}".format(_safe_str(ex)))
    return CsvTemplateReport(tool_folder, stamp), notices


class ValidatorOptionsWindow(WPFWindow):
    def __init__(self, selected_count, active_view_name, linked_room_choices):
        WPFWindow.__init__(self, "ValidatorOptionsWindow.xaml")
        self.result = None
        self._scope_choices = [("Active view - all eligible MEP elements", "view")]
        if selected_count:
            self._scope_choices.insert(0, ("Current selection - eligible MEP elements only", "selection"))
        self._target_choices = [
            ("Shared parameters only (recommended)", "SharedOnly"),
            ("Both shared and family/non-shared duplicates", "Both"),
            ("Family/non-shared parameters only", "FamilyOnly"),
            ("First available parameter (diagnostic fallback)", "FirstMatch"),
        ]
        self._output_choices = list(OUTPUT_FORMAT_CHOICES)
        self._linked_room_choices = list(linked_room_choices or [("Auto-detect from available architectural room links", "")])

        self.txtActiveView.Text = _safe_str(active_view_name)
        self.txtSelectedCount.Text = str(max(0, int(selected_count)))
        self._bind_combo(self.cmbProcessingScope, self._scope_choices, "selection" if selected_count else "view")
        self._bind_combo(self.cmbParameterTargetMode, self._target_choices, "SharedOnly")
        self._bind_combo(self.cmbOutputFormat, self._output_choices, "csv")
        self._bind_combo(self.cmbArchitecturalRoomLink, self._linked_room_choices, "")

        if len(self._linked_room_choices) <= 1:
            self.cmbArchitecturalRoomLink.IsEnabled = False

    def _bind_combo(self, combo, choices, selected_value):
        combo.Items.Clear()
        selected_index = 0
        for index, item in enumerate(choices):
            combo.Items.Add(item[0])
            if item[1] == selected_value:
                selected_index = index
        if combo.Items.Count:
            combo.SelectedIndex = selected_index

    def _selected_value(self, combo, choices):
        try:
            index = int(combo.SelectedIndex)
        except Exception:
            index = 0
        if index < 0 or index >= len(choices):
            index = 0
        return choices[index][1]

    def start_click(self, sender, e):
        self.result = {
            "scope": self._selected_value(self.cmbProcessingScope, self._scope_choices),
            "policy": self._selected_value(self.cmbParameterTargetMode, self._target_choices),
            "output_format": self._selected_value(self.cmbOutputFormat, self._output_choices),
            "arch_room_link_id": self._selected_value(self.cmbArchitecturalRoomLink, self._linked_room_choices),
        }
        self.Close()

    def cancel_click(self, sender, e):
        self.result = None
        self.Close()


def _show_options(selected_count, active_view_name, linked_room_choices):
    try:
        window = ValidatorOptionsWindow(selected_count, active_view_name, linked_room_choices)
        window.ShowDialog()
        return window.result
    except Exception as ex:
        forms.alert(
            "Could not open the validator setup window.\n\n{0}".format(_safe_str(ex)),
            title=TOOL_TITLE,
        )
        return None


def _selected_candidates(mapping, candidates, policy):
    if policy == "Both":
        return list(candidates)

    shared = [candidate for candidate in candidates if candidate.get("is_shared")]
    nonshared = [candidate for candidate in candidates if not candidate.get("is_shared")]
    if policy == "FamilyOnly":
        return nonshared
    if policy == "FirstMatch":
        preferred = shared[0] if shared else (candidates[0] if candidates else None)
        return [preferred] if preferred is not None else []
    return shared


def _candidate_detail(mapping, candidate):
    param = candidate.get("param")
    actual = mapping["_parameter_text"](param)
    return "{0} | {1} | {2} | ReadOnly={3} | Actual='{4}'".format(
        candidate.get("scope", ""),
        mapping["_parameter_kind_label"](candidate.get("is_shared", False)),
        candidate.get("id", ""),
        _safe_str(getattr(param, "IsReadOnly", "")),
        actual,
    )


def _expected_blank_reason(parameter_name, audit_context):
    if parameter_name == "NWB_Material":
        return "Workbook material is blank and the model has no material source."
    if parameter_name == "NWB_Department":
        source = _safe_str((audit_context or {}).get("linked_room_source", ""))
        return "Linked/direct room department unavailable{0}.".format(
            " ({0})".format(source) if source else ""
        )
    return "No expected value can be derived from the workbook and model context."


def _suggested_action(status, parameter_name, details):
    if status == "MISSING_PARAMETER":
        return "Bind {0} as a writable shared parameter to this category/family.".format(parameter_name)
    if status == "READ_ONLY_MISMATCH":
        return "Make the {0} binding writable, then run NWB_PARAMETERS AutoFill.".format(parameter_name)
    if status == "MISMATCH":
        return "Review the workbook match, then run NWB_PARAMETERS AutoFill if the expected value is correct."
    if status == "AMBIGUOUS_DUPLICATES":
        return "Resolve duplicate parameter bindings with conflicting stored values."
    if status == "NO_ASSET_MATCH":
        return "Add or refine an Asset Master row for this element type/category."
    if status == "NO_EXPECTED_VALUE":
        if parameter_name == "NWB_Material":
            return "Populate Asset Master NWB_Material or the element's Material source."
        if parameter_name == "NWB_Department":
            return "Place/connect the element to an architectural room or provide a direct department source."
        return "Provide the missing workbook or model source data."
    return ""


def _validate_parameter(mapping, element, parameter_name, expected_value, policy, audit_context):
    candidates = mapping["_collect_param_candidates"](element, parameter_name)
    selected = _selected_candidates(mapping, candidates, policy)
    details = "; ".join([_candidate_detail(mapping, candidate) for candidate in selected])
    actual_values = [mapping["_parameter_text"](candidate.get("param")) for candidate in selected]
    actual = " | ".join(actual_values)

    if not candidates:
        status = "MISSING_PARAMETER"
        reason = "No parameter candidate exists on the instance or its type."
    elif not selected:
        status = "MISSING_PARAMETER"
        reason = "Candidates exist but none meet the {0} validation policy.".format(policy)
    elif mapping["_is_blank"](expected_value):
        status = "NO_EXPECTED_VALUE"
        reason = _expected_blank_reason(parameter_name, audit_context)
    else:
        matches = [
            mapping["_parameter_value_matches_incoming"](candidate.get("param"), expected_value)
            for candidate in selected
        ]
        if all(matches):
            status = "PASS"
            reason = "All selected parameter targets match the derived value."
        elif any(matches):
            status = "AMBIGUOUS_DUPLICATES"
            reason = "Duplicate selected parameter targets do not all store the same expected value."
        elif all(bool(getattr(candidate.get("param"), "IsReadOnly", False)) for candidate in selected):
            status = "READ_ONLY_MISMATCH"
            reason = "All selected parameter targets are read-only and differ from the derived value."
        else:
            status = "MISMATCH"
            reason = "Stored value differs from the workbook-derived expected value."

    return {
        "actual_value": actual,
        "validation_status": status,
        "reason": reason,
        "candidate_count": len(candidates),
        "selected_target_count": len(selected),
        "target_details": details,
        "suggested_action": _suggested_action(status, parameter_name, details),
    }


def _collect_elements(mapping, scope):
    doc = mapping["doc"]
    active_view = doc.ActiveView
    view_mode = mapping["_view_label"](active_view)
    if view_mode == "schedule":
        if not mapping["_schedule_is_supported"](active_view):
            raise RuntimeError("The active schedule is not supported. Open a model view or a non-key schedule.")
        return mapping["_collect_active_schedule_elements"](active_view), active_view
    if scope == "selection":
        return mapping["_get_selected_mep_elements"](), active_view
    return mapping["_collect_active_model_view_elements"](active_view), active_view


def run():
    mapping = _load_mapping_module()
    selected = mapping["_get_selected_mep_elements"]()
    active_view = mapping["doc"].ActiveView
    options = _show_options(
        len(selected),
        _safe_str(getattr(active_view, "Name", "")),
        mapping["_collect_linked_room_choices"](),
    )
    if options is None:
        return

    scope = options["scope"]
    policy = options["policy"]
    output_format = options.get("output_format", "csv")
    arch_room_link_id = options["arch_room_link_id"]

    mapping["_set_run_options"](
        {
            "arch_room_link_id": arch_room_link_id,
            "site_override": "",
            "building_override": "",
            "design_pkg_override": "",
            "building_permit_override": "",
        }
    )
    elements, active_view = _collect_elements(mapping, scope)
    if not elements:
        forms.alert("No eligible MEP elements were found for the selected validation scope.", title=TOOL_TITLE)
        return

    workbook_path = mapping["_build_workbook_path"]()
    if not mapping["_validate_workbook_path"](workbook_path):
        forms.alert("Workbook not found:\n{0}".format(workbook_path), title=TOOL_TITLE)
        return

    try:
        workbook_data = mapping["_load_workbook_data"](workbook_path)
    except Exception as ex:
        forms.alert("Could not read the NWB workbook.\n\n{0}".format(_safe_str(ex)), title=TOOL_TITLE)
        return

    stamp = time.strftime("%Y%m%d_%H%M%S")
    report, export_notices = _create_report_writer(_tool_folder(), stamp, output_format)
    counts = {}
    no_asset_match = 0
    cancelled = False
    run_context = {
        "run_stamp": stamp,
        "workbook_path": workbook_path,
        "view_name": _safe_str(getattr(active_view, "Name", "")),
        "scope": scope,
    }

    try:
        with forms.ProgressBar(title="NWB data validation {value}/{max_value}", cancellable=True) as progress:
            total = len(elements)
            for index, element in enumerate(elements, 1):
                if progress.cancelled:
                    cancelled = True
                    break

                expected_values, debug = mapping["_build_target_values"](element, workbook_data)
                audit_context = mapping["_build_param_audit_context"](run_context, element, debug, expected_values)
                asset_score = int((debug or {}).get("asset_score", 0))
                asset_match_status = "MATCHED" if asset_score > 0 else "NO_ASSET_MATCH"
                asset_match_reason = (
                    "Asset Master row selected by the current AutoFill matching rules."
                    if asset_score > 0
                    else "No Asset Master row matched this element; asset-specific expected values may be unavailable."
                )
                if asset_score <= 0:
                    no_asset_match += 1

                for parameter_name in mapping["TARGET_PARAMETERS"]:
                    result = _validate_parameter(
                        mapping,
                        element,
                        parameter_name,
                        expected_values.get(parameter_name, ""),
                        policy,
                        audit_context,
                    )
                    row = dict(audit_context)
                    row.update(
                        {
                            "unique_id": _safe_str(getattr(element, "UniqueId", "")),
                            "target_parameter": parameter_name,
                            "expected_value": expected_values.get(parameter_name, ""),
                            "target_policy": policy,
                            "asset_match_status": asset_match_status,
                            "asset_match_reason": asset_match_reason,
                        }
                    )
                    row.update(result)
                    report.write_row(row)
                    status = row["validation_status"]
                    counts[status] = counts.get(status, 0) + 1

                progress.update_progress(index, total)
    finally:
        report.close()

    opened_report = _open_export_file(report.path)

    summary = [
        "NWB Data Validator completed." if not cancelled else "NWB Data Validator cancelled; partial report exported.",
        "Elements checked: {0}/{1}".format(sum(counts.values()) // len(mapping["TARGET_PARAMETERS"]), len(elements)),
        "No Asset Master match: {0}".format(no_asset_match),
        "Report file: {0}".format(report.path),
        "Report opened: {0}".format("Yes" if opened_report else "No"),
        "",
    ]
    for notice in export_notices:
        summary.append(notice)
    if export_notices:
        summary.append("")
    for status in sorted(counts.keys()):
        summary.append("{0}: {1}".format(status, counts[status]))
    forms.alert("\n".join(summary), title=TOOL_TITLE)


if __name__ == "__main__":
    try:
        run()
    except Exception as ex:
        forms.alert("{0} failed.\n\n{1}".format(TOOL_TITLE, _safe_str(ex)), title=TOOL_TITLE)