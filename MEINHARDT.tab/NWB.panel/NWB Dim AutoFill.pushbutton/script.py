# coding: utf8
from __future__ import print_function

import os
import re
import time

import clr

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Application as WinFormsApplication
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FailureProcessingResult,
    FilteredElementCollector,
    IFailuresPreprocessor,
    StorageType,
    Transaction,
    TransactionStatus,
    ViewSchedule,
)
from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow


doc = revit.doc
uidoc = revit.uidoc
config = script.get_config()


TOOL_TITLE = "NWB Dim AutoFill"
FAST_MODE_MAX_ELEMENTS = 2000

TARGET_PARAMETERS = [
    "NWB_DimWidth",
    "NWB_DimHeight",
    "NWB_DimDiameter",
]

DEFAULT_PARAMETER_TARGET_MODE = "SharedOnly"

WRITE_MODE_CHOICES = [
    ("Keep existing NWB dim values that are already set", "skip"),
    ("Override existing NWB dim values", "override"),
]

PARAMETER_TARGET_MODE_CHOICES = [
    ("{0} (recommended)".format(DEFAULT_PARAMETER_TARGET_MODE), DEFAULT_PARAMETER_TARGET_MODE),
    ("Both (shared + legacy duplicates)", "Both"),
    ("FamilyOnly / non-shared only", "FamilyOnly"),
    ("FirstMatch (diagnostic fallback)", "FirstMatch"),
]


class LiveMonitorWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        self.cancel_requested = False
        self._row_ids = []
        self._max_rows = 5000
        self._is_busy = False
        self._pending_focus_element_id = None

    def set_status(self, text):
        self.txtStatus.Text = _safe_str(text)

    def set_busy(self, is_busy):
        self._is_busy = bool(is_busy)

    def add_item(self, text, element_id=None):
        label = _safe_str(text)
        if element_id is not None:
            label = "ID {0} | {1}".format(element_id, label)
        self.lstItems.Items.Add(label)

        try:
            self._row_ids.append(int(element_id) if element_id is not None else None)
        except Exception:
            self._row_ids.append(None)

        while self.lstItems.Items.Count > self._max_rows:
            self.lstItems.Items.RemoveAt(0)
            if self._row_ids:
                del self._row_ids[0]

        try:
            self.lstItems.ScrollIntoView(self.lstItems.Items[self.lstItems.Items.Count - 1])
        except Exception:
            pass

    def _get_selected_element_id(self):
        try:
            idx = int(self.lstItems.SelectedIndex)
        except Exception:
            idx = -1
        if idx < 0 or idx >= len(self._row_ids):
            return self._parse_selected_element_id_from_label()

        eid = self._row_ids[idx]
        if eid is None:
            return self._parse_selected_element_id_from_label()
        return eid

    def _parse_selected_element_id_from_label(self):
        try:
            selected_item = self.lstItems.SelectedItem
        except Exception:
            selected_item = None
        label = _safe_str(selected_item)
        if not label:
            return None

        match = re.match(r"\s*ID\s+(-?\d+)\s*\|", label)
        if not match:
            return None

        try:
            return int(match.group(1))
        except Exception:
            return None

    def _focus_element_now(self, element_id_value):
        if element_id_value is None:
            return False, "Select a processed row with a valid ID first."

        try:
            eid_int = int(element_id_value)
        except Exception:
            return False, "Invalid element ID: {0}".format(element_id_value)

        try:
            element_id = ElementId(eid_int)
            element = doc.GetElement(element_id)
        except Exception:
            element = None

        if element is None:
            return False, "Element {0} could not be found.".format(eid_int)

        try:
            selection_ids = List[ElementId]()
            selection_ids.Add(element_id)
        except Exception as ex:
            return False, "Could not select element {0}: {1}".format(eid_int, _safe_str(ex))

        try:
            uidoc.Selection.SetElementIds(selection_ids)
        except Exception as ex:
            return False, "Could not select element {0}: {1}".format(eid_int, _safe_str(ex))

        centered = False
        center_error = ""
        try:
            uidoc.ShowElements(element_id)
            centered = True
        except Exception as ex:
            center_error = _safe_str(ex)

        try:
            try:
                uidoc.RefreshActiveView()
            except Exception:
                pass
            try:
                uidoc.Selection.SetElementIds(selection_ids)
            except Exception:
                pass
            try:
                uidoc.ShowElements(selection_ids)
                centered = True
            except Exception:
                pass
        except Exception:
            pass

        if centered:
            return True, "Selected and focused element {0} in active view.".format(eid_int)

        if center_error:
            return True, "Selected element {0}, but Revit could not center it: {1}".format(eid_int, center_error)

        return True, "Selected element {0}.".format(eid_int)

    def focus_element(self, element_id_value):
        if self._is_busy:
            self._pending_focus_element_id = element_id_value
            return True, "Element {0} queued. It will be centered at the next safe checkpoint.".format(element_id_value)
        return self._focus_element_now(element_id_value)

    def has_pending_focus(self):
        return self._pending_focus_element_id is not None

    def drain_pending_focus(self):
        if self._pending_focus_element_id is None:
            return
        eid = self._pending_focus_element_id
        self._pending_focus_element_id = None
        ok, msg = self._focus_element_now(eid)
        self.set_status(msg)
        return ok

    def monitor_selection_changed(self, sender, e):
        element_id_value = self._get_selected_element_id()
        if element_id_value is None:
            return
        ok, msg = self.focus_element(element_id_value)
        self.set_status(msg)

    def list_selection_changed(self, sender, e):
        self.monitor_selection_changed(sender, e)

    def list_item_double_click(self, sender, e):
        element_id_value = self._get_selected_element_id()
        if element_id_value is None:
            self.set_status("Select a processed row with a valid ID first.")
            return
        ok, msg = self.focus_element(element_id_value)
        self.set_status(msg)

    def show_selected_click(self, sender, e):
        element_id_value = self._get_selected_element_id()
        if element_id_value is None:
            self.set_status("Select a processed row with a valid ID first.")
            return
        ok, msg = self.focus_element(element_id_value)
        self.set_status(msg)

    def center_selected_click(self, sender, e):
        element_id_value = self._get_selected_element_id()
        if element_id_value is None:
            self.set_status("Select a processed row with a valid ID first.")
            return
        ok, msg = self.focus_element(element_id_value)
        self.set_status(msg)

    def clear_items_click(self, sender, e):
        self.lstItems.Items.Clear()
        self._row_ids = []
        self.set_status("Processed-item list cleared.")

    def cancel_click(self, sender, e):
        self.cancel_requested = True
        self.set_status("Cancel requested. Finishing current safe step and keeping processed items...")

    def close_click(self, sender, e):
        self.Hide()


MEP_CATEGORIES = [
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctInsulations,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PipeInsulations,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_Sprinklers,
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_CableTrayFitting,
    BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_ConduitFitting,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_LightingDevices,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_NurseCallDevices,
    BuiltInCategory.OST_SecurityDevices,
    BuiltInCategory.OST_TelephoneDevices,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_Casework,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_FurnitureSystems,
    BuiltInCategory.OST_MedicalEquipment,
]


class SizeParseResult(object):
    def __init__(self, width=None, height=None, diameter=None):
        self.width = width
        self.height = height
        self.diameter = diameter


def _build_worksharing_choices():
    if not getattr(doc, "IsWorkshared", False):
        return [("Process all elements", "all")]

    return [
        ("Skip elements owned by other users (recommended)", "skip-other-users"),
        ("Only process completely unowned elements", "unowned-only"),
        ("Try all elements (may show permission prompts)", "all"),
    ]


def _build_processing_scope_choices(selected_count):
    if selected_count <= 0:
        return [("All eligible MEP elements in the active view", "view")]

    return [
        ("Currently selected elements only", "selection"),
        ("All eligible MEP elements in the active view", "view"),
    ]


def _load_run_settings(selected_count):
    settings = {
        "write_mode": getattr(config, "nwb_dim_write_mode", "skip"),
        "parameter_target_mode": getattr(config, "nwb_dim_parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE),
        "worksharing_mode": getattr(config, "nwb_dim_worksharing_mode", "skip-other-users" if getattr(doc, "IsWorkshared", False) else "all"),
        "processing_scope": getattr(config, "nwb_dim_processing_scope", "selection" if selected_count > 0 else "view"),
        "export_audit": getattr(config, "nwb_dim_export_audit", True),
        "fast_mode": getattr(config, "nwb_dim_fast_mode", False),
    }

    valid_write_modes = [value for _, value in WRITE_MODE_CHOICES]
    if settings["write_mode"] not in valid_write_modes:
        settings["write_mode"] = "skip"

    valid_target_modes = [value for _, value in PARAMETER_TARGET_MODE_CHOICES]
    if settings["parameter_target_mode"] not in valid_target_modes:
        settings["parameter_target_mode"] = DEFAULT_PARAMETER_TARGET_MODE

    valid_worksharing = [value for _, value in _build_worksharing_choices()]
    if settings["worksharing_mode"] not in valid_worksharing:
        settings["worksharing_mode"] = valid_worksharing[0]

    valid_scopes = [value for _, value in _build_processing_scope_choices(selected_count)]
    if settings["processing_scope"] not in valid_scopes:
        settings["processing_scope"] = valid_scopes[0]

    return settings


def _save_run_settings(settings):
    try:
        config.nwb_dim_write_mode = settings.get("write_mode", "skip")
        config.nwb_dim_parameter_target_mode = settings.get("parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE)
        config.nwb_dim_worksharing_mode = settings.get("worksharing_mode", "all")
        config.nwb_dim_processing_scope = settings.get("processing_scope", "view")
        config.nwb_dim_export_audit = bool(settings.get("export_audit", True))
        config.nwb_dim_fast_mode = bool(settings.get("fast_mode", False))
        script.save_config()
    except Exception:
        pass


class RunSetupWindow(WPFWindow):
    def __init__(self, xaml_file_name, selected_count, active_view_name, defaults):
        WPFWindow.__init__(self, xaml_file_name)
        self.result = None
        self._write_choices = list(WRITE_MODE_CHOICES)
        self._target_choices = list(PARAMETER_TARGET_MODE_CHOICES)
        self._worksharing_choices = _build_worksharing_choices()
        self._scope_choices = _build_processing_scope_choices(selected_count)

        self.txtActiveView.Text = _safe_str(active_view_name)
        self.txtSelectedCount.Text = str(max(0, int(selected_count)))

        self._bind_combo(self.cmbWriteMode, self._write_choices, defaults.get("write_mode", "skip"))
        self._bind_combo(self.cmbParameterTargetMode, self._target_choices, defaults.get("parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE))
        self._bind_combo(self.cmbWorksharingMode, self._worksharing_choices, defaults.get("worksharing_mode", "all"))
        self._bind_combo(self.cmbProcessingScope, self._scope_choices, defaults.get("processing_scope", "view"))
        self.chkExportAudit.IsChecked = bool(defaults.get("export_audit", True))
        self.chkFastMode.IsChecked = bool(defaults.get("fast_mode", False))

        if len(self._worksharing_choices) == 1:
            self.cmbWorksharingMode.IsEnabled = False

    def _bind_combo(self, combo, choices, selected_value):
        combo.Items.Clear()
        selected_index = 0
        for index, (label, value) in enumerate(choices):
            combo.Items.Add(label)
            if value == selected_value:
                selected_index = index
        if combo.Items.Count > 0:
            combo.SelectedIndex = selected_index

    def _selected_combo_value(self, combo, choices):
        try:
            index = int(combo.SelectedIndex)
        except Exception:
            index = 0
        if index < 0 or index >= len(choices):
            index = 0
        return choices[index][1]

    def start_click(self, sender, e):
        self.result = {
            "write_mode": self._selected_combo_value(self.cmbWriteMode, self._write_choices),
            "parameter_target_mode": self._selected_combo_value(self.cmbParameterTargetMode, self._target_choices),
            "worksharing_mode": self._selected_combo_value(self.cmbWorksharingMode, self._worksharing_choices),
            "processing_scope": self._selected_combo_value(self.cmbProcessingScope, self._scope_choices),
            "export_audit": bool(self.chkExportAudit.IsChecked),
            "fast_mode": bool(self.chkFastMode.IsChecked),
        }
        self.Close()

    def cancel_click(self, sender, e):
        self.result = None
        self.Close()


class RunLogger(object):
    def __init__(self, tool_folder, tool_title):
        stamp = time.strftime("%Y%m%d_%H%M%S")
        logs_dir = os.path.join(tool_folder, "logs")
        try:
            if not os.path.isdir(logs_dir):
                os.makedirs(logs_dir)
        except Exception:
            logs_dir = tool_folder

        safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", _safe_str(tool_title))
        self.logs_dir = logs_dir
        self.safe_title = safe_title
        self.stamp = stamp
        self.file_path = os.path.join(logs_dir, "{0}_{1}.log".format(safe_title, stamp))
        self._buffer = []
        self._flush_every = 50

    def write(self, text):
        line = _safe_str(text)
        if not line:
            return
        self._buffer.append(line)
        if len(self._buffer) >= self._flush_every:
            self.flush()

    def flush(self):
        if not self._buffer:
            return
        try:
            with open(self.file_path, "a") as stream:
                for line in self._buffer:
                    stream.write(line)
                    stream.write("\n")
        except Exception:
            pass
        self._buffer = []

    def close(self):
        self.flush()

    def info(self, text):
        self.write("INFO | {0}".format(_safe_str(text)))

    def warning(self, text):
        self.write("WARN | {0}".format(_safe_str(text)))

    def error(self, text):
        self.write("ERROR | {0}".format(_safe_str(text)))


DIM_AUDIT_COLUMNS = [
    "run_stamp",
    "tool",
    "view_name",
    "view_mode",
    "processing_scope",
    "write_mode",
    "parameter_target_mode",
    "worksharing_mode",
    "element_id",
    "category",
    "family",
    "type",
    "level",
    "name",
    "size_text",
    "parsed_width",
    "parsed_height",
    "parsed_diameter",
    "target_parameter",
    "requested_value",
    "actual_before",
    "actual_after",
    "outcome",
    "reason",
    "target_scope",
    "target_kind",
    "target_param_id",
    "target_definition_name",
    "target_storage",
    "target_is_read_only",
    "candidate_count",
    "selected_target_count",
]


class RunAuditExporter(object):
    def __init__(self, logs_dir, safe_title, stamp, enabled, columns):
        self.enabled = bool(enabled)
        self.columns = list(columns or [])
        self.file_path = ""
        self._buffer = []
        self._flush_every = 80
        if not self.enabled:
            return

        try:
            if not os.path.isdir(logs_dir):
                os.makedirs(logs_dir)
        except Exception:
            self.enabled = False
            return

        self.file_path = os.path.join(logs_dir, "{0}_{1}_audit.csv".format(safe_title, stamp))
        try:
            with open(self.file_path, "w") as stream:
                stream.write(",".join([self._csv_escape(col) for col in self.columns]))
                stream.write("\n")
        except Exception:
            self.enabled = False
            self.file_path = ""

    def _csv_escape(self, value):
        text = _safe_str(value).replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
        return '"{0}"'.format(text.replace('"', '""'))

    def write_row(self, row):
        if not self.enabled:
            return
        row = row or {}
        values = []
        for column in self.columns:
            values.append(self._csv_escape(row.get(column, "")))
        self._buffer.append(",".join(values))
        if len(self._buffer) >= self._flush_every:
            self.flush()

    def flush(self):
        if not self.enabled or not self._buffer:
            return
        try:
            with open(self.file_path, "a") as stream:
                for line in self._buffer:
                    stream.write(line)
                    stream.write("\n")
        except Exception:
            pass
        self._buffer = []

    def close(self):
        self.flush()


def _build_dim_audit_context(run_context, element, size_text, parsed):
    context = dict(run_context or {})
    context.update(
        {
            "element_id": _element_id_text(element),
            "category": _get_category_name(element),
            "family": _get_family_name(element),
            "type": _get_type_name(element),
            "level": _get_level_name(element),
            "name": _get_element_name(element),
            "size_text": _safe_str(size_text),
            "parsed_width": _safe_str(getattr(parsed, "width", "")),
            "parsed_height": _safe_str(getattr(parsed, "height", "")),
            "parsed_diameter": _safe_str(getattr(parsed, "diameter", "")),
        }
    )
    return context


def _build_dim_audit_row(audit_context, element, target_param_name, incoming_value):
    row = dict(audit_context or {})
    row.update(
        {
            "element_id": row.get("element_id", _element_id_text(element)),
            "category": row.get("category", _get_category_name(element)),
            "family": row.get("family", _get_family_name(element)),
            "type": row.get("type", _get_type_name(element)),
            "level": row.get("level", _get_level_name(element)),
            "name": row.get("name", _get_element_name(element)),
            "target_parameter": target_param_name,
            "requested_value": _safe_str(incoming_value),
            "actual_before": "",
            "actual_after": "",
            "outcome": "",
            "reason": "",
            "target_scope": "",
            "target_kind": "",
            "target_param_id": "",
            "target_definition_name": target_param_name,
            "target_storage": "",
            "target_is_read_only": "",
            "candidate_count": "",
            "selected_target_count": "",
        }
    )
    return row


def _populate_dim_audit_target_fields(row, candidate):
    candidate = candidate or {}
    param = candidate.get("param")
    row["target_scope"] = candidate.get("scope", "")
    row["target_kind"] = _parameter_kind_label(candidate.get("is_shared"))
    row["target_param_id"] = candidate.get("id", "")
    row["target_definition_name"] = _safe_str(getattr(getattr(param, "Definition", None), "Name", "")) or row.get("target_parameter", "")
    row["target_storage"] = _safe_str(getattr(param, "StorageType", ""))
    row["target_is_read_only"] = _safe_str(getattr(param, "IsReadOnly", ""))
    row["actual_before"] = _parameter_text(param)
    return row


class _TransactionFailureCapture(IFailuresPreprocessor):
    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, failures_accessor):
        try:
            failure_messages = list(failures_accessor.GetFailureMessages())
        except Exception:
            failure_messages = []

        for failure_message in failure_messages:
            try:
                description = _safe_str(failure_message.GetDescriptionText()).strip()
            except Exception:
                description = ""
            try:
                severity = _safe_str(failure_message.GetSeverity())
            except Exception:
                severity = ""

            element_ids = []
            try:
                for failing_id in failure_message.GetFailingElementIds() or []:
                    try:
                        element_ids.append(str(failing_id.IntegerValue))
                    except Exception:
                        continue
            except Exception:
                pass

            self.messages.append(
                {
                    "severity": severity,
                    "description": description,
                    "element_ids": element_ids,
                }
            )

        return FailureProcessingResult.Continue


def _attach_transaction_failure_capture(tx):
    capture = _TransactionFailureCapture()
    try:
        options = tx.GetFailureHandlingOptions()
        options = options.SetFailuresPreprocessor(capture)
        tx.SetFailureHandlingOptions(options)
    except Exception:
        pass
    return capture


def _summarize_transaction_failure_messages(capture, max_items=8):
    items = []
    if capture is None:
        return items

    for message in capture.messages[:max_items]:
        severity = _safe_str(message.get("severity", "")).strip()
        description = _safe_str(message.get("description", "")).strip()
        element_ids = ",".join(message.get("element_ids", []) or [])
        parts = []
        if severity:
            parts.append("Severity={0}".format(severity))
        if description:
            parts.append(description)
        if element_ids:
            parts.append("Elements={0}".format(element_ids))
        text = " | ".join(parts).strip()
        if text:
            items.append(text)
    return items


def _log_transaction_failure_messages(logger, prefix, capture):
    if logger is None:
        return

    messages = _summarize_transaction_failure_messages(capture)
    if not messages:
        logger.error("{0} | No Revit failure messages were captured.".format(prefix))
        return

    for message in messages:
        logger.error("{0} | {1}".format(prefix, message))


def _mark_pending_audit_rows(pending_writes, outcome, reason, audit):
    if audit is None or not pending_writes:
        return
    for pending in pending_writes:
        row = dict((pending or {}).get("audit_row") or {})
        if not row:
            continue
        if not row.get("actual_after"):
            row["actual_after"] = row.get("actual_before", "")
        row["outcome"] = outcome
        row["reason"] = reason
        audit.write_row(row)


def _safe_str(value):
    try:
        if value is None:
            return ""
        return str(value)
    except Exception:
        return ""


def _normalize_text(value):
    text = _safe_str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def _normalize_parameter_text(value):
    return _safe_str(value).strip().replace("\r", "").replace("\n", "")


def _pump_ui():
    try:
        WinFormsApplication.DoEvents()
    except Exception:
        pass


def _monitor_set_status(monitor, text):
    if monitor is None:
        return
    try:
        monitor.set_status(text)
    except Exception:
        pass


def _monitor_add_item(monitor, text, element_id=None):
    if monitor is None:
        return
    try:
        monitor.add_item(text, element_id)
    except Exception:
        pass


def _is_blank(text):
    return not _safe_str(text).strip()


def _view_label(active_view):
    if isinstance(active_view, ViewSchedule):
        return "schedule"
    return "model"


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
    return True


def _collect_active_schedule_elements(schedule_view):
    elements = []
    try:
        for element in FilteredElementCollector(doc, schedule_view.Id).WhereElementIsNotElementType().ToElements():
            if element is None:
                continue
            if not _is_target_mep_category(element):
                continue
            try:
                _ = element.Id.IntegerValue
                elements.append(element)
            except Exception:
                continue
    except Exception:
        pass
    return elements


def _collect_active_model_view_elements(view):
    elements = []
    seen = set()

    for bic in MEP_CATEGORIES:
        try:
            collector = FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType()
        except Exception:
            continue

        try:
            for element in collector:
                if element is None:
                    continue
                try:
                    eid = int(element.Id.IntegerValue)
                except Exception:
                    continue
                if eid in seen:
                    continue
                seen.add(eid)
                elements.append(element)
        except Exception:
            continue

    return elements


def _is_target_mep_category(element):
    if element is None:
        return False

    try:
        category = element.Category
        if category is None or category.Id is None:
            return False
        category_id = int(category.Id.IntegerValue)
    except Exception:
        return False

    for bic in MEP_CATEGORIES:
        try:
            if int(ElementId(bic).IntegerValue) == category_id:
                return True
        except Exception:
            continue
    return False


def _get_selected_mep_elements():
    elements = []
    seen = set()

    try:
        selected_ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        selected_ids = []

    for element_id in selected_ids:
        try:
            element = doc.GetElement(element_id)
        except Exception:
            element = None

        if element is None or not _is_target_mep_category(element):
            continue

        try:
            eid = int(element.Id.IntegerValue)
        except Exception:
            continue

        if eid in seen:
            continue
        seen.add(eid)
        elements.append(element)

    return elements


def _first_largest_rect_pair(size_text):
    matches = list(re.finditer(r"(?i)(\d+(?:[\.,]\d+)?)\s*[x×]\s*(\d+(?:[\.,]\d+)?)", size_text))
    if not matches:
        return None

    best = None
    for match in matches:
        width = match.group(1).replace(",", ".")
        height = match.group(2).replace(",", ".")
        area = 0.0
        try:
            area = float(width) * float(height)
        except Exception:
            area = 0.0
        item = (width, height, area)
        if best is None or item[2] > best[2]:
            best = item

    return best


def _largest_diameter(size_text):
    candidates = []
    patterns = [
        r"(?i)(\d+(?:[\.,]\d+)?)\s*(?:o|O|dia\b|diameter\b|dn\b)",
        r"(?i)(?:o|O|dia\b|diameter\b|dn\s*)(\d+(?:[\.,]\d+)?)",
        r"(?i)(\d+(?:[\.,]\d+)?)\s*(?:ø|Ø)",
        r"(?i)(?:ø|Ø)\s*(\d+(?:[\.,]\d+)?)",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, size_text):
            raw = match.group(1).replace(",", ".")
            try:
                num = float(raw)
            except Exception:
                continue
            candidates.append((num, raw))

    if not candidates:
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def _parse_size(size_text):
    if _is_blank(size_text):
        return SizeParseResult()

    text = _safe_str(size_text).strip()
    width = None
    height = None
    diameter = None

    rect = _first_largest_rect_pair(text)
    if rect:
        width, height, _ = rect

    diameter = _largest_diameter(text)
    if diameter is None and width is None and height is None:
        match = re.match(r"^\s*(\d+(?:[\.,]\d+)?)\s*(?:mm)?\s*$", text, re.IGNORECASE)
        if match:
            diameter = match.group(1).replace(",", ".")

    return SizeParseResult(width=width, height=height, diameter=diameter)


def _extract_numeric_dimension_text(raw_text):
    text = _safe_str(raw_text).strip()
    if not text:
        return ""

    match = re.search(r"(\d+(?:[\.,]\d+)?)", text)
    if not match:
        return ""

    return match.group(1).replace(",", ".")


def _format_dimension_text(raw_number):
    try:
        number = float(raw_number)
    except Exception:
        return ""

    rounded = round(number, 3)
    if abs(rounded - round(rounded)) <= 1e-6:
        return str(int(round(rounded)))
    return ("{0:.3f}".format(rounded)).rstrip("0").rstrip(".")


def _read_numeric_dimension_from_param(param):
    if param is None:
        return ""

    try:
        if param.StorageType == StorageType.Double:
            return _format_dimension_text(float(param.AsDouble()) * 304.8)
        if param.StorageType == StorageType.Integer:
            return _format_dimension_text(param.AsInteger())
    except Exception:
        pass

    return _extract_numeric_dimension_text(_read_stringish(param))


def _read_numeric_dimension_by_name(element, names):
    for name in names:
        param = _find_param(element, name)
        value = _read_numeric_dimension_from_param(param)
        if value:
            return value

    type_el = _get_type_element(element)
    for name in names:
        param = _find_param(type_el, name)
        value = _read_numeric_dimension_from_param(param)
        if value:
            return value

    return ""


def _normalize_dimension_name(name):
    return re.sub(r"[^a-z0-9]+", "", _safe_str(name).strip().lower())


def _iter_named_dimension_candidates(element):
    owners = [("instance", element), ("type", _get_type_element(element))]
    seen = set()

    for scope, owner in owners:
        if owner is None:
            continue

        try:
            params = owner.Parameters
        except Exception:
            params = []

        for param in params:
            if param is None:
                continue

            try:
                definition = param.Definition
            except Exception:
                definition = None
            if definition is None:
                continue

            name = _safe_str(getattr(definition, "Name", "")).strip()
            if not name:
                continue

            value = _read_numeric_dimension_from_param(param)
            if not value:
                continue

            try:
                numeric_value = float(value)
            except Exception:
                continue
            if numeric_value <= 0:
                continue

            norm_name = _normalize_dimension_name(name)
            if not norm_name or norm_name.startswith("nwbdim"):
                continue

            try:
                param_id = _safe_str(getattr(param.Id, "IntegerValue", ""))
            except Exception:
                param_id = ""

            seen_key = (scope, norm_name, param_id)
            if seen_key in seen:
                continue
            seen.add(seen_key)

            yield {
                "scope": scope,
                "name": name,
                "norm_name": norm_name,
                "value": value,
                "numeric": numeric_value,
                "param": param,
            }


def _score_dimension_candidate(candidate, preferred_names, keywords, negative_keywords=None):
    norm_name = candidate.get("norm_name", "")
    score = 0

    for index, name in enumerate(preferred_names or []):
        norm_target = _normalize_dimension_name(name)
        if not norm_target:
            continue
        if norm_name == norm_target:
            score = max(score, 1000 - index)
        elif norm_target in norm_name:
            score = max(score, 900 - index)

    for index, keyword in enumerate(keywords or []):
        norm_keyword = _normalize_dimension_name(keyword)
        if not norm_keyword:
            continue
        if norm_name == norm_keyword:
            score = max(score, 800 - index)
        elif norm_keyword in norm_name:
            score = max(score, 700 - index)

    for keyword in negative_keywords or []:
        norm_keyword = _normalize_dimension_name(keyword)
        if norm_keyword and norm_keyword in norm_name:
            score -= 300

    if candidate.get("scope") == "instance":
        score += 10

    return score


def _pick_best_dimension_candidate(element, preferred_names, keywords, negative_keywords=None):
    exact_value = _read_numeric_dimension_by_name(element, preferred_names or [])
    if exact_value:
        return exact_value

    best = None
    best_score = 0
    for candidate in _iter_named_dimension_candidates(element):
        score = _score_dimension_candidate(candidate, preferred_names, keywords, negative_keywords)
        if score <= 0:
            continue
        if best is None or score > best_score:
            best = candidate
            best_score = score

    if best is None:
        return ""
    return best.get("value", "")


def _resolve_air_terminal_size_data(element, size_text, parsed):
    width_text = _pick_best_dimension_candidate(
        element,
        ["MHT_WIDTH", "Width", "Face Width", "Overall Width", "L_size", "L Size"],
        ["mhtwidth", "width", "facewidth", "overallwidth", "neckwidth", "corewidth", "modulewidth"],
        ["diameter", "dia", "radius", "flow", "pressure", "elevation", "offset"],
    )
    length_text = _pick_best_dimension_candidate(
        element,
        ["MHT_LENGTH", "Length", "Face Length", "Overall Length", "L_size", "L Size"],
        ["mhtlength", "length", "facelength", "overalllength", "modulelength"],
        ["ductlength", "linearlength", "diameter", "dia", "flow", "pressure", "elevation"],
    )
    height_text = _pick_best_dimension_candidate(
        element,
        ["Height", "Face Height", "Overall Height", "W_Size", "W Size"],
        ["height", "faceheight", "overallheight", "moduleheight"],
        ["ceilingheight", "mountingheight", "elevation", "diameter", "dia", "flow", "pressure"],
    )
    depth_text = _pick_best_dimension_candidate(
        element,
        ["MHT_DEPTH", "Depth", "Face Depth", "Overall Depth"],
        ["mhtdepth", "depth", "facedepth", "overalldepth", "boxdepth"],
        ["depthoffset", "insulationdepth", "elevation", "diameter", "dia", "flow", "pressure"],
    )
    diameter_text = _pick_best_dimension_candidate(
        element,
        ["MHT_DIAMETER", "Nominal Diameter", "Diameter", "Dia", "Neck Diameter", "Spigot Diameter"],
        ["mhtdiameter", "nominaldiameter", "diameter", "dia", "neckdiameter", "spigotdiameter", "connectordiameter"],
        ["radius", "flow", "pressure", "elevation"],
    )

    primary_width = width_text or length_text
    secondary_height = height_text or depth_text

    if primary_width and not secondary_height:
        alternate = length_text if length_text and length_text != primary_width else ""
        if not alternate and width_text and width_text != primary_width:
            alternate = width_text
        if not alternate:
            alternate = depth_text
        secondary_height = alternate

    if not primary_width and secondary_height:
        primary_width = width_text or length_text or depth_text

    if primary_width:
        parsed.width = primary_width
    if secondary_height:
        parsed.height = secondary_height
    if diameter_text and _is_blank(parsed.diameter):
        parsed.diameter = diameter_text

    if primary_width or secondary_height:
        if primary_width and secondary_height:
            size_text = "{0}x{1}".format(primary_width, secondary_height)
        else:
            size_text = primary_width or secondary_height
    elif diameter_text and _is_blank(size_text):
        size_text = diameter_text

    return size_text, parsed


def _read_numeric_dimension_by_bip(element, bip_name):
    try:
        bip = getattr(BuiltInParameter, bip_name)
    except Exception:
        return ""
    if bip is None:
        return ""

    try:
        return _read_numeric_dimension_from_param(element.get_Parameter(bip))
    except Exception:
        return ""


def _has_category(element, builtin_category):
    if element is None:
        return False

    try:
        category = element.Category
        if category is None or category.Id is None:
            return False
        return int(category.Id.IntegerValue) == int(ElementId(builtin_category).IntegerValue)
    except Exception:
        return False


def _get_element_size_data(element):
    size_text = _read_param_text_by_name(
        element,
        [
            "Size",
            "Calculated Size",
            "Overall Size",
            "Nominal Diameter",
            "Diameter",
            "Dia",
            "Duct Diameter",
            "Flex Duct Diameter",
            "Trade Size",
        ],
    )

    if _is_blank(size_text):
        for bip_name in ["RBS_CALCULATED_SIZE", "RBS_CURVE_DIAMETER_PARAM"]:
            size_text = _read_param_text_by_bip(element, bip_name)
            if not _is_blank(size_text):
                break

    parsed = _parse_size(size_text)

    if _is_blank(parsed.diameter):
        diameter_value = _read_numeric_dimension_by_name(
            element,
            ["Nominal Diameter", "Diameter", "Dia", "Duct Diameter", "Flex Duct Diameter", "Trade Size"],
        )
        if not diameter_value:
            diameter_value = _read_numeric_dimension_by_bip(element, "RBS_CURVE_DIAMETER_PARAM")
        if diameter_value:
            parsed.diameter = diameter_value
            if _is_blank(size_text):
                size_text = diameter_value

    if _has_category(element, BuiltInCategory.OST_DuctTerminal):
        size_text, parsed = _resolve_air_terminal_size_data(element, size_text, parsed)

    if _has_category(element, BuiltInCategory.OST_FlexDuctCurves) and _is_blank(parsed.diameter):
        diameter_text = _read_numeric_dimension_by_name(
            element,
            ["Diameter", "Duct Diameter", "Flex Duct Diameter", "Nominal Diameter", "Dia"]
        )
        if _is_blank(diameter_text):
            diameter_text = _read_numeric_dimension_by_bip(element, "RBS_CURVE_DIAMETER_PARAM")
        diameter_value = _extract_numeric_dimension_text(diameter_text)
        if diameter_value:
            parsed.diameter = diameter_value
            if _is_blank(size_text):
                size_text = diameter_text

    return size_text, parsed


def _find_param(element, name):
    if element is None or _is_blank(name):
        return None

    try:
        param = element.LookupParameter(name)
        if param is not None:
            return param
    except Exception:
        pass

    try:
        params = element.GetParameters(name)
        if params:
            return params[0]
    except Exception:
        pass

    try:
        for param in element.Parameters:
            try:
                if param.Definition and param.Definition.Name == name:
                    return param
            except Exception:
                continue
    except Exception:
        pass

    return None


def _get_type_element(element):
    if element is None:
        return None
    try:
        type_id = element.GetTypeId()
        if type_id is None or type_id == ElementId.InvalidElementId:
            return None
        return doc.GetElement(type_id)
    except Exception:
        return None


def _get_element_name(element):
    if element is None:
        return ""
    try:
        return _safe_str(getattr(element, "Name", "")).strip()
    except Exception:
        return ""


def _get_category_name(element):
    try:
        cat = element.Category
        if cat is not None:
            return _safe_str(cat.Name).strip()
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

    type_el = _get_type_element(element)
    if type_el is not None:
        try:
            return _safe_str(getattr(type_el, "FamilyName", "")).strip()
        except Exception:
            pass

    return ""


def _get_type_name(element):
    type_el = _get_type_element(element)
    if type_el is not None:
        try:
            return _safe_str(getattr(type_el, "Name", "")).strip()
        except Exception:
            pass
    return ""


def _get_level_name(element):
    try:
        lvl_id = element.LevelId
        if lvl_id is not None and lvl_id != ElementId.InvalidElementId and lvl_id.IntegerValue > 0:
            lvl = doc.GetElement(lvl_id)
            if lvl is not None:
                return _safe_str(getattr(lvl, "Name", "")).strip()
    except Exception:
        pass
    return ""


def _summarize_category_counts(elements):
    counts = {}
    for element in elements or []:
        name = _get_category_name(element) or "<No Category>"
        counts[name] = counts.get(name, 0) + 1
    if not counts:
        return ""
    return "; ".join(["{0}={1}".format(name, counts[name]) for name in sorted(counts.keys())])


def _summarize_element_ids(elements, max_ids=200):
    ids = []
    for element in elements or []:
        text = _element_id_text(element)
        if text and text != "?":
            ids.append(text)
    if not ids:
        return ""
    if len(ids) <= max_ids:
        return ", ".join(ids)
    return "{0} ... (+{1} more)".format(", ".join(ids[:max_ids]), len(ids) - max_ids)


def _log_size_runtime_context(logger, element, element_id_text, size_text, parsed, target_values):
    if logger is None:
        return

    logger.write(
        "SOURCE | Element {0} | Cat={1} | Family={2} | Type={3} | Level={4} | Name={5} | System={6} | SizeText={7} | Width={8} | Height={9} | Diameter={10}".format(
            element_id_text,
            _get_category_name(element),
            _get_family_name(element),
            _get_type_name(element),
            _get_level_name(element),
            _get_element_name(element),
            _read_param_text_by_bip(element, "RBS_SYSTEM_CLASSIFICATION_PARAM") or _read_param_text_by_name(element, ["System Classification", "System Type", "System Name"]),
            _safe_str(size_text),
            _safe_str(target_values.get("NWB_DimWidth", "")),
            _safe_str(target_values.get("NWB_DimHeight", "")),
            _safe_str(target_values.get("NWB_DimDiameter", "")),
        )
    )

    if _is_blank(target_values.get("NWB_DimWidth", "")) and _is_blank(target_values.get("NWB_DimHeight", "")) and _is_blank(target_values.get("NWB_DimDiameter", "")):
        logger.write("WARN | Element {0} | Size parsed to blank NWB dim targets".format(element_id_text))


def _iter_param_owners(element):
    owners = []
    seen = set()

    for scope_name, owner in [("instance", element), ("type", _get_type_element(element))]:
        if owner is None:
            continue
        try:
            owner_id = int(owner.Id.IntegerValue)
        except Exception:
            owner_id = id(owner)
        if owner_id in seen:
            continue
        seen.add(owner_id)
        owners.append((scope_name, owner))

    return owners


def _parameter_text(param):
    return _read_stringish(param)


def _param_debug_id(param):
    if param is None:
        return ""
    try:
        return _safe_str(param.Id)
    except Exception:
        return ""


def _param_is_shared(param):
    if param is None:
        return False
    try:
        return bool(param.IsShared)
    except Exception:
        return False


def _iter_named_params(owner, name):
    seen = set()

    try:
        param = owner.LookupParameter(name)
        if param is not None:
            key = _param_debug_id(param) or str(id(param))
            seen.add(key)
            yield param
    except Exception:
        pass

    try:
        params = owner.GetParameters(name)
    except Exception:
        params = []
    for param in params or []:
        if param is None:
            continue
        key = _param_debug_id(param) or str(id(param))
        if key in seen:
            continue
        seen.add(key)
        yield param

    try:
        for param in owner.Parameters:
            if param is None or param.Definition is None:
                continue
            if param.Definition.Name != name:
                continue
            key = _param_debug_id(param) or str(id(param))
            if key in seen:
                continue
            seen.add(key)
            yield param
    except Exception:
        pass


def _parameter_kind_label(is_shared):
    return "Shared Parameter" if is_shared else "Family/Other Parameter"


def _candidate_sort_key(candidate):
    param = candidate.get("param")

    storage_rank = 9
    try:
        storage_type = param.StorageType
        if storage_type == StorageType.String:
            storage_rank = 0
        elif storage_type == StorageType.Double:
            storage_rank = 1
        elif storage_type == StorageType.Integer:
            storage_rank = 2
        elif storage_type == StorageType.ElementId:
            storage_rank = 3
    except Exception:
        storage_rank = 9

    scope_priority = 0 if candidate["scope"] == "instance" else 1
    shared_priority = 0 if candidate["is_shared"] else 1
    return (
        scope_priority,
        shared_priority,
        storage_rank,
        candidate["id"],
    )


def _best_candidate(candidates):
    if not candidates:
        return None
    ordered = sorted(candidates, key=_candidate_sort_key)
    return ordered[0]


def _collect_param_candidates(element, name, logger=None):
    candidates = []

    for scope_name, owner in _iter_param_owners(element):
        for param in _iter_named_params(owner, name):
            candidate = {
                "param": param,
                "scope": scope_name,
                "owner": owner,
                "id": _param_debug_id(param),
                "is_shared": _param_is_shared(param),
            }
            candidates.append(candidate)
            if logger is not None:
                logger.info(
                    "Candidate: {0} | Scope={1} | Storage={2} | ReadOnly={3} | Shared={4} | Id={5}".format(
                        _safe_str(getattr(getattr(param, "Definition", None), "Name", "")),
                        scope_name,
                        _safe_str(getattr(param, "StorageType", "")),
                        _safe_str(getattr(param, "IsReadOnly", "")),
                        _param_is_shared(param),
                        _param_debug_id(param),
                    )
                )

    return sorted(candidates, key=_candidate_sort_key)


def _find_param_target(element, name, target_mode, logger=None):
    candidates = _collect_param_candidates(element, name, logger)
    if not candidates:
        return [], "missing", []

    writable = [candidate for candidate in candidates if not candidate["param"].IsReadOnly]
    if not writable:
        return [], "readonly", candidates

    shared = [candidate for candidate in writable if candidate["is_shared"]]
    nonshared = [candidate for candidate in writable if not candidate["is_shared"]]

    if target_mode == "Both":
        return writable, "both", candidates

    if target_mode == "FamilyOnly":
        if nonshared:
            best = _best_candidate(nonshared)
            return [best] if best is not None else [], "family-only", candidates
        return [], "no-family", candidates

    if target_mode == "FirstMatch":
        preferred = _best_candidate(shared) if shared else _best_candidate(writable)
        return [preferred], "first-match", candidates

    if shared:
        best = _best_candidate(shared)
        return [best] if best is not None else [], "shared-only", candidates
    if nonshared:
        return [], "no-shared", candidates

    return [], "missing", candidates


def _read_stringish(param):
    if param is None:
        return ""

    try:
        if param.StorageType == StorageType.String:
            return _safe_str(param.AsString())
        if param.StorageType == StorageType.Double:
            return _safe_str(param.AsValueString())
        if param.StorageType == StorageType.Integer:
            return _safe_str(param.AsInteger())
        if param.StorageType == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None:
                return ""
            return _safe_str(eid.IntegerValue)
    except Exception:
        pass

    try:
        return _safe_str(param.AsValueString())
    except Exception:
        return ""


def _read_raw_value(param):
    if param is None:
        return None

    try:
        storage_type = param.StorageType
        if storage_type == StorageType.String:
            return param.AsString() or ""
        if storage_type == StorageType.Integer:
            return param.AsInteger()
        if storage_type == StorageType.Double:
            return param.AsDouble()
        if storage_type == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None:
                return None
            return int(eid.IntegerValue)
    except Exception:
        pass

    return None


def _parse_numeric_text(raw_value):
    text = _safe_str(raw_value).strip()
    if not text:
        raise ValueError("blank")
    return float(text.replace(",", "."))


def _extract_first_number(text):
    match = re.search(r"[-+]?\d+(?:[\.,]\d+)?", _safe_str(text))
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", "."))
    except Exception:
        return None


def _is_effective_value_set(param):
    if param is None:
        return False

    try:
        storage_type = param.StorageType
    except Exception:
        storage_type = None

    raw_value = _read_raw_value(param)
    if storage_type == StorageType.String:
        return not _is_blank(raw_value)
    if storage_type == StorageType.Integer:
        try:
            return int(raw_value) != 0
        except Exception:
            return False
    if storage_type == StorageType.Double:
        try:
            return abs(float(raw_value)) > 1e-9
        except Exception:
            return False
    if storage_type == StorageType.ElementId:
        try:
            return int(raw_value) > 0
        except Exception:
            return False

    return raw_value is not None


def _parameter_value_matches_incoming(param, incoming_value):
    if param is None:
        return False

    incoming_text = _safe_str(incoming_value).strip()
    expected_text = _normalize_parameter_text(incoming_text)

    try:
        storage_type = param.StorageType
    except Exception:
        storage_type = None

    if storage_type == StorageType.String:
        return _normalize_parameter_text(param.AsString()) == expected_text

    if storage_type == StorageType.Integer:
        try:
            return int(param.AsInteger()) == int(round(_parse_numeric_text(incoming_text)))
        except Exception:
            try:
                current_text = _normalize_parameter_text(param.AsValueString())
            except Exception:
                current_text = _normalize_parameter_text(_parameter_text(param))
            if current_text == expected_text:
                return True
            expected_num = _extract_first_number(incoming_text)
            current_num = _extract_first_number(current_text)
            if expected_num is None or current_num is None:
                return False
            return int(round(current_num)) == int(round(expected_num))

    if storage_type == StorageType.Double:
        try:
            current_text = _normalize_parameter_text(param.AsValueString())
        except Exception:
            current_text = _normalize_parameter_text(_parameter_text(param))

        if current_text == expected_text:
            return True

        expected_num = _extract_first_number(incoming_text)
        current_num = _extract_first_number(current_text)
        if expected_num is not None and current_num is not None:
            if abs(current_num - expected_num) <= 1e-6:
                return True

        try:
            return abs(float(param.AsDouble()) - float(_parse_numeric_text(incoming_text))) <= 1e-6
        except Exception:
            return False

    if storage_type == StorageType.ElementId:
        try:
            eid = param.AsElementId()
            current_value = int(eid.IntegerValue) if eid is not None else None
            return current_value == int(round(_parse_numeric_text(incoming_text)))
        except Exception:
            return _normalize_parameter_text(param.AsValueString()) == expected_text

    return _normalize_parameter_text(_parameter_text(param)) == expected_text


def _write_stringish(param, value):
    if param is None:
        return False, "parameter missing"
    if param.IsReadOnly:
        return False, "Parameter is read-only"

    text_value = _safe_str(value).strip()

    def _verify_after_set(set_ok):
        if not set_ok:
            return False, "Set() returned False"

        try:
            doc.Regenerate()
        except Exception:
            pass

        actual = _parameter_text(param)
        if not _parameter_value_matches_incoming(param, text_value):
            return False, "Verification failed after Set() | actual='{0}' expected='{1}'".format(actual, text_value)
        return True, ""

    try:
        if param.StorageType == StorageType.String:
            return _verify_after_set(bool(param.Set(text_value)))

        if param.StorageType == StorageType.Integer:
            success = False
            if hasattr(param, "SetValueString"):
                try:
                    success = bool(param.SetValueString(text_value))
                except Exception:
                    success = False
            if not success:
                try:
                    success = bool(param.Set(int(round(_parse_numeric_text(text_value)))))
                except Exception as ex:
                    return False, _safe_str(ex)
            return _verify_after_set(success)

        if param.StorageType == StorageType.Double:
            success = False
            if hasattr(param, "SetValueString"):
                try:
                    success = bool(param.SetValueString(text_value))
                except Exception:
                    success = False
            if not success:
                try:
                    success = bool(param.Set(float(_parse_numeric_text(text_value))))
                except Exception as ex:
                    return False, _safe_str(ex)
            return _verify_after_set(success)

        if hasattr(param, "SetValueString"):
            try:
                success = bool(param.SetValueString(text_value))
            except Exception:
                success = False
            return _verify_after_set(success)

        return False, "unsupported storage type or SetValueString returned False"
    except Exception as ex:
        return False, _safe_str(ex)


def _choose_write_mode():
    choice = forms.CommandSwitchWindow.show(
        [
            "Keep existing NWB dim values that are already set",
            "Override existing NWB dim values",
        ],
        message="Choose write mode for NWB_Dim parameters:",
    )
    if not choice:
        return None
    return "skip" if choice.startswith("Keep existing") else "override"


def _choose_parameter_target_mode():
    choice = forms.CommandSwitchWindow.show(
        [
            "{0} (recommended)".format(DEFAULT_PARAMETER_TARGET_MODE),
            "Both (shared + legacy duplicates)",
            "FamilyOnly / non-shared only",
            "FirstMatch (diagnostic fallback)",
        ],
        message="Choose parameter target mode for duplicate NWB dim parameter names:",
    )
    if not choice:
        return None
    if choice.startswith(DEFAULT_PARAMETER_TARGET_MODE):
        return DEFAULT_PARAMETER_TARGET_MODE
    if choice.startswith("Both"):
        return "Both"
    if choice.startswith("FamilyOnly"):
        return "FamilyOnly"
    return "FirstMatch"


def _choose_worksharing_mode():
    if not getattr(doc, "IsWorkshared", False):
        return "all"

    choice = forms.CommandSwitchWindow.show(
        [
            "Skip elements owned by other users (recommended)",
            "Only process completely unowned elements",
            "Try all elements (may show permission prompts)",
        ],
        message="Choose worksharing handling:",
    )
    if not choice:
        return None
    if choice.startswith("Skip elements owned"):
        return "skip-other-users"
    if choice.startswith("Only process"):
        return "unowned-only"
    return "all"


def _choose_processing_scope(selected_count):
    if selected_count <= 0:
        return "view"

    choice = forms.CommandSwitchWindow.show(
        [
            "Currently selected elements only",
            "All eligible MEP elements in the active view",
        ],
        message="Choose processing scope ({0} eligible selected element(s) found):".format(selected_count),
    )
    if not choice:
        return None
    if choice.startswith("Currently selected"):
        return "selection"
    return "view"


def _prompt_run_settings(selected_count, active_view_name):
    defaults = _load_run_settings(selected_count)

    try:
        dialog = RunSetupWindow("RunOptionsWindow.xaml", selected_count, active_view_name, defaults)
        dialog.ShowDialog()
    except Exception as ex:
        forms.alert(
            "The NWB dim setup window could not be opened.\n\n{0}".format(_safe_str(ex)),
            title=TOOL_TITLE,
        )
        return None

    if not dialog.result:
        return None

    _save_run_settings(dialog.result)
    return dialog.result


def _current_username():
    try:
        return _safe_str(doc.Application.Username).strip()
    except Exception:
        return ""


def _read_param_text_by_name(element, names):
    for name in names:
        param = _find_param(element, name)
        text = _read_stringish(param).strip()
        if text:
            return text

    type_el = _get_type_element(element)
    for name in names:
        param = _find_param(type_el, name)
        text = _read_stringish(param).strip()
        if text:
            return text

    return ""


def _read_param_text_by_bip(element, bip_name):
    try:
        bip = getattr(BuiltInParameter, bip_name)
    except Exception:
        return ""
    if bip is None:
        return ""

    try:
        param = element.get_Parameter(bip)
        text = _read_stringish(param).strip()
        if text:
            return text
    except Exception:
        pass

    return ""


def _get_edited_by_text(element):
    text = _read_param_text_by_bip(element, "EDITED_BY")
    if text:
        return text.strip()
    return _read_param_text_by_name(element, ["Edited By"]).strip()


def _get_element_identifier(element):
    if element is None:
        return None
    try:
        return int(element.Id.IntegerValue)
    except Exception:
        return None


def _get_worksharing_status(element, owner_status_cache):
    if element is None or not getattr(doc, "IsWorkshared", False):
        return "not-workshared", ""

    element_id = _get_element_identifier(element)
    if element_id is not None and element_id in owner_status_cache:
        return owner_status_cache[element_id]

    edited_by = _get_edited_by_text(element)
    current_user_norm = _normalize_text(_current_username())
    edited_by_norm = _normalize_text(edited_by)

    if not edited_by_norm:
        result = ("unowned", "")
    elif edited_by_norm == current_user_norm:
        result = ("owned-by-me", edited_by)
    else:
        result = ("owned-by-other", edited_by)

    if element_id is not None:
        owner_status_cache[element_id] = result
    return result


def _is_element_allowed_by_worksharing(element, mode, owner_status_cache):
    if mode == "all" or not getattr(doc, "IsWorkshared", False):
        return True, ""

    for owner_scope, owner in _iter_param_owners(element):
        status, owner_name = _get_worksharing_status(owner, owner_status_cache)
        if mode == "skip-other-users" and status == "owned-by-other":
            return False, "{0} owned by {1}".format(owner_scope, owner_name or "another user")
        if mode == "unowned-only" and status != "unowned":
            if status == "owned-by-me":
                return False, "{0} already owned by current user".format(owner_scope)
            if status == "owned-by-other":
                return False, "{0} owned by {1}".format(owner_scope, owner_name or "another user")

    return True, ""


def _sample_writable_parameter_names(element, max_items=20):
    names = []
    seen = set()
    for _, owner in _iter_param_owners(element):
        try:
            for param in owner.Parameters:
                if param is None or param.Definition is None or param.IsReadOnly:
                    continue
                name = _safe_str(param.Definition.Name).strip()
                if not name or name in seen:
                    continue
                seen.add(name)
                names.append(name)
        except Exception:
            continue
        if len(names) >= max_items:
            break
    return sorted(names)[:max_items]


def _find_candidate_param_by_id(owner, param_name, param_id):
    if owner is None:
        return None

    for candidate in _iter_named_params(owner, param_name):
        if _param_debug_id(candidate) == _safe_str(param_id):
            return candidate

    return _find_param(owner, param_name)


def _element_id_text(element):
    try:
        return str(element.Id.IntegerValue)
    except Exception:
        return "?"


def _dump_element_parameters_for_debug(element, logger):
    if element is None or logger is None:
        return

    for scope_name, owner in _iter_param_owners(element):
        try:
            for param in owner.Parameters:
                if param is None or param.Definition is None:
                    continue
                logger.warning(
                    "ParamDump | Element {0} | Scope={1} | {2} | Storage={3} | ReadOnly={4} | Value='{5}'".format(
                        _element_id_text(element),
                        scope_name,
                        _safe_str(param.Definition.Name),
                        _safe_str(getattr(param, "StorageType", "")),
                        _safe_str(getattr(param, "IsReadOnly", "")),
                        _parameter_text(param),
                    )
                )
        except Exception:
            continue


def _apply_pending_write(pending):
    try:
        element = doc.GetElement(ElementId(int(pending.get("element_id_value"))))
    except Exception:
        element = None

    if element is None:
        return False, "element missing during retry"

    target_scope = pending.get("target_scope")
    owner = _get_type_element(element) if target_scope == "type" else element
    if owner is None:
        return False, "target owner missing during retry"

    param_name = pending.get("param_name")
    param_id = pending.get("param_id")
    param = _find_candidate_param_by_id(owner, param_name, param_id)
    if param is None:
        return False, "target parameter missing during retry"

    ok, err = _write_stringish(param, pending.get("incoming_value", ""))
    if ok:
        return True, ""
    return False, err or "retry write failed"


def _set_if_needed(element, target_param_name, incoming_value, mode, parameter_target_mode, counters, pending_writes, logger=None, audit=None, audit_context=None):
    if _is_blank(incoming_value):
        counters["skip_no_source"] += 1
        if audit is not None:
            row = _build_dim_audit_row(audit_context, element, target_param_name, incoming_value)
            row["outcome"] = "SKIP"
            row["reason"] = "no parsed value"
            audit.write_row(row)
        return "SKIP", "no parsed value"

    targets, resolution_reason, all_candidates = _find_param_target(element, target_param_name, parameter_target_mode, logger)
    candidate_count = len(all_candidates)
    if not targets:
        if resolution_reason == "readonly":
            counters["read_only"] += 1
        elif resolution_reason in ("no-family", "no-shared"):
            counters["skip_target_mode"] += 1
        else:
            counters["missing_target"] += 1
        if logger is not None:
            logger.error(
                "Element {0} | Category={1} | Family={2} | Parameter={3} | Reason={4} | Writable sample: {5}".format(
                    _element_id_text(element),
                    _safe_str(getattr(getattr(element, "Category", None), "Name", "")),
                    _get_element_name(_get_type_element(element) or element),
                    target_param_name,
                    resolution_reason,
                    ", ".join(_sample_writable_parameter_names(element)),
                )
            )
        if audit is not None:
            row = _build_dim_audit_row(audit_context, element, target_param_name, incoming_value)
            row["candidate_count"] = candidate_count
            row["selected_target_count"] = 0
            row["outcome"] = "SKIP"
            row["reason"] = resolution_reason
            audit.write_row(row)
        return "SKIP", resolution_reason

    if candidate_count > 1:
        counters["duplicate_parameter"] += 1
        if logger is not None:
            logger.info("Found {0} parameters named {1}".format(candidate_count, target_param_name))

    selected_ids = set([candidate["id"] for candidate in targets])
    if logger is not None and candidate_count > 1:
        for candidate in all_candidates:
            action = "SELECTED"
            if candidate["id"] not in selected_ids:
                action = "IGNORED ({0})".format(parameter_target_mode)
            logger.info(
                "{0} | {1} | Scope={2} | ReadOnly={3} | Id={4}".format(
                    action,
                    _parameter_kind_label(candidate["is_shared"]),
                    candidate["scope"],
                    _safe_str(getattr(candidate["param"], "IsReadOnly", "")),
                    candidate["id"],
                )
            )

    written_count = 0
    failed_count = 0
    messages = []

    for candidate in targets:
        param = candidate["param"]
        target_scope = candidate["scope"]
        is_shared = candidate["is_shared"]

        if param.IsReadOnly:
            counters["read_only"] += 1
            counters["failed"] += 1
            failed_count += 1
            messages.append("{0}:{1}:read-only".format(target_scope, candidate["id"]))
            if audit is not None:
                row = _build_dim_audit_row(audit_context, element, target_param_name, incoming_value)
                row["candidate_count"] = candidate_count
                row["selected_target_count"] = len(targets)
                _populate_dim_audit_target_fields(row, candidate)
                row["actual_after"] = row.get("actual_before", "")
                row["outcome"] = "FAIL"
                row["reason"] = "read-only"
                audit.write_row(row)
            continue

        current_value = _parameter_text(param)
        if mode == "skip" and _is_effective_value_set(param):
            counters["skip_existing"] += 1
            messages.append("{0}:{1}:existing-kept({2})".format(target_scope, candidate["id"], current_value or "set"))
            if audit is not None:
                row = _build_dim_audit_row(audit_context, element, target_param_name, incoming_value)
                row["candidate_count"] = candidate_count
                row["selected_target_count"] = len(targets)
                _populate_dim_audit_target_fields(row, candidate)
                row["actual_after"] = row.get("actual_before", "")
                row["outcome"] = "SKIP"
                row["reason"] = "existing-kept"
                audit.write_row(row)
            continue

        counters["write_attempted"] += 1
        if target_scope == "type":
            counters["type_writes"] += 1
        else:
            counters["instance_writes"] += 1
        if is_shared:
            counters["shared_writes"] += 1
        else:
            counters["family_writes"] += 1

        if logger is not None:
            logger.info(
                "Element {0} | {1} | Kind={2} | Scope={3} | Storage={4} | Old='{5}' | New='{6}'".format(
                    _element_id_text(element),
                    _safe_str(getattr(getattr(param, "Definition", None), "Name", "")),
                    _parameter_kind_label(is_shared),
                    target_scope,
                    _safe_str(getattr(param, "StorageType", "")),
                    current_value,
                    incoming_value,
                )
            )

        ok, err = _write_stringish(param, incoming_value)
        if ok:
            counters["pending_verify"] += 1
            audit_row = None
            if audit is not None:
                audit_row = _build_dim_audit_row(audit_context, element, target_param_name, incoming_value)
                audit_row["candidate_count"] = candidate_count
                audit_row["selected_target_count"] = len(targets)
                _populate_dim_audit_target_fields(audit_row, candidate)
            pending_writes.append(
                {
                    "element_id_value": element.Id.IntegerValue,
                    "param_name": target_param_name,
                    "incoming_value": incoming_value,
                    "target_scope": target_scope,
                    "param_id": candidate["id"],
                    "is_shared": is_shared,
                    "audit_row": audit_row,
                }
            )
            written_count += 1
            messages.append("{0}:{1}:pending-verify".format(target_scope, candidate["id"]))
            continue

        counters["failed"] += 1
        failed_count += 1
        messages.append("{0}:{1}:{2}".format(target_scope, candidate["id"], err or "write failed"))
        if logger is not None:
            logger.error(
                "Element {0} | Category={1} | Family={2} | Parameter={3} | Kind={4} | Scope={5} | Reason={6}".format(
                    _element_id_text(element),
                    _safe_str(getattr(getattr(element, "Category", None), "Name", "")),
                    _get_element_name(_get_type_element(element) or element),
                    target_param_name,
                    _parameter_kind_label(is_shared),
                    target_scope,
                    err or "write failed",
                )
            )
        if audit is not None:
            row = _build_dim_audit_row(audit_context, element, target_param_name, incoming_value)
            row["candidate_count"] = candidate_count
            row["selected_target_count"] = len(targets)
            _populate_dim_audit_target_fields(row, candidate)
            row["actual_after"] = _parameter_text(param)
            row["outcome"] = "FAIL"
            row["reason"] = err or "write failed"
            audit.write_row(row)

    if failed_count and written_count:
        return "PARTIAL", "; ".join(messages)
    if failed_count:
        return "FAIL", "; ".join(messages)
    if written_count:
        return "OK", "pending verification ({0} target(s))".format(written_count)

    return "SKIP", "; ".join(messages) if messages else resolution_reason


def _verify_pending_writes(pending_writes, counters, sample_fails, logger=None, audit=None):
    for pending in pending_writes:
        element_id_value = pending.get("element_id_value")
        param_name = pending.get("param_name")
        incoming_value = pending.get("incoming_value")
        target_scope = pending.get("target_scope")
        param_id = pending.get("param_id")
        is_shared = pending.get("is_shared")
        try:
            element = doc.GetElement(ElementId(int(element_id_value)))
        except Exception:
            element = None

        owner = _get_type_element(element) if target_scope == "type" else element
        param = _find_candidate_param_by_id(owner, param_name, param_id)
        if _parameter_value_matches_incoming(param, incoming_value):
            counters["written"] += 1
            if audit is not None:
                row = dict(pending.get("audit_row") or {})
                row["actual_after"] = _parameter_text(param)
                row["outcome"] = "OK"
                row["reason"] = "verified"
                audit.write_row(row)
            if logger is not None:
                logger.info(
                    "VERIFY OK | Element {0} | Parameter={1} | Kind={2} | Scope={3} | Id={4} | Actual='{5}'".format(
                        element_id_value,
                        param_name,
                        _parameter_kind_label(is_shared),
                        target_scope or "unknown",
                        param_id,
                        _parameter_text(param),
                    )
                )
            continue

        counters["failed"] += 1
        counters["verification_failed"] += 1
        if audit is not None:
            row = dict(pending.get("audit_row") or {})
            row["actual_after"] = _parameter_text(param)
            row["outcome"] = "FAIL"
            row["reason"] = "verification mismatch after commit"
            audit.write_row(row)
        if logger is not None:
            logger.error(
                "VERIFY FAIL | Element {0} | Parameter={1} | Kind={2} | Scope={3} | Id={4} | Actual='{5}' | Expected='{6}'".format(
                    element_id_value,
                    param_name,
                    _parameter_kind_label(is_shared),
                    target_scope or "unknown",
                    param_id,
                    _parameter_text(param),
                    incoming_value,
                )
            )
            _dump_element_parameters_for_debug(doc.GetElement(ElementId(int(element_id_value))), logger)
        if len(sample_fails) < 12:
            sample_fails.append(
                "Element {0} | {1} | verification mismatch after commit ({2})".format(
                    element_id_value,
                    param_name,
                    target_scope or "unknown",
                )
            )


def _retry_pending_writes_by_element(pending_writes, counters, sample_fails, logger=None, audit=None):
    grouped = []
    groups = {}

    for pending in pending_writes:
        element_id_value = pending.get("element_id_value")
        if element_id_value not in groups:
            groups[element_id_value] = []
            grouped.append((element_id_value, groups[element_id_value]))
        groups[element_id_value].append(pending)

    if logger is not None:
        logger.warning(
            "Retrying rolled-back chunk one element at a time | Elements={0} | PendingWrites={1}".format(
                len(grouped),
                len(pending_writes),
            )
        )

    for element_id_value, items in grouped:
        tx = Transaction(doc, "{0} Retry {1}".format(TOOL_TITLE, element_id_value))
        failure_capture = _attach_transaction_failure_capture(tx)
        tx.Start()

        try:
            failure_reason = ""
            for pending in items:
                ok, err = _apply_pending_write(pending)
                if ok:
                    continue
                failure_reason = err or "retry write failed"
                break

            if failure_reason:
                try:
                    tx.RollBack()
                except Exception:
                    pass
                counters["failed"] += len(items)
                if logger is not None:
                    logger.error("Element retry failed before commit | Element={0} | Reason={1}".format(element_id_value, failure_reason))
                _mark_pending_audit_rows(items, "FAIL", "retry write failed ({0})".format(failure_reason), audit)
                if len(sample_fails) < 12:
                    sample_fails.append("Element {0} | retry write failed ({1})".format(element_id_value, failure_reason))
                continue

            status = tx.Commit()
            if logger is not None:
                logger.info("Retry Transaction Status | Element={0} | Status={1}".format(element_id_value, status))
            if status != TransactionStatus.Committed:
                counters["failed"] += len(items)
                detail_messages = _summarize_transaction_failure_messages(failure_capture, 1)
                reason = "retry transaction not committed ({0})".format(status)
                if detail_messages:
                    reason = "{0} | {1}".format(reason, detail_messages[0])
                if logger is not None:
                    logger.error("Element retry transaction failed | Element={0} | Reason={1}".format(element_id_value, reason))
                    _log_transaction_failure_messages(logger, "RETRY FAILURE", failure_capture)
                _mark_pending_audit_rows(items, "FAIL", reason, audit)
                if len(sample_fails) < 12:
                    sample_fails.append("Element {0} | {1}".format(element_id_value, reason))
                continue

            _verify_pending_writes(items, counters, sample_fails, logger, audit)
        except Exception as ex:
            try:
                tx.RollBack()
            except Exception:
                pass
            counters["failed"] += len(items)
            reason = "retry exception ({0})".format(_safe_str(ex))
            if logger is not None:
                logger.error("Element retry exception | Element={0} | Reason={1}".format(element_id_value, reason))
            _mark_pending_audit_rows(items, "FAIL", reason, audit)
            if len(sample_fails) < 12:
                sample_fails.append("Element {0} | {1}".format(element_id_value, reason))


def run():
    active_view = doc.ActiveView
    if active_view is None:
        forms.alert("No active view found.", title=TOOL_TITLE)
        return

    view_mode = _view_label(active_view)
    if view_mode == "schedule" and not _schedule_is_supported(active_view):
        forms.alert(
            "Active schedule is not supported. Open a non-template, non-key schedule or use a model view.",
            title=TOOL_TITLE,
        )
        return

    selected_elements = _get_selected_mep_elements()
    run_settings = _prompt_run_settings(len(selected_elements), _safe_str(getattr(active_view, "Name", "")))
    if run_settings is None:
        forms.alert("Operation cancelled before processing.", title=TOOL_TITLE)
        return

    write_mode = run_settings.get("write_mode", "skip")
    parameter_target_mode = run_settings.get("parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE)
    worksharing_mode = run_settings.get("worksharing_mode", "all")
    processing_scope = run_settings.get("processing_scope", "view")
    export_audit = bool(run_settings.get("export_audit", True))
    fast_mode = bool(run_settings.get("fast_mode", False))

    logger = RunLogger(os.path.dirname(__file__), TOOL_TITLE)
    audit = RunAuditExporter(logger.logs_dir, logger.safe_title, logger.stamp, export_audit, DIM_AUDIT_COLUMNS)
    run_audit_context = {
        "run_stamp": logger.stamp,
        "tool": TOOL_TITLE,
        "view_name": _safe_str(getattr(active_view, "Name", "")),
        "view_mode": view_mode,
        "processing_scope": processing_scope,
        "write_mode": write_mode,
        "parameter_target_mode": parameter_target_mode,
        "worksharing_mode": worksharing_mode,
    }
    logger.write("START | Tool={0}".format(TOOL_TITLE))
    logger.write("START | View={0}".format(_safe_str(getattr(active_view, "Name", ""))))
    logger.write("START | ViewId={0}".format(_safe_str(getattr(getattr(active_view, "Id", None), "IntegerValue", ""))))
    logger.write("START | ViewMode={0}".format(view_mode))
    logger.write("START | WriteMode={0}".format(write_mode))
    logger.write("START | ParameterTargetMode={0}".format(parameter_target_mode))
    logger.write("START | WorksharingMode={0}".format(worksharing_mode))
    logger.write("START | ProcessingScope={0}".format(processing_scope))
    logger.write("START | SelectedInputCount={0}".format(len(selected_elements)))
    logger.write("START | ExportAudit={0}".format(export_audit))
    if export_audit and audit.file_path:
        logger.write("START | AuditFile={0}".format(audit.file_path))

    monitor = None
    try:
        if not fast_mode:
            monitor = LiveMonitorWindow("MonitorWindow.xaml")
            monitor.Show()
            monitor.Activate()
            _monitor_set_status(monitor, "Initializing {0}...".format(TOOL_TITLE))
            _monitor_add_item(monitor, "Preparing run context...")
            _pump_ui()
        else:
            logger.write("START | FastMode=True | Live monitor and per-element runtime logging disabled")
    except Exception as ex:
        monitor = None
        forms.alert(
            "Live monitor window could not be opened.\n"
            "Processing will continue with progress bar only.\n\n{0}".format(_safe_str(ex)),
            title=TOOL_TITLE,
        )

    if processing_scope == "selection":
        elements = selected_elements
    elif view_mode == "schedule":
        elements = _collect_active_schedule_elements(active_view)
    else:
        elements = _collect_active_model_view_elements(active_view)

    if not elements:
        logger.write("END | No eligible MEP elements found")
        logger.close()
        if processing_scope == "selection":
            forms.alert("No eligible MEP elements were found in the current selection.", title=TOOL_TITLE)
        else:
            forms.alert("No eligible MEP elements were found in the active {0} view.".format(view_mode), title=TOOL_TITLE)
        return

    counters = {
        "processed": 0,
        "write_attempted": 0,
        "written": 0,
        "failed": 0,
        "verification_failed": 0,
        "missing_target": 0,
        "skip_existing": 0,
        "skip_no_source": 0,
        "no_size": 0,
        "skip_worksharing": 0,
        "type_writes": 0,
        "instance_writes": 0,
        "read_only": 0,
        "duplicate_parameter": 0,
        "skip_target_mode": 0,
        "shared_writes": 0,
        "family_writes": 0,
        "pending_verify": 0,
        "transaction_failed": 0,
    }

    total = len(elements)
    batch_size = FAST_MODE_MAX_ELEMENTS
    logger.write("READY | Full MEP set accepted: {0} elements in batches of {1}".format(total, batch_size))
    chunk_size = 30 if fast_mode else 120
    cancelled = False
    sample_fails = []
    owner_status_cache = {}

    _monitor_set_status(
        monitor,
        "Ready. View mode: {0} | Elements: {1}".format(view_mode, total),
    )
    _monitor_add_item(
        monitor,
        "Starting processing in {0} scope / {1} view mode with write mode: {2} | target mode: {3} | worksharing: {4} | log: {5}".format(
            processing_scope,
            view_mode,
            write_mode,
            parameter_target_mode,
            worksharing_mode,
            logger.file_path,
        ),
    )
    logger.write("READY | Elements={0} | Log={1}".format(total, logger.file_path))
    logger.write("READY | CategoryBreakdown={0}".format(_summarize_category_counts(elements)))
    logger.write("READY | EligibleElementIds={0}".format(_summarize_element_ids(elements)))
    _pump_ui()

    with forms.ProgressBar(title="NWB dim auto-fill {value}/{max_value}", cancellable=True) as pb:
        pb.update_progress(0, total)

        index = 0
        while index < total:
            if pb.cancelled or (monitor is not None and monitor.cancel_requested):
                cancelled = True
                break

            chunk = elements[index:index + chunk_size]
            tx = Transaction(doc, TOOL_TITLE)
            failure_capture = _attach_transaction_failure_capture(tx)
            tx.Start()
            pending_writes = []
            processed_in_chunk = 0
            if monitor is not None:
                monitor.set_busy(True)

            try:
                for element in chunk:
                    if pb.cancelled or (monitor is not None and monitor.cancel_requested):
                        cancelled = True
                        break

                    processed_in_chunk += 1
                    counters["processed"] += 1
                    eid = _element_id_text(element)

                    allowed, ws_reason = _is_element_allowed_by_worksharing(element, worksharing_mode, owner_status_cache)
                    if not allowed:
                        counters["skip_worksharing"] += 1
                        _monitor_add_item(monitor, "[{0}] Skipped by worksharing rule: {1}".format(counters["processed"], ws_reason), eid)
                        logger.write("SKIP | Element {0} | Worksharing | {1}".format(eid, ws_reason))
                        if audit is not None:
                            row = _build_dim_audit_row(run_audit_context, element, "", "")
                            row["outcome"] = "SKIP"
                            row["reason"] = "worksharing: {0}".format(ws_reason)
                            audit.write_row(row)
                        continue

                    size_text, parsed = _get_element_size_data(element)
                    audit_context = _build_dim_audit_context(run_audit_context, element, size_text, parsed)
                    if _is_blank(size_text):
                        counters["no_size"] += 1
                        _monitor_add_item(monitor, "[{0}] Skipped | Size empty".format(counters["processed"]), eid)
                        logger.write(
                            "SKIP | Element {0} | Size empty | Cat={1} | Family={2} | Type={3} | Level={4} | Name={5}".format(
                                eid,
                                _get_category_name(element),
                                _get_family_name(element),
                                _get_type_name(element),
                                _get_level_name(element),
                                _get_element_name(element),
                            )
                        )
                        if audit is not None:
                            row = _build_dim_audit_row(audit_context, element, "", "")
                            row["outcome"] = "SKIP"
                            row["reason"] = "size empty"
                            audit.write_row(row)
                        if counters["processed"] % 10 == 0 or counters["processed"] == total:
                            pb.update_progress(counters["processed"], total)
                        continue

                    target_values = {
                        "NWB_DimWidth": parsed.width,
                        "NWB_DimHeight": parsed.height,
                        "NWB_DimDiameter": parsed.diameter,
                    }

                    statuses = []
                    for param_name in TARGET_PARAMETERS:
                        incoming = target_values.get(param_name, "")
                        status, message = _set_if_needed(
                            element,
                            param_name,
                            incoming,
                            write_mode,
                            parameter_target_mode,
                            counters,
                            pending_writes,
                            logger,
                            audit,
                            audit_context,
                        )
                        statuses.append("{0}:{1}".format(param_name, status))
                        if status in ("FAIL", "PARTIAL") and len(sample_fails) < 12:
                            sample_fails.append("Element {0} | {1} | {2}".format(eid, param_name, message))

                    if not fast_mode:
                        _log_size_runtime_context(logger, element, eid, size_text, parsed, target_values)

                    if not fast_mode:
                        logger.write(
                        "ELEMENT | {0} | Size='{1}' | Width='{2}' | Height='{3}' | Diameter='{4}' | Statuses={5}".format(
                            eid,
                            size_text,
                            _safe_str(parsed.width),
                            _safe_str(parsed.height),
                            _safe_str(parsed.diameter),
                            "; ".join(statuses),
                        )
                    )

                    if counters["processed"] <= 20 or (counters["processed"] % 10 == 0):
                        _monitor_add_item(
                            monitor,
                            "[{0}] Size='{1}' | {2}".format(counters["processed"], size_text, " ".join(statuses)),
                            eid,
                        )

                    if counters["processed"] % 10 == 0 or counters["processed"] == total:
                        pb.update_progress(counters["processed"], total)
                        _monitor_set_status(
                            monitor,
                            "Processed {0}/{1} | Written: {2} | Skipped existing: {3} | Missing target: {4} | No parsed value: {5} | Worksharing skip: {6} | Failed: {7}".format(
                                counters["processed"],
                                total,
                                counters["written"],
                                counters["skip_existing"],
                                counters["missing_target"],
                                counters["skip_no_source"],
                                counters["skip_worksharing"],
                                counters["failed"],
                            )
                        )

                    if counters["processed"] % 10 == 0:
                        _pump_ui()

                    if monitor is not None and monitor.has_pending_focus():
                        logger.info("Pausing chunk after element {0} to honor pending focus request.".format(eid))
                        break

                status = tx.Commit()
                logger.info("Transaction Status = {0}".format(status))
                if status != TransactionStatus.Committed:
                    counters["transaction_failed"] += 1
                    logger.error("Transaction failed: {0}".format(status))
                    _log_transaction_failure_messages(logger, "TRANSACTION FAILURE", failure_capture)
                    if pending_writes:
                        _retry_pending_writes_by_element(pending_writes, counters, sample_fails, logger, audit)
                    else:
                        _mark_pending_audit_rows(pending_writes, "FAIL", "transaction not committed ({0})".format(status), audit)
                    if len(sample_fails) < 12:
                        sample_fails.append("Transaction failed with status {0}".format(status))
                else:
                    if doc.IsModifiable:
                        try:
                            doc.Regenerate()
                        except Exception as ex:
                            logger.warning("Regenerate after commit failed: {0}".format(_safe_str(ex)))
                    logger.write("CHUNK COMMIT | Status={0} | Processed={1}".format(status, counters["processed"]))
                    _verify_pending_writes(pending_writes, counters, sample_fails, logger, audit)
                counters["pending_verify"] = max(0, counters["pending_verify"] - len(pending_writes))
            except Exception:
                try:
                    tx.RollBack()
                except Exception:
                    pass
                _mark_pending_audit_rows(pending_writes, "FAIL", "transaction rolled back by exception", audit)
                logger.write("CHUNK ROLLBACK | Processed={0}".format(counters["processed"]))
                raise
            finally:
                if monitor is not None:
                    monitor.set_busy(False)
                    monitor.drain_pending_focus()
                logger.flush()
                audit.flush()

            _pump_ui()
            if index // batch_size != (index + processed_in_chunk) // batch_size:
                logger.write("BATCH COMPLETE | Processed={0}/{1}".format(index + processed_in_chunk, total))
            if cancelled:
                break
            index += processed_in_chunk

    if cancelled:
        _monitor_set_status(monitor, "Cancelled. Kept already committed chunks.")
    else:
        _monitor_set_status(monitor, "Completed.")

    summary = [
        "{0} finished.".format(TOOL_TITLE),
        "View: {0} ({1})".format(_safe_str(getattr(active_view, "Name", "")), view_mode),
        "Processing scope: {0}".format(processing_scope),
        "Write mode: {0}".format(write_mode),
        "Parameter target mode: {0}".format(parameter_target_mode),
        "Worksharing mode: {0}".format(worksharing_mode),
        "Log file: {0}".format(logger.file_path),
        "Audit file: {0}".format(audit.file_path if audit.file_path else "disabled"),
        "Processed elements: {0}/{1}".format(counters["processed"], total),
        "Writes attempted: {0}".format(counters["write_attempted"]),
        "Parameter writes: {0}".format(counters["written"]),
        "Writes skipped: {0}".format(counters["skip_existing"] + counters["skip_no_source"] + counters["missing_target"] + counters["skip_worksharing"] + counters["skip_target_mode"]),
        "Skipped existing: {0}".format(counters["skip_existing"]),
        "Skipped no parsed value: {0}".format(counters["skip_no_source"]),
        "Missing target parameter: {0}".format(counters["missing_target"]),
        "Skipped by target mode policy: {0}".format(counters["skip_target_mode"]),
        "Skipped by worksharing rule: {0}".format(counters["skip_worksharing"]),
        "Elements with empty Size: {0}".format(counters["no_size"]),
        "Failed writes: {0}".format(counters["failed"]),
        "Verification failed: {0}".format(counters["verification_failed"]),
        "Shared-parameter writes: {0}".format(counters["shared_writes"]),
        "Family/non-shared writes: {0}".format(counters["family_writes"]),
        "Type writes: {0}".format(counters["type_writes"]),
        "Instance writes: {0}".format(counters["instance_writes"]),
        "Read only: {0}".format(counters["read_only"]),
        "Duplicate parameter: {0}".format(counters["duplicate_parameter"]),
        "Transaction failed: {0}".format(counters["transaction_failed"]),
        "Cancelled: {0}".format("Yes" if cancelled else "No"),
    ]

    if sample_fails:
        summary.append("")
        summary.append("Sample failures:")
        summary.extend(["- {0}".format(item) for item in sample_fails])

    logger.write(
        "END | Attempted={0} | Written={1} | Skipped={2} | VerifyFailed={3} | SharedWrites={4} | FamilyWrites={5} | TypeWrites={6} | InstanceWrites={7} | ReadOnly={8} | MissingTarget={9} | SkipTargetMode={10} | DuplicateParameter={11} | SkipWorksharing={12} | Failed={13} | Cancelled={14}".format(
            counters["write_attempted"],
            counters["written"],
            counters["skip_existing"] + counters["skip_no_source"] + counters["missing_target"] + counters["skip_worksharing"] + counters["skip_target_mode"],
            counters["verification_failed"],
            counters["shared_writes"],
            counters["family_writes"],
            counters["type_writes"],
            counters["instance_writes"],
            counters["read_only"],
            counters["missing_target"],
            counters["skip_target_mode"],
            counters["duplicate_parameter"],
            counters["skip_worksharing"],
            counters["failed"],
            cancelled,
        )
    )
    logger.close()
    audit.close()

    forms.alert("\n".join(summary), title=TOOL_TITLE)


if __name__ == "__main__":
    try:
        run()
    except Exception as ex:
        forms.alert("{0} failed.\n\n{1}".format(TOOL_TITLE, _safe_str(ex)), title=TOOL_TITLE)