# coding: utf8
from __future__ import print_function

import os
import re
import time

import clr

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Application as WinFormsApplication
from System import Type, Activator
from System.Collections.Generic import List
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    IFailuresPreprocessor,
    RevitLinkInstance,
    StorageType,
    Transaction,
    TransactionStatus,
    FailureProcessingResult,
    ViewSchedule,
    XYZ,
)
from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow
from pyrevitmep.meputils import NoConnectorManagerError, get_connector_manager


doc = revit.doc
uidoc = revit.uidoc
config = script.get_config()


TOOL_TITLE = "NWB_PARAMETERS AutoFill"
WORKBOOK_FILE_NAME = "NWB-WAL-GEN-DE-REG-0003.xlsx"

WBH_SCOPE_BOX_WBS01_RULES = [
    ("Companion Building", "CB"),
    ("Hospital", "HO"),
    ("Bridge", "LB"),
    ("Link Bridge", "LB"),
    ("Hub Building", "HB"),
    ("General", "GE"),
    ("External Area", "EA"),
    ("External Areas", "EA"),
    ("ZONE 222 - LINK BRIDGE", "LB"),
    ("ZONE 223 - LINK BRIDGE", "LB"),
    ("ZONE 251 - COMPANION", "CB"),
    ("ZONE 152 - HOSPITAL", "HO"),
    ("ZONE 153 - HOSPITAL", "HO"),
    ("ZONE 154N - EXTERNAL", "EA"),
    ("ZONE 154S - EXTERNAL", "EA"),
]

TARGET_PARAMETERS = [
    "NWB_AssetTypeCode",
    "NWB_AssetID",
    "NWB_AssetType",
    "NWB_Discipline",
    "NWB_Category",
    "NWB_Department",
    "NWB_SubDepartment",
    "NWB_Material",
    "NWB_System",
    "NWB_UniclassCode",
    "NWB_UniclassDescription",
    "NWB_DesignPkg",
    "NWB_BuildingPermitPkg",
    "NWB_WBS00",
    "NWB_WBS01",
    "NWB_WBS02",
    "NWB_WBS03",
    "NWB_WBS04",
    "NWB_WBS05",
    "NWB_WBSCode",
]

PARAMETER_TARGET_MODES = [
    "SharedOnly",
    "FamilyOnly",
    "Both",
    "FirstMatch",
]

DEFAULT_PARAMETER_TARGET_MODE = "SharedOnly"

WRITE_MODE_CHOICES = [
    ("Keep correct values and fix wrong ones", "smart"),
    ("Force rewrite mapped NWB values", "force"),
]

PARAMETER_TARGET_MODE_CHOICES = [
    ("{0} (recommended)".format(DEFAULT_PARAMETER_TARGET_MODE), DEFAULT_PARAMETER_TARGET_MODE),
    ("Both (shared + legacy duplicates)", "Both"),
    ("FamilyOnly / non-shared only", "FamilyOnly"),
    ("FirstMatch (diagnostic fallback)", "FirstMatch"),
]

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

    def _focus_element_now(self, element_id):
        try:
            eid_int = int(element_id)
        except Exception:
            return False, "Invalid element ID: {0}".format(element_id)

        try:
            element = doc.GetElement(ElementId(eid_int))
        except Exception:
            element = None

        if element is None:
            return False, "Element {0} could not be found.".format(eid_int)

        try:
            element_id_obj = ElementId(eid_int)
            selection_ids = List[ElementId]()
            selection_ids.Add(element_id_obj)
        except Exception as ex:
            return False, "Could not select element {0}: {1}".format(eid_int, _safe_str(ex))

        try:
            uidoc.Selection.SetElementIds(selection_ids)
        except Exception as ex:
            return False, "Could not select element {0}: {1}".format(eid_int, _safe_str(ex))

        centered = False
        center_error = ""
        try:
            uidoc.ShowElements(element_id_obj)
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

            # Best effort to bring the element into view even when Revit rejects ShowElements.
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

    def focus_element(self, element_id):
        if self._is_busy:
            self._pending_focus_element_id = element_id
            return True, "Element {0} queued. It will be centered at the next safe checkpoint.".format(element_id)
        return self._focus_element_now(element_id)

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
        eid = self._get_selected_element_id()
        if eid is None:
            return
        ok, msg = self.focus_element(eid)
        self.set_status(msg)

    def list_item_double_click(self, sender, e):
        eid = self._get_selected_element_id()
        if eid is None:
            self.set_status("Select a processed row with a valid ID first.")
            return
        ok, msg = self.focus_element(eid)
        self.set_status(msg)

    def show_selected_click(self, sender, e):
        eid = self._get_selected_element_id()
        if eid is None:
            self.set_status("Select a processed row with a valid ID first.")
            return
        ok, msg = self.focus_element(eid)
        self.set_status(msg)

    def center_selected_click(self, sender, e):
        eid = self._get_selected_element_id()
        if eid is None:
            self.set_status("Select a processed row with a valid ID first.")
            return
        ok, msg = self.focus_element(eid)
        self.set_status(msg)

    def clear_items_click(self, sender, e):
        self.lstItems.Items.Clear()
        self._row_ids = []
        self.set_status("Processed-item list cleared.")

    def cancel_click(self, sender, e):
        self.cancel_requested = True
        self.set_status("Cancel requested. Finishing current safe chunk and keeping committed changes...")

    def close_click(self, sender, e):
        self.Hide()


class WorkbookData(object):
    def __init__(self):
        self.asset_rows = []
        self.system_keywords = set()
        self.discipline_desc_to_code = {}
        self.discipline_code_to_desc = {}
        self.work_package_codes = set()
        self.work_packages = []
        self.building_permit_items = set()
        self.building_permit_by_key = {}
        self.wbs_maps = {
            "WBS00": {"code_to_desc": {}, "desc_to_code": {}, "default": ""},
            "WBS01": {"code_to_desc": {}, "desc_to_code": {}, "default": ""},
            "WBS02": {"code_to_desc": {}, "desc_to_code": {}, "default": ""},
            "WBS03": {"code_to_desc": {}, "desc_to_code": {}, "default": ""},
            "WBS04": {"code_to_desc": {}, "desc_to_code": {}, "default": ""},
            "WBS05": {"code_to_desc": {}, "desc_to_code": {}, "default": ""},
        }


_PROJECT_CONTEXT_CACHE = None
_RUN_OPTIONS = {}
_LINKED_ROOM_INDEX_CACHE = {}
_SCOPE_BOX_WBS01_CACHE = None
_PARAM_RESOLUTION_CACHE = {}  # Format: {(owner_id, param_name, target_mode): [candidates]}
_ASSET_MATCH_CACHE = {}  # Format: {(category_norm, family_norm, type_norm, name_norm, system_norm): (asset_row, score)}


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


def _collect_linked_room_choices():
    choices = [("Auto-detect from available architectural room links", "")]
    try:
        collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    except Exception:
        return choices

    seen = set()
    for link_inst in collector:
        if link_inst is None:
            continue
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue

        try:
            rooms = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
            has_rooms = False
            for room in rooms:
                if room is not None:
                    has_rooms = True
                    break
            if not has_rooms:
                continue
        except Exception:
            continue

        try:
            link_id = str(link_inst.Id.IntegerValue)
        except Exception:
            continue

        try:
            label = _safe_str(link_inst.Name).strip() or _safe_str(link_doc.Title).strip() or ("Link %s" % link_id)
        except Exception:
            label = "Link %s" % link_id

        if link_id in seen:
            continue
        seen.add(link_id)
        choices.append((label, link_id))

    return choices


def _load_run_settings(selected_count):
    settings = {
        "write_mode": getattr(config, "nwb_params_write_mode", "smart"),
        "parameter_target_mode": getattr(config, "nwb_params_parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE),
        "worksharing_mode": getattr(config, "nwb_params_worksharing_mode", "skip-other-users" if getattr(doc, "IsWorkshared", False) else "all"),
        "processing_scope": getattr(config, "nwb_params_processing_scope", "selection" if selected_count > 0 else "view"),
        "arch_room_link_id": getattr(config, "nwb_params_arch_room_link_id", ""),
        "export_audit": getattr(config, "nwb_params_export_audit", True),
        "site_override": "",
        "building_override": "",
        "design_pkg_override": getattr(config, "nwb_params_design_pkg_override", ""),
        "building_permit_override": getattr(config, "nwb_params_building_permit_override", ""),
    }

    valid_worksharing = [value for _, value in _build_worksharing_choices()]
    if settings["worksharing_mode"] not in valid_worksharing:
        settings["worksharing_mode"] = valid_worksharing[0]

    valid_scopes = [value for _, value in _build_processing_scope_choices(selected_count)]
    if settings["processing_scope"] not in valid_scopes:
        settings["processing_scope"] = valid_scopes[0]

    valid_write_modes = [value for _, value in WRITE_MODE_CHOICES]
    if settings["write_mode"] not in valid_write_modes:
        settings["write_mode"] = "smart"

    valid_target_modes = [value for _, value in PARAMETER_TARGET_MODE_CHOICES]
    if settings["parameter_target_mode"] not in valid_target_modes:
        settings["parameter_target_mode"] = DEFAULT_PARAMETER_TARGET_MODE

    return settings


def _save_run_settings(settings):
    try:
        config.nwb_params_write_mode = settings.get("write_mode", "smart")
        config.nwb_params_parameter_target_mode = settings.get("parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE)
        config.nwb_params_worksharing_mode = settings.get("worksharing_mode", "all")
        config.nwb_params_processing_scope = settings.get("processing_scope", "view")
        config.nwb_params_arch_room_link_id = settings.get("arch_room_link_id", "")
        config.nwb_params_export_audit = bool(settings.get("export_audit", True))
        config.nwb_params_site_override = ""
        config.nwb_params_building_override = ""
        config.nwb_params_design_pkg_override = settings.get("design_pkg_override", "")
        config.nwb_params_building_permit_override = settings.get("building_permit_override", "")
        script.save_config()
    except Exception:
        pass


def _set_run_options(settings):
    global _RUN_OPTIONS, _PROJECT_CONTEXT_CACHE
    _RUN_OPTIONS = dict(settings or {})
    _PROJECT_CONTEXT_CACHE = None


def _run_option_text(option_name):
    return _safe_str(_RUN_OPTIONS.get(option_name, "")).strip()


def _run_option_bool(option_name, default=False):
    value = _RUN_OPTIONS.get(option_name, default)
    if isinstance(value, bool):
        return value
    text = _safe_str(value).strip().lower()
    if text in ("1", "true", "yes", "on"):
        return True
    if text in ("0", "false", "no", "off"):
        return False
    return bool(value)


class RunSetupWindow(WPFWindow):
    def __init__(self, xaml_file_name, selected_count, active_view_name, defaults, project_snapshot):
        WPFWindow.__init__(self, xaml_file_name)
        self.result = None
        self._write_choices = list(WRITE_MODE_CHOICES)
        self._target_choices = list(PARAMETER_TARGET_MODE_CHOICES)
        self._worksharing_choices = _build_worksharing_choices()
        self._scope_choices = _build_processing_scope_choices(selected_count)
        self._linked_room_choices = _collect_linked_room_choices()

        self.txtActiveView.Text = _safe_str(active_view_name)
        self.txtSelectedCount.Text = str(max(0, int(selected_count)))
        self.txtDetectedProjectLocation.Text = self._display_value(project_snapshot.get("project_location", ""))
        self.txtDetectedBuildingName.Text = self._display_value(project_snapshot.get("building_name", ""))
        self.txtDetectedBuildingNumber.Text = self._display_value(project_snapshot.get("building_number", ""))
        self.txtDetectedProjectName.Text = self._display_value(project_snapshot.get("project_name", ""))
        self.txtDetectedProjectNumber.Text = self._display_value(project_snapshot.get("project_number", ""))

        self._bind_combo(self.cmbWriteMode, self._write_choices, defaults.get("write_mode", "smart"))
        self._bind_combo(self.cmbParameterTargetMode, self._target_choices, defaults.get("parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE))
        self._bind_combo(self.cmbWorksharingMode, self._worksharing_choices, defaults.get("worksharing_mode", "all"))
        self._bind_combo(self.cmbProcessingScope, self._scope_choices, defaults.get("processing_scope", "view"))
        self._bind_combo(self.cmbArchitecturalRoomLink, self._linked_room_choices, defaults.get("arch_room_link_id", ""))

        site_default = project_snapshot.get("project_location", "") or defaults.get("site_override", "")
        building_default = project_snapshot.get("building_name", "") or project_snapshot.get("building_number", "") or defaults.get("building_override", "")

        self.txtSiteOverride.Text = _safe_str(site_default)
        self.txtBuildingOverride.Text = _safe_str(building_default)
        self.txtDesignPkgOverride.Text = _safe_str(defaults.get("design_pkg_override", ""))
        self.txtBuildingPermitOverride.Text = _safe_str(defaults.get("building_permit_override", ""))
        self.chkExportAudit.IsChecked = bool(defaults.get("export_audit", True))

        if len(self._worksharing_choices) == 1:
            self.cmbWorksharingMode.IsEnabled = False

        if len(self._linked_room_choices) <= 1:
            self.cmbArchitecturalRoomLink.IsEnabled = False

    def _display_value(self, value):
        text = _safe_str(value).strip()
        return text or "-"

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
            "arch_room_link_id": self._selected_combo_value(self.cmbArchitecturalRoomLink, self._linked_room_choices),
            "export_audit": bool(self.chkExportAudit.IsChecked),
            "site_override": _safe_str(self.txtSiteOverride.Text).strip(),
            "building_override": _safe_str(self.txtBuildingOverride.Text).strip(),
            "design_pkg_override": _safe_str(self.txtDesignPkgOverride.Text).strip(),
            "building_permit_override": _safe_str(self.txtBuildingPermitOverride.Text).strip(),
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


PARAM_AUDIT_COLUMNS = [
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
    "system",
    "name",
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
    "asset_score",
    "asset_type_code",
    "derived_asset_id",
    "asset_id_rule",
    "asset_id_reason",
    "asset_type",
    "asset_category",
    "linked_room_source",
    "linked_room_link",
    "linked_room_number",
    "linked_room_name",
    "linked_room_department",
    "linked_room_subdepartment",
    "linked_room_building",
    "linked_room_site",
    "resolved_department",
    "resolved_subdepartment",
    "resolved_building",
    "resolved_site",
    "derived_wbs00",
    "derived_wbs01",
    "derived_wbs03",
    "derived_wbs05",
    "derived_wbs_code",
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


def _build_param_audit_context(run_context, element, debug, target_values):
    facts = (debug or {}).get("facts", {}) or {}
    context = dict(run_context or {})
    context.update(
        {
            "element_id": _element_id_text(element),
            "category": facts.get("category", "") or _get_category_name(element),
            "family": facts.get("family", "") or _get_family_name(element),
            "type": facts.get("type", "") or _get_type_name(element),
            "level": facts.get("level", "") or _get_level_name(element),
            "system": facts.get("system", "") or _get_system_classification(element),
            "name": facts.get("name", "") or _get_element_name(element),
            "asset_score": (debug or {}).get("asset_score", 0),
            "asset_type_code": (debug or {}).get("asset_type_code", ""),
            "derived_asset_id": _safe_str((target_values or {}).get("NWB_AssetID", "")),
            "asset_id_rule": (debug or {}).get("asset_id_rule", ""),
            "asset_id_reason": (debug or {}).get("asset_id_reason", ""),
            "asset_type": (debug or {}).get("asset_type", ""),
            "asset_category": (debug or {}).get("asset_category", ""),
            "linked_room_source": facts.get("linked_room_source", ""),
            "linked_room_link": facts.get("linked_room_link_name", ""),
            "linked_room_number": facts.get("linked_room_number", ""),
            "linked_room_name": facts.get("linked_room_name", ""),
            "linked_room_department": facts.get("linked_room_department", ""),
            "linked_room_subdepartment": facts.get("linked_room_subdepartment", ""),
            "linked_room_building": facts.get("linked_room_building", ""),
            "linked_room_site": facts.get("linked_room_site", ""),
            "resolved_department": facts.get("department", ""),
            "resolved_subdepartment": facts.get("subdepartment", ""),
            "resolved_building": facts.get("building_hint", ""),
            "resolved_site": facts.get("site_hint", ""),
            "derived_wbs00": _safe_str((target_values or {}).get("NWB_WBS00", "")),
            "derived_wbs01": _safe_str((target_values or {}).get("NWB_WBS01", "")),
            "derived_wbs03": _safe_str((target_values or {}).get("NWB_WBS03", "")),
            "derived_wbs05": _safe_str((target_values or {}).get("NWB_WBS05", "")),
            "derived_wbs_code": _safe_str((target_values or {}).get("NWB_WBSCode", "")),
        }
    )
    return context


def _build_param_audit_row(audit_context, element, target_param_name, incoming_value):
    row = dict(audit_context or {})
    row.update(
        {
            "element_id": row.get("element_id", _element_id_text(element)),
            "category": row.get("category", _get_category_name(element)),
            "family": row.get("family", _get_family_name(element)),
            "type": row.get("type", _get_type_name(element)),
            "level": row.get("level", _get_level_name(element)),
            "system": row.get("system", _get_system_classification(element)),
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


def _populate_param_audit_target_fields(row, candidate):
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


def _record_reason_stat(counters, element, param_name, reason):
    stats = counters.get("reason_stats")
    if stats is None:
        return

    key = (
        _get_category_name(element) or "<No Category>",
        _safe_str(param_name).strip() or "<Element>",
        _safe_str(reason).strip() or "<None>",
    )
    stats[key] = int(stats.get(key, 0)) + 1


def _top_reason_summary_lines(reason_stats, max_items=12):
    if not reason_stats:
        return []

    items = sorted(reason_stats.items(), key=lambda item: (-int(item[1]), item[0]))[:max_items]
    lines = []
    for (category, param_name, reason), count in items:
        lines.append("Count={0} | Category={1} | Parameter={2} | Reason={3}".format(count, category, param_name, reason))
    return lines


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


def _normalize_key(value):
    text = _safe_str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def _normalize_text(value):
    text = _safe_str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def _normalize_compare_text(value):
    return _normalize_text(value)


def _normalize_parameter_text(value):
    return _safe_str(value).strip().replace("\r", "").replace("\n", "")


def _is_blank(value):
    return not _safe_str(value).strip()


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


def _notify_status(callback, text):
    if callback is None:
        return
    try:
        callback(text)
    except Exception:
        pass


def _is_placeholder_text(value):
    text = _normalize_text(value)
    if not text:
        return True

    placeholders = [
        "refer to",
        "drofus",
        "from element geometry",
        "not applicable",
        "leave it empty",
        "assettypecode-building-level-elementid",
        "site-building-level",
        "material description",
    ]
    for phrase in placeholders:
        if phrase in text:
            return True
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
    return True


def _collect_active_schedule_elements(schedule_view):
    elements = []
    try:
        for element in FilteredElementCollector(doc, schedule_view.Id).WhereElementIsNotElementType().ToElements():
            if element is None:
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
                    eid = element.Id.IntegerValue
                except Exception:
                    continue
                if eid in seen:
                    continue
                seen.add(eid)
                elements.append(element)
        except Exception:
            continue

    return elements


def _find_param(element, name):
    if element is None or _is_blank(name):
        return None

    try:
        p = element.LookupParameter(name)
        if p:
            return p
    except Exception:
        pass

    try:
        params = element.GetParameters(name)
        if params:
            return params[0]
    except Exception:
        pass

    try:
        for p in element.Parameters:
            try:
                if p.Definition and p.Definition.Name == name:
                    return p
            except Exception:
                continue
    except Exception:
        pass

    return None


def _iter_param_owners(element):
    owners = []
    seen = set()

    for scope_name, owner in [("instance", element), ("type", getattr(element, "Symbol", None))]:
        if owner is None:
            continue
        try:
            owner_id = owner.Id.IntegerValue
        except Exception:
            owner_id = id(owner)
        if owner_id in seen:
            continue
        seen.add(owner_id)
        owners.append((scope_name, owner))

    type_el = _get_type_element(element)
    if type_el is not None:
        try:
            owner_id = type_el.Id.IntegerValue
        except Exception:
            owner_id = id(type_el)
        if owner_id not in seen:
            owners.append(("type", type_el))

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
    scope_priority = 0 if candidate["scope"] == "instance" else 1
    shared_priority = 0 if candidate["is_shared"] else 1
    return (
        scope_priority,
        shared_priority,
        _safe_str(getattr(candidate["param"], "StorageType", "")),
        candidate["id"],
    )


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
    # OPTIMIZATION (Phase 2.1): Check parameter resolution cache first
    global _PARAM_RESOLUTION_CACHE
    owner = _get_type_element(element) if element else None
    owner_id = int(owner.Id.IntegerValue) if owner else None
    cache_key = (owner_id, name, target_mode)
    
    if cache_key in _PARAM_RESOLUTION_CACHE:
        cached = _PARAM_RESOLUTION_CACHE[cache_key]
        if isinstance(cached, tuple) and len(cached) == 3:
            return cached  # Return cached (candidates, mode_status, all_candidates)
    
    candidates = _collect_param_candidates(element, name, logger)
    if not candidates:
        result = ([], "missing", [])
        if owner_id:
            _PARAM_RESOLUTION_CACHE[cache_key] = result
        return result

    writable = [candidate for candidate in candidates if not candidate["param"].IsReadOnly]
    if not writable:
        result = ([], "readonly", candidates)
        if owner_id:
            _PARAM_RESOLUTION_CACHE[cache_key] = result
        return result

    shared = [candidate for candidate in writable if candidate["is_shared"]]
    nonshared = [candidate for candidate in writable if not candidate["is_shared"]]

    if target_mode == "Both":
        result = (writable, "both", candidates)
        if owner_id:
            _PARAM_RESOLUTION_CACHE[cache_key] = result
        return result

    if target_mode == "FamilyOnly":
        if nonshared:
            result = (nonshared, "family-only", candidates)
            if owner_id:
                _PARAM_RESOLUTION_CACHE[cache_key] = result
            return result
        result = ([], "no-family", candidates)
        if owner_id:
            _PARAM_RESOLUTION_CACHE[cache_key] = result
        return result

    if target_mode == "FirstMatch":
        preferred = shared[0] if shared else writable[0]
        result = ([preferred], "first-match", candidates)
        if owner_id:
            _PARAM_RESOLUTION_CACHE[cache_key] = result
        return result

    if shared:
        result = (shared, "shared-only", candidates)
        if owner_id:
            _PARAM_RESOLUTION_CACHE[cache_key] = result
        return result
    if nonshared:
        result = ([], "no-shared", candidates)
        if owner_id:
            _PARAM_RESOLUTION_CACHE[cache_key] = result
        return result

    result = ([], "missing", candidates)
    if owner_id:
        _PARAM_RESOLUTION_CACHE[cache_key] = result
    return result


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
            return _normalize_parameter_text(param.AsValueString()) == expected_text

    if storage_type == StorageType.Double:
        try:
            return abs(float(param.AsDouble()) - float(_parse_numeric_text(incoming_text))) <= 1e-6
        except Exception:
            return _normalize_parameter_text(param.AsValueString()) == expected_text

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
        # OPTIMIZATION (Phase 1): Skip per-write regenerate and verification
        # Full verification happens at chunk-level post-commit via _verify_pending_writes()
        if not set_ok:
            return False, "Set() returned False"
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
            "Keep correct values and fix wrong ones",
            "Force rewrite mapped NWB values",
        ],
        message="Choose write mode for NWB_PARAMETERS values:",
    )
    if not choice:
        return None
    return "smart" if "Keep correct" in _safe_str(choice) else "force"


def _choose_parameter_target_mode():
    choice = forms.CommandSwitchWindow.show(
        [
            "{0} (recommended)".format(DEFAULT_PARAMETER_TARGET_MODE),
            "Both (shared + legacy duplicates)",
            "FamilyOnly / non-shared only",
            "FirstMatch (diagnostic fallback)",
        ],
        message="Choose parameter target mode for duplicate NWB parameter names:",
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
    project_snapshot = _get_project_info_snapshot()

    try:
        dialog = RunSetupWindow("RunOptionsWindow.xaml", selected_count, active_view_name, defaults, project_snapshot)
        dialog.ShowDialog()
    except Exception as ex:
        forms.alert(
            "The NWB parameters setup window could not be opened.\n\n{0}".format(_safe_str(ex)),
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


def _missing_source_reason(target_param_name, audit_context):
    if target_param_name == "NWB_AssetID":
        reason = _safe_str((audit_context or {}).get("asset_id_reason", "")).strip()
        return reason or "Asset Master NWB_AssetID is blank"

    if target_param_name == "NWB_Material":
        return "workbook material blank and model material missing"

    if target_param_name == "NWB_Department":
        room_source = _safe_str((audit_context or {}).get("linked_room_source", "")).strip()
        if room_source:
            return "linked/direct room department unavailable ({0})".format(room_source)
        return "linked/direct room department unavailable"

    if target_param_name == "NWB_SubDepartment":
        room_source = _safe_str((audit_context or {}).get("linked_room_source", "")).strip()
        if room_source:
            return "linked/direct room subdepartment unavailable ({0})".format(room_source)
        return "linked/direct room subdepartment unavailable"

    return "no source value"


def _set_if_needed(element, target_param_name, incoming_value, mode, parameter_target_mode, counters, pending_writes, logger=None, audit=None, audit_context=None):
    if _is_blank(incoming_value):
        source_reason = _missing_source_reason(target_param_name, audit_context)
        counters["skip_no_source"] += 1
        _record_reason_stat(counters, element, target_param_name, source_reason)
        if audit is not None:
            row = _build_param_audit_row(audit_context, element, target_param_name, incoming_value)
            row["outcome"] = "SKIP"
            row["reason"] = source_reason
            audit.write_row(row)
        return "SKIP", source_reason

    targets, resolution_reason, all_candidates = _find_param_target(element, target_param_name, parameter_target_mode, logger)
    candidate_count = len(all_candidates)
    if not targets:
        if resolution_reason == "readonly":
            counters["read_only"] += 1
        elif resolution_reason in ("no-family", "no-shared"):
            counters["skip_target_mode"] += 1
        else:
            counters["missing_target"] += 1
        _record_reason_stat(counters, element, target_param_name, resolution_reason)
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
            row = _build_param_audit_row(audit_context, element, target_param_name, incoming_value)
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
            _record_reason_stat(counters, element, target_param_name, "read-only")
            messages.append("{0}:{1}:read-only".format(target_scope, candidate["id"]))
            if audit is not None:
                row = _build_param_audit_row(audit_context, element, target_param_name, incoming_value)
                row["candidate_count"] = candidate_count
                row["selected_target_count"] = len(targets)
                _populate_param_audit_target_fields(row, candidate)
                row["actual_after"] = row.get("actual_before", "")
                row["outcome"] = "FAIL"
                row["reason"] = "read-only"
                audit.write_row(row)
            continue

        if mode != "force" and _parameter_value_matches_incoming(param, incoming_value):
            counters["skip_matching"] += 1
            messages.append("{0}:{1}:already-matches".format(target_scope, candidate["id"]))
            if audit is not None:
                row = _build_param_audit_row(audit_context, element, target_param_name, incoming_value)
                row["candidate_count"] = candidate_count
                row["selected_target_count"] = len(targets)
                _populate_param_audit_target_fields(row, candidate)
                row["actual_after"] = row.get("actual_before", "")
                row["outcome"] = "SKIP"
                row["reason"] = "already-matches"
                audit.write_row(row)
            continue

        counters["write_attempted"] += 1
        old_value = _parameter_text(param)
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
                "Writing {0} parameter '{1}' on element {2}".format(
                    target_scope.upper() if target_scope else "UNKNOWN",
                    _safe_str(getattr(getattr(param, "Definition", None), "Name", "")),
                    _element_id_text(element),
                )
            )
            logger.info(
                "Element {0} | {1} | Kind={2} | Scope={3} | Storage={4} | Old='{5}' | New='{6}'".format(
                    _element_id_text(element),
                    _safe_str(getattr(getattr(param, "Definition", None), "Name", "")),
                    _parameter_kind_label(is_shared),
                    target_scope,
                    _safe_str(getattr(param, "StorageType", "")),
                    old_value,
                    incoming_value,
                )
            )

        ok, err = _write_stringish(param, incoming_value)
        if ok:
            counters["pending_verify"] += 1
            audit_row = None
            if audit is not None:
                audit_row = _build_param_audit_row(audit_context, element, target_param_name, incoming_value)
                audit_row["candidate_count"] = candidate_count
                audit_row["selected_target_count"] = len(targets)
                _populate_param_audit_target_fields(audit_row, candidate)
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
            row = _build_param_audit_row(audit_context, element, target_param_name, incoming_value)
            row["candidate_count"] = candidate_count
            row["selected_target_count"] = len(targets)
            _populate_param_audit_target_fields(row, candidate)
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


def _element_id_text(element):
    try:
        return str(element.Id.IntegerValue)
    except Exception:
        return "?"


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


def _read_param_text_by_name(element, names):
    for name in names:
        p = _find_param(element, name)
        text = _read_stringish(p).strip()
        if text:
            return text

    type_el = _get_type_element(element)
    for name in names:
        p = _find_param(type_el, name)
        text = _read_stringish(p).strip()
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
        p = element.get_Parameter(bip)
        text = _read_stringish(p).strip()
        if text:
            return text
    except Exception:
        pass

    return ""


def _append_unique_probe_point(points, point, tol=0.01):
    if point is None:
        return
    for existing in points:
        try:
            if (
                abs(existing.X - point.X) <= tol
                and abs(existing.Y - point.Y) <= tol
                and abs(existing.Z - point.Z) <= tol
            ):
                return
        except Exception:
            continue
    points.append(point)


def _get_element_probe_points(element):
    points = []
    if element is None:
        return points

    try:
        loc = element.Location
        if loc is not None:
            if hasattr(loc, "Point") and loc.Point is not None:
                _append_unique_probe_point(points, loc.Point)
            if hasattr(loc, "Curve") and loc.Curve is not None:
                for t in (0.0, 0.25, 0.5, 0.75, 1.0):
                    try:
                        _append_unique_probe_point(points, loc.Curve.Evaluate(t, True))
                    except Exception:
                        continue
    except Exception:
        pass

    try:
        bb = element.get_BoundingBox(None)
        if bb is None:
            bb = element.get_BoundingBox(doc.ActiveView)
        if bb is not None:
            _append_unique_probe_point(
                points,
                XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5,
                ),
            )
    except Exception:
        pass

    return points


def _point_within_bounds(point, minx, miny, minz, maxx, maxy, maxz, tol=0.1):
    if point is None:
        return False

    try:
        return (
            point.X >= (minx - tol)
            and point.X <= (maxx + tol)
            and point.Y >= (miny - tol)
            and point.Y <= (maxy + tol)
            and point.Z >= (minz - tol)
            and point.Z <= (maxz + tol)
        )
    except Exception:
        return False


def _build_scope_box_wbs01_index(wb_data):
    global _SCOPE_BOX_WBS01_CACHE
    if _SCOPE_BOX_WBS01_CACHE is not None:
        return _SCOPE_BOX_WBS01_CACHE

    name_to_code = {}
    for scope_box_name, wbs01_code in WBH_SCOPE_BOX_WBS01_RULES:
        valid_code = _valid_wbs_code(wbs01_code, wb_data.wbs_maps["WBS01"])
        if valid_code:
            name_to_code[_normalize_text(scope_box_name)] = valid_code

    collected = {}
    try:
        scope_boxes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType()
    except Exception:
        scope_boxes = []

    for scope_box in scope_boxes:
        if scope_box is None:
            continue

        name = _safe_str(getattr(scope_box, "Name", "")).strip()
        code = name_to_code.get(_normalize_text(name), "")
        if not code:
            continue

        try:
            bb = scope_box.get_BoundingBox(None)
            if bb is None:
                bb = scope_box.get_BoundingBox(doc.ActiveView)
        except Exception:
            bb = None
        if bb is None:
            continue

        collected.setdefault(_normalize_text(name), []).append(
            {
                "name": name,
                "code": code,
                "minx": bb.Min.X,
                "miny": bb.Min.Y,
                "minz": bb.Min.Z,
                "maxx": bb.Max.X,
                "maxy": bb.Max.Y,
                "maxz": bb.Max.Z,
            }
        )

    ordered = []
    for scope_box_name, _ in WBH_SCOPE_BOX_WBS01_RULES:
        ordered.extend(collected.get(_normalize_text(scope_box_name), []))

    _SCOPE_BOX_WBS01_CACHE = ordered
    return ordered


def _derive_wbs01_from_scope_box(facts, wb_data):
    element = facts.get("element")
    if element is None:
        return ""

    probe_points = _get_element_probe_points(element)
    if not probe_points:
        return ""

    for scope_box in _build_scope_box_wbs01_index(wb_data):
        for probe_point in probe_points:
            if _point_within_bounds(
                probe_point,
                scope_box["minx"],
                scope_box["miny"],
                scope_box["minz"],
                scope_box["maxx"],
                scope_box["maxy"],
                scope_box["maxz"],
            ):
                return scope_box["code"]

    return ""


def _point_on_segment_2d(px, py, ax, ay, bx, by, tol):
    abx = bx - ax
    aby = by - ay
    apx = px - ax
    apy = py - ay
    ab2 = abx * abx + aby * aby
    if ab2 <= 1e-12:
        dx = px - ax
        dy = py - ay
        return (dx * dx + dy * dy) <= (tol * tol)

    t = (apx * abx + apy * aby) / ab2
    if t < 0.0:
        t = 0.0
    elif t > 1.0:
        t = 1.0

    cx = ax + t * abx
    cy = ay + t * aby
    dx = px - cx
    dy = py - cy
    return (dx * dx + dy * dy) <= (tol * tol)


def _point_in_polygon_2d(px, py, polygon, tol):
    if not polygon or len(polygon) < 3:
        return False

    count = len(polygon)
    for i in range(count):
        ax, ay = polygon[i]
        bx, by = polygon[(i + 1) % count]
        if _point_on_segment_2d(px, py, ax, ay, bx, by, tol):
            return True

    inside = False
    j = count - 1
    for i in range(count):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        crosses = ((yi > py) != (yj > py))
        if crosses:
            denom = (yj - yi)
            if abs(denom) >= 1e-12:
                x_int = (xj - xi) * (py - yi) / denom + xi
                if px < x_int:
                    inside = not inside
        j = i

    return inside


def _resolve_selected_arch_room_link_instance():
    selected_id = _run_option_text("arch_room_link_id")
    if not selected_id:
        return None
    try:
        target_id = int(selected_id)
    except Exception:
        return None
    try:
        return doc.GetElement(ElementId(target_id))
    except Exception:
        return None


def _build_linked_room_index(link_inst):
    if link_inst is None:
        return []

    try:
        link_key = link_inst.Id.IntegerValue
    except Exception:
        return []

    cached = _LINKED_ROOM_INDEX_CACHE.get(link_key)
    if cached is not None:
        return cached

    try:
        link_doc = link_inst.GetLinkDocument()
        inv_transform = link_inst.GetTotalTransform().Inverse
    except Exception:
        _LINKED_ROOM_INDEX_CACHE[link_key] = []
        return []

    if link_doc is None:
        _LINKED_ROOM_INDEX_CACHE[link_key] = []
        return []

    index = []
    try:
        rooms = FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
    except Exception:
        rooms = []

    for room in rooms:
        if room is None:
            continue

        minx = None
        miny = None
        minz = None
        maxx = None
        maxy = None
        maxz = None
        loops = []

        try:
            bb = room.get_BoundingBox(None)
            if bb is not None:
                minx = bb.Min.X
                miny = bb.Min.Y
                minz = bb.Min.Z
                maxx = bb.Max.X
                maxy = bb.Max.Y
                maxz = bb.Max.Z
        except Exception:
            pass

        room_level_name = ""
        try:
            level = link_doc.GetElement(room.LevelId)
            if level is not None:
                room_level_name = _safe_str(getattr(level, "Name", "")).strip()
        except Exception:
            pass

        index.append(
            {
                "room": room,
                "link_doc": link_doc,
                "inv_transform": inv_transform,
                "room_id": room.Id.IntegerValue,
                "minx": minx,
                "miny": miny,
                "minz": minz,
                "maxx": maxx,
                "maxy": maxy,
                "maxz": maxz,
                "loops": loops,
                "room_level_name": room_level_name,
            }
        )

    _LINKED_ROOM_INDEX_CACHE[link_key] = index
    return index


def _find_linked_room_for_element(element, selected_link=None, tol=1.0):
    link_candidates = []
    if selected_link is not None:
        link_candidates = [selected_link]
    else:
        try:
            link_candidates = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
        except Exception:
            link_candidates = []

    points = _get_element_probe_points(element)
    if not points:
        return None

    for link_inst in link_candidates:
        if link_inst is None:
            continue
        room_index = _build_linked_room_index(link_inst)
        if not room_index:
            continue

        for room_item in room_index:
            room = room_item.get("room")
            if room is None:
                continue

            for probe_point in points:
                try:
                    inv_transform = room_item.get("inv_transform")
                    p = inv_transform.OfPoint(probe_point) if inv_transform is not None else probe_point
                except Exception:
                    continue

                try:
                    minx = room_item.get("minx")
                    miny = room_item.get("miny")
                    minz = room_item.get("minz")
                    maxx = room_item.get("maxx")
                    maxy = room_item.get("maxy")
                    maxz = room_item.get("maxz")
                    if minx is not None and (p.X < (minx - tol) or p.X > (maxx + tol)):
                        continue
                    if miny is not None and (p.Y < (miny - tol) or p.Y > (maxy + tol)):
                        continue
                    if minz is not None and (p.Z < (minz - tol) or p.Z > (maxz + tol)):
                        continue
                except Exception:
                    pass

                try:
                    if hasattr(room, "IsPointInRoom") and room.IsPointInRoom(p):
                        return room_item
                except Exception:
                    pass

                loops = room_item.get("loops") or []
                if loops:
                    try:
                        px = p.X
                        py = p.Y
                        for poly in loops:
                            if _point_in_polygon_2d(px, py, poly, tol):
                                return room_item
                    except Exception:
                        pass

    return None


def _iter_connected_elements(element):
    if element is None:
        return

    seen = set()
    try:
        element_id_value = element.Id.IntegerValue
    except Exception:
        element_id_value = None

    try:
        connector_manager = get_connector_manager(element)
        connectors = connector_manager.Connectors if connector_manager is not None else None
    except NoConnectorManagerError:
        return
    except Exception:
        return

    if connectors is None:
        return

    for connector in connectors:
        try:
            references = connector.AllRefs
        except Exception:
            continue

        for other_connector in references:
            if other_connector is None or other_connector == connector:
                continue

            try:
                owner = other_connector.Owner
            except Exception:
                owner = None

            if owner is None:
                continue

            try:
                owner_id_value = owner.Id.IntegerValue
            except Exception:
                continue

            if element_id_value is not None and owner_id_value == element_id_value:
                continue
            if owner_id_value in seen:
                continue

            seen.add(owner_id_value)
            yield owner


def _find_connected_room_item_bfs(element, selected_link, max_depth=4, max_nodes=60):
    if element is None:
        return None, 0, None, 0

    seen = set()
    try:
        seen.add(element.Id.IntegerValue)
    except Exception:
        pass

    queue = [(element, 0)]
    queue_index = 0
    visited_count = 0

    while queue_index < len(queue) and visited_count < max_nodes:
        current, depth = queue[queue_index]
        queue_index += 1
        if depth >= max_depth:
            continue

        for connected_element in _iter_connected_elements(current):
            if connected_element is None:
                continue

            try:
                connected_id_value = connected_element.Id.IntegerValue
            except Exception:
                connected_id_value = None

            if connected_id_value is not None and connected_id_value in seen:
                continue
            if connected_id_value is not None:
                seen.add(connected_id_value)

            next_depth = depth + 1
            visited_count += 1
            queue.append((connected_element, next_depth))

            room_item = _find_linked_room_for_element(connected_element, selected_link)
            if room_item is not None:
                return room_item, visited_count, connected_element, next_depth

            if visited_count >= max_nodes:
                break

    return None, visited_count, None, 0


def _linked_room_context_from_room_item(room_item, source_label):
    if room_item is None:
        return {}

    room = room_item.get("room")
    if room is None:
        return {}

    room_department = _read_param_text_by_name(room, ["NWB_Department", "NWB Department", "Department", "Room Department", "Department Name", "Department Code", "Dept"])
    room_subdepartment = _read_param_text_by_name(room, ["NWB_SubDepartment", "NWB_Subdepartment", "NWB SubDepartment", "SubDepartment", "Subdepartment", "Room SubDepartment"])
    room_building = _read_param_text_by_name(room, ["Area/Building", "Area Building", "Building", "Building Name", "Building Code", "Block"])
    room_site = _read_param_text_by_name(room, ["Project Location", "Site", "Site Name", "Site Code", "Campus", "Facility"])

    return {
        "room": room,
        "department": room_department,
        "subdepartment": room_subdepartment,
        "building": room_building,
        "site": room_site,
        "room_name": _read_param_text_by_name(room, ["Name"]),
        "room_number": _read_param_text_by_name(room, ["Number", "Room Number"]),
        "link_name": _safe_str(getattr(room_item.get("link_doc"), "Title", "")).strip(),
        "source": source_label,
    }


def _get_linked_room_context(element):
    selected_link = _resolve_selected_arch_room_link_instance()
    room_item = _find_linked_room_for_element(element, selected_link)
    if room_item is not None:
        return _linked_room_context_from_room_item(room_item, "self")

    room_item, visited_count, connected_element, depth = _find_connected_room_item_bfs(element, selected_link)
    if room_item is not None:
        try:
            connected_id_value = connected_element.Id.IntegerValue
        except Exception:
            connected_id_value = ""
        connected_category = _get_category_name(connected_element)
        source_label = "connected:d{0}:{1}:{2}".format(depth, connected_id_value, connected_category)
        return _linked_room_context_from_room_item(room_item, source_label)

    return {
        "room": None,
        "department": "",
        "subdepartment": "",
        "building": "",
        "site": "",
        "room_name": "",
        "room_number": "",
        "link_name": "",
        "source": "connected-miss:d4:visited{0}".format(visited_count),
    }


def _get_category_name(element):
    try:
        cat = element.Category
        if cat is not None:
            return _safe_str(cat.Name).strip()
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


def _log_element_runtime_context(logger, element, element_id_text, facts, target_values, debug):
    if logger is None:
        return

    logger.write(
        "SOURCE | Element {0} | Cat={1} | Family={2} | Type={3} | Level={4} | System={5} | LinkedRoomLink={6} | LinkedRoomNumber={7} | LinkedRoomName={8} | LinkedRoomDept={9} | LinkedRoomBuilding={10} | LinkedRoomSite={11} | LinkedRoomSource={12} | ResolvedDept={13} | ResolvedBuilding={14} | ResolvedSite={15} | DisciplineHint={16} | AssetCandidates={17} | WBS00={18} | WBS01={19} | WBS03={20} | DesignPkg={21} | BuildingPermitPkg={22}".format(
            element_id_text,
            facts.get("category", ""),
            facts.get("family", ""),
            facts.get("type", ""),
            facts.get("level", ""),
            facts.get("system", ""),
            facts.get("linked_room_link_name", ""),
            facts.get("linked_room_number", ""),
            facts.get("linked_room_name", ""),
            facts.get("linked_room_department", ""),
            facts.get("linked_room_building", ""),
            facts.get("linked_room_site", ""),
            facts.get("linked_room_source", ""),
            facts.get("department", ""),
            facts.get("building_hint", ""),
            facts.get("site_hint", ""),
            facts.get("discipline_hint_norm", ""),
            ", ".join(facts.get("asset_category_candidates", []) or []),
            target_values.get("NWB_WBS00", ""),
            target_values.get("NWB_WBS01", ""),
            target_values.get("NWB_WBS03", ""),
            target_values.get("NWB_DesignPkg", ""),
            target_values.get("NWB_BuildingPermitPkg", ""),
        )
    )

    if _is_blank(target_values.get("NWB_Department", "")):
        logger.write(
            "WARN | Element {0} | Department unresolved | ExistingNWBDepartment={1} | ExistingWBS00={2} | ExistingWBS01={3} | ExistingDiscipline={4}".format(
                element_id_text,
                _read_param_text_by_name(element, ["NWB_Department"]),
                facts.get("existing_wbs00", ""),
                facts.get("existing_wbs01", ""),
                facts.get("existing_discipline", ""),
            )
        )


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

    for bip_name in [
        "INSTANCE_REFERENCE_LEVEL_PARAM",
        "FAMILY_LEVEL_PARAM",
        "RBS_START_LEVEL_PARAM",
    ]:
        try:
            bip = getattr(BuiltInParameter, bip_name)
        except Exception:
            bip = None
        if bip is None:
            continue
        try:
            p = element.get_Parameter(bip)
            if p is None:
                continue
            eid = p.AsElementId()
            if eid is None or eid == ElementId.InvalidElementId:
                continue
            lvl = doc.GetElement(eid)
            if lvl is not None:
                return _safe_str(getattr(lvl, "Name", "")).strip()
        except Exception:
            continue

    return ""


def _get_system_classification(element):
    text = _read_param_text_by_bip(element, "RBS_SYSTEM_CLASSIFICATION_PARAM")
    if text:
        return text
    return _read_param_text_by_name(element, ["System Classification", "System Type", "System Name", "NWB_System"])


def _get_site_text(element):
    return _read_param_text_by_name(
        element,
        [
            "Project Location",
            "NWB_Site",
            "Site",
            "Site Code",
            "Site Name",
            "Campus",
            "Facility",
        ],
    )


def _get_building_text(element):
    return _read_param_text_by_name(
        element,
        [
            "Building",
            "Building Name",
            "Building Code",
            "Building_No",
            "Building No",
            "Building Number",
            "Block",
            "Location",
        ],
    )


def _get_material_text(element):
    return _read_param_text_by_name(
        element,
        ["NWB_Material", "Material", "Material Description", "Structural Material", "Pipe Material", "Duct Material"],
    )


def _try_get_spatial_element_from_bip(element, bip_name):
    try:
        bip = getattr(BuiltInParameter, bip_name)
    except Exception:
        bip = None
    if bip is None:
        return None

    try:
        p = element.get_Parameter(bip)
        if p is None:
            return None
        eid = p.AsElementId()
        if eid is None or eid == ElementId.InvalidElementId:
            return None
        return doc.GetElement(eid)
    except Exception:
        return None


def _get_department_text(element):
    direct = _read_param_text_by_name(element, ["NWB_Department", "NWB Department", "Department", "Room Department", "Department Name", "Department Code", "Dept"])
    if direct:
        return direct

    for bip_name in ["RBS_SPACE_ASSOCIATED_ID", "ROOM_ID", "ELEM_ROOM_ID"]:
        spatial = _try_get_spatial_element_from_bip(element, bip_name)
        if spatial is None:
            continue
        value = _read_param_text_by_name(spatial, ["NWB_Department", "NWB Department", "Department", "Room Department", "Department Name", "Department Code", "Dept"])
        if value:
            return value

    return ""


def _get_subdepartment_text(element):
    parameter_names = ["NWB_SubDepartment", "NWB_Subdepartment", "NWB SubDepartment", "SubDepartment", "Subdepartment", "Room SubDepartment"]
    direct = _read_param_text_by_name(element, parameter_names)
    if direct:
        return direct

    for bip_name in ["RBS_SPACE_ASSOCIATED_ID", "ROOM_ID", "ELEM_ROOM_ID"]:
        spatial = _try_get_spatial_element_from_bip(element, bip_name)
        if spatial is None:
            continue
        value = _read_param_text_by_name(spatial, parameter_names)
        if value:
            return value

    return ""


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


def _excel_open_workbook(path):
    excel_type = Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise RuntimeError("Microsoft Excel COM is unavailable.")

    excel = Activator.CreateInstance(excel_type)
    _com_set(excel, "Visible", False)
    _com_set(excel, "DisplayAlerts", False)

    workbooks = None
    workbook = None
    try:
        workbooks = _com_get(excel, "Workbooks")
        workbook = _com_call(workbooks, "Open", path)
        return excel, workbook
    finally:
        _com_release(workbooks)


def _excel_close(excel, workbook):
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

    _com_release(workbook)
    _com_release(excel)


def _normalize_excel_values(value2):
    if value2 is None:
        return []

    try:
        rank = int(getattr(value2, "Rank", 0))
    except Exception:
        rank = 0

    if rank == 2:
        rows = []
        try:
            row_start = int(value2.GetLowerBound(0))
            row_end = int(value2.GetUpperBound(0))
            col_start = int(value2.GetLowerBound(1))
            col_end = int(value2.GetUpperBound(1))
            for r in range(row_start, row_end + 1):
                row_vals = []
                for c in range(col_start, col_end + 1):
                    try:
                        row_vals.append(value2.GetValue(r, c))
                    except Exception:
                        row_vals.append("")
                rows.append(row_vals)
        except Exception:
            rows = []
        if rows:
            return rows

    if rank == 1:
        try:
            start = int(value2.GetLowerBound(0))
            end = int(value2.GetUpperBound(0))
            return [[value2.GetValue(i)] for i in range(start, end + 1)]
        except Exception:
            pass

    if not hasattr(value2, "__iter__"):
        return [[value2]]

    rows = []
    try:
        for row in value2:
            if hasattr(row, "__iter__"):
                rows.append(list(row))
            else:
                rows.append([row])
    except Exception:
        return [[value2]]

    return rows


def _excel_value_to_text(value):
    if value is None:
        return ""

    try:
        if isinstance(value, float):
            if abs(value - round(value)) <= 1e-9:
                return str(int(round(value)))
            return _safe_str(value).strip()
    except Exception:
        pass

    try:
        if isinstance(value, int):
            return str(value)
    except Exception:
        pass

    return _safe_str(value).strip()


def _excel_find_sheet(workbook, sheet_name):
    sheets = None
    try:
        sheets = _com_get(workbook, "Worksheets")
        count = int(_com_get(sheets, "Count"))
        target_norm = _normalize_text(sheet_name)
        for i in range(1, count + 1):
            ws = None
            try:
                ws = _com_item(sheets, i)
                ws_name = _safe_str(_com_get(ws, "Name")).strip()
                ws_norm = _normalize_text(ws_name)
                if ws_norm == target_norm or target_norm in ws_norm or ws_norm in target_norm:
                    return ws
            except Exception:
                pass
            finally:
                ws_name_check = ""
                if ws is not None:
                    try:
                        ws_name_check = _normalize_text(_safe_str(_com_get(ws, "Name")))
                    except Exception:
                        ws_name_check = ""
                if ws is not None and not (
                    ws_name_check == target_norm
                    or target_norm in ws_name_check
                    or ws_name_check in target_norm
                ):
                    _com_release(ws)
    finally:
        _com_release(sheets)

    return None


def _read_sheet_matrix(workbook, sheet_name, max_rows=None, max_cols=None, prefer_display_text=False):
    ws = _excel_find_sheet(workbook, sheet_name)
    if ws is None:
        return []

    used = None
    rows_obj = None
    cols_obj = None
    cells_obj = None
    try:
        used = _com_get(ws, "UsedRange")
        rows_obj = _com_get(used, "Rows")
        cols_obj = _com_get(used, "Columns")
        row_count = int(_com_get(rows_obj, "Count"))
        col_count = int(_com_get(cols_obj, "Count"))

        if max_rows is not None:
            try:
                row_count = min(row_count, max(0, int(max_rows)))
            except Exception:
                pass
        if max_cols is not None:
            try:
                col_count = min(col_count, max(0, int(max_cols)))
            except Exception:
                pass

        if row_count <= 0 or col_count <= 0:
            return []

        if not prefer_display_text:
            try:
                bulk_values = _com_get(used, "Value2")
                raw_rows = _normalize_excel_values(bulk_values)
                matrix = []
                for r in range(min(row_count, len(raw_rows))):
                    raw_row = raw_rows[r] if r < len(raw_rows) else []
                    row_vals = []
                    for c in range(col_count):
                        raw_value = raw_row[c] if c < len(raw_row) else ""
                        row_vals.append(_excel_value_to_text(raw_value))
                    matrix.append(row_vals)
                if matrix:
                    return matrix
            except Exception:
                pass

        if cells_obj is None:
            cells_obj = _com_get(ws, "Cells")
        matrix = []
        for r in range(1, row_count + 1):
            row_vals = []
            for c in range(1, col_count + 1):
                cell_obj = None
                try:
                    cell_obj = _com_item2(cells_obj, r, c)
                    val = _com_get(cell_obj, "Text")
                except Exception:
                    val = ""
                finally:
                    _com_release(cell_obj)
                row_vals.append(_safe_str(val).strip())
            matrix.append(row_vals)

            if r % 25 == 0:
                _pump_ui()

        return matrix
    finally:
        _com_release(cells_obj)
        _com_release(rows_obj)
        _com_release(cols_obj)
        _com_release(used)
        _com_release(ws)


def _pick_header_row(matrix, scan_rows=6):
    if not matrix:
        return 0

    best_idx = 0
    best_score = -1
    limit = min(len(matrix), max(1, scan_rows))
    for i in range(limit):
        row = matrix[i]
        non_empty = 0
        alpha_like = 0
        for cell in row:
            text = _safe_str(cell).strip()
            if not text:
                continue
            non_empty += 1
            if re.search(r"[A-Za-z_]", text):
                alpha_like += 1
        score = (alpha_like * 3) + non_empty
        if score > best_score:
            best_score = score
            best_idx = i

    return best_idx


def _matrix_to_dict_rows(matrix, header_row_index=None):
    if not matrix:
        return []

    hdr_idx = _pick_header_row(matrix) if header_row_index is None else max(0, header_row_index)
    header = matrix[hdr_idx]

    headers = []
    seen = {}
    for idx, raw in enumerate(header):
        text = _safe_str(raw).strip()
        if not text:
            text = "Column_{0}".format(idx + 1)
        key = text
        dup = seen.get(key, 0)
        if dup > 0:
            key = "{0}_{1}".format(key, dup + 1)
        seen[text] = dup + 1
        headers.append(key)

    rows = []
    for row in matrix[hdr_idx + 1:]:
        if row is None:
            continue
        if len(row) == 0:
            continue

        item = {}
        has_any = False
        for i, h in enumerate(headers):
            value = row[i] if i < len(row) else ""
            text = _safe_str(value).strip()
            item[h] = text
            if text:
                has_any = True

        if has_any:
            rows.append(item)

    return rows


def _row_with_norm_keys(row):
    out = {}
    for k, v in row.items():
        out[_normalize_key(k)] = _safe_str(v).strip()
    return out


def _parse_work_packages(matrix, data):
    rows = _matrix_to_dict_rows(matrix)
    for row in rows:
        row_norm = _row_with_norm_keys(row)
        code = row_norm.get("subpackagenumber", "").strip().upper()
        sub_package = row_norm.get("subpackage", "").strip()
        associated_bp = row_norm.get("associatedbp", "").strip()
        if code and code.startswith("NWB-"):
            data.work_package_codes.add(code)
            data.work_packages.append(
                {
                    "code": code,
                    "sub_package": sub_package,
                    "sub_package_norm": _normalize_text(sub_package),
                    "associated_bp": associated_bp,
                    "associated_bp_key": _extract_building_permit_key(associated_bp),
                }
            )


def _parse_building_permit(matrix, data):
    rows = _matrix_to_dict_rows(matrix, header_row_index=0)
    for row in rows:
        row_norm = _row_with_norm_keys(row)
        item = row_norm.get("buildingpermititem", "").strip()
        if item:
            data.building_permit_items.add(item)
            key = _extract_building_permit_key(item)
            if key:
                data.building_permit_by_key[key] = item


def _load_workbook_data(workbook_path, status_callback=None):
    data = WorkbookData()

    excel = None
    workbook = None
    try:
        _notify_status(status_callback, "Opening workbook...")
        excel, workbook = _excel_open_workbook(workbook_path)

        _notify_status(status_callback, "Reading sheet: Asset Master")
        asset_matrix = _read_sheet_matrix(workbook, "Asset Master", max_cols=70)
        asset_rows_raw = _matrix_to_dict_rows(asset_matrix, header_row_index=0)
        data.asset_rows = [_row_with_norm_keys(x) for x in asset_rows_raw]

        _notify_status(status_callback, "Reading sheet: Systems Table")
        systems_matrix = _read_sheet_matrix(workbook, "Systems Table", max_cols=4)
        systems_rows = _matrix_to_dict_rows(systems_matrix)
        for row in systems_rows:
            for value in row.values():
                text = _safe_str(value).strip()
                if text and _normalize_text(text) != "asset system (keyword)":
                    data.system_keywords.add(_normalize_text(text))

        _notify_status(status_callback, "Reading sheet: Disciplines")
        disciplines_matrix = _read_sheet_matrix(workbook, "Disciplines", max_cols=4)
        disciplines_rows = _matrix_to_dict_rows(disciplines_matrix, header_row_index=0)
        for row in disciplines_rows:
            row_norm = _row_with_norm_keys(row)
            code = row_norm.get("disciplinecode", "").strip().upper()
            desc = row_norm.get("disciplinedescription", "").strip()
            if code:
                data.discipline_code_to_desc[code] = desc
            if desc:
                data.discipline_desc_to_code[_normalize_text(desc)] = code

        _notify_status(status_callback, "Reading sheet: WBS")
        wbs_matrix = _read_sheet_matrix(workbook, "WBS", max_cols=16, prefer_display_text=True)
        _parse_wbs_matrix_into_data(wbs_matrix, data)

        _notify_status(status_callback, "Reading sheet: Work Packages")
        work_packages_matrix = _read_sheet_matrix(workbook, "Work Packages", max_cols=40)
        _parse_work_packages(work_packages_matrix, data)

        _notify_status(status_callback, "Reading sheet: Building Permit")
        building_permit_matrix = _read_sheet_matrix(workbook, "Building Permit", max_cols=6)
        _parse_building_permit(building_permit_matrix, data)

        _notify_status(status_callback, "Validating workbook mapping data...")
        if not data.asset_rows:
            raise RuntimeError("Asset Master sheet is empty or unreadable.")
        if not data.wbs_maps["WBS03"]["code_to_desc"]:
            raise RuntimeError("WBS sheet is empty or unreadable.")
        if not data.work_package_codes:
            raise RuntimeError("Work Packages sheet is empty or unreadable.")
        if not data.building_permit_items:
            raise RuntimeError("Building Permit sheet is empty or unreadable.")

        _notify_status(
            status_callback,
            "Workbook ready. Asset rows={0}, WBS03 codes={1}, Work Packages={2}, Building Permit={3}".format(
                len(data.asset_rows),
                len(data.wbs_maps["WBS03"]["code_to_desc"]),
                len(data.work_package_codes),
                len(data.building_permit_items),
            ),
        )
        return data
    finally:
        _excel_close(excel, workbook)


def _parse_wbs_matrix_into_data(matrix, data):
    if not matrix:
        return

    header_idx = -1
    for i, row in enumerate(matrix):
        row_text = " | ".join([_normalize_text(c) for c in row])
        if "wbs00" in row_text and "wbs01" in row_text and "wbs05" in row_text:
            header_idx = i
            break

    if header_idx < 0:
        return

    pair_map = [
        ("WBS00", 0, 1),
        ("WBS01", 2, 3),
        ("WBS02", 4, 5),
        ("WBS03", 6, 7),
        ("WBS04", 8, 9),
        ("WBS05", 10, 11),
    ]

    code_row_idx = header_idx + 1
    data_start = header_idx + 2

    if code_row_idx >= len(matrix):
        return

    for r in range(data_start, len(matrix)):
        row = matrix[r]
        has_any = False
        for wbs_key, code_col, desc_col in pair_map:
            code = _safe_str(row[code_col] if code_col < len(row) else "").strip()
            desc = _safe_str(row[desc_col] if desc_col < len(row) else "").strip()
            if not code and not desc:
                continue

            has_any = True
            code_up = code.strip().upper()
            desc_norm = _normalize_text(desc)
            wbs_bucket = data.wbs_maps[wbs_key]

            if code_up:
                wbs_bucket["code_to_desc"][code_up] = desc
                if not wbs_bucket["default"]:
                    wbs_bucket["default"] = code_up
            if desc_norm and code_up:
                wbs_bucket["desc_to_code"][desc_norm] = code_up

        if not has_any:
            # Stop only after passing through initial sparse rows.
            if r > data_start + 4:
                break

    if "Z00" not in data.wbs_maps["WBS03"]["code_to_desc"]:
        data.wbs_maps["WBS03"]["code_to_desc"]["Z00"] = "Not applicable"
        data.wbs_maps["WBS03"]["desc_to_code"][_normalize_text("Not applicable")] = "Z00"


def _row_value(row_norm, key):
    return _safe_str(row_norm.get(_normalize_key(key), "")).strip()


def _match_wbs_code_from_text(text, wbs_bucket):
    raw = _safe_str(text).strip()
    if not raw:
        return ""

    up = raw.upper()
    if up in wbs_bucket["code_to_desc"]:
        return up

    text_norm = _normalize_text(raw)
    exact = wbs_bucket["desc_to_code"].get(text_norm, "")
    if exact:
        return exact

    best_code = ""
    best_len = -1
    for desc_norm, code in wbs_bucket["desc_to_code"].items():
        if not desc_norm:
            continue
        if desc_norm in text_norm or text_norm in desc_norm:
            if len(desc_norm) > best_len:
                best_len = len(desc_norm)
                best_code = code

    return best_code


def _valid_wbs_code(value, wbs_bucket):
    code = _safe_str(value).strip().upper()
    if code and code in wbs_bucket["code_to_desc"]:
        return code
    return ""


def _extract_building_permit_key(value):
    up = _safe_str(value).strip().upper()
    if not up:
        return ""

    match = re.search(r"\b(WBH|OPH)\s*[- ]\s*(BP\d+[A-Z]?)\b", up)
    if not match:
        return ""

    return "{0}-{1}".format(match.group(1), match.group(2))


def _canonical_building_permit(value, wb_data):
    if _is_blank(value):
        return ""

    key = _extract_building_permit_key(value)
    if key and key in wb_data.building_permit_by_key:
        return wb_data.building_permit_by_key[key]

    text_norm = _normalize_text(value)
    for item in wb_data.building_permit_items:
        item_norm = _normalize_text(item)
        if item_norm == text_norm or item_norm in text_norm or text_norm in item_norm:
            return item

    return _safe_str(value).strip()


def _get_project_information():
    try:
        return doc.ProjectInformation
    except Exception:
        return None


def _read_project_info_text(names):
    project_info = _get_project_information()
    if project_info is None:
        return ""
    return _read_param_text_by_name(project_info, names)


def _get_project_info_snapshot():
    project_name = _read_project_info_text(["Project Name"])
    if _is_blank(project_name):
        try:
            project_name = _safe_str(getattr(doc.ProjectInformation, "Name", "")).strip()
        except Exception:
            project_name = ""

    project_number = _read_project_info_text(["Project Number"])
    if _is_blank(project_number):
        try:
            project_number = _safe_str(getattr(doc.ProjectInformation, "Number", "")).strip()
        except Exception:
            project_number = ""

    return {
        "project_location": _read_project_info_text(["Project Location", "PROJECT LOCATION", "NWB_WBS00", "NWB_Site", "Site", "Site Code", "Site Name"]),
        "building_name": _read_project_info_text(["Building Name", "Building", "Project Building"]),
        "building_number": _read_project_info_text(["Building_No", "Building No", "Building Number", "Building Code", "NWB_WBS01"]),
        "project_name": project_name,
        "project_number": project_number,
    }


def _get_project_context(wb_data):
    global _PROJECT_CONTEXT_CACHE
    if _PROJECT_CONTEXT_CACHE is not None:
        return _PROJECT_CONTEXT_CACHE

    snapshot = _get_project_info_snapshot()

    project_info = _get_project_information()
    context = {
        "project_name": "",
        "project_number": "",
        "project_wbs00_raw": "",
        "project_wbs01_raw": "",
        "site_hint": "",
        "building_hint": "",
        "address": "",
        "client_name": "",
        "site_code": "",
        "building_code": "",
        "search_blob": "",
        "site_search_blob": "",
        "building_search_blob": "",
        "site_override": "",
        "building_override": "",
    }

    if project_info is None:
        _PROJECT_CONTEXT_CACHE = context
        return context

    context["project_name"] = _safe_str(snapshot.get("project_name", "")).strip()
    context["project_number"] = _safe_str(snapshot.get("project_number", "")).strip()
    context["site_override"] = _run_option_text("site_override") or _safe_str(snapshot.get("project_location", "")).strip()
    context["building_override"] = _run_option_text("building_override") or _safe_str(snapshot.get("building_name", "")).strip() or _safe_str(snapshot.get("building_number", "")).strip()
    context["project_wbs00_raw"] = _read_project_info_text(["NWB_WBS00"])
    context["project_wbs01_raw"] = _read_project_info_text(["NWB_WBS01"])
    context["site_hint"] = _read_project_info_text([
        "Project Location",
        "PROJECT LOCATION",
        "NWB_Site",
        "Site",
        "Site Code",
        "Site Name",
        "Project Site",
        "Campus",
        "Facility",
        "NWB_WBS00",
    ])
    context["building_hint"] = _read_project_info_text([
        "Building",
        "Building Name",
        "Building Code",
        "Building_No",
        "Building No",
        "Building Number",
        "Project Building",
        "Block",
        "NWB_WBS01",
    ])
    context["address"] = _read_project_info_text(["Address", "Project Address"])
    context["client_name"] = _read_project_info_text(["Client Name", "Client"])

    context["site_code"] = _match_wbs_code_from_text(context["site_override"], wb_data.wbs_maps["WBS00"])
    if not context["site_code"]:
        context["site_code"] = _match_wbs_code_from_text(context["project_wbs00_raw"], wb_data.wbs_maps["WBS00"])
    context["building_code"] = _match_wbs_code_from_text(context["building_override"], wb_data.wbs_maps["WBS01"])
    if not context["building_code"]:
        context["building_code"] = _match_wbs_code_from_text(context["project_wbs01_raw"], wb_data.wbs_maps["WBS01"])

    context["site_search_blob"] = " | ".join(
        [
            _normalize_text(context["site_override"]),
            _normalize_text(context["project_wbs00_raw"]),
            _normalize_text(context["project_name"]),
            _normalize_text(context["project_number"]),
            _normalize_text(context["site_hint"]),
            _normalize_text(context["address"]),
            _normalize_text(context["client_name"]),
        ]
    )
    context["building_search_blob"] = " | ".join(
        [
            _normalize_text(context["building_override"]),
            _normalize_text(context["project_wbs01_raw"]),
            _normalize_text(context["project_name"]),
            _normalize_text(context["project_number"]),
            _normalize_text(context["building_hint"]),
            _normalize_text(context["address"]),
        ]
    )

    context["search_blob"] = " | ".join(
        [
            _normalize_text(context["project_name"]),
            _normalize_text(context["project_number"]),
            _normalize_text(context["site_hint"]),
            _normalize_text(context["building_hint"]),
            _normalize_text(context["address"]),
            _normalize_text(context["client_name"]),
        ]
    )
    if not context["site_code"]:
        context["site_code"] = _match_wbs_code_from_text(context["site_hint"], wb_data.wbs_maps["WBS00"])
    if not context["site_code"]:
        context["site_code"] = _match_wbs_code_from_text(context["site_search_blob"], wb_data.wbs_maps["WBS00"])

    if not context["building_code"]:
        context["building_code"] = _match_wbs_code_from_text(context["building_hint"], wb_data.wbs_maps["WBS01"])
    if not context["building_code"]:
        context["building_code"] = _match_wbs_code_from_text(context["building_search_blob"], wb_data.wbs_maps["WBS01"])

    _PROJECT_CONTEXT_CACHE = context
    return context


def _derive_element_facts(element):
    facts = {}
    facts["element"] = element
    linked_room = _get_linked_room_context(element)
    direct_department = _get_department_text(element)
    direct_subdepartment = _get_subdepartment_text(element)
    direct_site = _get_site_text(element)
    direct_building = _get_building_text(element)

    facts["name"] = _get_element_name(element)
    facts["category"] = _get_category_name(element)
    facts["family"] = _get_family_name(element)
    facts["type"] = _get_type_name(element)
    facts["level"] = _get_level_name(element)
    facts["system"] = _get_system_classification(element)
    facts["material"] = _get_material_text(element)
    facts["linked_room_department"] = _safe_str(linked_room.get("department", "")).strip()
    facts["linked_room_subdepartment"] = _safe_str(linked_room.get("subdepartment", "")).strip()
    facts["linked_room_building"] = _safe_str(linked_room.get("building", "")).strip()
    facts["linked_room_site"] = _safe_str(linked_room.get("site", "")).strip()
    facts["linked_room_name"] = _safe_str(linked_room.get("room_name", "")).strip()
    facts["linked_room_number"] = _safe_str(linked_room.get("room_number", "")).strip()
    facts["linked_room_link_name"] = _safe_str(linked_room.get("link_name", "")).strip()
    facts["linked_room_source"] = _safe_str(linked_room.get("source", "")).strip()
    facts["department"] = facts["linked_room_department"] or direct_department
    facts["subdepartment"] = facts["linked_room_subdepartment"] or direct_subdepartment
    facts["site_hint"] = facts["linked_room_site"] or direct_site
    facts["building_hint"] = facts["linked_room_building"] or direct_building
    facts["existing_discipline"] = _read_param_text_by_name(element, ["NWB_Discipline"])
    facts["existing_wbs00"] = _read_param_text_by_name(element, ["NWB_WBS00"])
    facts["existing_wbs01"] = _read_param_text_by_name(element, ["NWB_WBS01"])
    facts["existing_wbs02"] = _read_param_text_by_name(element, ["NWB_WBS02"])
    facts["existing_wbs05"] = _read_param_text_by_name(element, ["NWB_WBS05"])

    facts["name_norm"] = _normalize_text(facts["name"])
    facts["category_norm"] = _normalize_text(facts["category"])
    facts["family_norm"] = _normalize_text(facts["family"])
    facts["type_norm"] = _normalize_text(facts["type"])
    facts["level_norm"] = _normalize_text(facts["level"])
    facts["system_norm"] = _normalize_text(facts["system"])
    facts["material_norm"] = _normalize_text(facts["material"])
    facts["department_norm"] = _normalize_text(facts["department"])
    facts["subdepartment_norm"] = _normalize_text(facts["subdepartment"])
    facts["site_hint_norm"] = _normalize_text(facts["site_hint"])
    facts["building_hint_norm"] = _normalize_text(facts["building_hint"])
    facts["linked_room_department_norm"] = _normalize_text(facts["linked_room_department"])
    facts["linked_room_building_norm"] = _normalize_text(facts["linked_room_building"])
    facts["linked_room_site_norm"] = _normalize_text(facts["linked_room_site"])
    facts["discipline_hint_norm"] = _normalize_text(_derive_discipline_hint(facts))
    facts["asset_category_candidates"] = _asset_category_candidates(facts)

    facts["search_blob"] = " | ".join([
        facts["name_norm"],
        facts["category_norm"],
        facts["family_norm"],
        facts["type_norm"],
        facts["level_norm"],
        facts["system_norm"],
        facts["material_norm"],
        facts["site_hint_norm"],
        facts["building_hint_norm"],
        facts["linked_room_site_norm"],
        facts["linked_room_building_norm"],
        facts["linked_room_department_norm"],
        _normalize_text(facts["linked_room_name"]),
        _normalize_text(facts["linked_room_number"]),
        _normalize_text(facts["linked_room_link_name"]),
        _normalize_text(facts["existing_discipline"]),
        _normalize_text(facts["existing_wbs00"]),
        _normalize_text(facts["existing_wbs01"]),
        _normalize_text(facts["existing_wbs02"]),
        _normalize_text(facts["existing_wbs05"]),
        facts["discipline_hint_norm"],
        " | ".join(facts["asset_category_candidates"]),
    ])

    return facts


def _derive_discipline_hint(facts):
    category_norm = facts.get("category_norm", "")
    search_blob = " ".join(
        [
            category_norm,
            facts.get("family_norm", ""),
            facts.get("type_norm", ""),
            facts.get("name_norm", ""),
            facts.get("system_norm", ""),
        ]
    )

    if "security" in search_blob or "cctv" in search_blob or "access control" in search_blob or "intruder" in search_blob:
        return "Security"
    if (
        "ict" in search_blob
        or "data" in search_blob
        or "communication" in search_blob
        or "telephone" in search_blob
        or "nurse call" in search_blob
        or "intercom" in search_blob
        or "audio visual" in search_blob
        or "a/v" in search_blob
        or "av system" in search_blob
        or "fibre" in search_blob
        or "fiber" in search_blob
    ):
        return "ICT Communications Technology"
    if "sprinkler" in category_norm or "fire" in search_blob:
        return "Fire and Fire Safety Systems"
    if (
        "cable tray" in category_norm
        or "conduit" in category_norm
        or "electrical" in category_norm
        or "lighting" in category_norm
        or "data" in category_norm
        or "communication" in category_norm
        or "telephone" in category_norm
    ):
        return "Electrical and Instrumentation"
    if "pipe" in category_norm or "plumbing" in category_norm or "hydraulic" in search_blob or "drainage" in search_blob:
        return "Hydraulics"
    if "utility" in search_blob:
        return "Utilities"
    if "duct" in category_norm or "air terminal" in category_norm or "mechanical equipment" in category_norm:
        return "Mechanical"
    return ""


def _asset_category_candidates(facts):
    category_norm = facts.get("category_norm", "")
    search_blob = " ".join(
        [
            category_norm,
            facts.get("family_norm", ""),
            facts.get("type_norm", ""),
            facts.get("name_norm", ""),
            facts.get("system_norm", ""),
        ]
    )

    candidates = []

    def _add(value):
        norm = _normalize_text(value)
        if norm and norm not in candidates:
            candidates.append(norm)

    _add(facts.get("category", ""))

    if "duct" in category_norm:
        _add("Ductwork")
    if "air terminal" in category_norm:
        _add("Air Terminal")
    if "duct fitting" in category_norm or "duct accessory" in category_norm:
        _add("Ductwork")
        if "damper" in search_blob:
            _add("Damper")
            _add("Motorised Damper")
        if any(token in search_blob for token in ["grille", "grill", "louvre", "louver", "cowl", "diffuser", "terminal", "register", "nozzle"]):
            _add("Air Terminal")
        if "heater" in search_blob or "coil" in search_blob:
            _add("Heating Coil")
        if "attenuator" in search_blob:
            _add("Attenuator")

    if "pipe fitting" in category_norm or "pipe accessory" in category_norm:
        _add("Hydraulics")
        _add("Pipe Fitting")
        _add("Pipework")
        if "valve" in search_blob:
            _add("Valve")
        if "pump" in search_blob:
            _add("Pump")
        if "strainer" in search_blob:
            _add("Strainer")
        if "filter" in search_blob:
            _add("Filter")
        if "gauge" in search_blob:
            _add("Gauge")
        if "meter" in search_blob:
            _add("Meter")
    elif "pipe" in category_norm:
        _add("Hydraulics")
        _add("Pipe")
        _add("Pipework")
        _add("Pipes and fittings")

    if "cable tray fitting" in category_norm:
        _add("Cable Tray Fittings")
        _add("Cable Ladder Fittings")
    elif "cable tray" in category_norm:
        _add("Cable Tray")
        _add("Cable Ladder")

    if "conduit fitting" in category_norm or "conduit" in category_norm:
        _add("Conduit")

    if "sprinkler" in category_norm:
        _add("Sprinkler")

    if "plumbing fixture" in category_norm:
        _add("Hydraulics")
        _add("Plumbing Fixture")

    if "mechanical equipment" in category_norm:
        _add("Equipment")

    if "hydraulic" in search_blob or "plumbing" in search_blob:
        _add("Hydraulics")

    return candidates


def _score_asset_row(row, facts):
    score = 0

    row_category = _normalize_text(_row_value(row, "NWB_Category"))
    row_assettype = _normalize_text(_row_value(row, "NWB_AssetType"))
    row_h3 = _normalize_text(_row_value(row, "Hierarchy - L3"))
    row_h2 = _normalize_text(_row_value(row, "Hierarchy - L2"))
    row_hierarchy = _normalize_text(_row_value(row, "Hierarchy"))
    row_system = _normalize_text(_row_value(row, "NWB_System"))
    category_candidates = facts.get("asset_category_candidates", [])

    if row_category:
        if row_category == facts["category_norm"] or row_category in category_candidates:
            score += 18
        elif row_category in facts["search_blob"] or facts["category_norm"] in row_category:
            score += 11

    if row_assettype:
        if row_assettype == facts["type_norm"] or row_assettype == facts.get("family_norm", "") or row_assettype == facts.get("name_norm", ""):
            score += 14
        elif row_assettype in facts["search_blob"]:
            score += 9

    if row_h3:
        if row_h3 == facts["type_norm"] or row_h3 in category_candidates:
            score += 10
        elif row_h3 in facts["search_blob"]:
            score += 6

    if row_h2:
        if row_h2 == facts.get("discipline_hint_norm", ""):
            score += 8
        elif row_h2 in facts["search_blob"]:
            score += 4

    if row_hierarchy and row_hierarchy in facts["search_blob"]:
        score += 7

    if row_system and not _is_placeholder_text(row_system):
        if row_system == facts["system_norm"]:
            score += 12
        elif row_system in facts["system_norm"] or facts["system_norm"] in row_system:
            score += 6

    return score


def _facts_cache_key(facts):
    """Create cache key from normalized facts (Pattern: MT12 Schedule Export)"""
    return (
        facts.get("category_norm", ""),
        facts.get("family_norm", ""),
        facts.get("type_norm", ""),
        facts.get("name_norm", ""),
        facts.get("system_norm", ""),
    )


def _select_asset_row(asset_rows, facts):
    # OPTIMIZATION (Phase 2.2): Check asset row matching cache first
    global _ASSET_MATCH_CACHE
    cache_key = _facts_cache_key(facts)

    if cache_key in _ASSET_MATCH_CACHE:
        return _ASSET_MATCH_CACHE[cache_key]

    best = None
    best_score = -1

    for row in asset_rows:
        score = _score_asset_row(row, facts)
        if score > best_score:
            best = row
            best_score = score

    if best_score <= 0:
        result = (None, 0)
        _ASSET_MATCH_CACHE[cache_key] = result
        return result

    result = (best, best_score)
    _ASSET_MATCH_CACHE[cache_key] = result
    return result


def _discipline_code_from_value(raw_value, wb_data):
    text = _safe_str(raw_value).strip()
    if not text:
        return ""

    up = text.upper()
    if up in wb_data.discipline_code_to_desc:
        return up
    if up in wb_data.wbs_maps["WBS04"]["code_to_desc"]:
        return up

    code = wb_data.discipline_desc_to_code.get(_normalize_text(text), "")
    if code:
        return code

    code = wb_data.wbs_maps["WBS04"]["desc_to_code"].get(_normalize_text(text), "")
    if code:
        return code

    return up if len(up) <= 4 else ""


def _derive_discipline_code(facts, asset_row, wb_data):
    from_asset = _row_value(asset_row, "NWB_Discipline") if asset_row else ""
    code = _discipline_code_from_value(from_asset, wb_data)
    if code:
        return code

    code = _discipline_code_from_value(facts.get("existing_discipline", ""), wb_data)
    if code:
        return code

    code = _discipline_code_from_value(_derive_discipline_hint(facts), wb_data)
    if code:
        return code

    search = " ".join(
        [
            facts.get("category_norm", ""),
            facts.get("family_norm", ""),
            facts.get("type_norm", ""),
            facts.get("name_norm", ""),
            facts.get("system_norm", ""),
            facts.get("material_norm", ""),
        ]
    )

    if "security" in search or "cctv" in search or "access control" in search or "intruder" in search:
        return "IS"
    if (
        "ict" in search
        or "data" in search
        or "communication" in search
        or "telephone" in search
        or "nurse call" in search
        or "intercom" in search
        or "audio visual" in search
        or "a/v" in search
        or "fiber" in search
        or "fibre" in search
    ):
        return "IC"
    if "sprinkler" in search or "fire" in search or "smoke" in search or "detector" in search or "alarm" in search:
        return "FR"
    if (
        "cable tray" in search
        or "cable ladder" in search
        or "conduit" in search
        or "electrical" in search
        or "lighting" in search
        or "switchboard" in search
        or "distribution board" in search
        or "panelboard" in search
        or "power" in search
        or "socket" in search
    ):
        return "EL"

    cat = facts["category_norm"]
    if "duct" in cat or "mechanical" in cat or "air" in cat:
        return "ME"
    if "pipe" in cat or "plumbing" in cat or "hydraulic" in search or "drainage" in search or "sanitary" in search:
        return "MH"
    if "sprinkler" in cat or "fire" in cat:
        return "FR"
    if "security" in cat:
        return "IS"
    if "data" in cat or "communication" in cat or "telephone" in cat or "nurse call" in cat:
        return "IC"
    if "cable" in cat or "conduit" in cat or "electrical" in cat or "lighting" in cat:
        return "EL"

    if "utility" in search:
        return "UT"

    return "ME"


def _match_wbs_code_by_level(level_name, wbs_bucket):
    if _is_blank(level_name):
        return ""

    lvl = _safe_str(level_name).strip().upper()
    code_to_desc = wbs_bucket["code_to_desc"]

    for code in code_to_desc.keys():
        if not code:
            continue
        pattern = r"(^|[^A-Z0-9]){0}([^A-Z0-9]|$)".format(re.escape(code))
        if re.search(pattern, lvl):
            return code

    desc_to_code = wbs_bucket["desc_to_code"]
    lvl_norm = _normalize_text(level_name)
    for desc_norm, code in desc_to_code.items():
        if not desc_norm:
            continue
        if desc_norm in lvl_norm or lvl_norm in desc_norm:
            return code

    if "ground" in lvl_norm and "GR" in code_to_desc:
        return "GR"

    return ""


def _match_wbs_code_token_by_level(level_name, wbs_bucket):
    if _is_blank(level_name):
        return ""

    lvl = _safe_str(level_name).strip().upper()
    for code in wbs_bucket["code_to_desc"].keys():
        if not code:
            continue
        pattern = r"(^|[^A-Z0-9]){0}([^A-Z0-9]|$)".format(re.escape(code))
        if re.search(pattern, lvl):
            return code

    return ""


def _derive_wbs00(facts, wb_data):
    project_context = _get_project_context(wb_data)

    manual_code = _match_wbs_code_from_text(project_context.get("site_override", ""), wb_data.wbs_maps["WBS00"])
    if manual_code:
        return manual_code

    explicit_project_code = _valid_wbs_code(project_context.get("site_code", ""), wb_data.wbs_maps["WBS00"])
    if explicit_project_code:
        return explicit_project_code

    for value in [
        facts.get("linked_room_site", ""),
        project_context.get("site_hint", ""),
        facts.get("site_hint", ""),
        project_context.get("site_search_blob", ""),
    ]:
        code = _match_wbs_code_from_text(value, wb_data.wbs_maps["WBS00"])
        if code:
            return code

    return ""


def _derive_wbs01(facts, wb_data):
    project_context = _get_project_context(wb_data)
    level_name = facts.get("level", "")

    scope_box_code = _derive_wbs01_from_scope_box(facts, wb_data)
    if scope_box_code:
        return scope_box_code

    manual_code = _match_wbs_code_from_text(project_context.get("building_override", ""), wb_data.wbs_maps["WBS01"])
    if manual_code:
        return manual_code

    linked_room_code = _match_wbs_code_from_text(facts.get("linked_room_building", ""), wb_data.wbs_maps["WBS01"])
    if linked_room_code:
        return linked_room_code

    explicit_project_code = _valid_wbs_code(project_context.get("building_code", ""), wb_data.wbs_maps["WBS01"])
    if explicit_project_code:
        return explicit_project_code

    for value in [
        facts.get("linked_room_building", ""),
        facts.get("building_hint", ""),
        project_context.get("building_hint", ""),
        level_name,
        project_context.get("building_search_blob", ""),
    ]:
        code = _match_wbs_code_from_text(value, wb_data.wbs_maps["WBS01"])
        if code:
            return code

    code = _match_wbs_code_by_level(level_name, wb_data.wbs_maps["WBS01"])
    if code:
        return code

    lvl_norm = _normalize_text(level_name)
    if "external" in lvl_norm:
        return "EA"

    return "EA"


def _derive_wbs02(facts, wb_data):
    level_name = facts.get("level", "")
    direct_code = _match_wbs_code_token_by_level(level_name, wb_data.wbs_maps["WBS02"])
    if direct_code:
        return direct_code

    existing_code = _valid_wbs_code(facts.get("existing_wbs02", ""), wb_data.wbs_maps["WBS02"])
    if existing_code:
        return existing_code

    code = _match_wbs_code_by_level(level_name, wb_data.wbs_maps["WBS02"])
    if code:
        return code

    return wb_data.wbs_maps["WBS02"]["default"] or "NS"


def _derive_wbs03(facts, wb_data):
    # User requirement: no room/department -> Z00.
    dept = _safe_str(facts.get("department", "")).strip()
    if not dept:
        return "Z00"

    code = _valid_wbs_code(dept, wb_data.wbs_maps["WBS03"])
    if code:
        return code

    code = _match_wbs_code_from_text(dept, wb_data.wbs_maps["WBS03"])
    if code:
        return code

    # Unknown department still resolves to Z00 to keep deterministic behavior.
    return "Z00"


def _derive_wbs04(facts, asset_row, wb_data):
    disc_code = _derive_discipline_code(facts, asset_row, wb_data)
    if disc_code in wb_data.wbs_maps["WBS04"]["code_to_desc"]:
        return disc_code

    desc = wb_data.discipline_code_to_desc.get(disc_code, "")
    if desc:
        code = wb_data.wbs_maps["WBS04"]["desc_to_code"].get(_normalize_text(desc), "")
        if code:
            return code

    return wb_data.wbs_maps["WBS04"]["default"] or disc_code or "ME"


def _derive_wbs05(facts, asset_row, wb_data):
    existing_code = _valid_wbs_code(facts.get("existing_wbs05", ""), wb_data.wbs_maps["WBS05"])
    wbs01 = _derive_wbs01(facts, wb_data)
    is_external = wbs01 == "EA" or "external" in facts.get("search_blob", "")
    discipline_code = _derive_discipline_code(facts, asset_row, wb_data)

    search = " ".join([
        _normalize_text(facts.get("category", "")),
        _normalize_text(facts.get("family", "")),
        _normalize_text(facts.get("type", "")),
        _normalize_text(facts.get("system", "")),
        _normalize_text(facts.get("discipline_hint_norm", "")),
        _normalize_text(_row_value(asset_row, "NWB_Category") if asset_row else ""),
        _normalize_text(_row_value(asset_row, "NWB_AssetType") if asset_row else ""),
    ])

    best_code = ""
    best_score = -1
    for desc_norm, code in wb_data.wbs_maps["WBS05"]["desc_to_code"].items():
        if not desc_norm:
            continue
        if desc_norm in search:
            score = len(desc_norm)
            if score > best_score:
                best_score = score
                best_code = code

    if best_code:
        return best_code

    service_suffix = ""
    if "communication" in search or "telephone" in search or "data" in search or "nurse call" in search:
        service_suffix = "11"
    elif "security" in search:
        service_suffix = "4"
    elif discipline_code == "MH":
        service_suffix = "1"
    elif discipline_code == "FR":
        service_suffix = "2"
    elif discipline_code == "EL":
        service_suffix = "3"
    elif discipline_code == "ME":
        service_suffix = "5"

    if service_suffix:
        service_code = "{0}{1}".format("SE" if is_external else "SI", service_suffix)
        if service_code in wb_data.wbs_maps["WBS05"]["code_to_desc"]:
            return service_code

    if existing_code:
        return existing_code

    return ""


def _compose_wbs_code(wbs00, wbs01, wbs02, wbs03, wbs04, wbs05):
    parts = [
        _safe_str(wbs00).strip(),
        _safe_str(wbs01).strip(),
        _safe_str(wbs02).strip(),
        _safe_str(wbs03).strip(),
        _safe_str(wbs04).strip(),
        _safe_str(wbs05).strip(),
    ]
    parts = [x for x in parts if x]
    return "-".join(parts)


def _derive_asset_id(element, asset_row, asset_type_code, building_code, level_code):
    if not asset_row:
        return "", "unmatched", "Asset Master row was not matched"

    source_value = _safe_str(_row_value(asset_row, "NWB_AssetID")).strip()
    normalized_source = _normalize_text(source_value)
    if source_value and "assettypecode-building-level-elementid" not in normalized_source:
        return source_value, "fixed", "fixed Asset Master NWB_AssetID"

    asset_id_rule = "template" if source_value else "default"
    asset_id_source = (
        "generated from AssetTypeCode-Building-Level-ElementID"
        if source_value
        else "generated from matched Asset Master row with blank NWB_AssetID"
    )

    asset_code = _safe_str(asset_type_code).strip().upper()
    building = _safe_str(building_code).strip().upper()
    level = _safe_str(level_code).strip().upper()
    element_id = _element_id_text(element).strip()
    if not asset_code:
        return "", asset_id_rule, "Asset ID generation requires NWB_AssetTypeCode"
    if not building:
        return "", asset_id_rule, "Asset ID generation requires a resolved WBS01 building code"
    if not level:
        return "", asset_id_rule, "Asset ID generation requires a resolved WBS02 level code"
    if not element_id or element_id == "?":
        return "", asset_id_rule, "Asset ID generation requires a valid Revit element ID"

    try:
        element_id = "{0:04d}".format(int(element_id))
    except Exception:
        pass
    return "{0}-{1}-{2}-{3}".format(asset_code, building, level, element_id), asset_id_rule, asset_id_source


def _clean_asset_value(value):
    text = _safe_str(value).strip()
    if _is_placeholder_text(text):
        return ""
    return text


def _score_work_package(pkg, discipline_code, site_code):
    score = 0
    code = _safe_str(pkg.get("code", "")).upper()
    bp_key = _safe_str(pkg.get("associated_bp_key", ""))
    sub_pkg = _safe_str(pkg.get("sub_package_norm", ""))

    if discipline_code and "-{0}-".format(discipline_code.upper()) in code:
        score += 12
    if site_code and code.startswith("NWB-{0}-".format(site_code.upper())):
        score += 7
    if site_code and bp_key.startswith(site_code.upper() + "-"):
        score += 5
    if discipline_code in ["ME", "MH", "EL", "FR", "IC", "IS", "UT"] and "service" in sub_pkg:
        score += 2

    return score


def _select_work_package(facts, asset_row, wb_data):
    if not wb_data.work_packages:
        return None

    discipline_code = _derive_discipline_code(facts, asset_row, wb_data)
    site_code = _derive_wbs00(facts, wb_data)
    best = None
    best_score = -1

    for pkg in wb_data.work_packages:
        score = _score_work_package(pkg, discipline_code, site_code)
        if score > best_score:
            best = pkg
            best_score = score

    if best_score <= 0:
        return None

    return best


def _derive_design_pkg(facts, asset_row, wb_data):
    override_value = _safe_str(_run_option_text("design_pkg_override")).strip().upper()
    if override_value:
        return override_value

    from_asset = _clean_asset_value(_row_value(asset_row, "NWB_DesignPkg") if asset_row else "")
    if from_asset:
        return from_asset

    pkg = _select_work_package(facts, asset_row, wb_data)
    if pkg is None:
        return ""

    return _safe_str(pkg.get("code", "")).strip().upper()


def _derive_building_permit_pkg(facts, asset_row, wb_data, design_pkg):
    override_value = _safe_str(_run_option_text("building_permit_override")).strip()
    if override_value:
        return _canonical_building_permit(override_value, wb_data)

    from_asset = _clean_asset_value(_row_value(asset_row, "NWB_BuildingPermitPkg") if asset_row else "")
    if from_asset:
        return _canonical_building_permit(from_asset, wb_data)

    if design_pkg:
        for pkg in wb_data.work_packages:
            if _safe_str(pkg.get("code", "")).strip().upper() == _safe_str(design_pkg).strip().upper():
                return _canonical_building_permit(pkg.get("associated_bp", ""), wb_data)

    pkg = _select_work_package(facts, asset_row, wb_data)
    if pkg is None:
        return ""

    return _canonical_building_permit(pkg.get("associated_bp", ""), wb_data)


def _build_target_values(element, wb_data):
    facts = _derive_element_facts(element)
    asset_row, asset_score = _select_asset_row(wb_data.asset_rows, facts)

    target = {}
    target["NWB_AssetTypeCode"] = _clean_asset_value(_row_value(asset_row, "NWB_AssetTypeCode") if asset_row else "")
    target["NWB_WBS01"] = _derive_wbs01(facts, wb_data)
    target["NWB_WBS02"] = _derive_wbs02(facts, wb_data)
    target["NWB_AssetID"], asset_id_rule, asset_id_reason = _derive_asset_id(
        element,
        asset_row,
        target["NWB_AssetTypeCode"],
        target["NWB_WBS01"],
        target["NWB_WBS02"],
    )
    target["NWB_AssetType"] = _clean_asset_value(_row_value(asset_row, "NWB_AssetType") if asset_row else "")

    discipline_code = _derive_discipline_code(facts, asset_row, wb_data)
    target["NWB_Discipline"] = discipline_code

    target["NWB_Category"] = _clean_asset_value(_row_value(asset_row, "NWB_Category") if asset_row else "")
    if _is_blank(target["NWB_Category"]):
        target["NWB_Category"] = _safe_str(facts.get("category", "")).strip()

    target["NWB_Department"] = _safe_str(facts.get("department", "")).strip()
    target["NWB_SubDepartment"] = _safe_str(facts.get("subdepartment", "")).strip()

    material = _clean_asset_value(_row_value(asset_row, "NWB_Material") if asset_row else "")
    if _is_blank(material):
        material = _safe_str(facts.get("material", "")).strip()
    target["NWB_Material"] = material

    system_text = _safe_str(facts.get("system", "")).strip()
    target["NWB_System"] = system_text

    target["NWB_UniclassCode"] = _clean_asset_value(_row_value(asset_row, "NWB_UniclassCode") if asset_row else "")
    target["NWB_UniclassDescription"] = _clean_asset_value(_row_value(asset_row, "NWB_UniclassDescription") if asset_row else "")

    target["NWB_DesignPkg"] = _derive_design_pkg(facts, asset_row, wb_data)
    target["NWB_BuildingPermitPkg"] = _derive_building_permit_pkg(facts, asset_row, wb_data, target["NWB_DesignPkg"])

    target["NWB_WBS00"] = _derive_wbs00(facts, wb_data)
    target["NWB_WBS03"] = _derive_wbs03(facts, wb_data)
    target["NWB_WBS04"] = _derive_wbs04(facts, asset_row, wb_data)
    target["NWB_WBS05"] = _derive_wbs05(facts, asset_row, wb_data)
    target["NWB_WBSCode"] = _compose_wbs_code(
        target["NWB_WBS00"],
        target["NWB_WBS01"],
        target["NWB_WBS02"],
        target["NWB_WBS03"],
        target["NWB_WBS04"],
        target["NWB_WBS05"],
    )

    debug = {
        "asset_score": asset_score,
        "asset_type": _row_value(asset_row, "NWB_AssetType") if asset_row else "",
        "asset_type_code": _row_value(asset_row, "NWB_AssetTypeCode") if asset_row else "",
        "asset_id_rule": asset_id_rule,
        "asset_id_reason": asset_id_reason,
        "asset_category": _row_value(asset_row, "NWB_Category") if asset_row else "",
        "elemental_code": target["NWB_WBS05"],
        "wbs_code": target["NWB_WBSCode"],
        "facts": facts,
    }
    return target, debug


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


def _build_workbook_path():
    here = os.path.dirname(__file__)
    return os.path.join(here.replace("NWB_PARAMETERS AutoFill.pushbutton", "NWB Dim AutoFill.pushbutton"), WORKBOOK_FILE_NAME)


def _validate_workbook_path(path):
    return os.path.exists(path)


def _view_label(active_view):
    if isinstance(active_view, ViewSchedule):
        return "schedule"
    return "model"


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

    _set_run_options(run_settings)

    write_mode = run_settings.get("write_mode", "smart")
    parameter_target_mode = run_settings.get("parameter_target_mode", DEFAULT_PARAMETER_TARGET_MODE)
    worksharing_mode = run_settings.get("worksharing_mode", "all")
    processing_scope = run_settings.get("processing_scope", "view")
    export_audit = bool(run_settings.get("export_audit", True))

    logger = RunLogger(os.path.dirname(__file__), TOOL_TITLE)
    audit = RunAuditExporter(logger.logs_dir, logger.safe_title, logger.stamp, export_audit, PARAM_AUDIT_COLUMNS)
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
    selected_link = _resolve_selected_arch_room_link_instance()
    logger.write("START | ArchRoomLink={0}".format(_safe_str(getattr(selected_link, "Name", "AUTO")) or "AUTO"))
    logger.write("START | SiteOverride={0}".format(_run_option_text("site_override")))
    logger.write("START | BuildingOverride={0}".format(_run_option_text("building_override")))
    logger.write("START | DesignPkgOverride={0}".format(_run_option_text("design_pkg_override")))
    logger.write("START | BuildingPermitOverride={0}".format(_run_option_text("building_permit_override")))
    logger.write("START | ExportAudit={0}".format(export_audit))
    if export_audit and audit.file_path:
        logger.write("START | AuditFile={0}".format(audit.file_path))

    monitor = None
    try:
        monitor = LiveMonitorWindow("MonitorWindow.xaml")
        monitor.Show()
        monitor.Activate()
        _monitor_set_status(monitor, "Initializing {0}...".format(TOOL_TITLE))
        _monitor_add_item(monitor, "Preparing run context...")
        _pump_ui()
    except Exception as ex:
        monitor = None
        forms.alert(
            "Live monitor window could not be opened.\n"
            "Processing will continue with progress bar only.\n\n{0}".format(_safe_str(ex)),
            title=TOOL_TITLE,
        )

    workbook_path = _build_workbook_path()
    _monitor_set_status(monitor, "Checking workbook path...")
    _monitor_add_item(monitor, "Workbook: {0}".format(workbook_path))
    logger.write("WORKBOOK | {0}".format(workbook_path))
    _pump_ui()
    if not _validate_workbook_path(workbook_path):
        logger.write("ERROR | Workbook not found")
        logger.close()
        forms.alert("Workbook not found:\n{0}".format(workbook_path), title=TOOL_TITLE)
        return

    _monitor_set_status(monitor, "Reading workbook mapping data...")
    _pump_ui()
    try:
        wb_data = _load_workbook_data(
            workbook_path,
            status_callback=lambda msg: (_monitor_set_status(monitor, msg), _pump_ui()),
        )
    except Exception as ex:
        logger.write("ERROR | Workbook read failed | {0}: {1}".format(_safe_str(type(ex).__name__), _safe_str(ex)))
        logger.close()
        forms.alert(
            "Failed to read workbook mapping data.\n\n"
            "Workbook: {0}\n"
            "Error type: {1}\n"
            "Error: {2}".format(
                workbook_path,
                _safe_str(type(ex).__name__),
                _safe_str(ex),
            ),
            title=TOOL_TITLE,
        )
        return

    _monitor_set_status(monitor, "Collecting target elements from active view...")
    _pump_ui()
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
        "skip_matching": 0,
        "skip_no_source": 0,
        "pending_verify": 0,
        "no_asset_match": 0,
        "skip_worksharing": 0,
        "type_writes": 0,
        "instance_writes": 0,
        "read_only": 0,
        "duplicate_parameter": 0,
        "skip_target_mode": 0,
        "shared_writes": 0,
        "family_writes": 0,
        "transaction_failed": 0,
        "reason_stats": {},
    }

    total = len(elements)
    chunk_size = 180
    cancelled = False
    sample_fails = []
    owner_status_cache = {}

    # OPTIMIZATION: Initialize caches for this run
    global _PARAM_RESOLUTION_CACHE, _ASSET_MATCH_CACHE
    _PARAM_RESOLUTION_CACHE = {}
    _ASSET_MATCH_CACHE = {}

    _monitor_set_status(
        monitor,
        "Ready. View mode: {0} | Elements: {1} | Asset rows: {2} | WBS03 codes: {3} | WP codes: {4} | BP items: {5}".format(
            view_mode,
            total,
            len(wb_data.asset_rows),
            len(wb_data.wbs_maps["WBS03"]["code_to_desc"]),
            len(wb_data.work_package_codes),
            len(wb_data.building_permit_items),
        )
    )
    _monitor_add_item(monitor, "Starting processing in {0} scope / {1} view mode with write mode: {2} | target mode: {3} | worksharing: {4} | log: {5}".format(processing_scope, view_mode, write_mode, parameter_target_mode, worksharing_mode, logger.file_path))
    logger.write("READY | Elements={0} | AssetRows={1} | Log={2}".format(total, len(wb_data.asset_rows), logger.file_path))
    logger.write("READY | CategoryBreakdown={0}".format(_summarize_category_counts(elements)))
    logger.write("READY | EligibleElementIds={0}".format(_summarize_element_ids(elements)))
    _pump_ui()

    with forms.ProgressBar(title="NWB parameters auto-fill {value}/{max_value}", cancellable=True) as pb:
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
                        _record_reason_stat(counters, element, "<Element>", "worksharing: {0}".format(ws_reason))
                        _monitor_add_item(monitor, "[{0}] Skipped by worksharing rule: {1}".format(counters["processed"], ws_reason), eid)
                        logger.write("SKIP | Element {0} | Worksharing | {1}".format(eid, ws_reason))
                        if audit is not None:
                            row = _build_param_audit_row(run_audit_context, element, "", "")
                            row["outcome"] = "SKIP"
                            row["reason"] = "worksharing: {0}".format(ws_reason)
                            audit.write_row(row)
                        continue

                    target_values, debug = _build_target_values(element, wb_data)
                    audit_context = _build_param_audit_context(run_audit_context, element, debug, target_values)
                    if int(debug.get("asset_score", 0)) <= 0:
                        counters["no_asset_match"] += 1
                        logger.write(
                            "NO ASSET MATCH | Element {0} | Cat={1} | Family={2} | Type={3} | System={4}".format(
                                eid,
                                debug.get("facts", {}).get("category", ""),
                                debug.get("facts", {}).get("family", ""),
                                debug.get("facts", {}).get("type", ""),
                                debug.get("facts", {}).get("system", ""),
                            )
                        )

                    statuses = []
                    for param_name in TARGET_PARAMETERS:
                        incoming = target_values.get(param_name, "")
                        st, msg = _set_if_needed(element, param_name, incoming, write_mode, parameter_target_mode, counters, pending_writes, logger, audit, audit_context)
                        statuses.append("{0}:{1}".format(param_name, st))
                        if st in ("FAIL", "PARTIAL") and len(sample_fails) < 12:
                            sample_fails.append("Element {0} | {1} | {2}".format(eid, param_name, msg))

                    _log_element_runtime_context(logger, element, eid, debug.get("facts", {}) or {}, target_values, debug)

                    logger.write(
                        "ELEMENT | {0} | Cat={1} | AssetScore={2} | AssetTypeCode={3} | AssetType={4} | AssetCategory={5} | Targets={6}".format(
                            eid,
                            debug.get("facts", {}).get("category", ""),
                            debug.get("asset_score", 0),
                            debug.get("asset_type_code", ""),
                            debug.get("asset_type", ""),
                            debug.get("asset_category", ""),
                            "; ".join(["{0}={1}".format(x, target_values.get(x, "")) for x in TARGET_PARAMETERS]),
                        )
                    )

                    if counters["processed"] <= 20 or (counters["processed"] % 10 == 0):
                        _monitor_add_item(monitor, "[{0}] AssetScore={1} | {2}".format(
                            counters["processed"],
                            debug.get("asset_score", 0),
                            " ".join(statuses[:5])
                        ), eid)

                    if counters["processed"] % 10 == 0 or counters["processed"] == total:
                        pb.update_progress(counters["processed"], total)
                        _monitor_set_status(
                            monitor,
                            "Processed {0}/{1} | Written: {2} | Already correct: {3} | Missing target: {4} | No source: {5} | Worksharing skip: {6} | Failed: {7}".format(
                                counters["processed"],
                                total,
                                counters["written"],
                                counters["skip_matching"],
                                counters["missing_target"],
                                counters["skip_no_source"],
                                counters["skip_worksharing"],
                                counters["failed"],
                            )
                        )

                    if counters["processed"] % 20 == 0:
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
        "Workbook: {0}".format(workbook_path),
        "Log file: {0}".format(logger.file_path),
        "Audit file: {0}".format(audit.file_path if audit.file_path else "disabled"),
        "Workbook checks: Asset rows={0}, WBS03 codes={1}, Work Packages={2}, Building Permit={3}".format(
            len(wb_data.asset_rows),
            len(wb_data.wbs_maps["WBS03"]["code_to_desc"]),
            len(wb_data.work_package_codes),
            len(wb_data.building_permit_items),
        ),
        "Processed elements: {0}/{1}".format(counters["processed"], total),
        "Writes attempted: {0}".format(counters["write_attempted"]),
        "Parameter writes: {0}".format(counters["written"]),
        "Writes skipped: {0}".format(counters["skip_matching"] + counters["skip_no_source"] + counters["missing_target"] + counters["skip_worksharing"] + counters["skip_target_mode"]),
        "Already correct: {0}".format(counters["skip_matching"]),
        "Skipped no source value: {0}".format(counters["skip_no_source"]),
        "Missing target parameter: {0}".format(counters["missing_target"]),
        "Skipped by target mode policy: {0}".format(counters["skip_target_mode"]),
        "Skipped by worksharing rule: {0}".format(counters["skip_worksharing"]),
        "Elements with no asset-row match: {0}".format(counters["no_asset_match"]),
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
        summary.extend(["- {0}".format(x) for x in sample_fails])

    reason_summary_lines = _top_reason_summary_lines(counters.get("reason_stats") or {})
    if reason_summary_lines:
        logger.write("SUMMARY | Top unresolved / skipped reasons")
        for line in reason_summary_lines:
            logger.write("SUMMARY | {0}".format(line))
        summary.append("")
        summary.append("Top unresolved / skipped reasons:")
        summary.extend(["- {0}".format(x) for x in reason_summary_lines[:8]])

    logger.write("END | Attempted={0} | Written={1} | Skipped={2} | VerifyFailed={3} | SharedWrites={4} | FamilyWrites={5} | TypeWrites={6} | InstanceWrites={7} | ReadOnly={8} | MissingTarget={9} | SkipTargetMode={10} | DuplicateParameter={11} | SkipWorksharing={12} | Failed={13} | Cancelled={14}".format(
        counters["write_attempted"],
        counters["written"],
        counters["skip_matching"] + counters["skip_no_source"] + counters["missing_target"] + counters["skip_worksharing"] + counters["skip_target_mode"],
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
    ))
    logger.close()
    audit.close()

    forms.alert("\n".join(summary), title=TOOL_TITLE)


if __name__ == "__main__":
    try:
        run()
    except Exception as ex:
        forms.alert("{0} failed.\n\n{1}".format(TOOL_TITLE, _safe_str(ex)), title=TOOL_TITLE)
