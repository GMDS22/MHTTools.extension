# coding: utf8
from __future__ import print_function

from collections import defaultdict
import csv
import io
import math
import os
import re
import tempfile
import time
from datetime import datetime

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    Level,
    RevitLinkInstance,
    SpatialElementBoundaryOptions,
    StorageType,
    Transaction,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow
from System.Windows.Controls import CheckBox
import System
from Microsoft.Win32 import SaveFileDialog


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
config = script.get_config()

__title__ = "Linked Room\nParameter Transfer"
__doc__ = "Select a room from a linked model and transfer room parameter values into selected host elements."

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


TARGET_CATEGORIES = [
    # MEP Spaces & Zones
    ("Spaces", BuiltInCategory.OST_MEPSpaces),
    ("Rooms", BuiltInCategory.OST_Rooms),
    ("HVAC Zones", BuiltInCategory.OST_HVAC_Zones),
    # Ducts
    ("Ducts", BuiltInCategory.OST_DuctCurves),
    ("Duct Fittings", BuiltInCategory.OST_DuctFitting),
    ("Duct Accessories", BuiltInCategory.OST_DuctAccessory),
    ("Duct Insulations", BuiltInCategory.OST_DuctInsulations),
    ("Flex Ducts", BuiltInCategory.OST_FlexDuctCurves),
    ("Air Terminals", BuiltInCategory.OST_DuctTerminal),
    ("Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment),
    # Pipes
    ("Pipes", BuiltInCategory.OST_PipeCurves),
    ("Pipe Fittings", BuiltInCategory.OST_PipeFitting),
    ("Pipe Accessories", BuiltInCategory.OST_PipeAccessory),
    ("Pipe Insulations", BuiltInCategory.OST_PipeInsulations),
    ("Flex Pipes", BuiltInCategory.OST_FlexPipeCurves),
    ("Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures),
    ("Sprinklers", BuiltInCategory.OST_Sprinklers),
    # Electrical
    ("Cable Trays", BuiltInCategory.OST_CableTray),
    ("Cable Tray Fittings", BuiltInCategory.OST_CableTrayFitting),
    ("Conduits", BuiltInCategory.OST_Conduit),
    ("Conduit Fittings", BuiltInCategory.OST_ConduitFitting),
    ("Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment),
    ("Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures),
    ("Lighting Fixtures", BuiltInCategory.OST_LightingFixtures),
    ("Lighting Devices", BuiltInCategory.OST_LightingDevices),
    # Low Voltage Devices
    ("Data Devices", BuiltInCategory.OST_DataDevices),
    ("Communication Devices", BuiltInCategory.OST_CommunicationDevices),
    ("Fire Alarm Devices", BuiltInCategory.OST_FireAlarmDevices),
    ("Nurse Call Devices", BuiltInCategory.OST_NurseCallDevices),
    ("Security Devices", BuiltInCategory.OST_SecurityDevices),
    ("Telephone Devices", BuiltInCategory.OST_TelephoneDevices),
    # Architectural / General
    ("Generic Models", BuiltInCategory.OST_GenericModel),
    ("Speciality Equipment", BuiltInCategory.OST_SpecialityEquipment),
    ("Casework", BuiltInCategory.OST_Casework),
    ("Furniture", BuiltInCategory.OST_Furniture),
    ("Furniture Systems", BuiltInCategory.OST_FurnitureSystems),
    ("Medical Equipment", BuiltInCategory.OST_MedicalEquipment),
]


def _load_persistent_settings():
    """Load all persistent settings from config."""
    settings = {
        "selected_categories": getattr(config, "selected_categories", []),
        "probe_offset_mm": getattr(config, "probe_offset_mm", 2000),
        "use_selected_room_fallback": getattr(config, "use_selected_room_fallback", False),
        "allow_type_writes_in_auto_room": getattr(config, "allow_type_writes_in_auto_room", False),
        "safe_mode": getattr(config, "safe_mode", True),
        "auto_match_exclude_keywords": getattr(config, "auto_match_exclude_keywords", "uniclass,class"),
        "export_mapped_preview_csv": getattr(config, "export_mapped_preview_csv", False),
        "last_preview_export_dir": getattr(config, "last_preview_export_dir", ""),
        "preview_height": getattr(config, "preview_height", 220),
    }
    return settings


def _save_persistent_settings(settings):
    """Save all persistent settings to config."""
    try:
        config.selected_categories = settings.get("selected_categories", [])
        config.probe_offset_mm = settings.get("probe_offset_mm", 2000)
        config.use_selected_room_fallback = settings.get("use_selected_room_fallback", False)
        config.allow_type_writes_in_auto_room = settings.get("allow_type_writes_in_auto_room", False)
        config.safe_mode = settings.get("safe_mode", True)
        config.auto_match_exclude_keywords = settings.get("auto_match_exclude_keywords", "uniclass,class")
        config.export_mapped_preview_csv = bool(settings.get("export_mapped_preview_csv", False))
        config.last_preview_export_dir = settings.get("last_preview_export_dir", "")
        config.preview_height = settings.get("preview_height", 220)
        script.save_config()
    except Exception:
        pass


def _format_duration_ms(ms_value):
    try:
        ms = int(ms_value)
    except Exception:
        ms = 0
    if ms < 0:
        ms = 0
    sec = float(ms) / 1000.0
    return "~{0:.2f}s ({1} ms)".format(sec, ms)


def normalize_name(name):
    if not name:
        return ""
    return "".join(ch for ch in name.lower() if ch.isalnum())


def get_link_instances(active_doc):
    return list(FilteredElementCollector(active_doc).OfClass(RevitLinkInstance))


def read_parameter_value(param, source_doc=None):
    if param is None:
        return None
    src_doc = source_doc or doc
    st = param.StorageType
    try:
        has_value = bool(param.HasValue)
    except Exception:
        has_value = False

    raw_value = None
    try:
        if st == StorageType.String:
            raw_value = param.AsString()
        elif st == StorageType.Integer:
            raw_value = param.AsInteger()
            if not has_value and raw_value == 0:
                raw_value = None
        elif st == StorageType.Double:
            raw_value = param.AsDouble()
            if not has_value and abs(float(raw_value)) <= 1e-12:
                raw_value = None
        elif st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None:
                raw_value = None
            else:
                try:
                    val = eid.IntegerValue
                    if val == ElementId.InvalidElementId.IntegerValue:
                        raw_value = None
                    else:
                        ref_el = None
                        try:
                            ref_el = src_doc.GetElement(eid)
                        except Exception:
                            ref_el = None

                        if ref_el is not None:
                            try:
                                ref_name = getattr(ref_el, "Name", None)
                                if ref_name:
                                    raw_value = ref_name
                                else:
                                    raw_value = val
                            except Exception:
                                raw_value = val
                        else:
                            raw_value = val
                except Exception:
                    try:
                        raw_value = int(eid)
                    except Exception:
                        raw_value = None
    except Exception:
        raw_value = None

    if raw_value not in (None, ""):
        return raw_value

    # Some linked-room parameters can present values in the Properties UI while
    # HasValue is false or storage readers return empty. Fall back to display text.
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


def _display_room_param_value(param, raw_value, source_doc=None):
    src_doc = source_doc or doc
    try:
        st = param.StorageType
    except Exception:
        st = None

    if st == StorageType.ElementId:
        try:
            eid = param.AsElementId()
            if eid is not None and eid != ElementId.InvalidElementId:
                try:
                    ref_el = src_doc.GetElement(eid)
                except Exception:
                    ref_el = None
                if ref_el is not None:
                    try:
                        ref_name = getattr(ref_el, "Name", None)
                        if ref_name:
                            return str(ref_name).strip()
                    except Exception:
                        pass
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

    if raw_value not in (None, ""):
        try:
            if isinstance(raw_value, str):
                s = raw_value.strip()
                if s:
                    return s
                return None
        except Exception:
            pass
        return raw_value

    return None


def _is_empty_value(value):
    if value is None:
        return True
    try:
        return str(value).strip() == ""
    except Exception:
        return False


def get_writable_parameters(element):
    names = set()
    storage_none = getattr(StorageType, "None")
    if element is None:
        return names

    owners = [element]
    try:
        symbol = getattr(element, "Symbol", None)
        if symbol is not None:
            owners.append(symbol)
        else:
            type_id = getattr(element, "GetTypeId", lambda: None)()
            if type_id is not None:
                type_el = doc.GetElement(type_id)
                if type_el is not None:
                    owners.append(type_el)
    except Exception:
        pass

    for owner in owners:
        try:
            for p in owner.Parameters:
                if p is None or p.Definition is None:
                    continue
                if p.IsReadOnly or p.StorageType == storage_none:
                    continue
                names.add(p.Definition.Name)
        except Exception:
            pass

    return names


def list_writable_parameter_names(element, max_items=50):
    names = get_writable_parameters(element)
    if not names:
        return []
    return sorted(names)[:max_items]


def values_equal(left, right, storage_type):
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    if storage_type == StorageType.Double:
        try:
            return abs(float(left) - float(right)) <= 1e-6
        except Exception:
            return False
    if storage_type == StorageType.Integer:
        try:
            return int(left) == int(right)
        except Exception:
            return False
    if storage_type == StorageType.ElementId:
        try:
            return int(left) == int(right)
        except Exception:
            return False
    return str(left).strip() == str(right).strip()


def set_parameter_value(param, value, duplicate_mode):
    if param is None:
        return False, "missing parameter"
    if param.IsReadOnly:
        return False, "read-only parameter"
    if value is None:
        return False, "empty value"

    try:
        existing = read_parameter_value(param)
        # Safe guard: treat identical values as a no-op, not a successful write.
        try:
            if values_equal(existing, value, param.StorageType):
                return False, "no change (already equal)"
        except Exception:
            pass

        if existing not in (None, "", 0):
            if duplicate_mode == "Skip":
                return False, "existing value skipped"
            if duplicate_mode == "Append" and param.StorageType == StorageType.String:
                old_s = existing if isinstance(existing, str) else str(existing)
                new_s = value if isinstance(value, str) else str(value)
                success = param.Set(old_s + "; " + new_s)
                if not success and hasattr(param, "SetValueString"):
                    try:
                        success = param.SetValueString(old_s + "; " + new_s)
                    except Exception:
                        success = False
                if not success:
                    return False, "failed to append value"
                try:
                    doc.Regenerate()
                except Exception:
                    pass
                verified = read_parameter_value(param)
                if not values_equal(verified, old_s + "; " + new_s, param.StorageType):
                    return False, "append not persisted"
                return True, "appended"

        st = param.StorageType
        if st == StorageType.String:
            value_str = value if isinstance(value, str) else str(value)
            success = param.Set(value_str)
            if not success and hasattr(param, "SetValueString"):
                try:
                    success = param.SetValueString(value_str)
                except Exception:
                    success = False
        elif st == StorageType.Integer:
            try:
                if isinstance(value, str):
                    value = value.strip()
                success = param.Set(int(value))
            except Exception:
                try:
                    success = param.Set(int(float(str(value).strip())))
                except Exception:
                    return False, "invalid integer value"
        elif st == StorageType.Double:
            try:
                if isinstance(value, str):
                    value = value.strip().replace(",", ".")
                success = param.Set(float(value))
            except Exception:
                return False, "invalid double value"
        elif st == StorageType.ElementId:
            # Best-effort handling: accept integer ids or numeric strings.
            try:
                if value is None:
                    return False, "empty element id value"
                if isinstance(value, (int, long)) if 'long' in globals() else isinstance(value, int):
                    eid = ElementId(int(value))
                    success = param.Set(eid)
                else:
                    sval = str(value).strip()
                    if sval.isdigit():
                        eid = ElementId(int(sval))
                        success = param.Set(eid)
                    else:
                        return False, "unsupported ElementId value"
            except Exception:
                return False, "invalid element id value"
        else:
            return False, "unsupported storage type"

        if not success:
            return False, "set returned False"

        try:
            doc.Regenerate()
        except Exception:
            pass

        verified = read_parameter_value(param)
        if not values_equal(verified, value, st):
            if st == StorageType.String and hasattr(param, "SetValueString"):
                try:
                    param.SetValueString(value if isinstance(value, str) else str(value))
                except Exception:
                    pass
                try:
                    doc.Regenerate()
                except Exception:
                    pass
                verified = read_parameter_value(param)
            if not values_equal(verified, value, st):
                return False, "write verification failed (expected {0}, got {1})".format(value, verified)

        return True, "written"
    except Exception as ex:
        return False, str(ex)


def get_best_match(room_param_name, target_param_names, strictness_mode="balanced"):
    if not target_param_names:
        return None

    room_raw = (room_param_name or "").strip()
    if not room_raw:
        return None

    if strictness_mode == "disabled":
        return None

    # Strict: exact case-sensitive, then exact case-insensitive.
    if strictness_mode == "strict":
        for name in target_param_names:
            if room_raw == name:
                return name
        room_lower = room_raw.lower()
        for name in target_param_names:
            if room_lower == (name or "").lower():
                return name
        return None

    room_norm = normalize_name(room_raw)
    if not room_norm:
        return None

    by_norm = {}
    for name in target_param_names:
        by_norm[normalize_name(name)] = name

    # Balanced: normalized exact only.
    if room_norm in by_norm:
        return by_norm[room_norm]

    if strictness_mode == "balanced":
        return None

    # Loose: allow contains similarity after normalized exact miss.
    for norm_name, raw_name in by_norm.items():
        if room_norm in norm_name or norm_name in room_norm:
            return raw_name

    return None


def _find_writable_parameter(target_element, target_param_name):
    if target_element is None or not target_param_name:
        return None, None

    name = target_param_name.strip()
    normalized_target_name = normalize_name(name)
    search_sources = []

    try:
        search_sources.append(("instance", target_element))
    except Exception:
        pass

    try:
        symbol = getattr(target_element, "Symbol", None)
        if symbol is not None:
            search_sources.append(("type", symbol))
    except Exception:
        pass

    read_only_found = False
    storage_none = getattr(StorageType, "None")

    for source_name, owner in search_sources:
        try:
            param = owner.LookupParameter(name)
            if param is not None:
                if not param.IsReadOnly and param.StorageType != storage_none:
                    return param, source_name
                read_only_found = True
        except Exception:
            pass

        try:
            params = list(owner.GetParameters(name))
        except Exception:
            params = []

        for param in params:
            if param is None:
                continue
            if not param.IsReadOnly and param.StorageType != storage_none:
                return param, source_name
            read_only_found = True

    if normalized_target_name:
        for source_name, owner in search_sources:
            try:
                for param in owner.Parameters:
                    if param is None or param.Definition is None:
                        continue
                    if normalize_name(param.Definition.Name) != normalized_target_name:
                        continue
                    if not param.IsReadOnly and param.StorageType != storage_none:
                        return param, source_name
                    read_only_found = True
            except Exception:
                pass

    if read_only_found:
        return None, "read-only"
    return None, None


def _find_parameter_for_verification(target_element, target_param_name, target_source=None):
    """Find parameter again from document state (instance/type) for post-commit verification."""
    if target_element is None or not target_param_name:
        return None

    name = target_param_name.strip()
    normalized_target_name = normalize_name(name)
    owners = []

    if target_source == "instance":
        owners = [target_element]
    elif target_source == "type":
        try:
            symbol = getattr(target_element, "Symbol", None)
            if symbol is not None:
                owners = [symbol]
            else:
                type_id = getattr(target_element, "GetTypeId", lambda: None)()
                if type_id is not None:
                    type_el = doc.GetElement(type_id)
                    if type_el is not None:
                        owners = [type_el]
        except Exception:
            owners = []
    else:
        owners = [target_element]
        try:
            symbol = getattr(target_element, "Symbol", None)
            if symbol is not None:
                owners.append(symbol)
        except Exception:
            pass

    for owner in owners:
        if owner is None:
            continue

        try:
            p = owner.LookupParameter(name)
            if p is not None:
                return p
        except Exception:
            pass

        try:
            plist = list(owner.GetParameters(name))
        except Exception:
            plist = []
        if plist:
            return plist[0]

        if normalized_target_name:
            try:
                for p in owner.Parameters:
                    if p is None or p.Definition is None:
                        continue
                    if normalize_name(p.Definition.Name) == normalized_target_name:
                        return p
            except Exception:
                pass

    return None


class CategorySelectionFilter(ISelectionFilter):
    def __init__(self, allowed_category_ids):
        self.allowed_category_ids = set(allowed_category_ids)

    def AllowElement(self, element):
        try:
            cat = element.Category
            if cat is None:
                return False
            return cat.Id.IntegerValue in self.allowed_category_ids
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return True


class SourceCategorySelectionFilter(ISelectionFilter):
    def __init__(self, category_id):
        self.category_id = int(category_id)

    def AllowElement(self, element):
        try:
            if element is None or element.Category is None:
                return False
            return element.Category.Id.IntegerValue == self.category_id
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return True


class LinkedRoomTransferWindow(WPFWindow):
    # Boundary extraction is expensive and has caused instability in some models.
    # Keep direct Revit room containment as the primary method.
    _ENABLE_BOUNDARY_FALLBACK = False
    _MAX_INDEXED_SOURCES = 50000
    _SAFE_MAX_INDEXED_SOURCES = 12000
    _SAFE_MAX_TARGET_ELEMENTS = 12000
    _SAFE_MAX_TRANSFER_ELEMENTS = 3000

    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)

        self.links = []
        self.selected_room = None
        self.selected_room_doc = None
        self.selected_room_link_inst = None
        self.selected_rooms = []
        self.selected_room_params = {}
        self.selected_elements = []
        self.common_params = set()
        self.mapping = {}
        self.mapping_auto_generated = False
        self.category_items = []
        self.last_transfer_summary = []
        self.element_room_map = {}
        self.room_detection_index = []
        self.room_detection_grid = {}
        self.room_detection_grid_fallback = []
        self.room_detection_grid_fallback_by_link = {}
        self.room_detection_link_contexts = {}
        self.room_detection_grid_size = 20.0
        self.room_detection_scope_mode = None
        self.room_detection_link_id = None
        self.room_index_capped = False
        self.runtime_breadcrumbs = []
        self.detect_levels = []
        self.selected_link_index = 0
        self.auto_map_target_filter_items = []

        self._load_links()
        if not self.links:
            return
        self._populate_link_selector()
        self._populate_detect_level_selector()
        self._build_category_list()
        self.persistent_settings = _load_persistent_settings()
        self._apply_persistent_settings()
        self.source_mode_changed(None, None)
        self.room_detection_mode_changed(None, None)
        self._apply_responsive_panel_heights()
        self._try_seed_current_selection()

    def _load_links(self):
        self.links = [lk for lk in get_link_instances(doc) if lk.GetLinkDocument() is not None]
        if not self.links:
            forms.alert("No loaded Revit links found. Load at least one linked model and retry.")
            self.Close()
            return

    def _confirm(self, message, title="Confirm"):
        """Show a safe yes/no confirmation that works across pyRevit versions."""
        try:
            result = forms.alert(message, title=title, yes=True, no=True)
            if isinstance(result, bool):
                return result
            text = (str(result) if result is not None else "").strip().lower()
            return text in ("yes", "y", "true", "ok", "1")
        except Exception:
            return False

    def _populate_link_selector(self):
        """Populate the link selector dropdown with available linked models."""
        try:
            combo = getattr(self, "cmbSelectLink", None)
            if combo is None:
                return
            
            combo.Items.Clear()
            combo.Items.Add("All Linked Models")
            for link_inst in self.links:
                try:
                    link_name = link_inst.Name
                except Exception:
                    link_name = "Link"
                combo.Items.Add(link_name)
            
            if self.links:
                combo.SelectedIndex = 0
                self.selected_link_index = 0
        except Exception:
            pass

    def _get_source_mode(self):
        try:
            combo = getattr(self, "cmbSourceMode", None)
            if combo is not None and combo.SelectedItem is not None:
                item = combo.SelectedItem
                text = str(item.Content) if hasattr(item, "Content") else str(item)
                t = (text or "").strip().lower()
                if "linked" in t and "space" in t:
                    return "linked-spaces"
                if "linked" in t and "room" in t:
                    return "linked-rooms"
                if "project" in t and "space" in t:
                    return "host-spaces"
                if "project" in t and "room" in t:
                    return "host-rooms"
        except Exception:
            pass
        return "linked-rooms"

    def _is_linked_source_mode(self):
        return self._get_source_mode().startswith("linked-")

    def _get_source_category_id(self):
        mode = self._get_source_mode()
        if mode.endswith("spaces"):
            return int(BuiltInCategory.OST_MEPSpaces)
        return int(BuiltInCategory.OST_Rooms)

    def _get_source_bic(self):
        mode = self._get_source_mode()
        if mode.endswith("spaces"):
            return BuiltInCategory.OST_MEPSpaces
        return BuiltInCategory.OST_Rooms

    def _get_source_kind_name(self):
        mode = self._get_source_mode()
        return "Space" if mode.endswith("spaces") else "Room"

    def source_mode_changed(self, sender, e):
        try:
            linked_mode = self._is_linked_source_mode()
            if hasattr(self, "cmbSelectLink") and self.cmbSelectLink is not None:
                self.cmbSelectLink.IsEnabled = linked_mode
            kind = self._get_source_kind_name()
            if hasattr(self, "btnDetectRooms") and self.btnDetectRooms is not None:
                self.btnDetectRooms.Content = "Detect {0}s".format(kind)
            if hasattr(self, "btnPickRooms") and self.btnPickRooms is not None:
                self.btnPickRooms.Content = "Pick {0}(s) In Current View".format(kind)
            if hasattr(self, "txtSelectedSourceLabel") and self.txtSelectedSourceLabel is not None:
                self.txtSelectedSourceLabel.Text = "Selected {0}".format(kind)
        except Exception:
            pass

    def _populate_detect_level_selector(self):
        try:
            combo = getattr(self, "cmbDetectLevel", None)
            if combo is None:
                return

            try:
                levels = list(
                    FilteredElementCollector(doc)
                    .OfClass(Level)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                levels = []

            def _level_sort_key(lvl):
                try:
                    return (float(lvl.Elevation), (lvl.Name or "").lower())
                except Exception:
                    return (0.0, "")

            self.detect_levels = sorted(levels, key=_level_sort_key)

            combo.Items.Clear()
            combo.Items.Add("Active View Level")
            for lvl in self.detect_levels:
                try:
                    combo.Items.Add("{0}".format(lvl.Name))
                except Exception:
                    combo.Items.Add("Level")

            combo.SelectedIndex = 0
        except Exception:
            pass

    def _is_per_level_detection_mode(self):
        mode = (self._get_room_detection_mode() or "").lower()
        return "per level" in mode

    def room_detection_mode_changed(self, sender, e):
        try:
            combo = getattr(self, "cmbDetectLevel", None)
            if combo is not None:
                combo.IsEnabled = self._is_per_level_detection_mode()
        except Exception:
            pass

    def detect_level_changed(self, sender, e):
        # Intentionally lightweight; selection is consumed at detect time.
        pass

        try:
            self.room_detection_index = []
            self.room_detection_scope_mode = None
            self.room_detection_link_id = None
            self._reset_selected_room()
            self.room_detection_mode_changed(None, None)
        except Exception:
            pass

    def _build_category_list(self):
        self.category_items = []
        for name, bic in TARGET_CATEGORIES:
            cb = CheckBox()
            cb.Content = name
            cb.IsChecked = True
            self.category_items.append({
                "name": name,
                "bic": bic,
                "cb": cb,
            })

        self._refresh_category_list("")

    def _apply_persistent_settings(self):
        """Apply loaded persistent settings to the UI controls."""
        settings = getattr(self, "persistent_settings", {}) or {}

        saved_categories = settings.get("selected_categories", [])
        if saved_categories:
            saved_set = set(saved_categories)
            for item in self.category_items:
                try:
                    item["cb"].IsChecked = item["name"] in saved_set
                except Exception:
                    item["cb"].IsChecked = False
        else:
            for item in self.category_items:
                item["cb"].IsChecked = True

        try:
            text = settings.get("probe_offset_mm", 2000)
            self.txtProbeOffset.Text = str(int(float(text)))
        except Exception:
            try:
                self.txtProbeOffset.Text = "2000"
            except Exception:
                pass

        try:
            self.chkUseSelectedRoomFallback.IsChecked = bool(settings.get("use_selected_room_fallback", False))
        except Exception:
            pass
        try:
            self.chkAllowTypeWritesInAutoRoomMode.IsChecked = bool(settings.get("allow_type_writes_in_auto_room", False))
        except Exception:
            pass
        try:
            self.chkSafeMode.IsChecked = bool(settings.get("safe_mode", True))
        except Exception:
            pass
        try:
            self.txtAutoMatchExclude.Text = str(settings.get("auto_match_exclude_keywords", "uniclass,class") or "")
        except Exception:
            pass
        try:
            self.chkExportMappedPreviewCsv.IsChecked = bool(settings.get("export_mapped_preview_csv", False))
        except Exception:
            pass
        try:
            preview_h = float(settings.get("preview_height", 220))
            if preview_h < 120:
                preview_h = 120
            if preview_h > 800:
                preview_h = 800
            self.sldPreviewHeight.Value = preview_h
        except Exception:
            pass

    def _save_persistent_settings_now(self):
        settings = {
            "selected_categories": [
                item["name"] for item in self.category_items if item["cb"].IsChecked
            ],
            "probe_offset_mm": self._get_probe_offset_mm(),
            "use_selected_room_fallback": bool(getattr(self, "chkUseSelectedRoomFallback", None) and self.chkUseSelectedRoomFallback.IsChecked),
            "allow_type_writes_in_auto_room": bool(getattr(self, "chkAllowTypeWritesInAutoRoomMode", None) and self.chkAllowTypeWritesInAutoRoomMode.IsChecked),
            "safe_mode": bool(getattr(self, "chkSafeMode", None) and self.chkSafeMode.IsChecked),
            "auto_match_exclude_keywords": str(getattr(getattr(self, "txtAutoMatchExclude", None), "Text", "") or ""),
            "export_mapped_preview_csv": bool(getattr(self, "chkExportMappedPreviewCsv", None) and self.chkExportMappedPreviewCsv.IsChecked),
            "last_preview_export_dir": str(getattr(config, "last_preview_export_dir", "") or ""),
            "preview_height": int(float(getattr(getattr(self, "sldPreviewHeight", None), "Value", 220) or 220)),
        }
        _save_persistent_settings(settings)

    def _get_probe_offset_mm(self):
        try:
            value = str(self.txtProbeOffset.Text).strip()
            if not value:
                return 2000.0
            return float(value)
        except Exception:
            return 2000.0

    def _is_safe_mode(self):
        try:
            chk = getattr(self, "chkSafeMode", None)
            if chk is not None:
                return bool(chk.IsChecked)
        except Exception:
            pass
        return True

    def _estimate_transfer_duration_ms(self, element_count, mapping_count, auto_room_mode):
        """Rough runtime estimate for user guidance before transfer starts."""
        try:
            ecount = max(0, int(element_count))
            mcount = max(0, int(mapping_count))
        except Exception:
            return 0

        # Calibrated heuristic: auto-room adds lookup overhead per mapping.
        per_map_ms = 1.35 if auto_room_mode else 0.9
        base_ms = 180.0
        est = base_ms + (ecount * mcount * per_map_ms)
        if self._is_safe_mode():
            est *= 1.1
        return int(max(0.0, est))

    def window_size_changed(self, sender, e):
        try:
            self._apply_responsive_panel_heights()
        except Exception:
            pass

    def _apply_responsive_panel_heights(self):
        """Scale key list areas with window height so panels expand when maximized."""
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

        list_h = _clamp((h * 0.22), 180.0, 430.0)
        map_h = _clamp((h * 0.18), 120.0, 340.0)
        preview_h = _clamp((h * 0.20), 150.0, 520.0)
        try:
            slider_h = float(getattr(getattr(self, "sldPreviewHeight", None), "Value", preview_h) or preview_h)
            preview_h = _clamp(slider_h, 120.0, 800.0)
        except Exception:
            pass

        for name, val in [
            ("lstCategories", list_h),
            ("lstRoomParameters", list_h),
            ("lstCommonParams", list_h),
            ("lstAutoMapTargetFilter", list_h),
            ("lstMappings", map_h),
            ("lstPreview", preview_h),
        ]:
            try:
                ctrl = getattr(self, name, None)
                if ctrl is not None:
                    ctrl.Height = val
            except Exception:
                pass

    def preview_height_changed(self, sender, e):
        try:
            h = float(getattr(getattr(self, "sldPreviewHeight", None), "Value", 220) or 220)
            if hasattr(self, "lstPreview") and self.lstPreview is not None:
                self.lstPreview.Height = max(120.0, min(800.0, h))
            self._save_persistent_settings_now()
        except Exception:
            pass

    def _breadcrumb(self, message):
        try:
            stamp = datetime.now().strftime("%H:%M:%S")
            line = "[{0}] {1}".format(stamp, message)
            self.runtime_breadcrumbs.append(line)
            if len(self.runtime_breadcrumbs) > 80:
                self.runtime_breadcrumbs = self.runtime_breadcrumbs[-80:]
            logger.info(line)
        except Exception:
            pass

    def _try_seed_current_selection(self):
        selected_ids = uidoc.Selection.GetElementIds()
        if not selected_ids:
            return

        allowed_ids = set(self._selected_category_ids())
        elements = []
        for eid in selected_ids:
            el = doc.GetElement(eid)
            if el is None or el.Category is None:
                continue
            if el.Category.Id.IntegerValue in allowed_ids:
                elements.append(el)

        if elements:
            self.selected_elements = elements
            self.element_room_map = {}
            self._set_element_summary()
            self._refresh_target_parameters()

    def _refresh_category_list(self, search_text):
        self.lstCategories.Items.Clear()
        q = (search_text or "").strip().lower()
        for item in self.category_items:
            if q and q not in item["name"].lower():
                continue
            self.lstCategories.Items.Add(item["cb"])

    def _selected_category_ids(self):
        ids = []
        for item in self.category_items:
            try:
                if item["cb"].IsChecked:
                    ids.append(int(item["bic"]))
            except Exception:
                continue
        return ids

    def _reset_selected_room(self):
        self.selected_room = None
        self.selected_room_doc = None
        self.selected_room_link_inst = None
        self.selected_rooms = []
        self.selected_room_params = {}
        try:
            if hasattr(self, "txtRoomInfo") and self.txtRoomInfo is not None:
                self.txtRoomInfo.Text = "No source selected. Use Step 1 Detect or Pick Source(s) In Current View."
            if hasattr(self, "pnlRoomInfo") and self.pnlRoomInfo is not None:
                self.pnlRoomInfo.Visibility = System.Windows.Visibility.Collapsed
            if hasattr(self, "lstRoomParameters") and self.lstRoomParameters is not None:
                self.lstRoomParameters.Items.Clear()
            if hasattr(self, "lstPreview") and self.lstPreview is not None:
                self.lstPreview.Items.Clear()
        except Exception:
            pass

    def _set_selected_rooms(self, room_items):
        """Persist selected source elements and use first source as mapping reference."""
        self.selected_rooms = list(room_items or [])
        if not self.selected_rooms:
            self._reset_selected_room()
            return

        link_inst, link_doc, room = self.selected_rooms[0]
        self.selected_room = room
        self.selected_room_doc = link_doc
        self.selected_room_link_inst = link_inst
        self.element_room_map = {}

        room_count = len(self.selected_rooms)
        header = self._room_header_text(room, link_doc, link_inst)
        if room_count > 1:
            header = "{0}\n{1}s selected: {2}".format(header, self._get_source_kind_name(), room_count)

        self.txtRoomInfo.Text = header
        self.pnlRoomInfo.Visibility = System.Windows.Visibility.Visible
        self._extract_room_parameters()
        self._refresh_target_parameters()

    def _refresh_mapping_controls(self):
        room_params = sorted(self.selected_room_params.keys())
        target_params = sorted(self.common_params)

        self.cmbMappingRoom.ItemsSource = room_params
        self.cmbMappingTarget.ItemsSource = target_params

        if room_params:
            self.cmbMappingRoom.SelectedIndex = 0
        else:
            self.cmbMappingRoom.SelectedIndex = -1

        if target_params:
            self.cmbMappingTarget.SelectedIndex = 0
        else:
            self.cmbMappingTarget.SelectedIndex = -1

    def _render_mapping_list(self):
        self.lstMappings.Items.Clear()
        if not self.mapping:
            return

        for room_param in sorted(self.mapping.keys()):
            target_param = self.mapping.get(room_param)
            if not target_param:
                continue

            self.lstMappings.Items.Add("[mapped] {0} -> {1}".format(room_param, target_param))

    def _refresh_preview_list(self):
        self.lstPreview.Items.Clear()
        rows = self._build_mapped_preview_rows(max_rows=250)
        for row in rows:
            self.lstPreview.Items.Add(
                "Element {0}: {1} -> {2} = {3}".format(
                    row.get("element_id", ""),
                    row.get("room_param", ""),
                    row.get("target_param", ""),
                    row.get("value", "<empty>"),
                )
            )

        try:
            self._last_preview_rows = list(rows)
        except Exception:
            pass

    def _build_mapped_preview_rows(self, max_rows=None):
        if not self.mapping or not self.selected_elements:
            return []

        rows = []
        is_auto_room = bool(getattr(self, "chkAutoRoomByElement", None) and self.chkAutoRoomByElement.IsChecked)
        selected_link = self._get_selected_link()
        scope_mode = self._get_auto_detect_scope()

        for el in self.selected_elements:
            if max_rows is not None and len(rows) >= int(max_rows):
                break

            if is_auto_room:
                room_item = self.element_room_map.get(el.Id.IntegerValue)
                if room_item is not None and selected_link is not None:
                    if room_item.get("link_id") != selected_link.Id.IntegerValue:
                        room_item = None
                if room_item is None:
                    room_item = self._find_linked_room_for_element(
                        el,
                        probe_offset_mm=self._get_probe_offset_mm(),
                        scope_mode=scope_mode,
                        selected_link=selected_link,
                    )
                    if room_item is not None:
                        self.element_room_map[el.Id.IntegerValue] = room_item
                room_values = self._extract_room_values(room_item["room"], room_item["link_doc"]) if room_item is not None else {}
            else:
                room_values = {name: data.get("value") for name, data in self.selected_room_params.items()}

            for room_pname, target_pname in sorted(self.mapping.items()):
                value = room_values.get(room_pname)
                if value is None:
                    value = "<empty>"
                rows.append({
                    "element_id": el.Id.IntegerValue,
                    "room_param": room_pname,
                    "target_param": target_pname,
                    "value": value,
                    "scope": scope_mode,
                    "source_mode": self._get_source_mode(),
                })
                if max_rows is not None and len(rows) >= int(max_rows):
                    break

        return rows

    def _selected_category_names(self):
        names = []
        for item in self.category_items:
            try:
                if item["cb"].IsChecked:
                    names.append(str(item.get("name", "")))
            except Exception:
                continue
        return names

    def _build_preview_export_filename(self):
        scope_name = self._get_auto_detect_scope()
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

        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        return "LinkedRoomMappedPreview_{0}_{1}_{2}.csv".format(scope_tag, categories_tag, timestamp)

    def _prompt_save_preview_csv_path(self, default_filename):
        dialog = SaveFileDialog()
        dialog.Title = "Save Mapped Preview CSV"
        dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        dialog.FileName = default_filename

        try:
            last_dir = str(getattr(config, "last_preview_export_dir", "") or "")
        except Exception:
            last_dir = ""
        if last_dir and os.path.isdir(last_dir):
            dialog.InitialDirectory = last_dir

        result = dialog.ShowDialog()
        if result:
            return dialog.FileName
        return None

    def _export_preview_rows_csv(self, export_path, rows):
        header = [
            "Row",
            "ElementId",
            "RoomParameter",
            "TargetParameter",
            "MappedValue",
            "SourceMode",
            "Scope",
            "Categories",
            "ExportedAt",
        ]

        categories_text = ", ".join(self._selected_category_names())
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(export_path, "wb") as stream:
            writer = csv.writer(stream)
            writer.writerow([_csv_cell(col) for col in header])
            idx = 1
            for row in rows:
                writer.writerow([
                    _csv_cell(idx),
                    _csv_cell(row.get("element_id", "")),
                    _csv_cell(row.get("room_param", "")),
                    _csv_cell(row.get("target_param", "")),
                    _csv_cell(row.get("value", "")),
                    _csv_cell(row.get("source_mode", "")),
                    _csv_cell(row.get("scope", "")),
                    _csv_cell(categories_text),
                    _csv_cell(export_time),
                ])
                idx += 1

    def _maybe_export_mapped_preview_csv(self, rows, required=False):
        try:
            export_enabled = bool(self.chkExportMappedPreviewCsv.IsChecked)
        except Exception:
            export_enabled = False

        if not export_enabled:
            return True

        if not rows:
            forms.alert("No mapped preview rows available to export.", title="Linked Room Parameter Transfer")
            return (not required)

        default_name = self._build_preview_export_filename()
        export_path = self._prompt_save_preview_csv_path(default_name)
        if not export_path:
            if required:
                forms.alert("Transfer cancelled because mapped preview CSV export was enabled and no file location was selected.", title="Linked Room Parameter Transfer")
                return False
            return True

        self._export_preview_rows_csv(export_path, rows)

        try:
            export_dir = os.path.dirname(export_path)
            if export_dir:
                config.last_preview_export_dir = export_dir
                script.save_config()
        except Exception:
            pass

        forms.alert("Mapped preview exported to:\n{0}".format(export_path), title="Linked Room Parameter Transfer")
        return True

    def _get_auto_match_exclusion_keywords(self):
        try:
            raw = str(getattr(getattr(self, "txtAutoMatchExclude", None), "Text", "") or "")
        except Exception:
            raw = ""

        parts = []
        token = []
        for ch in raw:
            if ch in [",", ";", "\n", "\r", "\t", "|"]:
                part = "".join(token).strip().lower()
                if part:
                    parts.append(part)
                token = []
            else:
                token.append(ch)
        last = "".join(token).strip().lower()
        if last:
            parts.append(last)

        # Preserve order but deduplicate.
        seen = set()
        uniq = []
        for p in parts:
            if p in seen:
                continue
            seen.add(p)
            uniq.append(p)
        return uniq

    def _is_excluded_for_auto_match(self, param_name, exclusion_keywords):
        if not param_name or not exclusion_keywords:
            return False
        text = str(param_name).strip().lower()
        if not text:
            return False
        for kw in exclusion_keywords:
            if kw and kw in text:
                return True
        return False

    def preview_click(self, sender, e):
        try:
            if not self.mapping:
                forms.alert("No mappings exist. Press Auto Match or add mappings manually first.")
                return
            if not self.selected_elements:
                forms.alert("No target elements selected. Choose target elements before previewing.")
                return
            self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Preview failed")
            forms.alert("Preview failed safely:\n{0}".format(str(ex)))

    def _auto_match_mappings(self):
        if not self.selected_room_params or not self.common_params:
            return {"matched": 0, "excluded_source": 0, "excluded_target": 0}

        mode = self._get_auto_map_mode()
        if mode == "disabled":
            return {"matched": 0, "excluded_source": 0, "excluded_target": 0}

        exclusion_keywords = self._get_auto_match_exclusion_keywords()
        allowed_target_params = self._get_auto_map_allowed_target_params()
        if not allowed_target_params:
            return {
                "matched": 0,
                "excluded_source": 0,
                "excluded_target": len(self.common_params),
                "excluded_target_by_filter": len(self.common_params),
            }

        filtered_targets = []
        excluded_target = 0
        excluded_target_by_filter = 0
        for target_name in sorted(self.common_params):
            if target_name not in allowed_target_params:
                excluded_target += 1
                excluded_target_by_filter += 1
                continue
            if self._is_excluded_for_auto_match(target_name, exclusion_keywords):
                excluded_target += 1
                continue
            filtered_targets.append(target_name)

        matched = 0
        excluded_source = 0

        for room_param in sorted(self.selected_room_params.keys()):
            if self._is_excluded_for_auto_match(room_param, exclusion_keywords):
                excluded_source += 1
                continue

            match = get_best_match(room_param, filtered_targets, mode)
            if match:
                self.mapping[room_param] = match
                matched += 1

        return {
            "matched": matched,
            "excluded_source": excluded_source,
            "excluded_target": excluded_target,
            "excluded_target_by_filter": excluded_target_by_filter,
        }

    def _get_auto_map_mode(self):
        try:
            combo = getattr(self, "cmbAutoMapStrictness", None)
            if combo is not None and combo.SelectedItem is not None:
                item = combo.SelectedItem
                text = str(item.Content) if hasattr(item, "Content") else str(item)
                t = text.lower()
                if "disabled" in t:
                    return "disabled"
                if "strict" in t:
                    return "strict"
                if "loose" in t:
                    return "loose"
        except Exception:
            pass
        return "balanced"

    def mapping_strictness_changed(self, sender, e):
        try:
            # Guard: event fires during XAML init before __init__ completes.
            if not getattr(self, "common_params", None):
                return
            if not getattr(self, "mapping_auto_generated", False):
                return
            self.mapping = {}
            self._auto_match_mappings()
            self._render_mapping_list()
            self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Auto-map strictness change failed")
            forms.alert("Auto-map strictness update failed safely:\n{0}".format(str(ex)))

    def auto_match_exclusions_changed(self, sender, e):
        try:
            if not getattr(self, "common_params", None):
                return
            if not getattr(self, "mapping_auto_generated", False):
                return
            self.mapping = {}
            self._auto_match_mappings()
            self._render_mapping_list()
            self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Auto-match exclusions change failed")
            forms.alert("Auto-match exclusion update failed safely:\n{0}".format(str(ex)))

    def _room_header_text(self, room, room_doc, link_inst):
        room_name = ""
        room_number = ""
        room_level = ""

        try:
            p = room.LookupParameter("Name")
            if p:
                room_name = p.AsString() or ""
        except Exception:
            pass

        try:
            p = room.LookupParameter("Number")
            if p:
                room_number = p.AsString() or ""
        except Exception:
            pass

        try:
            lvl = room_doc.GetElement(room.LevelId)
            if lvl is not None:
                room_level = lvl.Name
        except Exception:
            pass

        source_kind = self._get_source_kind_name()
        model_name = "Host Model"
        try:
            if link_inst is not None:
                model_name = link_inst.Name
        except Exception:
            pass

        return (
            "{0} Name: {1}\n"
            "{0} Number: {2}\n"
            "Level: {3}\n"
            "Model: {4}"
        ).format(source_kind, room_name, room_number, room_level, model_name)

    def _extract_room_parameters(self):
        self.selected_room_params = {}
        self.lstRoomParameters.Items.Clear()

        if self.selected_room is None:
            return

        show_empty = bool(self.chkShowEmpty.IsChecked)

        for p in self.selected_room.Parameters:
            if p is None or p.Definition is None:
                continue
            name = p.Definition.Name
            st = p.StorageType
            value = read_parameter_value(p, self.selected_room_doc)
            display_value = _display_room_param_value(p, value, self.selected_room_doc)

            if not show_empty and (display_value is None or display_value == ""):
                continue

            entry = {
                "value": value,
                "display": display_value,
                "storage": st,
                "has_value": p.HasValue,
            }

            existing = self.selected_room_params.get(name)
            if existing is None:
                self.selected_room_params[name] = entry
            else:
                # Duplicate parameter names can exist; keep whichever has a non-empty display.
                if _is_empty_value(existing.get("display")) and (not _is_empty_value(entry.get("display"))):
                    self.selected_room_params[name] = entry

        for pname in sorted(self.selected_room_params.keys()):
            data = self.selected_room_params[pname]
            self.lstRoomParameters.Items.Add(
                "{0} | {1} | {2}".format(
                    pname,
                    data.get("display") if data.get("display") is not None else "<empty>",
                    str(data["storage"]),
                )
            )

        self._refresh_mapping_controls()
        self._render_mapping_list()

    def _rebuild_auto_map_target_filter(self, target_params):
        try:
            lst = getattr(self, "lstAutoMapTargetFilter", None)
            if lst is None:
                return

            previous_state = {}
            for item in getattr(self, "auto_map_target_filter_items", []):
                try:
                    previous_state[item["name"]] = bool(item["cb"].IsChecked)
                except Exception:
                    continue

            lst.Items.Clear()
            self.auto_map_target_filter_items = []

            for pname in sorted(target_params):
                cb = CheckBox()
                cb.Content = pname
                cb.IsChecked = previous_state.get(pname, True)
                try:
                    cb.Checked += self.auto_map_target_filter_changed
                    cb.Unchecked += self.auto_map_target_filter_changed
                except Exception:
                    pass

                self.auto_map_target_filter_items.append({"name": pname, "cb": cb})
                lst.Items.Add(cb)
        except Exception:
            pass

    def _get_auto_map_allowed_target_params(self):
        allowed = set()
        for item in getattr(self, "auto_map_target_filter_items", []):
            try:
                if bool(item["cb"].IsChecked):
                    allowed.add(item["name"])
            except Exception:
                continue
        return allowed

    def _set_auto_map_target_filters(self, enabled):
        for item in getattr(self, "auto_map_target_filter_items", []):
            try:
                item["cb"].IsChecked = bool(enabled)
            except Exception:
                continue

    def auto_map_target_filter_changed(self, sender, e):
        try:
            if not getattr(self, "mapping_auto_generated", False):
                return
            self.mapping = {}
            self._auto_match_mappings()
            self._render_mapping_list()
            self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Auto-map target filter change failed")
            forms.alert("Auto-map target filter update failed safely:\n{0}".format(str(ex)))

    def select_all_auto_map_targets_click(self, sender, e):
        try:
            self._set_auto_map_target_filters(True)
            if getattr(self, "mapping_auto_generated", False):
                self.mapping = {}
                self._auto_match_mappings()
                self._render_mapping_list()
                self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Select-all auto-map targets failed")
            forms.alert("Select-all auto-map targets failed safely:\n{0}".format(str(ex)))

    def deselect_all_auto_map_targets_click(self, sender, e):
        try:
            self._set_auto_map_target_filters(False)
            if getattr(self, "mapping_auto_generated", False):
                self.mapping = {}
                self._auto_match_mappings()
                self._render_mapping_list()
                self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Deselect-all auto-map targets failed")
            forms.alert("Deselect-all auto-map targets failed safely:\n{0}".format(str(ex)))

    def _refresh_target_parameters(self):
        self.lstCommonParams.Items.Clear()
        try:
            self.lstAutoMapTargetFilter.Items.Clear()
        except Exception:
            pass
        self.lstMappings.Items.Clear()
        self.lstPreview.Items.Clear()
        self.common_params = set()
        self.auto_map_target_filter_items = []
        self.mapping = {}
        self.mapping_auto_generated = False

        if not self.selected_elements:
            return

        param_sets = [get_writable_parameters(e) for e in self.selected_elements]
        if not param_sets:
            return

        show_non_common = bool(self.chkShowNonCommon.IsChecked)

        if show_non_common:
            all_params = set()
            for pset in param_sets:
                all_params.update(pset)
            target_params = all_params
        else:
            common = set(param_sets[0])
            for pset in param_sets[1:]:
                common.intersection_update(pset)
            target_params = common

        self.common_params = set(target_params)
        self._rebuild_auto_map_target_filter(target_params)

        for pname in sorted(target_params):
            self.lstCommonParams.Items.Add(pname)

        self.mapping = {}
        self.mapping_auto_generated = False
        self._refresh_mapping_controls()
        self._render_mapping_list()
        self.lstPreview.Items.Clear()
        self._apply_responsive_panel_heights()

    def _set_element_summary(self):
        if not self.selected_elements:
            self.txtElementSummary.Text = "No elements selected."
            return

        by_cat = defaultdict(int)
        for e in self.selected_elements:
            try:
                cat_name = e.Category.Name if e.Category else "<No Category>"
            except Exception:
                cat_name = "<No Category>"
            by_cat[cat_name] += 1

        lines = ["Selected elements: {0}".format(len(self.selected_elements))]
        for cname in sorted(by_cat.keys()):
            lines.append("- {0}: {1}".format(cname, by_cat[cname]))

        if self.element_room_map:
            linked = len([1 for e in self.selected_elements if e.Id.IntegerValue in self.element_room_map])
            lines.append("- auto-room matched: {0}".format(linked))

        self.txtElementSummary.Text = "\n".join(lines)

    def _get_element_probe_point(self, element):
        points = self._get_element_probe_points(element)
        if points:
            return points[0]
        return None

    def _append_unique_probe_point(self, points, point, tol=1e-6):
        if point is None:
            return
        for p in points:
            try:
                if (
                    abs(float(p.X) - float(point.X)) <= tol
                    and abs(float(p.Y) - float(point.Y)) <= tol
                    and abs(float(p.Z) - float(point.Z)) <= tol
                ):
                    return
            except Exception:
                continue
        points.append(point)

    def _get_element_probe_points(self, element):
        points = []
        if element is None:
            return points

        # For MEP Spaces: Location is a LocationPoint; get it directly and avoid
        # an expensive bounding-box call on a view that may not show the space.
        try:
            cat = element.Category
            if cat is not None and cat.Id.IntegerValue == int(BuiltInCategory.OST_MEPSpaces):
                loc = element.Location
                if loc is not None and hasattr(loc, "Point") and loc.Point is not None:
                    self._append_unique_probe_point(points, loc.Point)
        except Exception:
            pass

        try:
            loc = element.Location
            if loc is not None:
                if hasattr(loc, "Point") and loc.Point is not None:
                    self._append_unique_probe_point(points, loc.Point)
                if hasattr(loc, "Curve") and loc.Curve is not None:
                    for t in (0.0, 0.25, 0.5, 0.75, 1.0):
                        try:
                            self._append_unique_probe_point(points, loc.Curve.Evaluate(t, True))
                        except Exception:
                            continue
        except Exception:
            pass

        try:
            bb = element.get_BoundingBox(None)
            if bb is None:
                bb = element.get_BoundingBox(doc.ActiveView)
            if bb is not None:
                self._append_unique_probe_point(points, XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5,
                ))
        except Exception:
            pass

        return points

    def _point_on_segment_2d(self, px, py, ax, ay, bx, by, tol):
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

    def _point_in_polygon_2d(self, px, py, polygon, tol):
        if not polygon or len(polygon) < 3:
            return False

        count = len(polygon)
        for i in range(count):
            ax, ay = polygon[i]
            bx, by = polygon[(i + 1) % count]
            if self._point_on_segment_2d(px, py, ax, ay, bx, by, tol):
                return True

        inside = False
        j = count - 1
        for i in range(count):
            xi, yi = polygon[i]
            xj, yj = polygon[j]
            crosses = ((yi > py) != (yj > py))
            if crosses:
                denom = (yj - yi)
                if abs(denom) < 1e-12:
                    j = i
                    continue
                x_int = (xj - xi) * (py - yi) / denom + xi
                if px < x_int:
                    inside = not inside
            j = i

        return inside

    def _get_auto_detect_scope(self):
        """Return user-selected auto-detect scope label."""
        try:
            item = getattr(self, "cmbAutoDetectScope", None)
            if item is not None and item.SelectedItem is not None:
                selected = item.SelectedItem
                if hasattr(selected, "Content"):
                    return str(selected.Content)
                return str(selected)
        except Exception:
            pass
        return "Active View Level"

    def _get_active_level(self):
        try:
            active_view = doc.ActiveView
            if active_view is not None:
                return active_view.GenLevel
        except Exception:
            pass
        return None

    def _element_matches_active_level(self, element):
        """Cross-category active-level check used to trim project-wide fallbacks."""
        active_level = self._get_active_level()
        if active_level is None:
            return True

        try:
            level_id = getattr(element, "LevelId", None)
            if level_id is not None and getattr(level_id, "IntegerValue", -1) > 0:
                return level_id.IntegerValue == active_level.Id.IntegerValue
        except Exception:
            pass

        # For family-based elements and many MEP categories.
        try:
            p = element.LookupParameter("Reference Level")
            if p is not None and p.HasValue:
                lvl_id = p.AsElementId()
                if lvl_id is not None and lvl_id.IntegerValue > 0:
                    return lvl_id.IntegerValue == active_level.Id.IntegerValue
        except Exception:
            pass

        # If we cannot determine a level, keep element instead of false-negative filtering.
        return True

    def _room_matches_active_level(self, room, room_doc, tol=0.01, target_level=None):
        """Match linked room to host active view level by elevation/name, not ElementId."""
        if target_level is None:
            try:
                active_view = doc.ActiveView
                active_level = active_view.GenLevel if active_view is not None else None
            except Exception:
                active_level = None
        else:
            active_level = target_level

        if active_level is None:
            return True

        try:
            room_level = room_doc.GetElement(room.LevelId)
        except Exception:
            room_level = None

        if room_level is None:
            return False

        try:
            if abs(float(room_level.Elevation) - float(active_level.Elevation)) <= tol:
                return True
        except Exception:
            pass

        try:
            return (room_level.Name or "").strip().lower() == (active_level.Name or "").strip().lower()
        except Exception:
            return False

    def _source_matches_active_level(self, source_el, source_doc, tol=0.01, target_level=None):
        if source_el is None or source_doc is None:
            return False

        cat_id = None
        try:
            cat_id = source_el.Category.Id.IntegerValue if source_el.Category is not None else None
        except Exception:
            cat_id = None

        if cat_id == int(BuiltInCategory.OST_Rooms):
            return self._room_matches_active_level(source_el, source_doc, tol=tol, target_level=target_level)

        active_level = target_level if target_level is not None else self._get_active_level()
        if active_level is None:
            return True

        source_level = None
        try:
            level_id = getattr(source_el, "LevelId", None)
            if level_id is not None and getattr(level_id, "IntegerValue", -1) > 0:
                source_level = source_doc.GetElement(level_id)
        except Exception:
            source_level = None

        if source_level is None:
            return True

        try:
            if abs(float(source_level.Elevation) - float(active_level.Elevation)) <= tol:
                return True
        except Exception:
            pass

        try:
            return (source_level.Name or "").strip().lower() == (active_level.Name or "").strip().lower()
        except Exception:
            return False

    def _build_room_detection_index(self, scope_mode=None, selected_link=None):
        self.room_detection_index = []
        self.room_detection_grid = {}
        self.room_detection_grid_fallback = []
        self.room_detection_grid_fallback_by_link = {}
        self.room_detection_link_contexts = {}
        self.room_index_capped = False
        opts = SpatialElementBoundaryOptions() if self._ENABLE_BOUNDARY_FALLBACK else None
        scope_mode = scope_mode or self._get_auto_detect_scope()
        self.room_detection_scope_mode = scope_mode
        self.room_detection_link_id = selected_link.Id.IntegerValue if selected_link is not None else None
        filter_by_active_level = (scope_mode != "Entire Project")
        max_indexed = self._SAFE_MAX_INDEXED_SOURCES if self._is_safe_mode() else self._MAX_INDEXED_SOURCES
        source_category_id = self._get_source_category_id()
        source_bic = self._get_source_bic()
        linked_mode = self._is_linked_source_mode()

        source_contexts = []
        if linked_mode:
            links_to_scan = [selected_link] if selected_link is not None else list(self.links)
            for link_inst in links_to_scan:
                link_doc = link_inst.GetLinkDocument()
                if link_doc is None:
                    continue
                try:
                    inv_transform = link_inst.GetTotalTransform().Inverse
                except Exception:
                    continue
                source_contexts.append((link_inst, link_doc, inv_transform, link_inst.Id.IntegerValue))
        else:
            source_contexts.append((None, doc, None, 0))

        for link_inst, src_doc, inv_transform, link_id in source_contexts:
            try:
                sources = (
                    FilteredElementCollector(src_doc)
                        .OfCategory(source_bic)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                continue

            for room in sources:
                if room is None:
                    continue

                if len(self.room_detection_index) >= max_indexed:
                    self.room_index_capped = True
                    self._breadcrumb("Source index cap reached ({0})".format(max_indexed))
                    return

                if filter_by_active_level and (not self._source_matches_active_level(room, src_doc)):
                    continue

                loops = []
                minx = None
                miny = None
                maxx = None
                maxy = None
                minz = None
                maxz = None

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

                if self._ENABLE_BOUNDARY_FALLBACK:
                    try:
                        seg_loops = room.GetBoundarySegments(opts)
                    except Exception as ex:
                        # Log the error and continue - don't let one room's boundary issue crash everything
                        logger.debug("Failed to get boundary segments for room {0}: {1}".format(room.Id.IntegerValue, str(ex)))
                        seg_loops = None

                    if seg_loops:
                        try:
                            for seg_loop in seg_loops:
                                if seg_loop is None:
                                    continue
                                poly = []
                                for seg in seg_loop:
                                    try:
                                        curve = seg.GetCurve()
                                        if curve is None:
                                            continue
                                        p0 = curve.GetEndPoint(0)
                                        if p0 is not None:
                                            poly.append((p0.X, p0.Y))
                                    except Exception:
                                        continue

                                if len(poly) >= 3:
                                    loops.append(poly)
                                    for x, y in poly:
                                        minx = x if minx is None else min(minx, x)
                                        miny = y if miny is None else min(miny, y)
                                        maxx = x if maxx is None else max(maxx, x)
                                        maxy = y if maxy is None else max(maxy, y)
                        except Exception as ex:
                            logger.debug("Error processing boundary loops for room {0}: {1}".format(room.Id.IntegerValue, str(ex)))

                has_boundary = bool(loops) and minx is not None

                room_item = {
                    "link_inst": link_inst,
                    "link_doc": src_doc,
                    "inv_transform": inv_transform,
                    "room": room,
                    "room_id": room.Id.IntegerValue,
                    "link_id": link_id,
                    "minx": minx,
                    "miny": miny,
                    "minz": minz,
                    "maxx": maxx,
                    "maxy": maxy,
                    "maxz": maxz,
                    "loops": loops,
                    "has_boundary": has_boundary,
                }
                self.room_detection_index.append(room_item)
                self.room_detection_link_contexts[link_id] = inv_transform

                # Index room bounding boxes so detection does not scan every room
                # for every element in large projects.
                if minx is None or miny is None or maxx is None or maxy is None:
                    self.room_detection_grid_fallback.append(room_item)
                    self.room_detection_grid_fallback_by_link.setdefault(link_id, []).append(room_item)
                    continue
                try:
                    cell_size = self.room_detection_grid_size
                    cell_min_x = int(math.floor(float(minx) / cell_size))
                    cell_max_x = int(math.floor(float(maxx) / cell_size))
                    cell_min_y = int(math.floor(float(miny) / cell_size))
                    cell_max_y = int(math.floor(float(maxy) / cell_size))
                    cell_count = (cell_max_x - cell_min_x + 1) * (cell_max_y - cell_min_y + 1)
                    if cell_count > 400:
                        self.room_detection_grid_fallback.append(room_item)
                        self.room_detection_grid_fallback_by_link.setdefault(link_id, []).append(room_item)
                        continue
                    for cell_x in range(cell_min_x, cell_max_x + 1):
                        for cell_y in range(cell_min_y, cell_max_y + 1):
                            self.room_detection_grid.setdefault((link_id, cell_x, cell_y), []).append(room_item)
                except Exception:
                    self.room_detection_grid_fallback.append(room_item)
                    self.room_detection_grid_fallback_by_link.setdefault(link_id, []).append(room_item)

    def _find_linked_room_for_host_point(self, host_point, probe_offset_mm=0.0, tol=1.0, scope_mode=None, selected_link=None):
        if host_point is None:
            return None

        resolved_scope = scope_mode or self._get_auto_detect_scope()
        resolved_link_id = selected_link.Id.IntegerValue if selected_link is not None else None
        if (
            (not self.room_detection_index)
            or (self.room_detection_scope_mode != resolved_scope)
            or (self.room_detection_link_id != resolved_link_id)
        ):
            self._build_room_detection_index(scope_mode=resolved_scope, selected_link=selected_link)

        points = [host_point]
        try:
            offset_ft = float(probe_offset_mm or 0.0) / 304.8
            if abs(offset_ft) > 1e-6:
                points.append(XYZ(host_point.X, host_point.Y, host_point.Z + offset_ft))
                points.append(XYZ(host_point.X, host_point.Y, host_point.Z - offset_ft))
        except Exception:
            pass

        candidate_rooms = []
        seen_room_keys = set()
        for probe_point in points:
            for link_id, inv_transform in self.room_detection_link_contexts.items():
                try:
                    p = inv_transform.OfPoint(probe_point) if inv_transform is not None else probe_point
                except Exception:
                    continue
                try:
                    cell_size = self.room_detection_grid_size
                    cell_x = int(math.floor(float(p.X) / cell_size))
                    cell_y = int(math.floor(float(p.Y) / cell_size))
                    nearby = []
                    for dx in (-1, 0, 1):
                        for dy in (-1, 0, 1):
                            nearby.extend(self.room_detection_grid.get((link_id, cell_x + dx, cell_y + dy), []))
                    nearby.extend(self.room_detection_grid_fallback_by_link.get(link_id, []))
                    for room_item in nearby:
                        room_key = (room_item["link_id"], room_item["room_id"])
                        if room_key not in seen_room_keys:
                            seen_room_keys.add(room_key)
                            candidate_rooms.append(room_item)
                except Exception:
                    continue

        # Rooms without usable bounding boxes are rare; preserve correctness by
        # checking them even when the spatial grid has no matching cell.
        for room_item in self.room_detection_grid_fallback:
            room_key = (room_item["link_id"], room_item["room_id"])
            if room_key not in seen_room_keys:
                seen_room_keys.add(room_key)
                candidate_rooms.append(room_item)

        for room_item in candidate_rooms:
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
                    if minx is not None and maxx is not None:
                        if p.X < (minx - tol) or p.X > (maxx + tol):
                            continue
                    if miny is not None and maxy is not None:
                        if p.Y < (miny - tol) or p.Y > (maxy + tol):
                            continue
                    if minz is not None and maxz is not None:
                        if p.Z < (minz - tol) or p.Z > (maxz + tol):
                            continue
                except Exception:
                    pass

                room = room_item["room"]

                cat_id = None
                try:
                    cat_id = room.Category.Id.IntegerValue if room.Category is not None else None
                except Exception:
                    cat_id = None

                # Prefer direct containment checks when API provides it.
                try:
                    if cat_id == int(BuiltInCategory.OST_Rooms) and hasattr(room, "IsPointInRoom") and room.IsPointInRoom(p):
                        return room_item
                    if cat_id == int(BuiltInCategory.OST_MEPSpaces) and hasattr(room, "IsPointInSpace") and room.IsPointInSpace(p):
                        return room_item
                except Exception:
                    pass

                # Fall back to boundary polygon tests when boundary data is available.
                if room_item.get("has_boundary"):
                    try:
                        px = p.X
                        py = p.Y
                        if px < (room_item["minx"] - tol) or px > (room_item["maxx"] + tol):
                            continue
                        if py < (room_item["miny"] - tol) or py > (room_item["maxy"] + tol):
                            continue

                        hit = False
                        for poly in room_item["loops"]:
                            if self._point_in_polygon_2d(px, py, poly, tol):
                                hit = True
                                break

                        if hit:
                            return room_item
                    except Exception:
                        continue

        return None

    def _find_linked_room_for_element(self, element, probe_offset_mm=0.0, scope_mode=None, selected_link=None):
        points = self._get_element_probe_points(element)
        if not points:
            return None

        vote_count = defaultdict(int)
        room_hits = {}

        for probe_point in points:
            room_item = self._find_linked_room_for_host_point(
                probe_point,
                probe_offset_mm=probe_offset_mm,
                scope_mode=scope_mode,
                selected_link=selected_link,
            )
            if room_item is None:
                continue
            key = (room_item["link_id"], room_item["room_id"])
            vote_count[key] += 1
            if key not in room_hits:
                room_hits[key] = room_item

        if not vote_count:
            return None

        best_key = None
        best_votes = -1
        for key, votes in vote_count.items():
            if votes > best_votes:
                best_key = key
                best_votes = votes
            elif votes == best_votes and best_key is not None and key < best_key:
                best_key = key

        if len(vote_count) > 1:
            try:
                logger.debug(
                    "Element %s matched multiple linked rooms via probe points; picked %s with %s/%s votes",
                    element.Id.IntegerValue,
                    "{0}:{1}".format(best_key[0], best_key[1]),
                    best_votes,
                    len(points),
                )
            except Exception:
                pass

        return room_hits.get(best_key)

    # Categories that do not support view-based FilteredElementCollector filtering.
    # These must always be collected project-wide even in "Active View Level" mode.
    _VIEW_UNSAFE_CATEGORIES = {
        int(BuiltInCategory.OST_HVAC_Zones),
        int(BuiltInCategory.OST_MEPSpaces),
        int(BuiltInCategory.OST_Rooms),
    }

    def _collect_category_elements(self, bic, scope_mode):
        """Safely collect elements for a single category, falling back to project-wide when needed."""
        # Some categories (HVAC Zones, Spaces) do not support view-based collection.
        use_project_wide = (
            scope_mode == "Entire Project"
            or int(bic) in self._VIEW_UNSAFE_CATEGORIES
        )
        try:
            if use_project_wide:
                return (
                    FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                )
            else:
                return (
                    FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                )
        except Exception:
            # Last-resort fallback: try project-wide if view-based failed
            try:
                return (
                    FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                return []

    def _iter_selected_category_elements(self, scope_mode=None):
        allowed_ids = set(self._selected_category_ids())
        if not allowed_ids:
            return []

        scope_mode = scope_mode or self._get_auto_detect_scope()

        items = []
        seen = set()

        # Safety limit: prevent unbounded collection when using "Entire Project"
        max_elements = 5000 if scope_mode == "Entire Project" else 50000
        if self._is_safe_mode():
            max_elements = min(max_elements, self._SAFE_MAX_TARGET_ELEMENTS)

        for _, bic in TARGET_CATEGORIES:
            if int(bic) not in allowed_ids:
                continue

            collected = self._collect_category_elements(bic, scope_mode)
            for el in collected:
                if el is None:
                    continue

                # Spaces / Zones are collected project-wide in view mode; trim to active level.
                if scope_mode != "Entire Project" and int(bic) in self._VIEW_UNSAFE_CATEGORIES:
                    if not self._element_matches_active_level(el):
                        continue

                try:
                    eid = el.Id.IntegerValue
                except Exception:
                    continue
                if eid in seen:
                    continue
                seen.add(eid)
                # Skip unplaced/redundant spaces (area == 0 or no valid location)
                if int(bic) == int(BuiltInCategory.OST_MEPSpaces):
                    try:
                        area = el.Area
                        if area <= 0.0:
                            continue
                    except Exception:
                        pass
                    try:
                        loc = el.Location
                        if loc is None:
                            continue
                    except Exception:
                        continue
                items.append(el)

                # Stop if we've collected too many elements
                if len(items) >= max_elements:
                    return items

        return items

    def _extract_room_values(self, room, room_doc):
        values = {}
        if room is None:
            return values

        for p in room.Parameters:
            if p is None or p.Definition is None:
                continue
            name = p.Definition.Name
            raw = read_parameter_value(p, room_doc)
            display = _display_room_param_value(p, raw, room_doc)
            chosen = raw if not _is_empty_value(raw) else display

            if name not in values:
                values[name] = chosen
            elif _is_empty_value(values.get(name)) and (not _is_empty_value(chosen)):
                values[name] = chosen

        return values

    def auto_detect_elements_click(self, sender, e):
        try:
            t0 = time.time()
            allowed_ids = self._selected_category_ids()
            if not allowed_ids:
                forms.alert("Select at least one target category.")
                return

            scope_mode = self._get_auto_detect_scope()
            selected_link = self._get_selected_link()
            self._breadcrumb("Auto-detect start | source_mode={0} | scope={1} | safe_mode={2}".format(self._get_source_mode(), scope_mode, self._is_safe_mode()))

            # Show warning for "Entire Project" scope
            if scope_mode == "Entire Project":
                warning_msg = (
                    "WARNING: 'Entire Project' scope will search through ALL MEP elements in the project.\n\n"
                    "This can be SLOW or may hang/crash Revit on large projects.\n\n"
                    "Recommendations:\n"
                    "1. Try 'Active View Level' scope first if elements are on one level\n"
                    "2. Pre-select elements in your view, then use 'Use Current Selection'\n"
                    "3. If project is very large, consider filtering by category first\n\n"
                    "Continue with 'Entire Project' search?"
                )
                if not self._confirm(warning_msg, title="Large Scope Warning"):
                    return

            candidates = self._iter_selected_category_elements(scope_mode)
            if not candidates:
                forms.alert(
                    "No elements found in {0} for selected categories.".format(
                        "entire project" if scope_mode == "Entire Project" else "active view level"
                    )
                )
                return

            self._breadcrumb("Auto-detect candidates collected: {0}".format(len(candidates)))

            self._build_room_detection_index(scope_mode=scope_mode, selected_link=selected_link)
            if not self.room_detection_index:
                forms.alert(
                    "No linked rooms were indexed for {0}.\n"
                    "Tip: try 'Entire Project' scope in Step 2B, or use Step 1 Detect Rooms to confirm links contain rooms."
                    .format("entire project" if scope_mode == "Entire Project" else "active view level")
                )
                return

            self._breadcrumb("Source index size: {0}".format(len(self.room_detection_index)))
            if self.room_index_capped:
                forms.alert(
                    "Source index reached safe cap ({0}).\n"
                    "Refine the scope (Active View Level / one linked model) and retry."
                    .format(self._SAFE_MAX_INDEXED_SOURCES if self._is_safe_mode() else self._MAX_INDEXED_SOURCES)
                )
                return

            self.element_room_map = {}
            matched_elements = []
            room_hit_count = defaultdict(int)

            for el in candidates:
                room_item = self._find_linked_room_for_element(
                    el,
                    probe_offset_mm=self._get_probe_offset_mm(),
                    scope_mode=scope_mode,
                    selected_link=selected_link,
                )
                if room_item is None:
                    continue

                self.element_room_map[el.Id.IntegerValue] = room_item
                matched_elements.append(el)
                room_hit_count[(room_item["link_id"], room_item["room_id"])] += 1

                if self._is_safe_mode() and len(matched_elements) >= self._SAFE_MAX_TARGET_ELEMENTS:
                    self._breadcrumb("Safe cap reached for matched elements ({0})".format(self._SAFE_MAX_TARGET_ELEMENTS))
                    break

            self.selected_elements = matched_elements
            self._set_element_summary()

            if not matched_elements:
                self._reset_selected_room()
                self._refresh_target_parameters()
                forms.alert(
                    "Auto-detect did not find category elements inside linked room footprints. "
                    "Tip: verify categories, selected scope, and linked room boundaries."
                )
                return

            top_key = None
            top_count = 0
            for key, count in room_hit_count.items():
                if count > top_count:
                    top_count = count
                    top_key = key

            if top_key is not None:
                for item in self.room_detection_index:
                    if (item["link_id"], item["room_id"]) == top_key:
                        self.selected_room = item["room"]
                        self.selected_room_doc = item["link_doc"]
                        self.selected_room_link_inst = item["link_inst"]
                        self.txtRoomInfo.Text = self._room_header_text(
                            item["room"], item["link_doc"], item["link_inst"]
                        )
                        self.pnlRoomInfo.Visibility = System.Windows.Visibility.Visible
                        break

            self._extract_room_parameters()
            self._refresh_target_parameters()
            self._breadcrumb("Auto-detect completed in {0:.2f}s | matched={1}".format(time.time() - t0, len(matched_elements)))
        except Exception as ex:
            logger.exception("Auto-detect elements failed")
            forms.alert("Auto-detect failed safely:\n{0}".format(str(ex)))

    def _get_selected_link(self):
        """Get selected linked model; return None when using all links."""
        if not self._is_linked_source_mode():
            return None
        try:
            combo = getattr(self, "cmbSelectLink", None)
            if combo is not None and combo.SelectedIndex >= 0:
                self.selected_link_index = combo.SelectedIndex
                if combo.SelectedIndex == 0:
                    return None
                idx = combo.SelectedIndex - 1
                if idx >= 0 and idx < len(self.links):
                    return self.links[idx]
        except Exception:
            pass

        return None

    def _get_room_detection_mode(self):
        """Get the currently selected room detection mode."""
        try:
            combo = getattr(self, "cmbRoomDetectionMode", None)
            if combo is not None and combo.SelectedItem is not None:
                item = combo.SelectedItem
                if hasattr(item, "Content"):
                    return str(item.Content)
                return str(item)
        except Exception:
            pass
        return "Auto-Detect Rooms In Current View"

    def _get_selected_detect_level(self):
        if not self._is_per_level_detection_mode():
            return None

        try:
            combo = getattr(self, "cmbDetectLevel", None)
            if combo is None:
                return None
            idx = combo.SelectedIndex
            if idx is None or idx <= 0:
                return self._get_active_level()

            level_idx = idx - 1
            if 0 <= level_idx < len(self.detect_levels):
                return self.detect_levels[level_idx]
        except Exception:
            pass

        return self._get_active_level()

    def _active_level_id(self):
        try:
            selected_level = self._get_selected_detect_level()
            if selected_level is not None:
                return selected_level.Id.IntegerValue
        except Exception:
            pass
        try:
            if doc.ActiveView is not None and doc.ActiveView.GenLevel is not None:
                return doc.ActiveView.GenLevel.Id.IntegerValue
        except Exception:
            pass
        return None

    def _collect_rooms_lightweight(self, links_to_scan, level_id=None):
        rooms_found = []
        source_category_id = self._get_source_category_id()
        source_bic = self._get_source_bic()
        linked_mode = self._is_linked_source_mode()
        target_level = self._get_selected_detect_level() if level_id is not None else None

        if linked_mode:
            for link_inst in links_to_scan:
                link_doc = link_inst.GetLinkDocument()
                if link_doc is None:
                    continue

                try:
                    sources = (
                        FilteredElementCollector(link_doc)
                        .OfCategory(source_bic)
                        .WhereElementIsNotElementType()
                    )
                except Exception:
                    continue

                for room in sources:
                    if room is None:
                        continue
                    if level_id is not None:
                        if not self._source_matches_active_level(room, link_doc, target_level=target_level):
                            continue
                    rooms_found.append((link_inst, link_doc, room))
        else:
            try:
                sources = (
                    FilteredElementCollector(doc)
                    .OfCategory(source_bic)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                sources = []

            for room in sources:
                if room is None:
                    continue
                if level_id is not None and (not self._source_matches_active_level(room, doc, target_level=target_level)):
                    continue
                rooms_found.append((None, doc, room))

        return rooms_found

    def _pick_representative_room(self, room_items):
        if not room_items:
            return None

        def sort_key(item):
            _, _, room = item
            number = ""
            name = ""
            try:
                p = room.LookupParameter("Number")
                if p:
                    number = p.AsString() or ""
            except Exception:
                pass
            try:
                p = room.LookupParameter("Name")
                if p:
                    name = p.AsString() or ""
            except Exception:
                pass
            return (number.lower(), name.lower(), room.Id.IntegerValue)

        return sorted(room_items, key=sort_key)[0]

    def _auto_detect_rooms_in_link(self):
        """Auto-detect sources by mode with optional selected-link filter."""
        detection_mode = self._get_room_detection_mode()
        selected_link = self._get_selected_link()
        source_kind = self._get_source_kind_name()

        links_to_scan = [selected_link] if selected_link is not None else list(self.links)
        if self._is_linked_source_mode() and (not links_to_scan):
            forms.alert("No loaded linked models found.")
            return

        level_id = None
        if "Entire Project" not in detection_mode:
            level_id = self._active_level_id()

        rooms = self._collect_rooms_lightweight(links_to_scan, level_id=level_id)

        # Fallback: if no rooms found at level, retry without level filter.
        fallback_used = False
        if not rooms and level_id is not None:
            rooms = self._collect_rooms_lightweight(links_to_scan, level_id=None)
            fallback_used = bool(rooms)

        if not rooms:
            forms.alert(
                "No {0}s detected for the selected mode.\n"
                "Try 'Auto-Detect Entire Project' or keep link as 'All Linked Models'.".format(source_kind)
            )
            return

        chosen = self._pick_representative_room(rooms)
        if chosen is None:
            forms.alert("No valid room candidate found.")
            return

        # Keep all detected rooms and promote representative room for parameter list display.
        ordered = [chosen]
        for item in rooms:
            if item is not chosen:
                ordered.append(item)
        self._set_selected_rooms(ordered)

        link_inst, _, _ = chosen

        if link_inst is not None:
            msg = (
                "Auto-detected {0} {1}(s) from {2} link(s).\n"
                "Selected source from: {3}"
            ).format(len(rooms), source_kind.lower(), len(links_to_scan), link_inst.Name)
        else:
            msg = "Auto-detected {0} {1}(s) from host model.".format(len(rooms), source_kind.lower())
        if fallback_used:
            msg += "\nNote: no rooms were found on current level, so full-link scan was used."
        forms.alert(msg)

    def detect_rooms_click(self, sender, e):
        try:
            self._auto_detect_rooms_in_link()
        except Exception as ex:
            logger.exception("Detect sources failed")
            forms.alert("Detect source failed safely:\n{0}".format(str(ex)))

    def pick_linked_rooms_click(self, sender, e):
        try:
            source_category_id = self._get_source_category_id()
            source_kind = self._get_source_kind_name()
            linked_mode = self._is_linked_source_mode()
            selected_link = self._get_selected_link()

            self.Hide()
            try:
                if linked_mode:
                    picked_refs = uidoc.Selection.PickObjects(
                        ObjectType.LinkedElement,
                        "Pick one or more linked {0}s in current view (ESC to finish)".format(source_kind.lower()),
                    )
                else:
                    picked_refs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        SourceCategorySelectionFilter(source_category_id),
                        "Pick one or more host {0}s in current view (ESC to finish)".format(source_kind.lower()),
                    )
            except Exception:
                self.Show()
                return

            self.Show()

            if not picked_refs:
                return

            room_items = []
            skipped_not_room = 0
            skipped_other_link = 0
            seen = set()

            for picked in picked_refs:
                if picked is None:
                    continue

                if linked_mode:
                    link_inst = doc.GetElement(picked.ElementId)
                    if link_inst is None:
                        continue

                    if selected_link is not None and link_inst.Id.IntegerValue != selected_link.Id.IntegerValue:
                        skipped_other_link += 1
                        continue

                    link_doc = link_inst.GetLinkDocument()
                    if link_doc is None:
                        continue

                    room = link_doc.GetElement(picked.LinkedElementId)
                    if room is None:
                        continue

                    try:
                        cat_id = room.Category.Id.IntegerValue if room.Category else None
                    except Exception:
                        cat_id = None

                    if cat_id != source_category_id:
                        skipped_not_room += 1
                        continue

                    key = (link_inst.Id.IntegerValue, room.Id.IntegerValue)
                    if key in seen:
                        continue
                    seen.add(key)
                    room_items.append((link_inst, link_doc, room))
                else:
                    room = doc.GetElement(picked.ElementId)
                    if room is None:
                        continue
                    try:
                        cat_id = room.Category.Id.IntegerValue if room.Category else None
                    except Exception:
                        cat_id = None
                    if cat_id != source_category_id:
                        skipped_not_room += 1
                        continue
                    key = (0, room.Id.IntegerValue)
                    if key in seen:
                        continue
                    seen.add(key)
                    room_items.append((None, doc, room))

            if not room_items:
                forms.alert("No valid {0}s were picked. Please pick {0} elements only.".format(source_kind))
                return

            representative = self._pick_representative_room(room_items)
            ordered = [representative] + [item for item in room_items if item is not representative]
            self._set_selected_rooms(ordered)

            if linked_mode:
                msg = "Picked {0} linked {1}(s).".format(len(room_items), source_kind.lower())
            else:
                msg = "Picked {0} host {1}(s).".format(len(room_items), source_kind.lower())
            if skipped_not_room:
                msg += "\nIgnored non-{0} picks: {1}".format(source_kind.lower(), skipped_not_room)
            if skipped_other_link:
                msg += "\nIgnored picks outside selected link: {0}".format(skipped_other_link)
            forms.alert(msg)
        except Exception as ex:
            try:
                self.Show()
            except Exception:
                pass
            logger.exception("Pick sources failed")
            forms.alert("Pick source failed safely:\n{0}".format(str(ex)))

    def category_search_changed(self, sender, e):
        try:
            self._refresh_category_list(self.txtCategorySearch.Text)
        except Exception as ex:
            logger.exception("Category search refresh failed")
            forms.alert("Category filter failed safely:\n{0}".format(str(ex)))

    def select_all_categories_click(self, sender, e):
        try:
            for item in self.category_items:
                item["cb"].IsChecked = True
            self._save_persistent_settings_now()
        except Exception as ex:
            logger.exception("Select-all categories failed")
            forms.alert("Select-all failed safely:\n{0}".format(str(ex)))

    def deselect_all_categories_click(self, sender, e):
        try:
            for item in self.category_items:
                item["cb"].IsChecked = False
            self._save_persistent_settings_now()
        except Exception as ex:
            logger.exception("Deselect-all categories failed")
            forms.alert("Deselect-all failed safely:\n{0}".format(str(ex)))

    def use_current_selection_click(self, sender, e):
        try:
            selected_ids = uidoc.Selection.GetElementIds()
            if not selected_ids:
                forms.alert("Current selection is empty.")
                return

            allowed_ids = set(self._selected_category_ids())
            if not allowed_ids:
                forms.alert("Select at least one target category.")
                return

            elements = []
            for eid in selected_ids:
                el = doc.GetElement(eid)
                if el is None or el.Category is None:
                    continue
                if el.Category.Id.IntegerValue in allowed_ids:
                    elements.append(el)

            self.selected_elements = elements
            self.element_room_map = {}
            self._set_element_summary()
            self._refresh_target_parameters()
        except Exception as ex:
            logger.exception("Use current selection failed")
            forms.alert("Use current selection failed safely:\n{0}".format(str(ex)))

    def select_elements_click(self, sender, e):
        try:
            allowed_ids = self._selected_category_ids()
            if not allowed_ids:
                forms.alert("Select at least one target category.")
                return

            self.Hide()
            try:
                refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    CategorySelectionFilter(allowed_ids),
                    "Select target elements from chosen categories",
                )
            except Exception:
                self.Show()
                return

            self.Show()

            self.selected_elements = [doc.GetElement(r.ElementId) for r in refs if r is not None]
            self.element_room_map = {}
            self._set_element_summary()
            self._refresh_target_parameters()
        except Exception as ex:
            try:
                self.Show()
            except Exception:
                pass
            logger.exception("Select elements failed")
            forms.alert("Element selection failed safely:\n{0}".format(str(ex)))

    def refresh_click(self, sender, e):
        try:
            self._extract_room_parameters()
            self._refresh_target_parameters()
        except Exception as ex:
            logger.exception("Refresh failed")
            forms.alert("Refresh failed safely:\n{0}".format(str(ex)))

    def add_update_mapping_click(self, sender, e):
        try:
            room_param = self.cmbMappingRoom.SelectedItem
            target_param = self.cmbMappingTarget.SelectedItem

            if not room_param:
                forms.alert("Pick a room parameter.")
                return
            if not target_param:
                forms.alert("Pick a target parameter.")
                return

            self.mapping[str(room_param).strip()] = str(target_param).strip()
            self.mapping_auto_generated = False
            self._render_mapping_list()
            self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Add/update mapping failed")
            forms.alert("Add/Update mapping failed safely:\n{0}".format(str(ex)))

    def remove_mapping_click(self, sender, e):
        try:
            removed_count = 0

            # Primary path: remove all selected rows from Active Mappings.
            selected_rows = []
            try:
                for item in self.lstMappings.SelectedItems:
                    selected_rows.append(item)
            except Exception:
                selected_rows = []

            for selected_row in selected_rows:
                try:
                    row_text = str(selected_row).strip()
                    prefix = "[mapped] "
                    if row_text.startswith(prefix):
                        row_text = row_text[len(prefix):]
                    if " -> " not in row_text:
                        continue
                    room_key = row_text.split(" -> ", 1)[0].strip()
                    if room_key in self.mapping:
                        del self.mapping[room_key]
                        removed_count += 1
                except Exception:
                    continue

            # Fallback path: remove by room parameter combo selection.
            if removed_count == 0:
                room_param = self.cmbMappingRoom.SelectedItem
                if room_param:
                    key = str(room_param).strip()
                    if key in self.mapping:
                        del self.mapping[key]
                        removed_count = 1

            if removed_count == 0:
                forms.alert("Select one or more rows in 'Active Mappings' (or pick a room parameter) to remove.")
                return

            self.mapping_auto_generated = False
            self._render_mapping_list()
            self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Remove mapping failed")
            forms.alert("Remove mapping failed safely:\n{0}".format(str(ex)))

    def auto_match_click(self, sender, e):
        try:
            if self._get_auto_map_mode() == "disabled":
                forms.alert("Auto-map is disabled. Change 'Auto-map strictness' to enable matching.")
                return
            self.mapping = {}
            stats = self._auto_match_mappings()
            self.mapping_auto_generated = True
            self._render_mapping_list()
            self._refresh_preview_list()
            self._save_persistent_settings_now()

            if stats is not None:
                forms.alert(
                    "Auto-match complete.\n"
                    "Mapped pairs: {0}\n"
                    "Excluded source params: {1}\n"
                    "Excluded target params: {2}\n"
                    "Excluded target params (checkbox filter): {3}".format(
                        int(stats.get("matched", 0)),
                        int(stats.get("excluded_source", 0)),
                        int(stats.get("excluded_target", 0)),
                        int(stats.get("excluded_target_by_filter", 0)),
                    )
                )
        except Exception as ex:
            logger.exception("Auto-match failed")
            forms.alert("Auto-match failed safely:\n{0}".format(str(ex)))

    def clear_mappings_click(self, sender, e):
        try:
            self.mapping = {}
            self.mapping_auto_generated = False
            self._render_mapping_list()
            self._refresh_preview_list()
        except Exception as ex:
            logger.exception("Clear mappings failed")
            forms.alert("Clear mappings failed safely:\n{0}".format(str(ex)))

    def _write_transfer_log(self, lines):
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        base = os.path.join(tempfile.gettempdir(), "LinkedRoomTransfer-{0}".format(ts))
        path = base + ".log"
        with io.open(path, "w", encoding="utf-8") as fp:
            fp.write("\n".join(lines))

        csv_path = None
        try:
            detail_rows = getattr(self, "_last_transfer_detail_rows", None)
            if detail_rows:
                csv_path = base + ".csv"
                with io.open(csv_path, "w", encoding="utf-8") as cf:
                    cf.write("timestamp,element_id,room_param,target_param,target_source,storage,before_value,attempted_value,after_value,result,message\n")
                    for r in detail_rows:
                        cf.write(r + "\n")
        except Exception:
            csv_path = None

        try:
            if hasattr(self, "_last_transfer_detail_rows"):
                del self._last_transfer_detail_rows
        except Exception:
            try:
                del self._last_transfer_detail_rows
            except Exception:
                pass

        return path, csv_path

    def transfer_click(self, sender, e):
        auto_room_mode = bool(getattr(self, "chkAutoRoomByElement", None) and self.chkAutoRoomByElement.IsChecked)
        selected_link = self._get_selected_link()
        scope_mode = self._get_auto_detect_scope()
        t0 = time.time()
        estimated_ms = self._estimate_transfer_duration_ms(
            len(getattr(self, "selected_elements", []) or []),
            len(getattr(self, "mapping", {}) or {}),
            auto_room_mode,
        )

        if not auto_room_mode and self.selected_room is None:
            forms.alert("Select a linked room first.")
            return
        if not self.selected_elements:
            forms.alert("Select target elements first.")
            return
        if not self.mapping:
            forms.alert("No smart mappings found. Adjust categories/selection and refresh.")
            return

        if self._is_safe_mode() and len(self.selected_elements) > self._SAFE_MAX_TRANSFER_ELEMENTS:
            forms.alert(
                "Safe Mode is ON and selected elements exceed stability cap ({0}).\n"
                "Please narrow selection or disable Safe Mode for this run."
                .format(self._SAFE_MAX_TRANSFER_ELEMENTS)
            )
            return

        self._breadcrumb(
            "Transfer start | mode={0} | source_mode={1} | elements={2} | mappings={3} | safe_mode={4}".format(
                "auto-room" if auto_room_mode else "single-source",
                self._get_source_mode(),
                len(self.selected_elements),
                len(self.mapping),
                self._is_safe_mode(),
            )
        )

        # Safety check: warn if processing many elements
        if len(self.selected_elements) > 1000:
            warning = (
                "Processing {0} elements.\n\n"
                "This may take a while or cause performance issues.\n"
                "Continue?"
            ).format(len(self.selected_elements))
            if not self._confirm(warning, title="Large Transfer Warning"):
                return

        # Safety: prevent multiple room params targeting the same destination param.
        by_target = defaultdict(list)
        for room_pname, target_pname in self.mapping.items():
            by_target[str(target_pname).strip()].append(str(room_pname).strip())

        duplicated_targets = [tp for tp, srcs in by_target.items() if len(srcs) > 1]
        if duplicated_targets:
            lines = [
                "Invalid mapping: each target parameter can only be mapped once.",
                "Fix these duplicated targets:",
            ]
            for tp in sorted(duplicated_targets)[:10]:
                lines.append("- {0} <= {1}".format(tp, ", ".join(sorted(by_target[tp]))))
            forms.alert("\n".join(lines))
            return

        if estimated_ms >= 12000:
            msg = (
                "Estimated transfer duration: {0}.\n"
                "Large transfer detected; this can take longer depending on model complexity.\n\n"
                "Continue?"
            ).format(_format_duration_ms(estimated_ms))
            if not self._confirm(msg, title="Transfer Time Estimate"):
                return
        elif estimated_ms >= 3000:
            forms.alert(
                "Estimated transfer duration: {0}.".format(_format_duration_ms(estimated_ms)),
                title="Transfer Time Estimate",
            )

        # Optional backup export: if enabled, force explicit save-location confirmation.
        backup_rows = self._build_mapped_preview_rows(max_rows=None)
        if not self._maybe_export_mapped_preview_csv(backup_rows, required=True):
            return

        duplicate_mode = "Overwrite"
        selected_mode = self.cmbDuplicateMode.SelectedItem
        try:
            duplicate_mode = selected_mode.Content
        except Exception:
            pass

        skip_empty = bool(self.chkSkipEmpty.IsChecked)

        log_lines = []
        log_lines.append("Linked Room Parameter Transfer started")
        log_lines.append("Selected room mode: {0}".format("auto-room" if auto_room_mode else "single-room"))
        log_lines.append("Selected elements: {0}".format(len(self.selected_elements)))
        log_lines.append("Mappings:")
        for room_pname, target_pname in self.mapping.items():
            log_lines.append("  {0} -> {1}".format(room_pname, target_pname))
        allow_selected_room_fallback = bool(getattr(self, "chkUseSelectedRoomFallback", None) and self.chkUseSelectedRoomFallback.IsChecked)
        log_lines.append("Selected-room fallback enabled: {0}".format(allow_selected_room_fallback))

        # Detailed CSV-style rows for per-attempt analysis
        detail_rows = []

        updated_elements = set()
        transferred = 0
        failed = 0
        skipped = 0
        skipped_type_dedup = 0
        blocked_type_auto = 0
        no_change_skips = 0
        instance_writes_applied = 0
        type_writes_applied = 0
        mapping_evaluated = 0
        fail_messages = []
        used_rooms = set()
        room_value_cache = {}
        type_write_keys = set()
        fallback_used_count = 0
        successful_write_attempts = []

        tx = Transaction(doc, "Linked Room Parameter Transfer")
        tx.Start()
        try:
            for el in self.selected_elements:
                try:
                    per_element_values = None
                    room_item = None

                    if auto_room_mode:
                        room_item = self.element_room_map.get(el.Id.IntegerValue)
                        if room_item is not None and selected_link is not None:
                            if room_item.get("link_id") != selected_link.Id.IntegerValue:
                                room_item = None
                        if room_item is None:
                            room_item = self._find_linked_room_for_element(
                                el,
                                probe_offset_mm=self._get_probe_offset_mm(),
                                scope_mode=scope_mode,
                                selected_link=selected_link,
                            )
                            if room_item is not None:
                                self.element_room_map[el.Id.IntegerValue] = room_item

                        if room_item is None and allow_selected_room_fallback and self.selected_room is not None:
                            # Fallback to manually selected linked room when auto-room detection fails.
                            room_item = {
                                "link_id": self.selected_room_link_inst.Id.IntegerValue,
                                "room_id": self.selected_room.Id.IntegerValue,
                                "room": self.selected_room,
                                "link_doc": self.selected_room_doc,
                            }
                            log_lines.append(
                                "Element {0}: auto-room failed, falling back to selected room.".format(
                                    el.Id.IntegerValue
                                )
                            )
                            fallback_used_count += 1

                        if room_item is None:
                            skipped += len(self.mapping)
                            log_lines.append(
                                "Element {0}: no linked room found, skipping {1} mappings.".format(
                                    el.Id.IntegerValue, len(self.mapping)
                                )
                            )
                            continue

                        room_key = (room_item["link_id"], room_item["room_id"])
                        used_rooms.add(room_key)
                        if room_key not in room_value_cache:
                            room_value_cache[room_key] = self._extract_room_values(
                                room_item["room"], room_item["link_doc"]
                            )
                        per_element_values = room_value_cache[room_key]

                    for room_pname, target_pname in self.mapping.items():
                        mapping_evaluated += 1
                        if auto_room_mode:
                            value = per_element_values.get(room_pname) if per_element_values else None
                        else:
                            room_data = self.selected_room_params.get(room_pname)
                            if not room_data:
                                skipped += 1
                                log_lines.append(
                                    "Element {0}: source room parameter '{1}' not found, skipping.".format(
                                        el.Id.IntegerValue, room_pname
                                    )
                                )
                                continue
                            value = room_data.get("value")
                            if value in (None, ""):
                                value = room_data.get("display")

                        if skip_empty and (value is None or value == ""):
                            skipped += 1
                            log_lines.append(
                                "Element {0}: value for {1} -> {2} is empty, skipping.".format(
                                    el.Id.IntegerValue, room_pname, target_pname
                                )
                            )
                            continue

                        target_param, target_source = _find_writable_parameter(el, target_pname)
                        log_lines.append(
                            "Element {0}: trying {1} -> {2} (target_source={3})".format(
                                el.Id.IntegerValue, room_pname, target_pname, target_source or "none"
                            )
                        )

                        if target_param is None:
                            failed += 1
                            if len(fail_messages) < 8:
                                if target_source == "read-only":
                                    msg_desc = "target parameter exists but is read-only"
                                else:
                                    msg_desc = "target parameter not found or not writable"
                                fail_messages.append(
                                    "{0} -> {1}: {2} on element {3}".format(
                                        room_pname, target_pname, msg_desc, el.Id.IntegerValue
                                    )
                                )
                            available = list_writable_parameter_names(el, max_items=25)
                            if available:
                                log_lines.append(
                                    "  Available writable params for element {0}: {1}".format(
                                        el.Id.IntegerValue, ", ".join(available)
                                    )
                                )
                            continue

                        current_value = read_parameter_value(target_param)
                        log_lines.append(
                            "  Found parameter '{0}' on element {1}, current='{2}', new='{3}', storage={4}".format(
                                target_pname,
                                el.Id.IntegerValue,
                                current_value,
                                value,
                                target_param.StorageType,
                            )
                        )

                        if target_source == "type":
                            if auto_room_mode:
                                allow_type_writes = bool(getattr(self, "chkAllowTypeWritesInAutoRoomMode", None) and self.chkAllowTypeWritesInAutoRoomMode.IsChecked)
                                if not allow_type_writes:
                                    blocked_type_auto += 1
                                    skipped += 1
                                    if len(fail_messages) < 8:
                                        fail_messages.append(
                                            "{0} -> {1}: skipped because target is a TYPE parameter in auto-room mode".format(
                                                room_pname, target_pname
                                            )
                                        )
                                    log_lines.append(
                                        "  Element {0}: skipping type parameter {1} in auto-room mode.".format(
                                            el.Id.IntegerValue, target_pname
                                        )
                                    )
                                    continue
                                log_lines.append(
                                    "  Element {0}: allowing type parameter {1} in auto-room mode because user enabled the option.".format(
                                        el.Id.IntegerValue, target_pname
                                    )
                                )
                            try:
                                owner_id = target_param.Element.Id.IntegerValue
                            except Exception:
                                try:
                                    owner_id = el.GetTypeId().IntegerValue
                                except Exception:
                                    owner_id = None

                            if owner_id is not None:
                                write_key = (owner_id, target_pname)
                                if write_key in type_write_keys:
                                    skipped_type_dedup += 1
                                    skipped += 1
                                    log_lines.append(
                                        "  Skipping duplicate type write for {0} on type {1}".format(
                                            target_pname, owner_id
                                        )
                                    )
                                    continue
                                type_write_keys.add(write_key)

                        ok, msg = set_parameter_value(target_param, value, duplicate_mode)
                        post_value_for_log = None
                        try:
                            post_value_for_log = read_parameter_value(target_param)
                        except Exception:
                            post_value_for_log = None

                        # Add a CSV-safe detailed row for this attempt
                        msg_l = (msg or "").lower()
                        is_skip_result = (
                            (not ok)
                            and (
                                "skipped" in msg_l
                                or "empty" in msg_l
                                or "no change" in msg_l
                                or "already equal" in msg_l
                            )
                        )
                        try:
                            tsnow = datetime.now().isoformat()
                            storage = str(target_param.StorageType)
                            source_kind = str(target_source or "")
                            before_value = str(current_value) if current_value is not None else ""
                            attempted = str(value) if value is not None else ""
                            after_value = str(post_value_for_log) if post_value_for_log is not None else ""
                            # escape double quotes
                            before_safe = '"' + before_value.replace('"', '""') + '"'
                            attempted_safe = '"' + attempted.replace('"', '""') + '"'
                            after_safe = '"' + after_value.replace('"', '""') + '"'
                            if ok:
                                result = "SUCCESS"
                            elif is_skip_result:
                                result = "SKIP"
                            else:
                                result = "FAIL"
                            msg_safe = '"' + str(msg).replace('"', '""') + '"'
                            row = ",".join([
                                tsnow,
                                str(el.Id.IntegerValue),
                                str(room_pname),
                                str(target_pname),
                                source_kind,
                                storage,
                                before_safe,
                                attempted_safe,
                                after_safe,
                                result,
                                msg_safe,
                            ])
                            detail_rows.append(row)
                        except Exception:
                            pass
                        if ok:
                            transferred += 1
                            updated_elements.add(el.Id.IntegerValue)
                            if target_source == "type":
                                type_writes_applied += 1
                            else:
                                instance_writes_applied += 1
                            successful_write_attempts.append({
                                "element_id": el.Id.IntegerValue,
                                "room_param": room_pname,
                                "target_param": target_pname,
                                "target_source": target_source,
                                "storage": target_param.StorageType,
                                "expected": value,
                                "result_msg": msg,
                            })
                            log_lines.append(
                                "  SUCCESS: {0} -> {1} on element {2} ({3})".format(
                                    room_pname, target_pname, el.Id.IntegerValue, msg
                                )
                            )
                        else:
                            if is_skip_result:
                                skipped += 1
                                if "no change" in msg_l or "already equal" in msg_l:
                                    no_change_skips += 1
                                log_lines.append(
                                    "  SKIPPED: {0} -> {1} on element {2}: {3}".format(
                                        room_pname, target_pname, el.Id.IntegerValue, msg
                                    )
                                )
                            else:
                                failed += 1
                                if len(fail_messages) < 8:
                                    fail_messages.append(
                                        "{0} -> {1} [{2}]: {3}".format(
                                            room_pname, target_pname, target_source, msg
                                        )
                                    )
                                log_lines.append(
                                    "  FAIL: {0} -> {1} on element {2}: {3}".format(
                                        room_pname, target_pname, el.Id.IntegerValue, msg
                                    )
                                )
                except Exception as ex:
                    failed += 1
                    logger.exception("Element transfer failed for element {0}".format(el.Id.IntegerValue))
                    log_lines.append(
                        "Element {0}: unexpected error, skipped remaining mappings ({1})".format(
                            el.Id.IntegerValue, ex
                        )
                    )
                    continue

            tx.Commit()
            try:
                doc.Regenerate()
            except Exception:
                pass
            try:
                uidoc.RefreshActiveView()
            except Exception:
                pass

            verified_transferred = 0
            verified_elements = set()
            unverified_count = 0
            unverified_samples = []

            for rec in successful_write_attempts:
                try:
                    el_id = rec.get("element_id")
                    el_for_verify = doc.GetElement(ElementId(int(el_id))) if el_id is not None else None
                    if el_for_verify is None:
                        unverified_count += 1
                        if len(unverified_samples) < 8:
                            unverified_samples.append(
                                "Element {0}: could not reload element for verification".format(el_id)
                            )
                        continue

                    p = _find_parameter_for_verification(
                        el_for_verify,
                        rec.get("target_param"),
                        rec.get("target_source"),
                    )
                    if p is None:
                        unverified_count += 1
                        if len(unverified_samples) < 8:
                            unverified_samples.append(
                                "{0} on element {1}: target parameter not found during verification".format(
                                    rec.get("target_param"),
                                    el_id,
                                )
                            )
                        continue

                    actual = read_parameter_value(p)
                    expected = rec.get("expected")
                    storage = rec.get("storage")
                    result_msg = str(rec.get("result_msg") or "").lower()

                    ok_verify = False
                    if "append" in result_msg:
                        try:
                            ok_verify = str(expected).strip() in str(actual)
                        except Exception:
                            ok_verify = False
                    else:
                        ok_verify = values_equal(actual, expected, storage)

                    if ok_verify:
                        verified_transferred += 1
                        verified_elements.add(el_id)
                    else:
                        unverified_count += 1
                        if len(unverified_samples) < 8:
                            unverified_samples.append(
                                "{0} -> {1} on element {2} [{3}] expected='{4}' actual='{5}'".format(
                                    rec.get("room_param"),
                                    rec.get("target_param"),
                                    rec.get("element_id"),
                                    rec.get("target_source"),
                                    expected,
                                    actual,
                                )
                            )
                except Exception as vex:
                    unverified_count += 1
                    if len(unverified_samples) < 8:
                        unverified_samples.append(
                            "Verification error on element {0}: {1}".format(rec.get("element_id"), vex)
                        )
        except Exception as ex:
            tx.RollBack()
            forms.alert("Transfer failed and transaction was rolled back:\n{0}".format(ex))
            return

        elapsed_s = time.time() - t0
        elapsed_ms = int(elapsed_s * 1000.0)

        summary = [
            "Transfer Summary",
            "Mode: {0}".format("Auto detect room per element" if auto_room_mode else "Single selected room"),
            "Estimated duration: {0}".format(_format_duration_ms(estimated_ms)),
            "Room context: {0}".format(
                "{0} detected room(s)".format(len(used_rooms))
                if auto_room_mode
                else self.txtRoomInfo.Text.replace("\n", " | ")
            ),
            "Elements updated (verified): {0}".format(len(verified_elements)),
            "Mappings evaluated: {0}".format(mapping_evaluated),
            "Parameter writes (verified): {0}".format(verified_transferred),
            "Write calls returned success (pre-verify): {0}".format(transferred),
            "Instance writes applied: {0}".format(instance_writes_applied),
            "Type writes applied: {0}".format(type_writes_applied),
            "Failed writes: {0}".format(failed),
            "Skipped: {0}".format(skipped),
            "Elapsed: {0:.2f}s ({1} ms)".format(elapsed_s, elapsed_ms),
        ]

        if verified_transferred == 0:
            summary.append("No persisted writes were detected.")

        if unverified_count:
            summary.append("Successes not persisted after commit: {0}".format(unverified_count))
            summary.append("\nSample non-persisted writes:")
            summary.extend(unverified_samples)

        if no_change_skips:
            summary.append("No-change skips (value already equal): {0}".format(no_change_skips))

        if transferred == 0 and failed == 0 and no_change_skips > 0:
            summary.append("Outcome: all mapped parameters were already up-to-date; no write was required.")

        if fallback_used_count:
            summary.append("Fallback used for {0} element(s)".format(fallback_used_count))

        if blocked_type_auto:
            summary.append("Type-parameter writes blocked in auto-room mode: {0}".format(blocked_type_auto))
        if skipped_type_dedup:
            summary.append("Duplicate type writes skipped: {0}".format(skipped_type_dedup))

        if fail_messages:
            summary.append("\nSample failures:")
            summary.extend(fail_messages)

        try:
            self._last_transfer_detail_rows = detail_rows
        except Exception:
            pass
        full_log_lines = list(summary)
        full_log_lines.append("")
        full_log_lines.append("Detailed transfer trace:")
        full_log_lines.extend(log_lines)
        log_path, csv_path = self._write_transfer_log(full_log_lines)
        summary.append("\nLog file: {0}".format(log_path))
        if csv_path:
            summary.append("Detailed CSV file: {0}".format(csv_path))

        if self.runtime_breadcrumbs:
            summary.append("\nRuntime Breadcrumbs (latest):")
            summary.extend(self.runtime_breadcrumbs[-12:])

        self.last_transfer_summary = list(summary)
        self._save_persistent_settings_now()

        # Inform user if fallback to the manually selected linked room was used
        try:
            if fallback_used_count:
                forms.alert("Selected-room fallback used for {0} element(s).\nResults summary will be shown next.".format(fallback_used_count), title="Linked Room Parameter Transfer")
        except Exception:
            pass

        forms.alert("\n".join(summary), title="Linked Room Parameter Transfer")

    def cancel_click(self, sender, e):
        try:
            self.Close()
        except Exception as ex:
            logger.exception("Cancel/Close failed")
            forms.alert("Close failed safely:\n{0}".format(str(ex)))


try:
    ui = LinkedRoomTransferWindow("WPFWindow.xaml")
    if getattr(ui, "links", None):
        ui.ShowDialog()
except Exception as ex:
    try:
        logger.exception("Linked Room Parameter Transfer crashed during startup")
    except Exception:
        pass
    forms.alert("Tool startup failed safely:\n{0}".format(str(ex)), title="Linked Room Parameter Transfer")
