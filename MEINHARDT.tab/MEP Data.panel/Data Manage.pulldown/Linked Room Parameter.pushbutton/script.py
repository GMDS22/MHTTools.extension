# coding: utf8
from __future__ import print_function

from collections import defaultdict
import io
import os
import tempfile
from datetime import datetime

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
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


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
config = script.get_config()

__title__ = "Linked Room\nParameter Transfer"
__doc__ = "Select a room from a linked model and transfer room parameter values into selected host elements."


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
    }
    return settings


def _save_persistent_settings(settings):
    """Save all persistent settings to config."""
    try:
        config.selected_categories = settings.get("selected_categories", [])
        config.probe_offset_mm = settings.get("probe_offset_mm", 2000)
        script.save_config()
    except Exception:
        pass


def normalize_name(name):
    if not name:
        return ""
    return "".join(ch for ch in name.lower() if ch.isalnum())


def get_link_instances(active_doc):
    return list(FilteredElementCollector(active_doc).OfClass(RevitLinkInstance))


def read_parameter_value(param, source_doc=None):
    if param is None or not param.HasValue:
        return None
    src_doc = source_doc or doc
    st = param.StorageType
    try:
        if st == StorageType.String:
            return param.AsString()
        if st == StorageType.Integer:
            return param.AsInteger()
        if st == StorageType.Double:
            return param.AsDouble()
        if st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid and eid.IntegerValue > 0:
                e = src_doc.GetElement(eid)
                if e is not None:
                    try:
                        return e.Name
                    except Exception:
                        return eid.IntegerValue
            return None
    except Exception:
        return None
    return None


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


class LinkedRoomTransferWindow(WPFWindow):
    # Boundary extraction is expensive and has caused instability in some models.
    # Keep direct Revit room containment as the primary method.
    _ENABLE_BOUNDARY_FALLBACK = False

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
        self.selected_link_index = 0

        self._load_links()
        if not self.links:
            return
        self._populate_link_selector()
        self._build_category_list()
        self.persistent_settings = _load_persistent_settings()
        self._apply_persistent_settings()
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

    def _save_persistent_settings_now(self):
        settings = {
            "selected_categories": [
                item["name"] for item in self.category_items if item["cb"].IsChecked
            ],
            "probe_offset_mm": self._get_probe_offset_mm(),
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
        self.txtRoomInfo.Text = "No room selected. Use Step 1 Detect Rooms or Pick Room(s) In Current View."
        self.pnlRoomInfo.Visibility = System.Windows.Visibility.Collapsed
        self.lstRoomParameters.Items.Clear()
        self.lstPreview.Items.Clear()

    def _set_selected_rooms(self, room_items):
        """Persist selected rooms and use first room as parameter source for mapping UI."""
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
            header = "{0}\nRooms selected: {1}".format(header, room_count)

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
        if not self.mapping or not self.selected_elements:
            return

        preview_rows = []
        max_rows = 250
        is_auto_room = bool(getattr(self, "chkAutoRoomByElement", None) and self.chkAutoRoomByElement.IsChecked)

        for el in self.selected_elements:
            if len(preview_rows) >= max_rows:
                break

            if is_auto_room:
                room_item = self.element_room_map.get(el.Id.IntegerValue)
                if room_item is None:
                    pt = self._get_element_probe_point(el)
                    room_item = self._find_linked_room_for_host_point(pt, probe_offset_mm=self._get_probe_offset_mm())
                    if room_item is not None:
                        self.element_room_map[el.Id.IntegerValue] = room_item
                room_values = self._extract_room_values(room_item["room"], room_item["link_doc"]) if room_item is not None else {}
            else:
                room_values = {name: data.get("value") for name, data in self.selected_room_params.items()}

            for room_pname, target_pname in sorted(self.mapping.items()):
                value = room_values.get(room_pname)
                if value is None:
                    value = "<empty>"
                preview_rows.append("Element {0}: {1} -> {2} = {3}".format(el.Id.IntegerValue, room_pname, target_pname, value))
                if len(preview_rows) >= max_rows:
                    break

        if len(preview_rows) >= max_rows:
            preview_rows.append("...showing first {0} preview rows".format(max_rows))

        for row in preview_rows:
            self.lstPreview.Items.Add(row)

    def preview_click(self, sender, e):
        if not self.mapping:
            forms.alert("No mappings exist. Press Auto Match or add mappings manually first.")
            return
        if not self.selected_elements:
            forms.alert("No target elements selected. Choose target elements before previewing.")
            return
        self._refresh_preview_list()

    def _auto_match_mappings(self):
        if not self.selected_room_params or not self.common_params:
            return

        mode = self._get_auto_map_mode()
        if mode == "disabled":
            return

        for room_param in sorted(self.selected_room_params.keys()):
            match = get_best_match(room_param, self.common_params, mode)
            if match:
                self.mapping[room_param] = match

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
        # Guard: event fires during XAML init before __init__ completes.
        if not getattr(self, "common_params", None):
            return
        if not getattr(self, "mapping_auto_generated", False):
            return
        self.mapping = {}
        self._auto_match_mappings()
        self._render_mapping_list()
        self._refresh_preview_list()

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

        return (
            "Room Name: {0}\n"
            "Room Number: {1}\n"
            "Level: {2}\n"
            "Link Name: {3}"
        ).format(room_name, room_number, room_level, link_inst.Name)

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

            if not show_empty and (value is None or value == ""):
                continue

            self.selected_room_params[name] = {
                "value": value,
                "storage": st,
                "has_value": p.HasValue,
            }

        for pname in sorted(self.selected_room_params.keys()):
            data = self.selected_room_params[pname]
            self.lstRoomParameters.Items.Add(
                "{0} | {1} | {2}".format(
                    pname,
                    data["value"] if data["value"] is not None else "<empty>",
                    str(data["storage"]),
                )
            )

        self._refresh_mapping_controls()
        self._render_mapping_list()

    def _refresh_target_parameters(self):
        self.lstCommonParams.Items.Clear()
        self.lstMappings.Items.Clear()
        self.lstPreview.Items.Clear()
        self.common_params = set()
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

        for pname in sorted(target_params):
            self.lstCommonParams.Items.Add(pname)

        self.mapping = {}
        self.mapping_auto_generated = False
        self._refresh_mapping_controls()
        self._render_mapping_list()
        self.lstPreview.Items.Clear()

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
        if element is None:
            return None

        # For MEP Spaces: Location is a LocationPoint; get it directly and avoid
        # an expensive bounding-box call on a view that may not show the space.
        try:
            cat = element.Category
            if cat is not None and cat.Id.IntegerValue == int(BuiltInCategory.OST_MEPSpaces):
                loc = element.Location
                if loc is not None and hasattr(loc, "Point") and loc.Point is not None:
                    return loc.Point
        except Exception:
            pass

        try:
            loc = element.Location
            if loc is not None:
                if hasattr(loc, "Point") and loc.Point is not None:
                    return loc.Point
                if hasattr(loc, "Curve") and loc.Curve is not None:
                    return loc.Curve.Evaluate(0.5, True)
        except Exception:
            pass

        try:
            bb = element.get_BoundingBox(None)
            if bb is None:
                bb = element.get_BoundingBox(doc.ActiveView)
            if bb is not None:
                return XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5,
                )
        except Exception:
            pass

        return None

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

    def _room_matches_active_level(self, room, room_doc, tol=0.01):
        """Match linked room to host active view level by elevation/name, not ElementId."""
        try:
            active_view = doc.ActiveView
            active_level = active_view.GenLevel if active_view is not None else None
        except Exception:
            active_level = None

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

    def _build_room_detection_index(self, scope_mode=None):
        self.room_detection_index = []
        opts = SpatialElementBoundaryOptions() if self._ENABLE_BOUNDARY_FALLBACK else None
        scope_mode = scope_mode or self._get_auto_detect_scope()
        filter_by_active_level = (scope_mode != "Entire Project")

        for link_inst in self.links:
            link_doc = link_inst.GetLinkDocument()
            if link_doc is None:
                continue

            try:
                inv_transform = link_inst.GetTotalTransform().Inverse
            except Exception:
                continue

            rooms = (
                FilteredElementCollector(link_doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
            )

            for room in rooms:
                if room is None:
                    continue

                # Filter linked rooms by active level using elevation/name (cross-doc safe).
                if filter_by_active_level and (not self._room_matches_active_level(room, link_doc)):
                    continue

                loops = []
                minx = None
                miny = None
                maxx = None
                maxy = None

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

                self.room_detection_index.append({
                    "link_inst": link_inst,
                    "link_doc": link_doc,
                    "inv_transform": inv_transform,
                    "room": room,
                    "room_id": room.Id.IntegerValue,
                    "link_id": link_inst.Id.IntegerValue,
                    "minx": minx,
                    "miny": miny,
                    "maxx": maxx,
                    "maxy": maxy,
                    "loops": loops,
                    "has_boundary": has_boundary,
                })

    def _find_linked_room_for_host_point(self, host_point, probe_offset_mm=0.0, tol=1.0):
        if host_point is None:
            return None

        if not self.room_detection_index:
            self._build_room_detection_index()

        points = [host_point]
        try:
            offset_ft = float(probe_offset_mm or 0.0) / 304.8
            if abs(offset_ft) > 1e-6:
                points.append(XYZ(host_point.X, host_point.Y, host_point.Z + offset_ft))
                points.append(XYZ(host_point.X, host_point.Y, host_point.Z - offset_ft))
        except Exception:
            pass

        for room_item in self.room_detection_index:
            for probe_point in points:
                try:
                    p = room_item["inv_transform"].OfPoint(probe_point)
                except Exception:
                    continue

                room = room_item["room"]

                # Prefer direct room containment when the API provides it.
                try:
                    if hasattr(room, "IsPointInRoom") and room.IsPointInRoom(p):
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
            values[p.Definition.Name] = read_parameter_value(p, room_doc)

        return values

    def auto_detect_elements_click(self, sender, e):
        try:
            allowed_ids = self._selected_category_ids()
            if not allowed_ids:
                forms.alert("Select at least one target category.")
                return

            scope_mode = self._get_auto_detect_scope()

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

            self._build_room_detection_index(scope_mode)
            if not self.room_detection_index:
                forms.alert(
                    "No linked rooms were indexed for {0}.\n"
                    "Tip: try 'Entire Project' scope in Step 2B, or use Step 1 Detect Rooms to confirm links contain rooms."
                    .format("entire project" if scope_mode == "Entire Project" else "active view level")
                )
                return

            self.element_room_map = {}
            matched_elements = []
            room_hit_count = defaultdict(int)

            for el in candidates:
                pt = self._get_element_probe_point(el)
                room_item = self._find_linked_room_for_host_point(
                    pt, probe_offset_mm=self._get_probe_offset_mm()
                )
                if room_item is None:
                    continue

                self.element_room_map[el.Id.IntegerValue] = room_item
                matched_elements.append(el)
                room_hit_count[(room_item["link_id"], room_item["room_id"])] += 1

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
        except Exception as ex:
            logger.exception("Auto-detect elements failed")
            forms.alert("Auto-detect failed safely:\n{0}".format(str(ex)))

    def _get_selected_link(self):
        """Get selected linked model; return None when using all links."""
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

    def _active_level_id(self):
        try:
            if doc.ActiveView is not None and doc.ActiveView.GenLevel is not None:
                return doc.ActiveView.GenLevel.Id.IntegerValue
        except Exception:
            pass
        return None

    def _collect_rooms_lightweight(self, links_to_scan, level_id=None):
        rooms_found = []
        for link_inst in links_to_scan:
            link_doc = link_inst.GetLinkDocument()
            if link_doc is None:
                continue

            try:
                rooms = (
                    FilteredElementCollector(link_doc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                continue

            for room in rooms:
                if room is None:
                    continue
                if level_id is not None:
                    if not self._room_matches_active_level(room, link_doc):
                        continue
                rooms_found.append((link_inst, link_doc, room))

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
        """Auto-detect rooms by mode with optional selected-link filter."""
        detection_mode = self._get_room_detection_mode()
        selected_link = self._get_selected_link()

        links_to_scan = [selected_link] if selected_link is not None else list(self.links)
        if not links_to_scan:
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
                "No rooms detected for the selected mode.\n"
                "Try 'Auto-Detect Entire Project' or keep link as 'All Linked Models'."
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

        msg = (
            "Auto-detected {0} room(s) from {1} link(s).\n"
            "Selected room from: {2}"
        ).format(len(rooms), len(links_to_scan), link_inst.Name)
        if fallback_used:
            msg += "\nNote: no rooms were found on current level, so full-link scan was used."
        forms.alert(msg)

    def detect_rooms_click(self, sender, e):
        self._auto_detect_rooms_in_link()

    def pick_linked_rooms_click(self, sender, e):
        selected_link = self._get_selected_link()

        self.Hide()
        try:
            picked_refs = uidoc.Selection.PickObjects(
                ObjectType.LinkedElement,
                "Pick one or more linked rooms in current view (ESC to finish)",
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

            if cat_id != int(BuiltInCategory.OST_Rooms):
                skipped_not_room += 1
                continue

            key = (link_inst.Id.IntegerValue, room.Id.IntegerValue)
            if key in seen:
                continue
            seen.add(key)
            room_items.append((link_inst, link_doc, room))

        if not room_items:
            forms.alert("No linked rooms were picked. Please pick Room elements only.")
            return

        representative = self._pick_representative_room(room_items)
        ordered = [representative] + [item for item in room_items if item is not representative]
        self._set_selected_rooms(ordered)

        msg = "Picked {0} linked room(s).".format(len(room_items))
        if skipped_not_room:
            msg += "\nIgnored non-room picks: {0}".format(skipped_not_room)
        if skipped_other_link:
            msg += "\nIgnored picks outside selected link: {0}".format(skipped_other_link)
        forms.alert(msg)

    def category_search_changed(self, sender, e):
        self._refresh_category_list(self.txtCategorySearch.Text)

    def select_all_categories_click(self, sender, e):
        for item in self.category_items:
            item["cb"].IsChecked = True
        self._save_persistent_settings_now()

    def deselect_all_categories_click(self, sender, e):
        for item in self.category_items:
            item["cb"].IsChecked = False
        self._save_persistent_settings_now()

    def use_current_selection_click(self, sender, e):
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

    def select_elements_click(self, sender, e):
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

    def refresh_click(self, sender, e):
        self._extract_room_parameters()
        self._refresh_target_parameters()

    def add_update_mapping_click(self, sender, e):
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

    def remove_mapping_click(self, sender, e):
        room_param = self.cmbMappingRoom.SelectedItem
        if not room_param:
            forms.alert("Pick a room parameter to remove mapping.")
            return

        key = str(room_param).strip()
        if key in self.mapping:
            del self.mapping[key]
        self.mapping_auto_generated = False
        self._render_mapping_list()
        self._refresh_preview_list()

    def auto_match_click(self, sender, e):
        if self._get_auto_map_mode() == "disabled":
            forms.alert("Auto-map is disabled. Change 'Auto-map strictness' to enable matching.")
            return
        self.mapping = {}
        self._auto_match_mappings()
        self.mapping_auto_generated = True
        self._render_mapping_list()
        self._refresh_preview_list()

    def clear_mappings_click(self, sender, e):
        self.mapping = {}
        self.mapping_auto_generated = False
        self._render_mapping_list()
        self._refresh_preview_list()

    def _write_transfer_log(self, lines):
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            base = os.path.join(tempfile.gettempdir(), "LinkedRoomTransfer-{0}".format(ts))
            path = base + ".log"
            with io.open(path, "w", encoding="utf-8") as fp:
                fp.write("\n".join(lines))

            # Optional detailed CSV (if a detailed rows list was provided as last element)
            try:
                # If caller appended a detail_rows list as an attribute on self, write it too.
                detail_rows = getattr(self, "_last_transfer_detail_rows", None)
                if detail_rows:
                    csv_path = base + ".csv"
                    with io.open(csv_path, "w", encoding="utf-8") as cf:
                        cf.write("timestamp,element_id,room_param,target_param,storage,attempted_value,result,message\n")
                        for r in detail_rows:
                            cf.write(r + "\n")
                    path = path + "; CSV: {0}".format(csv_path)
            except Exception:
                pass

            # Clear any cached detail rows
            try:
                if hasattr(self, "_last_transfer_detail_rows"):
                    delattr(self, "_last_transfer_detail_rows")
            except Exception:
                try:
                    del self._last_transfer_detail_rows
                except Exception:
                    pass

            return path

    def transfer_click(self, sender, e):
        auto_room_mode = bool(getattr(self, "chkAutoRoomByElement", None) and self.chkAutoRoomByElement.IsChecked)

        if auto_room_mode and not self.selected_elements:
            self.auto_detect_elements_click(sender, e)

        if not auto_room_mode and self.selected_room is None:
            forms.alert("Select a linked room first.")
            return
        if not self.selected_elements:
            forms.alert("Select target elements first.")
            return
        if not self.mapping:
            forms.alert("No smart mappings found. Adjust categories/selection and refresh.")
            return

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

        # Allow user to opt-out of falling back to the manually selected linked room
        allow_selected_room_fallback = bool(getattr(self, "chkUseSelectedRoomFallback", None) and self.chkUseSelectedRoomFallback.IsChecked)

        # Detailed CSV-style rows for per-attempt analysis
        detail_rows = []

        updated_elements = set()
        transferred = 0
        failed = 0
        skipped = 0
        skipped_type_dedup = 0
        blocked_type_auto = 0
        fail_messages = []
        used_rooms = set()
        room_value_cache = {}
        type_write_keys = set()
        fallback_used_count = 0

        tx = Transaction(doc, "Linked Room Parameter Transfer")
        tx.Start()
        try:
            for el in self.selected_elements:
                try:
                    per_element_values = None
                    room_item = None

                    if auto_room_mode:
                        room_item = self.element_room_map.get(el.Id.IntegerValue)
                        if room_item is None:
                            pt = self._get_element_probe_point(el)
                            room_item = self._find_linked_room_for_host_point(
                                pt, probe_offset_mm=self._get_probe_offset_mm()
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

                        # Add a CSV-safe detailed row for this attempt
                        try:
                            tsnow = datetime.now().isoformat()
                            storage = str(target_param.StorageType)
                            attempted = str(value) if value is not None else ""
                            # escape double quotes
                            attempted_safe = '"' + attempted.replace('"', '""') + '"'
                            result = "SUCCESS" if ok else "FAIL"
                            msg_safe = '"' + str(msg).replace('"', '""') + '"'
                            row = ",".join([tsnow, str(el.Id.IntegerValue), str(room_pname), str(target_pname), storage, attempted_safe, result, msg_safe])
                            detail_rows.append(row)
                        except Exception:
                            pass
                        if ok:
                            transferred += 1
                            updated_elements.add(el.Id.IntegerValue)
                            log_lines.append(
                                "  SUCCESS: {0} -> {1} on element {2} ({3})".format(
                                    room_pname, target_pname, el.Id.IntegerValue, msg
                                )
                            )
                        else:
                            if "skipped" in msg.lower() or "empty" in msg.lower():
                                skipped += 1
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
        except Exception as ex:
            tx.RollBack()
            forms.alert("Transfer failed and transaction was rolled back:\n{0}".format(ex))
            return

        summary = [
            "Transfer Summary",
            "Mode: {0}".format("Auto detect room per element" if auto_room_mode else "Single selected room"),
            "Room context: {0}".format(
                "{0} detected room(s)".format(len(used_rooms))
                if auto_room_mode
                else self.txtRoomInfo.Text.replace("\n", " | ")
            ),
            "Elements updated: {0}".format(len(updated_elements)),
            "Parameter writes: {0}".format(transferred),
            "Failed writes: {0}".format(failed),
            "Skipped: {0}".format(skipped),
        ]

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
        log_path = self._write_transfer_log(summary)
        summary.append("\nLog file: {0}".format(log_path))
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
        self.Close()


ui = LinkedRoomTransferWindow("WPFWindow.xaml")
if getattr(ui, "links", None):
    ui.ShowDialog()
