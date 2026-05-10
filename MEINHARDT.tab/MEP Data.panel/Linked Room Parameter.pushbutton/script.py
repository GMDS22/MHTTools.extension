# coding: utf8
from __future__ import print_function

from collections import defaultdict
import io
import os
import tempfile
from datetime import datetime

from Autodesk.Revit.DB import (
    BuiltInCategory,
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

__title__ = "Linked Room\nParameter Transfer"
__doc__ = "Select a room from a linked model and transfer room parameter values into selected host elements."


TARGET_CATEGORIES = [
    ("Ducts", BuiltInCategory.OST_DuctCurves),
    ("Duct Fittings", BuiltInCategory.OST_DuctFitting),
    ("Duct Accessories", BuiltInCategory.OST_DuctAccessory),
    ("Air Terminals", BuiltInCategory.OST_DuctTerminal),
    ("Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment),
    ("Flex Ducts", BuiltInCategory.OST_FlexDuctCurves),
    ("Pipes", BuiltInCategory.OST_PipeCurves),
    ("Pipe Fittings", BuiltInCategory.OST_PipeFitting),
    ("Pipe Accessories", BuiltInCategory.OST_PipeAccessory),
    ("Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures),
    ("Sprinklers", BuiltInCategory.OST_Sprinklers),
    ("Flex Pipes", BuiltInCategory.OST_FlexPipeCurves),
    ("Duct Insulations", BuiltInCategory.OST_DuctInsulations),
    ("Pipe Insulations", BuiltInCategory.OST_PipeInsulations),
    ("Lighting Fixtures", BuiltInCategory.OST_LightingFixtures),
    ("Lighting Devices", BuiltInCategory.OST_LightingDevices),
    ("Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment),
    ("Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures),
    ("Data Devices", BuiltInCategory.OST_DataDevices),
    ("Communication Devices", BuiltInCategory.OST_CommunicationDevices),
    ("Fire Alarm Devices", BuiltInCategory.OST_FireAlarmDevices),
    ("Nurse Call Devices", BuiltInCategory.OST_NurseCallDevices),
    ("Security Devices", BuiltInCategory.OST_SecurityDevices),
    ("Telephone Devices", BuiltInCategory.OST_TelephoneDevices),
    ("Cable Trays", BuiltInCategory.OST_CableTray),
    ("Conduits", BuiltInCategory.OST_Conduit),
    ("Generic Models", BuiltInCategory.OST_GenericModel),
    ("Speciality Equipment", BuiltInCategory.OST_SpecialityEquipment),
]


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
    try:
        for p in element.Parameters:
            if p is None or p.Definition is None:
                continue
            if p.IsReadOnly or p.StorageType == storage_none:
                continue
            names.add(p.Definition.Name)
    except Exception:
        pass
    return names


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


def set_parameter_value(param, value, duplicate_mode):
    if param is None or param.IsReadOnly:
        return False, "read-only or missing"

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
                param.Set(old_s + "; " + new_s)
                return True, "appended"

        st = param.StorageType
        if st == StorageType.String:
            param.Set(value if isinstance(value, str) else str(value))
            return True, "written"
        if st == StorageType.Integer:
            param.Set(int(value))
            return True, "written"
        if st == StorageType.Double:
            param.Set(float(value))
            return True, "written"

        return False, "unsupported storage type"
    except Exception as ex:
        return False, str(ex)


def _find_writable_parameter(target_element, target_param_name):
    if target_element is None or not target_param_name:
        return None, None

    name = target_param_name.strip()
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

    for source_name, owner in search_sources:
        try:
            param = owner.LookupParameter(name)
            if param is not None and not param.IsReadOnly:
                return param, source_name
        except Exception:
            pass

        try:
            params = list(owner.GetParameters(name))
        except Exception:
            params = []

        for param in params:
            if param is not None and not param.IsReadOnly:
                return param, source_name

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
        self.category_items = []
        self.last_transfer_summary = []
        self.element_room_map = {}
        self.room_detection_index = []
        self.selected_link_index = 0

        self._load_links()
        self._populate_link_selector()
        self._build_category_list()
        self._try_seed_current_selection()

    def _load_links(self):
        self.links = [lk for lk in get_link_instances(doc) if lk.GetLinkDocument() is not None]
        if not self.links:
            forms.alert("No loaded Revit links found. Load at least one linked model and retry.", exitscript=True)
            return

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
        if not self.selected_room_params:
            return

        for room_param in sorted(self.selected_room_params.keys()):
            target_param = self.mapping.get(room_param)
            if not target_param:
                continue

            self.lstMappings.Items.Add("[mapped] {0} -> {1}".format(room_param, target_param))

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
        if not getattr(self, "selected_room_params", None) or not getattr(self, "common_params", None):
            return
        self.mapping = {}
        self._auto_match_mappings()
        self._render_mapping_list()

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
        self.common_params = set()
        self.mapping = {}

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

        self._auto_match_mappings()
        self._refresh_mapping_controls()
        self._render_mapping_list()

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
            bb = element.get_BoundingBox(doc.ActiveView)
            if bb is None:
                bb = element.get_BoundingBox(None)
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
        opts = SpatialElementBoundaryOptions()
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

                try:
                    seg_loops = room.GetBoundarySegments(opts)
                except Exception:
                    seg_loops = None

                if seg_loops:
                    for seg_loop in seg_loops:
                        poly = []
                        for seg in seg_loop:
                            try:
                                p0 = seg.GetCurve().GetEndPoint(0)
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

    def _find_linked_room_for_host_point(self, host_point, tol=1.0):
        if host_point is None:
            return None

        if not self.room_detection_index:
            self._build_room_detection_index()

        for room_item in self.room_detection_index:
            try:
                p = room_item["inv_transform"].OfPoint(host_point)
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

        return None

    def _iter_selected_category_elements(self, scope_mode=None):
        allowed_ids = set(self._selected_category_ids())
        if not allowed_ids:
            return []

        scope_mode = scope_mode or self._get_auto_detect_scope()

        items = []
        seen = set()
        for _, bic in TARGET_CATEGORIES:
            if int(bic) not in allowed_ids:
                continue

            try:
                if scope_mode == "Entire Project":
                    coll = (
                        FilteredElementCollector(doc)
                        .OfCategory(bic)
                        .WhereElementIsNotElementType()
                    )
                else:
                    coll = (
                        FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfCategory(bic)
                        .WhereElementIsNotElementType()
                    )
            except Exception:
                coll = []

            for el in coll:
                if el is None:
                    continue
                eid = el.Id.IntegerValue
                if eid in seen:
                    continue
                seen.add(eid)
                items.append(el)

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
        allowed_ids = self._selected_category_ids()
        if not allowed_ids:
            forms.alert("Select at least one target category.")
            return

        scope_mode = self._get_auto_detect_scope()

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
            room_item = self._find_linked_room_for_host_point(pt)
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

    def deselect_all_categories_click(self, sender, e):
        for item in self.category_items:
            item["cb"].IsChecked = False

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
        self._render_mapping_list()

    def remove_mapping_click(self, sender, e):
        room_param = self.cmbMappingRoom.SelectedItem
        if not room_param:
            forms.alert("Pick a room parameter to remove mapping.")
            return

        key = str(room_param).strip()
        if key in self.mapping:
            del self.mapping[key]
        self._render_mapping_list()

    def auto_match_click(self, sender, e):
        if self._get_auto_map_mode() == "disabled":
            forms.alert("Auto-map is disabled. Change 'Auto-map strictness' to enable matching.")
            return
        self._auto_match_mappings()
        self._render_mapping_list()

    def clear_mappings_click(self, sender, e):
        self.mapping = {}
        self._render_mapping_list()

    def _write_transfer_log(self, lines):
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        path = os.path.join(tempfile.gettempdir(), "LinkedRoomTransfer-{0}.log".format(ts))
        with io.open(path, "w", encoding="utf-8") as fp:
            fp.write("\n".join(lines))
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

        tx = Transaction(doc, "Linked Room Parameter Transfer")
        tx.Start()
        try:
            for el in self.selected_elements:
                per_element_values = None

                if auto_room_mode:
                    room_item = self.element_room_map.get(el.Id.IntegerValue)
                    if room_item is None:
                        pt = self._get_element_probe_point(el)
                        room_item = self._find_linked_room_for_host_point(pt)
                        if room_item is not None:
                            self.element_room_map[el.Id.IntegerValue] = room_item

                    if room_item is None:
                        skipped += len(self.mapping)
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
                            continue
                        value = room_data.get("value")

                    if skip_empty and (value is None or value == ""):
                        skipped += 1
                        continue

                    target_param, target_source = _find_writable_parameter(el, target_pname)

                    if target_param is None:
                        failed += 1
                        if len(fail_messages) < 8:
                            fail_messages.append(
                                "{0} -> {1}: target parameter not found or not writable on element {2}".format(
                                    room_pname, target_pname, el.Id.IntegerValue
                                )
                            )
                        continue

                    # Writing type parameters per element can cause inflated counts and wrong behavior
                    # in auto-room mode where different elements may belong to different rooms.
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
                                continue
                            type_write_keys.add(write_key)

                    ok, msg = set_parameter_value(target_param, value, duplicate_mode)
                    if ok:
                        transferred += 1
                        updated_elements.add(el.Id.IntegerValue)
                    else:
                        if "skipped" in msg.lower() or "empty" in msg.lower():
                            skipped += 1
                        else:
                            failed += 1
                            if len(fail_messages) < 8:
                                fail_messages.append(
                                    "{0} -> {1} [{2}]: {3}".format(room_pname, target_pname, target_source, msg)
                                )

            tx.Commit()
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

        if blocked_type_auto:
            summary.append("Type-parameter writes blocked in auto-room mode: {0}".format(blocked_type_auto))
        if skipped_type_dedup:
            summary.append("Duplicate type writes skipped: {0}".format(skipped_type_dedup))

        if fail_messages:
            summary.append("\nSample failures:")
            summary.extend(fail_messages)

        log_path = self._write_transfer_log(summary)
        summary.append("\nLog file: {0}".format(log_path))
        self.last_transfer_summary = list(summary)

        forms.alert("\n".join(summary), title="Linked Room Parameter Transfer")

    def cancel_click(self, sender, e):
        self.Close()


ui = LinkedRoomTransferWindow("WPFWindow.xaml")
ui.ShowDialog()
