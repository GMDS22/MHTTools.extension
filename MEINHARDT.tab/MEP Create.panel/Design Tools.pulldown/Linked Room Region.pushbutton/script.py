# coding: utf8
from __future__ import print_function

import clr
import json
import math
import os
import System

clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationCore")
from System.Collections.Generic import List
from System.Drawing import Color as DrawingColor
from System.Windows.Forms import ColorDialog, DialogResult
from System.Windows.Media import BrushConverter
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from Autodesk.Revit.DB import (
    Color,
    BuiltInParameter,
    BuiltInCategory,
    ElementId,
    FillPatternElement,
    FillPatternTarget,
    CurveLoop,
    FilledRegion,
    FilledRegionType,
    FilteredElementCollector,
    GraphicsStyle,
    GraphicsStyleType,
    Line,
    RevitLinkInstance,
    SpatialElementBoundaryLocation,
    SpatialElementBoundaryOptions,
    SubTransaction,
    Transform,
    Transaction,
    XYZ,
)

from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
config = script.get_config()
_BRUSH_CONVERTER = BrushConverter()
MIN_CURVE_LEN_FT = 0.005
_LAST_SUCCESS_STATE_KEY = "linked_room_region_last_success_state_json"


class RoomSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return element.Category is not None and element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms)
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return True


class LinkedRoomSelectionFilter(ISelectionFilter):
    def __init__(self, host_doc):
        self._host_doc = host_doc

    def AllowElement(self, element):
        # Allow only Revit link instances to be targetable for linked picks.
        try:
            return isinstance(element, RevitLinkInstance)
        except Exception:
            return False

    def AllowReference(self, reference, point):
        try:
            link_inst = self._host_doc.GetElement(reference.ElementId)
            if link_inst is None:
                return False
            link_doc = link_inst.GetLinkDocument()
            if link_doc is None:
                return False

            linked_id = getattr(reference, "LinkedElementId", None)
            if linked_id is None or linked_id == ElementId.InvalidElementId:
                return False

            linked_elem = link_doc.GetElement(linked_id)
            if linked_elem is None or linked_elem.Category is None:
                return False
            return linked_elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms)
        except Exception:
            return False


class FilledRegionSelectionFilter(ISelectionFilter):
    def __init__(self, active_view_id):
        self._active_view_id = active_view_id

    def AllowElement(self, element):
        try:
            if not isinstance(element, FilledRegion):
                return False
        except Exception:
            return False

        try:
            owner_view_id = getattr(element, "OwnerViewId", None)
            if owner_view_id is not None and owner_view_id != ElementId.InvalidElementId:
                return owner_view_id == self._active_view_id
        except Exception:
            pass

        return True

    def AllowReference(self, reference, point):
        return True


def _safe_elem_name(elem, fallback=""):
    if elem is None:
        return fallback
    try:
        return elem.Name or fallback
    except Exception:
        pass
    try:
        return elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or fallback
    except Exception:
        pass
    return fallback


def _get_active_level_elevation(active_view):
    try:
        lvl = active_view.GenLevel
        if lvl is not None:
            return float(lvl.Elevation)
    except Exception:
        pass
    return None


def _is_room_on_active_view_level(room, room_doc, link_inst, active_view, tol=0.5):
    active_elev = _get_active_level_elevation(active_view)
    if active_elev is None:
        return True

    try:
        room_level = room_doc.GetElement(room.LevelId)
        if room_level is None:
            return False
    except Exception:
        return False

    try:
        room_elev = float(room_level.Elevation)
    except Exception:
        return False

    try:
        if link_inst is not None:
            trf = link_inst.GetTotalTransform() if hasattr(link_inst, "GetTotalTransform") else link_inst.GetTransform()
            host_pt = trf.OfPoint(XYZ(0, 0, room_elev))
            host_elev = float(host_pt.Z)
        else:
            host_elev = room_elev
    except Exception:
        host_elev = room_elev

    return abs(host_elev - active_elev) <= tol


def _room_host_probe_point(room, room_doc, link_inst, active_view):
    point = None
    try:
        loc = room.Location
        if loc is not None and hasattr(loc, "Point") and loc.Point is not None:
            point = loc.Point
    except Exception:
        point = None

    if point is None:
        try:
            bb = room.get_BoundingBox(None)
            if bb is not None:
                point = XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5,
                )
        except Exception:
            point = None

    if point is None:
        return None

    if link_inst is not None:
        try:
            trf = link_inst.GetTotalTransform() if hasattr(link_inst, "GetTotalTransform") else link_inst.GetTransform()
            return trf.OfPoint(point)
        except Exception:
            return None

    return point


def _point_inside_active_view_crop(host_point, active_view):
    if host_point is None or active_view is None:
        return False

    try:
        if hasattr(active_view, "CropBoxActive") and not active_view.CropBoxActive:
            return True
    except Exception:
        pass

    try:
        crop = active_view.CropBox
    except Exception:
        crop = None

    if crop is None:
        return True

    try:
        inv = crop.Transform.Inverse
        p = inv.OfPoint(host_point)
        tol = 1e-6
        return (
            (crop.Min.X - tol) <= p.X <= (crop.Max.X + tol)
            and (crop.Min.Y - tol) <= p.Y <= (crop.Max.Y + tol)
            and (crop.Min.Z - tol) <= p.Z <= (crop.Max.Z + tol)
        )
    except Exception:
        return True


class RoomItem(object):
    def __init__(self, room, room_doc, link_inst=None):
        self.room = room
        self.room_doc = room_doc
        self.link_inst = link_inst
        self.room_id = room.Id.IntegerValue
        self.number = ""
        self.name = ""
        self.level_name = ""
        self.link_name = "Host Model"

        if self.link_inst is not None:
            self.link_name = _safe_elem_name(self.link_inst, "Linked Model")

        try:
            self.number = room.Number or ""
        except Exception:
            pass

        try:
            self.name = room.Name or ""
        except Exception:
            pass

        try:
            lvl = room_doc.GetElement(room.LevelId)
            if lvl is not None:
                self.level_name = lvl.Name or ""
        except Exception:
            pass

        parts = []
        if self.number:
            parts.append(self.number)
        if self.name:
            parts.append(self.name)
        display = " - ".join(parts) if parts else "Room {0}".format(self.room_id)
        if self.level_name:
            display = "{0}  |  {1}".format(display, self.level_name)
        if self.link_name:
            display = "{0}  |  {1}".format(display, self.link_name)
        self.display = display


class RegionTypeItem(object):
    def __init__(self, fr_type):
        self.fr_type = fr_type
        self.name = _safe_elem_name(fr_type, "Filled Region Type")


class FillPatternItem(object):
    def __init__(self, pattern_elem):
        self.pattern_elem = pattern_elem
        self.pattern_id = pattern_elem.Id
        self.name = _safe_elem_name(pattern_elem, "Fill Pattern")


class LineStyleItem(object):
    def __init__(self, graphics_style):
        self.graphics_style = graphics_style
        self.style_id = graphics_style.Id
        self.name = _line_style_name(graphics_style)


class LinkedRoomRegionWindow(WPFWindow):
    def __init__(self, xaml_path, initial_state=None):
        WPFWindow.__init__(self, xaml_path)
        self._room_items = []
        self._filtered_room_items = []
        self._region_type_items = []
        self._fill_pattern_items = []
        self._line_style_items = []
        self._foreground_rgb = (17, 119, 187)
        self._background_rgb = (42, 42, 46)
        self._foreground_pattern_id = _solid_fill_pattern_id()
        self._background_pattern_id = _solid_fill_pattern_id()
        self._show_foreground = True
        self._show_background = True
        self._is_masking = False
        self._boundary_line_style_id = _invisible_line_style_id()
        self._existing_region_ids = []
        self._existing_region_line_style_id = None
        self._selected_host_room_ids = set()
        self._selected_link_room_pairs = set()
        self._is_updating_color_ui = False
        self._is_restoring_state = False
        self._is_syncing_room_selection = False
        self.result = None

        self._reload_rooms()
        self._load_region_types()
        self._load_fill_patterns()
        self._load_line_styles()
        self._load_defaults()
        self._apply_initial_state(initial_state)

    def _load_defaults(self):
        last_name = getattr(config, "linked_room_region_last_type_name", "Linked Room Region")
        self.txtNewTypeName.Text = last_name

    def _apply_initial_state(self, state):
        state = state or {}
        self._is_restoring_state = True

        try:
            self._foreground_rgb = tuple(state.get("foreground_rgb", self._foreground_rgb))
            self._background_rgb = tuple(state.get("background_rgb", self._background_rgb))
            saved_fg_pattern_id = _deserialize_element_id(state.get("foreground_pattern_id"))
            saved_bg_pattern_id = _deserialize_element_id(state.get("background_pattern_id"))
            if saved_fg_pattern_id is not None:
                self._foreground_pattern_id = saved_fg_pattern_id
            if saved_bg_pattern_id is not None:
                self._background_pattern_id = saved_bg_pattern_id
            saved_boundary_line_style_id = _deserialize_element_id(state.get("boundary_line_style_id"))
            if saved_boundary_line_style_id is not None:
                self._boundary_line_style_id = saved_boundary_line_style_id
            self._existing_region_ids = _deserialize_element_id_list(state.get("existing_region_ids", []))
            self._existing_region_line_style_id = _deserialize_element_id(state.get("existing_line_style_id"))
            self._show_foreground = bool(state.get("show_foreground", self._show_foreground))
            self._show_background = bool(state.get("show_background", self._show_background))
            self._is_masking = bool(state.get("is_masking", self._is_masking))

            replace_existing = bool(state.get("replace_existing", False))
            create_new_type = state.get("create_new_type")
            apply_colors = state.get("apply_colors")
            new_type_name = state.get("new_type_name")
            room_filter = state.get("room_filter", "")

            if create_new_type is not None:
                self.chkCreateNewType.IsChecked = bool(create_new_type)
            if apply_colors is not None:
                self.chkApplyColors.IsChecked = bool(apply_colors)
            self.chkShowForeground.IsChecked = self._show_foreground
            self.chkShowBackground.IsChecked = self._show_background
            self.chkMasking.IsChecked = self._is_masking

            self.chkReplaceExisting.IsChecked = replace_existing

            if new_type_name:
                self.txtNewTypeName.Text = new_type_name

            remembered_type_name = state.get("selected_region_type_name")
            if remembered_type_name:
                for idx, item in enumerate(self._region_type_items):
                    if item.name == remembered_type_name:
                        self.cmbRegionTypes.SelectedIndex = idx
                        break

            if saved_boundary_line_style_id is None and self._existing_region_line_style_id is not None:
                self._boundary_line_style_id = self._existing_region_line_style_id

            if not state.get("foreground_pattern_id") and not state.get("background_pattern_id") and "is_masking" not in state:
                self._load_selected_region_type_graphics()

            self._refresh_line_style_ui()
            self._refresh_color_ui()
            self._update_existing_region_ui()

            self.txtRoomFilter.Text = room_filter
            self._apply_room_filter()
            self._apply_saved_room_selection(state)
            self._update_summary()
        finally:
            self._is_restoring_state = False

    def _apply_saved_room_selection(self, state):
        self._selected_host_room_ids = set(state.get("selected_host_room_ids", []))
        self._selected_link_room_pairs = set(tuple(pair) for pair in state.get("selected_link_room_pairs", []))
        if not self._selected_host_room_ids and not self._selected_link_room_pairs:
            return

        self._ensure_picked_rooms_present(self._selected_host_room_ids, self._selected_link_room_pairs)
        if self.txtRoomFilter.Text:
            self.txtRoomFilter.Text = ""
            self._apply_room_filter()

        self._restore_room_selection_to_ui()

    def _sync_selected_room_state_from_ui(self):
        if self.lstRooms is None:
            return

        self._selected_host_room_ids = set()
        self._selected_link_room_pairs = set()

        for selected in self.lstRooms.SelectedItems:
            if selected.link_inst is not None:
                self._selected_link_room_pairs.add((selected.link_inst.Id.IntegerValue, selected.room_id))
            else:
                self._selected_host_room_ids.add(selected.room_id)

    def _restore_room_selection_to_ui(self):
        if self.lstRooms is None:
            return

        self._is_syncing_room_selection = True
        try:
            self.lstRooms.UnselectAll()
            for item in self._filtered_room_items:
                if self._room_item_is_picked(item, self._selected_host_room_ids, self._selected_link_room_pairs):
                    self.lstRooms.SelectedItems.Add(item)
        finally:
            self._is_syncing_room_selection = False

    def _selected_room_items(self):
        self._sync_selected_room_state_from_ui()

        selected_items = []
        seen = set()
        for item in self._room_items:
            if not self._room_item_is_picked(item, self._selected_host_room_ids, self._selected_link_room_pairs):
                continue

            key = (item.link_inst.Id.IntegerValue, item.room_id) if item.link_inst is not None else (None, item.room_id)
            if key in seen:
                continue
            seen.add(key)
            selected_items.append(item)

        return selected_items

    def _capture_state(self):
        self._sync_selected_room_state_from_ui()
        selected_host_ids = sorted(self._selected_host_room_ids)
        selected_link_room_pairs = sorted(list(self._selected_link_room_pairs))

        selected_type_name = None
        selected_type_item = self.cmbRegionTypes.SelectedItem
        if selected_type_item is not None:
            selected_type_name = selected_type_item.name

        return {
            "view_id": _serialize_element_id(doc.ActiveView.Id),
            "room_filter": self.txtRoomFilter.Text or "",
            "selected_host_room_ids": selected_host_ids,
            "selected_link_room_pairs": selected_link_room_pairs,
            "selected_region_type_name": selected_type_name,
            "create_new_type": bool(self.chkCreateNewType.IsChecked),
            "new_type_name": (self.txtNewTypeName.Text or "").strip(),
            "apply_colors": bool(self.chkApplyColors.IsChecked),
            "foreground_rgb": tuple(self._foreground_rgb),
            "background_rgb": tuple(self._background_rgb),
            "foreground_pattern_id": _serialize_element_id(self._foreground_pattern_id),
            "background_pattern_id": _serialize_element_id(self._background_pattern_id),
            "show_foreground": bool(self.chkShowForeground.IsChecked),
            "show_background": bool(self.chkShowBackground.IsChecked),
            "is_masking": bool(self.chkMasking.IsChecked),
            "boundary_line_style_id": _serialize_element_id(self._boundary_line_style_id),
            "replace_existing": bool(self.chkReplaceExisting.IsChecked),
            "existing_region_ids": _serialize_element_id_list(self._existing_region_ids),
            "existing_line_style_id": _serialize_element_id(self._existing_region_line_style_id),
        }

    def _request_external_action(self, action_name):
        self.result = {
            "action": action_name,
            "state": self._capture_state(),
        }
        try:
            self.DialogResult = True
        except Exception:
            pass
        self.Close()

    def _refresh_color_ui(self):
        fg_hex = _rgb_to_hex(self._foreground_rgb)
        bg_hex = _rgb_to_hex(self._background_rgb)
        self._is_updating_color_ui = True
        try:
            self.txtForegroundColor.Text = fg_hex
            self.txtBackgroundColor.Text = bg_hex
            self.chkShowForeground.IsChecked = bool(self._show_foreground)
            self.chkShowBackground.IsChecked = bool(self._show_background)
            self.chkMasking.IsChecked = bool(self._is_masking)
            self._select_pattern_item(self.cmbForegroundPattern, self._foreground_pattern_id)
            self._select_pattern_item(self.cmbBackgroundPattern, self._background_pattern_id)
            try:
                self.brdForegroundPreview.Background = _BRUSH_CONVERTER.ConvertFromString(fg_hex)
                self.brdBackgroundPreview.Background = _BRUSH_CONVERTER.ConvertFromString(bg_hex)
                self.brdForegroundPreview.Opacity = 1.0 if self._show_foreground else 0.35
                self.brdBackgroundPreview.Opacity = 1.0 if self._show_background else 0.35
            except Exception:
                pass
        finally:
            self._is_updating_color_ui = False

        apply_graphics = bool(self.chkApplyColors.IsChecked)
        fg_enabled = apply_graphics and self._show_foreground
        bg_enabled = apply_graphics and self._show_background
        self.txtForegroundColor.IsEnabled = fg_enabled
        self.btnPickForeground.IsEnabled = fg_enabled
        self.cmbForegroundPattern.IsEnabled = fg_enabled
        self.txtBackgroundColor.IsEnabled = bg_enabled
        self.btnPickBackground.IsEnabled = bg_enabled
        self.cmbBackgroundPattern.IsEnabled = bg_enabled

    def _refresh_line_style_ui(self):
        self._select_line_style_item(self.cmbBoundaryLineStyle, self._boundary_line_style_id)

    def _update_existing_region_ui(self):
        replace_existing = bool(getattr(self.chkReplaceExisting, "IsChecked", False))
        self.btnSelectExistingRegions.IsEnabled = replace_existing

        if not replace_existing:
            self.txtExistingRegionsSummary.Text = "Create new filled regions from the selected linked rooms."
            return

        count = len(self._existing_region_ids)
        if count == 0:
            self.txtExistingRegionsSummary.Text = "Select existing filled regions in the active view to replace."
        else:
            self.txtExistingRegionsSummary.Text = "Existing filled regions selected: {0}".format(count)

    def _load_region_types(self):
        self._region_type_items = []
        self.cmbRegionTypes.Items.Clear()

        region_types = (
            FilteredElementCollector(doc)
            .OfClass(FilledRegionType)
            .ToElements()
        )
        for fr_type in sorted(region_types, key=lambda t: _safe_elem_name(t, "").lower()):
            item = RegionTypeItem(fr_type)
            self._region_type_items.append(item)
            self.cmbRegionTypes.Items.Add(item)

        self.cmbRegionTypes.DisplayMemberPath = "name"

        if self._region_type_items:
            remembered = getattr(config, "linked_room_region_last_type_name", "")
            found_index = -1
            if remembered:
                for idx, item in enumerate(self._region_type_items):
                    if item.name == remembered:
                        found_index = idx
                        break
            self.cmbRegionTypes.SelectedIndex = found_index if found_index >= 0 else 0

    def _load_fill_patterns(self):
        self._fill_pattern_items = []
        self.cmbForegroundPattern.Items.Clear()
        self.cmbBackgroundPattern.Items.Clear()

        patterns = []
        try:
            patterns = list(FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements())
        except Exception:
            patterns = []

        drafting_patterns = []
        for pattern in patterns:
            try:
                fill_pattern = pattern.GetFillPattern()
                if fill_pattern is not None and fill_pattern.Target == FillPatternTarget.Drafting:
                    drafting_patterns.append(pattern)
            except Exception:
                continue

        for pattern in sorted(drafting_patterns, key=lambda p: _safe_elem_name(p, "").lower()):
            item = FillPatternItem(pattern)
            self._fill_pattern_items.append(item)
            self.cmbForegroundPattern.Items.Add(item)
            self.cmbBackgroundPattern.Items.Add(item)

        self.cmbForegroundPattern.DisplayMemberPath = "name"
        self.cmbBackgroundPattern.DisplayMemberPath = "name"

    def _load_line_styles(self):
        self._line_style_items = []
        self.cmbBoundaryLineStyle.Items.Clear()

        styles = []
        try:
            styles = list(FilteredElementCollector(doc).OfClass(GraphicsStyle).ToElements())
        except Exception:
            styles = []

        visible_styles = []
        lines_parent_id = None
        try:
            lines_parent_id = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines).Id.IntegerValue
        except Exception:
            lines_parent_id = None

        seen = set()
        for style in styles:
            try:
                if style.GraphicsStyleType != GraphicsStyleType.Projection:
                    continue
            except Exception:
                continue

            try:
                category = style.GraphicsStyleCategory
            except Exception:
                category = None

            include = False
            try:
                if category is not None and category.Id.IntegerValue == -2000064:
                    include = True
                elif category is not None and category.Parent is not None and lines_parent_id is not None:
                    include = category.Parent.Id.IntegerValue == lines_parent_id
            except Exception:
                include = False

            if not include:
                continue

            key = _serialize_element_id(style.Id)
            if key in seen:
                continue
            seen.add(key)
            visible_styles.append(style)

        invisible_id = _serialize_element_id(_invisible_line_style_id())
        visible_styles.sort(key=lambda style: ((0 if _serialize_element_id(style.Id) == invisible_id else 1), _line_style_name(style).lower()))

        for style in visible_styles:
            item = LineStyleItem(style)
            self._line_style_items.append(item)
            self.cmbBoundaryLineStyle.Items.Add(item)

        self.cmbBoundaryLineStyle.DisplayMemberPath = "name"
        self._refresh_line_style_ui()

    def _select_line_style_item(self, combo, line_style_id):
        if combo is None or combo.Items.Count == 0:
            return

        target_id = line_style_id
        if target_id is None or target_id == ElementId.InvalidElementId:
            target_id = _invisible_line_style_id()

        found_index = -1
        target_value = _serialize_element_id(target_id)
        for idx, item in enumerate(self._line_style_items):
            if _serialize_element_id(item.style_id) == target_value:
                found_index = idx
                break

        combo.SelectedIndex = found_index if found_index >= 0 else 0

    def _get_selected_line_style_id(self):
        item = self.cmbBoundaryLineStyle.SelectedItem if self.cmbBoundaryLineStyle is not None else None
        if item is None:
            return _invisible_line_style_id()
        return item.style_id

    def _select_pattern_item(self, combo, pattern_id):
        if combo is None or combo.Items.Count == 0:
            return

        target_id = pattern_id
        if target_id is None or target_id == ElementId.InvalidElementId:
            target_id = _solid_fill_pattern_id()

        found_index = -1
        target_value = _serialize_element_id(target_id)
        for idx, item in enumerate(self._fill_pattern_items):
            if _serialize_element_id(item.pattern_id) == target_value:
                found_index = idx
                break

        combo.SelectedIndex = found_index if found_index >= 0 else 0

    def _get_selected_pattern_id(self, combo):
        item = combo.SelectedItem if combo is not None else None
        if item is None:
            return ElementId.InvalidElementId
        return item.pattern_id

    def _load_selected_region_type_graphics(self):
        item = self.cmbRegionTypes.SelectedItem
        if item is None:
            return

        graphics = _read_region_type_graphics(item.fr_type)
        self._foreground_rgb = graphics["foreground_rgb"]
        self._background_rgb = graphics["background_rgb"]
        self._foreground_pattern_id = graphics["foreground_pattern_id"]
        self._background_pattern_id = graphics["background_pattern_id"]
        self._show_foreground = graphics["show_foreground"]
        self._show_background = graphics["show_background"]
        self._is_masking = _read_region_type_is_masking(item.fr_type)

    def _pick_color(self, current_rgb, title):
        color_value = forms.ask_for_string(
            default=_rgb_to_hex(current_rgb),
            prompt="Enter color as #RRGGBB or R,G,B.",
            title=title,
        )
        if not color_value:
            return None

        parsed_rgb = _parse_rgb_text(color_value)
        if parsed_rgb is None:
            forms.alert("Invalid color value. Use #RRGGBB or R,G,B.", title="Linked Room Region")
            return None

        return parsed_rgb

    def on_pick_foreground(self, sender, args):
        self._request_external_action("pick_foreground_color")

    def on_pick_background(self, sender, args):
        self._request_external_action("pick_background_color")

    def on_region_type_changed(self, sender, args):
        if self._is_restoring_state:
            return
        self._load_selected_region_type_graphics()
        self._refresh_color_ui()

    def on_apply_graphics_changed(self, sender, args):
        self._refresh_color_ui()

    def on_show_foreground_changed(self, sender, args):
        self._show_foreground = bool(self.chkShowForeground.IsChecked)
        self._refresh_color_ui()

    def on_show_background_changed(self, sender, args):
        self._show_background = bool(self.chkShowBackground.IsChecked)
        self._refresh_color_ui()

    def on_masking_changed(self, sender, args):
        self._is_masking = bool(self.chkMasking.IsChecked)
        self._refresh_color_ui()

    def on_boundary_line_style_changed(self, sender, args):
        if self.cmbBoundaryLineStyle is None:
            return
        self._boundary_line_style_id = self._get_selected_line_style_id()

    def on_foreground_pattern_changed(self, sender, args):
        if self._is_updating_color_ui:
            return
        self._foreground_pattern_id = self._get_selected_pattern_id(self.cmbForegroundPattern)

    def on_background_pattern_changed(self, sender, args):
        if self._is_updating_color_ui:
            return
        self._background_pattern_id = self._get_selected_pattern_id(self.cmbBackgroundPattern)

    def _apply_color_text_input(self, textbox, attr_name):
        if self._is_updating_color_ui:
            return

        parsed_rgb = _parse_rgb_text(textbox.Text)
        if parsed_rgb is None:
            return

        if getattr(self, attr_name) == parsed_rgb:
            return

        setattr(self, attr_name, parsed_rgb)
        self._refresh_color_ui()

    def on_foreground_text_changed(self, sender, args):
        self._apply_color_text_input(self.txtForegroundColor, "_foreground_rgb")

    def on_background_text_changed(self, sender, args):
        self._apply_color_text_input(self.txtBackgroundColor, "_background_rgb")

    def _reload_rooms(self):
        self._room_items = []
        active_view = doc.ActiveView

        # Primary: rooms from all loaded links filtered to current view level.
        link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements())
        for link_inst in link_instances:
            link_doc = None
            try:
                link_doc = link_inst.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc is None:
                continue

            link_rooms = (
                FilteredElementCollector(link_doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .ToElements()
            )

            visible_items = []
            level_items = []
            all_items = []

            for room in link_rooms:
                try:
                    if hasattr(room, "Area") and room.Area <= 0:
                        continue
                except Exception:
                    continue

                item = RoomItem(room, link_doc, link_inst)
                all_items.append(item)

                host_pt = _room_host_probe_point(room, link_doc, link_inst, active_view)
                if _point_inside_active_view_crop(host_pt, active_view):
                    visible_items.append(item)

                if _is_room_on_active_view_level(room, link_doc, link_inst, active_view):
                    level_items.append(item)

            if visible_items:
                self._room_items.extend(visible_items)
            elif level_items:
                self._room_items.extend(level_items)
            else:
                self._room_items.extend(all_items)

        # Fallback: include host rooms if no linked rooms are found.
        if not self._room_items:
            host_rooms = (
                FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .ToElements()
            )
            for room in host_rooms:
                try:
                    if hasattr(room, "Area") and room.Area <= 0:
                        continue
                except Exception:
                    continue

                item = RoomItem(room, doc, None)
                host_pt = _room_host_probe_point(room, doc, None, active_view)
                if _point_inside_active_view_crop(host_pt, active_view):
                    self._room_items.append(item)

        self._room_items.sort(key=lambda r: (r.level_name.lower(), r.number.lower(), r.name.lower()))
        self._apply_room_filter()

        if not self._room_items:
            self.txtSummary.Text = "No rooms detected in active view."

    def _apply_room_filter(self):
        query = (self.txtRoomFilter.Text or "").strip().lower()
        if not query:
            filtered = list(self._room_items)
        else:
            filtered = []
            for item in self._room_items:
                hay = "{0} {1} {2} {3}".format(item.number, item.name, item.level_name, item.link_name).lower()
                if query in hay:
                    filtered.append(item)
        self._render_rooms(filtered)

    def _render_rooms(self, room_items):
        if not self._is_syncing_room_selection:
            self._sync_selected_room_state_from_ui()
        self.lstRooms.Items.Clear()
        self._filtered_room_items = list(room_items)
        for item in room_items:
            self.lstRooms.Items.Add(item)
        self.lstRooms.DisplayMemberPath = "display"
        self._restore_room_selection_to_ui()
        self._update_summary()

    def _update_summary(self):
        room_count = len(self._selected_host_room_ids) + len(self._selected_link_room_pairs)
        total_count = len(self._filtered_room_items)
        self.txtSummary.Text = "Selected rooms: {0} | Visible rooms: {1}".format(room_count, total_count)

    def on_filter_text_changed(self, sender, args):
        self._apply_room_filter()

    def on_refresh_rooms(self, sender, args):
        self._reload_rooms()

    def _room_item_is_picked(self, item, picked_host_ids, picked_link_pairs):
        if item is None:
            return False

        if item.link_inst is not None:
            pair = (item.link_inst.Id.IntegerValue, item.room_id)
            return pair in picked_link_pairs

        return item.room_id in picked_host_ids

    def _ensure_picked_rooms_present(self, picked_host_ids, picked_link_pairs):
        if not picked_host_ids and not picked_link_pairs:
            return False

        existing_host_ids = set()
        existing_link_pairs = set()
        for item in self._room_items:
            if item.link_inst is not None:
                existing_link_pairs.add((item.link_inst.Id.IntegerValue, item.room_id))
            else:
                existing_host_ids.add(item.room_id)

        added = False

        missing_host_ids = picked_host_ids.difference(existing_host_ids)
        for room_id in missing_host_ids:
            try:
                room = doc.GetElement(ElementId(room_id))
                if room is None:
                    continue
                self._room_items.append(RoomItem(room, doc, None))
                added = True
            except Exception:
                continue

        missing_link_pairs = picked_link_pairs.difference(existing_link_pairs)
        link_cache = {}
        for link_id, room_id in missing_link_pairs:
            try:
                if link_id not in link_cache:
                    link_inst = doc.GetElement(ElementId(link_id))
                    link_doc = link_inst.GetLinkDocument() if link_inst is not None else None
                    link_cache[link_id] = (link_inst, link_doc)
                link_inst, link_doc = link_cache.get(link_id, (None, None))
                if link_inst is None or link_doc is None:
                    continue

                room = link_doc.GetElement(ElementId(room_id))
                if room is None:
                    continue

                self._room_items.append(RoomItem(room, link_doc, link_inst))
                added = True
            except Exception:
                continue

        if added:
            self._room_items.sort(key=lambda r: (r.level_name.lower(), r.number.lower(), r.name.lower()))

        return added

    def on_select_rooms(self, sender, args):
        self._request_external_action("pick_rooms")

    def on_rooms_selection_changed(self, sender, args):
        if not self._is_syncing_room_selection:
            self._sync_selected_room_state_from_ui()
        self._update_summary()

    def on_replace_existing_changed(self, sender, args):
        if not bool(self.chkReplaceExisting.IsChecked):
            self._existing_region_ids = []
            self._existing_region_line_style_id = None
        self._update_existing_region_ui()

    def on_select_existing_regions(self, sender, args):
        self._request_external_action("pick_existing_regions")

    def on_cancel(self, sender, args):
        self.result = None
        try:
            self.DialogResult = False
        except Exception:
            pass
        self.Close()

    def on_create(self, sender, args):
        try:
            selected_type_item = self.cmbRegionTypes.SelectedItem

            if selected_type_item is None:
                forms.alert("No Filled Region Type is available.", title="Linked Room Region")
                return

            room_items = self._selected_room_items()

            if not room_items:
                forms.alert("Select at least one linked room.", title="Linked Room Region")
                return

            create_new = bool(self.chkCreateNewType.IsChecked)
            new_name = (self.txtNewTypeName.Text or "").strip()

            if create_new and not new_name:
                forms.alert("Enter a new Filled Region Type name.", title="Linked Room Region")
                return

            replace_existing = bool(self.chkReplaceExisting.IsChecked)
            if replace_existing and not self._existing_region_ids:
                forms.alert("Select at least one existing filled region to replace.", title="Linked Room Region")
                return

            self.result = {
                "action": "create",
                "state": self._capture_state(),
                "data": {
                    "room_items": room_items,
                    "base_type": selected_type_item.fr_type,
                    "create_new_type": create_new,
                    "new_type_name": new_name,
                    "apply_colors": bool(self.chkApplyColors.IsChecked),
                    "foreground_rgb": self._foreground_rgb,
                    "background_rgb": self._background_rgb,
                    "foreground_pattern_id": self._foreground_pattern_id,
                    "background_pattern_id": self._background_pattern_id,
                    "show_foreground": bool(self.chkShowForeground.IsChecked),
                    "show_background": bool(self.chkShowBackground.IsChecked),
                    "is_masking": bool(self.chkMasking.IsChecked),
                    "boundary_line_style_id": self._get_selected_line_style_id(),
                    "replace_existing": replace_existing,
                    "existing_region_ids": list(self._existing_region_ids),
                    "existing_line_style_id": self._existing_region_line_style_id,
                },
            }

            try:
                self.DialogResult = True
            except Exception:
                pass
            self.Close()
        except Exception as ex:
            logger.debug("Create button failed: {0}".format(ex))
            forms.alert("Failed to prepare the filled region request.\n\n{0}".format(ex), title="Linked Room Region")


def _rgb_to_hex(rgb):
    return "#{0:02X}{1:02X}{2:02X}".format(int(rgb[0]), int(rgb[1]), int(rgb[2]))


def _serialize_element_id(element_id):
    if element_id is None:
        return None
    try:
        return element_id.IntegerValue
    except Exception:
        return None


def _deserialize_element_id(element_id_value):
    if element_id_value in (None, ""):
        return None
    try:
        return ElementId(int(element_id_value))
    except Exception:
        return None


def _serialize_element_id_list(element_ids):
    serialized = []
    for element_id in element_ids or []:
        value = _serialize_element_id(element_id)
        if value is not None:
            serialized.append(value)
    return serialized


def _deserialize_element_id_list(element_id_values):
    deserialized = []
    for value in element_id_values or []:
        element_id = _deserialize_element_id(value)
        if element_id is not None:
            deserialized.append(element_id)
    return deserialized


def _sanitize_persisted_state(state):
    state = state or {}
    allowed_keys = [
        "view_id",
        "selected_region_type_name",
        "create_new_type",
        "new_type_name",
        "apply_colors",
        "foreground_rgb",
        "background_rgb",
        "foreground_pattern_id",
        "background_pattern_id",
        "show_foreground",
        "show_background",
        "is_masking",
        "boundary_line_style_id",
    ]

    sanitized = {}
    for key in allowed_keys:
        if key in state:
            sanitized[key] = state.get(key)

    sanitized["foreground_rgb"] = list(sanitized.get("foreground_rgb", (17, 119, 187)))
    sanitized["background_rgb"] = list(sanitized.get("background_rgb", (42, 42, 46)))
    sanitized["new_type_name"] = sanitized.get("new_type_name", "") or ""
    return sanitized


def _load_persisted_success_state(current_view_id=None):
    raw_value = getattr(config, _LAST_SUCCESS_STATE_KEY, "")
    if not raw_value:
        return {}

    try:
        state = json.loads(raw_value)
    except Exception as ex:
        logger.debug("Failed to parse Linked Room Region saved state: {0}".format(ex))
        return {}

    if not isinstance(state, dict):
        return {}

    state = _sanitize_persisted_state(state)
    if current_view_id is not None and state.get("view_id") != current_view_id:
        state["view_id"] = current_view_id

    return state


def _save_persisted_success_state(state):
    sanitized = _sanitize_persisted_state(state)
    try:
        setattr(config, _LAST_SUCCESS_STATE_KEY, json.dumps(sanitized))
        if sanitized.get("new_type_name"):
            config.linked_room_region_last_type_name = sanitized.get("new_type_name")
        script.save_config()
    except Exception as ex:
        logger.debug("Failed to save Linked Room Region state: {0}".format(ex))


def _parse_rgb_text(value):
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    if text.startswith("#"):
        hex_text = text[1:]
        if len(hex_text) != 6:
            return None
        try:
            return (
                int(hex_text[0:2], 16),
                int(hex_text[2:4], 16),
                int(hex_text[4:6], 16),
            )
        except Exception:
            return None

    parts = [part.strip() for part in text.replace(";", ",").split(",") if part.strip()]
    if len(parts) != 3:
        return None

    try:
        rgb = tuple(max(0, min(255, int(part))) for part in parts)
    except Exception:
        return None

    if len(rgb) != 3:
        return None

    return rgb


def _make_db_color(rgb):
    return Color(int(rgb[0]), int(rgb[1]), int(rgb[2]))


def _solid_fill_pattern_id():
    try:
        pattern = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, "<Solid fill>")
        if pattern is not None:
            return pattern.Id
    except Exception:
        pass

    try:
        for pattern in FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements():
            try:
                fill_pattern = pattern.GetFillPattern()
                if fill_pattern is not None and fill_pattern.IsSolidFill:
                    return pattern.Id
            except Exception:
                continue
    except Exception:
        pass

    return ElementId.InvalidElementId


def _line_style_name(graphics_style):
    if graphics_style is None:
        return "Line Style"

    try:
        category = graphics_style.GraphicsStyleCategory
        if category is not None and category.Name:
            name = category.Name
            if name.startswith("<") and name.endswith(">"):
                return name[1:-1]
            return name
    except Exception:
        pass

    return _safe_elem_name(graphics_style, "Line Style")


def _invisible_line_style_id():
    try:
        for style in FilteredElementCollector(doc).OfClass(GraphicsStyle).ToElements():
            try:
                if style.GraphicsStyleType != GraphicsStyleType.Projection:
                    continue
            except Exception:
                continue

            try:
                category = style.GraphicsStyleCategory
                if category is not None and category.Id.IntegerValue == -2000064:
                    return style.Id
                if _line_style_name(style).lower() == "invisible lines":
                    return style.Id
            except Exception:
                continue
    except Exception:
        pass

    return ElementId.InvalidElementId


def _get_region_type_pattern_id(region_type, attr_name):
    if region_type is None:
        return ElementId.InvalidElementId

    try:
        pattern_id = getattr(region_type, attr_name)
        if pattern_id is not None:
            return pattern_id
    except Exception:
        pass

    getter_name = "Get{0}".format(attr_name)
    try:
        getter = getattr(region_type, getter_name)
        pattern_id = getter()
        if pattern_id is not None:
            return pattern_id
    except Exception:
        pass

    return ElementId.InvalidElementId


def _set_region_type_pattern_id(region_type, attr_name, pattern_id):
    if region_type is None:
        return

    value = pattern_id if pattern_id is not None else ElementId.InvalidElementId

    try:
        setattr(region_type, attr_name, value)
        return
    except Exception:
        pass

    setter_name = "Set{0}".format(attr_name)
    try:
        setter = getattr(region_type, setter_name)
        setter(value)
    except Exception:
        pass


def _read_region_type_graphics(region_type):
    default_fg = (17, 119, 187)
    default_bg = (42, 42, 46)
    solid_fill_id = _solid_fill_pattern_id()
    if region_type is None:
        return {
            "foreground_rgb": default_fg,
            "background_rgb": default_bg,
            "foreground_pattern_id": solid_fill_id,
            "background_pattern_id": solid_fill_id,
            "show_foreground": True,
            "show_background": True,
        }

    fg = default_fg
    bg = default_bg
    try:
        c = region_type.ForegroundPatternColor
        fg = (int(c.Red), int(c.Green), int(c.Blue))
    except Exception:
        pass
    try:
        c = region_type.BackgroundPatternColor
        bg = (int(c.Red), int(c.Green), int(c.Blue))
    except Exception:
        pass

    fg_pattern_id = _get_region_type_pattern_id(region_type, "ForegroundPatternId")
    bg_pattern_id = _get_region_type_pattern_id(region_type, "BackgroundPatternId")
    show_foreground = fg_pattern_id is not None and fg_pattern_id != ElementId.InvalidElementId
    show_background = bg_pattern_id is not None and bg_pattern_id != ElementId.InvalidElementId

    if not show_foreground:
        fg_pattern_id = solid_fill_id
    if not show_background:
        bg_pattern_id = solid_fill_id

    return {
        "foreground_rgb": fg,
        "background_rgb": bg,
        "foreground_pattern_id": fg_pattern_id,
        "background_pattern_id": bg_pattern_id,
        "show_foreground": show_foreground,
        "show_background": show_background,
    }


def _read_region_type_is_masking(region_type):
    if region_type is None:
        return False

    try:
        return bool(region_type.IsMasking)
    except Exception:
        pass

    masking_param = getattr(BuiltInParameter, "FILLED_REGION_MASKING", None)
    if masking_param is not None:
        try:
            param = region_type.get_Parameter(masking_param)
            if param is not None:
                return bool(param.AsInteger())
        except Exception:
            pass

    return False


def _line_style_name_from_id(line_style_id):
    if line_style_id is None or line_style_id == ElementId.InvalidElementId:
        return "Invisible Lines"

    try:
        return _line_style_name(doc.GetElement(line_style_id))
    except Exception:
        return "Line Style"


def _is_invisible_line_style_id(line_style_id):
    if line_style_id is None or line_style_id == ElementId.InvalidElementId:
        return True

    try:
        style = doc.GetElement(line_style_id)
    except Exception:
        style = None

    if style is None:
        return False

    try:
        category = style.GraphicsStyleCategory
        if category is not None and category.Id.IntegerValue == -2000064:
            return True
    except Exception:
        pass

    try:
        return _line_style_name(style).strip().lower() == "invisible lines"
    except Exception:
        return False


def _resolve_visible_region_graphics(foreground_pattern_id, background_pattern_id, show_foreground, show_background, line_style_id):
    resolved_fg_pattern_id = foreground_pattern_id
    resolved_bg_pattern_id = background_pattern_id

    if resolved_fg_pattern_id is None or resolved_fg_pattern_id == ElementId.InvalidElementId:
        resolved_fg_pattern_id = _solid_fill_pattern_id()
    if resolved_bg_pattern_id is None or resolved_bg_pattern_id == ElementId.InvalidElementId:
        resolved_bg_pattern_id = _solid_fill_pattern_id()

    resolved_show_foreground = bool(show_foreground)
    resolved_show_background = bool(show_background)

    # Prevent silently creating fully invisible regions when both fills are off
    # and the boundary line style is Invisible Lines.
    if not resolved_show_foreground and not resolved_show_background and _is_invisible_line_style_id(line_style_id):
        resolved_show_foreground = True

    return resolved_fg_pattern_id, resolved_bg_pattern_id, resolved_show_foreground, resolved_show_background


def _set_region_type_is_masking(region_type, is_masking):
    if region_type is None:
        return

    try:
        region_type.IsMasking = bool(is_masking)
        return
    except Exception:
        pass

    masking_param = getattr(BuiltInParameter, "FILLED_REGION_MASKING", None)
    if masking_param is not None:
        try:
            param = region_type.get_Parameter(masking_param)
            if param is not None and not param.IsReadOnly:
                param.Set(1 if is_masking else 0)
        except Exception:
            pass


def _read_region_type_colors(region_type):
    graphics = _read_region_type_graphics(region_type)
    return graphics["foreground_rgb"], graphics["background_rgb"]


def _apply_region_type_colors(region_type, foreground_rgb, background_rgb, foreground_pattern_id=None, background_pattern_id=None, show_foreground=True, show_background=True):
    if region_type is None:
        return

    fg = _make_db_color(foreground_rgb)
    bg = _make_db_color(background_rgb)

    try:
        region_type.ForegroundPatternColor = fg
    except Exception:
        try:
            region_type.SetForegroundPatternColor(fg)
        except Exception:
            pass

    try:
        region_type.BackgroundPatternColor = bg
    except Exception:
        try:
            region_type.SetBackgroundPatternColor(bg)
        except Exception:
            pass

    _set_region_type_pattern_id(
        region_type,
        "ForegroundPatternId",
        foreground_pattern_id if show_foreground else ElementId.InvalidElementId,
    )
    _set_region_type_pattern_id(
        region_type,
        "BackgroundPatternId",
        background_pattern_id if show_background else ElementId.InvalidElementId,
    )


def _curve_point_key(pt, tol=0.01):
    return (
        round(float(pt.X) / tol) * tol,
        round(float(pt.Y) / tol) * tol,
        round(float(pt.Z) / tol) * tol,
    )


def _curve_key_undirected(curve, tol=0.01):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return None

    a = _curve_point_key(p0, tol)
    b = _curve_point_key(p1, tol)

    curve_type = "Curve"
    try:
        curve_type = curve.GetType().Name
    except Exception:
        pass

    try:
        mid = curve.Evaluate(0.5, True)
        m = _curve_point_key(mid, tol)
    except Exception:
        m = (None, None, None)

    if a <= b:
        return (curve_type, a, b, m)
    else:
        return (curve_type, b, a, m)


def _distance(p1, p2):
    dx = float(p1.X) - float(p2.X)
    dy = float(p1.Y) - float(p2.Y)
    dz = float(p1.Z) - float(p2.Z)
    return (dx * dx + dy * dy + dz * dz) ** 0.5


def _orient_curve_to_start(curve, start_point, tol=0.01):
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    if _distance(p0, start_point) <= tol:
        return curve
    if _distance(p1, start_point) <= tol:
        try:
            return curve.CreateReversed()
        except Exception:
            return None
    return None


def _is_effectively_line(curve):
    try:
        return curve.GetType().Name == "Line"
    except Exception:
        return False


def _closed_curve_sequence_points(curves, tol=0.02):
    if not curves:
        return None

    points = []
    try:
        first_start = curves[0].GetEndPoint(0)
    except Exception:
        return None

    points.append(first_start)
    cursor = first_start
    for curve in curves:
        try:
            start_pt = curve.GetEndPoint(0)
            end_pt = curve.GetEndPoint(1)
        except Exception:
            return None

        if _distance(start_pt, cursor) > tol:
            return None
        points.append(end_pt)
        cursor = end_pt

    if _distance(points[0], points[-1]) > tol:
        return None
    return points


def _point_is_collinear_2d(prev_pt, curr_pt, next_pt, tol=0.01):
    ax = float(curr_pt.X) - float(prev_pt.X)
    ay = float(curr_pt.Y) - float(prev_pt.Y)
    bx = float(next_pt.X) - float(curr_pt.X)
    by = float(next_pt.Y) - float(curr_pt.Y)
    len_a = math.sqrt((ax * ax) + (ay * ay))
    len_b = math.sqrt((bx * bx) + (by * by))

    if len_a <= tol or len_b <= tol:
        return True

    cross_val = abs(_cross_2d(ax, ay, bx, by))
    return cross_val <= (tol * max(len_a, len_b, 1.0))


def _simplify_closed_curve_sequence(curves, tol=0.02):
    if not curves or len(curves) < 3:
        return curves
    if any(not _is_effectively_line(curve) for curve in curves):
        return curves

    points = _closed_curve_sequence_points(curves, tol)
    if not points or len(points) < 4:
        return curves

    vertices = list(points[:-1])
    changed = True
    while changed and len(vertices) > 3:
        changed = False
        for idx in range(len(vertices)):
            prev_pt = vertices[idx - 1]
            curr_pt = vertices[idx]
            next_pt = vertices[(idx + 1) % len(vertices)]

            if _distance(prev_pt, curr_pt) <= tol or _distance(curr_pt, next_pt) <= tol or _point_is_collinear_2d(prev_pt, curr_pt, next_pt, tol):
                vertices.pop(idx)
                changed = True
                break

    simplified = []
    for idx in range(len(vertices)):
        start_pt = vertices[idx]
        end_pt = vertices[(idx + 1) % len(vertices)]
        if _distance(start_pt, end_pt) <= tol:
            continue
        try:
            simplified.append(Line.CreateBound(start_pt, end_pt))
        except Exception:
            return curves

    return simplified if len(simplified) >= 3 else curves


def _curve_loop_from_curves(curves, tol=0.02):
    simplified_curves = _simplify_closed_curve_sequence(curves, tol)
    loop = CurveLoop()
    try:
        for curve in simplified_curves:
            loop.Append(curve)
    except Exception:
        return None

    if _curve_loop_area_xy(loop) <= 1e-6:
        return None
    return loop


def _project_point_to_elevation(point, elevation):
    return XYZ(float(point.X), float(point.Y), float(elevation))


def _curve_to_planar_segments(curve, elevation, z_tol=1e-4, min_len=MIN_CURVE_LEN_FT):
    if curve is None:
        return []

    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        if abs(float(p0.Z) - float(p1.Z)) <= z_tol:
            q0 = _project_point_to_elevation(p0, elevation)
            q1 = _project_point_to_elevation(p1, elevation)
            if _distance(q0, q1) > min_len:
                return [Line.CreateBound(q0, q1)]
    except Exception:
        pass

    try:
        pts = list(curve.Tessellate())
    except Exception:
        pts = []

    if len(pts) < 2:
        return []

    projected = []
    for pt in pts:
        q = _project_point_to_elevation(pt, elevation)
        if projected and _distance(projected[-1], q) <= min_len:
            continue
        projected.append(q)

    segments = []
    for idx in range(len(projected) - 1):
        p0 = projected[idx]
        p1 = projected[idx + 1]
        if _distance(p0, p1) <= min_len:
            continue
        try:
            segments.append(Line.CreateBound(p0, p1))
        except Exception:
            continue

    return segments


def _active_view_plane_elevation(view):
    try:
        level = view.GenLevel
        if level is not None:
            return level.Elevation
    except Exception:
        pass

    try:
        origin = view.Origin
        if origin is not None:
            return origin.Z
    except Exception:
        pass

    return 0.0


def _boundary_segment_to_host_curve(seg, link_inst=None):
    curve = None
    try:
        curve = seg.GetCurve()
    except Exception:
        try:
            curve = seg.Curve
        except Exception:
            curve = None

    if curve is None:
        return None

    if link_inst is None:
        return curve

    try:
        trf = link_inst.GetTotalTransform() if hasattr(link_inst, "GetTotalTransform") else link_inst.GetTransform()
        return curve.CreateTransformed(trf)
    except Exception:
        return None


def _build_ordered_curve_loop(curves, tol=0.02):
    if not curves or len(curves) < 3:
        return None

    ordered = []
    first_curve = curves[0]
    ordered.append(first_curve)

    try:
        start_pt = first_curve.GetEndPoint(0)
        cursor = first_curve.GetEndPoint(1)
    except Exception:
        return None

    for curve in curves[1:]:
        oriented = _orient_curve_to_start(curve, cursor, tol)
        if oriented is None:
            return None
        ordered.append(oriented)
        cursor = oriented.GetEndPoint(1)

    if _distance(cursor, start_pt) > tol:
        return None

    return _curve_loop_from_curves(ordered, tol)


def _collect_selected_room_loops(room_items, target_elevation, tol=0.02):
    opts = SpatialElementBoundaryOptions()
    try:
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center
    except Exception:
        pass

    all_curves = []
    room_loop_records = []
    skipped_loops = 0

    for item in room_items:
        room = item.room
        link_inst = item.link_inst
        try:
            seglists = room.GetBoundarySegments(opts)
        except Exception:
            seglists = None

        if not seglists:
            continue

        room_loop_candidates = []
        for seglist in seglists:
            planar_curves = []
            for seg in seglist:
                host_curve = _boundary_segment_to_host_curve(seg, link_inst)
                if host_curve is None:
                    continue
                planar_segments = _curve_to_planar_segments(host_curve, target_elevation)
                if planar_segments:
                    planar_curves.extend(planar_segments)

            if len(planar_curves) < 3:
                skipped_loops += 1
                continue

            loop = _build_ordered_curve_loop(planar_curves, tol)
            if loop is None:
                skipped_loops += 1
                continue

            room_loop_candidates.append((loop, planar_curves))

        if not room_loop_candidates:
            continue

        room_loop_candidates.sort(key=lambda data: _curve_loop_area_xy(data[0]), reverse=True)
        selected_loop, selected_curves = room_loop_candidates[0]
        all_curves.extend(selected_curves)
        room_loop_records.append({
            "room_item": item,
            "loop": selected_loop,
            "curves": selected_curves,
        })
        if len(room_loop_candidates) > 1:
            skipped_loops += len(room_loop_candidates) - 1

    return all_curves, room_loop_records, skipped_loops


def _collect_center_boundary_curves(room_items, target_elevation):
    opts = SpatialElementBoundaryOptions()
    try:
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center
    except Exception:
        pass

    all_curves = []

    for item in room_items:
        room = item.room
        link_inst = item.link_inst
        try:
            seglists = room.GetBoundarySegments(opts)
        except Exception:
            seglists = None

        if not seglists:
            continue

        for seglist in seglists:
            for seg in seglist:
                curve = None
                try:
                    curve = seg.GetCurve()
                except Exception:
                    try:
                        curve = seg.Curve
                    except Exception:
                        curve = None

                if curve is None:
                    continue

                host_curve = curve
                if link_inst is not None:
                    try:
                        trf = link_inst.GetTotalTransform() if hasattr(link_inst, "GetTotalTransform") else link_inst.GetTransform()
                        host_curve = curve.CreateTransformed(trf)
                    except Exception:
                        continue

                planar_segments = _curve_to_planar_segments(host_curve, target_elevation)
                if planar_segments:
                    all_curves.extend(planar_segments)

    return all_curves


def _iter_column_elements(target_doc):
    if target_doc is None:
        return []

    categories = [BuiltInCategory.OST_StructuralColumns]
    maybe_arch_col = getattr(BuiltInCategory, "OST_Columns", None)
    if maybe_arch_col is not None:
        categories.append(maybe_arch_col)

    seen = set()
    results = []
    for bic in categories:
        try:
            elems = (
                FilteredElementCollector(target_doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception:
            elems = []

        for elem in elems:
            try:
                key = elem.Id.IntegerValue
            except Exception:
                continue
            if key in seen:
                continue
            seen.add(key)
            results.append(elem)

    return results


def _get_element_probe_point(elem, view=None):
    if elem is None:
        return None

    try:
        loc = elem.Location
        if loc is not None and hasattr(loc, "Point") and loc.Point is not None:
            return loc.Point
    except Exception:
        pass

    try:
        bb = elem.get_BoundingBox(view)
        if bb is None:
            bb = elem.get_BoundingBox(None)
        if bb is None:
            return None
        return XYZ(
            (bb.Min.X + bb.Max.X) * 0.5,
            (bb.Min.Y + bb.Max.Y) * 0.5,
            (bb.Min.Z + bb.Max.Z) * 0.5,
        )
    except Exception:
        return None


def _curve_loop_from_4_points(p0, p1, p2, p3):
    loop = CurveLoop()
    loop.Append(Line.CreateBound(p0, p1))
    loop.Append(Line.CreateBound(p1, p2))
    loop.Append(Line.CreateBound(p2, p3))
    loop.Append(Line.CreateBound(p3, p0))
    return loop


def _polygon_area_2d(points):
    if not points or len(points) < 3:
        return 0.0
    area2 = 0.0
    count = len(points)
    for i in range(count):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % count]
        area2 += (x1 * y2) - (x2 * y1)
    return abs(area2) * 0.5


def _build_column_loop_from_bbox_host(bb, target_elevation):
    if bb is None:
        return None, None

    p0 = XYZ(bb.Min.X, bb.Min.Y, target_elevation)
    p1 = XYZ(bb.Max.X, bb.Min.Y, target_elevation)
    p2 = XYZ(bb.Max.X, bb.Max.Y, target_elevation)
    p3 = XYZ(bb.Min.X, bb.Max.Y, target_elevation)

    polygon = [(p0.X, p0.Y), (p1.X, p1.Y), (p2.X, p2.Y), (p3.X, p3.Y)]
    if _polygon_area_2d(polygon) <= 1e-6:
        return None, None

    center = XYZ((p0.X + p2.X) * 0.5, (p0.Y + p2.Y) * 0.5, target_elevation)
    return _curve_loop_from_4_points(p0, p1, p2, p3), center


def _build_column_loop_from_bbox_link(bb, link_inst, target_elevation):
    if bb is None or link_inst is None:
        return None, None

    try:
        trf = link_inst.GetTotalTransform() if hasattr(link_inst, "GetTotalTransform") else link_inst.GetTransform()
    except Exception:
        return None, None

    try:
        q0 = trf.OfPoint(XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z))
        q1 = trf.OfPoint(XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z))
        q2 = trf.OfPoint(XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z))
        q3 = trf.OfPoint(XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z))
    except Exception:
        return None, None

    p0 = XYZ(q0.X, q0.Y, target_elevation)
    p1 = XYZ(q1.X, q1.Y, target_elevation)
    p2 = XYZ(q2.X, q2.Y, target_elevation)
    p3 = XYZ(q3.X, q3.Y, target_elevation)

    polygon = [(p0.X, p0.Y), (p1.X, p1.Y), (p2.X, p2.Y), (p3.X, p3.Y)]
    if _polygon_area_2d(polygon) <= 1e-6:
        return None, None

    center = XYZ((p0.X + p2.X) * 0.5, (p0.Y + p2.Y) * 0.5, target_elevation)
    return _curve_loop_from_4_points(p0, p1, p2, p3), center


def _room_item_contains_host_point(room_item, host_point):
    if room_item is None or host_point is None:
        return False

    test_point = host_point
    if room_item.link_inst is not None:
        try:
            trf = room_item.link_inst.GetTotalTransform() if hasattr(room_item.link_inst, "GetTotalTransform") else room_item.link_inst.GetTransform()
            test_point = trf.Inverse.OfPoint(host_point)
        except Exception:
            return False

    try:
        return room_item.room.IsPointInRoom(test_point)
    except Exception:
        return False


def _room_items_contain_host_point(room_items, host_point):
    if host_point is None:
        return False

    for room_item in room_items:
        if _room_item_contains_host_point(room_item, host_point):
            return True

    return False


def _collect_column_hole_loops(room_items, target_elevation, view):
    if not room_items:
        return []
    hole_loops = []

    # Host columns in current model/view.
    for col in _iter_column_elements(doc):
        point_host = _get_element_probe_point(col, view)
        if not _room_items_contain_host_point(room_items, point_host):
            continue

        try:
            bb = col.get_BoundingBox(view)
            if bb is None:
                bb = col.get_BoundingBox(None)
        except Exception:
            bb = None

        loop, center = _build_column_loop_from_bbox_host(bb, target_elevation)
        if loop is not None and center is not None:
            hole_loops.append((loop, center))

    # Linked columns from all loaded links (including links different from room links)
    link_instances = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements())
    for link_inst in link_instances:
        link_doc = None
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue

        for col in _iter_column_elements(link_doc):
            try:
                bb = col.get_BoundingBox(None)
            except Exception:
                bb = None

            loop, center = _build_column_loop_from_bbox_link(bb, link_inst, target_elevation)
            if not _room_items_contain_host_point(room_items, center):
                continue
            if loop is not None and center is not None:
                hole_loops.append((loop, center))

    return hole_loops


def _curve_loop_polygon_xy(curve_loop):
    points = []
    try:
        for curve in curve_loop:
            p0 = curve.GetEndPoint(0)
            points.append((float(p0.X), float(p0.Y)))
    except Exception:
        return []
    return points


def _point_in_polygon_2d(px, py, polygon):
    if not polygon or len(polygon) < 3:
        return False

    inside = False
    j = len(polygon) - 1
    for i in range(len(polygon)):
        xi, yi = polygon[i]
        xj, yj = polygon[j]

        intersects = ((yi > py) != (yj > py))
        if intersects:
            denom = (yj - yi)
            if abs(denom) > 1e-12:
                xint = (xj - xi) * (py - yi) / denom + xi
                if px < xint:
                    inside = not inside
        j = i

    return inside


def _assign_column_holes_to_outer_loops(outer_loops, column_hole_loops):
    hole_map = {}
    outer_polys = []
    for idx, loop in enumerate(outer_loops):
        hole_map[idx] = []
        outer_polys.append(_curve_loop_polygon_xy(loop))

    for hole_loop, center in column_hole_loops:
        try:
            cx = float(center.X)
            cy = float(center.Y)
        except Exception:
            continue

        for idx, poly in enumerate(outer_polys):
            if _point_in_polygon_2d(cx, cy, poly):
                hole_map[idx].append(hole_loop)
                break

    return hole_map


def _remove_shared_interior_segments(curves):
    fragmented = _fragment_line_segments_at_overlaps(curves)
    buckets = {}
    for curve in fragmented:
        key = _curve_key_undirected(curve)
        if key is None:
            continue
        if key not in buckets:
            buckets[key] = []
        buckets[key].append(curve)

    result = []
    for key, items in buckets.items():
        count = len(items)
        if count % 2 == 1:
            result.append(items[0])
    return result


def _cross_2d(ax, ay, bx, by):
    return (ax * by) - (ay * bx)


def _segment_points_xy(curve):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return p0, p1
    except Exception:
        return None, None


def _point_on_segment_2d(point, seg_start, seg_end, tol=0.01):
    if point is None or seg_start is None or seg_end is None:
        return False

    vx = float(seg_end.X) - float(seg_start.X)
    vy = float(seg_end.Y) - float(seg_start.Y)
    wx = float(point.X) - float(seg_start.X)
    wy = float(point.Y) - float(seg_start.Y)
    seg_len = math.sqrt((vx * vx) + (vy * vy))
    if seg_len <= 1e-9:
        return False

    cross_val = abs(_cross_2d(vx, vy, wx, wy))
    if cross_val > (tol * max(seg_len, 1.0)):
        return False

    dot_val = (wx * vx) + (wy * vy)
    if dot_val < -(tol * seg_len):
        return False
    if dot_val > ((seg_len * seg_len) + (tol * seg_len)):
        return False

    return True


def _segment_param_2d(point, seg_start, seg_end):
    dx = float(seg_end.X) - float(seg_start.X)
    dy = float(seg_end.Y) - float(seg_start.Y)
    if abs(dx) >= abs(dy) and abs(dx) > 1e-9:
        return (float(point.X) - float(seg_start.X)) / dx
    if abs(dy) > 1e-9:
        return (float(point.Y) - float(seg_start.Y)) / dy
    return 0.0


def _segment_point_at_param(seg_start, seg_end, t):
    return XYZ(
        float(seg_start.X) + ((float(seg_end.X) - float(seg_start.X)) * t),
        float(seg_start.Y) + ((float(seg_end.Y) - float(seg_start.Y)) * t),
        float(seg_start.Z) + ((float(seg_end.Z) - float(seg_start.Z)) * t),
    )


def _segments_collinear_2d(start_a, end_a, start_b, end_b, tol=0.01):
    ax = float(end_a.X) - float(start_a.X)
    ay = float(end_a.Y) - float(start_a.Y)
    bx = float(end_b.X) - float(start_b.X)
    by = float(end_b.Y) - float(start_b.Y)
    seg_len_a = math.sqrt((ax * ax) + (ay * ay))
    seg_len_b = math.sqrt((bx * bx) + (by * by))
    if seg_len_a <= 1e-9 or seg_len_b <= 1e-9:
        return False

    if abs(_cross_2d(ax, ay, bx, by)) > (tol * max(seg_len_a, seg_len_b, 1.0)):
        return False

    off_x = float(start_b.X) - float(start_a.X)
    off_y = float(start_b.Y) - float(start_a.Y)
    return abs(_cross_2d(ax, ay, off_x, off_y)) <= (tol * max(seg_len_a, 1.0))


def _unique_sorted_params(params, tol=1e-6):
    sorted_params = sorted(params)
    unique = []
    for value in sorted_params:
        if not unique or abs(value - unique[-1]) > tol:
            unique.append(value)
    return unique


def _fragment_line_segments_at_overlaps(curves, tol=0.01):
    segment_data = []
    for curve in curves:
        p0, p1 = _segment_points_xy(curve)
        if p0 is None or p1 is None:
            continue
        if _distance(p0, p1) <= tol:
            continue
        segment_data.append((curve, p0, p1))

    fragments = []
    for idx, (_curve, start_pt, end_pt) in enumerate(segment_data):
        split_params = [0.0, 1.0]
        for other_idx, (_other_curve, other_start, other_end) in enumerate(segment_data):
            if idx == other_idx:
                continue
            if not _segments_collinear_2d(start_pt, end_pt, other_start, other_end, tol):
                continue

            for point in (other_start, other_end):
                if not _point_on_segment_2d(point, start_pt, end_pt, tol):
                    continue
                param = _segment_param_2d(point, start_pt, end_pt)
                if tol < param < (1.0 - tol):
                    split_params.append(param)

        split_params = _unique_sorted_params(split_params)
        for param_index in range(len(split_params) - 1):
            t0 = split_params[param_index]
            t1 = split_params[param_index + 1]
            if (t1 - t0) <= 1e-9:
                continue

            frag_start = _segment_point_at_param(start_pt, end_pt, t0)
            frag_end = _segment_point_at_param(start_pt, end_pt, t1)
            if _distance(frag_start, frag_end) <= tol:
                continue

            try:
                fragments.append(Line.CreateBound(frag_start, frag_end))
            except Exception:
                continue

    return fragments


def _curve_direction_xy(curve):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return (float(p1.X) - float(p0.X), float(p1.Y) - float(p0.Y))
    except Exception:
        return (0.0, 0.0)


def _vector_length_xy(vec):
    return math.sqrt((vec[0] * vec[0]) + (vec[1] * vec[1]))


def _score_curve_continuation(prev_curve, next_curve):
    prev_vec = _curve_direction_xy(prev_curve)
    next_vec = _curve_direction_xy(next_curve)
    prev_len = _vector_length_xy(prev_vec)
    next_len = _vector_length_xy(next_vec)

    if prev_len <= 1e-9 or next_len <= 1e-9:
        return -999999.0

    return ((prev_vec[0] * next_vec[0]) + (prev_vec[1] * next_vec[1])) / (prev_len * next_len)


def _curve_loop_area_xy(curve_loop):
    return _polygon_area_2d(_curve_loop_polygon_xy(curve_loop))


def _curve_loop_signed_area_xy(curve_loop):
    points = _curve_loop_polygon_xy(curve_loop)
    if not points or len(points) < 3:
        return 0.0

    area2 = 0.0
    count = len(points)
    for idx in range(count):
        x1, y1 = points[idx]
        x2, y2 = points[(idx + 1) % count]
        area2 += (x1 * y2) - (x2 * y1)
    return area2 * 0.5


def _reverse_curve_loop(curve_loop):
    reversed_curves = []
    try:
        curves = list(curve_loop)
    except Exception:
        return curve_loop

    for curve in reversed(curves):
        try:
            reversed_curves.append(curve.CreateReversed())
        except Exception:
            return curve_loop

    reversed_loop = CurveLoop()
    try:
        for curve in reversed_curves:
            reversed_loop.Append(curve)
        return reversed_loop
    except Exception:
        return curve_loop


def _normalize_curve_loop_orientation(curve_loop, clockwise=False):
    signed_area = _curve_loop_signed_area_xy(curve_loop)
    if abs(signed_area) <= 1e-9:
        return curve_loop

    is_clockwise = signed_area < 0.0
    if is_clockwise == clockwise:
        return curve_loop
    return _reverse_curve_loop(curve_loop)


def _group_loops_into_boundary_sets(loops):
    if not loops:
        return []

    loop_infos = []
    for idx, loop in enumerate(loops):
        polygon = _curve_loop_polygon_xy(loop)
        if len(polygon) < 3:
            continue
        loop_infos.append({
            "index": idx,
            "loop": loop,
            "polygon": polygon,
            "area": _curve_loop_area_xy(loop),
            "parent": None,
            "depth": 0,
        })

    loop_infos.sort(key=lambda info: info["area"], reverse=True)

    for idx, info in enumerate(loop_infos):
        point = info["polygon"][0]
        parent_idx = None
        parent_area = None
        for candidate_idx, candidate in enumerate(loop_infos):
            if candidate_idx == idx:
                continue
            if candidate["area"] <= info["area"]:
                continue
            if _point_in_polygon_2d(point[0], point[1], candidate["polygon"]):
                if parent_idx is None or candidate["area"] < parent_area:
                    parent_idx = candidate_idx
                    parent_area = candidate["area"]

        info["parent"] = parent_idx

    for idx, info in enumerate(loop_infos):
        depth = 0
        parent_idx = info["parent"]
        while parent_idx is not None:
            depth += 1
            parent_idx = loop_infos[parent_idx]["parent"]
        info["depth"] = depth

    boundary_sets = []
    outer_map = {}
    for idx, info in enumerate(loop_infos):
        if (info["depth"] % 2) == 0:
            normalized_outer = _normalize_curve_loop_orientation(info["loop"], clockwise=False)
            outer_map[idx] = len(boundary_sets)
            boundary_sets.append([normalized_outer])

    for idx, info in enumerate(loop_infos):
        if (info["depth"] % 2) == 0:
            continue

        parent_idx = info["parent"]
        while parent_idx is not None and (loop_infos[parent_idx]["depth"] % 2) == 1:
            parent_idx = loop_infos[parent_idx]["parent"]

        if parent_idx is None:
            continue

        set_index = outer_map.get(parent_idx)
        if set_index is None:
            continue

        boundary_sets[set_index].append(_normalize_curve_loop_orientation(info["loop"], clockwise=True))

    return boundary_sets


def _try_trace_closed_loop(seed_curve, curves, used_indexes, endpoint_map, tol):
    start_pt = seed_curve.GetEndPoint(0)
    cursor = seed_curve.GetEndPoint(1)
    current_curve = seed_curve
    current_loop = [seed_curve]
    local_indexes = set()
    max_steps = max(20, len(curves) + 20)

    for idx, candidate in enumerate(curves):
        if candidate is seed_curve:
            local_indexes.add(idx)
            break

    for _ in range(max_steps):
        if _distance(cursor, start_pt) <= tol and len(current_loop) >= 3:
            loop = _curve_loop_from_curves(current_loop, tol)
            if loop is None:
                return None, None
            return loop, local_indexes

        cursor_key = _curve_point_key(cursor, tol)
        candidate_indexes = endpoint_map.get(cursor_key, [])
        best_index = None
        best_curve = None
        best_score = -999999.0

        for candidate_index in candidate_indexes:
            if candidate_index in used_indexes or candidate_index in local_indexes:
                continue
            candidate = curves[candidate_index]
            oriented = _orient_curve_to_start(candidate, cursor, tol)
            if oriented is None:
                continue

            score = _score_curve_continuation(current_curve, oriented)
            if best_index is None or score > best_score:
                best_index = candidate_index
                best_curve = oriented
                best_score = score

        if best_index is None or best_curve is None:
            return None, None

        current_loop.append(best_curve)
        local_indexes.add(best_index)
        current_curve = best_curve
        cursor = best_curve.GetEndPoint(1)

    return None, None


def _build_closed_loops(curves, tol=0.02):
    loops = []
    used_indexes = set()
    endpoint_map = {}

    for idx, curve in enumerate(curves):
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
        except Exception:
            continue

        for pt in (p0, p1):
            key = _curve_point_key(pt, tol)
            if key not in endpoint_map:
                endpoint_map[key] = []
            endpoint_map[key].append(idx)

    for idx, curve in enumerate(curves):
        if idx in used_indexes:
            continue

        loop, local_indexes = _try_trace_closed_loop(curve, curves, used_indexes, endpoint_map, tol)
        if loop is None:
            try:
                reversed_seed = curve.CreateReversed()
            except Exception:
                reversed_seed = None
            if reversed_seed is not None:
                loop, local_indexes = _try_trace_closed_loop(reversed_seed, curves, used_indexes, endpoint_map, tol)

        if loop is None or not local_indexes:
            continue

        if _curve_loop_area_xy(loop) <= 1e-6:
            continue

        used_indexes.update(local_indexes)
        loops.append(loop)

    loops.sort(key=_curve_loop_area_xy, reverse=True)
    return loops


def _group_connected_room_loop_records(room_loop_records, tol=0.01):
    if not room_loop_records:
        return []

    adjacency = {}
    key_to_indexes = {}
    segment_data = []
    for idx, record in enumerate(room_loop_records):
        adjacency[idx] = set()
        for curve in record.get("curves", []):
            key = _curve_key_undirected(curve, tol)
            if key is None:
                continue
            existing = key_to_indexes.get(key, [])
            for other_idx in existing:
                adjacency[idx].add(other_idx)
                adjacency[other_idx].add(idx)
            existing.append(idx)
            key_to_indexes[key] = existing

            p0, p1 = _segment_points_xy(curve)
            if p0 is None or p1 is None:
                continue
            if _distance(p0, p1) <= tol:
                continue
            segment_data.append((idx, p0, p1))

    fragment_owner_map = {}
    for seg_idx, (room_idx, start_pt, end_pt) in enumerate(segment_data):
        split_params = [0.0, 1.0]
        for other_seg_idx, (_other_room_idx, other_start, other_end) in enumerate(segment_data):
            if seg_idx == other_seg_idx:
                continue
            if not _segments_collinear_2d(start_pt, end_pt, other_start, other_end, tol):
                continue

            for point in (other_start, other_end):
                if not _point_on_segment_2d(point, start_pt, end_pt, tol):
                    continue
                param = _segment_param_2d(point, start_pt, end_pt)
                if tol < param < (1.0 - tol):
                    split_params.append(param)

        split_params = _unique_sorted_params(split_params)
        for param_index in range(len(split_params) - 1):
            t0 = split_params[param_index]
            t1 = split_params[param_index + 1]
            if (t1 - t0) <= 1e-9:
                continue

            frag_start = _segment_point_at_param(start_pt, end_pt, t0)
            frag_end = _segment_point_at_param(start_pt, end_pt, t1)
            if _distance(frag_start, frag_end) <= tol:
                continue

            try:
                fragment = Line.CreateBound(frag_start, frag_end)
            except Exception:
                continue

            key = _curve_key_undirected(fragment, tol)
            if key is None:
                continue

            owners = fragment_owner_map.get(key)
            if owners is None:
                owners = set()
                fragment_owner_map[key] = owners
            owners.add(room_idx)

    for owners in fragment_owner_map.values():
        if len(owners) < 2:
            continue

        owner_list = list(owners)
        for idx, room_idx in enumerate(owner_list):
            for other_room_idx in owner_list[idx + 1:]:
                adjacency[room_idx].add(other_room_idx)
                adjacency[other_room_idx].add(room_idx)

    groups = []
    visited = set()
    for idx in range(len(room_loop_records)):
        if idx in visited:
            continue

        stack = [idx]
        component = []
        visited.add(idx)
        while stack:
            current = stack.pop()
            component.append(room_loop_records[current])
            for neighbor in adjacency.get(current, set()):
                if neighbor in visited:
                    continue
                visited.add(neighbor)
                stack.append(neighbor)

        groups.append(component)

    return groups


def _merge_connected_room_loops(room_loop_records, tol=0.02):
    merged_curves = []
    merged_loops = []

    for group in _group_connected_room_loop_records(room_loop_records, tol):
        if len(group) == 1:
            merged_curves.extend(group[0].get("curves", []))
            merged_loops.append(group[0].get("loop"))
            continue

        group_curves = []
        for record in group:
            group_curves.extend(record.get("curves", []))

        outer_group_curves = _remove_shared_interior_segments(group_curves)
        group_loops = _build_closed_loops(outer_group_curves, tol)
        if group_loops:
            merged_curves.extend(outer_group_curves)
            merged_loops.extend(group_loops)
            continue

        for record in group:
            merged_curves.extend(record.get("curves", []))
            merged_loops.append(record.get("loop"))

    merged_loops = [loop for loop in merged_loops if loop is not None]
    merged_loops.sort(key=_curve_loop_area_xy, reverse=True)
    return merged_curves, merged_loops


def _resolve_region_type(base_type, create_new_type, new_type_name):
    if not create_new_type:
        return base_type

    if not new_type_name:
        return base_type

    existing = None
    region_types = FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements()
    for fr_type in region_types:
        try:
            if _safe_elem_name(fr_type, "") == new_type_name:
                existing = fr_type
                break
        except Exception:
            continue

    if existing is not None:
        return existing

    try:
        duplicate_result = base_type.Duplicate(new_type_name)
        if isinstance(duplicate_result, ElementId):
            return doc.GetElement(duplicate_result)
        if isinstance(duplicate_result, FilledRegionType):
            return duplicate_result
        try:
            return doc.GetElement(duplicate_result)
        except Exception:
            return duplicate_result
    except Exception as ex:
        forms.alert("Failed to create Filled Region Type: {0}".format(ex), title="Linked Room Region")
        return None


def _get_filled_region_line_style_id(region):
    if region is None:
        return None

    try:
        line_style_id = region.GetLineStyleId()
        if line_style_id is not None and line_style_id != ElementId.InvalidElementId:
            return line_style_id
    except Exception:
        pass

    try:
        line_style_id = getattr(region, "LineStyleId", None)
        if line_style_id is not None and line_style_id != ElementId.InvalidElementId:
            return line_style_id
    except Exception:
        pass

    return None


def _pick_rooms_for_state(state):
    new_host_ids = set()
    new_link_pairs = set()
    prefer_linked_pick = any(FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements())

    try:
        if prefer_linked_pick:
            refs = uidoc.Selection.PickObjects(
                ObjectType.LinkedElement,
                LinkedRoomSelectionFilter(doc),
                "Pick linked room(s) in active view"
            )
            if refs:
                for rf in refs:
                    try:
                        new_link_pairs.add((rf.ElementId.IntegerValue, rf.LinkedElementId.IntegerValue))
                    except Exception:
                        continue
        else:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                RoomSelectionFilter(),
                "Pick host room(s) in active view"
            )
            if refs:
                for rf in refs:
                    try:
                        new_host_ids.add(rf.ElementId.IntegerValue)
                    except Exception:
                        continue
    except Exception as ex:
        if "cancel" not in ex.__class__.__name__.lower():
            logger.debug("Room selection failed: {0}".format(ex))
        return state

    if new_host_ids or new_link_pairs:
        state["selected_host_room_ids"] = sorted(new_host_ids)
        state["selected_link_room_pairs"] = sorted(list(new_link_pairs))
        state["room_filter"] = ""
    return state


def _pick_existing_regions_for_state(state):
    picked_ids = []
    line_style_id = None
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            FilledRegionSelectionFilter(doc.ActiveView.Id),
            "Pick filled region(s) in the active view to replace"
        )
        for rf in refs or []:
            try:
                region = doc.GetElement(rf.ElementId)
                if region is None:
                    continue
                picked_ids.append(region.Id)
                if line_style_id is None:
                    line_style_id = _get_filled_region_line_style_id(region)
            except Exception:
                continue
    except Exception as ex:
        if "cancel" not in ex.__class__.__name__.lower():
            logger.debug("Filled region selection failed: {0}".format(ex))
        return state

    unique_ids = []
    seen = set()
    for region_id in picked_ids:
        try:
            key = region_id.IntegerValue
        except Exception:
            continue
        if key in seen:
            continue
        seen.add(key)
        unique_ids.append(region_id)

    if unique_ids:
        state["existing_region_ids"] = _serialize_element_id_list(unique_ids)
        state["existing_line_style_id"] = _serialize_element_id(line_style_id)
    return state


def _pick_color_for_state(state, state_key, title):
    current_rgb = tuple(state.get(state_key, (17, 119, 187)))
    try:
        dlg = ColorDialog()
        try:
            dlg.FullOpen = True
        except Exception:
            pass

        try:
            dlg.Color = DrawingColor.FromArgb(int(current_rgb[0]), int(current_rgb[1]), int(current_rgb[2]))
        except Exception:
            pass

        res = dlg.ShowDialog()
        if res == DialogResult.OK:
            picked = dlg.Color
            state[state_key] = (int(picked.R), int(picked.G), int(picked.B))
            return state
        return state
    except Exception as ex:
        logger.debug("Color dialog failed: {0}".format(ex))

    color_value = forms.ask_for_string(
        default=_rgb_to_hex(current_rgb),
        prompt="Enter color as #RRGGBB or R,G,B.",
        title=title,
    )
    if not color_value:
        return state

    parsed_rgb = _parse_rgb_text(color_value)
    if parsed_rgb is None:
        forms.alert("Invalid color value. Use #RRGGBB or R,G,B.", title="Linked Room Region")
        return state

    state[state_key] = parsed_rgb
    return state


def _create_filled_regions(view, region_type, loops, hole_map=None, line_style_id=None):
    created = 0
    failed = 0
    errors = []
    created_region_ids = []

    hole_map = hole_map or {}

    if loops and not hole_map:
        boundary_sets = _group_loops_into_boundary_sets(loops)
        if not boundary_sets:
            return 0, 1, [{
                "loop_index": 0,
                "boundary_count": 0,
                "hole_count": 0,
                "message": "No valid boundary sets were resolved from the merged room loops.",
            }], created_region_ids

        for set_index, boundary_set in enumerate(boundary_sets):
            ok_set, err_set = _try_create_region_boundaries(view, region_type, boundary_set, line_style_id)
            if ok_set:
                try:
                    region = _commit_region_with_boundaries(view, region_type, boundary_set, line_style_id)
                    created += 1
                    if region is not None:
                        created_region_ids.append(region.Id)
                except Exception as ex:
                    failed += 1
                    errors.append({
                        "loop_index": set_index,
                        "boundary_count": len(boundary_set),
                        "hole_count": max(0, len(boundary_set) - 1),
                        "message": str(ex),
                    })
                continue

            failed += 1
            errors.append({
                "loop_index": set_index,
                "boundary_count": len(boundary_set),
                "hole_count": max(0, len(boundary_set) - 1),
                "message": "Boundary set create failed: {0}".format(err_set),
            })

        return created, failed, errors, created_region_ids

    for idx, loop in enumerate(loops):
        outer_only = [loop]
        candidate_holes = list(hole_map.get(idx, []))
        all_boundaries = outer_only + candidate_holes

        ok_all, err_all = _try_create_region_boundaries(view, region_type, all_boundaries, line_style_id)
        if ok_all:
            try:
                region = _commit_region_with_boundaries(view, region_type, all_boundaries, line_style_id)
                created += 1
                if region is not None:
                    created_region_ids.append(region.Id)
            except Exception as ex:
                failed += 1
                errors.append({
                    "loop_index": idx,
                    "boundary_count": len(all_boundaries),
                    "hole_count": len(candidate_holes),
                    "message": str(ex),
                })
            continue

        ok_outer, err_outer = _try_create_region_boundaries(view, region_type, outer_only, line_style_id)
        if not ok_outer:
            failed += 1
            logger.debug("Filled region outer boundary failed: {0}".format(err_outer))
            errors.append({
                "loop_index": idx,
                "boundary_count": len(outer_only),
                "hole_count": 0,
                "message": str(err_outer),
            })
            continue

        accepted_holes = []
        rejected_holes = []
        for hole_index, hole_loop in enumerate(candidate_holes):
            test_boundaries = outer_only + accepted_holes + [hole_loop]
            ok_hole, err_hole = _try_create_region_boundaries(view, region_type, test_boundaries, line_style_id)
            if ok_hole:
                accepted_holes.append(hole_loop)
            else:
                rejected_holes.append({
                    "hole_index": hole_index,
                    "message": str(err_hole),
                })

        final_boundaries = outer_only + accepted_holes
        try:
            region = _commit_region_with_boundaries(view, region_type, final_boundaries, line_style_id)
            created += 1
            if region is not None:
                created_region_ids.append(region.Id)
            if rejected_holes or err_all is not None:
                errors.append({
                    "loop_index": idx,
                    "boundary_count": len(final_boundaries),
                    "hole_count": len(accepted_holes),
                    "message": "Recovered by excluding {0} invalid hole loop(s). Initial error: {1}".format(
                        len(rejected_holes),
                        err_all,
                    ),
                    "rejected_holes": rejected_holes,
                })
        except Exception as ex:
            failed += 1
            logger.debug("Filled region final create failed: {0}".format(ex))
            errors.append({
                "loop_index": idx,
                "boundary_count": len(final_boundaries),
                "hole_count": len(accepted_holes),
                "message": str(ex),
                "rejected_holes": rejected_holes,
            })

    return created, failed, errors, created_region_ids


def _curve_loop_list(curve_loops):
    boundaries = List[CurveLoop]()
    for curve_loop in curve_loops:
        boundaries.Add(curve_loop)
    return boundaries


def _try_create_region_boundaries(view, region_type, curve_loops, line_style_id=None):
    sub_tx = SubTransaction(doc)
    try:
        sub_tx.Start()
        region = FilledRegion.Create(doc, region_type.Id, view.Id, _curve_loop_list(curve_loops))
        if line_style_id is not None:
            try:
                region.SetLineStyleId(line_style_id)
            except Exception:
                pass
        sub_tx.RollBack()
        return True, None
    except Exception as ex:
        try:
            sub_tx.RollBack()
        except Exception:
            pass
        return False, ex


def _commit_region_with_boundaries(view, region_type, curve_loops, line_style_id=None):
    region = FilledRegion.Create(doc, region_type.Id, view.Id, _curve_loop_list(curve_loops))
    if line_style_id is not None:
        try:
            region.SetLineStyleId(line_style_id)
        except Exception:
            pass
    return region


def _build_operation_log_lines(view, data, region_type, curves, outer_curves, loops, column_hole_loops, created, failed, errors, skipped_loops=0):
    resolved_fg_pattern_id, resolved_bg_pattern_id, resolved_show_foreground, resolved_show_background = _resolve_visible_region_graphics(
        data.get("foreground_pattern_id"),
        data.get("background_pattern_id"),
        data.get("show_foreground", True),
        data.get("show_background", True),
        data.get("boundary_line_style_id"),
    )
    lines = [
        "Linked Room Region Debug",
        "Geometry mode: planar projected boundary segments",
        "Boundary mode: connected selected room perimeters merged (column holes disabled)",
        "View: {0}".format(getattr(view, "Name", "Active View")),
        "Selected rooms: {0}".format(len(data.get("room_items", []))),
        "Boundary curves collected: {0}".format(len(curves)),
        "Boundary segments used: {0}".format(len(outer_curves)),
        "Closed loops built: {0}".format(len(loops)),
        "Secondary/invalid loops skipped: {0}".format(skipped_loops),
        "Column hole loops: disabled",
        "Filled region type: {0}".format(_safe_elem_name(region_type, "Filled Region Type") if region_type else "None"),
        "Masking region type: {0}".format(bool(data.get("is_masking", False))),
        "Boundary line style: {0}".format(_line_style_name_from_id(data.get("boundary_line_style_id"))),
        "Show foreground pattern: {0}".format(resolved_show_foreground),
        "Show background pattern: {0}".format(resolved_show_background),
        "Replace existing mode: {0}".format(bool(data.get("replace_existing", False))),
        "Result: created {0} | failed {1}".format(created, failed),
        "",
        "FilledRegion.Create Errors",
    ]

    if not errors:
        lines.append("No create errors were recorded.")
        return lines

    for err in errors:
        lines.append(
            "Loop {0} | Boundaries {1} | Holes {2} | Error {3}".format(
                err.get("loop_index", 0) + 1,
                err.get("boundary_count", 0),
                err.get("hole_count", 0),
                err.get("message", ""),
            )
        )
        for rejected in err.get("rejected_holes", []):
            lines.append(
                "  Rejected hole {0}: {1}".format(
                    rejected.get("hole_index", 0) + 1,
                    rejected.get("message", ""),
                )
            )

    return lines


def _write_operation_log_file(lines):
    temp_dir = os.environ.get("TEMP") or os.environ.get("TMP") or script.get_bundle_file(".")
    log_path = os.path.join(temp_dir, "LinkedRoomRegionDebug.txt")
    try:
        with open(log_path, "w") as log_file:
            log_file.write("\n".join(lines))
        return log_path
    except Exception as ex:
        logger.debug("Failed to write Linked Room Region log file: {0}".format(ex))
        return None


def _write_operation_output(view, data, region_type, curves, outer_curves, loops, column_hole_loops, hole_map, created, failed, errors, skipped_loops=0):
    log_lines = _build_operation_log_lines(view, data, region_type, curves, outer_curves, loops, column_hole_loops, created, failed, errors, skipped_loops)
    should_write_debug_log = bool(errors) or failed > 0
    log_path = _write_operation_log_file(log_lines) if should_write_debug_log else None
    first_error = errors[0].get("message", "") if errors else ""

    if should_write_debug_log:
        for line in log_lines:
            logger.debug(line)

    return {
        "log_path": log_path,
        "first_error": first_error,
    }


def _execute_room_region_operation(view, data):
    if view is None:
        forms.alert("No active view found.", title="Linked Room Region")
        return False

    if not data:
        return False

    target_elev = _active_view_plane_elevation(view)
    room_items = data["room_items"]

    curves, room_loop_records, skipped_loops = _collect_selected_room_loops(room_items, target_elev)
    if not curves:
        forms.alert("No room boundary curves were found for selected rooms.", title="Linked Room Region")
        return False
    if not room_loop_records:
        forms.alert("Failed to build closed boundaries from selected rooms.", title="Linked Room Region")
        return False

    outer_curves, loops = _merge_connected_room_loops(room_loop_records)
    if not loops:
        forms.alert("Failed to build closed boundaries from selected rooms.", title="Linked Room Region")
        return False

    column_hole_loops = []
    hole_map = {}
    replace_existing = bool(data.get("replace_existing", False))
    existing_region_ids = data.get("existing_region_ids", []) or []
    existing_line_style_id = data.get("existing_line_style_id")
    boundary_line_style_id = data.get("boundary_line_style_id") or existing_line_style_id or _invisible_line_style_id()
    resolved_fg_pattern_id, resolved_bg_pattern_id, resolved_show_foreground, resolved_show_background = _resolve_visible_region_graphics(
        data.get("foreground_pattern_id"),
        data.get("background_pattern_id"),
        data.get("show_foreground", True),
        data.get("show_background", True),
        boundary_line_style_id,
    )
    created_region_ids = []

    t = Transaction(doc, "Create or Update Filled Regions from Linked Rooms")
    try:
        t.Start()
        region_type = _resolve_region_type(
            data["base_type"],
            data["create_new_type"],
            data["new_type_name"],
        )
        if region_type is None:
            t.RollBack()
            return False

        _set_region_type_is_masking(region_type, data.get("is_masking", False))

        if data.get("apply_colors", False):
            _apply_region_type_colors(
                region_type,
                data.get("foreground_rgb", (17, 119, 187)),
                data.get("background_rgb", (42, 42, 46)),
                resolved_fg_pattern_id,
                resolved_bg_pattern_id,
                resolved_show_foreground,
                resolved_show_background,
            )

        if replace_existing and existing_region_ids:
            try:
                doc.Delete(List[ElementId](existing_region_ids))
            except Exception:
                for region_id in existing_region_ids:
                    try:
                        doc.Delete(region_id)
                    except Exception:
                        pass

        created, failed, errors, created_region_ids = _create_filled_regions(view, region_type, loops, hole_map, boundary_line_style_id)
        debug_result = _write_operation_output(view, data, region_type, curves, outer_curves, loops, column_hole_loops, hole_map, created, failed, errors, skipped_loops)

        if created == 0:
            t.RollBack()
            first_error = (debug_result or {}).get("first_error", "")
            log_path = (debug_result or {}).get("log_path", None)
            message_lines = ["No filled regions were created."]
            if first_error:
                message_lines.append("")
                message_lines.append("First Revit error:")
                message_lines.append(first_error)
            if log_path:
                message_lines.append("")
                message_lines.append("Debug log file:")
                message_lines.append(log_path)
            if errors:
                forms.alert("\n".join(message_lines), title="Linked Room Region")
            else:
                forms.alert("\n".join(message_lines), title="Linked Room Region")
            return False

        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        forms.alert("Failed while creating regions: {0}".format(ex), title="Linked Room Region")
        return False

    if created_region_ids:
        try:
            region_ids = List[ElementId](created_region_ids)
            uidoc.Selection.SetElementIds(region_ids)
            try:
                uidoc.ShowElements(region_ids)
            except Exception:
                pass
        except Exception as ex:
            logger.debug("Failed to select created filled regions: {0}".format(ex))

    forms.alert(
        "\n".join([
            "{0} {1} filled region(s). Failed: {2}.".format(
                "Updated" if replace_existing else "Created",
                created,
                failed,
            ),
            "Boundary groups detected: {0}.".format(len(loops)),
            "Column boundaries detected: 0.",
            "Debug log file: {0}".format(debug_result.get("log_path")) if failed > 0 and debug_result and debug_result.get("log_path") else "",
        ]).strip(),
        title="Linked Room Region",
    )
    return True


def main():
    view = doc.ActiveView
    if view is None:
        forms.alert("No active view found.", title="Linked Room Region")
        return

    current_view_id = _serialize_element_id(view.Id)
    persisted_state = _load_persisted_success_state(current_view_id)

    xaml_file = script.get_bundle_file("LinkedRoomRegionWindow.xaml")
    state = {
        "view_id": current_view_id,
        "foreground_rgb": (17, 119, 187),
        "background_rgb": (42, 42, 46),
        "show_foreground": True,
        "show_background": True,
        "boundary_line_style_id": _serialize_element_id(_invisible_line_style_id()),
        "selected_host_room_ids": [],
        "selected_link_room_pairs": [],
        "existing_region_ids": [],
    }
    state.update(persisted_state)
    state["view_id"] = current_view_id

    while True:
        active_view = doc.ActiveView
        if active_view is None:
            forms.alert("No active view found.", title="Linked Room Region")
            return

        if state.get("view_id") and _serialize_element_id(active_view.Id) != state.get("view_id"):
            forms.alert("Active view changed while using Linked Room Region. Reopen the tool from the target view.", title="Linked Room Region")
            return

        win = None
        try:
            win = LinkedRoomRegionWindow(xaml_file, state)
            win.ShowDialog()
        except Exception as ex:
            forms.alert("Failed to open Linked Room Region window: {0}".format(ex), title="Linked Room Region")
            return

        if win is None or not win.result:
            return

        result = win.result
        action = result.get("action")
        state = result.get("state", state)

        if action == "pick_rooms":
            state = _pick_rooms_for_state(state)
            continue

        if action == "pick_existing_regions":
            state = _pick_existing_regions_for_state(state)
            continue

        if action == "pick_foreground_color":
            state = _pick_color_for_state(state, "foreground_rgb", "Set Foreground Color")
            continue

        if action == "pick_background_color":
            state = _pick_color_for_state(state, "background_rgb", "Set Background Color")
            continue

        if action == "create":
            if _execute_room_region_operation(active_view, result.get("data")):
                _save_persisted_success_state(result.get("state", state))
                return
            continue

        return


if __name__ == "__main__":
    main()
