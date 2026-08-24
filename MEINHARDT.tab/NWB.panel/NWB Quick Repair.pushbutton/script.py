# coding: utf8
from __future__ import print_function

import math
import re

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    RevitLinkInstance,
    StorageType,
    Transaction,
    TransactionStatus,
    XYZ,
)
from pyrevit import forms, revit


doc = revit.doc
uidoc = revit.uidoc

TOOL_TITLE = "NWB Quick Repair"

TARGET_PARAMS = [
    "NWB_Department",
    "NWB_SubDepartment",
    "NWB_RoomID",
    "NWB_RoomName",
]

MEP_CATEGORIES = [
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeAccessory,
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

_LINKED_ROOM_CACHE = None

ROOM_PARAM_ALIASES = {
    "NWB_RoomID": ["NWB_RoomID", "Room Number", "Number", "Room ID"],
    "NWB_RoomName": ["NWB_RoomName", "Room Name", "Name"],
    "NWB_Department": ["NWB_Department", "Department", "Room Department", "Department Name"],
    "NWB_SubDepartment": ["NWB_SubDepartment", "SubDepartment", "Subdepartment", "Room SubDepartment"],
}


def _safe_str(value):
    try:
        if value is None:
            return ""
        return str(value)
    except Exception:
        return ""


def _norm(value):
    text = _safe_str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def _is_blank(value):
    return _norm(value) == ""


def _find_param(element, name):
    if element is None or not name:
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
            if param is None:
                continue
            try:
                if param.Definition and param.Definition.Name == name:
                    return param
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

    try:
        type_el = element.Symbol
    except Exception:
        type_el = None
    if type_el is not None:
        try:
            owner_id = type_el.Id.IntegerValue
        except Exception:
            owner_id = id(type_el)
        if owner_id not in seen:
            owners.append(("type", type_el))

    return owners


def _read_param_text(param):
    if param is None:
        return ""

    try:
        storage = param.StorageType
    except Exception:
        storage = None

    if storage == StorageType.String:
        try:
            text = param.AsString()
            return _safe_str(text).strip()
        except Exception:
            pass

    if storage == StorageType.Integer:
        try:
            text = param.AsInteger()
            return _safe_str(text).strip()
        except Exception:
            pass

    if storage == StorageType.Double:
        try:
            text = param.AsValueString()
            return _safe_str(text).strip()
        except Exception:
            pass

    if storage == StorageType.ElementId:
        try:
            eid = param.AsElementId()
            if eid is not None and eid != ElementId.InvalidElementId:
                return _safe_str(eid.IntegerValue).strip()
        except Exception:
            pass

    try:
        value = param.AsValueString()
        if value is not None:
            return _safe_str(value).strip()
    except Exception:
        pass

    try:
        value = param.AsString()
        if value is not None:
            return _safe_str(value).strip()
    except Exception:
        pass

    return ""


def _get_room_value_from_room(room, param_name, source_doc):
    aliases = ROOM_PARAM_ALIASES.get(param_name, [param_name])
    for alias in aliases:
        param = _find_param(room, alias)
        text = _read_param_text(param)
        if text:
            return text

    type_owner = None
    try:
        type_id = room.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            type_owner = source_doc.GetElement(type_id)
    except Exception:
        type_owner = None

    if type_owner is not None:
        for alias in aliases:
            param = _find_param(type_owner, alias)
            text = _read_param_text(param)
            if text:
                return text
    if param_name == "NWB_RoomID":
        try:
            return _safe_str(room.Number).strip()
        except Exception:
            pass
    if param_name == "NWB_RoomName":
        try:
            return _safe_str(room.Name).strip()
        except Exception:
            pass
    return ""


def _element_center_point(element):
    if element is None:
        return None

    try:
        loc = element.Location
        if loc is not None and hasattr(loc, "Point"):
            return loc.Point
    except Exception:
        pass

    try:
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


def _find_room_for_point(point, link_data):
    if point is None or link_data is None:
        return None

    try:
        transformed = link_data["transform"].OfPoint(point)
    except Exception:
        transformed = point

    candidate_rooms = []
    try:
        cell_size = link_data["cell_size"]
        cell_x = int(math.floor(float(transformed.X) / cell_size))
        cell_y = int(math.floor(float(transformed.Y) / cell_size))
        seen = set()
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for room_item in link_data["grid"].get((cell_x + dx, cell_y + dy), []):
                    room_id = room_item["id"]
                    if room_id not in seen:
                        seen.add(room_id)
                        candidate_rooms.append(room_item)
        candidate_rooms.extend(link_data["fallback"])
    except Exception:
        candidate_rooms = link_data["rooms"]

    for room_item in candidate_rooms:
        room = room_item["room"]
        if room is None:
            continue
        bounds = room_item.get("bounds")
        if bounds is not None:
            if not (
                bounds[0] <= transformed.X <= bounds[3]
                and bounds[1] <= transformed.Y <= bounds[4]
                and bounds[2] <= transformed.Z <= bounds[5]
            ):
                continue
        try:
            if hasattr(room, "IsPointInRoom") and room.IsPointInRoom(transformed):
                return room
        except Exception:
            pass
        return room

    return None


def _get_linked_room_cache():
    global _LINKED_ROOM_CACHE
    if _LINKED_ROOM_CACHE is not None:
        return _LINKED_ROOM_CACHE

    linked_rooms = []
    try:
        links = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    except Exception:
        links = []

    for link in links:
        if link is None:
            continue
        try:
            link_doc = link.GetLinkDocument()
            transform = link.GetTotalTransform().Inverse
            rooms = list(FilteredElementCollector(link_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType())
        except Exception:
            continue
        if link_doc is None or not rooms:
            continue

        room_items = []
        grid = {}
        fallback = []
        cell_size = 20.0
        for room in rooms:
            try:
                bb = room.get_BoundingBox(None)
                bounds = None
                if bb is not None:
                    bounds = (bb.Min.X, bb.Min.Y, bb.Min.Z, bb.Max.X, bb.Max.Y, bb.Max.Z)
            except Exception:
                bounds = None

            room_item = {
                "room": room,
                "id": room.Id.IntegerValue,
                "bounds": bounds,
            }
            room_items.append(room_item)
            if bounds is None:
                fallback.append(room_item)
                continue

            try:
                x0 = int(math.floor(float(bounds[0]) / cell_size))
                x1 = int(math.floor(float(bounds[3]) / cell_size))
                y0 = int(math.floor(float(bounds[1]) / cell_size))
                y1 = int(math.floor(float(bounds[4]) / cell_size))
                if (x1 - x0 + 1) * (y1 - y0 + 1) > 400:
                    fallback.append(room_item)
                    continue
                for cell_x in range(x0, x1 + 1):
                    for cell_y in range(y0, y1 + 1):
                        grid.setdefault((cell_x, cell_y), []).append(room_item)
            except Exception:
                fallback.append(room_item)

        linked_rooms.append(
            {
                "doc": link_doc,
                "transform": transform,
                "rooms": room_items,
                "grid": grid,
                "fallback": fallback,
                "cell_size": cell_size,
                "link": link,
            }
        )

    _LINKED_ROOM_CACHE = linked_rooms
    return linked_rooms


def _build_room_context(element):
    point = _element_center_point(element)
    if point is None:
        return {
            "room": None,
            "room_name": "",
            "room_number": "",
            "department": "",
            "subdepartment": "",
            "source": "no-point",
        }

    for link_data in _get_linked_room_cache():
        room = _find_room_for_point(point, link_data)
        if room is not None:
            source_doc = link_data["doc"]
            room_name = _get_room_value_from_room(room, "NWB_RoomName", source_doc)
            room_number = _get_room_value_from_room(room, "NWB_RoomID", source_doc)
            department = _get_room_value_from_room(room, "NWB_Department", source_doc)
            subdepartment = _get_room_value_from_room(room, "NWB_SubDepartment", source_doc)
            return {
                "room": room,
                "room_name": room_name,
                "room_number": room_number,
                "department": department,
                "subdepartment": subdepartment,
                "source": "linked-room",
            }

    return {
        "room": None,
        "room_name": "",
        "room_number": "",
        "department": "",
        "subdepartment": "",
        "source": "no-linked-room",
    }


def _resolve_target_value(element, param_name, room_context):
    if param_name == "NWB_Department":
        return room_context.get("department", "")
    if param_name == "NWB_SubDepartment":
        return room_context.get("subdepartment", "")
    if param_name == "NWB_RoomID":
        return room_context.get("room_number", "")
    if param_name == "NWB_RoomName":
        return room_context.get("room_name", "")

    for _, owner in _iter_param_owners(element):
        param = _find_param(owner, param_name)
        if param is None:
            continue
        text = _read_param_text(param)
        if text:
            return text

    return ""


def _values_match(current, expected):
    return _norm(current) == _norm(expected)


def _write_text_param(param, value):
    if param is None:
        return False, "missing parameter"
    if param.IsReadOnly:
        return False, "read-only"

    try:
        text = _safe_str(value).strip()
        if param.StorageType == StorageType.String:
            ok = param.Set(text)
            return bool(ok), "OK" if ok else "Set() returned False"

        if param.StorageType == StorageType.Integer:
            try:
                number = int(float(text))
            except Exception:
                return False, "not integer"
            ok = param.Set(number)
            return bool(ok), "OK" if ok else "Set() returned False"

        if param.StorageType == StorageType.Double:
            try:
                number = float(text)
            except Exception:
                return False, "not double"
            ok = param.Set(number)
            return bool(ok), "OK" if ok else "Set() returned False"

        return False, "unsupported storage"
    except Exception as ex:
        return False, _safe_str(ex)


def _pick_write_target(element, param_name):
    for _, owner in _iter_param_owners(element):
        candidate = _find_param(owner, param_name)
        if candidate is not None and not candidate.IsReadOnly:
            return candidate
    return None


def _is_supported_element(element):
    if element is None:
        return False
    try:
        category_id = element.Category.Id.IntegerValue
        return any(category_id == int(category) for category in MEP_CATEGORIES)
    except Exception:
        return False


def _has_repair_parameter(element):
    for _, owner in _iter_param_owners(element):
        for param_name in TARGET_PARAMS:
            if _find_param(owner, param_name) is not None:
                return True
    return False


def _repair_nwb_values(elements):
    repaired = 0
    skipped = 0
    failures = []
    pending_writes = []

    for element in elements or []:
        if element is None:
            continue
        try:
            room_context = _build_room_context(element)
            for param_name in TARGET_PARAMS:
                target_value = _resolve_target_value(element, param_name, room_context)
                if _is_blank(target_value):
                    skipped += 1
                    continue

                target_param = _pick_write_target(element, param_name)
                if target_param is None:
                    continue

                current_value = _read_param_text(target_param)
                if not _is_blank(current_value) and _values_match(current_value, target_value):
                    continue

                pending_writes.append((element, param_name, target_param, target_value))
        except Exception as ex:
            failures.append("{0} | read/resolve failed | {1}".format(_safe_str(getattr(getattr(element, "Id", None), "IntegerValue", "?")), _safe_str(ex)))

    if not pending_writes:
        return repaired, skipped, failures

    tx = Transaction(doc, "Quick Repair NWB Parameters")
    try:
        tx.Start()
        for element, param_name, target_param, target_value in pending_writes:
            try:
                ok, message = _write_text_param(target_param, target_value)
            except Exception as ex:
                ok, message = False, _safe_str(ex)
            if ok:
                repaired += 1
            else:
                failures.append("{0} | {1} | {2}".format(element.Id.IntegerValue, param_name, message))
        tx.Commit()
    except Exception as ex:
        try:
            if tx.GetStatus() == TransactionStatus.Started:
                tx.RollBack()
        except Exception:
            pass
        failures.append("transaction failed | {0}".format(_safe_str(ex)))
        repaired = 0

    return repaired, skipped, failures


def _collect_target_elements():
    try:
        ids = uidoc.Selection.GetElementIds()
    except Exception:
        ids = []

    elements = []
    seen = set()
    for eid in ids:
        el = doc.GetElement(eid)
        if _is_supported_element(el) and _has_repair_parameter(el):
            try:
                if el.Id.IntegerValue in seen:
                    continue
                seen.add(el.Id.IntegerValue)
            except Exception:
                pass
            elements.append(el)

    if elements:
        return elements

    try:
        view = doc.ActiveView
        if view is None:
            return []
        for category in MEP_CATEGORIES:
            collector = FilteredElementCollector(doc, view.Id).OfCategory(category).WhereElementIsNotElementType()
            for el in collector:
                if not _is_supported_element(el) or not _has_repair_parameter(el):
                    continue
                try:
                    if el.Id.IntegerValue in seen:
                        continue
                    seen.add(el.Id.IntegerValue)
                except Exception:
                    pass
                elements.append(el)
    except Exception:
        pass

    return elements


def run():
    selected = _collect_target_elements()
    if not selected:
        forms.alert("Select elements first or open a model view with elements visible.", title=TOOL_TITLE)
        return

    repaired, skipped, failures = _repair_nwb_values(selected)

    message_lines = [
        "NWB Quick Repair complete.",
        "Repaired: {0}".format(repaired),
        "Skipped (no resolved value): {0}".format(skipped),
    ]
    if failures:
        message_lines.append("Failures:")
        for fail in failures[:10]:
            message_lines.append("- {0}".format(fail))

    forms.alert("\n".join(message_lines), title=TOOL_TITLE)


run()
