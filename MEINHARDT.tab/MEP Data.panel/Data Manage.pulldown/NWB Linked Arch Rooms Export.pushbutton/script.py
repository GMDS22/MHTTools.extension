# coding: utf8
from __future__ import print_function

import os
import re
import time
import codecs

import clr

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    RevitLinkInstance,
    StorageType,
)
from pyrevit import forms, revit, script


doc = revit.doc
logger = script.get_logger()

__title__ = "NWB Linked\nArch Rooms"
__doc__ = "Export linked architectural room information to a CSV file for NWB mapping and QA."

try:
    _TEXT_TYPE = unicode
except NameError:
    _TEXT_TYPE = str


class LinkExportOption(object):
    def __init__(self, link_inst, link_doc, rooms):
        self.link_inst = link_inst
        self.link_doc = link_doc
        self.rooms = list(rooms or [])
        self.link_id = self._get_link_id(link_inst)
        self.link_name = self._get_link_name(link_inst, link_doc)
        self.name = "{0} ({1} rooms)".format(self.link_name, len(self.rooms))

    def _get_link_id(self, link_inst):
        try:
            return link_inst.Id.IntegerValue
        except Exception:
            return -1

    def _get_link_name(self, link_inst, link_doc):
        for candidate in [
            getattr(link_inst, "Name", None),
            getattr(link_doc, "Title", None),
        ]:
            text = _safe_str(candidate).strip()
            if text:
                return text
        return "Link {0}".format(self.link_id)


def _safe_str(value):
    if value is None:
        return ""
    try:
        if isinstance(value, _TEXT_TYPE):
            return value
    except Exception:
        pass
    try:
        return _TEXT_TYPE(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return ""


def _sanitize_filename_part(value, max_len=80):
    text = _safe_str(value).strip()
    if not text:
        return "NA"
    text = re.sub(r"[^A-Za-z0-9._-]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-._")
    if not text:
        text = "NA"
    if len(text) > max_len:
        text = text[:max_len].rstrip("-._")
    return text or "NA"


def _read_parameter_value(param, source_doc=None):
    if param is None:
        return ""
    src_doc = source_doc or doc

    value = ""
    try:
        st = param.StorageType
    except Exception:
        st = None

    try:
        if st == StorageType.String:
            value = param.AsString() or ""
        elif st == StorageType.Integer:
            value = param.AsValueString() or _safe_str(param.AsInteger())
        elif st == StorageType.Double:
            value = param.AsValueString() or _safe_str(param.AsDouble())
        elif st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is not None and eid != ElementId.InvalidElementId:
                ref_el = None
                try:
                    ref_el = src_doc.GetElement(eid)
                except Exception:
                    ref_el = None
                if ref_el is not None:
                    value = _safe_str(getattr(ref_el, "Name", ""))
                if not value:
                    value = _safe_str(getattr(eid, "IntegerValue", ""))
        else:
            value = ""
    except Exception:
        value = ""

    if not value:
        try:
            value = param.AsValueString() or ""
        except Exception:
            value = ""
    if not value:
        try:
            value = param.AsString() or ""
        except Exception:
            value = ""

    return _safe_str(value).strip()


def _get_room_parameter_map(room, room_doc):
    values = {}
    try:
        iterator = room.Parameters
    except Exception:
        iterator = []

    for param in iterator:
        if param is None:
            continue
        definition = None
        try:
            definition = param.Definition
        except Exception:
            definition = None
        if definition is None:
            continue
        name = _safe_str(getattr(definition, "Name", "")).strip()
        if not name:
            continue
        value = _read_parameter_value(param, room_doc)
        if value and name not in values:
            values[name] = value

    return values


def _first_non_blank(values, names):
    for name in names:
        text = _safe_str(values.get(name, "")).strip()
        if text:
            return text
    return ""


def _get_level_name(room, room_doc):
    try:
        level = room_doc.GetElement(room.LevelId)
        if level is not None:
            return _safe_str(getattr(level, "Name", "")).strip()
    except Exception:
        pass
    return ""


def _get_phase_name(room, room_doc):
    for attr_name in ["PhaseId", "CreatedPhaseId"]:
        try:
            phase_id = getattr(room, attr_name, None)
            if phase_id is None or phase_id == ElementId.InvalidElementId:
                continue
            phase = room_doc.GetElement(phase_id)
            if phase is not None:
                text = _safe_str(getattr(phase, "Name", "")).strip()
                if text:
                    return text
        except Exception:
            continue
    return ""


def _get_room_point(room):
    try:
        location = room.Location
        if location is not None and hasattr(location, "Point") and location.Point is not None:
            return location.Point
    except Exception:
        pass

    try:
        bb = room.get_BoundingBox(None)
        if bb is not None:
            return bb.Min.Add(bb.Max).Multiply(0.5)
    except Exception:
        pass
    return None


def _collect_link_options():
    options = []
    try:
        links = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    except Exception:
        links = []

    for link_inst in links:
        if link_inst is None:
            continue
        try:
            link_doc = link_inst.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue

        try:
            rooms = list(
                FilteredElementCollector(link_doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
            )
        except Exception:
            rooms = []

        if not rooms:
            continue
        options.append(LinkExportOption(link_inst, link_doc, rooms))

    options.sort(key=lambda item: item.link_name.lower())
    return options


def _select_link_options(options):
    if not options:
        return []
    if len(options) == 1:
        return options

    selected = forms.SelectFromList.show(
        options,
        name_attr="name",
        title="Select linked architectural room models to export",
        multiselect=True,
        button_name="Export Linked Rooms",
    )
    if not selected:
        return []
    if isinstance(selected, list):
        return selected
    return [selected]


def _collect_room_rows(selected_options):
    rows = []
    extra_headers = []
    extra_header_set = set()
    exported_at = time.strftime("%Y-%m-%d %H:%M:%S")

    for option in selected_options:
        link_doc = option.link_doc
        for room in option.rooms:
            if room is None:
                continue
            param_map = _get_room_parameter_map(room, link_doc)
            point = _get_room_point(room)

            row = {
                "ExportedAt": exported_at,
                "LinkInstanceId": _safe_str(option.link_id),
                "LinkInstanceName": option.link_name,
                "LinkDocumentTitle": _safe_str(getattr(link_doc, "Title", "")).strip(),
                "RoomId": _safe_str(getattr(room.Id, "IntegerValue", "")),
                "RoomUniqueId": _safe_str(getattr(room, "UniqueId", "")),
                "RoomNumber": _first_non_blank(param_map, ["Number", "Room Number"]),
                "RoomName": _first_non_blank(param_map, ["Name", "Room Name"]),
                "Level": _get_level_name(room, link_doc),
                "Phase": _get_phase_name(room, link_doc),
                "Department": _first_non_blank(param_map, ["Department", "Room Department", "Department Name"]),
                "ResolvedDepartment": _first_non_blank(param_map, ["NWB_Department", "Department", "Room Department", "Department Name"]),
                "ResolvedAreaBuilding": _first_non_blank(param_map, ["Area/Building", "Area Building"]),
                "ResolvedBuilding": _first_non_blank(param_map, ["Building", "Building Name", "Building Code", "Block"]),
                "ResolvedSite": _first_non_blank(param_map, ["Project Location", "Site", "Site Name", "Site Code", "Campus", "Facility"]),
                "Area": _first_non_blank(param_map, ["Area"]),
                "Perimeter": _first_non_blank(param_map, ["Perimeter"]),
                "LocationX": "",
                "LocationY": "",
                "LocationZ": "",
            }

            if point is not None:
                try:
                    row["LocationX"] = _safe_str(round(point.X, 6))
                    row["LocationY"] = _safe_str(round(point.Y, 6))
                    row["LocationZ"] = _safe_str(round(point.Z, 6))
                except Exception:
                    pass

            for header, value in param_map.items():
                row[header] = value
                if header not in extra_header_set:
                    extra_header_set.add(header)
                    extra_headers.append(header)

            rows.append(row)

    extra_headers.sort(key=lambda text: text.lower())
    return rows, extra_headers


def _csv_escape(value):
    text = _safe_str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace('"', '""')
    if any(token in text for token in [",", '"', "\n"]):
        return '"{0}"'.format(text)
    return text


def _write_csv(path, headers, rows):
    try:
        with codecs.open(path, "w", "utf-8-sig") as stream:
            stream.write(",".join([_csv_escape(header) for header in headers]) + "\r\n")
            for row in rows:
                values = [_csv_escape(row.get(header, "")) for header in headers]
                stream.write(",".join(values) + "\r\n")
    except Exception as ex:
        raise RuntimeError("CSV export failed: {0}".format(_safe_str(ex)))


def _open_export_file(path):
    try:
        os.startfile(path)
        return True
    except Exception as ex:
        logger.warning("Could not open exported file automatically: {0}".format(_safe_str(ex)))
        return False


def _build_output_path(selected_options):
    script_dir = os.path.dirname(__file__)
    tool_root = os.path.dirname(script_dir)
    target_dir = os.path.join(tool_root, "NWB_PARAMETERS AutoFill.pushbutton")
    if not os.path.isdir(target_dir):
        target_dir = script_dir

    stamp = time.strftime("%Y%m%d_%H%M%S")
    if len(selected_options) == 1:
        name_part = _sanitize_filename_part(selected_options[0].link_name, max_len=40)
        file_name = "NWB_Linked_ARCH_Rooms_{0}_{1}.csv".format(name_part, stamp)
    else:
        file_name = "NWB_Linked_ARCH_Rooms_AllLinks_{0}.csv".format(stamp)
    return os.path.join(target_dir, file_name)


def main():
    options = _collect_link_options()
    if not options:
        forms.alert(
            "No loaded Revit links with rooms were found. Load the architectural model link first and try again.",
            title="NWB Linked Arch Rooms",
        )
        return

    selected_options = _select_link_options(options)
    if not selected_options:
        return

    rows, extra_headers = _collect_room_rows(selected_options)
    if not rows:
        forms.alert(
            "No room rows could be collected from the selected links.",
            title="NWB Linked Arch Rooms",
        )
        return

    headers = [
        "ExportedAt",
        "LinkInstanceId",
        "LinkInstanceName",
        "LinkDocumentTitle",
        "RoomId",
        "RoomUniqueId",
        "RoomNumber",
        "RoomName",
        "Level",
        "Phase",
        "Department",
        "ResolvedDepartment",
        "ResolvedAreaBuilding",
        "ResolvedBuilding",
        "ResolvedSite",
        "Area",
        "Perimeter",
        "LocationX",
        "LocationY",
        "LocationZ",
    ]
    for header in extra_headers:
        if header not in headers:
            headers.append(header)

    output_path = _build_output_path(selected_options)
    _write_csv(output_path, headers, rows)
    opened = _open_export_file(output_path)

    logger.info("Exported {0} linked rooms to {1}".format(len(rows), output_path))
    forms.alert(
        "Exported {0} linked rooms to:\n\n{1}{2}".format(
            len(rows),
            output_path,
            "\n\nThe CSV was opened automatically." if opened else "\n\nOpen the CSV manually if it did not launch.",
        ),
        title="NWB Linked Arch Rooms",
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        logger.exception("Linked room export failed")
        forms.alert(
            "Linked room export failed.\n\n{0}".format(_safe_str(ex)),
            title="NWB Linked Arch Rooms",
        )
