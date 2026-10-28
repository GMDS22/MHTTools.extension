# -*- coding: utf-8 -*-
"""MHT Fire & Smoke Penetration Detector.

Native pyRevit implementation. Detection is read-only; marker placement is
blocked until an approved family/symbol has been configured and validated.
"""

import clr
import os
import sys
import time
import traceback
from datetime import datetime

clr.AddReference("System")
clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")

from Autodesk.Revit.DB import (
    BuiltInCategory, ElementId, FilteredElementCollector, FamilySymbol,
    Options, SolidUtils, StorageType, Transform, Wall
)
from Autodesk.Revit.DB import Mechanical as Mechanical
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import DB, forms, revit, script

TOOL_VERSION = "0.1.0-audit-gated"
THIS_DIR = os.path.dirname(__file__)
if THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)
import core

logger = script.get_logger()
uidoc = revit.uidoc
doc = uidoc.Document if uidoc else None


def _safe_param_text(element, name):
    if not element or not name:
        return ""
    try:
        p = element.LookupParameter(name)
        if not p:
            try:
                element_type = getattr(element, "WallType", None)
                p = element_type.LookupParameter(name) if element_type else None
            except Exception:
                p = None
        if not p:
            return ""
        value = p.AsValueString() or p.AsString()
        if value:
            return value
        if p.StorageType == StorageType.Integer:
            return str(p.AsInteger())
        if p.StorageType == StorageType.Double:
            return str(p.AsDouble())
        if p.StorageType == StorageType.ElementId:
            element_id = p.AsElementId()
            if element_id and element_id != ElementId.InvalidElementId:
                return str(_id_value(element_id))
    except Exception:
        pass
    return ""


def _safe_param_alias_text(element, aliases_text, default_alias=""):
    aliases = core.split_parameter_aliases(aliases_text)
    if default_alias and default_alias not in aliases:
        aliases.insert(0, default_alias)
    for alias in aliases:
        value = _safe_param_text(element, alias)
        if value and value.strip():
            return value, alias
    return "", ""


def _id_value(element_id):
    try:
        return element_id.IntegerValue
    except Exception:
        return int(element_id)


def _uid(element):
    try:
        return element.UniqueId
    except Exception:
        return ""


def _point_tuple(point):
    return (point.X, point.Y, point.Z)


def _host_document_guid(document):
    try:
        return str(document.ProjectInformation.Id.IntegerValue) if document else "UNKNOWN"
    except Exception:
        return "UNKNOWN"


def _solids(element, transform=None, options=None):
    result = []
    if element is None:
        return result
    try:
        geometry = element.get_Geometry(options or Options())
    except Exception:
        return result
    for obj in geometry:
        if isinstance(obj, DB.Solid) and obj.Volume > 1e-9:
            result.append(obj)
        elif isinstance(obj, DB.GeometryInstance):
            nested = obj.GetInstanceGeometry() if obj.GetInstanceGeometry else obj.SymbolGeometry
            for solid in nested:
                if isinstance(solid, DB.Solid) and solid.Volume > 1e-9:
                    result.append(solid)
    return result


def _point_from_geometry_result(result):
    try:
        return _point_tuple(result[0]) if hasattr(result[0], "X") else None
    except Exception:
        return None


def _reference_translation(element):
    return element.LookupParameter("Reference Level") if element else None


def _wall_class(element, fire_names, smoke_names, combined_names):
    fire_text, fire_alias = _safe_param_alias_text(element, fire_names)
    smoke_text, smoke_alias = _safe_param_alias_text(element, smoke_names)
    combined_text, combined_alias = _safe_param_alias_text(element, combined_names)

    if combined_text:
        classification = core.classify_wall_material(combined_text)
        if classification != "UNKNOWN":
            return classification, combined_text, smoke_text
    if fire_text or smoke_text:
        classification = core.wall_classification(fire_text, smoke_text, "")
        return classification, fire_text, smoke_text

    nwb_text, nwb_alias = _safe_param_alias_text(element, "NWB_Material", "NWB_Material")
    return core.classify_wall_material(nwb_text), nwb_text, ""


def _wall_transform(transform):
    return transform


class DetectorWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, os.path.join(THIS_DIR, "FireSmokePenetration.xaml"))
        self.results = []
        self.findings = []
        self.close_button.Click += self.close_click
        self.analyze_button.Click += self.analyze_click
        self.export_button.Click += self.export_click
        self.place_button.Click += self.place_click

    def close_click(self, sender, args):
        self.Close()

    def _float(self, control, default):
        try:
            return float(control.Text.strip())
        except Exception:
            return default

    def _selected_scope(self):
        try:
            return str(self.scope_box.SelectedItem.Content).lower()
        except Exception:
            return "active model"

    def analyze_click(self, sender, args):
        start = time.time()
        self.results_box.Text = "Analyzing..."
        try:
            scope = "active view" if "view" in self._selected_scope() else "active model"
            include_links = bool(self.include_links_box.IsChecked)
            tolerance_mm = self._float(self.tolerance_box, 3.0)
            tolerance_ft = core.mm_to_ft(tolerance_mm)
            fire_name = self.fire_param_box.Text.strip()
            smoke_name = self.smoke_param_box.Text.strip()
            combined_name = self.combined_param_box.Text.strip()
            missing_config = not any((fire_name, smoke_name, combined_name))

            walls, ducts, diagnostics = self._collect(scope, include_links)
            pairs = 0
            verified = []
            if not missing_config:
                for wall_item in walls:
                    wall = wall_item["element"]
                    classification, fire_text, smoke_text = _wall_class(wall, fire_name, smoke_name, combined_name)
                    wall_item.update(classification=classification, fire_text=fire_text, smoke_text=smoke_text)
                    if classification == "UNKNOWN" or classification == "NOT_RATED":
                        continue
                    for duct in ducts:
                        pairs += 1
                        if self._confirm_intersection(wall_item, duct, tolerance_ft):
                            verified.append(self._finding(wall_item, duct, classification, fire_text, smoke_text))
            self.findings = verified
            elapsed = time.time() - start
            lines = [
                "MHT Fire & Smoke Penetration Detector {}".format(TOOL_VERSION),
                "Result: {} verified candidate penetrations".format(len(verified)),
                "Walls considered: {}".format(len(walls)),
                "Eligible ducts: {}".format(len(ducts)),
                "Candidate pairs evaluated: {}".format(pairs),
                "Elapsed: {:.2f} seconds".format(elapsed),
            ]
            if missing_config:
                lines.append("STOPPED: No wall classification parameter is configured. Unknown ratings are not treated as negative.")
            lines.extend(diagnostics)
            self.results_box.Text = "\n".join(lines)
        except Exception as exc:
            logger.error(traceback.format_exc())
            self.results_box.Text = "Analysis failed:\n{}".format(exc)

    def _collect(self, scope, include_links):
        walls = []
        ducts = []
        diagnostics = []
        wall_cats = [BuiltInCategory.OST_Walls]
        wall_ids = ElementId.CreateList(wall_cats, False) if hasattr(ElementId, "CreateList") else None
        if scope == "active view":
            wall_elements = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
        else:
            wall_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        wall_elements = [e for e in wall_elements if e.Category and e.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Walls)]
        for wall in wall_elements:
            if isinstance(wall, Wall):
                walls.append({"element": wall, "document": doc, "link_uid": "HOST-WALL", "source_document": doc.Title, "transform": Transform.Identity})

        if scope == "active view":
            duct_elements = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
        else:
            duct_elements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
        ducts.extend(list(duct_elements))

        if include_links:
            links = FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).WhereElementIsNotElementType().ToElements()
            for link in links:
                linked_doc = link.GetLinkDocument() if hasattr(link, "GetLinkDocument") else None
                if not linked_doc:
                    diagnostics.append("Link unloaded or inaccessible: {}".format(_uid(link)))
                    continue
                transform = link.GetTotalTransform() if hasattr(link, "GetTotalTransform") else link.GetTransform()
                linked_walls = FilteredElementCollector(linked_doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
                for wall in linked_walls:
                    if isinstance(wall, Wall):
                        walls.append({"element": wall, "document": linked_doc, "link_uid": _uid(link), "source_document": linked_doc.Title, "transform": transform})
        diagnostics.append("Nested links are not followed. Unsupported nested instances are omitted and reported only as model coverage limitation.")
        return walls, ducts, diagnostics

    def _confirm_intersection(self, wall_item, duct, tolerance_ft):
        try:
            duct_solids = _solids(duct)
            wall_solids = _solids(wall_item["element"])
            if not duct_solids or not wall_solids:
                return False
            transform = wall_item["transform"]
            for duct_solid in duct_solids:
                for wall_solid in wall_solids:
                    candidate = wall_solid
                    if transform is not None and transform != Transform.Identity:
                        candidate = SolidUtils.CreateTransformed(wall_solid, transform)
                    intersection = duct_solid.Intersect(candidate)
                    if intersection == DB.SetComparisonResult.Disjoint:
                        continue
                    if intersection == DB.SetComparisonResult.Subset or intersection == DB.SetComparisonResult.Equal:
                        return True
                    if intersection == DB.SetComparisonResult.Overlap:
                        return True
        except Exception as exc:
            logger.debug("Intersection check skipped: {}".format(exc))
        return False

    def _finding(self, wall_item, duct, classification, fire_text, smoke_text):
        wall = wall_item["element"]
        link_uid = wall_item.get("link_uid", "HOST-WALL")
        source_identity = "{}|{}|{}|{}".format(_host_document_guid(doc), link_uid, _uid(wall), _uid(duct))
        geometry_signature = "{}:{}:{}".format(_id_value(wall.Id), _id_value(duct.Id), classification)
        finding_id = core.make_finding_id(_host_document_guid(doc), link_uid, _uid(wall), _uid(duct), geometry_signature)
        duct_center = getattr(duct, "Location", None)
        point = ""
        if duct_center and hasattr(duct_center, "Curve") and hasattr(duct_center.Curve, "GetEndPoint"):
            curve = duct_center.Curve
            p = curve.GetEndPoint(0)
            q = curve.GetEndPoint(1)
            point = "{:.6f},{:.6f},{:.6f}".format((p.X + q.X) / 2.0, (p.Y + q.Y) / 2.0, (p.Z + q.Z) / 2.0)
        coords = (point.split(",") + ["", "", ""][:3])[:3]
        return {
            "finding_id": finding_id,
            "status": "CONFIRMED_GEOMETRY" if classification in ("FIRE", "SMOKE", "BOTH") else "UNKNOWN_CLASSIFICATION",
            "review_status": "Not Reviewed",
            "comments": "",
            "host_document_guid": _host_document_guid(doc),
            "link_instance_uid": link_uid,
            "source_document": wall_item.get("source_document", ""),
            "wall_uid": _uid(wall),
            "duct_uid": _uid(duct),
            "wall_classification": classification,
            "point_x": coords[0],
            "point_y": coords[1],
            "point_z": coords[2],
            "wall_thickness_mm": "",
            "source_identity": source_identity,
            "warnings": "Precise entry/exit point requires Revit solid boundary sampling; engineering approval not determined.",
            "detected_at": datetime.now().isoformat(timespec="seconds"),
        }

    def export_click(self, sender, args):
        if not self.findings:
            forms.alert("Run analysis first. No findings are available to export.", exitscript=False)
            return
        path = forms.save_file(file_ext="csv", title="Export fire and smoke penetration findings")
        if path:
            try:
                core.write_csv(path, self.findings)
                forms.alert("Exported {} findings.".format(len(self.findings)), exitscript=False)
            except Exception as exc:
                forms.alert("Export failed: {}".format(exc), exitscript=False)

    def place_click(self, sender, args):
        family_name = self.marker_family_box.Text.strip()
        symbol_name = self.marker_symbol_box.Text.strip()
        if not family_name or not symbol_name:
            forms.alert("Marker placement is blocked. Supply both an approved marker family and symbol.", exitscript=False)
            return
        families = FilteredElementCollector(doc).OfClass(DB.Family).ToElements()
        family = next((f for f in families if f.Name == family_name), None)
        if family is None:
            forms.alert("Approved marker family was not found: {}".format(family_name), exitscript=False)
            return
        symbol = next((s for s in family.GetFamilySymbolIds() if doc.GetElement(s).Name == symbol_name), None)
        if symbol is None:
            forms.alert("Approved marker symbol was not found: {}".format(symbol_name), exitscript=False)
            return
        forms.alert("The family and symbol exist, but placement parameters and reconciliation have not been approved. No model elements were changed.", exitscript=False)


def _ensure_project_environment():
    if not doc:
        forms.alert("This tool requires an active Revit project.", exitscript=True)


_ensure_project_environment()
window = DetectorWindow()
window.ShowDialog()
