# coding: utf8
from math import pi, acos

from Autodesk.Revit.DB import (
    Line,
    ViewSection,
    XYZ,
    FilteredElementCollector,
    Grid,
    ReferencePlane,
    FamilyInstance,
    BuiltInParameter,
    ElevationMarker,
    ViewType,
    Options,
    Element,
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit import Exceptions

from pyrevit import forms, script, revit

__doc__ = """Updated version of pyRevit MEP's Make Parallel tool. Make multiple elements parallel to a reference in the XY plane. First element selected is reference. Subsequent selections are rotated to match."""
__title__ = "Make Parallel (XY)"
__author__ = "GM"

uidoc = revit.uidoc
doc = revit.doc
logger = script.get_logger()


# Removed old element_selection function


def get_view_from(element):
    # type: (Element) -> ViewSection
    sketch_parameter = element.get_Parameter(BuiltInParameter.VIEW_FIXED_SKETCH_PLANE)  # type: SketchPlane
    return doc.GetElement(doc.GetElement(sketch_parameter.AsElementId()).OwnerViewId)


def get_elevation_marker(element):
    # type: (Element) -> ElevationMarker
    view = get_view_from(element)
    for elevation_marker in FilteredElementCollector(doc).OfClass(ElevationMarker):  # type: ElevationMarker
        for i in range(4):
            id = elevation_marker.GetViewId(i)
            if view.Id == id:
                return elevation_marker


def section_direction(element):
    # type: (Element) -> XYZ
    sketch_parameter = element.get_Parameter(BuiltInParameter.VIEW_FIXED_SKETCH_PLANE)  # type: SketchPlane
    view = doc.GetElement(doc.GetElement(sketch_parameter.AsElementId()).OwnerViewId)  # type: ViewSection
    return view.RightDirection


def grid_direction(element):
    # type: (Grid) -> XYZ
    return element.Curve.Direction


def plane_direction(element):
    # type: (ReferencePlane) -> XYZ
    return element.Direction


def line_direction(element):
    return element.Location.Curve.Direction


def family_direction(element):
    # type: (FamilyInstance) -> XYZ
    return element.FacingOrientation


def scope_box(element):
    # type: (Element) -> XYZ
    options = Options()
    options.View = doc.ActiveView
    for geom in element.Geometry[options]:
        return geom.Direction


def direction(element):
    direction_funcs = (
        grid_direction,
        plane_direction,
        line_direction,
        family_direction,
        section_direction,
        scope_box,
    )

    for func in direction_funcs:
        try:
            return func(element)
        except AttributeError:
            pass
    else:
        logger.debug("DIRECTION : type {}".format(type(element)))


def section_origin(element):
    # type: (Element) -> XYZ
    sketch_parameter = element.get_Parameter(BuiltInParameter.VIEW_FIXED_SKETCH_PLANE)  # type: SketchPlane
    view = doc.GetElement(doc.GetElement(sketch_parameter.AsElementId()).OwnerViewId)  # type: ViewSection
    return view.Origin


def grid_origin(element):
    # type: (Grid) -> XYZ
    return element.Curve.Origin


def plane_origin(element):
    # type: (ReferencePlane) -> XYZ
    return element.GetPlane().Origin


def line_origin(element):
    return element.Location.Curve.Origin


def family_origin(element):
    # type: (FamilyInstance) -> XYZ
    return element.GetTransform().Origin


def scope_box_origin(element):
    # type: (Element) -> XYZ
    options = Options()
    options.View = doc.ActiveView
    for geom in element.Geometry[options]:
        return geom.Origin


def origin(element):
    origin_funcs = (
        grid_origin,
        plane_origin,
        line_origin,
        family_origin,
        section_origin,
        scope_box_origin,
    )

    for func in origin_funcs:
        try:
            return func(element)
        except AttributeError:
            continue
    else:
        logger.debug("ORIGIN : type {}".format(type(element)))


while element_selection():
    pass
