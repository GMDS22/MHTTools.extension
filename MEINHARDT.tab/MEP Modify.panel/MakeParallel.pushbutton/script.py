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


# Pick reference
try:
    with forms.WarningBar(title="Pick reference element"):
        ref_pick = uidoc.Selection.PickObject(ObjectType.Element, "Pick reference element")
except Exceptions.OperationCanceledException:
    exit()

reference_element = doc.GetElement(ref_pick)
v1 = direction(reference_element)
if v1 is None:
    forms.alert("Cannot determine direction for reference element.", title="Make Parallel")
    exit()

xy_v1 = XYZ(v1.X, v1.Y, 0)

# Now pick targets
while True:
    try:
        with forms.WarningBar(title="Pick target element (ESC to finish)"):
            target_pick = uidoc.Selection.PickObject(ObjectType.Element, "Pick target element")
    except Exceptions.OperationCanceledException:
        break

    element2 = doc.GetElement(target_pick)
    if element2.Id == reference_element.Id:
        forms.alert("Reference and target cannot be the same element.", title="Make Parallel")
        continue

    v2 = direction(element2)
    if v2 is None:
        forms.alert("Cannot determine direction for target element.", title="Make Parallel")
        continue

    xy_v2 = XYZ(v2.X, v2.Y, 0)

    angle = xy_v2.AngleTo(xy_v1)
    if angle > pi / 2:
        angle = angle - pi
    normal = xy_v2.CrossProduct(xy_v1)

    try:
        short_tol = doc.Application.ShortCurveTolerance
    except Exception:
        short_tol = 1e-6

    try:
        normal_len = normal.GetLength()
    except Exception:
        normal_len = 0.0

    if abs(angle) < 1e-9 or normal_len <= 1e-9:
        forms.alert(
            "Elements are already parallel in XY (or rotation axis is too small).",
            title="Make Parallel",
            warn_icon=False,
        )
        continue

    logger.debug("ANGLE : {}".format(angle))
    logger.debug("NORMAL : {}".format(normal))
    logger.debug("DIRECTION \n 1: {} \n 2: {}".format(direction(reference_element), direction(element2)))

    axis_origin = origin(element2)
    if axis_origin is None:
        forms.alert("Cannot determine origin for the target element.", title="Make Parallel")
        continue

    try:
        axis_dir = normal.Normalize()
    except Exception:
        axis_dir = XYZ(normal.X / normal_len, normal.Y / normal_len, normal.Z / normal_len)

    axis_len = max(short_tol * 100.0, 1.0)
    axis = Line.CreateBound(axis_origin, axis_origin + axis_dir.Multiply(axis_len))

    # Need to rotate elevation marker if it is an elevation
    try:
        if get_view_from(element2).ViewType == ViewType.Elevation:
            element2 = get_elevation_marker(element2)
    except AttributeError:
        pass

    if not hasattr(element2.Location, 'Rotate'):
        forms.alert("The target element does not support rotation.", title="Make Parallel")
        continue

    with revit.Transaction("Make parallel", doc):
        element2.Location.Rotate(axis, angle)
