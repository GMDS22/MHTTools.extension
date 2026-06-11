# coding: utf8
from __future__ import print_function

import traceback

from Autodesk.Revit.DB import (
    BooleanOperationsType,
    BooleanOperationsUtils,
    CylindricalFace,
    Document,
    ElementId,
    FilledRegion,
    FilledRegionType,
    FilteredElementCollector,
    GeometryCreationUtilities,
    GraphicsStyle,
    GraphicsStyleType,
    Transaction,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType, Selection
from pyrevit import forms


TOOL_TITLE = "Combine Filled Regions"

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document  # type: Document
selection = uidoc.Selection  # type: Selection


def _line_style_name(graphics_style):
    if graphics_style is None:
        return ""

    try:
        category = graphics_style.GraphicsStyleCategory
        if category is not None and category.Name:
            return category.Name
    except Exception:
        pass

    return ""


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


def find_top_faces(solid):
    """Find top faces on the merged solid."""
    faces = []
    for face in solid.Faces:
        if isinstance(face, CylindricalFace):
            continue

        try:
            if face.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ):
                faces.append(face)
        except Exception:
            pass

    return faces


class RegionSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, FilledRegion)

    def AllowReference(self, reference, point):
        return True


def pick_filled_regions():
    refs = None
    with forms.WarningBar(title='Pick Filled Regions and click "Finish"'):
        try:
            refs = selection.PickObjects(ObjectType.Element, RegionSelectionFilter())
        except Exception:
            refs = None

    if not refs:
        forms.alert("Filled regions were not selected. Please try again.", title=TOOL_TITLE, exitscript=True)

    regions = [doc.GetElement(ref) for ref in refs]
    regions = [region for region in regions if isinstance(region, FilledRegion)]

    if len(regions) < 2:
        forms.alert("You need to select at least 2 filled regions.", title=TOOL_TITLE, exitscript=True)

    return regions


def merge_regions(selected_regions):
    transaction = Transaction(doc, TOOL_TITLE)
    transaction.Start()
    try:
        shapes = []
        for region in selected_regions:
            boundaries = region.GetBoundaries()
            shape = GeometryCreationUtilities.CreateExtrusionGeometry(boundaries, XYZ(0, 0, 1), 10)
            shapes.append(shape)

        merged_shape = shapes.pop()
        for shape in shapes:
            merged_shape = BooleanOperationsUtils.ExecuteBooleanOperation(
                merged_shape,
                shape,
                BooleanOperationsType.Union,
            )

        top_faces = find_top_faces(merged_shape)
        if not top_faces:
            raise Exception("No top face could be resolved from the merged filled-region geometry.")

        final_outline = None
        for index, face in enumerate(top_faces):
            outline = face.GetEdgesAsCurveLoops()
            if index == 0:
                final_outline = outline
            else:
                final_outline.AddRange(outline)

        source_type_id = selected_regions[0].GetTypeId()
        source_type = doc.GetElement(source_type_id)
        if source_type is None or not isinstance(source_type, FilledRegionType):
            raise Exception("The selected filled regions do not share a valid source filled region type.")

        new_region = FilledRegion.Create(doc, source_type_id, doc.ActiveView.Id, final_outline)
        if new_region is None:
            raise Exception("Revit did not create a merged filled region.")

        invisible_line_style_id = _invisible_line_style_id()
        if invisible_line_style_id != ElementId.InvalidElementId:
            try:
                new_region.SetLineStyleId(invisible_line_style_id)
            except Exception:
                pass

        for region in selected_regions:
            doc.Delete(region.Id)

        transaction.Commit()
    except Exception:
        try:
            transaction.RollBack()
        except Exception:
            pass
        forms.alert(traceback.format_exc(), title=TOOL_TITLE)


def main():
    if doc.ActiveView is None:
        forms.alert("No active view found.", title=TOOL_TITLE, exitscript=True)

    selected_regions = pick_filled_regions()
    merge_regions(selected_regions)


if __name__ == "__main__":
    main()
