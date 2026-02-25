# -*- coding: utf-8 -*-
from __future__ import print_function

from pyrevit import revit, DB, forms, script

logger = script.get_logger()
output = script.get_output()

# Revit internal units are feet.
MIN_CURVE_LEN_FT = 0.005  # ~1/16"


def _chunks(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def _curve_key(curve, tol=0.01):
    # Dedup in 2D (XY) with rounding tolerance.
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        pts = list(curve.Tessellate())
        if len(pts) < 2:
            return None
        p0, p1 = pts[0], pts[-1]

    def _r(p):
        return (round(p.X / tol) * tol, round(p.Y / tol) * tol)

    a = _r(p0)
    b = _r(p1)
    return (a, b) if a <= b else (b, a)


def _flatten_curve_to_elevation(curve, elevation, z_tol=1e-4):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        if abs(p0.Z - p1.Z) > z_tol:
            return None
        dz = elevation - p0.Z
        if abs(dz) < z_tol:
            return curve
        t = DB.Transform.CreateTranslation(DB.XYZ(0, 0, dz))
        return curve.CreateTransformed(t)
    except Exception:
        return None


def _get_param_as_string(elem, param_name):
    try:
        p = elem.LookupParameter(param_name)
        if p and p.HasValue:
            return p.AsString()
    except Exception:
        return None
    return None


def _set_param_string(elem, param_name, value):
    if value is None:
        return False
    try:
        p = elem.LookupParameter(param_name)
        if p and not p.IsReadOnly:
            p.Set(value)
            return True
    except Exception:
        return False
    return False


def _get_space_name_number(space):
    name_value = _get_param_as_string(space, 'Name')
    number_value = _get_param_as_string(space, 'Number')

    if not name_value:
        try:
            p = space.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
            if p and p.HasValue:
                name_value = p.AsString()
        except Exception:
            pass
    if not name_value:
        try:
            p = space.get_Parameter(DB.BuiltInParameter.ELEM_ROOM_NAME)
            if p and p.HasValue:
                name_value = p.AsString()
        except Exception:
            pass

    if not number_value:
        try:
            p = space.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER)
            if p and p.HasValue:
                number_value = p.AsString()
        except Exception:
            pass
    if not number_value:
        try:
            p = space.get_Parameter(DB.BuiltInParameter.ELEM_ROOM_NUMBER)
            if p and p.HasValue:
                number_value = p.AsString()
        except Exception:
            pass

    return name_value, number_value


def _get_area_plan_views_for_level(doc, level):
    area_plans = []
    for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan).ToElements():
        try:
            if v.IsTemplate:
                continue
            if v.ViewType != DB.ViewType.AreaPlan:
                continue
            if v.GenLevel and v.GenLevel.Id == level.Id:
                area_plans.append(v)
        except Exception:
            continue
    return area_plans


def _pick_area_plan_view(doc, level):
    area_plans = _get_area_plan_views_for_level(doc, level)
    if not area_plans:
        return None
    if len(area_plans) == 1:
        return area_plans[0]
    return forms.SelectFromList.show(
        area_plans,
        name_attr='Name',
        title='Select Area Plan View',
        button_name='Use View',
        multiselect=False
    )


def _pick_area_scheme(doc):
    schemes = list(DB.FilteredElementCollector(doc).OfClass(DB.AreaScheme).ToElements())
    if not schemes:
        forms.alert('No Area Schemes exist in this project. Create an Area Scheme first.', exitscript=True)
    if len(schemes) == 1:
        return schemes[0]
    return forms.SelectFromList.show(
        schemes,
        name_attr='Name',
        title='Select Area Scheme',
        button_name='Use Scheme',
        multiselect=False
    )


def _get_default_area_tag_type(doc):
    try:
        types = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_AreaTags)
                 .WhereElementIsElementType()
                 .ToElements())
        return types[0] if types else None
    except Exception:
        return None


def _create_area_tag(doc, view, area, point_xyz, tag_type):
    uv = DB.UV(point_xyz.X, point_xyz.Y)
    creator = doc.Create

    tag = None
    for args in (
        (view, area, uv),
        (view.Id, area.Id, uv),
    ):
        try:
            if hasattr(creator, 'NewAreaTag'):
                tag = creator.NewAreaTag(*args)
                break
        except Exception:
            continue

    if tag is not None and tag_type is not None:
        try:
            tag.ChangeTypeId(tag_type.Id)
        except Exception:
            pass

    return tag


def _create_area_boundary_lines(doc, area_view, sketch_plane, curves):
    creator = doc.Create
    created = 0
    failed = 0

    if hasattr(creator, 'NewAreaBoundaryLines'):
        for batch in _chunks(curves, 200):
            ca = DB.CurveArray()
            for c in batch:
                ca.Append(c)
            try:
                creator.NewAreaBoundaryLines(sketch_plane, ca, area_view)
                created += len(batch)
            except Exception as ex:
                logger.debug('Batch boundary creation failed, per-curve: %s', ex)
                for c in batch:
                    ca2 = DB.CurveArray()
                    ca2.Append(c)
                    try:
                        creator.NewAreaBoundaryLines(sketch_plane, ca2, area_view)
                        created += 1
                    except Exception:
                        failed += 1
        return created, failed

    if hasattr(creator, 'NewAreaBoundaryLine'):
        for c in curves:
            try:
                creator.NewAreaBoundaryLine(sketch_plane, c, area_view)
                created += 1
            except Exception:
                failed += 1
        return created, failed

    forms.alert('This Revit version/API does not expose NewAreaBoundaryLine(s).', exitscript=True)


def _create_area(doc, area_view, point_xyz):
    uv = DB.UV(point_xyz.X, point_xyz.Y)
    creator = doc.Create

    for args in (
        (area_view, uv),
        (area_view.Id, uv),
    ):
        try:
            if hasattr(creator, 'NewArea'):
                a = creator.NewArea(*args)
                if a is not None:
                    return a
        except Exception:
            continue

    return None


def main():
    doc = revit.doc
    view = revit.active_view

    if view is None or view.IsTemplate:
        forms.alert('Open a plan view (not a template) and try again.', exitscript=True)

    plan_like = set([
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.EngineeringPlan,
        DB.ViewType.AreaPlan,
    ])
    try:
        if view.ViewType not in plan_like:
            forms.alert('Active view must be a plan view (Floor/Ceiling/Engineering/Area Plan).', exitscript=True)
    except Exception:
        forms.alert('Active view must be a plan view.', exitscript=True)

    level = None
    try:
        level = view.GenLevel
    except Exception:
        pass
    if level is None and hasattr(view, 'LevelId'):
        try:
            level = doc.GetElement(view.LevelId)
        except Exception:
            level = None

    if level is None:
        forms.alert('Could not determine the level of the active plan view.', exitscript=True)

    # Collect Spaces visible in current view
    try:
        spaces = (DB.FilteredElementCollector(doc, view.Id)
                  .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
                  .WhereElementIsNotElementType()
                  .ToElements())
    except Exception:
        spaces = []

    if not spaces:
        forms.alert('No MEP Spaces found in the current view.', exitscript=True)

    # Choose/create an Area Plan view for this level
    target_view = view
    created_view = None
    scheme = None

    if view.ViewType != DB.ViewType.AreaPlan:
        picked = _pick_area_plan_view(doc, level)
        if picked:
            target_view = picked
        else:
            scheme = _pick_area_scheme(doc)
            if not scheme:
                return

    tag_type = _get_default_area_tag_type(doc)
    elev = level.Elevation

    opts = DB.SpatialElementBoundaryOptions()
    if hasattr(DB, 'SpatialElementBoundaryLocation'):
        try:
            opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
        except Exception:
            pass

    # Collect curves + placement points
    raw_curves = []
    space_points = []  # (XYZ, space)

    for sp in spaces:
        try:
            # Skip unplaced/invalid spaces
            try:
                if hasattr(sp, 'Area') and sp.Area <= 0:
                    continue
            except Exception:
                pass

            lp = getattr(sp, 'Location', None)
            pt = getattr(lp, 'Point', None) if lp else None
            if pt is not None:
                space_points.append((pt, sp))

            seglists = sp.GetBoundarySegments(opts)
            if not seglists:
                continue
            for seglist in seglists:
                for seg in seglist:
                    try:
                        raw_curves.append(seg.GetCurve())
                    except Exception:
                        try:
                            raw_curves.append(seg.Curve)
                        except Exception:
                            pass
        except Exception as ex:
            logger.debug('Space boundary read failed: %s', ex)

    if not raw_curves:
        forms.alert('No Space boundary curves found in the current view Spaces.', exitscript=True)

    flattened = []
    skipped = 0
    skipped_too_short = 0
    for c in raw_curves:
        if c is None:
            continue
        try:
            if hasattr(c, 'Length') and c.Length <= MIN_CURVE_LEN_FT:
                skipped_too_short += 1
                continue
        except Exception:
            pass

        flat = _flatten_curve_to_elevation(c, elev)
        if flat is None:
            skipped += 1
            continue
        flattened.append(flat)

    uniq = {}
    for c in flattened:
        k = _curve_key(c)
        if k is None:
            continue
        uniq[k] = c

    curves = list(uniq.values())
    if not curves:
        forms.alert('No usable Space boundary curves found (after flattening/dedup).', exitscript=True)

    created_bounds = 0
    failed_bounds = 0
    created_areas = 0
    failed_areas = 0
    created_tags = 0
    failed_tags = 0

    with revit.Transaction('Spaces (current view) to Areas + Tags'):
        # Create Area Plan if needed
        if target_view.ViewType != DB.ViewType.AreaPlan:
            try:
                created_view = DB.ViewPlan.CreateAreaPlan(doc, scheme.Id, level.Id)
                target_view = created_view
            except Exception as ex:
                forms.alert('Failed to create Area Plan view: {}'.format(ex), exitscript=True)

        # Ensure sketch plane for target view.
        sketch_plane = None
        try:
            sketch_plane = target_view.SketchPlane
        except Exception:
            sketch_plane = None
        if sketch_plane is None:
            plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
            sketch_plane = DB.SketchPlane.Create(doc, plane)

        created_bounds, failed_bounds = _create_area_boundary_lines(doc, target_view, sketch_plane, curves)

        try:
            doc.Regenerate()
        except Exception:
            pass

        for pt, sp in space_points:
            try:
                hp = DB.XYZ(pt.X, pt.Y, elev)
                area = _create_area(doc, target_view, hp)
                if area is None:
                    failed_areas += 1
                    continue

                sp_name, sp_number = _get_space_name_number(sp)
                _set_param_string(area, 'Name', sp_name)
                _set_param_string(area, 'Number', sp_number)

                created_areas += 1

                tag = _create_area_tag(doc, target_view, area, hp, tag_type)
                if tag is None:
                    failed_tags += 1
                else:
                    created_tags += 1
            except Exception:
                failed_areas += 1

    # Switch to area plan after commit
    try:
        revit.uidoc.ActiveView = target_view
    except Exception:
        pass

    output.print_md('## Spaces (Current View) → Area Boundaries + Areas + Area Tags')
    output.print_md('* Source view: `{}`'.format(getattr(view, 'Name', 'Active View')))
    output.print_md('* Target Area Plan: `{}`'.format(getattr(target_view, 'Name', 'Area Plan')))
    output.print_md('* Spaces found in view: `{}`'.format(len(spaces)))
    output.print_md('* Space boundary curves read: `{}`'.format(len(raw_curves)))
    output.print_md('* Area boundary curves created (deduped): `{}` | failed: `{}`'.format(created_bounds, failed_bounds))
    output.print_md('* Skipped (flatten): `{}`'.format(skipped))
    output.print_md('* Skipped (too short): `{}`'.format(skipped_too_short))
    output.print_md('* Areas created: `{}` | failed: `{}`'.format(created_areas, failed_areas))
    output.print_md('* Tags created: `{}` | failed: `{}`'.format(created_tags, failed_tags))
    if tag_type is None:
        output.print_md('> Note: No Area Tag type found in host. Load an Area Tag family/type, then re-run to tag.')


if __name__ == '__main__':
    main()
