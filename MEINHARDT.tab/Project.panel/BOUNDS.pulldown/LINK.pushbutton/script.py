# -*- coding: utf-8 -*-
from __future__ import print_function

from pyrevit import revit, DB, forms, script


logger = script.get_logger()
output = script.get_output()


# Revit internal units are feet.
# Levels are typically separated by several feet; a small tolerance safely filters other levels.
LEVEL_ELEV_TOL_FT = 0.5   # 6 inches

# Avoid creating tiny boundary segments that can trigger "Line is too short" warnings.
MIN_CURVE_LEN_FT = 0.005  # ~1/16 inch


class _LevelItem(object):
    def __init__(self, level, host_elev):
        self.level = level
        self.host_elev = host_elev
        try:
            self.name = '{}  (host Z ≈ {:.3f} ft)'.format(level.Name, host_elev)
        except Exception:
            self.name = 'Level  (host Z ≈ {:.3f} ft)'.format(host_elev)


class _LinkItem(object):
    def __init__(self, link_instance):
        self.link_instance = link_instance
        self.name = getattr(link_instance, 'Name', 'Revit Link')


class _AreaPickItem(object):
    def __init__(self, area):
        self.area = area
        try:
            self.number = getattr(area, 'Number', None)
        except Exception:
            self.number = None
        try:
            self.name = getattr(area, 'Name', None)
        except Exception:
            self.name = None

        if not self.number:
            self.number = _get_param_as_string(area, 'Number')
        if not self.name:
            self.name = _get_param_as_string(area, 'Name')

        parts = []
        if self.number:
            parts.append(str(self.number))
        if self.name:
            parts.append(str(self.name))
        self.display = ' - '.join(parts) if parts else 'Area'


def _get_link_transform(link_instance):
    # Prefer total transform (includes shared coordinates).
    if hasattr(link_instance, 'GetTotalTransform'):
        try:
            return link_instance.GetTotalTransform()
        except Exception:
            pass
    try:
        return link_instance.GetTransform()
    except Exception:
        return DB.Transform.Identity


def _get_levels(doc):
    try:
        return list(DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements())
    except Exception:
        return []


def _pick_link_level_for_host_elevation(link_doc, link_transform, host_level_elev, always_prompt=False):
    levels = _get_levels(link_doc)
    if not levels:
        return None, None

    items = []
    for lvl in levels:
        try:
            p = DB.XYZ(0, 0, lvl.Elevation)
            hp = link_transform.OfPoint(p)
            items.append(_LevelItem(lvl, hp.Z))
        except Exception:
            continue

    if not items:
        return None, None

    # Auto-pick nearest by transformed elevation.
    best = min(items, key=lambda it: abs(it.host_elev - host_level_elev))

    # Let user override if multiple levels exist.
    if always_prompt and len(items) > 1:
        picked = forms.SelectFromList.show(
            items,
            name_attr='name',
            title='Select Linked Level (for filtering)',
            button_name='Use Level',
            multiselect=False
        )
        if picked:
            best = picked
        else:
            return None, None
    elif len(items) > 1:
        picked = forms.SelectFromList.show(
            items,
            name_attr='name',
            title='Select Linked Level (for filtering)',
            button_name='Use Level',
            multiselect=False
        )
        if picked:
            best = picked

    return best.level, best.host_elev


def _curve_key(curve, tol=0.01):
    # Dedup in 2D (XY) with rounding tolerance.
    # tol is in feet (Revit internal units). 0.01 ft ~ 1/8".
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
    if a <= b:
        return (a, b)
    else:
        return (b, a)


def _flatten_curve_to_elevation(curve, elevation, z_tol=1e-4):
    # Space separator lines must lie on the view plane.
    # Boundary segments from Rooms/Areas are expected to be horizontal.
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
        # Fallback: attempt translate by first tessellated point
        try:
            pts = list(curve.Tessellate())
            if not pts:
                return None
            dz = elevation - pts[0].Z
            t = DB.Transform.CreateTranslation(DB.XYZ(0, 0, dz))
            return curve.CreateTransformed(t)
        except Exception:
            return None


def _chunks(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def _collect_boundary_curves_from_spatial_elements(link_doc, use_rooms=True, use_areas=False, level_id=None):
    curves = []
    opts = DB.SpatialElementBoundaryOptions()

    # Prefer Finish boundaries when available.
    if hasattr(DB, 'SpatialElementBoundaryLocation'):
        try:
            opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
        except Exception:
            pass

    if use_rooms:
        room_col = (DB.FilteredElementCollector(link_doc)
                    .OfCategory(DB.BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .ToElements())
        for room in room_col:
            try:
                if level_id is not None:
                    try:
                        if hasattr(room, 'LevelId') and room.LevelId and room.LevelId != level_id:
                            continue
                    except Exception:
                        pass
                if hasattr(room, 'Area') and room.Area <= 0:
                    continue
                seglists = room.GetBoundarySegments(opts)
                if not seglists:
                    continue
                for seglist in seglists:
                    for seg in seglist:
                        try:
                            curves.append(seg.GetCurve())
                        except Exception:
                            try:
                                curves.append(seg.Curve)
                            except Exception:
                                pass
            except Exception as ex:
                logger.debug('Room boundary read failed: %s', ex)

    if use_areas:
        area_col = (DB.FilteredElementCollector(link_doc)
                    .OfCategory(DB.BuiltInCategory.OST_Areas)
                    .WhereElementIsNotElementType()
                    .ToElements())
        for area in area_col:
            try:
                if level_id is not None:
                    try:
                        if hasattr(area, 'LevelId') and area.LevelId and area.LevelId != level_id:
                            continue
                    except Exception:
                        pass
                if hasattr(area, 'Area') and area.Area <= 0:
                    continue
                seglists = area.GetBoundarySegments(opts)
                if not seglists:
                    continue
                for seglist in seglists:
                    for seg in seglist:
                        try:
                            curves.append(seg.GetCurve())
                        except Exception:
                            try:
                                curves.append(seg.Curve)
                            except Exception:
                                pass
            except Exception as ex:
                logger.debug('Area boundary read failed: %s', ex)

    return curves


def _collect_view_owned_elements(doc, view, bic):
    # Some categories don't reliably return results via FilteredElementCollector(doc, view.Id)
    # across versions. Collect globally and filter by OwnerViewId as a robust fallback.
    results = []
    seen = set()

    # First try: view-filtered collector
    try:
        elems = (DB.FilteredElementCollector(doc, view.Id)
                 .OfCategory(bic)
                 .WhereElementIsNotElementType()
                 .ToElements())
        for e in elems:
            try:
                if e.Id.IntegerValue in seen:
                    continue
                seen.add(e.Id.IntegerValue)
                results.append(e)
            except Exception:
                continue
    except Exception:
        pass

    # Second try: global collector + OwnerViewId
    try:
        elems2 = (DB.FilteredElementCollector(doc)
                  .OfCategory(bic)
                  .WhereElementIsNotElementType()
                  .ToElements())
    except Exception:
        elems2 = []

    for e in elems2:
        try:
            if e.Id.IntegerValue in seen:
                continue
            ovid = getattr(e, 'OwnerViewId', None)
            if ovid and ovid.IntegerValue == view.Id.IntegerValue:
                seen.add(e.Id.IntegerValue)
                results.append(e)
        except Exception:
            continue

    return results


def _get_element_linestyle_name(elem):
    # Many line-based elements (ModelCurve/DetailCurve/CurveElement) expose LineStyle.
    try:
        ls = getattr(elem, 'LineStyle', None)
        if ls is not None:
            return getattr(ls, 'Name', None)
    except Exception:
        pass

    # Fallback: try to resolve GraphicsStyle from GraphicsStyleId
    try:
        gs_id = getattr(elem, 'GraphicsStyleId', None)
        if gs_id is not None and hasattr(gs_id, 'IntegerValue') and gs_id.IntegerValue != -1:
            gs = elem.Document.GetElement(gs_id)
            if gs is not None:
                return getattr(gs, 'Name', None)
    except Exception:
        pass

    return None


def _collect_curves_in_view_by_linestyle(doc, view, name_contains):
    # Collect curve-based elements in view and filter by line style name.
    curves = []
    try:
        curve_elems = (DB.FilteredElementCollector(doc, view.Id)
                       .OfClass(DB.CurveElement)
                       .WhereElementIsNotElementType()
                       .ToElements())
    except Exception:
        curve_elems = []

    for e in curve_elems:
        try:
            ls_name = _get_element_linestyle_name(e)
            if not ls_name:
                continue
            if name_contains not in ls_name:
                continue

            c = getattr(e, 'GeometryCurve', None)
            if c is None:
                c = getattr(e, 'Curve', None)
            if c is None and hasattr(e, 'GetCurve'):
                c = e.GetCurve()
            if c is not None:
                curves.append((e, c))
        except Exception:
            continue

    return curves


def _collect_area_boundary_curves_from_view(doc, view):
    # Area boundary lines are drawn as "Area Scheme Lines" in the view.
    curves = []
    elems = _collect_view_owned_elements(doc, view, DB.BuiltInCategory.OST_AreaSchemeLines)

    for e in elems:
        try:
            c = getattr(e, 'GeometryCurve', None)
            if c is None:
                c = getattr(e, 'Curve', None)
            if c is None and hasattr(e, 'GetCurve'):
                c = e.GetCurve()
            if c is not None:
                curves.append(c)
        except Exception:
            continue

    # Fallback: in some projects these present as curve elements with a line style.
    if not curves:
        for e, c in _collect_curves_in_view_by_linestyle(doc, view, 'Area Boundary'):
            try:
                curves.append(c)
            except Exception:
                continue

    return curves


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


def _try_unhide_category(view, bic):
    try:
        cat = view.Document.Settings.Categories.get_Item(bic)
        if cat is None:
            return
        view.SetCategoryHidden(cat.Id, False)
    except Exception:
        pass


def _get_space_separator_ids_in_view(doc, view):
    ids = []
    seen = set()

    # Primary: dedicated category
    elems = _collect_view_owned_elements(doc, view, DB.BuiltInCategory.OST_MEPSpaceSeparationLines)
    for e in elems:
        try:
            if e.Id.IntegerValue in seen:
                continue
            seen.add(e.Id.IntegerValue)
            ids.append(e.Id)
        except Exception:
            continue

    # Fallback: sometimes these show up as curve elements with "Space Separation" line style
    for e, _c in _collect_curves_in_view_by_linestyle(doc, view, 'Space Separation'):
        try:
            if e.Id.IntegerValue in seen:
                continue
            seen.add(e.Id.IntegerValue)
            ids.append(e.Id)
        except Exception:
            continue

    return ids


def _get_room_separator_ids_in_view(doc, view):
    ids = []
    seen = set()

    try:
        elems = _collect_view_owned_elements(doc, view, DB.BuiltInCategory.OST_RoomSeparationLines)
    except Exception:
        elems = []

    for e in elems:
        try:
            if e.Id.IntegerValue in seen:
                continue
            seen.add(e.Id.IntegerValue)
            ids.append(e.Id)
        except Exception:
            continue

    # Fallback: line style name
    for e, _c in _collect_curves_in_view_by_linestyle(doc, view, 'Room Separation'):
        try:
            if e.Id.IntegerValue in seen:
                continue
            seen.add(e.Id.IntegerValue)
            ids.append(e.Id)
        except Exception:
            continue

    return ids


def _get_area_boundary_line_ids_in_view(doc, view):
    ids = []
    seen = set()

    # Primary: area scheme lines (area boundary lines live here)
    elems = _collect_view_owned_elements(doc, view, DB.BuiltInCategory.OST_AreaSchemeLines)
    for e in elems:
        try:
            if e.Id.IntegerValue in seen:
                continue
            seen.add(e.Id.IntegerValue)
            ids.append(e.Id)
        except Exception:
            continue

    # Fallback: curve elements with Area Boundary linestyle
    for e, _c in _collect_curves_in_view_by_linestyle(doc, view, 'Area Boundary'):
        try:
            if e.Id.IntegerValue in seen:
                continue
            seen.add(e.Id.IntegerValue)
            ids.append(e.Id)
        except Exception:
            continue

    return ids


def _get_view_owned_ids_by_category(doc, view, bic):
    ids = []
    try:
        elems = _collect_view_owned_elements(doc, view, bic)
    except Exception:
        elems = []

    for e in elems:
        try:
            ids.append(e.Id)
        except Exception:
            continue

    return ids


def _get_visible_element_ids_in_view_by_category(doc, view, bic):
    # Elements visible in a view are often returned by FilteredElementCollector(doc, view.Id)
    # even when they are not "view-owned" (e.g., Zones).
    try:
        return list(
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElementIds()
        )
    except Exception:
        return []


def _start_over_in_current_view(doc, view):
    # Deletes common generated elements in the current view so the user can rerun safely.
    # Intentionally does NOT delete MEP Spaces themselves (those are model elements).

    to_delete = []

    # Space/room separator lines
    to_delete.extend(_get_space_separator_ids_in_view(doc, view) or [])
    to_delete.extend(_get_room_separator_ids_in_view(doc, view) or [])

    # Area boundary lines in view
    to_delete.extend(_get_area_boundary_line_ids_in_view(doc, view) or [])

    # Tags (view-owned)
    try:
        to_delete.extend(_get_view_owned_ids_by_category(doc, view, DB.BuiltInCategory.OST_MEPSpaceTags) or [])
    except Exception:
        pass
    try:
        to_delete.extend(_get_view_owned_ids_by_category(doc, view, DB.BuiltInCategory.OST_AreaTags) or [])
    except Exception:
        pass
    try:
        to_delete.extend(_get_view_owned_ids_by_category(doc, view, DB.BuiltInCategory.OST_ZoneTags) or [])
    except Exception:
        pass

    # Zones (not view-owned; delete those visible in the view)
    try:
        to_delete.extend(_get_visible_element_ids_in_view_by_category(doc, view, DB.BuiltInCategory.OST_Zone) or [])
    except Exception:
        pass

    # De-dup ids
    uniq = []
    seen = set()
    for eid in to_delete:
        try:
            iv = eid.IntegerValue
        except Exception:
            continue
        if iv in seen:
            continue
        seen.add(iv)
        uniq.append(eid)

    if not uniq:
        forms.alert('Nothing to delete in the current view (no separators/boundaries/zones/tags found).')
        return 0

    action = forms.CommandSwitchWindow.show(
        ['Delete in current view', 'Cancel'],
        message=(
            'START OVER (Current View)\n\n'
            'This will delete {} element(s) visible/owned by the current view, including:\n'
            '- Space Separation lines\n'
            '- Room Separation lines\n'
            '- Area Boundary lines\n'
            '- Space Tags / Area Tags / Zone Tags (if present)\n'
            '- Zones visible in this view\n\n'
            'MEP Spaces themselves are NOT deleted.\n\n'
            'Proceed?'
        ).format(len(uniq))
    )
    if not action or action == 'Cancel':
        return None

    deleted = 0
    try:
        res = doc.Delete(uniq)
        deleted = len(res) if res else 0
    except Exception:
        for eid in uniq:
            try:
                res = doc.Delete(eid)
                if res:
                    deleted += 1
            except Exception:
                pass

    return deleted


def _prompt_delete_existing_space_separators(doc, view):
    ids = list(_get_space_separator_ids_in_view(doc, view) or [])
    if not ids:
        return 0

    action = forms.CommandSwitchWindow.show(
        ['Delete existing in current view', 'Keep existing', 'Cancel'],
        message='Found {} existing Space Separation lines in the current view.\n\nDelete them before creating new ones?'.format(len(ids))
    )
    if not action or action == 'Cancel':
        return None
    if action == 'Keep existing':
        return 0

    deleted = 0
    try:
        res = doc.Delete(ids)
        deleted = len(res) if res else 0
    except Exception:
        # Fallback: try delete one-by-one
        for eid in ids:
            try:
                res = doc.Delete(eid)
                if res:
                    deleted += 1
            except Exception:
                pass

    return deleted


def _get_area_tag_type_from_link(link_doc):
    # Find the most common Area Tag type used in the linked model.
    # NOTE: On some Revit/IronPython combos, DB.AreaTag exists in API but is not
    # exposed through Revit's native object model for OfClass(). Use category filtering.
    type_counts = {}
    try:
        tags = (DB.FilteredElementCollector(link_doc)
                .OfCategory(DB.BuiltInCategory.OST_AreaTags)
                .WhereElementIsNotElementType()
                .ToElements())
    except Exception:
        tags = []

    for t in tags:
        try:
            tid = t.GetTypeId()
            if tid and tid.IntegerValue != -1:
                type_counts[tid] = type_counts.get(tid, 0) + 1
        except Exception:
            continue

    if not type_counts:
        return None

    best_type_id = max(type_counts, key=type_counts.get)
    try:
        return link_doc.GetElement(best_type_id)
    except Exception:
        return None


def _find_matching_area_tag_type_in_host(doc, link_tag_type):
    if link_tag_type is None:
        return None

    try:
        desired_type_name = link_tag_type.Name
    except Exception:
        desired_type_name = None

    desired_family_name = None
    try:
        if hasattr(link_tag_type, 'FamilyName'):
            desired_family_name = link_tag_type.FamilyName
        else:
            fam = getattr(link_tag_type, 'Family', None)
            desired_family_name = fam.Name if fam else None
    except Exception:
        desired_family_name = None

    # Area tag types are element types under OST_AreaTags.
    candidates = (DB.FilteredElementCollector(doc)
                  .OfCategory(DB.BuiltInCategory.OST_AreaTags)
                  .WhereElementIsElementType()
                  .ToElements())

    best = None
    for c in candidates:
        try:
            if desired_type_name and c.Name != desired_type_name:
                continue
            if desired_family_name and hasattr(c, 'FamilyName') and c.FamilyName != desired_family_name:
                continue
            best = c
            break
        except Exception:
            continue

    return best


def _create_area_boundary_lines(doc, view, sketch_plane, curves):
    creator = doc.Create
    created = 0
    failed = 0

    # Try bulk API if available.
    if hasattr(creator, 'NewAreaBoundaryLines'):
        for batch in _chunks(curves, 200):
            ca = DB.CurveArray()
            for c in batch:
                ca.Append(c)
            try:
                creator.NewAreaBoundaryLines(sketch_plane, ca, view)
                created += len(batch)
            except Exception:
                for c in batch:
                    try:
                        creator.NewAreaBoundaryLines(sketch_plane, DB.CurveArray([c]), view)
                        created += 1
                    except Exception:
                        failed += 1
        return created, failed

    # Fall back to per-curve creation.
    if hasattr(creator, 'NewAreaBoundaryLine'):
        for c in curves:
            try:
                creator.NewAreaBoundaryLine(sketch_plane, c, view)
                created += 1
            except Exception:
                failed += 1
        return created, failed

    forms.alert('This Revit version/API does not expose NewAreaBoundaryLine(s).', exitscript=True)


def _create_area(doc, view, point_xyz):
    # Create an Area at point (uses UV in view plane).
    uv = DB.UV(point_xyz.X, point_xyz.Y)
    creator = doc.Create
    if hasattr(creator, 'NewArea'):
        try:
            return creator.NewArea(view, uv)
        except Exception:
            try:
                return creator.NewArea(view.Id, uv)
            except Exception:
                return None
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


def _set_bip_string(elem, bip, value):
    if value is None:
        return False
    try:
        p = elem.get_Parameter(bip)
        if p and not p.IsReadOnly:
            p.Set(value)
            return True
    except Exception:
        return False
    return False


def _set_space_name_number(space, name_value, number_value):
    # Spaces often expose name/number through Room built-in parameters.
    ok_name = _set_param_string(space, 'Name', name_value)
    ok_num = _set_param_string(space, 'Number', number_value)

    if not ok_name:
        ok_name = _set_bip_string(space, DB.BuiltInParameter.ROOM_NAME, name_value)
    if not ok_name:
        ok_name = _set_bip_string(space, DB.BuiltInParameter.ELEM_ROOM_NAME, name_value)

    if not ok_num:
        ok_num = _set_bip_string(space, DB.BuiltInParameter.ROOM_NUMBER, number_value)
    if not ok_num:
        ok_num = _set_bip_string(space, DB.BuiltInParameter.ELEM_ROOM_NUMBER, number_value)

    return ok_name, ok_num


def _get_space_name_number(space):
    # Mirror the setter fallbacks so matching works across families/templates.
    name_value = None
    number_value = None

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


def _get_area_name_number(area):
    name_value = None
    number_value = None

    try:
        name_value = getattr(area, 'Name', None)
    except Exception:
        name_value = None

    try:
        number_value = getattr(area, 'Number', None)
    except Exception:
        number_value = None

    if not name_value:
        name_value = _get_param_as_string(area, 'Name')
    if not number_value:
        number_value = _get_param_as_string(area, 'Number')

    return name_value, number_value


def _collect_areas_in_current_view(doc, view, level=None):
    try:
        areas = (DB.FilteredElementCollector(doc, view.Id)
                 .OfCategory(DB.BuiltInCategory.OST_Areas)
                 .WhereElementIsNotElementType()
                 .ToElements())
    except Exception:
        areas = []

    if level is None:
        return areas

    filtered = []
    for a in areas:
        try:
            if hasattr(a, 'LevelId') and a.LevelId and a.LevelId.IntegerValue != -1:
                if a.LevelId.IntegerValue != level.Id.IntegerValue:
                    continue
            filtered.append(a)
        except Exception:
            filtered.append(a)
    return filtered


def _build_space_lookup_for_level(doc, level=None):
    # Returns multiple lookups for best-effort matching.
    by_num_name = {}
    by_num = {}
    by_name = {}

    try:
        spaces = (DB.FilteredElementCollector(doc)
                  .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
                  .WhereElementIsNotElementType()
                  .ToElements())
    except Exception:
        spaces = []

    for s in spaces:
        try:
            if level is not None and hasattr(s, 'LevelId') and s.LevelId and s.LevelId.IntegerValue != -1:
                if s.LevelId.IntegerValue != level.Id.IntegerValue:
                    continue
        except Exception:
            pass

        name_value, number_value = _get_space_name_number(s)
        if number_value and name_value:
            by_num_name[(str(number_value), str(name_value))] = s
        if number_value and str(number_value) not in by_num:
            by_num[str(number_value)] = s
        if name_value and str(name_value) not in by_name:
            by_name[str(name_value)] = s

    return by_num_name, by_num, by_name


def _make_zone_name(area_name, area_number):
    try:
        if area_name:
            return str(area_name)
    except Exception:
        pass
    return None


def _collect_zones_by_name(doc, phase=None):
    zones = []
    try:
        zones = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_Zone)
                 .WhereElementIsNotElementType()
                 .ToElements())
    except Exception:
        zones = []

    by_name = {}
    for z in zones:
        try:
            if phase is not None:
                try:
                    zph = getattr(z, 'Phase', None)
                    if zph is not None and hasattr(zph, 'Id') and zph.Id.IntegerValue != phase.Id.IntegerValue:
                        continue
                except Exception:
                    pass

            try:
                nm = z.Name
            except Exception:
                nm = _get_param_as_string(z, 'Name')
            if not nm:
                continue
            by_name.setdefault(str(nm), []).append(z)
        except Exception:
            continue
    return by_name


def _get_view_phase(doc, view):
    # Spaces are phase-dependent. Try to use the view phase.
    try:
        p = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
        if p:
            pid = p.AsElementId()
            if pid and pid.IntegerValue != -1:
                ph = doc.GetElement(pid)
                if ph is not None:
                    return ph
    except Exception:
        pass

    # Fallback: last project phase
    try:
        phases = list(doc.Phases)
        if phases:
            return phases[-1]
    except Exception:
        pass
    return None


def _point_key_xy(xyz, tol=0.2):
    # tol in feet; used to avoid creating duplicates at almost-same locations.
    try:
        return (round(xyz.X / tol) * tol, round(xyz.Y / tol) * tol)
    except Exception:
        return None


def _create_space(doc, level, phase, point_xyz):
    uv = DB.UV(point_xyz.X, point_xyz.Y)
    creator = doc.Create

    # Try common signatures across versions.
    for args in (
        (level, uv),
        (level.Id, uv),
        (phase, uv, level),
        (phase.Id, uv, level.Id),
    ):
        try:
            if hasattr(creator, 'NewSpace'):
                sp = creator.NewSpace(*args)
                if sp is not None:
                    return sp
        except Exception:
            continue

    return None


def _create_space_tag(doc, view, space, point_xyz, tag_type):
    uv = DB.UV(point_xyz.X, point_xyz.Y)
    creator = doc.Create
    tag = None

    # Try creator.NewSpaceTag variants
    for args in (
        # Per Revit API: NewSpaceTag(space, point(UV), view)
        (space, uv, view),
        (uv, space, view),
        (view, space, uv),
        (view.Id, space.Id, uv),
        (uv, space.Id, view.Id),
    ):
        try:
            if hasattr(creator, 'NewSpaceTag'):
                tag = creator.NewSpaceTag(*args)
                break
        except Exception:
            continue

    # Fallback: SpatialElementTag.Create
    if tag is None:
        try:
            if hasattr(DB, 'SpatialElementTag') and hasattr(DB.SpatialElementTag, 'Create'):
                tag = DB.SpatialElementTag.Create(doc, view.Id, space.Id, uv)
        except Exception:
            tag = None

    # Fallback: IndependentTag.Create
    if tag is None:
        try:
            if hasattr(DB, 'IndependentTag') and hasattr(DB.IndependentTag, 'Create'):
                ref = DB.Reference(space)
                tag = DB.IndependentTag.Create(
                    doc,
                    view.Id,
                    ref,
                    True,
                    DB.TagMode.TM_ADDBY_CATEGORY,
                    DB.TagOrientation.Horizontal,
                    DB.XYZ(point_xyz.X, point_xyz.Y, point_xyz.Z)
                )
        except Exception:
            tag = None

    if tag is not None and tag_type is not None:
        try:
            tag.ChangeTypeId(tag_type.Id)
        except Exception:
            try:
                if hasattr(tag, 'SpaceTagType'):
                    tag.SpaceTagType = tag_type
            except Exception:
                pass

    return tag


def _get_default_space_tag_type(doc):
    try:
        types = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_MEPSpaceTags)
                 .WhereElementIsElementType()
                 .ToElements())
        return types[0] if types else None
    except Exception:
        return None


def _create_area_tag(doc, view, area, point_xyz, tag_type):
    uv = DB.UV(point_xyz.X, point_xyz.Y)
    creator = doc.Create

    tag = None
    # Try a few common signatures.
    for args in (
        (uv, area, view),
        (view, area, uv),
        (view.Id, area.Id, uv),
        (uv, area.Id, view.Id),
    ):
        try:
            if hasattr(creator, 'NewAreaTag'):
                tag = creator.NewAreaTag(*args)
                break
        except Exception:
            continue

    # Fallback: try SpatialElementTag.Create (some versions).
    if tag is None:
        try:
            if hasattr(DB, 'SpatialElementTag') and hasattr(DB.SpatialElementTag, 'Create'):
                # Create needs a LinkElementId for linked tags; for host areas try plain.
                tag = DB.SpatialElementTag.Create(doc, view.Id, area.Id, uv)
        except Exception:
            tag = None

    if tag is not None and tag_type is not None:
        try:
            tag.ChangeTypeId(tag_type.Id)
        except Exception:
            try:
                if hasattr(tag, 'AreaTagType'):
                    tag.AreaTagType = tag_type
            except Exception:
                pass

    return tag


def main():
    doc = revit.doc
    view = revit.active_view

    if view is None or view.IsTemplate:
        forms.alert('Open a plan view (not a template) and try again.', exitscript=True)

    # Some Revit plan-like views may not pass isinstance(ViewPlan) checks consistently across versions.
    # Validate by view type instead.
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

    mode = forms.CommandSwitchWindow.show(
        [
            'Start Over in Current View (Delete Separators, Boundaries, Zones, Tags)',
            'Create HVAC Zones (from Selected Areas in Current View)',
            'Create Space Boundaries + Spaces + Space Tags (from Linked Rooms)',
            'Create Area Boundaries + Areas + Area Tags (from Linked Rooms)',
            'Create Areas + Area Boundaries + Tags (from Link)',
            'Create Space Separators (from Link)',
            'Create Space Separators (from Area Boundaries in Current View)',
            'Create Spaces + Space Tags (from Link Areas)',
        ],
        message='What do you want to create in the host model?'
    )
    if not mode:
        return

    if mode.startswith('Start Over in Current View'):
        with revit.Transaction('Start over in current view'):
            deleted = _start_over_in_current_view(doc, view)
            if deleted is None:
                return
        forms.alert('Deleted {} element(s) in the current view.'.format(deleted))
        return

    if mode == 'Create HVAC Zones (from Selected Areas in Current View)':
        phase = _get_view_phase(doc, view)
        if phase is None:
            forms.alert('Could not determine a phase for this view/project. Zones are phase-dependent.', exitscript=True)

        areas = _collect_areas_in_current_view(doc, view, level=level)
        if not areas:
            forms.alert('No Areas found in the current view.', exitscript=True)

        items = [_AreaPickItem(a) for a in areas]
        picked = forms.SelectFromList.show(
            items,
            name_attr='display',
            title='Select Areas to Create HVAC Zones',
            button_name='Create Zones',
            multiselect=True
        )
        if not picked:
            return

        by_num_name, by_num, by_name = _build_space_lookup_for_level(doc, level=level)
        existing_zones_by_name = _collect_zones_by_name(doc, phase=phase)

        # Precompute matches and intended names so we can prompt once.
        rows = []
        missing_space = 0
        already_zoned = 0
        name_conflicts = set()
        duplicate_selected_names = 0
        seen_selected_names = set()

        for it in picked:
            area = it.area
            area_name, area_number = _get_area_name_number(area)
            zone_name = _make_zone_name(area_name, area_number)
            if not zone_name:
                zone_name = 'Zone'

            if zone_name in seen_selected_names:
                duplicate_selected_names += 1
            else:
                seen_selected_names.add(zone_name)

            if zone_name in existing_zones_by_name:
                name_conflicts.add(zone_name)

            key = None
            if area_number and area_name:
                key = (str(area_number), str(area_name))

            space = None
            if key and key in by_num_name:
                space = by_num_name.get(key)
            if space is None and area_number:
                space = by_num.get(str(area_number))
            if space is None and area_name:
                space = by_name.get(str(area_name))

            if space is None:
                missing_space += 1
            else:
                try:
                    z = getattr(space, 'Zone', None)
                    if z is not None:
                        try:
                            # Treat default zone as "not really assigned".
                            if hasattr(z, 'IsDefaultZone') and z.IsDefaultZone:
                                pass
                            else:
                                already_zoned += 1
                        except Exception:
                            already_zoned += 1
                except Exception:
                    pass

            rows.append({
                'area': area,
                'area_name': area_name,
                'area_number': area_number,
                'zone_name': zone_name,
                'space': space,
            })

        # If zones with same name already exist, ask what to do.
        conflict_policy = 'Keep existing (skip)'
        if name_conflicts:
            conflict_policy = forms.CommandSwitchWindow.show(
                ['Keep existing (skip)', 'Delete existing and recreate', 'Cancel'],
                message=(
                    'Found existing Zone(s) with the same name as {} selected Area(s).\n\n'
                    'How should duplicates be handled?\n\n'
                    'Note: Default Zone(s) will never be deleted.'
                ).format(len(name_conflicts))
            )
            if not conflict_policy or conflict_policy == 'Cancel':
                return

        missing_policy = 'Skip'
        if missing_space:
            missing_policy = forms.CommandSwitchWindow.show(
                ['Skip', 'Create zone anyway (no spaces added)', 'Cancel'],
                message=(
                    '{} selected Area(s) have no matching Space on this level (by Number/Name).\n\n'
                    'Zones are defined by Spaces. What should I do for those Areas?'
                ).format(missing_space)
            )
            if not missing_policy or missing_policy == 'Cancel':
                return

        zoned_space_policy = 'Skip'
        if already_zoned:
            zoned_space_policy = forms.CommandSwitchWindow.show(
                ['Skip', 'Move Space to new Zone', 'Cancel'],
                message=(
                    '{} matching Space(s) are already assigned to a Zone.\n\n'
                    'To put them in the new Zone, they must be removed from their current Zone first.'
                ).format(already_zoned)
            )
            if not zoned_space_policy or zoned_space_policy == 'Cancel':
                return

        created = 0
        created_empty = 0
        skipped_existing_name = 0
        skipped_duplicate_selected_name = 0
        skipped_no_space = 0
        skipped_space_already_zoned = 0
        moved_spaces = 0
        failed = 0

        with revit.Transaction('Create HVAC Zones from Areas (current view)'):
            creator = doc.Create
            if not hasattr(creator, 'NewZone'):
                forms.alert('This Revit version/API does not expose NewZone.', exitscript=True)

            # Delete conflicting existing zones up-front if requested.
            if conflict_policy == 'Delete existing and recreate' and name_conflicts:
                ids_to_delete = []
                for nm in name_conflicts:
                    for z in existing_zones_by_name.get(nm, []):
                        try:
                            if hasattr(z, 'IsDefaultZone') and z.IsDefaultZone:
                                continue
                        except Exception:
                            pass
                        try:
                            ids_to_delete.append(z.Id)
                        except Exception:
                            continue
                if ids_to_delete:
                    try:
                        doc.Delete(ids_to_delete)
                    except Exception:
                        for zid in ids_to_delete:
                            try:
                                doc.Delete(zid)
                            except Exception:
                                pass

                # Refresh mapping after deletion attempt.
                existing_zones_by_name = _collect_zones_by_name(doc, phase=phase)

            seen_created_names = set()

            for r in rows:
                try:
                    zone_name = r['zone_name']
                    space = r['space']

                    # Prevent duplicates within this run
                    if zone_name in seen_created_names:
                        skipped_duplicate_selected_name += 1
                        continue

                    if zone_name in existing_zones_by_name and conflict_policy == 'Keep existing (skip)':
                        skipped_existing_name += 1
                        continue

                    if space is None:
                        if missing_policy == 'Skip':
                            skipped_no_space += 1
                            continue

                    # Create zone
                    zone = creator.NewZone(level, phase)
                    try:
                        zone.Name = zone_name
                    except Exception:
                        pass

                    # If no space, keep empty zone
                    if space is None:
                        created_empty += 1
                        created += 1
                        seen_created_names.add(zone_name)
                        continue

                    # If space is already in another zone, handle based on policy
                    old_zone = None
                    try:
                        old_zone = getattr(space, 'Zone', None)
                    except Exception:
                        old_zone = None

                    if old_zone is not None:
                        try:
                            if hasattr(old_zone, 'IsDefaultZone') and old_zone.IsDefaultZone:
                                old_zone = None
                        except Exception:
                            pass

                    if old_zone is not None:
                        if zoned_space_policy == 'Skip':
                            skipped_space_already_zoned += 1
                            # If we created a zone for this area, delete it to avoid empty duplicates
                            try:
                                doc.Delete(zone.Id)
                            except Exception:
                                pass
                            continue
                        else:
                            try:
                                ss_old = DB.Mechanical.SpaceSet()
                                ss_old.Insert(space)
                                old_zone.RemoveSpaces(ss_old)
                                moved_spaces += 1
                            except Exception:
                                # If we can't remove, we can't reliably add
                                try:
                                    doc.Delete(zone.Id)
                                except Exception:
                                    pass
                                failed += 1
                                continue

                    ss = DB.Mechanical.SpaceSet()
                    try:
                        ss.Insert(space)
                    except Exception:
                        pass

                    ok = False
                    try:
                        ok = zone.AddSpaces(ss)
                    except Exception:
                        ok = False

                    if ok:
                        created += 1
                        seen_created_names.add(zone_name)
                    else:
                        failed += 1
                        try:
                            doc.Delete(zone.Id)
                        except Exception:
                            pass

                except Exception as ex:
                    logger.debug('Zone creation failed: %s', ex)
                    failed += 1

        forms.alert(
            'HVAC Zones (from current view Areas) complete:\n\n'
            'Selected Areas: {}\n'
            'Zones created: {} (empty: {})\n'
            'Skipped (existing zone name): {}\n'
            'Skipped (duplicate zone name in selection): {}\n'
            'Skipped (no matching Space): {}\n'
            'Skipped (Space already zoned): {}\n'
            'Spaces moved from old zone: {}\n'
            'Failed: {}'.format(
                len(picked),
                created,
                created_empty,
                skipped_existing_name,
                skipped_duplicate_selected_name,
                skipped_no_space,
                skipped_space_already_zoned,
                moved_spaces,
                failed
            )
        )
        return

    # CURRENT VIEW AREA BOUNDARIES -> SPACE SEPARATORS
    if 'Area Boundaries in Current View' in mode:
        elev = level.Elevation

        raw_curves = _collect_area_boundary_curves_from_view(doc, view)
        if not raw_curves:
            forms.alert('No Area Boundary lines found in the current view.', exitscript=True)

        flattened = []
        skipped = 0
        skipped_too_short = 0
        for c in raw_curves:
            if c is None:
                continue

            # Skip tiny segments.
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

        # Dedup segments
        uniq = {}
        for c in flattened:
            k = _curve_key(c)
            if k is None:
                continue
            uniq[k] = c

        curves = list(uniq.values())
        if not curves:
            forms.alert('No usable Area Boundary curves found in the current view (after filtering).', exitscript=True)

        # Ensure we have a sketch plane.
        sketch_plane = None
        try:
            sketch_plane = view.SketchPlane
        except Exception:
            sketch_plane = None

        if sketch_plane is None:
            plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
            sketch_plane = DB.SketchPlane.Create(doc, plane)

        created = 0
        failed = 0
        with revit.Transaction('Area boundaries (current view) to space separators'):
            creator = doc.Create
            if not hasattr(creator, 'NewSpaceBoundaryLines'):
                forms.alert('This Revit version/API does not expose NewSpaceBoundaryLines.', exitscript=True)

            deleted_existing = _prompt_delete_existing_space_separators(doc, view)
            if deleted_existing is None:
                return

            for batch in _chunks(curves, 200):
                ca = DB.CurveArray()
                for cc in batch:
                    ca.Append(cc)
                try:
                    creator.NewSpaceBoundaryLines(sketch_plane, ca, view)
                    created += len(batch)
                except Exception as ex:
                    logger.debug('Batch creation failed, falling back per-curve: %s', ex)
                    for cc in batch:
                        ca2 = DB.CurveArray()
                        ca2.Append(cc)
                        try:
                            creator.NewSpaceBoundaryLines(sketch_plane, ca2, view)
                            created += 1
                        except Exception:
                            failed += 1

            _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceSeparationLines)

        output.print_md('## Current View Area Boundaries → Space Separators')
        output.print_md('* View: `{}`'.format(getattr(view, 'Name', 'Active View')))
        output.print_md('* Existing space separators deleted: `{}`'.format(deleted_existing))
        output.print_md('* Area boundary curves read: `{}`'.format(len(raw_curves)))
        output.print_md('* Space separator curves created (deduped): `{}`'.format(created))
        output.print_md('* Skipped (flatten): `{}`'.format(skipped))
        output.print_md('* Skipped (too short): `{}`'.format(skipped_too_short))
        output.print_md('* Failed to create: `{}`'.format(failed))
        return

    # LINK-BASED MODES
    link_instances = [li for li in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()
                      if li.GetLinkDocument() is not None]

    if not link_instances:
        forms.alert('No loaded Revit links found in this model.', exitscript=True)

    link_items = [_LinkItem(li) for li in link_instances]
    selected = forms.SelectFromList.show(
        link_items,
        name_attr='name',
        title='Select Revit Link',
        button_name='Use Link',
        multiselect=False
    )

    if not selected:
        return

    link_instance = selected.link_instance
    link_doc = link_instance.GetLinkDocument()
    if link_doc is None:
        forms.alert('Selected link is not loaded.', exitscript=True)

    # Special mode: Space boundaries + spaces + tags from linked Rooms
    if mode == 'Create Space Boundaries + Spaces + Space Tags (from Linked Rooms)':
        link_t = _get_link_transform(link_instance)
        elev = level.Elevation

        # Pick the linked level that corresponds to this host level.
        link_level, link_level_host_elev = _pick_link_level_for_host_elevation(link_doc, link_t, elev)
        filter_elev = link_level_host_elev if link_level_host_elev is not None else elev

        # Collect linked rooms + placement points
        room_points = []  # list of (XYZ point, linked_room)
        for r in (DB.FilteredElementCollector(link_doc)
                  .OfCategory(DB.BuiltInCategory.OST_Rooms)
                  .WhereElementIsNotElementType()
                  .ToElements()):
            try:
                if hasattr(r, 'Area') and r.Area <= 0:
                    continue

                if link_level is not None:
                    try:
                        if hasattr(r, 'LevelId') and r.LevelId and r.LevelId != link_level.Id:
                            continue
                    except Exception:
                        pass

                lp = getattr(r, 'Location', None)
                if lp is None:
                    continue
                pt = getattr(lp, 'Point', None)
                if pt is None:
                    continue
                room_points.append((pt, r))
            except Exception:
                continue

        if not room_points:
            forms.alert('No placed Rooms found in the selected link (after level filtering).', exitscript=True)

        # Collect linked room boundary curves
        raw_curves = _collect_boundary_curves_from_spatial_elements(
            link_doc,
            use_rooms=True,
            use_areas=False,
            level_id=(link_level.Id if link_level is not None else None)
        )
        if not raw_curves:
            forms.alert('No Room boundary curves found in the selected link (after level filtering).', exitscript=True)

        transformed = []
        skipped = 0
        skipped_too_short = 0
        for c in raw_curves:
            if c is None:
                continue
            try:
                host_c = c.CreateTransformed(link_t)
            except Exception:
                skipped += 1
                continue

            try:
                if hasattr(host_c, 'Length') and host_c.Length <= MIN_CURVE_LEN_FT:
                    skipped_too_short += 1
                    continue
            except Exception:
                pass

            flat = _flatten_curve_to_elevation(host_c, elev)
            if flat is None:
                skipped += 1
                continue
            transformed.append(flat)

        uniq = {}
        for c in transformed:
            k = _curve_key(c)
            if k is None:
                continue
            uniq[k] = c
        curves = list(uniq.values())
        if not curves:
            forms.alert('No usable Room boundary curves found in the selected link (after filtering).', exitscript=True)

        # Phase for spaces
        phase = _get_view_phase(doc, view)
        if phase is None:
            forms.alert('Could not determine a Phase for Space creation.', exitscript=True)

        # Avoid duplicate spaces by XY key
        existing_keys = set()
        try:
            existing_spaces = (DB.FilteredElementCollector(doc)
                               .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
                               .WhereElementIsNotElementType()
                               .ToElements())
        except Exception:
            existing_spaces = []

        for sp in existing_spaces:
            try:
                if hasattr(sp, 'LevelId') and sp.LevelId and sp.LevelId != level.Id:
                    continue
                loc = getattr(sp, 'Location', None)
                pt = getattr(loc, 'Point', None)
                if pt is None:
                    continue
                k = _point_key_xy(pt)
                if k is not None:
                    existing_keys.add(k)
            except Exception:
                continue

        tag_type = _get_default_space_tag_type(doc)

        created_bounds = 0
        failed_bounds = 0
        created_spaces = 0
        failed_spaces = 0
        skipped_existing = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Create Space Boundaries + Spaces + Tags from linked Rooms'):
            creator = doc.Create
            if not hasattr(creator, 'NewSpaceBoundaryLines'):
                forms.alert('This Revit version/API does not expose NewSpaceBoundaryLines.', exitscript=True)

            # Ensure sketch plane
            sketch_plane = None
            try:
                sketch_plane = view.SketchPlane
            except Exception:
                sketch_plane = None
            if sketch_plane is None:
                plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
                sketch_plane = DB.SketchPlane.Create(doc, plane)

            deleted_existing = _prompt_delete_existing_space_separators(doc, view)
            if deleted_existing is None:
                return

            for batch in _chunks(curves, 200):
                ca = DB.CurveArray()
                for cc in batch:
                    ca.Append(cc)
                try:
                    creator.NewSpaceBoundaryLines(sketch_plane, ca, view)
                    created_bounds += len(batch)
                except Exception as ex:
                    logger.debug('Batch boundary creation failed, falling back per-curve: %s', ex)
                    for cc in batch:
                        ca2 = DB.CurveArray()
                        ca2.Append(cc)
                        try:
                            creator.NewSpaceBoundaryLines(sketch_plane, ca2, view)
                            created_bounds += 1
                        except Exception:
                            failed_bounds += 1

            try:
                doc.Regenerate()
            except Exception:
                pass

            for p, src_room in room_points:
                try:
                    hp = link_t.OfPoint(p)
                    if link_level is None and abs(hp.Z - filter_elev) > LEVEL_ELEV_TOL_FT:
                        continue
                    hp = DB.XYZ(hp.X, hp.Y, elev)

                    k = _point_key_xy(hp)
                    if k is not None and k in existing_keys:
                        skipped_existing += 1
                        continue

                    sp = _create_space(doc, level, phase, hp)
                    if sp is None:
                        failed_spaces += 1
                        continue

                    # Copy name/number from linked room.
                    try:
                        src_name = src_room.Name
                    except Exception:
                        src_name = _get_param_as_string(src_room, 'Name')
                    try:
                        src_number = src_room.Number
                    except Exception:
                        src_number = _get_param_as_string(src_room, 'Number')

                    _set_space_name_number(sp, src_name, src_number)
                    created_spaces += 1
                    if k is not None:
                        existing_keys.add(k)

                    tag = _create_space_tag(doc, view, sp, hp, tag_type)
                    if tag is None:
                        failed_tags += 1
                    else:
                        created_tags += 1
                except Exception:
                    failed_spaces += 1

            _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceSeparationLines)
            _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceTags)

        output.print_md('## Linked Rooms → Space Boundaries + Spaces + Space Tags')
        output.print_md('* Link: `{}`'.format(getattr(link_instance, 'Name', 'Revit Link')))
        if link_level is not None:
            try:
                output.print_md('* Linked level filter: `{}` (host Z ≈ `{:.3f} ft`)'.format(link_level.Name, filter_elev))
            except Exception:
                output.print_md('* Linked level filter applied')
        output.print_md('* Rooms processed: `{}`'.format(len(room_points)))
        output.print_md('* Room boundary curves read: `{}`'.format(len(raw_curves)))
        output.print_md('* Space boundary curves created (deduped): `{}` | failed: `{}`'.format(created_bounds, failed_bounds))
        output.print_md('* Skipped (transform/flatten): `{}`'.format(skipped))
        output.print_md('* Skipped (too short): `{}`'.format(skipped_too_short))
        output.print_md('* Spaces created: `{}` | failed: `{}` | skipped existing: `{}`'.format(created_spaces, failed_spaces, skipped_existing))
        output.print_md('* Tags created: `{}` | failed: `{}`'.format(created_tags, failed_tags))
        if tag_type is None:
            output.print_md('> Note: No Space Tag type found in host. Load a Space Tag family/type, then re-run to tag.')
        return

    # Special mode: Spaces + tags from linked Areas
    if mode == 'Create Spaces + Space Tags (from Link Areas)':
        link_t = _get_link_transform(link_instance)
        elev = level.Elevation

        # Pick the linked level that corresponds to this host level.
        link_level, link_level_host_elev = _pick_link_level_for_host_elevation(link_doc, link_t, elev)
        filter_elev = link_level_host_elev if link_level_host_elev is not None else elev

        # Collect linked areas + points
        area_points = []  # list of (XYZ point, linked_area)
        for a in (DB.FilteredElementCollector(link_doc)
                  .OfCategory(DB.BuiltInCategory.OST_Areas)
                  .WhereElementIsNotElementType()
                  .ToElements()):
            try:
                if hasattr(a, 'Area') and a.Area <= 0:
                    continue
                if link_level is not None:
                    try:
                        if hasattr(a, 'LevelId') and a.LevelId and a.LevelId != link_level.Id:
                            continue
                    except Exception:
                        pass
                lp = a.Location
                if lp is None:
                    continue
                pt = getattr(lp, 'Point', None)
                if pt is None:
                    continue
                area_points.append((pt, a))
            except Exception:
                continue

        if not area_points:
            forms.alert('No placed Areas found in the selected link (after level filtering).', exitscript=True)

        phase = _get_view_phase(doc, view)
        if phase is None:
            forms.alert('Could not determine a Phase for Space creation.', exitscript=True)

        # Avoid duplicates
        existing_keys = set()
        try:
            existing_spaces = (DB.FilteredElementCollector(doc)
                               .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
                               .WhereElementIsNotElementType()
                               .ToElements())
        except Exception:
            existing_spaces = []

        for sp in existing_spaces:
            try:
                if hasattr(sp, 'LevelId') and sp.LevelId and sp.LevelId != level.Id:
                    continue
                loc = getattr(sp, 'Location', None)
                pt = getattr(loc, 'Point', None)
                if pt is None:
                    continue
                k = _point_key_xy(pt)
                if k is not None:
                    existing_keys.add(k)
            except Exception:
                continue

        tag_type = _get_default_space_tag_type(doc)

        created_spaces = 0
        failed_spaces = 0
        skipped_existing = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Create Spaces + Tags from linked Areas'):
            for p, src_area in area_points:
                try:
                    hp = link_t.OfPoint(p)
                    if link_level is None and abs(hp.Z - filter_elev) > LEVEL_ELEV_TOL_FT:
                        continue
                    hp = DB.XYZ(hp.X, hp.Y, elev)

                    k = _point_key_xy(hp)
                    if k is not None and k in existing_keys:
                        skipped_existing += 1
                        continue

                    sp = _create_space(doc, level, phase, hp)
                    if sp is None:
                        failed_spaces += 1
                        continue

                    # Copy name/number from linked area.
                    src_name = _get_param_as_string(src_area, 'Name')
                    src_number = _get_param_as_string(src_area, 'Number')
                    _set_space_name_number(sp, src_name, src_number)

                    created_spaces += 1
                    if k is not None:
                        existing_keys.add(k)

                    tag = _create_space_tag(doc, view, sp, hp, tag_type)
                    if tag is None:
                        failed_tags += 1
                    else:
                        created_tags += 1
                except Exception:
                    failed_spaces += 1

            # Attempt to unhide space tags in this view (a view template can still override).
            _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceTags)

        output.print_md('## Linked Areas → Spaces + Space Tags')
        output.print_md('* Link: `{}`'.format(getattr(link_instance, 'Name', 'Revit Link')))
        if link_level is not None:
            try:
                output.print_md('* Linked level filter: `{}` (host Z ≈ `{:.3f} ft`)'.format(link_level.Name, filter_elev))
            except Exception:
                output.print_md('* Linked level filter applied')
        output.print_md('* Areas processed: `{}`'.format(len(area_points)))
        output.print_md('* Spaces created: `{}` | failed: `{}` | skipped existing: `{}`'.format(created_spaces, failed_spaces, skipped_existing))
        output.print_md('* Tags created: `{}` | failed: `{}`'.format(created_tags, failed_tags))
        if tag_type is None:
            output.print_md('> Note: No Space Tag type found in host. Load a Space Tag family/type, then re-run to tag.')
        output.print_md('> Note: Spaces require enclosing boundaries (walls room-bounding and/or Space Separation lines). If many spaces failed, verify boundaries on this level.')
        return

    source_choice = None
    if mode == 'Create Area Boundaries + Areas + Area Tags (from Linked Rooms)':
        source_choice = 'Rooms'
    else:
        source_choice = forms.CommandSwitchWindow.show(
            ['Rooms', 'Areas', 'Rooms + Areas'],
            message='Use boundaries from which linked elements?'
        )
        if not source_choice:
            return

    use_rooms = source_choice in ('Rooms', 'Rooms + Areas')
    use_areas = source_choice in ('Areas', 'Rooms + Areas')

    link_t = _get_link_transform(link_instance)
    elev = level.Elevation

    # Pick the linked level that corresponds to this host level.
    # This avoids flattening other linked levels onto the current host level.
    link_level, link_level_host_elev = _pick_link_level_for_host_elevation(link_doc, link_t, elev)
    # If we can't determine linked levels, fall back to host level elevation filtering.
    filter_elev = link_level_host_elev if link_level_host_elev is not None else elev

    def _collect_link_curves(level_id):
        return _collect_boundary_curves_from_spatial_elements(
            link_doc,
            use_rooms=use_rooms,
            use_areas=use_areas,
            level_id=level_id
        )

    raw_curves = _collect_link_curves(link_level.Id if link_level is not None else None)

    if not raw_curves:
        # Check if curves exist at all (unfiltered). If yes, level mapping is likely wrong.
        raw_any = _collect_link_curves(None)

        if raw_any:
            msg = (
                'No boundary curves found after level filtering.\n\n'
                'This usually means the linked Level chosen for this host Level is not the one that contains the Rooms/Areas you want.\n\n'
                'Choose an action:'
            )
            action = forms.CommandSwitchWindow.show(
                ['Pick a different linked level', 'Proceed without level filter', 'Cancel'],
                message=msg
            )
            if not action or action == 'Cancel':
                return

            if action == 'Pick a different linked level':
                link_level, link_level_host_elev = _pick_link_level_for_host_elevation(
                    link_doc, link_t, elev, always_prompt=True
                )
                if link_level is None:
                    return
                filter_elev = link_level_host_elev if link_level_host_elev is not None else elev
                raw_curves = _collect_link_curves(link_level.Id)
            else:
                # Proceed without filtering (may include other linked levels).
                link_level = None
                link_level_host_elev = None
                filter_elev = elev
                raw_curves = raw_any

        if not raw_curves:
            forms.alert(
                'No boundary curves found in the selected link.\n\n'
                'Common causes:\n'
                '- You picked Rooms but the link only has Areas (or vice-versa).\n'
                '- Rooms/Areas in the link are unplaced or not enclosed (no boundaries).\n'
                '- The link is not the model you expect.\n',
                exitscript=True
            )

    transformed = []
    skipped = 0
    skipped_too_short = 0

    for c in raw_curves:
        if c is None:
            continue
        try:
            host_c = c.CreateTransformed(link_t)
        except Exception:
            skipped += 1
            continue

        # NOTE: Do not filter by curve Z.
        # For Areas (especially in non-shared links), boundary curves can report an
        # unexpected elevation even when they belong to the selected level.
        # Level filtering is handled by linked element LevelId during collection.

        # Skip tiny segments.
        try:
            if hasattr(host_c, 'Length') and host_c.Length <= MIN_CURVE_LEN_FT:
                skipped_too_short += 1
                continue
        except Exception:
            pass

        # Always flatten to the host target level elevation.
        flat = _flatten_curve_to_elevation(host_c, elev)
        if flat is None:
            skipped += 1
            continue

        transformed.append(flat)

    # Dedup segments (shared boundaries show up twice).
    uniq = {}
    for c in transformed:
        k = _curve_key(c)
        if k is None:
            continue
        uniq[k] = c

    curves = list(uniq.values())

    # Ensure we have a sketch plane.
    sketch_plane = None
    try:
        sketch_plane = view.SketchPlane
    except Exception:
        sketch_plane = None

    if sketch_plane is None:
        plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
        sketch_plane = DB.SketchPlane.Create(doc, plane)

    created = 0
    failed = 0

    # AREA MODE: boundaries + areas + tags (works best in Area Plan view).
    if mode == 'Create Areas + Area Boundaries + Tags (from Link)':
        target_view = view
        created_view = None
        if view.ViewType != DB.ViewType.AreaPlan:
            # Prefer existing Area Plan view; if none, we'll create one inside the transaction.
            picked_area_view = _pick_area_plan_view(doc, level)
            if picked_area_view:
                target_view = picked_area_view
            else:
                scheme = _pick_area_scheme(doc)
                if not scheme:
                    return

        # Determine tag type to match link (by name/family), if available in host.
        link_tag_type = _get_area_tag_type_from_link(link_doc)
        host_tag_type = _find_matching_area_tag_type_in_host(doc, link_tag_type)

        # Collect placement points + source elements (Areas and/or Rooms)
        area_points = []  # list of (XYZ point, src_element)
        if use_areas:
            for a in (DB.FilteredElementCollector(link_doc)
                      .OfCategory(DB.BuiltInCategory.OST_Areas)
                      .WhereElementIsNotElementType()
                      .ToElements()):
                try:
                    if hasattr(a, 'Area') and a.Area <= 0:
                        continue

                    if link_level is not None:
                        try:
                            if hasattr(a, 'LevelId') and a.LevelId and a.LevelId != link_level.Id:
                                continue
                        except Exception:
                            pass

                    lp = a.Location
                    if lp is None:
                        continue
                    pt = getattr(lp, 'Point', None)
                    if pt is None:
                        continue
                    area_points.append((pt, a))
                except Exception:
                    continue

        if use_rooms:
            for r in (DB.FilteredElementCollector(link_doc)
                      .OfCategory(DB.BuiltInCategory.OST_Rooms)
                      .WhereElementIsNotElementType()
                      .ToElements()):
                try:
                    if hasattr(r, 'Area') and r.Area <= 0:
                        continue

                    if link_level is not None:
                        try:
                            if hasattr(r, 'LevelId') and r.LevelId and r.LevelId != link_level.Id:
                                continue
                        except Exception:
                            pass

                    lp = getattr(r, 'Location', None)
                    if lp is None:
                        continue
                    pt = getattr(lp, 'Point', None)
                    if pt is None:
                        continue
                    area_points.append((pt, r))
                except Exception:
                    continue

        # De-dup points (linked Rooms+Areas can overlap)
        dedup_points = []
        seen_pt = set()
        for p, se in area_points:
            try:
                k = _point_key_xy(p)
            except Exception:
                k = None
            if k is not None:
                if k in seen_pt:
                    continue
                seen_pt.add(k)
            dedup_points.append((p, se))
        area_points = dedup_points

        created_areas = 0
        failed_areas = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Link boundaries to area boundaries + tags'):
            # If we're not in an Area Plan view yet and none was picked, create it now.
            if target_view.ViewType != DB.ViewType.AreaPlan:
                try:
                    created_view = DB.ViewPlan.CreateAreaPlan(doc, scheme.Id, level.Id)
                    target_view = created_view
                except Exception as ex:
                    forms.alert('Failed to create Area Plan view: {}'.format(ex), exitscript=True)

            # Ensure sketch plane for target view.
            try:
                sketch_plane = target_view.SketchPlane
            except Exception:
                sketch_plane = None
            if sketch_plane is None:
                plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
                sketch_plane = DB.SketchPlane.Create(doc, plane)

            # Create Area Boundary Lines
            created, failed = _create_area_boundary_lines(doc, target_view, sketch_plane, curves)

            # IMPORTANT: Revit needs a regen after sketch/area boundary creation
            # before NewArea can properly find enclosed regions.
            try:
                doc.Regenerate()
            except Exception:
                pass

            # Create Areas at linked element points
            link_t = _get_link_transform(link_instance)
            for p, src_area in area_points:
                try:
                    hp = link_t.OfPoint(p)
                    # Filter out Areas from other linked levels.
                    # Primary filter is linked element LevelId; only fall back to Z filtering.
                    if link_level is None and abs(hp.Z - filter_elev) > LEVEL_ELEV_TOL_FT:
                        continue
                    hp = DB.XYZ(hp.X, hp.Y, elev)
                    area = _create_area(doc, target_view, hp)
                    if area is None:
                        failed_areas += 1
                        continue

                    # Copy name/number from linked element (Area or Room).
                    try:
                        src_name = src_area.Name
                    except Exception:
                        src_name = _get_param_as_string(src_area, 'Name')
                    try:
                        src_number = src_area.Number
                    except Exception:
                        src_number = _get_param_as_string(src_area, 'Number')
                    _set_param_string(area, 'Name', src_name)
                    _set_param_string(area, 'Number', src_number)

                    created_areas += 1

                    tag = _create_area_tag(doc, target_view, area, hp, host_tag_type)
                    if tag is None:
                        failed_tags += 1
                    else:
                        created_tags += 1
                except Exception:
                    failed_areas += 1

        # Switch UI active view after transaction commits.
        if target_view is not None:
            try:
                revit.uidoc.ActiveView = target_view
            except Exception:
                pass

        output.print_md('## Link Boundaries → Area Boundaries + Areas + Tags')
        output.print_md('* Link: `{}`'.format(getattr(link_instance, 'Name', 'Revit Link')))
        output.print_md('* Source: `{}`'.format(source_choice))
        if link_level is not None:
            try:
                output.print_md('* Linked level filter: `{}` (host Z ≈ `{:.3f} ft`)'.format(link_level.Name, filter_elev))
            except Exception:
                output.print_md('* Linked level filter applied')
        output.print_md('* Curves read: `{}`'.format(len(raw_curves)))
        output.print_md('* Area boundary curves created (deduped): `{}`'.format(created))
        output.print_md('* Skipped (transform/flatten): `{}`'.format(skipped))
        output.print_md('* Skipped (too short): `{}`'.format(skipped_too_short))
        output.print_md('* Failed boundary curves: `{}`'.format(failed))
        output.print_md('* Areas created: `{}` | failed: `{}`'.format(created_areas, failed_areas))
        output.print_md('* Tags created: `{}` | failed: `{}`'.format(created_tags, failed_tags))
        if link_tag_type is not None and host_tag_type is None:
            output.print_md('> Note: Matching Area Tag type from link was not found in host. Load the same Area Tag family/type, then re-run to match.')
        return

    # SPACE SEPARATOR MODE
    with revit.Transaction('Link boundaries to space separators'):
        creator = doc.Create

        if not hasattr(creator, 'NewSpaceBoundaryLines'):
            forms.alert('This Revit version/API does not expose NewSpaceBoundaryLines.', exitscript=True)

        deleted_existing = _prompt_delete_existing_space_separators(doc, view)
        if deleted_existing is None:
            return

        # Create in batches for speed, then fall back per-curve on failures.
        for batch in _chunks(curves, 200):
            ca = DB.CurveArray()
            for c in batch:
                ca.Append(c)

            try:
                creator.NewSpaceBoundaryLines(sketch_plane, ca, view)
                created += len(batch)
            except Exception as ex:
                logger.debug('Batch creation failed, falling back per-curve: %s', ex)
                for c in batch:
                    ca2 = DB.CurveArray()
                    ca2.Append(c)
                    try:
                        creator.NewSpaceBoundaryLines(sketch_plane, ca2, view)
                        created += 1
                    except Exception:
                        failed += 1

        # Attempt to unhide the category in this view (view template can still override).
        _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceSeparationLines)

    output.print_md('## Link Boundaries → Space Separators')
    output.print_md('* Link: `{}`'.format(getattr(link_instance, 'Name', 'Revit Link')))
    output.print_md('* Source: `{}`'.format(source_choice))
    output.print_md('* Existing space separators deleted: `{}`'.format(deleted_existing))
    output.print_md('* Curves read: `{}`'.format(len(raw_curves)))
    output.print_md('* Curves created (deduped): `{}`'.format(created))
    output.print_md('* Skipped (transform/flatten): `{}`'.format(skipped))
    output.print_md('* Failed to create: `{}`'.format(failed))


if __name__ == '__main__':
    main()
