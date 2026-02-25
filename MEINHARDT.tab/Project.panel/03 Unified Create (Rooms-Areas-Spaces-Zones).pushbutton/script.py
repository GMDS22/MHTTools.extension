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


def _pick_elements_in_view_by_bic(uidoc, prompt, bic):
    """Prompt user to multi-select elements in the active view by category."""
    try:
        from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
    except Exception:
        return []

    class _BicSelectionFilter(ISelectionFilter):
        def __init__(self, bic_val):
            try:
                self._bic_int = int(bic_val)
            except Exception:
                self._bic_int = None

        def AllowElement(self, e):
            try:
                c = getattr(e, 'Category', None)
                if c is None or self._bic_int is None:
                    return False
                return c.Id.IntegerValue == self._bic_int
            except Exception:
                return False

        def AllowReference(self, reference, position):
            return False

    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, _BicSelectionFilter(bic), prompt)
    except Exception:
        return []

    picked = []
    for r in refs or []:
        try:
            e = uidoc.Document.GetElement(r.ElementId)
            if e is not None:
                picked.append(e)
        except Exception:
            continue
    return picked


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


def _get_default_area_tag_type(doc):
    try:
        types = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_AreaTags)
                 .WhereElementIsElementType()
                 .ToElements())
        return types[0] if types else None
    except Exception:
        return None


def _get_default_room_tag_type(doc):
    try:
        types = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_RoomTags)
                 .WhereElementIsElementType()
                 .ToElements())
        return types[0] if types else None
    except Exception:
        return None


def _set_room_name_number(room, name_value, number_value):
    ok_name = _set_param_string(room, 'Name', name_value)
    ok_num = _set_param_string(room, 'Number', number_value)
    return ok_name, ok_num


def _create_room(doc, level, point_xyz):
    uv = DB.UV(point_xyz.X, point_xyz.Y)
    creator = doc.Create

    for args in (
        (level, uv),
        (level.Id, uv),
    ):
        try:
            if hasattr(creator, 'NewRoom'):
                r = creator.NewRoom(*args)
                if r is not None:
                    return r
        except Exception:
            continue
    return None


def _create_room_tag(doc, view, room, point_xyz, tag_type):
    uv = DB.UV(point_xyz.X, point_xyz.Y)
    creator = doc.Create
    tag = None

    # Revit API: NewRoomTag(roomId: LinkElementId, point: UV, viewId: ElementId)
    try:
        rid = DB.LinkElementId(room.Id)
    except Exception:
        rid = None

    for args in (
        (rid, uv, view.Id),
        (rid, uv, view.Id),
    ):
        try:
            if rid is not None and hasattr(creator, 'NewRoomTag'):
                tag = creator.NewRoomTag(*args)
                break
        except Exception:
            continue

    if tag is not None and tag_type is not None:
        try:
            tag.ChangeTypeId(tag_type.Id)
        except Exception:
            try:
                if hasattr(tag, 'RoomTagType'):
                    tag.RoomTagType = tag_type
            except Exception:
                pass

    return tag


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

    # Unified UI: exactly 4 top-level options, then per-item creation methods.
    source_choice_override = None

    target = forms.CommandSwitchWindow.show(
        ['ROOMS', 'AREAS', 'SPACES', 'HVAC ZONES'],
        message='What do you want to create?'
    )
    if not target:
        return

    mode = None

    if target == 'ROOMS':
        action = forms.CommandSwitchWindow.show(
            [
                'From Link Rooms → Room Separation Lines + Rooms + Room Tags',
                'From Current View Area Boundaries → Room Separation Lines + Rooms + Room Tags',
                'From Selected Areas (Current View) → Room Separation Lines + Rooms + Room Tags',
            ],
            message='How do you want to create Rooms?'
        )
        if not action:
            return
        if action.startswith('From Link'):
            mode = 'UNIFIED_ROOMS_FROM_LINK'
        elif action.startswith('From Selected Areas'):
            mode = 'UNIFIED_ROOMS_FROM_SELECTED_AREAS'
        else:
            mode = 'UNIFIED_ROOMS_FROM_VIEW_AREA_BOUNDS'

    elif target == 'AREAS':
        action = forms.CommandSwitchWindow.show(
            [
                'From Link Rooms → Area Boundaries + Areas + Area Tags',
                'From Link Areas → Area Boundaries + Areas + Area Tags',
                'From Current View Spaces → Area Boundaries + Areas + Area Tags',
                'From Selected Rooms (Current View) → Area Boundaries + Areas + Area Tags',
            ],
            message='How do you want to create Areas?'
        )
        if not action:
            return
        if action.startswith('From Link Rooms'):
            mode = 'Create Area Boundaries + Areas + Area Tags (from Linked Rooms)'
        elif action.startswith('From Link Areas'):
            # Use the existing link-based Area mode but force source = Areas.
            mode = 'Create Areas + Area Boundaries + Tags (from Link)'
            source_choice_override = 'Areas'
        elif action.startswith('From Selected Rooms'):
            mode = 'UNIFIED_AREAS_FROM_SELECTED_ROOMS'
        else:
            mode = 'UNIFIED_AREAS_FROM_VIEW_SPACES'

    elif target == 'SPACES':
        action = forms.CommandSwitchWindow.show(
            [
                'From Link Areas → Spaces + Space Tags',
                'From Link Rooms → Space Separation Lines + Spaces + Space Tags',
                'From Current View Area Boundaries → Space Separation Lines + Spaces + Space Tags',
                'From Selected Areas (Current View) → Spaces + Space Tags (no separators)',
                'From Selected Rooms (Current View) → Space Separation Lines + Spaces + Space Tags',
            ],
            message='How do you want to create Spaces?'
        )
        if not action:
            return
        if action.startswith('From Link Areas'):
            mode = 'Create Spaces + Space Tags (from Link Areas)'
        elif action.startswith('From Link Rooms'):
            mode = 'Create Space Boundaries + Spaces + Space Tags (from Linked Rooms)'
        elif action.startswith('From Current View Area Boundaries'):
            mode = 'Create Space Separators (from Area Boundaries in Current View)'
        elif action.startswith('From Selected Rooms'):
            mode = 'UNIFIED_SPACES_FROM_SELECTED_ROOMS'
        else:
            mode = 'UNIFIED_SPACES_FROM_SELECTED_AREAS'

    elif target == 'HVAC ZONES':
        action = forms.CommandSwitchWindow.show(
            [
                'From Selected Areas (Current View) → Zones',
                'From Selected Spaces (Current View) → Zones',
            ],
            message='How do you want to create HVAC Zones?'
        )
        if not action:
            return
        if action.startswith('From Selected Areas'):
            mode = 'Create HVAC Zones (from Selected Areas in Current View)'
        else:
            mode = 'UNIFIED_ZONES_FROM_SELECTED_SPACES'

    if not mode:
        return

    if mode.startswith('Start Over in Current View'):
        with revit.Transaction('Start over in current view'):
            deleted = _start_over_in_current_view(doc, view)
            if deleted is None:
                return
        forms.alert('Deleted {} element(s) in the current view.'.format(deleted))
        return

    # UNIFIED: Zones from selected Spaces (current view)
    if mode == 'UNIFIED_ZONES_FROM_SELECTED_SPACES':
        phase = _get_view_phase(doc, view)
        if phase is None:
            forms.alert('Could not determine a phase for this view/project. Zones are phase-dependent.', exitscript=True)

        picked_space_elems = _pick_elements_in_view_by_bic(
            revit.uidoc,
            'Select Spaces to Create HVAC Zones',
            DB.BuiltInCategory.OST_MEPSpaces
        )
        if not picked_space_elems:
            return

        # Group spaces by intended Zone name so duplicates merge into one Zone.
        groups = {}
        for sp in picked_space_elems:
            nm, num = _get_space_name_number(sp)
            zone_name = _make_zone_name(nm, num) or 'Zone'
            g = groups.setdefault(zone_name, {'spaces': [], 'space_ids': set()})
            try:
                sid = sp.Id.IntegerValue
            except Exception:
                sid = None
            if sid is None or sid not in g['space_ids']:
                g['spaces'].append(sp)
                if sid is not None:
                    g['space_ids'].add(sid)

        existing_zones_by_name = _collect_zones_by_name(doc, phase=phase)

        name_conflicts = set([zn for zn in groups.keys() if zn in existing_zones_by_name])
        conflict_policy = 'Keep existing (skip)'
        if name_conflicts:
            conflict_policy = forms.CommandSwitchWindow.show(
                ['Keep existing (skip)', 'Delete existing and recreate', 'Cancel'],
                message='Found existing Zone(s) with the same name as {} selection(s).\n\nHow should duplicates be handled?'.format(len(name_conflicts))
            )
            if not conflict_policy or conflict_policy == 'Cancel':
                return

        created = 0
        skipped_existing_name = 0
        failed = 0

        with revit.Transaction('Create HVAC Zones from selected Spaces (current view)'):
            creator = doc.Create
            if not hasattr(creator, 'NewZone'):
                forms.alert('This Revit version/API does not expose NewZone.', exitscript=True)

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
                        pass
                existing_zones_by_name = _collect_zones_by_name(doc, phase=phase)

            for zn, g in groups.items():
                if zn in existing_zones_by_name and conflict_policy == 'Keep existing (skip)':
                    skipped_existing_name += 1
                    continue

                try:
                    zone = creator.NewZone(level, phase)
                    try:
                        zone.Name = zn
                    except Exception:
                        pass

                    ss = DB.Mechanical.SpaceSet()
                    for sp in g.get('spaces', []):
                        try:
                            ss.Insert(sp)
                        except Exception:
                            pass
                    ok = False
                    try:
                        ok = zone.AddSpaces(ss)
                    except Exception:
                        ok = False
                    if ok:
                        created += 1
                    else:
                        failed += 1
                        try:
                            doc.Delete(zone.Id)
                        except Exception:
                            pass
                except Exception:
                    failed += 1

        forms.alert(
            'Zones from selected Spaces complete:\n\n'
            'Selected Spaces: {}\n'
            'Zones created: {}\n'
            'Skipped (existing zone name): {}\n'
            'Failed: {}'.format(len(picked_space_elems), created, skipped_existing_name, failed)
        )
        return

    # UNIFIED: Spaces + tags from selected Areas (no separators)
    if mode == 'UNIFIED_SPACES_FROM_SELECTED_AREAS':
        phase = _get_view_phase(doc, view)
        if phase is None:
            forms.alert('Could not determine a Phase for Space creation.', exitscript=True)

        areas = _collect_areas_in_current_view(doc, view, level=level)
        if not areas:
            forms.alert('No Areas found in the current view.', exitscript=True)

        items = [_AreaPickItem(a) for a in areas]
        picked = forms.SelectFromList.show(
            items,
            name_attr='display',
            title='Select Areas to Create Spaces',
            button_name='Create Spaces',
            multiselect=True
        )
        if not picked:
            return

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

        with revit.Transaction('Create Spaces + Tags from selected Areas (current view)'):
            for it in picked:
                a = it.area
                try:
                    lp = getattr(a, 'Location', None)
                    pt = getattr(lp, 'Point', None) if lp else None
                    if pt is None:
                        continue

                    hp = DB.XYZ(pt.X, pt.Y, level.Elevation)
                    k = _point_key_xy(hp)
                    if k is not None and k in existing_keys:
                        skipped_existing += 1
                        continue

                    sp = _create_space(doc, level, phase, hp)
                    if sp is None:
                        failed_spaces += 1
                        continue

                    area_name, area_number = _get_area_name_number(a)
                    _set_space_name_number(sp, area_name, area_number)
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

            _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceTags)

        forms.alert(
            'Spaces from selected Areas complete:\n\n'
            'Selected Areas: {}\n'
            'Spaces created: {} | failed: {} | skipped existing: {}\n'
            'Tags created: {} | failed: {}'.format(len(picked), created_spaces, failed_spaces, skipped_existing, created_tags, failed_tags)
        )
        return

    # UNIFIED: Spaces from selected Rooms (current view)
    if mode == 'UNIFIED_SPACES_FROM_SELECTED_ROOMS':
        phase = _get_view_phase(doc, view)
        if phase is None:
            forms.alert('Could not determine a phase for this view/project. Spaces are phase-dependent.', exitscript=True)

        try:
            rooms = (DB.FilteredElementCollector(doc, view.Id)
                     .OfCategory(DB.BuiltInCategory.OST_Rooms)
                     .WhereElementIsNotElementType()
                     .ToElements())
        except Exception:
            rooms = []

        if not rooms:
            forms.alert('No Rooms found in the current view.', exitscript=True)

        class _RoomPick(object):
            def __init__(self, rm):
                self.rm = rm
                try:
                    self.name = rm.Name
                except Exception:
                    self.name = _get_param_as_string(rm, 'Name')
                try:
                    self.number = rm.Number
                except Exception:
                    self.number = _get_param_as_string(rm, 'Number')
                parts = []
                if self.number:
                    parts.append(str(self.number))
                if self.name:
                    parts.append(str(self.name))
                self.display = ' - '.join(parts) if parts else 'Room'

        rm_items = [_RoomPick(r) for r in rooms]
        picked_rooms = forms.SelectFromList.show(
            rm_items,
            name_attr='display',
            title='Select Rooms to Create Spaces',
            button_name='Create Spaces',
            multiselect=True
        )
        if not picked_rooms:
            return

        opts = DB.SpatialElementBoundaryOptions()
        if hasattr(DB, 'SpatialElementBoundaryLocation'):
            try:
                opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
            except Exception:
                pass

        elev = level.Elevation
        raw_curves = []
        room_points = []
        for it in picked_rooms:
            rm = it.rm
            try:
                lp = getattr(rm, 'Location', None)
                pt = getattr(lp, 'Point', None) if lp else None
                if pt is not None:
                    room_points.append((pt, rm))

                seglists = rm.GetBoundarySegments(opts)
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
            except Exception:
                continue

        if not raw_curves:
            forms.alert('No boundary curves found on the selected Rooms.', exitscript=True)

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
            forms.alert('No usable boundary curves found on the selected Rooms.', exitscript=True)

        # Ensure sketch plane
        sketch_plane = None
        try:
            sketch_plane = view.SketchPlane
        except Exception:
            sketch_plane = None
        if sketch_plane is None:
            plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
            sketch_plane = DB.SketchPlane.Create(doc, plane)

        # Avoid duplicate Spaces by XY
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
                pt = getattr(loc, 'Point', None) if loc else None
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

        with revit.Transaction('Selected Rooms (current view) to Space separators + Spaces + Tags'):
            creator = doc.Create
            if not hasattr(creator, 'NewSpaceBoundaryLines'):
                forms.alert('This Revit version/API does not expose NewSpaceBoundaryLines.', exitscript=True)

            for batch in _chunks(curves, 200):
                ca = DB.CurveArray()
                for cc in batch:
                    ca.Append(cc)
                try:
                    creator.NewSpaceBoundaryLines(sketch_plane, ca, view)
                    created_bounds += len(batch)
                except Exception as ex:
                    logger.debug('Batch space boundary creation failed, per-curve: %s', ex)
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

            for pt, rm in room_points:
                try:
                    hp = DB.XYZ(pt.X, pt.Y, elev)
                    k = _point_key_xy(hp)
                    if k is not None and k in existing_keys:
                        skipped_existing += 1
                        continue

                    sp = _create_space(doc, level, phase, hp)
                    if sp is None:
                        failed_spaces += 1
                        continue

                    try:
                        nm = rm.Name
                    except Exception:
                        nm = _get_param_as_string(rm, 'Name')
                    try:
                        num = rm.Number
                    except Exception:
                        num = _get_param_as_string(rm, 'Number')
                    _set_space_name_number(sp, nm, num)
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

            _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceTags)

        forms.alert(
            'Spaces from selected Rooms complete:\n\n'
            'Selected Rooms: {}\n'
            'Space separator curves created: {} | failed: {}\n'
            'Spaces created: {} | failed: {} | skipped existing: {}\n'
            'Tags created: {} | failed: {}'.format(
                len(picked_rooms), created_bounds, failed_bounds, created_spaces, failed_spaces, skipped_existing, created_tags, failed_tags
            )
        )
        return

    # UNIFIED: Areas from selected Rooms (current view)
    if mode == 'UNIFIED_AREAS_FROM_SELECTED_ROOMS':
        elev = level.Elevation

        try:
            rooms = (DB.FilteredElementCollector(doc, view.Id)
                     .OfCategory(DB.BuiltInCategory.OST_Rooms)
                     .WhereElementIsNotElementType()
                     .ToElements())
        except Exception:
            rooms = []

        if not rooms:
            forms.alert('No Rooms found in the current view.', exitscript=True)

        class _RoomPick2(object):
            def __init__(self, rm):
                self.rm = rm
                try:
                    self.name = rm.Name
                except Exception:
                    self.name = _get_param_as_string(rm, 'Name')
                try:
                    self.number = rm.Number
                except Exception:
                    self.number = _get_param_as_string(rm, 'Number')
                parts = []
                if self.number:
                    parts.append(str(self.number))
                if self.name:
                    parts.append(str(self.name))
                self.display = ' - '.join(parts) if parts else 'Room'

        rm_items = [_RoomPick2(r) for r in rooms]
        picked_rooms = forms.SelectFromList.show(
            rm_items,
            name_attr='display',
            title='Select Rooms to Create Areas',
            button_name='Create Areas',
            multiselect=True
        )
        if not picked_rooms:
            return

        # Select/create an Area Plan view for this level
        target_view = view
        created_view = None
        scheme = None

        if view.ViewType != DB.ViewType.AreaPlan:
            picked_view = _pick_area_plan_view(doc, level)
            if picked_view:
                target_view = picked_view
            else:
                scheme = _pick_area_scheme(doc)
                if not scheme:
                    return

        tag_type = _get_default_area_tag_type(doc)
        opts = DB.SpatialElementBoundaryOptions()
        if hasattr(DB, 'SpatialElementBoundaryLocation'):
            try:
                opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
            except Exception:
                pass

        raw_curves = []
        room_points = []
        for it in picked_rooms:
            rm = it.rm
            try:
                lp = getattr(rm, 'Location', None)
                pt = getattr(lp, 'Point', None) if lp else None
                if pt is not None:
                    room_points.append((pt, rm))

                seglists = rm.GetBoundarySegments(opts)
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
            except Exception:
                continue

        if not raw_curves:
            forms.alert('No Room boundary curves found on the selected Rooms.', exitscript=True)

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
            forms.alert('No usable Room boundary curves found on the selected Rooms.', exitscript=True)

        created_bounds = 0
        failed_bounds = 0
        created_areas = 0
        failed_areas = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Selected Rooms (current view) to Areas + Tags'):
            if target_view.ViewType != DB.ViewType.AreaPlan:
                try:
                    created_view = DB.ViewPlan.CreateAreaPlan(doc, scheme.Id, level.Id)
                    target_view = created_view
                except Exception as ex:
                    forms.alert('Failed to create Area Plan view: {}'.format(ex), exitscript=True)

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

            for pt, rm in room_points:
                try:
                    hp = DB.XYZ(pt.X, pt.Y, elev)
                    area = _create_area(doc, target_view, hp)
                    if area is None:
                        failed_areas += 1
                        continue

                    try:
                        nm = rm.Name
                    except Exception:
                        nm = _get_param_as_string(rm, 'Name')
                    try:
                        num = rm.Number
                    except Exception:
                        num = _get_param_as_string(rm, 'Number')
                    _set_param_string(area, 'Name', nm)
                    _set_param_string(area, 'Number', num)
                    created_areas += 1

                    tag = _create_area_tag(doc, target_view, area, hp, tag_type)
                    if tag is None:
                        failed_tags += 1
                    else:
                        created_tags += 1
                except Exception:
                    failed_areas += 1

        try:
            revit.uidoc.ActiveView = target_view
        except Exception:
            pass

        forms.alert(
            'Areas from selected Rooms complete:\n\n'
            'Selected Rooms: {}\n'
            'Area boundary curves created: {} | failed: {}\n'
            'Areas created: {} | failed: {}\n'
            'Tags created: {} | failed: {}'.format(
                len(picked_rooms), created_bounds, failed_bounds, created_areas, failed_areas, created_tags, failed_tags
            )
        )
        return

    # UNIFIED: Rooms from selected Areas (current view)
    if mode == 'UNIFIED_ROOMS_FROM_SELECTED_AREAS':
        elev = level.Elevation

        areas = _collect_areas_in_current_view(doc, view, level=level)
        if not areas:
            forms.alert('No Areas found in the current view.', exitscript=True)

        items = [_AreaPickItem(a) for a in areas]
        picked = forms.SelectFromList.show(
            items,
            name_attr='display',
            title='Select Areas to Create Rooms',
            button_name='Create Rooms',
            multiselect=True
        )
        if not picked:
            return

        opts = DB.SpatialElementBoundaryOptions()
        if hasattr(DB, 'SpatialElementBoundaryLocation'):
            try:
                opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
            except Exception:
                pass

        raw_curves = []
        area_points = []
        for it in picked:
            a = it.area
            try:
                lp = getattr(a, 'Location', None)
                pt = getattr(lp, 'Point', None) if lp else None
                if pt is not None:
                    area_points.append((pt, a))

                seglists = a.GetBoundarySegments(opts)
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
            except Exception:
                continue

        if not raw_curves:
            forms.alert('No boundary curves found on the selected Areas.', exitscript=True)

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
            forms.alert('No usable boundary curves found on the selected Areas.', exitscript=True)

        # Ensure sketch plane
        sketch_plane = None
        try:
            sketch_plane = view.SketchPlane
        except Exception:
            sketch_plane = None
        if sketch_plane is None:
            plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
            sketch_plane = DB.SketchPlane.Create(doc, plane)

        # Avoid duplicate rooms by XY
        existing_keys = set()
        try:
            existing_rooms = (DB.FilteredElementCollector(doc)
                              .OfCategory(DB.BuiltInCategory.OST_Rooms)
                              .WhereElementIsNotElementType()
                              .ToElements())
        except Exception:
            existing_rooms = []
        for r in existing_rooms:
            try:
                loc = getattr(r, 'Location', None)
                pt = getattr(loc, 'Point', None) if loc else None
                if pt is None:
                    continue
                k = _point_key_xy(pt)
                if k is not None:
                    existing_keys.add(k)
            except Exception:
                continue

        tag_type = _get_default_room_tag_type(doc)
        created_bounds = 0
        failed_bounds = 0
        created_rooms = 0
        failed_rooms = 0
        skipped_existing = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Selected Areas (current view) to room separation lines + rooms + tags'):
            creator = doc.Create
            if not hasattr(creator, 'NewRoomBoundaryLines'):
                forms.alert('This Revit version/API does not expose NewRoomBoundaryLines.', exitscript=True)

            for batch in _chunks(curves, 200):
                ca = DB.CurveArray()
                for cc in batch:
                    ca.Append(cc)
                try:
                    creator.NewRoomBoundaryLines(sketch_plane, ca, view)
                    created_bounds += len(batch)
                except Exception as ex:
                    logger.debug('Batch room boundary creation failed, per-curve: %s', ex)
                    for cc in batch:
                        ca2 = DB.CurveArray()
                        ca2.Append(cc)
                        try:
                            creator.NewRoomBoundaryLines(sketch_plane, ca2, view)
                            created_bounds += 1
                        except Exception:
                            failed_bounds += 1

            try:
                doc.Regenerate()
            except Exception:
                pass

            for pt, a in area_points:
                try:
                    hp = DB.XYZ(pt.X, pt.Y, elev)
                    k = _point_key_xy(hp)
                    if k is not None and k in existing_keys:
                        skipped_existing += 1
                        continue

                    rm = _create_room(doc, level, hp)
                    if rm is None:
                        failed_rooms += 1
                        continue

                    an, anum = _get_area_name_number(a)
                    _set_room_name_number(rm, an, anum)
                    created_rooms += 1
                    if k is not None:
                        existing_keys.add(k)

                    tag = _create_room_tag(doc, view, rm, hp, tag_type)
                    if tag is None:
                        failed_tags += 1
                    else:
                        created_tags += 1
                except Exception:
                    failed_rooms += 1

            _try_unhide_category(view, DB.BuiltInCategory.OST_RoomTags)

        forms.alert(
            'Rooms from selected Areas complete:\n\n'
            'Selected Areas: {}\n'
            'Room separation curves created: {} | failed: {}\n'
            'Rooms created: {} | failed: {} | skipped existing: {}\n'
            'Tags created: {} | failed: {}'.format(
                len(picked), created_bounds, failed_bounds, created_rooms, failed_rooms, skipped_existing, created_tags, failed_tags
            )
        )
        return

    # UNIFIED: Areas from Spaces in current view
    if mode == 'UNIFIED_AREAS_FROM_VIEW_SPACES':
        elev = level.Elevation

        try:
            spaces = (DB.FilteredElementCollector(doc, view.Id)
                      .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
                      .WhereElementIsNotElementType()
                      .ToElements())
        except Exception:
            spaces = []

        if not spaces:
            forms.alert('No MEP Spaces found in the current view.', exitscript=True)

        # Select/create an Area Plan view for this level
        target_view = view
        created_view = None
        scheme = None

        if view.ViewType != DB.ViewType.AreaPlan:
            picked_view = _pick_area_plan_view(doc, level)
            if picked_view:
                target_view = picked_view
            else:
                scheme = _pick_area_scheme(doc)
                if not scheme:
                    return

        tag_type = _get_default_area_tag_type(doc)
        opts = DB.SpatialElementBoundaryOptions()
        if hasattr(DB, 'SpatialElementBoundaryLocation'):
            try:
                opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
            except Exception:
                pass

        raw_curves = []
        space_points = []

        for sp in spaces:
            try:
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
            except Exception:
                continue

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

        created_bounds = 0
        failed_bounds = 0
        created_areas = 0
        failed_areas = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Spaces (current view) to Areas + Tags'):
            if target_view.ViewType != DB.ViewType.AreaPlan:
                try:
                    created_view = DB.ViewPlan.CreateAreaPlan(doc, scheme.Id, level.Id)
                    target_view = created_view
                except Exception as ex:
                    forms.alert('Failed to create Area Plan view: {}'.format(ex), exitscript=True)

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

        try:
            revit.uidoc.ActiveView = target_view
        except Exception:
            pass

        forms.alert(
            'Areas from Spaces (current view) complete:\n\n'
            'Spaces in view: {}\n'
            'Area boundary curves created: {} | failed: {}\n'
            'Areas created: {} | failed: {}\n'
            'Tags created: {} | failed: {}'.format(
                len(spaces), created_bounds, failed_bounds, created_areas, failed_areas, created_tags, failed_tags
            )
        )
        return

    if mode == 'Create HVAC Zones (from Selected Areas in Current View)':
        phase = _get_view_phase(doc, view)
        if phase is None:
            forms.alert('Could not determine a phase for this view/project. Zones are phase-dependent.', exitscript=True)

        picked_areas = _pick_elements_in_view_by_bic(
            revit.uidoc,
            'Select Areas to Create HVAC Zones',
            DB.BuiltInCategory.OST_Areas
        )
        if not picked_areas:
            return

        by_num_name, by_num, by_name = _build_space_lookup_for_level(doc, level=level)
        existing_zones_by_name = _collect_zones_by_name(doc, phase=phase)

        # Precompute matches and intended names so we can prompt once.
        # IMPORTANT: If multiple Areas have the same intended Zone name, they merge into ONE Zone.
        groups = {}
        missing_space_areas = 0
        already_zoned_spaces = 0
        name_conflicts = set()

        for area in picked_areas:
            area_name, area_number = _get_area_name_number(area)
            zone_name = _make_zone_name(area_name, area_number) or 'Zone'

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

            # Capture placement point so we can create a Space if needed.
            area_pt = None
            try:
                lp = getattr(area, 'Location', None)
                area_pt = getattr(lp, 'Point', None) if lp else None
            except Exception:
                area_pt = None

            if space is None:
                missing_space_areas += 1
            else:
                try:
                    z = getattr(space, 'Zone', None)
                    if z is not None:
                        try:
                            if hasattr(z, 'IsDefaultZone') and z.IsDefaultZone:
                                pass
                            else:
                                already_zoned_spaces += 1
                        except Exception:
                            already_zoned_spaces += 1
                except Exception:
                    pass

            g = groups.setdefault(zone_name, {
                'areas': [],
                'spaces': [],
                'space_ids': set(),
                'missing_spaces': 0,
                'missing_area_infos': [],  # list of {pt, name, number}
            })
            g['areas'].append(area)
            if space is None:
                g['missing_spaces'] += 1
                if area_pt is not None:
                    g['missing_area_infos'].append({'pt': area_pt, 'name': area_name, 'number': area_number})
            else:
                try:
                    sid = space.Id.IntegerValue
                except Exception:
                    sid = None
                if sid is None or sid not in g['space_ids']:
                    g['spaces'].append(space)
                    if sid is not None:
                        g['space_ids'].add(sid)

        # If zones with same name already exist, ask what to do.
        conflict_policy = 'Keep existing (add spaces)'
        if name_conflicts:
            conflict_policy = forms.CommandSwitchWindow.show(
                ['Keep existing (add spaces)', 'Delete existing and recreate', 'Cancel'],
                message=(
                    'Found existing Zone(s) with the same name as {} selected Area(s).\n\n'
                    'How should duplicates be handled?\n\n'
                    'Note: Default Zone(s) will never be deleted.'
                ).format(len(name_conflicts))
            )
            if not conflict_policy or conflict_policy == 'Cancel':
                return

        zoned_space_policy = 'Skip'
        if already_zoned_spaces:
            zoned_space_policy = forms.CommandSwitchWindow.show(
                ['Skip', 'Move Space to new Zone', 'Cancel'],
                message=(
                    '{} matching Space(s) are already assigned to a Zone.\n\n'
                    'To put them in the new Zone, they must be removed from their current Zone first.'
                ).format(already_zoned_spaces)
            )
            if not zoned_space_policy or zoned_space_policy == 'Cancel':
                return

        created = 0
        created_empty = 0
        updated_existing = 0
        spaces_created = 0
        spaces_create_failed = 0
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

                # Refresh space lookup too (optional, but keeps matching consistent if Spaces existed before).
                by_num_name, by_num, by_name = _build_space_lookup_for_level(doc, level=level)

            for zone_name, g in groups.items():
                try:
                    # Determine target Zone: reuse existing (add spaces) or create new
                    zone = None
                    if zone_name in existing_zones_by_name and conflict_policy == 'Keep existing (add spaces)':
                        # Prefer a non-default zone if present
                        for zc in existing_zones_by_name.get(zone_name, []):
                            try:
                                if hasattr(zc, 'IsDefaultZone') and zc.IsDefaultZone:
                                    continue
                            except Exception:
                                pass
                            zone = zc
                            break
                        if zone is None:
                            try:
                                zone = existing_zones_by_name.get(zone_name, [None])[0]
                            except Exception:
                                zone = None
                    if zone is None:
                        zone = creator.NewZone(level, phase)
                        try:
                            zone.Name = zone_name
                        except Exception:
                            pass

                    # Ensure we have Spaces to add. If missing, try to create Spaces at each Area point.
                    spaces_to_add = list(g.get('spaces', []) or [])
                    seen_ids = set(g.get('space_ids', set()) or set())
                    for mi in g.get('missing_area_infos', []) or []:
                        try:
                            pt = mi.get('pt')
                            if pt is None:
                                continue
                            hp = DB.XYZ(pt.X, pt.Y, level.Elevation)
                            sp = _create_space(doc, level, phase, hp)
                            if sp is None:
                                spaces_create_failed += 1
                                continue
                            spaces_created += 1
                            _set_space_name_number(sp, mi.get('name'), mi.get('number'))
                            try:
                                sid = sp.Id.IntegerValue
                            except Exception:
                                sid = None
                            if sid is None or sid not in seen_ids:
                                spaces_to_add.append(sp)
                                if sid is not None:
                                    seen_ids.add(sid)
                        except Exception:
                            spaces_create_failed += 1

                    ss = DB.Mechanical.SpaceSet()
                    inserted_any = False

                    for space in spaces_to_add:
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

                        # If it's already in the target zone, skip it
                        if old_zone is not None:
                            try:
                                if hasattr(old_zone, 'Id') and hasattr(zone, 'Id') and old_zone.Id == zone.Id:
                                    continue
                            except Exception:
                                pass

                        if old_zone is not None:
                            if zoned_space_policy == 'Skip':
                                skipped_space_already_zoned += 1
                                continue
                            else:
                                try:
                                    ss_old = DB.Mechanical.SpaceSet()
                                    ss_old.Insert(space)
                                    old_zone.RemoveSpaces(ss_old)
                                    moved_spaces += 1
                                except Exception:
                                    # can't safely move this one
                                    failed += 1
                                    continue

                        try:
                            ss.Insert(space)
                            inserted_any = True
                        except Exception:
                            pass

                    if not inserted_any:
                        # No spaces could be added (all missing or already-zoned). Keep the Zone as empty.
                        created += 1
                        created_empty += 1
                        if zone_name in existing_zones_by_name and conflict_policy == 'Keep existing (add spaces)':
                            updated_existing += 1
                        continue

                    ok = False
                    try:
                        ok = zone.AddSpaces(ss)
                    except Exception:
                        ok = False

                    if ok:
                        if zone_name in existing_zones_by_name and conflict_policy == 'Keep existing (add spaces)':
                            updated_existing += 1
                        else:
                            created += 1
                    else:
                        failed += 1
                        # If we created a brand new zone and failed to add, delete it; if existing, keep.
                        try:
                            if zone_name not in existing_zones_by_name:
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
            'Existing Zones updated: {}\n'
            'Spaces auto-created: {} (failed: {})\n'
            'Skipped (Space already zoned): {}\n'
            'Spaces moved from old zone: {}\n'
            'Failed: {}'.format(
                len(picked_areas),
                created,
                created_empty,
                updated_existing,
                spaces_created,
                spaces_create_failed,
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

        # Collect Areas in this view for Space placement/name/number
        areas_in_view = _collect_areas_in_current_view(doc, view, level=level)

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

        phase = _get_view_phase(doc, view)
        if phase is None:
            forms.alert('Could not determine a Phase for Space creation.', exitscript=True)

        tag_type = _get_default_space_tag_type(doc)

        created_bounds = 0
        failed_bounds = 0
        created_spaces = 0
        failed_spaces = 0
        skipped_existing_spaces = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Area boundaries (current view) to space separators + spaces + tags'):
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
                    created_bounds += len(batch)
                except Exception as ex:
                    logger.debug('Batch creation failed, falling back per-curve: %s', ex)
                    for cc in batch:
                        ca2 = DB.CurveArray()
                        ca2.Append(cc)
                        try:
                            creator.NewSpaceBoundaryLines(sketch_plane, ca2, view)
                            created_bounds += 1
                        except Exception:
                            failed_bounds += 1

            _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceSeparationLines)

            # Regen so boundaries are recognized before creating Spaces.
            try:
                doc.Regenerate()
            except Exception:
                pass

            # Create Spaces + Tags from Areas in current view
            for a in areas_in_view:
                try:
                    lp = getattr(a, 'Location', None)
                    pt = getattr(lp, 'Point', None) if lp else None
                    if pt is None:
                        continue

                    hp = DB.XYZ(pt.X, pt.Y, elev)

                    k = _point_key_xy(hp)
                    if k is not None and k in existing_keys:
                        skipped_existing_spaces += 1
                        continue

                    sp = _create_space(doc, level, phase, hp)
                    if sp is None:
                        failed_spaces += 1
                        continue

                    # Copy name/number from host Area
                    area_name, area_number = _get_area_name_number(a)
                    _set_space_name_number(sp, area_name, area_number)

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

            _try_unhide_category(view, DB.BuiltInCategory.OST_MEPSpaceTags)

        output.print_md('## Current View Area Boundaries → Space Separators + Spaces + Space Tags')
        output.print_md('* View: `{}`'.format(getattr(view, 'Name', 'Active View')))
        output.print_md('* Existing space separators deleted: `{}`'.format(deleted_existing))
        output.print_md('* Area boundary curves read: `{}`'.format(len(raw_curves)))
        output.print_md('* Space separator curves created (deduped): `{}` | failed: `{}`'.format(created_bounds, failed_bounds))
        output.print_md('* Skipped (flatten): `{}`'.format(skipped))
        output.print_md('* Skipped (too short): `{}`'.format(skipped_too_short))
        output.print_md('* Areas found in view: `{}`'.format(len(areas_in_view)))
        output.print_md('* Spaces created: `{}` | failed: `{}` | skipped existing: `{}`'.format(created_spaces, failed_spaces, skipped_existing_spaces))
        output.print_md('* Tags created: `{}` | failed: `{}`'.format(created_tags, failed_tags))
        if tag_type is None:
            output.print_md('> Note: No Space Tag type found in host. Load a Space Tag family/type, then re-run to tag.')
        return

    # UNIFIED: Rooms from current view Area Boundaries (uses host Areas for placement)
    if mode == 'UNIFIED_ROOMS_FROM_VIEW_AREA_BOUNDS':
        elev = level.Elevation

        raw_curves = _collect_area_boundary_curves_from_view(doc, view)
        if not raw_curves:
            forms.alert('No Area Boundary lines found in the current view.', exitscript=True)

        areas = _collect_areas_in_current_view(doc, view, level=level)
        if not areas:
            forms.alert('No Areas found in the current view. Place Areas first to use them as Room placement points.', exitscript=True)

        items = [_AreaPickItem(a) for a in areas]
        picked = forms.SelectFromList.show(
            items,
            name_attr='display',
            title='Select Areas to Create Rooms',
            button_name='Create Rooms',
            multiselect=True
        )
        if not picked:
            return

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
            forms.alert('No usable boundary curves found in the current view.', exitscript=True)

        # Ensure sketch plane
        sketch_plane = None
        try:
            sketch_plane = view.SketchPlane
        except Exception:
            sketch_plane = None
        if sketch_plane is None:
            plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
            sketch_plane = DB.SketchPlane.Create(doc, plane)

        # Avoid duplicate rooms by XY
        existing_keys = set()
        try:
            existing_rooms = (DB.FilteredElementCollector(doc)
                              .OfCategory(DB.BuiltInCategory.OST_Rooms)
                              .WhereElementIsNotElementType()
                              .ToElements())
        except Exception:
            existing_rooms = []

        for r in existing_rooms:
            try:
                loc = getattr(r, 'Location', None)
                pt = getattr(loc, 'Point', None) if loc else None
                if pt is None:
                    continue
                k = _point_key_xy(pt)
                if k is not None:
                    existing_keys.add(k)
            except Exception:
                continue

        tag_type = _get_default_room_tag_type(doc)
        created_bounds = 0
        failed_bounds = 0
        created_rooms = 0
        failed_rooms = 0
        skipped_existing = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Area boundaries (current view) to room separation lines + rooms + tags'):
            creator = doc.Create
            if not hasattr(creator, 'NewRoomBoundaryLines'):
                forms.alert('This Revit version/API does not expose NewRoomBoundaryLines.', exitscript=True)

            for batch in _chunks(curves, 200):
                ca = DB.CurveArray()
                for cc in batch:
                    ca.Append(cc)
                try:
                    creator.NewRoomBoundaryLines(sketch_plane, ca, view)
                    created_bounds += len(batch)
                except Exception as ex:
                    logger.debug('Batch room boundary creation failed, per-curve: %s', ex)
                    for cc in batch:
                        ca2 = DB.CurveArray()
                        ca2.Append(cc)
                        try:
                            creator.NewRoomBoundaryLines(sketch_plane, ca2, view)
                            created_bounds += 1
                        except Exception:
                            failed_bounds += 1

            try:
                doc.Regenerate()
            except Exception:
                pass

            for it in picked:
                a = it.area
                try:
                    lp = getattr(a, 'Location', None)
                    pt = getattr(lp, 'Point', None) if lp else None
                    if pt is None:
                        continue
                    hp = DB.XYZ(pt.X, pt.Y, elev)
                    k = _point_key_xy(hp)
                    if k is not None and k in existing_keys:
                        skipped_existing += 1
                        continue

                    rm = _create_room(doc, level, hp)
                    if rm is None:
                        failed_rooms += 1
                        continue

                    an, anum = _get_area_name_number(a)
                    _set_room_name_number(rm, an, anum)
                    created_rooms += 1
                    if k is not None:
                        existing_keys.add(k)

                    tag = _create_room_tag(doc, view, rm, hp, tag_type)
                    if tag is None:
                        failed_tags += 1
                    else:
                        created_tags += 1
                except Exception:
                    failed_rooms += 1

        forms.alert(
            'Rooms from current view Area Boundaries complete:\n\n'
            'Room separation curves created: {} | failed: {}\n'
            'Rooms created: {} | failed: {} | skipped existing: {}\n'
            'Tags created: {} | failed: {}'.format(
                created_bounds, failed_bounds, created_rooms, failed_rooms, skipped_existing, created_tags, failed_tags
            )
        )
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

    # UNIFIED: Rooms from linked Rooms
    if mode == 'UNIFIED_ROOMS_FROM_LINK':
        link_t = _get_link_transform(link_instance)
        elev = level.Elevation

        link_level, link_level_host_elev = _pick_link_level_for_host_elevation(link_doc, link_t, elev)
        filter_elev = link_level_host_elev if link_level_host_elev is not None else elev

        # Placement points from linked rooms
        room_points = []
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
                pt = getattr(lp, 'Point', None) if lp else None
                if pt is None:
                    continue
                room_points.append((pt, r))
            except Exception:
                continue

        if not room_points:
            forms.alert('No placed Rooms found in the selected link (after level filtering).', exitscript=True)

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
            forms.alert('No usable Room boundary curves found (after transform/flatten).', exitscript=True)

        # Ensure sketch plane
        sketch_plane = None
        try:
            sketch_plane = view.SketchPlane
        except Exception:
            sketch_plane = None
        if sketch_plane is None:
            plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, elev))
            sketch_plane = DB.SketchPlane.Create(doc, plane)

        # Avoid duplicate rooms by XY
        existing_keys = set()
        try:
            existing_rooms = (DB.FilteredElementCollector(doc)
                              .OfCategory(DB.BuiltInCategory.OST_Rooms)
                              .WhereElementIsNotElementType()
                              .ToElements())
        except Exception:
            existing_rooms = []

        for rm in existing_rooms:
            try:
                loc = getattr(rm, 'Location', None)
                pt = getattr(loc, 'Point', None) if loc else None
                if pt is None:
                    continue
                k = _point_key_xy(pt)
                if k is not None:
                    existing_keys.add(k)
            except Exception:
                continue

        tag_type = _get_default_room_tag_type(doc)
        created_bounds = 0
        failed_bounds = 0
        created_rooms = 0
        failed_rooms = 0
        skipped_existing = 0
        created_tags = 0
        failed_tags = 0

        with revit.Transaction('Link Rooms to room separation lines + rooms + tags'):
            creator = doc.Create
            if not hasattr(creator, 'NewRoomBoundaryLines'):
                forms.alert('This Revit version/API does not expose NewRoomBoundaryLines.', exitscript=True)

            for batch in _chunks(curves, 200):
                ca = DB.CurveArray()
                for cc in batch:
                    ca.Append(cc)
                try:
                    creator.NewRoomBoundaryLines(sketch_plane, ca, view)
                    created_bounds += len(batch)
                except Exception as ex:
                    logger.debug('Batch room boundary creation failed, per-curve: %s', ex)
                    for cc in batch:
                        ca2 = DB.CurveArray()
                        ca2.Append(cc)
                        try:
                            creator.NewRoomBoundaryLines(sketch_plane, ca2, view)
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

                    rm = _create_room(doc, level, hp)
                    if rm is None:
                        failed_rooms += 1
                        continue

                    try:
                        src_name = src_room.Name
                    except Exception:
                        src_name = _get_param_as_string(src_room, 'Name')
                    try:
                        src_number = src_room.Number
                    except Exception:
                        src_number = _get_param_as_string(src_room, 'Number')
                    _set_room_name_number(rm, src_name, src_number)
                    created_rooms += 1
                    if k is not None:
                        existing_keys.add(k)

                    tag = _create_room_tag(doc, view, rm, hp, tag_type)
                    if tag is None:
                        failed_tags += 1
                    else:
                        created_tags += 1
                except Exception:
                    failed_rooms += 1

        forms.alert(
            'Rooms from linked Rooms complete:\n\n'
            'Room separation curves created: {} | failed: {}\n'
            'Rooms created: {} | failed: {} | skipped existing: {}\n'
            'Tags created: {} | failed: {}'.format(
                created_bounds, failed_bounds, created_rooms, failed_rooms, skipped_existing, created_tags, failed_tags
            )
        )
        return

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
    elif source_choice_override is not None:
        source_choice = source_choice_override
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
