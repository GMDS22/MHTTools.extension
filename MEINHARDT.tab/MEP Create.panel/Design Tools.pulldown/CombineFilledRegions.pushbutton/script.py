# coding: utf8
from __future__ import print_function

import clr
import math

clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationCore")

from System.Collections.Generic import List
from System.Drawing import Color as DrawingColor
from System.Windows.Forms import ColorDialog, DialogResult
from System.Windows.Media import BrushConverter

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Color,
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
    SubTransaction,
    Transaction,
    XYZ,
)

from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
_BRUSH_CONVERTER = BrushConverter()
MIN_CURVE_LEN_FT = 0.005


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


def _serialize_element_id(element_id):
    if element_id is None:
        return None
    try:
        return element_id.IntegerValue
    except Exception:
        return None


def _rgb_to_hex(rgb):
    return "#{0:02X}{1:02X}{2:02X}".format(int(rgb[0]), int(rgb[1]), int(rgb[2]))


def _parse_rgb_text(text):
    value = (text or "").strip()
    if not value:
        return None

    if value.startswith("#"):
        hex_value = value[1:]
        if len(hex_value) != 6:
            return None
        try:
            return (
                int(hex_value[0:2], 16),
                int(hex_value[2:4], 16),
                int(hex_value[4:6], 16),
            )
        except Exception:
            return None

    parts = [part.strip() for part in value.split(",") if part.strip()]
    if len(parts) != 3:
        return None

    try:
        rgb = tuple(max(0, min(255, int(part))) for part in parts)
    except Exception:
        return None

    return rgb if len(rgb) == 3 else None


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


def _distance(p1, p2):
    dx = float(p1.X) - float(p2.X)
    dy = float(p1.Y) - float(p2.Y)
    dz = float(p1.Z) - float(p2.Z)
    return (dx * dx + dy * dy + dz * dz) ** 0.5


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
    return (curve_type, b, a, m)


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


def _is_effectively_line(curve):
    try:
        return curve.GetType().Name == "Line"
    except Exception:
        return False


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


def _curve_loop_polygon_xy(curve_loop):
    points = []
    try:
        for curve in curve_loop:
            p0 = curve.GetEndPoint(0)
            points.append((float(p0.X), float(p0.Y)))
    except Exception:
        return []
    return points


def _polygon_area_2d(points):
    if not points or len(points) < 3:
        return 0.0
    area2 = 0.0
    count = len(points)
    for idx in range(count):
        x1, y1 = points[idx]
        x2, y2 = points[(idx + 1) % count]
        area2 += (x1 * y2) - (x2 * y1)
    return abs(area2) * 0.5


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
    for items in buckets.values():
        if len(items) % 2 == 1:
            result.append(items[0])
    return result


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


def _group_connected_region_loop_records(region_loop_records, tol=0.01):
    if not region_loop_records:
        return []

    adjacency = {}
    key_to_indexes = {}
    segment_data = []
    for idx, record in enumerate(region_loop_records):
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
    for seg_idx, (region_idx, start_pt, end_pt) in enumerate(segment_data):
        split_params = [0.0, 1.0]
        for other_seg_idx, (_other_region_idx, other_start, other_end) in enumerate(segment_data):
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
            owners.add(region_idx)

    for owners in fragment_owner_map.values():
        if len(owners) < 2:
            continue

        owner_list = list(owners)
        for idx, region_idx in enumerate(owner_list):
            for other_region_idx in owner_list[idx + 1:]:
                adjacency[region_idx].add(other_region_idx)
                adjacency[other_region_idx].add(region_idx)

    groups = []
    visited = set()
    for idx in range(len(region_loop_records)):
        if idx in visited:
            continue

        stack = [idx]
        component = []
        visited.add(idx)
        while stack:
            current = stack.pop()
            component.append(region_loop_records[current])
            for neighbor in adjacency.get(current, set()):
                if neighbor in visited:
                    continue
                visited.add(neighbor)
                stack.append(neighbor)

        groups.append(component)

    return groups


def _build_boundary_sets_for_region_groups(region_loop_records, tol=0.02):
    boundary_sets = []

    for group in _group_connected_region_loop_records(region_loop_records, tol):
        if len(group) == 1:
            group_loops = list(group[0].get("loops", []))
            if not group_loops:
                group_loops = _build_closed_loops(group[0].get("curves", []), tol)
        else:
            group_curves = []
            for record in group:
                group_curves.extend(record.get("curves", []))

            outer_group_curves = _remove_shared_interior_segments(group_curves)
            group_loops = _build_closed_loops(outer_group_curves, tol)
            if not group_loops:
                group_loops = _build_closed_loops(group_curves, tol)

        if not group_loops:
            continue

        grouped_boundary_sets = _group_loops_into_boundary_sets(group_loops)
        if grouped_boundary_sets:
            boundary_sets.extend(grouped_boundary_sets)
            continue

        for loop in group_loops:
            boundary_sets.append([loop])

    return boundary_sets


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


def _create_filled_regions(view, region_type, boundary_sets, line_style_id=None):
    created_regions = []
    errors = []
    for set_index, boundary_set in enumerate(boundary_sets):
        ok_set, err_set = _try_create_region_boundaries(view, region_type, boundary_set, line_style_id)
        if not ok_set:
            errors.append({
                "set_index": set_index,
                "boundary_count": len(boundary_set),
                "message": "Boundary set create failed: {0}".format(err_set),
            })
            continue

        try:
            created_regions.append(_commit_region_with_boundaries(view, region_type, boundary_set, line_style_id))
        except Exception as ex:
            errors.append({
                "set_index": set_index,
                "boundary_count": len(boundary_set),
                "message": str(ex),
            })

    return created_regions, errors


def _flatten_boundary_sets(boundary_sets):
    flattened = []
    for boundary_set in boundary_sets or []:
        for loop in boundary_set or []:
            flattened.append(loop)
    return flattened


def _resolve_region_type(base_type, create_new_type, new_type_name):
    if not create_new_type or not new_type_name:
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
        forms.alert("Failed to create Filled Region Type: {0}".format(ex), title="Combine Filled Regions")
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


class SelectedRegionItem(object):
    def __init__(self, region):
        self.region = region
        self.region_id = region.Id
        self.region_key = region.Id.IntegerValue
        self.type_id = region.GetTypeId()
        self.type_name = "Filled Region Type"
        self.owner_view_name = "Active View"
        self.line_style_id = _get_filled_region_line_style_id(region)

        try:
            region_type = doc.GetElement(self.type_id)
            self.type_name = _safe_elem_name(region_type, self.type_name)
        except Exception:
            pass

        try:
            owner_view = doc.GetElement(region.OwnerViewId)
            if owner_view is not None:
                self.owner_view_name = _safe_elem_name(owner_view, self.owner_view_name)
        except Exception:
            pass

        self.display = "ID {0}  |  {1}  |  {2}".format(self.region_key, self.type_name, self.owner_view_name)


class CombineFilledRegionsWindow(WPFWindow):
    def __init__(self, xaml_path, selected_items):
        WPFWindow.__init__(self, xaml_path)
        self._selected_items = list(selected_items or [])
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
        self._is_updating_color_ui = False
        self.result = None

        self._load_region_types()
        self._load_fill_patterns()
        self._load_line_styles()
        self._load_selected_regions()
        self._apply_selection_defaults()
        self._refresh_color_ui()
        self._refresh_line_style_ui()
        self._update_summary()

    def _load_region_types(self):
        self._region_type_items = []
        self.cmbRegionTypes.Items.Clear()

        region_types = FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements()
        for fr_type in sorted(region_types, key=lambda t: _safe_elem_name(t, "").lower()):
            item = RegionTypeItem(fr_type)
            self._region_type_items.append(item)
            self.cmbRegionTypes.Items.Add(item)

        self.cmbRegionTypes.DisplayMemberPath = "name"
        if self._region_type_items:
            self.cmbRegionTypes.SelectedIndex = 0

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

    def _load_selected_regions(self):
        self.lstRegions.Items.Clear()
        for item in self._selected_items:
            self.lstRegions.Items.Add(item)
        self.lstRegions.DisplayMemberPath = "display"

    def _apply_selection_defaults(self):
        if not self._selected_items:
            return

        first_item = self._selected_items[0]
        preferred_type_key = _serialize_element_id(first_item.type_id)
        for idx, item in enumerate(self._region_type_items):
            if _serialize_element_id(item.fr_type.Id) == preferred_type_key:
                self.cmbRegionTypes.SelectedIndex = idx
                break

        if first_item.line_style_id is not None:
            self._boundary_line_style_id = first_item.line_style_id

        base_type = doc.GetElement(first_item.type_id)
        graphics = _read_region_type_graphics(base_type)
        self._foreground_rgb = graphics["foreground_rgb"]
        self._background_rgb = graphics["background_rgb"]
        self._foreground_pattern_id = graphics["foreground_pattern_id"]
        self._background_pattern_id = graphics["background_pattern_id"]
        self._show_foreground = graphics["show_foreground"]
        self._show_background = graphics["show_background"]
        self._is_masking = _read_region_type_is_masking(base_type)
        self.chkShowForeground.IsChecked = self._show_foreground
        self.chkShowBackground.IsChecked = self._show_background
        self.chkMasking.IsChecked = self._is_masking
        self.txtNewTypeName.Text = "{0} - Combined".format(first_item.type_name)

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

    def _update_summary(self):
        count = len(self._selected_items)
        type_names = sorted(set(item.type_name for item in self._selected_items))
        source_summary = type_names[0] if len(type_names) == 1 else "Mixed source types ({0})".format(len(type_names))
        self.txtSelectionSummary.Text = "Selected filled regions in the active view: {0}. Source type: {1}.".format(count, source_summary)

        selected_type_item = self.cmbRegionTypes.SelectedItem
        target_type_name = selected_type_item.name if selected_type_item is not None else "None"
        if bool(self.chkCreateNewType.IsChecked):
            target_type_name = (self.txtNewTypeName.Text or "").strip() or target_type_name

        self.txtSummary.Text = (
            "The selected filled regions will be replaced after the merged region is created successfully. "
            "Target type: {0}. Boundary line style: {1}."
        ).format(target_type_name, _line_style_name(doc.GetElement(self._get_selected_line_style_id())))

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
                return (int(picked.R), int(picked.G), int(picked.B))
        except Exception as ex:
            logger.debug("Color dialog failed: {0}".format(ex))

        color_value = forms.ask_for_string(
            default=_rgb_to_hex(current_rgb),
            prompt="Enter color as #RRGGBB or R,G,B.",
            title=title,
        )
        if not color_value:
            return None

        parsed_rgb = _parse_rgb_text(color_value)
        if parsed_rgb is None:
            forms.alert("Invalid color value. Use #RRGGBB or R,G,B.", title="Combine Filled Regions")
            return None

        return parsed_rgb

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

    def on_region_type_changed(self, sender, args):
        self._load_selected_region_type_graphics()
        self._refresh_color_ui()
        self._update_summary()

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
        self._boundary_line_style_id = self._get_selected_line_style_id()
        self._update_summary()

    def on_foreground_pattern_changed(self, sender, args):
        if self._is_updating_color_ui:
            return
        self._foreground_pattern_id = self._get_selected_pattern_id(self.cmbForegroundPattern)

    def on_background_pattern_changed(self, sender, args):
        if self._is_updating_color_ui:
            return
        self._background_pattern_id = self._get_selected_pattern_id(self.cmbBackgroundPattern)

    def on_foreground_text_changed(self, sender, args):
        self._apply_color_text_input(self.txtForegroundColor, "_foreground_rgb")

    def on_background_text_changed(self, sender, args):
        self._apply_color_text_input(self.txtBackgroundColor, "_background_rgb")

    def on_pick_foreground(self, sender, args):
        picked = self._pick_color(self._foreground_rgb, "Set Foreground Color")
        if picked is not None:
            self._foreground_rgb = picked
            self._refresh_color_ui()

    def on_pick_background(self, sender, args):
        picked = self._pick_color(self._background_rgb, "Set Background Color")
        if picked is not None:
            self._background_rgb = picked
            self._refresh_color_ui()

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
                forms.alert("No Filled Region Type is available.", title="Combine Filled Regions")
                return

            if not self._selected_items:
                forms.alert("Select at least one filled region before running the tool.", title="Combine Filled Regions")
                return

            create_new = bool(self.chkCreateNewType.IsChecked)
            new_name = (self.txtNewTypeName.Text or "").strip()
            if create_new and not new_name:
                forms.alert("Enter a new Filled Region Type name.", title="Combine Filled Regions")
                return

            self.result = {
                "selected_region_ids": [item.region_id for item in self._selected_items],
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
            }

            try:
                self.DialogResult = True
            except Exception:
                pass
            self.Close()
        except Exception as ex:
            logger.debug("Create button failed: {0}".format(ex))
            forms.alert("Failed to prepare the combine request.\n\n{0}".format(ex), title="Combine Filled Regions")


def _get_selected_regions(view):
    selected_items = []
    seen = set()
    try:
        selected_ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        selected_ids = []

    for element_id in selected_ids:
        region = doc.GetElement(element_id)
        if region is None:
            continue
        try:
            if not isinstance(region, FilledRegion):
                continue
        except Exception:
            continue

        try:
            owner_view_id = getattr(region, "OwnerViewId", None)
            if owner_view_id is not None and owner_view_id != ElementId.InvalidElementId and owner_view_id != view.Id:
                continue
        except Exception:
            pass

        key = _serialize_element_id(region.Id)
        if key in seen:
            continue
        seen.add(key)
        selected_items.append(SelectedRegionItem(region))

    return selected_items


def _collect_region_boundary_curves(region_items, target_elevation):
    region_loop_records = []
    failed_regions = []

    for item in region_items:
        try:
            boundaries = list(item.region.GetBoundaries())
        except Exception as ex:
            failed_regions.append("ID {0}: {1}".format(item.region_key, ex))
            continue

        region_curves = []
        region_loops = []
        for curve_loop in boundaries:
            try:
                curves = list(curve_loop)
            except Exception:
                curves = []

            for curve in curves:
                region_curves.extend(_curve_to_planar_segments(curve, target_elevation))

        if len(region_curves) < 3:
            failed_regions.append("ID {0}: not enough valid boundary segments".format(item.region_key))
            continue

        for curve_loop in boundaries:
            planar_loop_curves = []
            try:
                ordered_curves = list(curve_loop)
            except Exception:
                ordered_curves = []

            for curve in ordered_curves:
                planar_loop_curves.extend(_curve_to_planar_segments(curve, target_elevation))

            if len(planar_loop_curves) < 3:
                continue

            rebuilt_loop = _curve_loop_from_curves(planar_loop_curves)
            if rebuilt_loop is not None:
                region_loops.append(rebuilt_loop)
                continue

            rebuilt_loops = _build_closed_loops(planar_loop_curves)
            if rebuilt_loops:
                region_loops.extend(rebuilt_loops)

        region_loop_records.append({
            "region_item": item,
            "curves": region_curves,
            "loops": region_loops,
        })

    return region_loop_records, failed_regions


def _delete_regions(region_ids):
    try:
        doc.Delete(List[ElementId](region_ids))
        return
    except Exception:
        pass

    for region_id in region_ids:
        try:
            doc.Delete(region_id)
        except Exception:
            continue


def _execute_combine_regions_operation(view, data):
    target_elev = _active_view_plane_elevation(view)
    selected_region_ids = data.get("selected_region_ids") or []
    region_items = []
    for region_id in selected_region_ids:
        region = doc.GetElement(region_id)
        if region is None:
            continue
        try:
            if not isinstance(region, FilledRegion):
                continue
        except Exception:
            continue
        region_items.append(SelectedRegionItem(region))

    if not region_items:
        forms.alert("No selected filled regions were available in the active view.", title="Combine Filled Regions")
        return False

    region_loop_records, failed_regions = _collect_region_boundary_curves(region_items, target_elev)
    if failed_regions:
        forms.alert(
            "One or more selected filled regions could not be read.\n\n{0}".format("\n".join(failed_regions)),
            title="Combine Filled Regions",
        )
        return False

    if not region_loop_records:
        forms.alert("The selected filled regions did not produce enough valid boundary segments.", title="Combine Filled Regions")
        return False

    boundary_sets = _build_boundary_sets_for_region_groups(region_loop_records)
    if not boundary_sets:
        forms.alert("Failed to build closed merged boundaries from the selected filled regions.", title="Combine Filled Regions")
        return False

    t = Transaction(doc, "Combine Filled Regions")
    try:
        t.Start()
        region_type = _resolve_region_type(
            data.get("base_type"),
            data.get("create_new_type"),
            data.get("new_type_name"),
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
                data.get("foreground_pattern_id"),
                data.get("background_pattern_id"),
                data.get("show_foreground", True),
                data.get("show_background", True),
            )

        created_regions = []
        errors = []
        line_style_id = data.get("boundary_line_style_id") or _invisible_line_style_id()

        combined_boundaries = _flatten_boundary_sets(boundary_sets)
        combined_ok, combined_error = _try_create_region_boundaries(
            view,
            region_type,
            combined_boundaries,
            line_style_id,
        )
        if combined_ok:
            try:
                created_regions.append(_commit_region_with_boundaries(view, region_type, combined_boundaries, line_style_id))
            except Exception as ex:
                errors.append({
                    "message": str(ex),
                    "boundary_count": len(combined_boundaries),
                    "set_index": 0,
                })
        else:
            if combined_error is not None:
                errors.append({
                    "message": "Single combined region rejected by Revit: {0}".format(combined_error),
                    "boundary_count": len(combined_boundaries),
                    "set_index": 0,
                })

            fallback_created, fallback_errors = _create_filled_regions(
                view,
                region_type,
                boundary_sets,
                line_style_id,
            )
            created_regions.extend(fallback_created)
            errors.extend(fallback_errors)

        if not created_regions:
            t.RollBack()
            first_error = errors[0]["message"] if errors else "Unknown FilledRegion.Create failure."
            forms.alert(
                "Failed to create a stable merged filled region.\n\n{0}".format(first_error),
                title="Combine Filled Regions",
            )
            return False

        _delete_regions([item.region_id for item in region_items])
        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        forms.alert("Failed while combining filled regions: {0}".format(ex), title="Combine Filled Regions")
        return False

    forms.alert(
        "Created {0} merged filled region(s) from {1} selected region(s).{2}".format(
            len(created_regions),
            len(region_items),
            "\n\nRevit rejected the single combined region, so the tool fell back to separate created groups." if (errors and len(created_regions) > 1) else ("\n\nSome boundary groups were skipped because Revit rejected them." if errors else ""),
        ),
        title="Combine Filled Regions",
    )
    return True


def main():
    view = doc.ActiveView
    if view is None:
        forms.alert("No active view found.", title="Combine Filled Regions")
        return

    selected_items = _get_selected_regions(view)
    if not selected_items:
        forms.alert(
            "Select one or more filled regions in the active view before running this tool.",
            title="Combine Filled Regions",
        )
        return

    xaml_file = script.get_bundle_file("CombineFilledRegionsWindow.xaml")
    win = None
    try:
        win = CombineFilledRegionsWindow(xaml_file, selected_items)
        win.ShowDialog()
    except Exception as ex:
        forms.alert("Failed to open Combine Filled Regions window: {0}".format(ex), title="Combine Filled Regions")
        return

    if win is None or not win.result:
        return

    _execute_combine_regions_operation(view, win.result)


if __name__ == "__main__":
    main()