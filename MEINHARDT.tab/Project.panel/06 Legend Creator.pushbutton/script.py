# -*- coding: utf-8 -*-
"""Legend Creator

Copies/extends pyChilizer's Populate Legend concept and adds:
- All MEP categories (text legend)
- Color Fill Scheme legend (filled region + text)
- Spaces legend (text)

Author: GM
"""

from __future__ import print_function

from collections import OrderedDict

from pyrevit import revit, DB, forms, script
from pyrevit.framework import List
from Autodesk.Revit import Exceptions

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

_PREFERRED_LABEL_PARAM = None

BIC = DB.BuiltInCategory


# -----------------------------
# Units + basic graphics helpers
# -----------------------------

def _mm_to_internal(mm_val):
    try:
        mm_val = float(mm_val)
    except Exception:
        mm_val = 0.0

    # Revit 2022+ (ForgeTypeId)
    try:
        return DB.UnitUtils.ConvertToInternalUnits(mm_val, DB.UnitTypeId.Millimeters)
    except Exception:
        pass

    # Older API fallback
    try:
        return DB.UnitUtils.ConvertToInternalUnits(mm_val, DB.DisplayUnitType.DUT_MILLIMETERS)
    except Exception:
        return mm_val


def _invis_style():
    # Invisible lines graphics style category id is -2000064
    try:
        for gs in DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle):
            try:
                if gs.GraphicsStyleCategory.Id.IntegerValue == -2000064:
                    return gs
            except Exception:
                continue
    except Exception:
        pass
    return None


def _any_filled_region_type():
    return DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).FirstElement()


def _solid_fill_pattern_id():
    # Prefer API helper if available
    try:
        fp = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Drafting, "<Solid fill>")
        if fp:
            return fp.Id
    except Exception:
        pass

    # Fallback: search by pattern name
    try:
        for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
            try:
                if fp.GetFillPattern().IsSolidFill:
                    return fp.Id
            except Exception:
                continue
    except Exception:
        pass

    return DB.ElementId.InvalidElementId


def _default_text_type_id():
    try:
        return doc.GetDefaultElementTypeId(DB.ElementTypeGroup.TextNoteType)
    except Exception:
        # last resort
        t = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).FirstElement()
        return t.Id if t else DB.ElementId.InvalidElementId


def _text_note_types_dict():
    d = {}
    try:
        for t in DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType):
            try:
                nm = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            except Exception:
                nm = None
            nm = nm or getattr(t, 'Name', None) or 'Text'
            d[nm] = t
    except Exception:
        pass
    return d


def _make_rectangle(pt, width, height):
    p1 = DB.XYZ(pt.X, pt.Y, 0)
    p2 = DB.XYZ(pt.X + width, pt.Y, 0)
    p3 = DB.XYZ(pt.X + width, pt.Y + height, 0)
    p4 = DB.XYZ(pt.X, pt.Y + height, 0)
    l1 = DB.Line.CreateBound(p1, p2)
    l2 = DB.Line.CreateBound(p2, p3)
    l3 = DB.Line.CreateBound(p3, p4)
    l4 = DB.Line.CreateBound(p4, p1)
    return [l1, l2, l3, l4]


def _translate_curves(curves, dx=0.0, dy=0.0):
    v = DB.XYZ(dx, dy, 0)
    t = DB.Transform.CreateTranslation(v)
    return [c.CreateTransformed(t) for c in curves]


def _draw_filled_box(view, base_rect, y_offset, fill_type_id, line_style_id):
    curves = _translate_curves(base_rect, dx=0.0, dy=-y_offset)
    loop = DB.CurveLoop.Create(List[DB.Curve](curves))
    region = DB.FilledRegion.Create(doc, fill_type_id, view.Id, [loop])
    if line_style_id and line_style_id != DB.ElementId.InvalidElementId:
        region.SetLineStyleId(line_style_id)
    return region


def _ensure_unique_view_name(name):
    existing = set()
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType():
        try:
            existing.add(v.Name)
        except Exception:
            pass

    if name not in existing:
        return name

    i = 2
    while True:
        cand = "{} ({})".format(name, i)
        if cand not in existing:
            return cand
        i += 1


# -----------------------------
# View creation/targeting
# -----------------------------

def _get_any_legend_view():
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType():
        try:
            if v.ViewType == DB.ViewType.Legend and not v.IsTemplate:
                return v
        except Exception:
            continue
    return None


def _legend_has_component(view):
    try:
        lc = DB.FilteredElementCollector(doc, view.Id).OfCategory(BIC.OST_LegendComponents).FirstElement()
        return lc is not None
    except Exception:
        return False


def _pick_template_legend_view():
    # Prefer a legend view that already contains a legend component.
    legends = []
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType():
        try:
            if v.ViewType == DB.ViewType.Legend and not v.IsTemplate:
                legends.append(v)
        except Exception:
            continue

    if not legends:
        forms.alert(
            "No Legend views exist in this model. Create a Legend view manually once, then retry.",
            exitscript=True,
        )

    with_comp = [v for v in legends if _legend_has_component(v)]
    if with_comp:
        # Stable pick: first by name
        return sorted(with_comp, key=lambda x: x.Name)[0]

    picked = forms.SelectFromList.show(
        sorted(legends, key=lambda x: x.Name),
        name_attr='Name',
        multiselect=False,
        title='Pick Template Legend View',
        button_name='Select'
    )
    if not picked:
        script.exit()

    if not _legend_has_component(picked):
        forms.alert(
            "That legend has no Legend Components to use as a template.\n\n"
            "Open that legend view and place ONE Legend Component using:"
            "\nAnnotate → Component → Legend Component\n\n"
            "Then run Legend Creator again.",
            exitscript=True,
        )

    return picked


def _get_drafting_viewfamilytype_id():
    for vft in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        try:
            if vft.ViewFamily == DB.ViewFamily.Drafting:
                return vft.Id
        except Exception:
            continue
    return DB.ElementId.InvalidElementId


def _activate_view(view):
    try:
        uidoc.ActiveView = view
        return True
    except Exception:
        return False


def _get_target_view(target_mode, base_name):
    av = revit.active_view

    if target_mode == "Current View":
        return av

    if target_mode == "New Drafting View":
        vft_id = _get_drafting_viewfamilytype_id()
        if vft_id == DB.ElementId.InvalidElementId:
            forms.alert("Could not find Drafting ViewFamilyType.", exitscript=True)
        with revit.Transaction("Create Drafting View"):
            dv = DB.ViewDrafting.Create(doc, vft_id)
            dv.Name = _ensure_unique_view_name(base_name)
        _activate_view(dv)
        return dv

    # New Legend View via duplicate (Revit API has no simple Create for legends in many versions)
    base = av if av.ViewType == DB.ViewType.Legend and _legend_has_component(av) else _pick_template_legend_view()

    with revit.Transaction("Create Legend View"):
        new_id = base.Duplicate(DB.ViewDuplicateOption.Duplicate)
        lv = doc.GetElement(new_id)
        lv.Name = _ensure_unique_view_name(base_name)
    return lv


# -----------------------------
# Data collection
# -----------------------------

def _get_category_name(bic):
    try:
        return DB.LabelUtils.GetLabelFor(bic)
    except Exception:
        return str(bic)


def _collect_symbols_for_category(bic, only_used, used_scope_view):
    # Types
    syms = list(
        DB.FilteredElementCollector(doc)
        .OfCategory(bic)
        .WhereElementIsElementType()
        .ToElements()
    )

    if not only_used:
        return syms

    # Used scope: active view only, else entire model
    if used_scope_view is not None:
        insts = DB.FilteredElementCollector(doc, used_scope_view.Id).OfCategory(bic).WhereElementIsNotElementType()
    else:
        insts = DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()

    used_type_ids = set()
    for inst in insts:
        try:
            tid = inst.GetTypeId()
            if tid and tid != DB.ElementId.InvalidElementId:
                used_type_ids.add(tid)
        except Exception:
            continue

    return [s for s in syms if s.Id in used_type_ids]


def _group_symbols_by_family(symbols):
    ordered = OrderedDict()
    for sym in symbols:
        try:
            fam = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
        except Exception:
            fam = None
        try:
            typ = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        except Exception:
            typ = None

        fam = fam or "<Family>"
        typ = typ or "<Type>"

        if fam not in ordered:
            ordered[fam] = OrderedDict()
        ordered[fam][typ] = sym

    # deterministic: sort families, then types
    sorted_out = OrderedDict()
    for fam in sorted(ordered.keys()):
        type_map = ordered[fam]
        sorted_out[fam] = OrderedDict()
        for typ in sorted(type_map.keys()):
            sorted_out[fam][typ] = type_map[typ]
    return sorted_out


# -----------------------------
# Legend creation: Family Symbols (Legend Components) or Text
# -----------------------------

def _category_view_directions(bic):
    # Copied from pyChilizer Populate Legend logic
    if bic in [BIC.OST_Walls]:
        return {"Section": -5, "Floor Plan": -8}
    if bic in [BIC.OST_Roofs, BIC.OST_Ceilings, BIC.OST_Floors]:
        return {"Section": -5}
    if bic in [BIC.OST_Windows, BIC.OST_Doors]:
        return {"Back": -6, "Front": -7, "Floor Plan": -8}
    if bic in [
        BIC.OST_GenericModel,
        BIC.OST_Casework,
        BIC.OST_ElectricalEquipment,
        BIC.OST_ElectricalFixtures,
        BIC.OST_Furniture,
        BIC.OST_FurnitureSystems,
        BIC.OST_PlumbingFixtures,
        BIC.OST_Entourage,
        BIC.OST_MechanicalEquipment,
    ]:
        return {"Back": -6, "Front": -7, "Floor Plan": -8, "Right": -9, "Left": -10}

    return {"Section": -5, "Back": -6, "Front": -7, "Floor Plan": -8, "Right": -9, "Left": -10}


def _pick_point_in_active_view(title, restore_view=None):
    with forms.WarningBar(title=title):
        try:
            return uidoc.Selection.PickPoint()
        except Exceptions.OperationCanceledException:
            if restore_view is not None:
                try:
                    uidoc.ActiveView = restore_view
                except Exception:
                    pass
            forms.alert("Cancelled", ok=True, exitscript=True)


def _place_text_list(view, pt, lines, line_height_mm=6.0, text_type_id=None):
    ttype_id = text_type_id or _default_text_type_id()
    if ttype_id == DB.ElementId.InvalidElementId:
        forms.alert("No TextNoteType found.", exitscript=True)

    line_height = _mm_to_internal(line_height_mm) * (float(view.Scale) / 100.0)

    with revit.Transaction("Place Legend Text"):
        y = 0.0
        for txt in lines:
            pos = DB.XYZ(pt.X, pt.Y - y, 0)
            DB.TextNote.Create(doc, view.Id, pos, txt, ttype_id)
            y += line_height


def _family_label(symbol):
    try:
        fam = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
        if fam:
            return fam
    except Exception:
        pass
    try:
        return symbol.FamilyName
    except Exception:
        return None


def _populate_legend_components(view, bic, ordered_symbols, view_direction, pt, spacing_internal, text_type_id):
    # Need at least one source legend component in view to copy
    source_lc = DB.FilteredElementCollector(doc, view.Id).OfCategory(BIC.OST_LegendComponents).FirstElement()
    forms.alert_ifnot(
        source_lc,
        "The target legend must have at least one source Legend Component to copy. Place any Legend Component once, then retry.",
        exitscript=True,
    )

    y_offset = 0.0
    ttype_id = text_type_id or _default_text_type_id()
    text_offset = _mm_to_internal(100.0) * (float(view.Scale) / 100.0)

    with revit.Transaction("Populate Legend"):
        for fam in ordered_symbols:
            # One type per family (first type)
            try:
                symbol = next(iter(ordered_symbols[fam].values()))
            except Exception:
                continue

            try:
                # Copy in place then move to an explicit row point.
                copy_id = DB.ElementTransformUtils.CopyElement(doc, source_lc.Id, DB.XYZ(0, 0, 0))[0]
                new_lc = doc.GetElement(copy_id)
                new_lc.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT).Set(symbol.Id)
                new_lc.get_Parameter(DB.BuiltInParameter.LEGEND_COMPONENT_VIEW).Set(view_direction)
                doc.Regenerate()

                bb = new_lc.get_BoundingBox(view)
                if not bb:
                    # still advance to avoid infinite overlap
                    y_offset += spacing_internal
                    continue

                row_pt = DB.XYZ(pt.X, pt.Y - y_offset, 0)
                move_vec = row_pt - DB.XYZ(bb.Min.X, bb.Min.Y, 0)
                DB.ElementTransformUtils.MoveElement(doc, copy_id, move_vec)
                doc.Regenerate()

                bb = new_lc.get_BoundingBox(view)
                if bb and ttype_id and ttype_id != DB.ElementId.InvalidElementId:
                    label = _family_label(symbol)
                    if label:
                        label_pos = DB.XYZ(bb.Max.X + text_offset, (bb.Min.Y + bb.Max.Y) / 2.0, 0)
                        DB.TextNote.Create(doc, view.Id, label_pos, label, ttype_id)

                # Next row (down)
                bb_h = (bb.Max.Y - bb.Min.Y) if bb else 0.0
                y_offset += (bb_h + spacing_internal)
            except Exception:
                continue


    


# -----------------------------
# Legend creation: Color Fill Scheme
# -----------------------------

def _collect_color_fill_schemes():
    try:
        return list(DB.FilteredElementCollector(doc).OfClass(DB.ColorFillScheme).ToElements())
    except Exception:
        return []


def _scheme_category_name(scheme):
    try:
        cat = doc.Settings.Categories.get_Item(scheme.CategoryId)
        if cat:
            return cat.Name
    except Exception:
        pass
    try:
        return str(scheme.CategoryId)
    except Exception:
        return "Category"


def _scheme_entries(scheme):
    try:
        return list(scheme.GetEntries() or [])
    except Exception:
        return []


def _entry_label(entry, idx):
    for attr in ["Caption", "Value", "StringValue", "Name"]:
        try:
            v = getattr(entry, attr)
            if v:
                return str(v)
        except Exception:
            pass
    return "Entry {}".format(idx + 1)


def _entry_color(entry):
    try:
        return entry.Color
    except Exception:
        return None


def _entry_visible(entry):
    for attr in ["IsVisible", "Visible"]:
        try:
            v = getattr(entry, attr)
            return bool(v)
        except Exception:
            continue
    return True


def _draw_color_scheme_legend(view, pt, scheme, only_visible=True, box_w_mm=1000, box_h_mm=240, box_off_mm=80):
    entries = _scheme_entries(scheme)
    if only_visible:
        entries = [e for e in entries if _entry_visible(e)]

    if not entries:
        forms.alert("No entries found for the selected Color Fill Scheme.", exitscript=True)

    scale = float(view.Scale) / 100.0
    w = _mm_to_internal(box_w_mm) * scale
    h = _mm_to_internal(box_h_mm) * scale
    text_offset = 1 * scale
    shift = _mm_to_internal(box_off_mm + box_h_mm) * scale

    base_rect = _make_rectangle(pt, w, h)

    line_style = _invis_style()
    line_style_id = line_style.Id if line_style else DB.ElementId.InvalidElementId

    fr_type = _any_filled_region_type()
    forms.alert_ifnot(fr_type, "No FilledRegionType found in this model.", exitscript=True)

    solid_pat_id = _solid_fill_pattern_id()
    if solid_pat_id == DB.ElementId.InvalidElementId:
        # still draw, but pattern might be missing in very unusual templates
        solid_pat_id = fr_type.ForegroundPatternId

    ttype_id = _default_text_type_id()

    with revit.Transaction("Draw Color Scheme Legend"):
        offset = 0.0
        for idx, entry in enumerate(entries):
            color = _entry_color(entry)
            label = _entry_label(entry, idx)

            reg = _draw_filled_box(view, base_rect, offset, fr_type.Id, line_style_id)

            if color:
                ogs = DB.OverrideGraphicSettings()
                ogs.SetSurfaceForegroundPatternColor(color)
                ogs.SetSurfaceForegroundPatternId(solid_pat_id)
                view.SetElementOverrides(reg.Id, ogs)

            label_pos = DB.XYZ(pt.X + w + text_offset, pt.Y - (offset - h), 0)
            DB.TextNote.Create(doc, view.Id, label_pos, label, ttype_id)

            offset += shift


# -----------------------------
# Legend creation: Spaces
# -----------------------------

def _space_label(space):
    num = None
    name = None

    # Spaces use ROOM_* parameters in some versions; SPACE_* may not exist.
    candidates_number = []
    for attr in ['ROOM_NUMBER', 'SPACE_NUMBER']:
        try:
            candidates_number.append(getattr(DB.BuiltInParameter, attr))
        except Exception:
            pass

    candidates_name = []
    for attr in ['ROOM_NAME', 'SPACE_NAME']:
        try:
            candidates_name.append(getattr(DB.BuiltInParameter, attr))
        except Exception:
            pass

    for bip in candidates_number:
        try:
            p = space.get_Parameter(bip)
            if p and p.AsString():
                num = p.AsString()
                break
        except Exception:
            continue

    for bip in candidates_name:
        try:
            p = space.get_Parameter(bip)
            if p and p.AsString():
                name = p.AsString()
                break
        except Exception:
            continue

    # Final fallback: common parameter names
    if not num:
        try:
            p = space.LookupParameter('Number')
            if p and p.AsString():
                num = p.AsString()
        except Exception:
            pass
    if not name:
        try:
            p = space.LookupParameter('Name')
            if p and p.AsString():
                name = p.AsString()
        except Exception:
            pass

    if num and name:
        return "{} - {}".format(num, name)
    if name:
        return name
    if num:
        return num
    try:
        return space.Name
    except Exception:
        return "Space"


def _collect_spaces(only_used, used_scope_view):
    bic_spaces = BIC.OST_MEPSpaces

    if only_used and used_scope_view is not None:
        col = DB.FilteredElementCollector(doc, used_scope_view.Id).OfCategory(bic_spaces).WhereElementIsNotElementType()
    else:
        col = DB.FilteredElementCollector(doc).OfCategory(bic_spaces).WhereElementIsNotElementType()

    return list(col.ToElements())


# -----------------------------
# UI + main
# -----------------------------

def _common_categories():
    # Start with pyChilizer's list + key MEP categories.
    cats = [
        BIC.OST_Windows,
        BIC.OST_Doors,
        BIC.OST_Floors,
        BIC.OST_Walls,
        BIC.OST_GenericModel,
        BIC.OST_Casework,
        BIC.OST_Furniture,
        BIC.OST_FurnitureSystems,
        BIC.OST_PlumbingFixtures,
        BIC.OST_Roofs,
        BIC.OST_ElectricalEquipment,
        BIC.OST_ElectricalFixtures,
        BIC.OST_Ceilings,
        # MEP
        BIC.OST_MechanicalEquipment,
        BIC.OST_DuctAccessory,
        BIC.OST_DuctFitting,
        BIC.OST_DuctTerminal,
        BIC.OST_PipeAccessory,
        BIC.OST_PipeFitting,
        BIC.OST_Sprinklers,
        BIC.OST_LightingFixtures,
        BIC.OST_CommunicationDevices,
        BIC.OST_DataDevices,
        BIC.OST_SecurityDevices,
        BIC.OST_FireAlarmDevices,
        BIC.OST_CableTrayFitting,
        BIC.OST_ConduitFitting,
    ]

    out = {}
    for c in cats:
        out[_get_category_name(c)] = c
    # de-dupe by name
    return out


def main():
    orig_view = revit.active_view
    modes = [
        "Family Legend (single category)",
        "Family Legend (All MEP categories)",
        "Legend from Color Fill Scheme",
        "Legend from Spaces",
    ]

    target_modes = [
        "Current View",
        "New Drafting View",
        "New Legend View",
    ]

    used_modes = [
        "All (ignore view usage)",
        "Only used in Active View",
    ]

    cat_dict = _common_categories()
    schemes = _collect_color_fill_schemes()
    scheme_dict = {}
    for s in schemes:
        try:
            disp = "{} | {}".format(_scheme_category_name(s), s.Name)
        except Exception:
            disp = "Color Fill Scheme"
        scheme_dict[disp] = s

    # Quick form (kept simple, consistent with existing scripts)
    from rpw.ui.forms import FlexForm, Label, ComboBox, Button, TextBox

    text_types = _text_note_types_dict()
    forms.alert_ifnot(text_types, "No TextNoteTypes found in this model.", exitscript=True)

    components = [
        Label("Mode"),
        ComboBox("mode", options={m: m for m in modes}),
        Label("Target"),
        ComboBox("target", options={t: t for t in target_modes}),
        Label("Content"),
        ComboBox("used", options={u: u for u in used_modes}),
        Label("Text Style"),
        ComboBox("textstyle", options=text_types),
        Label("Category (for single-category family legend)"),
        ComboBox("cat", options=cat_dict),
        Label("Color Fill Scheme (for scheme legend)"),
        ComboBox("scheme", options=scheme_dict),
        Label("Box Width [mm] (scheme legend)"),
        TextBox("box_w", Text="1000"),
        Label("Box Height [mm] (scheme legend)"),
        TextBox("box_h", Text="240"),
        Label("Offset [mm] (scheme legend)"),
        TextBox("box_off", Text="80"),
        Button("Create"),
    ]

    form = FlexForm("Legend Creator", components)
    ok = form.show()
    if not ok:
        return

    mode = form.values.get("mode")
    target_mode = form.values.get("target")
    used_mode = form.values.get("used")

    only_used = used_mode == "Only used in Active View"
    used_scope_view = revit.active_view if only_used else None

    chosen_text_type = form.values.get("textstyle")
    chosen_text_type_id = chosen_text_type.Id if chosen_text_type else _default_text_type_id()

    base_name = "GM - Legend"
    if mode == "Family Legend (single category)":
        # Smart default: family legends must be created in a Legend view.
        target_mode = "New Legend View"
        bic = form.values.get("cat")
        if bic is None:
            forms.alert("Pick a category.", exitscript=True)
        base_name = "GM - Family Legend - {}".format(_get_category_name(bic))

        target_view = _get_target_view(target_mode, base_name)
        _activate_view(target_view)
        pt = _pick_point_in_active_view("Pick Placement Point", restore_view=orig_view)

        symbols = _collect_symbols_for_category(bic, only_used=only_used, used_scope_view=used_scope_view)
        if not symbols:
            forms.alert("No family types found for that category.", exitscript=True)

        ordered = _group_symbols_by_family(symbols)

        if target_view.ViewType == DB.ViewType.Legend:
            # Ask view direction (copied from pyChilizer Populate Legend)
            view_dirs = _category_view_directions(bic)
            view_dir = forms.SelectFromList.show(
                sorted(view_dirs.keys()),
                title="View Direction",
                button_name="Select",
                multiselect=False,
            )
            if not view_dir:
                return
            dir_code = view_dirs[view_dir]

            spacing_internal = _mm_to_internal(50.0) * (float(target_view.Scale) / 100.0)
            _populate_legend_components(
                target_view,
                bic,
                ordered,
                dir_code,
                pt,
                spacing_internal=spacing_internal,
                text_type_id=chosen_text_type_id,
            )
        else:
            # Drafting/current view: text-only (legend components are legend-only)
            lines = []
            for fam in ordered:
                lines.append(fam)
                # one type per family
                try:
                    first_typ = next(iter(ordered[fam].keys()))
                    lines.append("    {}".format(first_typ))
                except Exception:
                    continue
            _place_text_list(target_view, pt, lines, text_type_id=chosen_text_type_id)

        return

    if mode == "Family Legend (All MEP categories)":
        mep_cats = [
            BIC.OST_MechanicalEquipment,
            BIC.OST_DuctTerminal,
            BIC.OST_DuctFitting,
            BIC.OST_DuctAccessory,
            BIC.OST_PipeFitting,
            BIC.OST_PipeAccessory,
            BIC.OST_PlumbingFixtures,
            BIC.OST_Sprinklers,
            BIC.OST_ElectricalEquipment,
            BIC.OST_ElectricalFixtures,
            BIC.OST_LightingFixtures,
            BIC.OST_CommunicationDevices,
            BIC.OST_DataDevices,
            BIC.OST_SecurityDevices,
            BIC.OST_FireAlarmDevices,
            BIC.OST_CableTrayFitting,
            BIC.OST_ConduitFitting,
        ]

        base_name = "GM - MEP Family Legend"
        target_view = _get_target_view(target_mode, base_name)
        _activate_view(target_view)
        pt = _pick_point_in_active_view("Pick Placement Point", restore_view=orig_view)

        # Text-only for multi-category (reliable across view types)
        lines = []
        for bic in mep_cats:
            try:
                symbols = _collect_symbols_for_category(bic, only_used=only_used, used_scope_view=used_scope_view)
            except Exception:
                symbols = []
            if not symbols:
                continue

            ordered = _group_symbols_by_family(symbols)
            lines.append("[{}]".format(_get_category_name(bic)))
            for fam in ordered:
                lines.append("  {}".format(fam))
                for typ in ordered[fam]:
                    lines.append("      {}".format(typ))
            lines.append("")

        forms.alert_ifnot(lines, "No MEP family types found (with the selected scope).", exitscript=True)
        _place_text_list(target_view, pt, lines, text_type_id=chosen_text_type_id)
        return

    if mode == "Legend from Color Fill Scheme":
        # Smart behavior: allow scheme legends either in current view or a new Legend view.
        if target_mode == "New Drafting View":
            target_mode = "Current View"
        scheme = form.values.get("scheme")
        if scheme is None:
            forms.alert("Pick a Color Fill Scheme.", exitscript=True)

        base_name = "GM - Scheme Legend - {}".format(getattr(scheme, "Name", "Scheme"))
        target_view = _get_target_view(target_mode, base_name)
        if target_view.Id != orig_view.Id:
            _activate_view(target_view)
        pt = _pick_point_in_active_view("Pick Placement Point", restore_view=orig_view)

        try:
            box_w = float(form.values.get("box_w") or 1000)
            box_h = float(form.values.get("box_h") or 240)
            box_off = float(form.values.get("box_off") or 80)
        except Exception:
            box_w, box_h, box_off = 1000, 240, 80

        # "Only used in views" is interpreted here as "only visible entries" (best-effort).
        only_visible = True if only_used else False

        _draw_color_scheme_legend(
            target_view,
            pt,
            scheme,
            only_visible=only_visible,
            box_w_mm=box_w,
            box_h_mm=box_h,
            box_off_mm=box_off,
        )
        return

    if mode == "Legend from Spaces":
        base_name = "GM - Spaces Legend"
        target_view = _get_target_view(target_mode, base_name)
        if target_view.Id != orig_view.Id:
            _activate_view(target_view)
        pt = _pick_point_in_active_view("Pick Placement Point", restore_view=orig_view)

        spaces = _collect_spaces(only_used=only_used, used_scope_view=used_scope_view)
        forms.alert_ifnot(spaces, "No Spaces found (with the selected scope).", exitscript=True)

        # Sort by label for stable ordering
        labels = sorted([_space_label(s) for s in spaces if s is not None])
        _place_text_list(target_view, pt, labels, text_type_id=chosen_text_type_id)
        return


if __name__ == "__main__":
    main()
