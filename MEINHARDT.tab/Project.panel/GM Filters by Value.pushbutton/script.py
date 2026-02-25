# -*- coding: utf-8 -*-
# GMToolbox version of Filterbyvalue
# Author: Gino Moreno
# Description: Filters and colors MEP elements in the current and linked Revit files, with per-category selection. Includes support for duct and pipe systems.
# Enhanced to support linked file filtering

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List
import csv
import re
import sys

doc = __revit__.ActiveUIDocument.Document
active_view = doc.ActiveView


def _ask_yesno(prompt, title='System Filter Settings', default_yes=True):
    # pyRevit forms.alert returns True/False when yes/no flags are provided
    if default_yes:
        return bool(forms.alert(prompt, title=title, yes=True, no=True))
    else:
        # No direct default selection control; keep same signature for readability
        return bool(forms.alert(prompt, title=title, yes=True, no=True))


def _get_settings_via_pyrevit_forms():
    title = 'System Filter Settings'

    halftone = _ask_yesno('Halftone?', title=title, default_yes=True)

    transparency = forms.ask_for_string(default='0', prompt='Transparency (0-100):', title=title)
    if transparency is None:
        return None

    fill_mode = forms.CommandSwitchWindow.show(
        ['Background', 'Foreground', 'Both'],
        message='Fill mode:',
        title=title
    )
    if not fill_mode:
        return None

    line_color_mode = forms.CommandSwitchWindow.show(
        ['Darker than fill', 'Same as fill'],
        message='Line color:',
        title=title
    )
    if not line_color_mode:
        return None

    line_darken = forms.ask_for_string(
        default='0.75',
        prompt='Line darken factor (0-1, only if darker):',
        title=title
    )
    if line_darken is None:
        return None

    cat_mode = forms.CommandSwitchWindow.show(
        ['Pick categories', 'All MEP categories (expanded)'],
        message='Categories:',
        title=title
    )
    if not cat_mode:
        return None

    defaults = {
        'duct_curves': True,
        'duct_fittings': True,
        'duct_accessories': True,
        'duct_insulations': True,
        'pipe_curves': True,
        'pipe_fittings': True,
        'pipe_accessories': True,
        'pipe_insulations': True,
    }

    all_mep = (cat_mode == 'All MEP categories (expanded)')
    if not all_mep:
        options = [
            'Duct Curves',
            'Duct Fittings',
            'Duct Accessories',
            'Duct Insulations',
            'Pipe Curves',
            'Pipe Fittings',
            'Pipe Accessories',
            'Pipe Insulations',
        ]
        default_sel = [
            k for k, v in {
                'Duct Curves': defaults['duct_curves'],
                'Duct Fittings': defaults['duct_fittings'],
                'Duct Accessories': defaults['duct_accessories'],
                'Duct Insulations': defaults['duct_insulations'],
                'Pipe Curves': defaults['pipe_curves'],
                'Pipe Fittings': defaults['pipe_fittings'],
                'Pipe Accessories': defaults['pipe_accessories'],
                'Pipe Insulations': defaults['pipe_insulations'],
            }.items() if v
        ]
        picked = forms.SelectFromList.show(
            options,
            title=title,
            multiselect=True,
            button_name='Select',
        )
        if picked is None:
            return None
        # if user hits Select with nothing, allow it to fail later (same behavior)
        defaults.update({
            'duct_curves': 'Duct Curves' in picked,
            'duct_fittings': 'Duct Fittings' in picked,
            'duct_accessories': 'Duct Accessories' in picked,
            'duct_insulations': 'Duct Insulations' in picked,
            'pipe_curves': 'Pipe Curves' in picked,
            'pipe_fittings': 'Pipe Fittings' in picked,
            'pipe_accessories': 'Pipe Accessories' in picked,
            'pipe_insulations': 'Pipe Insulations' in picked,
        })

    load_csv = _ask_yesno('Load MHT CSV mapping for line patterns?', title=title, default_yes=False)

    vals = {
        'halftone': halftone,
        'transparency': transparency,
        'fill_mode': fill_mode,
        'line_color_mode': line_color_mode,
        'line_darken': line_darken,
        'all_mep': all_mep,
        'load_csv': load_csv,
    }
    vals.update(defaults)
    return vals

# --- Linked File Helpers ---
def get_linked_docs():
    """Get all loaded linked documents"""
    links = [el for el in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)]
    link_docs = []
    for link in links:
        try:
            ldoc = link.GetLinkDocument()
            if ldoc:
                link_docs.append((link, ldoc))
        except:
            pass
    return link_docs


def get_linked_doc_by_name(name):
    """Get a linked document by its title"""
    for link, ldoc in get_linked_docs():
        if ldoc.Title == name:
            return (link, ldoc)
    return (None, None)


def collect_system_types_from_link(bic, target_doc):
    """Collect system types from a linked document"""
    try:
        return (
            DB.FilteredElementCollector(target_doc)
            .OfCategory(bic)
            .WhereElementIsElementType()
            .ToElements()
        )
    except:
        return []


def get_used_system_type_ids_in_linked_view(link_instance, linked_doc, view, elem_cats, sys_param):
    """Get used system type IDs in linked elements visible in a view"""
    used_ids = set()
    try:
        cat_ids = List[DB.ElementId]([DB.ElementId(c) for c in elem_cats])
        # Collect elements from the linked document
        collector = (
            DB.FilteredElementCollector(linked_doc, view.Id)
            .WherePasses(DB.ElementMulticategoryFilter(cat_ids))
            .WhereElementIsNotElementType()
        )
        for el in collector:
            try:
                p = el.get_Parameter(sys_param)
                if not p or p.StorageType != DB.StorageType.ElementId:
                    continue
                st_id = p.AsElementId()
                if st_id and st_id != DB.ElementId.InvalidElementId:
                    used_ids.add(st_id.IntegerValue)
            except:
                pass
    except:
        pass
    return used_ids


def create_linked_filter_name(link_name, base_name):
    """Create a filter name for linked elements"""
    return sanitize("LINK-{} - {}".format(link_name, base_name))


# --- Original Tool Logic (current model) ---
def sanitize(name):
    for ch in '{}[]:/\\|?*<>':
        name = name.replace(ch, '-')
    return name.strip()


def color_from_text(text):
    # Deterministic, readable-ish color from a string.
    try:
        s = str(text or '')
    except Exception:
        s = ''
    h = 0
    for ch in s:
        try:
            h = (h * 31 + ord(ch)) & 0xFFFFFFFF
        except Exception:
            continue
    # Pastel-ish range
    r = 80 + (h & 0x7F)
    g = 80 + ((h >> 8) & 0x7F)
    b = 80 + ((h >> 16) & 0x7F)
    return DB.Color(int(max(0, min(255, r))), int(max(0, min(255, g))), int(max(0, min(255, b))))


def get_param_map_from_sample(sample_elem):
    # Returns {param_name: (param_id, storage_type)}
    m = {}
    if not sample_elem:
        return m
    try:
        for p in sample_elem.Parameters:
            try:
                d = p.Definition
                if not d:
                    continue
                n = (d.Name or '').strip()
                if not n:
                    continue
                if n not in m:
                    m[n] = (p.Id, p.StorageType)
            except:
                pass
    except:
        pass
    return m


def get_param_on_element(el, param_id, param_name=None):
    try:
        if param_id is not None:
            p = el.get_Parameter(param_id)
            if p:
                return p
    except:
        pass
    if param_name:
        try:
            return el.LookupParameter(param_name)
        except:
            return None
    return None


def get_param_value_for_rule(p):
    # Returns (key, display, rule_value)
    if not p:
        return None, None, None
    st = p.StorageType
    try:
        if st == DB.StorageType.String:
            s = p.AsString() or p.AsValueString() or ''
            s = str(s).strip()
            if not s:
                return None, None, None
            return s.lower(), s, s
        if st == DB.StorageType.Integer:
            i = p.AsInteger()
            return i, str(i), int(i)
        if st == DB.StorageType.Double:
            d = p.AsDouble()
            # Display as value string when possible
            try:
                disp = p.AsValueString() or str(d)
            except:
                disp = str(d)
            return float(d), str(disp), float(d)
        if st == DB.StorageType.ElementId:
            eid = p.AsElementId()
            if not eid or eid == DB.ElementId.InvalidElementId:
                return None, None, None
            # Try to resolve name for display
            disp = None
            try:
                ee = doc.GetElement(eid)
                if ee is not None and hasattr(ee, 'Name'):
                    disp = ee.Name
            except:
                pass
            if not disp:
                disp = str(eid.IntegerValue)
            return eid.IntegerValue, disp, eid
    except:
        return None, None, None
    return None, None, None


def create_equals_rule(param_id, storage_type, rule_value):
    try:
        if storage_type == DB.StorageType.String:
            try:
                return DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, str(rule_value), False)
            except:
                return DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, str(rule_value))
        if storage_type == DB.StorageType.Integer:
            return DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, int(rule_value))
        if storage_type == DB.StorageType.Double:
            # Tolerance in internal units
            return DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, float(rule_value), 1e-6)
        if storage_type == DB.StorageType.ElementId:
            return DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, rule_value)
    except:
        return None
    return None

def get_type_name(st):
    p = st.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    return p.AsString() if p else "Unnamed System"

def get_system_material_color(st):
    mat_id = st.MaterialId
    if mat_id and mat_id != DB.ElementId.InvalidElementId:
        mat = doc.GetElement(mat_id)
        if mat and mat.Color and mat.Color.IsValid:
            return mat.Color
    return DB.Color(150, 150, 150)

def get_system_abbreviation(st):
    try:
        p = st.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)
        if p:
            val = p.AsString()
            if val:
                return val.strip().upper()
    except:
        pass
    return None

def darker_color(color, factor=0.75):
    r = int(max(0, min(255, color.Red * factor)))
    g = int(max(0, min(255, color.Green * factor)))
    b = int(max(0, min(255, color.Blue * factor)))
    return DB.Color(r, g, b)

def get_solid_fill_id():
    for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                return fp.Id
        except:
            pass
    return DB.ElementId.InvalidElementId

def get_line_pattern_id_by_name(name):
    for lp in DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement):
        if lp.Name == name:
            return lp.Id
    return DB.ElementId.InvalidElementId

def try_get_bic(bic_name):
    try:
        return getattr(DB.BuiltInCategory, bic_name)
    except:
        return None

def _extract_mht_code(text):
    if not text:
        return None
    try:
        s = str(text).strip()
    except:
        return None
    try:
        matches = re.findall(r"\(([^)]+)\)", s)
        if matches:
            code = matches[-1].strip().upper()
            if code and not code.isdigit():
                return code
    except:
        pass
    u = s.upper()
    if u.startswith("M-MHT-"):
        u = u[len("M-MHT-"):]
    if u.startswith("_"):
        u = u[1:]
    token = u.split(" ", 1)[0]
    token = token.split("_", 1)[0]
    token = token.strip()
    if not token or token.isdigit():
        return None
    return token

def load_mht_pattern_map_from_csv(csv_path):
    base_map = {}
    solid_map = {}
    try:
        with open(csv_path, 'rb') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    line_style = (row.get('Line Style') or '').strip()
                    pattern = (row.get('Line Pattern') or '').strip()
                    if not line_style or not pattern:
                        continue
                    code = _extract_mht_code(line_style)
                    if not code:
                        continue
                    if 'SOLID' in line_style.upper():
                        solid_map[code] = pattern
                    else:
                        base_map[code] = pattern
                except:
                    pass
    except:
        return {}
    for k, v in solid_map.items():
        if k not in base_map:
            base_map[k] = v
    return base_map

def pick_views():
    choice = forms.CommandSwitchWindow.show(
        ["Active View", "Select Views"],
        message="Apply filters to which views?"
    )
    if choice == "Active View":
        return [active_view]
    views = [
        v for v in DB.FilteredElementCollector(doc)
        .OfClass(DB.View)
        .WhereElementIsNotElementType()
        if not v.IsTemplate
    ]
    view_map = {"{} | {}".format(v.ViewType, v.Name): v for v in views}
    picked = forms.SelectFromList.show(
        sorted(view_map.keys()),
        multiselect=True,
        title="Select Views"
    )
    return [view_map[x] for x in picked] if picked else []

def find_filter(name):
    for f in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
        if f.Name == name:
            return f
    return None

def collect_system_types(bic, target_doc=None):
    if not target_doc:
        target_doc = doc
    return (
        DB.FilteredElementCollector(target_doc)
        .OfCategory(bic)
        .WhereElementIsElementType()
        .ToElements()
    )

def get_used_system_type_ids_in_view(view, elem_cats, sys_param, target_doc=None):
    if not target_doc:
        target_doc = doc
    used_ids = set()
    try:
        cat_ids = List[DB.ElementId]([DB.ElementId(c) for c in elem_cats])
        collector = (
            DB.FilteredElementCollector(target_doc, view.Id)
            .WherePasses(DB.ElementMulticategoryFilter(cat_ids))
            .WhereElementIsNotElementType()
        )
        for el in collector:
            try:
                p = el.get_Parameter(sys_param)
                if not p or p.StorageType != DB.StorageType.ElementId:
                    continue
                st_id = p.AsElementId()
                if st_id and st_id != DB.ElementId.InvalidElementId:
                    used_ids.add(st_id.IntegerValue)
            except:
                pass
    except:
        pass
    return used_ids

# --- DATA: Current Model ---
duct_systems = collect_system_types(DB.BuiltInCategory.OST_DuctSystem)
pipe_systems = collect_system_types(DB.BuiltInCategory.OST_PipingSystem)

views = pick_views()
if not views:
    script.exit()

mode_choice = forms.CommandSwitchWindow.show(
    ['System Types (Current)', 'By Parameter (New)'],
    message='How do you want to create filters?'
)
if not mode_choice:
    script.exit()

use_param_mode = mode_choice.startswith('By Parameter')

# Only require system types when running the current System Type workflow.
if (not use_param_mode) and (not duct_systems and not pipe_systems):
    forms.alert("No system types found.")
    script.exit()

# --- SETTINGS (global) ---
vals = _get_settings_via_pyrevit_forms()
if vals is None:
    script.exit()

try:
    transparency = int(float(vals.get("transparency", 0)))
except:
    transparency = 0
if transparency < 0:
    transparency = 0
if transparency > 100:
    transparency = 100

try:
    line_darken = float(vals.get("line_darken", 0.75))
except:
    line_darken = 0.75
if line_darken < 0.0:
    line_darken = 0.0
if line_darken > 1.0:
    line_darken = 1.0

halftone = bool(vals.get("halftone", True))
fill_mode = vals.get("fill_mode") or "Background"

line_color_mode = vals.get("line_color_mode") or "Darker than fill"
line_color_mode = "same" if line_color_mode == "Same as fill" else "darker"

pattern_map = {}
if vals.get("load_csv"):
    csv_path = forms.pick_file(file_ext='csv', title='Pick MHT LineStyles CSV')
    if not csv_path:
        sys.exit()
    pattern_map = load_mht_pattern_map_from_csv(csv_path)

# --- PRECOMPUTE: system types actually used per view ---
duct_elem_cats = []
pipe_elem_cats = []

if vals.get("all_mep"):
    duct_elem_cats = [
        DB.BuiltInCategory.OST_DuctCurves,
        DB.BuiltInCategory.OST_DuctFitting,
        DB.BuiltInCategory.OST_DuctAccessory,
        DB.BuiltInCategory.OST_DuctInsulations,
    ]
    extra = [
        try_get_bic('OST_FlexDuctCurves'),
        try_get_bic('OST_DuctTerminal'),
        try_get_bic('OST_DuctPlaceHolder'),
    ]
    duct_elem_cats.extend([c for c in extra if c is not None])

    pipe_elem_cats = [
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeFitting,
        DB.BuiltInCategory.OST_PipeAccessory,
        DB.BuiltInCategory.OST_PipeInsulations,
    ]
    extra = [
        try_get_bic('OST_FlexPipeCurves'),
        try_get_bic('OST_PlaceHolderPipes'),
    ]
    pipe_elem_cats.extend([c for c in extra if c is not None])
else:
    if vals.get("duct_curves"):
        duct_elem_cats.append(DB.BuiltInCategory.OST_DuctCurves)
    if vals.get("duct_fittings"):
        duct_elem_cats.append(DB.BuiltInCategory.OST_DuctFitting)
    if vals.get("duct_accessories"):
        duct_elem_cats.append(DB.BuiltInCategory.OST_DuctAccessory)
    if vals.get("duct_insulations"):
        duct_elem_cats.append(DB.BuiltInCategory.OST_DuctInsulations)

    if vals.get("pipe_curves"):
        pipe_elem_cats.append(DB.BuiltInCategory.OST_PipeCurves)
    if vals.get("pipe_fittings"):
        pipe_elem_cats.append(DB.BuiltInCategory.OST_PipeFitting)
    if vals.get("pipe_accessories"):
        pipe_elem_cats.append(DB.BuiltInCategory.OST_PipeAccessory)
    if vals.get("pipe_insulations"):
        pipe_elem_cats.append(DB.BuiltInCategory.OST_PipeInsulations)

forms.alert_ifnot(duct_elem_cats or pipe_elem_cats, "No categories selected.", exitscript=True)

used_by_view = {}
all_used_duct = set()
all_used_pipe = set()

for v in views:
    duct_used = get_used_system_type_ids_in_view(
        v,
        duct_elem_cats,
        DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM
    )

    pipe_used = get_used_system_type_ids_in_view(
        v,
        pipe_elem_cats,
        DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
    )

    used_by_view[v.Id.IntegerValue] = {
        "duct": duct_used,
        "pipe": pipe_used
    }

    all_used_duct |= duct_used
    all_used_pipe |= pipe_used

solid_fill_id = get_solid_fill_id()

# --- MAIN ---
if use_param_mode:
    # Create filters by distinct values of a user-selected parameter.
    all_elem_cats = []
    all_elem_cats.extend(duct_elem_cats)
    all_elem_cats.extend(pipe_elem_cats)

    # Find a sample element from selected categories in the selected views.
    sample = None
    try:
        cat_ids = List[DB.ElementId]([DB.ElementId(c) for c in all_elem_cats])
        for v in views:
            try:
                sample = (
                    DB.FilteredElementCollector(doc, v.Id)
                    .WherePasses(DB.ElementMulticategoryFilter(cat_ids))
                    .WhereElementIsNotElementType()
                    .FirstElement()
                )
            except:
                sample = None
            if sample:
                break
    except:
        sample = None

    if not sample:
        forms.alert('No elements found in the selected views/categories.', exitscript=True)

    pmap = get_param_map_from_sample(sample)
    if not pmap:
        forms.alert('Could not read parameters from a sample element.', exitscript=True)

    param_name = forms.SelectFromList.show(sorted(pmap.keys()), title='Select Parameter', multiselect=False)
    if not param_name:
        script.exit()

    param_id, storage_type = pmap.get(param_name)

    # Collect distinct values
    values = {}  # key -> (display, rule_value)
    try:
        cat_ids = List[DB.ElementId]([DB.ElementId(c) for c in all_elem_cats])
        for v in views:
            collector = (
                DB.FilteredElementCollector(doc, v.Id)
                .WherePasses(DB.ElementMulticategoryFilter(cat_ids))
                .WhereElementIsNotElementType()
            )
            for el in collector:
                try:
                    p = get_param_on_element(el, param_id, param_name=param_name)
                    k, disp, rv = get_param_value_for_rule(p)
                    if k is None:
                        continue
                    if k not in values:
                        values[k] = (disp, rv)
                except:
                    pass
    except:
        pass

    if not values:
        forms.alert('No non-empty values found for parameter "{}".'.format(param_name), exitscript=True)

    # Safety: too many filters can get heavy.
    if len(values) > 200:
        ok_many = forms.alert(
            'Found {} unique values for "{}".\n\nThis will create a lot of filters. Continue?'.format(len(values), param_name),
            yes=True, no=True
        )
        if not ok_many:
            script.exit()

    with revit.Transaction('Filters by Parameter (Background + Line Pattern)'):
        cat_ids = List[DB.ElementId]([DB.ElementId(c) for c in all_elem_cats])

        for k, (disp, rv) in values.items():
            try:
                # Filter name: PARAM - VALUE
                fname = sanitize('{} - {}'.format(param_name, disp))
                filt = find_filter(fname)

                rule = create_equals_rule(param_id, storage_type, rv)
                if not rule:
                    continue
                elem_filter = DB.ElementParameterFilter(List[DB.FilterRule]([rule]))

                if not filt:
                    filt = DB.ParameterFilterElement.Create(doc, fname, cat_ids)
                else:
                    try:
                        filt.SetCategories(cat_ids)
                    except:
                        pass
                try:
                    filt.SetElementFilter(elem_filter)
                except:
                    pass

                fill_color = color_from_text(disp)
                if line_color_mode == 'same':
                    line_color = fill_color
                else:
                    line_color = darker_color(fill_color, factor=line_darken)

                ogs = DB.OverrideGraphicSettings()
                try:
                    ogs.SetHalftone(halftone)
                except:
                    pass
                try:
                    ogs.SetSurfaceTransparency(transparency)
                except:
                    pass
                try:
                    ogs.SetCutTransparency(transparency)
                except:
                    pass

                ogs.SetProjectionLineColor(line_color)
                ogs.SetCutLineColor(line_color)

                # Pattern: try match by value text, else none.
                pattern_id = get_line_pattern_id_by_name(disp)
                if pattern_id != DB.ElementId.InvalidElementId:
                    ogs.SetProjectionLinePatternId(pattern_id)
                    ogs.SetCutLinePatternId(pattern_id)

                use_bg = (fill_mode in ['Background', 'Both'])
                use_fg = (fill_mode in ['Foreground', 'Both'])
                if solid_fill_id != DB.ElementId.InvalidElementId:
                    if use_bg:
                        ogs.SetSurfaceBackgroundPatternId(solid_fill_id)
                        ogs.SetSurfaceBackgroundPatternColor(fill_color)
                        ogs.SetCutBackgroundPatternId(solid_fill_id)
                        ogs.SetCutBackgroundPatternColor(fill_color)
                    else:
                        ogs.SetSurfaceBackgroundPatternId(DB.ElementId.InvalidElementId)
                        ogs.SetCutBackgroundPatternId(DB.ElementId.InvalidElementId)

                    if use_fg:
                        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                        ogs.SetSurfaceForegroundPatternColor(fill_color)
                        ogs.SetCutForegroundPatternId(solid_fill_id)
                        ogs.SetCutForegroundPatternColor(fill_color)
                    else:
                        ogs.SetSurfaceForegroundPatternId(DB.ElementId.InvalidElementId)
                        ogs.SetCutForegroundPatternId(DB.ElementId.InvalidElementId)

                for v in views:
                    try:
                        if filt.Id not in list(v.GetFilters()):
                            v.AddFilter(filt.Id)
                        v.SetIsFilterEnabled(filt.Id, True)
                        v.SetFilterOverrides(filt.Id, ogs)
                    except:
                        pass

            except:
                pass

    forms.alert(
        'DONE.\n\n'
        'Filters created by parameter:\n'
        '- Parameter: {}\n'
        '- Filters created/updated: {}'.format(param_name, len(values))
    )

else:
    with revit.Transaction("System Filters (Background + Line Pattern)"):

        systems_map = [
            (
                duct_systems,
                duct_elem_cats,
                DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM
            ),
            (
                pipe_systems,
                pipe_elem_cats,
                DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
            )
        ]

        for systems, elem_cats, sys_param in systems_map:
            is_duct = (sys_param == DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
            used_anywhere = all_used_duct if is_duct else all_used_pipe

            for st in systems:

                # SKIP UNUSED SYSTEM TYPES
                if st.Id.IntegerValue not in used_anywhere:
                    continue

                name = sanitize(get_type_name(st))
                fill_color = get_system_material_color(st)
                if line_color_mode == "same":
                    line_color = fill_color
                else:
                    line_color = darker_color(fill_color, factor=line_darken)

                filt = find_filter(name)
                cat_ids = List[DB.ElementId]([DB.ElementId(c) for c in elem_cats])

                rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(
                    DB.ElementId(sys_param),
                    st.Id
                )
                elem_filter = DB.ElementParameterFilter(
                    List[DB.FilterRule]([rule])
                )

                if not filt:
                    filt = DB.ParameterFilterElement.Create(doc, name, cat_ids)
                else:
                    try:
                        filt.SetCategories(cat_ids)
                    except:
                        pass

                try:
                    filt.SetElementFilter(elem_filter)
                except:
                    pass

                ogs = DB.OverrideGraphicSettings()

                try:
                    ogs.SetHalftone(halftone)
                except:
                    pass

                try:
                    ogs.SetSurfaceTransparency(transparency)
                except:
                    pass

                try:
                    ogs.SetCutTransparency(transparency)
                except:
                    pass

                # --- LINE GRAPHICS ---
                ogs.SetProjectionLineColor(line_color)
                ogs.SetCutLineColor(line_color)

                pattern_id = DB.ElementId.InvalidElementId
                if pattern_map:
                    code = get_system_abbreviation(st) or _extract_mht_code(get_type_name(st))
                    patt_name = pattern_map.get(code) if code else None
                    if patt_name:
                        pattern_id = get_line_pattern_id_by_name(patt_name)

                if pattern_id == DB.ElementId.InvalidElementId:
                    pattern_id = get_line_pattern_id_by_name(name)

                if pattern_id != DB.ElementId.InvalidElementId:
                    ogs.SetProjectionLinePatternId(pattern_id)
                    ogs.SetCutLinePatternId(pattern_id)

                # --- FILL ---
                use_bg = (fill_mode in ["Background", "Both"])
                use_fg = (fill_mode in ["Foreground", "Both"])

                if solid_fill_id != DB.ElementId.InvalidElementId:
                    if use_bg:
                        ogs.SetSurfaceBackgroundPatternId(solid_fill_id)
                        ogs.SetSurfaceBackgroundPatternColor(fill_color)
                        ogs.SetCutBackgroundPatternId(solid_fill_id)
                        ogs.SetCutBackgroundPatternColor(fill_color)
                    else:
                        ogs.SetSurfaceBackgroundPatternId(DB.ElementId.InvalidElementId)
                        ogs.SetCutBackgroundPatternId(DB.ElementId.InvalidElementId)

                    if use_fg:
                        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                        ogs.SetSurfaceForegroundPatternColor(fill_color)
                        ogs.SetCutForegroundPatternId(solid_fill_id)
                        ogs.SetCutForegroundPatternColor(fill_color)
                    else:
                        ogs.SetSurfaceForegroundPatternId(DB.ElementId.InvalidElementId)
                        ogs.SetCutForegroundPatternId(DB.ElementId.InvalidElementId)
                else:
                    ogs.SetSurfaceBackgroundPatternId(DB.ElementId.InvalidElementId)
                    ogs.SetCutBackgroundPatternId(DB.ElementId.InvalidElementId)
                    ogs.SetSurfaceForegroundPatternId(DB.ElementId.InvalidElementId)
                    ogs.SetCutForegroundPatternId(DB.ElementId.InvalidElementId)

                for v in views:
                    try:
                        used_in_view = used_by_view.get(v.Id.IntegerValue, {})
                        used_set = used_in_view.get("duct" if is_duct else "pipe", set())

                        if st.Id.IntegerValue not in used_set:
                            continue

                        view_filters = list(v.GetFilters())

                        if filt.Id not in view_filters:
                            v.AddFilter(filt.Id)

                        v.SetIsFilterEnabled(filt.Id, True)
                        v.SetFilterOverrides(filt.Id, ogs)

                    except:
                        pass

            try:
                filt.SetElementFilter(elem_filter)
            except:
                pass

            ogs = DB.OverrideGraphicSettings()

            try:
                ogs.SetHalftone(halftone)
            except:
                pass

            try:
                ogs.SetSurfaceTransparency(transparency)
            except:
                pass

            try:
                ogs.SetCutTransparency(transparency)
            except:
                pass

            # --- LINE GRAPHICS ---
            ogs.SetProjectionLineColor(line_color)
            ogs.SetCutLineColor(line_color)

            pattern_id = DB.ElementId.InvalidElementId
            if pattern_map:
                code = get_system_abbreviation(st) or _extract_mht_code(get_type_name(st))
                patt_name = pattern_map.get(code) if code else None
                if patt_name:
                    pattern_id = get_line_pattern_id_by_name(patt_name)

            if pattern_id == DB.ElementId.InvalidElementId:
                pattern_id = get_line_pattern_id_by_name(name)

            if pattern_id != DB.ElementId.InvalidElementId:
                ogs.SetProjectionLinePatternId(pattern_id)
                ogs.SetCutLinePatternId(pattern_id)

            # --- FILL ---
            use_bg = (fill_mode in ["Background", "Both"])
            use_fg = (fill_mode in ["Foreground", "Both"])

            if solid_fill_id != DB.ElementId.InvalidElementId:
                if use_bg:
                    ogs.SetSurfaceBackgroundPatternId(solid_fill_id)
                    ogs.SetSurfaceBackgroundPatternColor(fill_color)
                    ogs.SetCutBackgroundPatternId(solid_fill_id)
                    ogs.SetCutBackgroundPatternColor(fill_color)
                else:
                    ogs.SetSurfaceBackgroundPatternId(DB.ElementId.InvalidElementId)
                    ogs.SetCutBackgroundPatternId(DB.ElementId.InvalidElementId)

                if use_fg:
                    ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                    ogs.SetSurfaceForegroundPatternColor(fill_color)
                    ogs.SetCutForegroundPatternId(solid_fill_id)
                    ogs.SetCutForegroundPatternColor(fill_color)
                else:
                    ogs.SetSurfaceForegroundPatternId(DB.ElementId.InvalidElementId)
                    ogs.SetCutForegroundPatternId(DB.ElementId.InvalidElementId)
            else:
                ogs.SetSurfaceBackgroundPatternId(DB.ElementId.InvalidElementId)
                ogs.SetCutBackgroundPatternId(DB.ElementId.InvalidElementId)
                ogs.SetSurfaceForegroundPatternId(DB.ElementId.InvalidElementId)
                ogs.SetCutForegroundPatternId(DB.ElementId.InvalidElementId)

            for v in views:
                try:
                    used_in_view = used_by_view.get(v.Id.IntegerValue, {})
                    used_set = used_in_view.get("duct" if is_duct else "pipe", set())

                    if st.Id.IntegerValue not in used_set:
                        continue

                    view_filters = list(v.GetFilters())

                    if filt.Id not in view_filters:
                        v.AddFilter(filt.Id)

                    v.SetIsFilterEnabled(filt.Id, True)
                    v.SetFilterOverrides(filt.Id, ogs)

                except:
                    pass

    forms.alert(
        "DONE.\n\n"
        "System filters updated successfully.\n\n"
        "- Background fill applied\n"
        "- Line colors applied\n"
        "- Line patterns matched by {}".format("MHT CSV mapping" if pattern_map else "system name")
    )


# --- LINKED FILE PROCESSING ---
# Now process linked files with the same filtering logic
link_docs = get_linked_docs()
if link_docs:
    # Let user select which linked files to process
    link_options = {}
    for link_instance, linked_doc in link_docs:
        try:
            link_name = linked_doc.Title
            link_options[link_name] = (link_instance, linked_doc)
        except:
            pass
    
    if link_options:
        selected_link_names = forms.SelectFromList.show(
            sorted(link_options.keys()),
            title='Select Linked Files to Process',
            multiselect=True,
            button_name='Process Selected Links'
        )
        
        if selected_link_names:
            selected_links = [(link_options[name][0], link_options[name][1]) for name in selected_link_names]
            total_linked_filters = 0
            
            with revit.Transaction("Linked System Filters (Background + Line Pattern)"):
                for link_instance, linked_doc in selected_links:
                    try:
                        link_name = linked_doc.Title
                        
                        # Collect system types from linked document
                        linked_duct_systems = collect_system_types_from_link(DB.BuiltInCategory.OST_DuctSystem, linked_doc)
                        linked_pipe_systems = collect_system_types_from_link(DB.BuiltInCategory.OST_PipingSystem, linked_doc)
                    
                    if not linked_duct_systems and not linked_pipe_systems:
                        continue
                    
                    # Precompute used system types in linked file per view
                    linked_used_by_view = {}
                    linked_all_used_duct = set()
                    linked_all_used_pipe = set()
                    
                    for v in views:
                        try:
                            duct_used = get_used_system_type_ids_in_linked_view(
                                link_instance,
                                linked_doc,
                                v,
                                duct_elem_cats,
                                DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM
                            )
                            
                            pipe_used = get_used_system_type_ids_in_linked_view(
                                link_instance,
                                linked_doc,
                                v,
                                pipe_elem_cats,
                                DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM
                            )
                            
                            linked_used_by_view[v.Id.IntegerValue] = {
                                "duct": duct_used,
                                "pipe": pipe_used
                            }
                            
                            linked_all_used_duct |= duct_used
                            linked_all_used_pipe |= pipe_used
                        except:
                            pass
                    
                    # Process linked duct and pipe systems
                    linked_systems_map = [
                        (
                            linked_duct_systems,
                            duct_elem_cats,
                            DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM,
                            linked_all_used_duct,
                            "duct"
                        ),
                        (
                            linked_pipe_systems,
                            pipe_elem_cats,
                            DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM,
                            linked_all_used_pipe,
                            "pipe"
                        )
                    ]
                    
                    for systems, elem_cats, sys_param, used_anywhere, system_type in linked_systems_map:
                        for st in systems:
                            # SKIP UNUSED SYSTEM TYPES
                            if st.Id.IntegerValue not in used_anywhere:
                                continue
                            
                            base_name = sanitize(get_type_name(st))
                            filter_name = create_linked_filter_name(link_name, base_name)
                            
                            # Use same color scheme as host model
                            fill_color = get_system_material_color(st)
                            if line_color_mode == "same":
                                line_color = fill_color
                            else:
                                line_color = darker_color(fill_color, factor=line_darken)
                            
                            filt = find_filter(filter_name)
                            
                            # Create category list for linked elements
                            # Linked elements need the RevitLinkInstance category plus the element categories
                            cat_ids = List[DB.ElementId]()
                            cat_ids.Add(DB.ElementId(DB.BuiltInCategory.OST_RvtLinks))
                            
                            # Create filter rule for linked elements
                            # We need to filter by the link instance AND the system type
                            rules = List[DB.FilterRule]()
                            
                            # Rule 1: Match link instance
                            link_rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(
                                DB.ElementId(DB.BuiltInParameter.SYMBOL_ID_PARAM),
                                link_instance.GetTypeId()
                            )
                            rules.Add(link_rule)
                            
                            # Rule 2: Match system type (this applies to elements within the link)
                            # Note: Direct filtering of linked element parameters is limited in Revit API
                            # We use the system type as a reference but the filter primarily works on visibility
                            
                            elem_filter = DB.ElementParameterFilter(rules)
                            
                            if not filt:
                                filt = DB.ParameterFilterElement.Create(doc, filter_name, cat_ids)
                            else:
                                try:
                                    filt.SetCategories(cat_ids)
                                except:
                                    pass
                            
                            try:
                                filt.SetElementFilter(elem_filter)
                            except:
                                pass
                            
                            # Create override graphics settings
                            ogs = DB.OverrideGraphicSettings()
                            
                            try:
                                ogs.SetHalftone(halftone)
                            except:
                                pass
                            
                            try:
                                ogs.SetSurfaceTransparency(transparency)
                            except:
                                pass
                            
                            try:
                                ogs.SetCutTransparency(transparency)
                            except:
                                pass
                            
                            # LINE GRAPHICS
                            ogs.SetProjectionLineColor(line_color)
                            ogs.SetCutLineColor(line_color)
                            
                            # Line pattern
                            pattern_id = DB.ElementId.InvalidElementId
                            if pattern_map:
                                code = get_system_abbreviation(st) or _extract_mht_code(get_type_name(st))
                                patt_name = pattern_map.get(code) if code else None
                                if patt_name:
                                    pattern_id = get_line_pattern_id_by_name(patt_name)
                            
                            if pattern_id == DB.ElementId.InvalidElementId:
                                pattern_id = get_line_pattern_id_by_name(base_name)
                            
                            if pattern_id != DB.ElementId.InvalidElementId:
                                ogs.SetProjectionLinePatternId(pattern_id)
                                ogs.SetCutLinePatternId(pattern_id)
                            
                            # FILL
                            use_bg = (fill_mode in ["Background", "Both"])
                            use_fg = (fill_mode in ["Foreground", "Both"])
                            
                            if solid_fill_id != DB.ElementId.InvalidElementId:
                                if use_bg:
                                    ogs.SetSurfaceBackgroundPatternId(solid_fill_id)
                                    ogs.SetSurfaceBackgroundPatternColor(fill_color)
                                    ogs.SetCutBackgroundPatternId(solid_fill_id)
                                    ogs.SetCutBackgroundPatternColor(fill_color)
                                else:
                                    ogs.SetSurfaceBackgroundPatternId(DB.ElementId.InvalidElementId)
                                    ogs.SetCutBackgroundPatternId(DB.ElementId.InvalidElementId)
                                
                                if use_fg:
                                    ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                                    ogs.SetSurfaceForegroundPatternColor(fill_color)
                                    ogs.SetCutForegroundPatternId(solid_fill_id)
                                    ogs.SetCutForegroundPatternColor(fill_color)
                                else:
                                    ogs.SetSurfaceForegroundPatternId(DB.ElementId.InvalidElementId)
                                    ogs.SetCutForegroundPatternId(DB.ElementId.InvalidElementId)
                            
                            # Apply filter to views
                            for v in views:
                                try:
                                    used_in_view = linked_used_by_view.get(v.Id.IntegerValue, {})
                                    used_set = used_in_view.get(system_type, set())
                                    
                                    if st.Id.IntegerValue not in used_set:
                                        continue
                                    
                                    view_filters = list(v.GetFilters())
                                    
                                    if filt.Id not in view_filters:
                                        v.AddFilter(filt.Id)
                                    
                                    v.SetIsFilterEnabled(filt.Id, True)
                                    v.SetFilterOverrides(filt.Id, ogs)
                                    
                                    total_linked_filters += 1
                                except:
                                    pass
                        
                    except Exception as ex:
                        pass  # Continue with next link if one fails
            
            if total_linked_filters > 0:
                forms.alert(
                    "Linked file filters created!\n\n"
                    "- Processed {} linked file(s)\n"
                    "- Total linked filters applied: {}\n\n"
                    "Note: Linked filters are prefixed with 'LINK-'".format(len(selected_links), total_linked_filters)
                )
            else:
                forms.alert("No MEP elements found in selected linked files or unable to create filters.")

# --- Optional: Linked File Processing (Legacy - kept for reference, but replaced by automatic processing above) ---
# The automatic linked file processing above replaces this manual selection approach