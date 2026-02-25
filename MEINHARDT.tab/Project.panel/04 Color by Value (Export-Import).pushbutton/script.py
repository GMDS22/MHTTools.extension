# -*- coding: utf-8 -*-
from __future__ import print_function

import csv
import os

import clr
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
import System
from System.Drawing import Color as DrawingColor
from System.Windows.Forms import ColorDialog, DialogResult

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List


logger = script.get_logger()
output = script.get_output()

doc = revit.doc
uidoc = revit.uidoc


PREFIX = 'GMColorByValue'
DELIM_NEW = '__'


def _clamp_int(v, lo, hi, default=None):
    try:
        iv = int(v)
        if iv < lo:
            return lo
        if iv > hi:
            return hi
        return iv
    except Exception:
        return default


def _color_from_text(text):
    # Deterministic but pleasant-ish colors per value.
    try:
        s = (text or '').strip().lower()
    except Exception:
        s = ''
    if not s:
        return DB.Color(200, 200, 200)

    h = 0
    for ch in s:
        h = (h * 33 + ord(ch)) & 0xFFFFFFFF

    # Pastel-ish range
    r = 120 + (h & 0x7F)
    g = 120 + ((h >> 8) & 0x7F)
    b = 120 + ((h >> 16) & 0x7F)
    return DB.Color(int(r), int(g), int(b))


def _sanitize(name):
    try:
        s = str(name)
    except Exception:
        s = 'Value'
    for ch in '{}[]:/\\|?*<>':
        s = s.replace(ch, '-')
    return s.strip()


def _get_solid_fill_id():
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


def _pick_fill_pattern_id():
    choice = forms.CommandSwitchWindow.show(
        ['Solid Fill', 'Pick Fill Pattern'],
        message='Fill pattern for filter overrides:'
    )
    if not choice:
        return None

    if choice == 'Solid Fill':
        return _get_solid_fill_id()

    fps = []
    try:
        fps = list(DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement))
    except Exception:
        fps = []

    if not fps:
        return _get_solid_fill_id()

    fp_map = {}
    for fp in fps:
        try:
            n = fp.Name
        except Exception:
            continue
        key = n
        if key in fp_map:
            key = '{} ({})'.format(n, fp.Id.IntegerValue)
        fp_map[key] = fp

    picked = forms.SelectFromList.show(
        sorted(fp_map.keys()),
        title='Select Fill Pattern',
        multiselect=False
    )
    if not picked:
        return None
    return fp_map[picked].Id


def _pick_views():
    choice = forms.CommandSwitchWindow.show(
        ['Active View', 'Select Views'],
        message='Apply/export filters for which views?'
    )
    if choice == 'Active View':
        return [doc.ActiveView]

    views = [
        v for v in DB.FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType()
        if not getattr(v, 'IsTemplate', False)
    ]
    view_map = {'{} | {}'.format(getattr(v, 'ViewType', ''), getattr(v, 'Name', 'View')): v for v in views}
    picked = forms.SelectFromList.show(sorted(view_map.keys()), multiselect=True, title='Select Views')
    return [view_map[x] for x in picked] if picked else []


def _collect_parameter_filters_by_name():
    m = {}
    for f in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
        try:
            m[f.Name] = f
        except Exception:
            continue
    return m


def _get_first_element_of_category(bic):
    try:
        return (
            DB.FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .FirstElement()
        )
    except Exception:
        return None


def _find_parameter_id_on_element(elem, param_name):
    if elem is None:
        return None
    target = (param_name or '').strip().lower()
    if not target:
        return None

    try:
        for p in elem.Parameters:
            try:
                d = p.Definition
                if d is None:
                    continue
                if (d.Name or '').strip().lower() == target:
                    return p.Id
            except Exception:
                continue
    except Exception:
        pass

    # Fallback: some built-in parameters are only reachable via LookupParameter
    try:
        p = elem.LookupParameter(param_name)
        if p:
            return p.Id
    except Exception:
        pass

    return None


def _make_filter_name(category_label, param_name, value_text):
    # Revit can reject ':' in names. Use a safe delimiter.
    return '{}{}{}{}{}{}{}'.format(
        PREFIX,
        DELIM_NEW,
        _sanitize(category_label),
        DELIM_NEW,
        _sanitize(param_name),
        DELIM_NEW,
        _sanitize(value_text)
    )


def _parse_filter_name(name):
    # Returns (category_label, param_name, value_text) or None
    if not name or not name.startswith(PREFIX):
        return None

    # Legacy format
    if name.startswith(PREFIX + '::'):
        parts = name.split('::')
        if len(parts) < 4:
            return None
        return parts[1], parts[2], '::'.join(parts[3:])

    # New format
    if name.startswith(PREFIX + DELIM_NEW):
        parts = name.split(DELIM_NEW)
        if len(parts) < 4:
            return None
        return parts[1], parts[2], DELIM_NEW.join(parts[3:])

    return None


def _set_overrides(view, filter_id, color, transparency=0, halftone=False, pattern_id=None, line_weight=None):
    ogs = DB.OverrideGraphicSettings()

    try:
        ogs.SetHalftone(bool(halftone))
    except Exception:
        pass

    try:
        t = int(transparency)
        if t < 0:
            t = 0
        if t > 100:
            t = 100
        ogs.SetSurfaceTransparency(t)
    except Exception:
        pass

    pid = pattern_id if pattern_id is not None else _get_solid_fill_id()
    if pid and pid != DB.ElementId.InvalidElementId:
        try:
            ogs.SetSurfaceForegroundPatternId(pid)
            ogs.SetSurfaceForegroundPatternColor(color)
        except Exception:
            # Older APIs
            try:
                ogs.SetProjectionFillPatternId(pid)
                ogs.SetProjectionFillColor(color)
            except Exception:
                pass

    try:
        ogs.SetProjectionLineColor(color)
    except Exception:
        pass

    if line_weight is not None:
        try:
            ogs.SetProjectionLineWeight(int(line_weight))
        except Exception:
            pass

    try:
        view.SetFilterOverrides(filter_id, ogs)
        view.SetFilterVisibility(filter_id, True)
    except Exception:
        pass


def _ensure_filter_applied(view, pfe):
    try:
        if pfe.Id not in view.GetFilters():
            view.AddFilter(pfe.Id)
    except Exception:
        try:
            view.AddFilter(pfe.Id)
        except Exception:
            pass


def _build_equals_rule(param_id, value_text):
    # String equals (case-insensitive by default)
    try:
        return DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, value_text, False)
    except Exception:
        try:
            return DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, value_text)
        except Exception:
            return None


def _create_or_update_filter(filter_name, cat_bic, param_id, value_text, existing_by_name):
    cat_ids = List[DB.ElementId]([DB.ElementId(int(cat_bic))])

    pfe = existing_by_name.get(filter_name)
    if pfe is None:
        try:
            pfe = DB.ParameterFilterElement.Create(doc, filter_name, cat_ids)
        except Exception as ex:
            logger.debug('Failed to create filter %s: %s', filter_name, ex)
            return None
    else:
        # Keep categories in sync (best-effort)
        try:
            pfe.SetCategories(cat_ids)
        except Exception:
            pass

    rule = _build_equals_rule(param_id, value_text)
    if rule is None:
        return pfe

    try:
        elem_filter = DB.ElementParameterFilter(rule)
        pfe.SetElementFilter(elem_filter)
    except Exception as ex:
        logger.debug('Failed to set rule for %s: %s', filter_name, ex)

    return pfe


def _choose_category():
    options = [
        ('MEP Spaces', DB.BuiltInCategory.OST_MEPSpaces),
        ('HVAC Zones', DB.BuiltInCategory.OST_HVAC_Zones),
        ('Rooms', DB.BuiltInCategory.OST_Rooms),
        ('Areas', DB.BuiltInCategory.OST_Areas),
    ]

    labels = [o[0] for o in options]
    picked = forms.SelectFromList.show(labels, title='Select Category', multiselect=False)
    if not picked:
        return None, None

    for label, bic in options:
        if label == picked:
            return label, bic
    return None, None


def _ask_param_and_values(category_label):
    param_name = forms.ask_for_string(
        default='Name',
        prompt='Parameter to color by (exact parameter name):',
        title='Color by Value'
    )
    if not param_name:
        return None, None

    values_raw = forms.ask_for_string(
        default='',
        prompt='Enter value list (comma-separated).\n\nExample: Zone A, Zone B, Zone C',
        title='Values'
    )
    if values_raw is None:
        return None, None

    values = [v.strip() for v in (values_raw or '').split(',') if v.strip()]
    if not values:
        forms.alert('No values provided.', exitscript=True)

    return param_name, values


def _ask_color_for_value(value_text):
    # Minimal: prompt user for RGB
    rgb = forms.ask_for_string(
        default='200,200,200',
        prompt='Enter RGB for "{}" as R,G,B (0-255):'.format(value_text),
        title='Pick Color'
    )
    if not rgb:
        return None
    try:
        parts = [int(x.strip()) for x in rgb.split(',')]
        if len(parts) != 3:
            return None
        r, g, b = parts
        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))
        return DB.Color(r, g, b)
    except Exception:
        return None


def _parse_cschn_lines(lines):
    # Format: VALUE::R###G###B###
    mapping = {}
    for raw in (lines or []):
        try:
            line = (raw or '').strip()
        except Exception:
            continue
        if not line or '::' not in line:
            continue
        value, rgb = line.split('::', 1)
        value = (value or '').strip()
        if not value:
            continue

        try:
            rgb = rgb.strip()
            # tolerate lowercase
            rgb_u = rgb.upper()
            r_i = rgb_u.index('R')
            g_i = rgb_u.index('G')
            b_i = rgb_u.index('B')
            r = int(rgb_u[r_i + 1:g_i])
            g = int(rgb_u[g_i + 1:b_i])
            b = int(rgb_u[b_i + 1:])
            r = max(0, min(255, r))
            g = max(0, min(255, g))
            b = max(0, min(255, b))
            mapping[value] = DB.Color(r, g, b)
        except Exception:
            continue
    return mapping


def _load_cschn(path):
    try:
        with open(path, 'rb') as f:
            # IronPython: bytes -> str
            content = f.read()
        try:
            text = content.decode('utf-8')
        except Exception:
            text = content.decode('utf-8', 'ignore')
        return _parse_cschn_lines(text.splitlines())
    except Exception:
        return {}


def _find_zone_param_on_space(space):
    # Best-effort parameter discovery on MEP Space.
    if space is None:
        return None

    candidates = []
    try:
        for p in space.Parameters:
            try:
                d = p.Definition
                if d is None:
                    continue
                n = (d.Name or '').strip()
                if not n:
                    continue
                nl = n.lower()
                if 'zone' in nl:
                    candidates.append(n)
            except Exception:
                continue
    except Exception:
        candidates = []

    # Prioritize common names
    priority = ['hvac zone', 'zone name', 'zone']
    for key in priority:
        for n in candidates:
            if n.lower() == key:
                return n

    if candidates:
        picked = forms.SelectFromList.show(
            sorted(set(candidates)),
            title='Select Space Parameter for Zone',
            multiselect=False
        )
        return picked

    # Fall back to asking user
    return forms.ask_for_string(
        default='Zone Name',
        prompt='Zone parameter name on Spaces (exact):',
        title='HVAC Zone Filters'
    )


def _collect_unique_values(cat_bic, param_name, views=None):
    # If views provided, try collecting in those views (reduces noise).
    values = set()
    try:
        collectors = []
        if views:
            for v in views:
                try:
                    collectors.append(DB.FilteredElementCollector(doc, v.Id))
                except Exception:
                    continue
        else:
            collectors = [DB.FilteredElementCollector(doc)]

        for cl in collectors:
            try:
                elems = cl.OfCategory(cat_bic).WhereElementIsNotElementType().ToElements()
            except Exception:
                continue
            for e in elems:
                try:
                    p = e.LookupParameter(param_name)
                    if not p:
                        continue
                    s = None
                    try:
                        s = p.AsString()
                    except Exception:
                        s = None
                    if s is None:
                        try:
                            s = p.AsValueString()
                        except Exception:
                            s = None
                    if s:
                        values.add(s.strip())
                except Exception:
                    continue
    except Exception:
        pass

    return sorted([v for v in values if v])


def _views_have_templates(views):
    for v in views or []:
        try:
            if getattr(v, 'IsTemplate', False):
                continue
            if v.ViewTemplateId and v.ViewTemplateId != DB.ElementId.InvalidElementId:
                return True
        except Exception:
            continue
    return False


def _resolve_target_views(doc, views, template_policy):
    """Return list of views/templates to edit.

    template_policy:
      - 'Apply to templates'
      - 'Detach templates'
      - 'Skip templated views'
    """
    targets = []
    templates_seen = set()

    for v in views or []:
        if not v:
            continue

        try:
            if getattr(v, 'IsTemplate', False):
                if v.Id.IntegerValue not in templates_seen:
                    templates_seen.add(v.Id.IntegerValue)
                    targets.append(v)
                continue
        except Exception:
            pass

        vtid = None
        try:
            vtid = v.ViewTemplateId
        except Exception:
            vtid = DB.ElementId.InvalidElementId

        if vtid and vtid != DB.ElementId.InvalidElementId:
            if template_policy == 'Apply to templates':
                vt = doc.GetElement(vtid)
                if vt and vt.Id.IntegerValue not in templates_seen:
                    templates_seen.add(vt.Id.IntegerValue)
                    targets.append(vt)
                continue
            elif template_policy == 'Detach templates':
                try:
                    v.ViewTemplateId = DB.ElementId.InvalidElementId
                except Exception:
                    continue
                targets.append(v)
            else:
                # Skip templated
                continue
        else:
            targets.append(v)

    return targets


class _ZoneRow(object):
    def __init__(self, value, color=None):
        self.Include = True
        self.Value = value
        if color is None:
            color = _color_from_text(value)
        self.R = int(getattr(color, 'Red', 200))
        self.G = int(getattr(color, 'Green', 200))
        self.B = int(getattr(color, 'Blue', 200))

    def to_color(self):
        r = _clamp_int(self.R, 0, 255, default=200)
        g = _clamp_int(self.G, 0, 255, default=200)
        b = _clamp_int(self.B, 0, 255, default=200)
        return DB.Color(int(r), int(g), int(b))


class _HVACZonesWindow(forms.WPFWindow):
    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), 'HVACZonesFilters.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        self._picked_views = []
        self._cschn_map = {}
        self._fill_pattern_map = {}
        self._zone_param_candidates = []

        # init template policy
        self.cbTemplatePolicy.ItemsSource = ['Apply to templates', 'Detach templates', 'Skip templated views']
        self.cbTemplatePolicy.SelectedIndex = 0

        # init color source
        self.cbColorSource.ItemsSource = ['Deterministic from Zone Name', 'Colors from CSCHN file', 'One global color']
        self.cbColorSource.SelectedIndex = 0

        # fill patterns
        self._populate_fill_patterns()

        # zone params
        self._populate_zone_params()

        self._update_views_text()
        self._update_zone_count()
        self._toggle_cschn_controls()
        self._toggle_global_rgb_controls()

        # events
        self.btnPickViews.Click += self._on_pick_views
        self.rbActiveView.Checked += self._on_view_mode_changed
        self.rbSelectViews.Checked += self._on_view_mode_changed
        self.btnPickCschn.Click += self._on_pick_cschn
        self.btnPickGlobal.Click += self._on_pick_global
        self.cbColorSource.SelectionChanged += self._on_color_source_changed
        self.btnReload.Click += self._on_reload
        self.btnApplyColors.Click += self._on_apply_colors
        self.btnAllOn.Click += self._on_all_on
        self.btnAllOff.Click += self._on_all_off
        self.btnRun.Click += self._on_run
        self.btnCancel.Click += self._on_cancel

    def _on_cancel(self, sender, args):
        self.Close()

    def _on_view_mode_changed(self, sender, args):
        self._update_views_text()

    def _on_pick_views(self, sender, args):
        views = forms.select_views(title='Select Views', multiple=True, use_selection=True)
        self._picked_views = list(views or [])
        self.rbSelectViews.IsChecked = True
        self._update_views_text()

    def _on_pick_cschn(self, sender, args):
        path = forms.pick_file(file_ext='cschn', title='Pick .cschn file (VALUE::R###G###B###)')
        if not path:
            return
        self.txtCschn.Text = path
        self._cschn_map = _load_cschn(path)

    def _on_color_source_changed(self, sender, args):
        self._toggle_cschn_controls()
        self._toggle_global_rgb_controls()

    def _toggle_cschn_controls(self):
        mode = self.cbColorSource.SelectedItem
        enabled = (mode == 'Colors from CSCHN file')
        self.txtCschn.IsEnabled = enabled
        self.btnPickCschn.IsEnabled = enabled

    def _toggle_global_rgb_controls(self):
        mode = self.cbColorSource.SelectedItem
        enabled = (mode == 'One global color')
        self.txtR.IsEnabled = enabled
        self.txtG.IsEnabled = enabled
        self.txtB.IsEnabled = enabled
        try:
            self.btnPickGlobal.IsEnabled = enabled
        except Exception:
            pass

    def _pick_color_dialog(self, initial_color=None):
        dlg = ColorDialog()
        try:
            dlg.FullOpen = True
        except Exception:
            pass

        if initial_color is not None:
            try:
                dlg.Color = DrawingColor.FromArgb(
                    int(initial_color.Red), int(initial_color.Green), int(initial_color.Blue)
                )
            except Exception:
                pass

        try:
            res = dlg.ShowDialog()
        except Exception:
            # Some environments require owner; fallback
            res = dlg.ShowDialog()

        if res == DialogResult.OK:
            c = dlg.Color
            return DB.Color(int(c.R), int(c.G), int(c.B))
        return None

    def _on_pick_global(self, sender, args):
        c0 = None
        try:
            c0 = DB.Color(
                _clamp_int(self.txtR.Text, 0, 255, default=200),
                _clamp_int(self.txtG.Text, 0, 255, default=200),
                _clamp_int(self.txtB.Text, 0, 255, default=200)
            )
        except Exception:
            c0 = None

        c = self._pick_color_dialog(initial_color=c0)
        if c is None:
            return
        self.txtR.Text = str(c.Red)
        self.txtG.Text = str(c.Green)
        self.txtB.Text = str(c.Blue)

    # Called from XAML Button Click
    def pick_row_color(self, sender, args):
        try:
            zr = sender.Tag
        except Exception:
            zr = None
        if zr is None:
            return

        c0 = None
        try:
            c0 = DB.Color(
                _clamp_int(getattr(zr, 'R', 200), 0, 255, default=200),
                _clamp_int(getattr(zr, 'G', 200), 0, 255, default=200),
                _clamp_int(getattr(zr, 'B', 200), 0, 255, default=200)
            )
        except Exception:
            c0 = None

        c = self._pick_color_dialog(initial_color=c0)
        if c is None:
            return

        zr.R = int(c.Red)
        zr.G = int(c.Green)
        zr.B = int(c.Blue)
        try:
            self.dgZones.Items.Refresh()
        except Exception:
            pass

    def _update_views_text(self):
        if self.rbActiveView.IsChecked:
            self.txtViews.Text = 'Active view'
        else:
            self.txtViews.Text = '{} view(s) selected'.format(len(self._picked_views))

    def _update_zone_count(self):
        try:
            cnt = len(list(self.dgZones.ItemsSource or []))
        except Exception:
            cnt = 0
        self.txtZoneCount.Text = '{} zones'.format(cnt)

    def _populate_fill_patterns(self):
        items = []
        items.append('Solid Fill')
        try:
            fps = list(DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement))
        except Exception:
            fps = []

        self._fill_pattern_map = {'Solid Fill': _get_solid_fill_id()}
        for fp in fps:
            try:
                name = fp.Name
            except Exception:
                continue
            key = name
            if key in self._fill_pattern_map:
                key = '{} ({})'.format(name, fp.Id.IntegerValue)
            self._fill_pattern_map[key] = fp.Id
            items.append(key)

        self.cbFillPattern.ItemsSource = sorted(set(items), key=lambda x: (x != 'Solid Fill', x))
        self.cbFillPattern.SelectedItem = 'Solid Fill'

    def _populate_zone_params(self):
        sample_space = _get_first_element_of_category(DB.BuiltInCategory.OST_MEPSpaces)
        candidates = []
        if sample_space:
            try:
                for p in sample_space.Parameters:
                    try:
                        d = p.Definition
                        if not d:
                            continue
                        n = (d.Name or '').strip()
                        if n and ('zone' in n.lower()):
                            candidates.append(n)
                    except Exception:
                        continue
            except Exception:
                pass

        # Add common defaults at top if present
        common = ['HVAC Zone', 'Zone Name', 'Zone']
        ordered = []
        for c in common:
            if c in candidates and c not in ordered:
                ordered.append(c)
        for c in sorted(set(candidates)):
            if c not in ordered:
                ordered.append(c)

        self._zone_param_candidates = ordered
        self.cbZoneParam.ItemsSource = ordered
        if ordered:
            self.cbZoneParam.SelectedIndex = 0

    def _get_selected_views(self):
        if self.rbActiveView.IsChecked:
            return [doc.ActiveView]
        return list(self._picked_views)

    def _get_selected_zone_param(self):
        try:
            return self.cbZoneParam.SelectedItem
        except Exception:
            return None

    def _load_zones(self):
        views = self._get_selected_views()
        if not views:
            forms.alert('No views selected.')
            return

        zone_param = self._get_selected_zone_param()
        if not zone_param:
            forms.alert('Pick a zone parameter first.')
            return

        values = _collect_unique_values(DB.BuiltInCategory.OST_MEPSpaces, zone_param, views=views)
        if not values:
            # fallback: all spaces in model
            values = _collect_unique_values(DB.BuiltInCategory.OST_MEPSpaces, zone_param, views=None)

        rows = []
        for v in values:
            rows.append(_ZoneRow(v))
        self.dgZones.ItemsSource = rows
        self._update_zone_count()

    def _apply_colors_to_rows(self):
        rows = list(self.dgZones.ItemsSource or [])
        if not rows:
            return

        mode = self.cbColorSource.SelectedItem
        if mode == 'Colors from CSCHN file':
            path = (self.txtCschn.Text or '').strip()
            if not path:
                forms.alert('Select a .cschn file first.')
                return
            if not self._cschn_map:
                self._cschn_map = _load_cschn(path)
            if not self._cschn_map:
                forms.alert('No valid colors found in that .cschn file.')
                return
            for r in rows:
                c = self._cschn_map.get(r.Value)
                if c is None:
                    continue
                r.R = int(c.Red)
                r.G = int(c.Green)
                r.B = int(c.Blue)

        elif mode == 'One global color':
            r0 = _clamp_int(self.txtR.Text, 0, 255, default=200)
            g0 = _clamp_int(self.txtG.Text, 0, 255, default=200)
            b0 = _clamp_int(self.txtB.Text, 0, 255, default=200)
            for r in rows:
                r.R = int(r0)
                r.G = int(g0)
                r.B = int(b0)

        else:
            # Deterministic
            for r in rows:
                c = _color_from_text(r.Value)
                r.R = int(c.Red)
                r.G = int(c.Green)
                r.B = int(c.Blue)

        self.dgZones.Items.Refresh()

    def _on_reload(self, sender, args):
        self._load_zones()

    def _on_apply_colors(self, sender, args):
        self._apply_colors_to_rows()

    def _on_all_on(self, sender, args):
        for r in list(self.dgZones.ItemsSource or []):
            r.Include = True
        self.dgZones.Items.Refresh()

    def _on_all_off(self, sender, args):
        for r in list(self.dgZones.ItemsSource or []):
            r.Include = False
        self.dgZones.Items.Refresh()

    def _on_run(self, sender, args):
        views = self._get_selected_views()
        if not views:
            forms.alert('No views selected.')
            return

        zone_param = self._get_selected_zone_param()
        if not zone_param:
            forms.alert('Pick a zone parameter first.')
            return

        rows = [r for r in list(self.dgZones.ItemsSource or []) if getattr(r, 'Include', True)]
        if not rows:
            forms.alert('No zones selected.')
            return

        sample_space = _get_first_element_of_category(DB.BuiltInCategory.OST_MEPSpaces)
        if sample_space is None:
            forms.alert('No MEP Spaces found in this model.')
            return

        param_id = _find_parameter_id_on_element(sample_space, zone_param)
        if param_id is None:
            forms.alert('Parameter "{}" not found on MEP Spaces.'.format(zone_param))
            return

        # graphics
        transparency = _clamp_int(self.txtTransparency.Text, 0, 100, default=0)
        halftone = bool(self.chkHalftone.IsChecked)
        lw = (self.txtLineWeight.Text or '').strip()
        line_weight = _clamp_int(lw, 1, 16, default=None) if lw else None

        pat_key = self.cbFillPattern.SelectedItem
        pattern_id = self._fill_pattern_map.get(pat_key, _get_solid_fill_id())

        template_policy = self.cbTemplatePolicy.SelectedItem or 'Apply to templates'
        targets = _resolve_target_views(doc, views, template_policy)
        if not targets:
            forms.alert('No editable target views/templates found for the selected policy.')
            return

        existing_by_name = _collect_parameter_filters_by_name()

        created = 0
        applied_ok = 0
        applied_fail = 0
        skipped_views = 0

        with revit.Transaction('HVAC Zone Filters (Create + Apply)'):
            for zr in rows:
                val = zr.Value
                c = zr.to_color()

                fname = _make_filter_name('HVAC Zones (Spaces)', zone_param, val)
                pfe = _create_or_update_filter(fname, DB.BuiltInCategory.OST_MEPSpaces, param_id, val, existing_by_name)
                if pfe is None:
                    continue
                existing_by_name[pfe.Name] = pfe
                created += 1

                for v in targets:
                    # Don't swallow failures; count them.
                    try:
                        _ensure_filter_applied(v, pfe)
                        _set_overrides(v, pfe.Id, c, transparency=transparency, halftone=halftone, pattern_id=pattern_id, line_weight=line_weight)
                        applied_ok += 1
                    except Exception:
                        applied_fail += 1
                        continue

        forms.alert(
            'Done.\n\n'
            'Filters created/updated: {}\n'
            'Applied to targets: OK={}  Failed={}\n'
            'Targets: {} (from {} selected view(s))'.format(created, applied_ok, applied_fail, len(targets), len(views))
        )
        self.Close()


def _export_csv(path, rows):
    with open(path, 'wb') as f:
        writer = csv.writer(f)
        writer.writerow(['FilterName', 'Category', 'Parameter', 'Value', 'R', 'G', 'B', 'Transparency', 'Halftone'])
        for r in rows:
            writer.writerow(r)


def _import_csv(path):
    rows = []
    with open(path, 'rb') as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def _get_filter_color_from_view(view, filter_id):
    try:
        ogs = view.GetFilterOverrides(filter_id)
    except Exception:
        return None, None, None, None

    # Try the common getters
    try:
        c = ogs.ProjectionLineColor
        if c and c.IsValid:
            return c.Red, c.Green, c.Blue, getattr(ogs, 'SurfaceTransparency', 0)
    except Exception:
        pass

    # Fall back: best-effort
    return None, None, None, None


def main():
    action = forms.CommandSwitchWindow.show(
        ['Create/Update (Interactive)', 'Create Filters for HVAC Zones', 'Export (CSV for Excel)', 'Import (CSV from Excel)'],
        message='What do you want to do?'
    )
    if not action:
        return

    # HVAC Zones special mode (single UI)
    if action == 'Create Filters for HVAC Zones':
        w = _HVACZonesWindow()
        w.ShowDialog()
        return

    views = _pick_views()
    if not views:
        return

    # Normal flow: choose category
    category_label, cat_bic = _choose_category()
    if not cat_bic:
        return

    if action.startswith('Export'):
        out_path = forms.save_file(file_ext='csv', default_name='{}_{}_filters.csv'.format(PREFIX, _sanitize(category_label)))
        if not out_path:
            return

        # Export filters from the FIRST chosen view only (avoids duplicates).
        view = views[0]
        rows = []
        for fid in view.GetFilters():
            f = doc.GetElement(fid)
            if f is None:
                continue
            try:
                name = f.Name
            except Exception:
                continue
            parsed = _parse_filter_name(name)
            if not parsed:
                continue

            cat_lbl, param_name, value_text = parsed
            if cat_lbl != category_label:
                continue

            r, g, b, transparency = _get_filter_color_from_view(view, fid)
            if r is None:
                continue

            # Halftone isn't reliably readable across versions; export blank.
            rows.append([name, cat_lbl, param_name, value_text, r, g, b, transparency or 0, ''])

        _export_csv(out_path, rows)
        forms.alert('Exported {} filter rows to:\n\n{}'.format(len(rows), out_path))
        return

    if action.startswith('Import'):
        in_path = forms.pick_file(file_ext='csv', title='Select CSV to Import')
        if not in_path:
            return

        data = _import_csv(in_path)
        if not data:
            forms.alert('No rows found in CSV.', exitscript=True)

        clear_existing = forms.alert('Clear existing {} filters in target views first?'.format(PREFIX), yes=True, no=True)

        existing_by_name = _collect_parameter_filters_by_name()

        sample_elem = _get_first_element_of_category(cat_bic)
        if sample_elem is None:
            forms.alert('No elements found in category {} in this model.'.format(category_label), exitscript=True)

        with revit.Transaction('Import Color-by-Value Filters'):
            if clear_existing:
                for v in views:
                    try:
                        for fid in list(v.GetFilters()):
                            f = doc.GetElement(fid)
                            if f is None:
                                continue
                            n = getattr(f, 'Name', '')
                            if n.startswith(PREFIX + '::') or n.startswith(PREFIX + DELIM_NEW):
                                try:
                                    v.RemoveFilter(fid)
                                except Exception:
                                    pass
                    except Exception:
                        pass

            for row in data:
                try:
                    fname = (row.get('FilterName') or '').strip()
                    cat_lbl = (row.get('Category') or '').strip()
                    param_name = (row.get('Parameter') or '').strip()
                    value_text = (row.get('Value') or '').strip()

                    if not value_text or not param_name:
                        continue

                    # Respect current category selection; skip others.
                    if cat_lbl and cat_lbl != category_label:
                        continue

                    if not fname:
                        fname = _make_filter_name(category_label, param_name, value_text)

                    param_id = _find_parameter_id_on_element(sample_elem, param_name)
                    if param_id is None:
                        logger.debug('Parameter not found: %s', param_name)
                        continue

                    r = int(row.get('R') or 0)
                    g = int(row.get('G') or 0)
                    b = int(row.get('B') or 0)
                    r = max(0, min(255, r))
                    g = max(0, min(255, g))
                    b = max(0, min(255, b))
                    color = DB.Color(r, g, b)

                    transparency = int(row.get('Transparency') or 0)
                    halftone = str(row.get('Halftone') or '').strip().lower() in ['1', 'true', 'yes', 'y']

                    pfe = _create_or_update_filter(fname, cat_bic, param_id, value_text, existing_by_name)
                    if pfe is None:
                        continue
                    existing_by_name[pfe.Name] = pfe

                    for v in views:
                        _ensure_filter_applied(v, pfe)
                        _set_overrides(v, pfe.Id, color, transparency=transparency, halftone=halftone)
                except Exception as ex:
                    logger.debug('Row import failed: %s', ex)
                    continue

        forms.alert('Import complete. Check the target view(s) Filters dialog.')
        return

    # Interactive create/update
    param_name, values = _ask_param_and_values(category_label)
    if not param_name:
        return

    sample_elem = _get_first_element_of_category(cat_bic)
    if sample_elem is None:
        forms.alert('No elements found in category {} in this model.'.format(category_label), exitscript=True)

    param_id = _find_parameter_id_on_element(sample_elem, param_name)
    if param_id is None:
        forms.alert('Parameter "{}" not found on category {}.'.format(param_name, category_label), exitscript=True)

    transparency = forms.ask_for_string(default='0', prompt='Transparency 0-100:', title='Graphics')
    halftone = forms.alert('Halftone?', yes=True, no=True)

    existing_by_name = _collect_parameter_filters_by_name()

    created = 0
    with revit.Transaction('Create/Update Color-by-Value Filters'):
        for val in values:
            c = _ask_color_for_value(val)
            if c is None:
                continue

            fname = _make_filter_name(category_label, param_name, val)
            pfe = _create_or_update_filter(fname, cat_bic, param_id, val, existing_by_name)
            if pfe is None:
                continue
            existing_by_name[pfe.Name] = pfe

            for v in views:
                _ensure_filter_applied(v, pfe)
                _set_overrides(v, pfe.Id, c, transparency=transparency or 0, halftone=halftone)

            created += 1

    forms.alert('Created/updated {} filters in {} view(s).'.format(created, len(views)))


if __name__ == '__main__':
    main()
