# -*- coding: utf-8 -*-
from __future__ import print_function

import csv
import os
import re
import colorsys

from pyrevit import revit, DB, forms, script


logger = script.get_logger()
output = script.get_output()

doc = revit.doc


def _clamp255(v):
    try:
        v = int(v)
    except Exception:
        v = 0
    return max(0, min(255, v))


def _lighten_color(c, pct):
    # pct in [0..100]
    try:
        p = max(0.0, min(1.0, float(pct) / 100.0))
    except Exception:
        p = 0.0
    try:
        r = int(round(c.Red + (255 - c.Red) * p))
        g = int(round(c.Green + (255 - c.Green) * p))
        b = int(round(c.Blue + (255 - c.Blue) * p))
        return DB.Color(_clamp255(r), _clamp255(g), _clamp255(b))
    except Exception:
        return c


def _darken_color(c, pct):
    # pct in [0..100]
    try:
        p = max(0.0, min(1.0, float(pct) / 100.0))
    except Exception:
        p = 0.0
    try:
        r = int(round(c.Red * (1.0 - p)))
        g = int(round(c.Green * (1.0 - p)))
        b = int(round(c.Blue * (1.0 - p)))
        return DB.Color(_clamp255(r), _clamp255(g), _clamp255(b))
    except Exception:
        return c


def _adjust_saturation_color(c, pct, increase=True):
    """Adjust saturation while preserving value/brightness.

    pct in [0..100]. If increase=True, saturation moves toward 1.0.
    If increase=False, saturation moves toward 0.0.
    """
    try:
        p = max(0.0, min(1.0, float(pct) / 100.0))
    except Exception:
        p = 0.0

    try:
        rf = max(0.0, min(1.0, float(c.Red) / 255.0))
        gf = max(0.0, min(1.0, float(c.Green) / 255.0))
        bf = max(0.0, min(1.0, float(c.Blue) / 255.0))

        h, s, v = colorsys.rgb_to_hsv(rf, gf, bf)
        if increase:
            s = min(1.0, s * (1.0 + p))
        else:
            s = max(0.0, s * (1.0 - p))

        nr, ng, nb = colorsys.hsv_to_rgb(h, s, v)
        return DB.Color(
            _clamp255(int(round(nr * 255.0))),
            _clamp255(int(round(ng * 255.0))),
            _clamp255(int(round(nb * 255.0))),
        )
    except Exception:
        return c


def _adjust_value_color(c, pct, increase=True):
    """Adjust brightness/value while preserving hue.

    pct in [0..100]. If increase=True, value moves toward 1.0.
    If increase=False, value moves toward 0.0.
    """
    try:
        p = max(0.0, min(1.0, float(pct) / 100.0))
    except Exception:
        p = 0.0

    try:
        rf = max(0.0, min(1.0, float(c.Red) / 255.0))
        gf = max(0.0, min(1.0, float(c.Green) / 255.0))
        bf = max(0.0, min(1.0, float(c.Blue) / 255.0))

        h, s, v = colorsys.rgb_to_hsv(rf, gf, bf)
        if increase:
            v = min(1.0, v + (1.0 - v) * p)
        else:
            v = max(0.0, v * (1.0 - p))

        nr, ng, nb = colorsys.hsv_to_rgb(h, s, v)
        return DB.Color(
            _clamp255(int(round(nr * 255.0))),
            _clamp255(int(round(ng * 255.0))),
            _clamp255(int(round(nb * 255.0))),
        )
    except Exception:
        return c


def _adjust_rgb_offsets(c, dr, dg, db):
    try:
        return DB.Color(
            _clamp255(int(c.Red) + int(dr)),
            _clamp255(int(c.Green) + int(dg)),
            _clamp255(int(c.Blue) + int(db)),
        )
    except Exception:
        return c


def _ask_rgb_delta(default='0,0,0', title='RGB Offset', prompt='Enter RGB offsets as dR,dG,dB (e.g. 10,-5,0):'):
    s = forms.ask_for_string(default=default, title=title, prompt=prompt)
    if not s:
        return None
    try:
        parts = [int(x.strip()) for x in s.split(',')]
        if len(parts) != 3:
            return None
        return parts[0], parts[1], parts[2]
    except Exception:
        return None


def _category_name_from_id(cat_id):
    try:
        cat = doc.Settings.Categories.get_Item(cat_id)
        if cat:
            return cat.Name
    except Exception:
        pass
    try:
        cat = doc.GetElement(cat_id)
        if cat is not None and hasattr(cat, 'Name'):
            return cat.Name
    except Exception:
        pass
    return 'Category'


def _collect_color_fill_schemes():
    try:
        return list(DB.FilteredElementCollector(doc).OfClass(DB.ColorFillScheme).ToElements())
    except Exception:
        return []


class _SchemeItem(object):
    def __init__(self, scheme):
        self.scheme = scheme
        try:
            self.name = scheme.Name
        except Exception:
            self.name = 'Color Fill Scheme'

        try:
            self.category_id = scheme.CategoryId
        except Exception:
            self.category_id = None

        self.category_name = _category_name_from_id(self.category_id) if self.category_id else 'Category'

        # Best-effort info (not always available)
        try:
            self.entry_count = len(list(scheme.GetEntries() or []))
        except Exception:
            self.entry_count = 0

        self.display = '{} | {} ({} entries)'.format(self.category_name, self.name, self.entry_count)


def _try_get_entries(scheme):
    try:
        return list(scheme.GetEntries() or [])
    except Exception:
        return []


def _entry_label(entry, idx=None):
    # A stable identifier for the user to recognize (caption/value)
    for attr in ['Caption', 'caption', 'Value', 'value', 'StringValue', 'stringValue', 'Name', 'name']:
        try:
            v = getattr(entry, attr)
            if v:
                return str(v)
        except Exception:
            pass

    # Some entries expose a method
    for meth in ['GetCaption', 'get_Caption', 'GetStringValue', 'AsString']:
        try:
            if hasattr(entry, meth):
                v = getattr(entry, meth)()
                if v:
                    return str(v)
        except Exception:
            pass

    if idx is not None:
        return 'Entry {}'.format(idx + 1)
    return 'Entry'


def _entry_color(entry):
    try:
        return entry.Color
    except Exception:
        try:
            return entry.get_Color()
        except Exception:
            return None


def _entry_visible(entry):
    for attr in ['IsVisible', 'Visible', 'isVisible', 'visible']:
        try:
            v = getattr(entry, attr)
            if v is None:
                continue
            return bool(v)
        except Exception:
            continue
    for meth in ['get_IsVisible', 'GetIsVisible']:
        try:
            if hasattr(entry, meth):
                return bool(getattr(entry, meth)())
        except Exception:
            continue
    return True


def _set_entry_visible(entry, is_visible):
    for attr in ['IsVisible', 'Visible']:
        try:
            setattr(entry, attr, bool(is_visible))
            return True
        except Exception:
            continue
    for meth in ['set_IsVisible', 'SetIsVisible']:
        try:
            if hasattr(entry, meth):
                getattr(entry, meth)(bool(is_visible))
                return True
        except Exception:
            continue
    return False


def _set_entry_color(entry, new_color):
    # Some APIs allow direct set, some require SetEntries.
    try:
        entry.Color = new_color
        return True
    except Exception:
        try:
            entry.set_Color(new_color)
            return True
        except Exception:
            return False


def _commit_entries(scheme, entries):
    # Prefer setting entries back to scheme when API supports it.
    for meth in ['SetEntries', 'set_Entries']:
        try:
            if hasattr(scheme, meth):
                getattr(scheme, meth)(entries)
                return True
        except Exception:
            pass
    return False


def _pick_scheme(category_filter=None):
    schemes = _collect_color_fill_schemes()
    if not schemes:
        forms.alert('No Color Fill Schemes found in this model.', exitscript=True)

    items = []
    for s in schemes:
        it = _SchemeItem(s)
        if category_filter:
            if it.category_name != category_filter:
                continue
        items.append(it)

    if not items:
        forms.alert('No Color Fill Schemes found for that category.', exitscript=True)

    picked = forms.SelectFromList.show(
        items,
        name_attr='display',
        multiselect=False,
        title='Select Color Fill Scheme'
    )
    return picked.scheme if picked else None


def _pick_category_name():
    # Only categories that actually have at least one scheme.
    schemes = _collect_color_fill_schemes()
    cats = []
    seen = set()
    for s in schemes:
        try:
            cid = s.CategoryId
        except Exception:
            cid = None
        if not cid:
            continue
        nm = _category_name_from_id(cid)
        if nm not in seen:
            seen.add(nm)
            cats.append(nm)

    cats = sorted(cats)
    if not cats:
        return None

    # Put common ones first (if present)
    preferred = ['HVAC Zones', 'Spaces', 'Rooms', 'Areas']
    ordered = []
    for p in preferred:
        if p in cats:
            ordered.append(p)
    for c in cats:
        if c not in ordered:
            ordered.append(c)

    picked = forms.SelectFromList.show(ordered, multiselect=False, title='Select Scheme Category')
    return picked


def _ask_rgb(default='200,200,200', title='Set Color', prompt='Enter RGB as R,G,B (0-255):'):
    s = forms.ask_for_string(default=default, title=title, prompt=prompt)
    if not s:
        return None
    try:
        parts = [int(x.strip()) for x in s.split(',')]
        if len(parts) != 3:
            return None
        return DB.Color(_clamp255(parts[0]), _clamp255(parts[1]), _clamp255(parts[2]))
    except Exception:
        return None


def _export_scheme_csv(path, scheme, entries):
    rows = []
    scheme_name = getattr(scheme, 'Name', 'Scheme')
    cat_name = _category_name_from_id(getattr(scheme, 'CategoryId', None))

    for i, e in enumerate(entries):
        c = _entry_color(e)
        if c is None:
            continue
        vis = _entry_visible(e)
        rows.append([
            cat_name,
            scheme_name,
            _entry_label(e, idx=i),
            1 if vis else 0,
            c.Red,
            c.Green,
            c.Blue,
        ])

    with open(path, 'wb') as f:
        w = csv.writer(f)
        w.writerow(['Category', 'Scheme', 'Value', 'Visible', 'R', 'G', 'B'])
        for r in rows:
            w.writerow(r)


def _import_scheme_csv(path):
    with open(path, 'rb') as f:
        return list(csv.DictReader(f))


def _export_scheme_cschn(path, scheme, entries):
    # Observed format:
    #   VALUE::R127G153B197
    # One entry per line.
    lines = []
    for i, e in enumerate(entries):
        c = _entry_color(e)
        if c is None:
            continue
        value = _entry_label(e, idx=i)
        lines.append('{}::R{}G{}B{}'.format(value, int(c.Red), int(c.Green), int(c.Blue)))

    # Keep file stable/orderly
    try:
        lines = sorted(lines)
    except Exception:
        pass

    with open(path, 'wb') as f:
        for ln in lines:
            try:
                f.write((ln + '\r\n').encode('utf-8'))
            except Exception:
                try:
                    f.write(ln + '\r\n')
                except Exception:
                    pass


def _safe_filename(text):
    try:
        t = (text or '').strip()
    except Exception:
        t = ''
    if not t:
        return 'unnamed'
    t = re.sub(r'[\\/:*?"<>|]+', '_', t)
    t = re.sub(r'\s+', ' ', t)
    return t[:120]


def _backup_dir():
    return os.path.join(os.path.dirname(__file__), 'backups')


def _backup_path_for_scheme(scheme):
    try:
        sid = str(scheme.Id.IntegerValue)
    except Exception:
        sid = '0'

    try:
        sname = _safe_filename(getattr(scheme, 'Name', 'Scheme'))
    except Exception:
        sname = 'Scheme'

    try:
        cname = _safe_filename(_category_name_from_id(getattr(scheme, 'CategoryId', None)))
    except Exception:
        cname = 'Category'

    fname = '{}__{}__{}.cschn'.format(cname, sname, sid)
    return os.path.join(_backup_dir(), fname)


def _backup_scheme_colors(scheme, entries):
    try:
        bdir = _backup_dir()
        if not os.path.isdir(bdir):
            os.makedirs(bdir)
        bpath = _backup_path_for_scheme(scheme)
        _export_scheme_cschn(bpath, scheme, entries)
        return bpath
    except Exception:
        return None


def _import_scheme_cschn(path):
    # Returns dict: value_lower -> (r,g,b)
    data = {}
    try:
        with open(path, 'rb') as f:
            raw = f.read()
    except Exception:
        return data

    try:
        text = raw.decode('utf-8')
    except Exception:
        try:
            text = raw.decode('utf-16')
        except Exception:
            try:
                text = raw.decode('cp1252')
            except Exception:
                text = str(raw)

    for line in (text or '').splitlines():
        try:
            ln = (line or '').strip()
            if not ln:
                continue

            # VALUE::R127G153B197
            if '::' in ln:
                value_part, rgb_part = ln.split('::', 1)
            elif ':' in ln:
                value_part, rgb_part = ln.split(':', 1)
            else:
                continue

            value = (value_part or '').strip()
            if not value:
                continue

            m = re.search(r'R\s*(\d+)\s*G\s*(\d+)\s*B\s*(\d+)', rgb_part, re.IGNORECASE)
            if not m:
                continue
            r, g, b = m.group(1), m.group(2), m.group(3)
            data[value.strip().lower()] = (_clamp255(r), _clamp255(g), _clamp255(b))
        except Exception:
            continue

    return data


def _generate_contrasting_colors(num_colors, saturation=0.8, value=0.9):
    """Generate a list of contrasting colors by evenly spacing hues in HSV space."""
    colors = []
    for i in range(num_colors):
        hue = (i * 360.0) / num_colors  # Evenly distribute hues
        r, g, b = colorsys.hsv_to_rgb(hue / 360.0, saturation, value)
        colors.append(DB.Color(
            _clamp255(int(round(r * 255))),
            _clamp255(int(round(g * 255))),
            _clamp255(int(round(b * 255)))
        ))
    return colors


def main():
    action = forms.CommandSwitchWindow.show(
        [
            'Global: Adjust colors (Saturation/Brightness/RGB)',
            'Reset scheme colors (from last backup)',
            'Reset to Contrasting Colors',
            'Edit: Change one entry (Color/Visible)',
            'Export scheme colors (CSV)',
            'Import scheme colors (CSV)',
            'Export scheme colors (.cschn)',
            'Import scheme colors (.cschn)',
        ],
        message='What do you want to do?'
    )
    if not action:
        return

    cat_name = _pick_category_name()
    if not cat_name:
        forms.alert('No scheme categories found in this model.', exitscript=True)

    scheme = _pick_scheme(category_filter=cat_name)
    if not scheme:
        return

    entries = _try_get_entries(scheme)
    if not entries:
        forms.alert('Selected scheme has no entries.', exitscript=True)

    if action.startswith('Reset scheme colors'):
        bpath = _backup_path_for_scheme(scheme)
        if not os.path.isfile(bpath):
            forms.alert('No backup found for this scheme.\n\nBackup path:\n{}'.format(bpath), exitscript=True)

        cschn_map = _import_scheme_cschn(bpath)
        if not cschn_map:
            forms.alert('Backup file exists but no valid entries were found.\n\nBackup path:\n{}'.format(bpath), exitscript=True)

        # Build lookup by the scheme "Value" label shown in Revit UI
        entry_by_label = {}
        for i, e in enumerate(entries):
            entry_by_label[_entry_label(e, idx=i).strip().lower()] = e

        changed = 0
        missing = 0
        with revit.Transaction('Reset Color Fill Scheme colors'):
            for k, rgb in cschn_map.items():
                try:
                    e = entry_by_label.get(k)
                    if e is None:
                        missing += 1
                        continue
                    r, g, b = rgb
                    c = DB.Color(_clamp255(r), _clamp255(g), _clamp255(b))
                    _set_entry_color(e, c)
                    changed += 1
                except Exception as ex:
                    logger.debug('Reset entry failed: %s', ex)
                    continue

            _commit_entries(scheme, entries)

        if missing:
            forms.alert('Reset {} entries from backup.\nMissing (not found in scheme): {}\n\nBackup:\n{}'.format(changed, missing, bpath))
        else:
            forms.alert('Reset {} entries from backup.\n\nBackup:\n{}'.format(changed, bpath))
        return

    if action.startswith('Reset to Contrasting Colors'):
        # Backup current colors first
        bpath = _backup_scheme_colors(scheme, entries)

        # Generate contrasting colors
        contrasting_colors = _generate_contrasting_colors(len(entries))

        changed = 0
        with revit.Transaction('Reset to Contrasting Colors'):
            for i, e in enumerate(entries):
                color = contrasting_colors[i % len(contrasting_colors)]  # Cycle if more entries than colors
                if _set_entry_color(e, color):
                    changed += 1
            _commit_entries(scheme, entries)

        if bpath:
            forms.alert('Reset {} entries to contrasting colors.\n\nBackup saved to:\n{}'.format(changed, bpath))
        else:
            forms.alert('Reset {} entries to contrasting colors.'.format(changed))
        return

    if action.startswith('Export'):
        if action.endswith('(.cschn)'):
            out_path = forms.save_file(file_ext='cschn', default_name='{}_{}.cschn'.format(_category_name_from_id(scheme.CategoryId), getattr(scheme, 'Name', 'Scheme')))
        else:
            out_path = forms.save_file(file_ext='csv', default_name='{}_{}.csv'.format(_category_name_from_id(scheme.CategoryId), getattr(scheme, 'Name', 'Scheme')))
        if not out_path:
            return
        if action.endswith('(.cschn)'):
            _export_scheme_cschn(out_path, scheme, entries)
        else:
            _export_scheme_csv(out_path, scheme, entries)
        forms.alert('Exported {} entries to:\n\n{}'.format(len(entries), out_path))
        return

    if action.startswith('Import'):
        if action.endswith('(.cschn)'):
            in_path = forms.pick_file(file_ext='cschn', title='Select .cschn to import')
        else:
            in_path = forms.pick_file(file_ext='csv', title='Select CSV to import')
        if not in_path:
            return

        rows = None
        cschn_map = None
        if action.endswith('(.cschn)'):
            cschn_map = _import_scheme_cschn(in_path)
            if not cschn_map:
                forms.alert('No valid entries found in .cschn.', exitscript=True)
        else:
            rows = _import_scheme_csv(in_path)
            if not rows:
                forms.alert('No rows found in CSV.', exitscript=True)

        # Build lookup by the scheme "Value" label shown in Revit UI
        entry_by_label = {}
        for i, e in enumerate(entries):
            entry_by_label[_entry_label(e, idx=i).strip().lower()] = e

        changed = 0
        missing = 0
        with revit.Transaction('Import Color Fill Scheme colors'):
            if cschn_map is not None:
                for k, rgb in cschn_map.items():
                    try:
                        e = entry_by_label.get(k)
                        if e is None:
                            missing += 1
                            continue
                        r, g, b = rgb
                        c = DB.Color(_clamp255(r), _clamp255(g), _clamp255(b))
                        _set_entry_color(e, c)
                        changed += 1
                    except Exception as ex:
                        logger.debug('CSCHN import entry failed: %s', ex)
                        continue
            else:
                for r in rows:
                    try:
                        entry_name = (r.get('Value') or r.get('Entry') or '').strip().lower()
                        if not entry_name:
                            continue
                        e = entry_by_label.get(entry_name)
                        if e is None:
                            missing += 1
                            continue

                        vis_raw = (r.get('Visible') or '').strip().lower()
                        if vis_raw:
                            is_vis = vis_raw in ['1', 'true', 'yes', 'y', 'on']
                            _set_entry_visible(e, is_vis)

                        c = DB.Color(_clamp255(r.get('R')), _clamp255(r.get('G')), _clamp255(r.get('B')))
                        if not _set_entry_color(e, c):
                            # will rely on SetEntries below
                            pass
                        changed += 1
                    except Exception as ex:
                        logger.debug('CSV import row failed: %s', ex)
                        continue

            # Ensure scheme receives the updated entries for APIs that require it
            _commit_entries(scheme, entries)

        if missing:
            forms.alert('Imported colors for {} matching entries.\nMissing (not found in scheme): {}'.format(changed, missing))
        else:
            forms.alert('Imported colors for {} matching entries.'.format(changed))
        return

    if action.startswith('Edit:'):
        class _EntryItem(object):
            def __init__(self, entry, idx):
                self.entry = entry
                self.idx = idx
                self.label = _entry_label(entry, idx=idx)
                c = _entry_color(entry)
                v = _entry_visible(entry)
                if c is not None:
                    self.display = '{}  [{}]  (RGB {},{}, {})'.format(self.label, 'Visible' if v else 'Hidden', c.Red, c.Green, c.Blue)
                else:
                    self.display = '{}  [{}]'.format(self.label, 'Visible' if v else 'Hidden')

        entry_items = [_EntryItem(e, i) for i, e in enumerate(entries)]
        picked = forms.SelectFromList.show(entry_items, name_attr='display', multiselect=False, title='Select Scheme Entry')
        if not picked:
            return

        old = _entry_color(picked.entry)
        default = '200,200,200'
        if old is not None:
            default = '{},{},{}'.format(old.Red, old.Green, old.Blue)

        newc = _ask_rgb(default=default, title='Set Entry Color', prompt='Enter RGB for "{}" as R,G,B:'.format(picked.label))
        if newc is None:
            return

        make_visible = forms.CommandSwitchWindow.show(
            ['Keep Visible setting', 'Set Visible', 'Set Hidden'],
            message='Visibility for this scheme entry:'
        )
        if not make_visible:
            return

        _backup_scheme_colors(scheme, entries)

        with revit.Transaction('Edit Color Fill Scheme entry'):
            _set_entry_color(picked.entry, newc)
            if make_visible == 'Set Visible':
                _set_entry_visible(picked.entry, True)
            elif make_visible == 'Set Hidden':
                _set_entry_visible(picked.entry, False)
            _commit_entries(scheme, entries)

        forms.alert('Updated entry color.')
        return

    # Global adjust (saturation/brightness/RGB)
    mode = forms.CommandSwitchWindow.show(
        [
            'Increase Saturation',
            'Decrease Saturation',
            'Increase Brightness',
            'Decrease Brightness',
            'RGB Offset (Add/Subtract)',
            'Cancel',
        ],
        message='Adjust all scheme colors. A backup is saved automatically so you can Reset later.'
    )
    if mode == 'Cancel' or not mode:
        return

    pct = None
    rgb_delta = None
    if mode == 'RGB Offset (Add/Subtract)':
        rgb_delta = _ask_rgb_delta()
        if rgb_delta is None:
            return
    else:
        pct = forms.ask_for_string(default='10', title='Global Adjust', prompt='Amount percent (0-100):')
        if pct is None:
            return

    bpath = _backup_scheme_colors(scheme, entries)

    changed = 0
    with revit.Transaction('Adjust Color Fill Scheme colors'):
        for e in entries:
            c = _entry_color(e)
            if c is None:
                continue

            if mode == 'Increase Saturation':
                nc = _adjust_saturation_color(c, pct, increase=True)
            elif mode == 'Decrease Saturation':
                nc = _adjust_saturation_color(c, pct, increase=False)
            elif mode == 'Increase Brightness':
                nc = _adjust_value_color(c, pct, increase=True)
            elif mode == 'Decrease Brightness':
                nc = _adjust_value_color(c, pct, increase=False)
            else:
                dr, dg, db = rgb_delta
                nc = _adjust_rgb_offsets(c, dr, dg, db)

            if _set_entry_color(e, nc):
                changed += 1
            else:
                # still count; commit below may apply
                changed += 1

        _commit_entries(scheme, entries)

    if bpath:
        forms.alert('Adjusted {} entry colors.\n\nBackup saved to:\n{}'.format(changed, bpath))
    else:
        forms.alert('Adjusted {} entry colors.'.format(changed))


if __name__ == '__main__':
    main()
