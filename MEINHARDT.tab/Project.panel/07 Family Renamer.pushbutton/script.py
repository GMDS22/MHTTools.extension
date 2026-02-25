# -*- coding: utf-8 -*-
"""
MHT Family Namer - pyRevit tool

Scans loaded families and suggests standardized names based on templates in "naming_rules.json".

This script is intentionally minimal and prints results to the Revit Python Console. It can be
extended to provide a WinForms UI for review/apply.
"""
from __future__ import print_function


import os
import json

from pyrevit import revit, script
from Autodesk.Revit import DB

# Windows Forms / .NET imports are available when running inside Revit / IronPython.
# Wrap imports in try/except to avoid static analyzer errors in editors.
try:
    from System import Array
    from System.Windows.Forms import Application, Form, DataGridView, DataGridViewCheckBoxColumn, DataGridViewTextBoxColumn, Button, DockStyle, DialogResult, FormStartPosition, ComboBox, Label, TextBox, MessageBox, MessageBoxButtons, MessageBoxIcon, FlatStyle, DataGridViewAutoSizeColumnsMode, DataGridViewRowHeadersWidthSizeMode, AutoScaleMode, FormWindowState, FlowDirection
    from System.Drawing import Size, Color, SystemColors, SizeF
except Exception:
    Application = None
    Form = object
    DataGridView = None
    DataGridViewCheckBoxColumn = None
    DataGridViewTextBoxColumn = None
    Button = None
    DockStyle = None
    DialogResult = None
    FormStartPosition = None
    Size = None
    ComboBox = None
    Label = None
    Color = None
    SizeF = None
    DataGridViewAutoSizeColumnsMode = None
    DataGridViewRowHeadersWidthSizeMode = None
    AutoScaleMode = None
    FormWindowState = None
    FlowDirection = None

logger = script.get_logger()

# UI behavior:
# - The main review UI stays.
# - Warning/info/question popups are suppressed by default.
#   Only errors should interrupt users.
SHOW_NON_ERROR_DIALOGS = False

SCRIPT_DIR = os.path.dirname(__file__)
RULES_FILE = os.path.join(SCRIPT_DIR, 'naming_rules.json')

__author__ = 'gmoreno'
__description__ = 'Scan loaded families, suggest standardized names and optionally rename families and types safely.'


def _show_error_dialog(message, title='Error'):
    try:
        if MessageBox is not None:
            MessageBox.Show(str(message), str(title), MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
    except Exception:
        pass
    try:
        logger.error('{}: {}'.format(title, message))
    except Exception:
        pass


def _maybe_dialog(message, title='Info', buttons=None, icon=None, default_result=None):
    """Show a MessageBox only when SHOW_NON_ERROR_DIALOGS is True.

    Returns MessageBox.Show result, or default_result when suppressed.
    """
    if not SHOW_NON_ERROR_DIALOGS:
        return default_result
    try:
        if MessageBox is None:
            return default_result
        b = buttons if buttons is not None else MessageBoxButtons.OK
        i = icon if icon is not None else MessageBoxIcon.Information
        return MessageBox.Show(str(message), str(title), b, i)
    except Exception:
        return default_result


def load_rules(path):
    # Try the provided path first, then a set of sensible fallbacks so tests
    # and different workspace layouts can find the rules file.
    candidates = []
    try:
        if path:
            candidates.append(path)
    except Exception:
        pass
    try:
        # try current working directory (tests may run from repo root)
        candidates.append(os.path.join(os.getcwd(), 'naming_rules.json'))
    except Exception:
        pass
    try:
        # primary rules location next to this script
        candidates.append(os.path.join(SCRIPT_DIR, 'naming_rules.json'))
    except Exception:
        pass
    try:
        # also try inside the package folder relative to SCRIPT_DIR
        candidates.append(os.path.join(SCRIPT_DIR, 'Family_Rnmr.pushbutton', 'naming_rules.json'))
    except Exception:
        pass
    try:
        # parent folder (e.g., gm-tools.panel/naming_rules.json)
        candidates.append(os.path.join(SCRIPT_DIR, '..', 'naming_rules.json'))
    except Exception:
        pass

    for p in candidates:
        try:
            if not p:
                continue
            p_norm = os.path.normpath(p)
            if os.path.exists(p_norm):
                try:
                    with open(p_norm, 'r') as f:
                        return json.load(f)
                except Exception as e:
                    logger.warning('Failed to load naming rules from {}: {}'.format(p_norm, e))
        except Exception:
            continue
    return {}


def get_families(doc):
    """Return a list of Family objects from the given Revit document.

    When running inside Revit this uses a FilteredElementCollector. When not
    available (unit tests / offline), return an empty list so higher-level
    logic can operate safely.
    """
    try:
        from Autodesk.Revit.DB import FilteredElementCollector, Family
        col = FilteredElementCollector(doc).OfClass(Family)
        try:
            return list(col.ToElements())
        except Exception:
            # some IronPython/Revit versions expose IEnumerable directly
            return [f for f in col]
    except Exception:
        # Not running inside Revit or collector failed; return empty list
        return []


def _extract_param_map(params_iter):
    """Extract a simple param-name -> string map from a Revit Parameters iterator."""
    out = {}
    try:
        for p in params_iter:
            try:
                name = p.Definition.Name
            except Exception:
                continue
            try:
                if p.StorageType == DB.StorageType.String:
                    out[name] = p.AsString() or ''
                elif p.StorageType == DB.StorageType.Double:
                    val = p.AsDouble()
                    try:
                        val_mm = DB.UnitUtils.ConvertFromInternalUnits(val, DB.UnitTypeId.Millimeters)
                    except Exception:
                        try:
                            val_mm = val * 304.8
                        except Exception:
                            val_mm = val
                    try:
                        out[name] = str(int(round(val_mm)))
                    except Exception:
                        out[name] = str(val_mm)
                else:
                    out[name] = ''
            except Exception:
                out[name] = ''
    except Exception:
        pass
    return out


def _build_first_instance_params_cache(doc, families):
    """Build {familyIdInt: param_map} using at most one instance per family.

    Avoids the previous O(Families * Instances) behavior.
    """
    cache = {}
    try:
        if not families:
            return cache
        remaining = set()
        for fam in families:
            try:
                remaining.add(fam.Id.IntegerValue)
            except Exception:
                pass
        if not remaining:
            return cache

        collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance)
        for inst in collector:
            try:
                fid = inst.Symbol.Family.Id.IntegerValue
            except Exception:
                continue
            if fid not in remaining:
                continue
            try:
                cache[fid] = _extract_param_map(inst.Parameters)
            except Exception:
                cache[fid] = {}
            try:
                remaining.remove(fid)
            except Exception:
                pass
            if not remaining:
                break
    except Exception as e:
        logger.warning('Failed building instance param cache: {}'.format(e))
    return cache


def _write_text_file(path, text, encoding='utf-8'):
    """Write text to file in a way that works on CPython and IronPython.

    Tries codecs.open first (works with encoding in IronPython), then built-in open
    without encoding. Returns True on success, False on failure.
    """
    try:
        # Try codecs.open first - this works better with IronPython
        try:
            import codecs
            with codecs.open(path, 'w', encoding=encoding) as f:
                f.write(text)
            return True
        except Exception:
            pass
        # Fallback: built-in open without encoding (IronPython default)
        try:
            with open(path, 'w') as f:
                f.write(text)
            return True
        except Exception:
            return False
    except Exception:
        return False
def _extract_first_number_with_units(txt):
    """Extract the first numeric size from text and convert to millimeters.

    Returns an integer millimeter value or None.
    """
    if not txt:
        return None
    try:
        import re
        s = str(txt)
        # find numeric tokens (with optional decimals)
        nums = re.findall(r"(\d+(?:\.\d+)?)", s)
        if not nums:
            return None

        candidates = []
        low = s.lower()
        for token in nums:
            try:
                num = float(token)
            except Exception:
                continue
            mm = None
            # unit detection (prefer explicit units found in the surrounding text)
            if 'mm' in low:
                mm = int(round(num))
            elif 'cm' in low:
                mm = int(round(num * 10))
            elif 'dn' in low:
                mm = int(round(num))
            elif '"' in low or ' in' in low or low.endswith('in'):
                mm = int(round(num * 25.4))
            elif 'm' in low and num < 10:
                mm = int(round(num * 1000))
            else:
                mm = int(round(num))

            # clamp to plausible engineering range
            try:
                if mm is not None and 5 <= mm <= 5000:
                    return mm
                if mm is not None:
                    candidates.append(mm)
            except Exception:
                continue

        return candidates[0] if candidates else None
    except Exception:
        return None


def _scan_for_token_in_text(tokens, text):
    """Return first token (as provided in tokens) that is found in text (case-insensitive), or None."""
    if not tokens or not text:
        return None
    try:
        import re
        t = text or ''
        # First try word-boundary matches to catch abbreviations like 'EA', 'RA', 'SA' reliably
        for tok in tokens:
            try:
                if not tok:
                    continue
                pattern = r"\b" + re.escape(tok) + r"\b"
                if re.search(pattern, t, flags=re.I):
                    return tok
            except Exception:
                continue
        # Fallback: substring match (looser)
        t_lower = t.lower()
        for tok in tokens:
            try:
                if not tok:
                    continue
                if tok.lower() in t_lower:
                    return tok
            except Exception:
                continue
    except Exception:
        pass
    return None

# ---------------------------------------------------------------------------------


def _write_export_files(csv_text, results, rules, script_dir, open_file=False):
    """Centralized CSV + debug JSON writer. Returns (csv_path, dbg_path or None).

    Uses `_write_text_file` for cross-runtime compatibility. If `open_file` is True
    attempts to open the CSV after writing (Windows: os.startfile, otherwise webbrowser).
    """
    try:
        from datetime import datetime
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fname = 'family_name_suggestions_{}.csv'.format(ts)
        csv_path = os.path.join(script_dir or SCRIPT_DIR, fname)
        ok = _write_text_file(csv_path, csv_text, encoding='utf-8')
        if not ok:
            raise Exception('Failed to write CSV via helper')

        # build debug JSON for problematic rows
        debug_rows = []
        try:
            for r in results:
                info = r.get('info', {}) or {}
                is_problem = False
                reasons = []
                if not r.get('suggested'):
                    is_problem = True
                    reasons.append('no_suggestion')
                try:
                    cat = (r.get('category') or '').lower()
                except Exception:
                    cat = ''
                if 'air' in cat or 'terminal' in cat or 'air terminal' in cat:
                    sysv = info.get('_classified_system')
                    sizev = info.get('_classified_size')
                    shapev = info.get('_classified_shape')
                    if not sysv or (isinstance(sysv, str) and (sysv.upper() == 'UND' or sysv.strip() == '')):
                        is_problem = True
                        reasons.append('missing_system')
                    if not sizev:
                        is_problem = True
                        reasons.append('missing_size')
                    if not shapev:
                        reasons.append('missing_shape')
                if 'duct' in cat:
                    s = (r.get('suggested') or '').strip()
                    if not s or s.upper() == 'MHT-ME-DF' or '---' in s:
                        is_problem = True
                        reasons.append('incomplete_df_template')
                    if not info.get('_classified_shape'):
                        reasons.append('missing_shape')
                    if not info.get('_classified_fitting'):
                        reasons.append('missing_fitting')
                    if not info.get('_classified_size'):
                        reasons.append('missing_size')
                if is_problem:
                    debug_rows.append({
                        'family': r.get('family'),
                        'category': r.get('category'),
                        'suggested': r.get('suggested'),
                        'reasons': reasons,
                        'classified_system': info.get('_classified_system'),
                        'classified_size': info.get('_classified_size'),
                        'classified_shape': info.get('_classified_shape'),
                        'classified_fitting': info.get('_classified_fitting'),
                        'classified_modifier': info.get('_classified_modifier'),
                        'type_params_first': (info.get('types') or [])[:1],
                        'instance_params': info.get('instance_params', {})
                    })
        except Exception:
            debug_rows = []

        dbg_path = None
        if debug_rows:
            try:
                dbg_fname = 'family_name_suggestions_debug_{}.json'.format(ts)
                dbg_path = os.path.join(script_dir or SCRIPT_DIR, dbg_fname)
                dbg_text = json.dumps({'generated': ts, 'problems': debug_rows}, indent=2, ensure_ascii=False)
                _write_text_file(dbg_path, dbg_text, encoding='utf-8')
            except Exception:
                dbg_path = None

        # Optionally open file for user convenience
        try:
            if open_file:
                if os.name == 'nt':
                    try:
                        os.startfile(csv_path)
                    except Exception:
                        import webbrowser
                        webbrowser.open('file://' + os.path.realpath(csv_path))
                else:
                    import webbrowser
                    webbrowser.open('file://' + os.path.realpath(csv_path))
        except Exception:
            pass

        return csv_path, dbg_path
    except Exception as e:
        logger.warning('Export failed: {}'.format(e))
        return None, None


def export_results_now(rows, rules=None, script_dir=None, open_file=False):
    """Build CSV from `rows` (list of dicts) and write CSV + debug JSON using helper.

    Returns tuple (csv_path, dbg_path) or (None, None) on failure.
    """
    try:
        if rules is None:
            try:
                rules = load_rules(RULES_FILE)
            except Exception:
                rules = {}
        # Clone rows so we don't mutate caller data
        results = [dict(r) for r in (rows or [])]

        # compute counts for suggested names and attempt to disambiguate by size
        counts = {}
        for r in results:
            try:
                s = (r.get('suggested') or '').strip()
                counts[s] = counts.get(s, 0) + 1
            except Exception:
                pass

        # try to append size when duplicates exist
        try:
            for name, cnt in list(counts.items()):
                if cnt > 1:
                    dup_results = [r for r in results if (r.get('suggested') or '') == name]
                    for r in dup_results:
                        try:
                            size = r.get('info', {}).get('_classified_size')
                            if size:
                                candidate = "%s-%s" % (name, size)
                                if candidate.lower() not in [x.lower() for x in counts.keys()] and candidate.lower() not in [rr.get('suggested','').lower() for rr in results]:
                                    r['suggested'] = candidate
                                    counts[candidate] = counts.get(candidate, 0) + 1
                                    counts[name] = counts.get(name, 1) - 1
                        except Exception:
                            pass
        except Exception:
            pass

        # auto-resolve duplicates by appending -01, -02, etc.
        used = {}
        for r in results:
            try:
                s = r.get('suggested') or ''
                if not s:
                    r['conflict'] = False
                    continue
                if counts.get(s, 0) > 1:
                    idx = used.get(s, 0) + 1
                    used[s] = idx
                    new_name = "%s-%02d" % (s, idx)
                    r['suggested'] = new_name
                    r['conflict'] = True
                else:
                    r['conflict'] = False
            except Exception:
                r['conflict'] = False

        # build CSV text
        out = ['Family,Category,SuggestedName,Conflict,ProblemReasons']
        for r in results:
            try:
                cur = r.get('family') or ''
                cat = r.get('category') or ''
                sug = r.get('suggested') or ''
                conf = 'DUP' if r.get('conflict') else ''
                reasons = []
                info = r.get('info', {}) or {}
                if not sug:
                    reasons.append('no_suggestion')
                cat_l = (cat or '').lower()
                if 'air' in cat_l or 'terminal' in cat_l:
                    sysv = info.get('_classified_system')
                    sizev = info.get('_classified_size')
                    shapev = info.get('_classified_shape')
                    if not sysv or (isinstance(sysv, str) and (sysv.upper() == 'UND' or sysv.strip() == '')):
                        reasons.append('missing_system')
                    if not sizev:
                        reasons.append('missing_size')
                    if not shapev:
                        reasons.append('missing_shape')
                if 'duct' in cat_l:
                    s = (sug or '').strip()
                    if not s or s.upper() == 'MHT-ME-DF' or '---' in s:
                        reasons.append('incomplete_df_template')
                    if not info.get('_classified_shape'):
                        reasons.append('missing_shape')
                    if not info.get('_classified_fitting'):
                        reasons.append('missing_fitting')
                    if not info.get('_classified_size'):
                        reasons.append('missing_size')
                pr = ';'.join(reasons)
                out.append('"{}","{}","{}","{}","{}"'.format(cur, cat, sug, conf, pr))
            except Exception:
                continue

        csv_text = '\n'.join(out)

        return _write_export_files(csv_text, results, rules, script_dir or SCRIPT_DIR, open_file=open_file)
    except Exception as e:
        logger.warning('export_results_now failed: {}'.format(e))
        return None, None


def remove_existing_prefixes(name, rules):
    """Strip known leading company/discipline prefixes from a name.

    Uses rules['COMPANY'] and keys from rules['DISCIPLINE'] when available.
    This prevents duplicating the company prefix (e.g. MHT-) when applying templates.
    """
    try:
        if not name:
            return name
        import re
        out = name.strip()
        # drop leading exclamation markers often used in families (e.g., !MHT_...)
        out = re.sub(r'^[!]+', '', out)

        # build candidate prefixes: company + discipline keys
        prefixes = []
        try:
            comp = (rules.get('COMPANY') or '').strip()
            if comp:
                prefixes.append(comp)
        except Exception:
            pass
        try:
            disc = rules.get('DISCIPLINE', {}) or {}
            for k in disc.keys():
                if k and k not in prefixes:
                    prefixes.append(k)
        except Exception:
            pass

        # Iteratively strip any of these prefixes when they appear at start followed by a delimiter
        changed = True
        while changed:
            changed = False
            for p in prefixes:
                if not p:
                    continue
                # match prefix at start with optional separators (dash, underscore, space)
                low = out.lower()
                p_low = p.lower()
                if low.startswith(p_low):
                    # remove the prefix and any immediately following separators
                    rem = out[len(p):]
                    rem = re.sub(r'^[\s\-_]+', '', rem)
                    out = rem
                    changed = True
                    break
        # After iterative stripping, also strip a short vendor-style prefix (1-3 letters)
        # followed by underscore/dash/space (e.g. "M_Round", "TG-Name"). Do not strip
        # if it exactly matches the company code or a discipline key.
        try:
            m = re.match(r'^([A-Za-z]{1,3})[_\-\s]+(.*)', out)
            if m:
                token = m.group(1)
                skip = False
                try:
                    comp = (rules.get('COMPANY') or '').strip()
                    if comp and token.lower() == comp.lower():
                        skip = True
                except Exception:
                    pass
                try:
                    disc = rules.get('DISCIPLINE', {}) or {}
                    for k in disc.keys():
                        if k and token.lower() == k.lower():
                            skip = True
                            break
                except Exception:
                    pass
                if not skip:
                    out = m.group(2)
        except Exception:
            pass
    except Exception:
        return name
    return out


def safe_get_name(obj):
    """Safe name extraction for Revit objects or plain strings used in tests."""
    try:
        if obj is None:
            return ''
        if isinstance(obj, str):
            return obj
        # Revit FamilySymbol and Family objects expose a Name property
        name = getattr(obj, 'Name', None)
        if name:
            return str(name)
        # fallback: avoid returning raw object reprs from wrapped Revit elements
        s = str(obj)
        # If it looks like a Revit wrapper repr, return empty string to avoid leaking memory addresses
        if 'autodesk.revit' in s.lower() or s.startswith('<Autodesk.Revit') or 'object at 0x' in s.lower():
            return ''
        return s
    except Exception:
        return ''


def clean_name(s):
    """Normalize a name to a cleaned token form: remove common words, non-alphanumerics -> dashes, TitleCase parts."""
    try:
        if not s:
            return ''
        import re
        out = str(s).strip()
        # remove common noise words like 'type'
        out = re.sub(r'\btype\b', '', out, flags=re.I)
        # replace non-alphanumeric with dash
        out = re.sub(r'[^A-Za-z0-9]+', '-', out)
        # collapse multiple dashes
        out = re.sub(r'-{2,}', '-', out)
        # strip leading/trailing dashes
        out = out.strip('-')
        # Title case each hyphen-separated component but keep all-upper for known acronyms (simple heuristic)
        parts = [p for p in out.split('-') if p]
        def fix_part(p):
            if p.isupper() or p.isdigit():
                return p
            return p.title()
        parts = [fix_part(p) for p in parts]
        out = '-'.join(parts)
        return out
        out = re.sub(r'-{2,}', '-', out)
        out = out.strip('-')
        parts = [p.title() for p in out.split('-') if p]
        return '-'.join(parts)
    except Exception:
        try:
            return str(s).strip()
        except Exception:
            return ''
            for p in prefixes:
                try:
                    if not p:
                        continue
                    # match prefix at start, case-insensitive, optionally followed by delimiters
                    m = re.match(r'^(%s)[\-\_\.: ]*' % re.escape(p), out, flags=re.I)
                    if m:
                        out = out[m.end():]
                        changed = True
                        break
                except Exception:
                    continue

            # If nothing matched from rules, also strip common vendor/discipline-like prefixes
            # e.g., WSP_DI-, ABC-XY_, TG_DI., DETItem_ etc. These are usually organization tokens
            # and should not be preserved. Pattern: 2-5 uppercase letters optionally followed by
            # a token like _DI or -DI, then a delimiter.
            if not changed:
                try:
                    m2 = re.match(r'^([A-Z]{2,5}(?:[\-_][A-Z]{1,4})?)[\-\_\.: ]+', out)
                    if m2:
                        # don't strip if this token equals the company's code or known disciplines
                        token = m2.group(1)
                        skip = False
                        try:
                            comp = (rules.get('COMPANY') or '').strip()
                            if comp and token.lower() == comp.lower():
                                skip = True
                        except Exception:
                            pass
                        try:
                            disc = rules.get('DISCIPLINE', {}) or {}
                            for k in disc.keys():
                                if k and token.lower() == k.lower():
                                    skip = True
                                    break
                        except Exception:
                            pass
                        if not skip:
                            out = out[m2.end():]
                            changed = True
                except Exception:
                    pass

        # remove leftover duplicate separators and return
        out = out.replace('--', '-').replace('__', '_')
        return out.strip(' -_.')
    except Exception:
        return name


def _collect_existing_family_names(doc):
    names = set()
    try:
        for f in DB.FilteredElementCollector(doc).OfClass(DB.Family):
            try:
                n = safe_get_name(f)
                if n:
                    names.add(n.strip().lower())
            except Exception:
                pass
    except Exception:
        pass
    return names


def _collect_existing_type_names(doc, fam_name):
    names = set()
    try:
        for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
            try:
                fam = safe_get_name(s.Family) or ''
                if fam.strip().lower() == (fam_name or '').strip().lower():
                    n = safe_get_name(s)
                    if n:
                        names.add(n.strip().lower())
            except Exception:
                pass
    except Exception:
        pass
    return names


def _make_unique(base, existing, used=None):
    try:
        if used is None:
            used = set()
        candidate = base
        idx = 1
        while candidate.strip().lower() in existing or candidate.strip().lower() in used:
            candidate = "%s-%02d" % (base, idx)
            idx += 1
        used.add(candidate.strip().lower())
        return candidate
    except Exception:
        return base


# Category to short-code mapping (used by templates via <CAT>)
CATEGORY_CODE = {
    "Pipe Fittings": "PF",
    "Duct Fittings": "DF",
    "Mechanical Equipment": "EQ",
    "Generic Models": "GM",
    "Generic Annotations": "FAM",
    "Lighting Fixtures": "LFX",
    "Electrical Fixtures": "EFX",
    "Plumbing Fixtures": "PLF",
}


def family_primary_category_name(family):
    try:
        cat = family.FamilyCategory
        return cat.Name if cat is not None else 'Unknown'
    except Exception:
        return 'Unknown'


def gather_family_info(family, instance_params_cache=None):
    info = {}
    info['id'] = family.Id.IntegerValue
    # Resolve family name safely across Revit versions / API wrappers
    fam_name = None
    try:
        # Direct property on Family object (some Revit API wrappers expose this)
        if hasattr(family, 'Name'):
            try:
                fam_name = family.Name
            except Exception:
                fam_name = None

        # Some hosts expose FamilyName on symbols
        if not fam_name:
            try:
                symbol_ids = list(family.GetFamilySymbolIds())
                if symbol_ids:
                    sym = family.Document.GetElement(symbol_ids[0])
                    if sym is not None:
                        # try common attributes in order
                        for attr in ('FamilyName', 'Name'):
                            if hasattr(sym, attr):
                                try:
                                    fam_name = getattr(sym, attr)
                                    break
                                except Exception:
                                    fam_name = None
            except Exception:
                fam_name = None

        # fallback: try built-in parameter for family name
        if not fam_name:
            try:
                param = family.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if param:
                    fam_name = param.AsString()
            except Exception:
                fam_name = None

    except Exception:
        fam_name = None

    info['family_name'] = fam_name if fam_name else 'Unknown Family'
    info['category'] = family_primary_category_name(family)
    # collect type names and parameters from the first symbol
    symbols = list(family.GetFamilySymbolIds())
    types = []
    for sid in symbols:
        sym = family.Document.GetElement(sid)
        t = {'type_name': safe_get_name(sym) or ''}
        # collect simple parameters available on the type
        try:
            t['params'] = _extract_param_map(sym.Parameters)
        except Exception:
            t['params'] = {}
        types.append(t)
    info['types'] = types
    # gather instance-level parameters for fallback
    inst_params = {}
    try:
        fid_int = None
        try:
            fid_int = family.Id.IntegerValue
        except Exception:
            fid_int = None

        if instance_params_cache is not None and fid_int is not None:
            try:
                inst_params = instance_params_cache.get(fid_int, {}) or {}
            except Exception:
                inst_params = {}
        else:
            # Fallback (slower): scan for the first placed instance
            doc = family.Document
            collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance)
            for inst in collector:
                try:
                    if inst.Symbol.Family.Id == family.Id:
                        try:
                            inst_params = _extract_param_map(inst.Parameters)
                        except Exception:
                            inst_params = {}
                        break
                except Exception:
                    continue
    except Exception:
        inst_params = {}
    info['instance_params'] = inst_params
    return info


def classify_family(info, rules, used_generic_names=None):
    """
    Classify a family to find a generic name based on a set of rules.
    Returns the generic name (e.g., "Pump", "AHU") or None if no match.
    """
    classification_rules = rules.get("CLASSIFICATION_RULES", [])
    if not classification_rules:
        return None

    if used_generic_names is None:
        used_generic_names = set()

    fam_cat = (info.get('category') or '').lower()
    fam_name = (info.get('family_name') or '').lower()
    
    # Combine all parameter names from types and instance for searching
    all_param_names = set(info.get('instance_params', {}).keys())
    for t in info.get('types', []):
        all_param_names.update(t.get('params', {}).keys())

    all_param_values = {}
    for t in info.get('types', []):
        for p_name, p_val in t.get('params', {}).items():
            all_param_values[p_name.lower()] = (p_val or '').lower()
    # include instance-level parameter values as well (useful for system/type detection)
    for p_name, p_val in info.get('instance_params', {}).items():
        try:
            all_param_values[p_name.lower()] = (p_val or '').lower()
        except Exception:
            pass
    # build a combined searchable text from parameter values, parameter names, type names and family name
    try:
        type_names_text = ' '.join([(t.get('type_name') or '').lower() for t in info.get('types', []) if t.get('type_name')])
    except Exception:
        type_names_text = ''
    try:
        param_names_text = ' '.join([p.lower() for p in all_param_names if p])
    except Exception:
        param_names_text = ''
    # Build a combined searchable text containing parameter values, param names,
    # family name, type names and several variants of the family name so that
    # classification rules can match tokens that appear only in the current
    # family name (including names with prefixes, hyphens or underscores).
    fam_variants = [fam_name or '']
    try:
        # include a version with existing company/discipline prefixes removed
        stripped = remove_existing_prefixes(info.get('family_name') or '', rules) if rules else info.get('family_name') or ''
        fam_variants.append((stripped or '').lower())
    except Exception:
        pass
    try:
        # include a cleaned variant (safe characters, dashes normalized)
        fam_variants.append(clean_name(info.get('family_name') or '').lower())
    except Exception:
        pass
    # also include simple whitespace-separated variants
    try:
        raw = (info.get('family_name') or '')
        fam_variants.append(raw.replace('-', ' ').lower())
        fam_variants.append(raw.replace('_', ' ').lower())
    except Exception:
        pass

    combined_text = ' '.join([v for v in all_param_values.values() if v]) + ' ' + param_names_text + ' ' + ' '.join([v for v in fam_variants if v]) + ' ' + type_names_text
    
    # Note: Getting system classification requires a placed instance, which is complex.
    # For now, we will rely on category, parameter names, and family name keywords.

    for rule in classification_rules:
        generic_name = rule.get("generic_name")
        clues = rule.get("clues", {})
        if not generic_name or not clues:
            continue

        # 1. Check Category
        cat_clues = [c.lower() for c in clues.get("category", [])]
        if cat_clues and fam_cat not in cat_clues:
            continue # This rule doesn't apply to this category

        # 2. Check Parameter Names
        param_clues = [p.lower() for p in clues.get("param_names", [])]
        if param_clues:
            found_param = False
            for p_clue in param_clues:
                for existing_param in all_param_names:
                    if p_clue in existing_param.lower():
                        found_param = True
                        break
                if found_param:
                    break
            if found_param:
                # Pre-scan parameter values, type names, and family name for system, size, and subtype tokens
                sys_param_clues = clues.get("system_type_param_keywords", [])
                sys_val_clues = clues.get("system_type_values", [])
                system_type = None
                if sys_val_clues:
                    system_type = _scan_for_token_in_text(sys_val_clues, combined_text)
                if not system_type and sys_param_clues and sys_val_clues:
                    for p_clue in sys_param_clues:
                        p_val = all_param_values.get(p_clue.lower())
                        if p_val:
                            system_type = _scan_for_token_in_text(sys_val_clues, p_val)
                            if system_type:
                                break

                # size detection
                size = None
                size_param_clues = clues.get("size_param_keywords", [])
                if size_param_clues:
                    for p_clue in size_param_clues:
                        p_val = all_param_values.get(p_clue.lower())
                        if p_val:
                            mm = _extract_first_number_with_units(p_val)
                            if mm:
                                size = 'DN{}'.format(mm)
                                break
                if not size:
                    mm = _extract_first_number_with_units(combined_text)
                    if mm:
                        size = 'DN{}'.format(mm)

                # subtype detection (shape, fitting, modifier)
                shape = None
                fitting = None
                modifier = None
                for s_k in clues.get('sub_type_keywords', []):
                    try:
                        if s_k and s_k.lower() in combined_text:
                            shape = s_k.title()
                            break
                    except Exception:
                        continue
                # fitting keywords may be in family_name_keywords
                for f_k in clues.get('family_name_keywords', []):
                    try:
                        if f_k and f_k.lower() in combined_text:
                            fitting = f_k.title()
                            break
                    except Exception:
                        continue
                for m_k in clues.get('fitting_modifier_keywords', []):
                    try:
                        if m_k and m_k.lower() in combined_text:
                            mk = m_k.lower()
                            if 'long' in mk and 'short' not in mk:
                                modifier = 'LongRad'
                            elif 'short' in mk and 'long' not in mk:
                                modifier = 'ShortRad'
                            else:
                                modifier = m_k.title().replace(' ', '')
                            break
                    except Exception:
                        continue

                info['_classified_system'] = system_type or 'UND'
                if size:
                    info['_classified_size'] = size
                if shape:
                    info['_classified_shape'] = shape
                if fitting:
                    info['_classified_fitting'] = fitting
                if modifier:
                    info['_classified_modifier'] = modifier

                # Debug: if this rule appears to be for Air Terminals but key tokens weren't found,
                # emit a detailed log to help expand rules or find parameters containing tokens.
                try:
                    is_air_rule = False
                    try:
                        # check generic name or category clues
                        if generic_name and 'air' in (generic_name or '').lower():
                            is_air_rule = True
                    except Exception:
                        pass
                    try:
                        if not is_air_rule:
                            cat_clues = clues.get('category', []) or []
                            for cc in cat_clues:
                                try:
                                    if 'air term' in (cc or '').lower() or 'air terminal' in (cc or '').lower():
                                        is_air_rule = True
                                        break
                                except Exception:
                                    continue
                    except Exception:
                        pass

                    if is_air_rule:
                        missing = []
                        if not system_type:
                            missing.append('system')
                        if not size:
                            missing.append('size')
                        if not shape:
                            missing.append('shape')
                        if missing:
                            try:
                                log = getattr(script, 'get_logger')()
                            except Exception:
                                log = None
                            dbg = {
                                'family': info.get('family_name'),
                                'category': info.get('category'),
                                'missing': missing,
                                'combined_text_sample': combined_text[:300],
                                'instance_params': info.get('instance_params', {}),
                                'type_params_first': (info.get('types') or [])[:1]
                            }
                            try:
                                if log:
                                    log.warning('Air Terminal classification incomplete: {}'.format(dbg))
                                else:
                                    print('Air Terminal classification incomplete:', dbg)
                            except Exception:
                                try:
                                    print('Air Terminal debug:', dbg)
                                except Exception:
                                    pass
                except Exception:
                    pass

                try:
                    specific = build_classified_name(generic_name, fam_name, clues, all_param_values, used_generic_names)
                    return specific
                except Exception:
                    return generic_name

        # 3. Check Family Name Keywords (as a fallback)
        keyword_clues = [k.lower() for k in clues.get("family_name_keywords", [])]
        if keyword_clues:
            for keyword in keyword_clues:
                if keyword in fam_name:
                    try:
                        # Pre-scan combined text for system and size, then build specific name
                        try:
                            sys_param_clues = clues.get("system_type_param_keywords", [])
                            sys_val_clues = clues.get("system_type_values", [])
                            system_type = None
                            if sys_val_clues:
                                system_type = _scan_for_token_in_text(sys_val_clues, combined_text)
                            if not system_type and sys_param_clues and sys_val_clues:
                                for p_clue in sys_param_clues:
                                    p_val = all_param_values.get(p_clue.lower())
                                    if p_val:
                                        for v_clue in sys_val_clues:
                                            if v_clue.lower() in p_val:
                                                system_type = v_clue
                                                break
                                    if system_type:
                                        break

                            size = None
                            size_param_clues = clues.get("size_param_keywords", [])
                            if size_param_clues:
                                for p_clue in size_param_clues:
                                    p_val = all_param_values.get(p_clue.lower())
                                    if p_val:
                                        mm = _extract_first_number_with_units(p_val)
                                        if mm:
                                            size = 'DN{}'.format(mm)
                                            break
                            if not size:
                                mm = _extract_first_number_with_units(combined_text)
                                if mm:
                                    size = 'DN{}'.format(mm)

                            info['_classified_system'] = system_type or 'UND'
                            if size:
                                info['_classified_size'] = size
                        except Exception:
                            info['_classified_system'] = 'UND'
                        specific = build_classified_name(generic_name, fam_name, clues, all_param_values, used_generic_names)
                        return specific
                    except Exception:
                        return generic_name

    return None


def build_classified_name(generic_name, fam_name, clues, all_param_values, used_names):
    """Build a classified name, trying for more specificity if conflicts exist."""
    # Level 1: Just the generic name
    try:
        if generic_name and generic_name.lower() not in used_names:
            return generic_name
    except Exception:
        pass

    # Level 2: Generic name + System Type (scan parameter values, type names and family name)
    system_type = None
    sys_param_clues = clues.get("system_type_param_keywords", [])
    sys_val_clues = clues.get("system_type_values", [])
    try:
        # combined searchable text
        combined_text = ' '.join([v for v in all_param_values.values() if v]) + ' ' + (fam_name or '')
    except Exception:
        combined_text = (fam_name or '')
    if sys_val_clues:
        system_type = _scan_for_token_in_text(sys_val_clues, combined_text)
    # fallback to parameter-name based checks
    if not system_type and sys_param_clues and sys_val_clues:
        for p_clue in sys_param_clues:
            p_val = all_param_values.get(p_clue.lower())
            if p_val:
                for v_clue in sys_val_clues:
                    try:
                        if v_clue.lower() in p_val:
                            system_type = v_clue
                            break
                    except Exception:
                        continue
            if system_type:
                break
    try:
        if system_type:
            name_with_system = "{}-{}".format(generic_name, system_type)
            if name_with_system.lower() not in used_names:
                return name_with_system
    except Exception:
        pass

    # Level 3: Generic name + Size (if available) - prefer size before UND or subtype
    size = None
    size_param_clues = clues.get("size_param_keywords", [])
    if size_param_clues:
        for p_clue in size_param_clues:
            p_val = all_param_values.get(p_clue.lower())
            if p_val:
                mm = _extract_first_number_with_units(p_val)
                if mm:
                    size = 'DN{}'.format(mm)
                    break
    # fallback: scan any param values and family/type names
    if not size:
        try:
            for v in all_param_values.values():
                mm = _extract_first_number_with_units(v)
                if mm:
                    size = 'DN{}'.format(mm)
                    break
            if not size:
                mm = _extract_first_number_with_units(fam_name)
                if mm:
                    size = 'DN{}'.format(mm)
        except Exception:
            pass

    try:
        if size:
            name_with_size = "{}-{}".format(generic_name, size)
            if name_with_size.lower() not in used_names:
                return name_with_size
    except Exception:
        pass

    # Level 4: Fallback to sub-type keywords from family name (shape distinctions like round/rect)
    try:
        # richer subtype assembly: try to compose shape-fittingtype-modifier names
        fam_lower = (fam_name or '').lower()
        sub_type_keywords = clues.get("sub_type_keywords", [])
        fitting_mods = clues.get("fitting_modifier_keywords", [])

        # simple match-first: if sub_type_keywords contains matching tokens, prefer them
        found_parts = []
        # shape mapping
        shape_map = {
            'rect': 'Rec', 'rectangular': 'Rec', 'rec': 'Rec', 'round': 'Round'
        }
        # fitting type map
        fit_map = {
            'tee': 'Tee', 'elbow': 'Elbow', 'transition': 'Transition', 'wye': 'Wye', 'reducer': 'Reducer', 'takeoff': 'Takeoff'
        }
        # modifier map
        mod_map = {
            'long radius': 'LongRad', 'longrad': 'LongRad', 'lr': 'LongRad', 'long': 'LongRad',
            'short radius': 'ShortRad', 'shortrad': 'ShortRad', 'sr': 'ShortRad', 'short': 'ShortRad'
        }

        # detect shape
        shape = None
        for s_k, s_v in shape_map.items():
            if s_k in fam_lower:
                shape = s_v
                break

        # detect fitting type
        fit = None
        for f_k, f_v in fit_map.items():
            if f_k in fam_lower:
                fit = f_v
                break

        # detect modifier
        modifier = None
        # check modifiers provided in rule first
        for m in fitting_mods:
            try:
                if m.lower() in fam_lower:
                    modifier = mod_map.get(m.lower(), None) or m.replace(' ', '').title()
                    break
            except Exception:
                continue
        # fallback: search mod_map keys
        if not modifier:
            for m_k, m_v in mod_map.items():
                if m_k in fam_lower:
                    modifier = m_v
                    break

        # Build composite name in the order shape-fitting-modifier (if found)
        composite_parts = []
        if shape:
            composite_parts.append(shape)
        if fit:
            composite_parts.append(fit)
        if modifier:
            composite_parts.append(modifier)

        if composite_parts:
            specific_name = '-'.join(composite_parts)
            # if this composite doesn't clash with used names, return it
            if specific_name.lower() not in used_names:
                return specific_name
        # as a fallback, also consider single sub_type keywords
        for sub_keyword in sub_type_keywords:
            try:
                if sub_keyword.lower() in fam_lower:
                    specific_name = "{}-{}".format(generic_name, sub_keyword)
                    if specific_name.lower() not in used_names:
                        return specific_name
            except Exception:
                continue
    except Exception:
        pass

    # Finally, if no other disambiguator found, use UND suffix (Undefined system) to keep names predictable
    try:
        name_with_und = "{}-UND".format(generic_name)
        if name_with_und.lower() not in used_names:
            return name_with_und
    except Exception:
        pass

    # If all else fails, return the original generic name and let the suffixer handle it.
    return generic_name


def apply_template(template, info, rules):
    """Apply template with support for <COMPANY>, <DISC:...>, <Family>, <Type>, <Param:...>

    rules: loaded naming_rules JSON (contains COMPANY and DISCIPLINE mapping)
    """
    out = template
    company = rules.get('COMPANY', '')
    out = out.replace('<COMPANY>', company)

    # discipline token: <DISC:Name> -> lookup in rules['DISCIPLINE']
    import re
    for m in re.findall(r'<DISC:([^>]+)>', out):
        disc_map = rules.get('DISCIPLINE', {})
        out = out.replace('<DISC:%s>' % m, disc_map.get(m, m))

    # category code replacement
    try:
        cat_code = CATEGORY_CODE.get(info.get('category', ''), None)
        if cat_code:
            out = out.replace('<CAT>', cat_code)
    except Exception:
        pass

    # --- Smart Classification ---
    # Try to find a generic name first (e.g., "Pump", "AHU")
    generic_name = info.get('_classified_name') # Get pre-calculated name from classifier

    ignore_type_name = False
    if generic_name:
        fam_clean = generic_name
    else:
        # Fallback to the old behavior if no classification rule matches
        fam_raw = info.get('family_name', '')
        # remove existing company/discipline prefixes to avoid duplication
        try:
            fam_stripped = remove_existing_prefixes(fam_raw, rules) if rules else fam_raw
        except Exception:
            fam_stripped = fam_raw
        fam_clean = clean_name(fam_stripped)

    # use first type for placeholders (handle empty list safely)
    types_list = info.get('types') or []
    first_type = types_list[0] if len(types_list) > 0 else None
    typ_clean = ''
    if first_type:
        typ_raw = first_type.get('type_name', '')
        try:
            typ_stripped = remove_existing_prefixes(typ_raw, rules) if rules else typ_raw
        except Exception:
            typ_stripped = typ_raw
        typ_clean = clean_name(typ_stripped)

    # Deduplicate overlapping parts between family and type
    try:
        lf = fam_clean.lower() if fam_clean else ''
        lt = typ_clean.lower() if typ_clean else ''
        if lf and lt:
            # exact match -> keep family, clear type
            if lf == lt:
                typ_clean = ''
            else:
                import re as _re
                # if family is contained in type, remove it from type
                if lf in lt:
                    try:
                        typ_clean = _re.sub(_re.escape(fam_clean), '', typ_clean, flags=_re.I)
                    except Exception:
                        typ_clean = typ_clean.replace(fam_clean, '')
                # if type is contained in family, remove it from family
                elif lt in lf:
                    try:
                        fam_clean = _re.sub(_re.escape(typ_clean), '', fam_clean, flags=_re.I)
                    except Exception:
                        fam_clean = fam_clean.replace(typ_clean, '')
        # normalize leftover separators
        fam_clean = fam_clean.replace('--', '-').strip('- ').strip()
        typ_clean = typ_clean.replace('--', '-').strip('- ').strip()
    except Exception:
        pass

    # For Air Terminals, prefer a short family token by removing literal words like
    # 'air terminal', 'diffuser', 'register', 'grille' so names become e.g. 'Air-RA-DN600'
    try:
        if (info.get('category') or '').lower() == 'air terminals':
            try:
                fk = fam_clean or ''
                # remove commonly repeated words
                for w in ['air terminal', 'air-terminal', 'airterminal', 'diffuser', 'diff', 'register', 'grille', 'slot', 'linear', 'diffuser slot', 'square diffuser', 'rectangular diffuser', 'terminal']:
                    try:
                        fk = fk.replace(w, '')
                        fk = fk.replace(w.title(), '')
                    except Exception:
                        pass
                # fallback: if empty, use 'Air'
                fk = fk.replace('--', '-').replace('__', '_').strip(' -_')
                if not fk:
                    fk = 'Air'
                # Title-case tokens separated by dash
                try:
                    fk = '-'.join([p.title() for p in fk.split('-') if p])
                except Exception:
                    fk = fk
                fam_clean = fk
            except Exception:
                pass
    except Exception:
        pass

    out = out.replace('<Family>', fam_clean)
    # replace type and parameter placeholders depending on whether a type exists
    # Prefer classified system and size if classifier stored them into info
    classified_system = info.get('_classified_system')
    classified_size = info.get('_classified_size')

    # Determine type token value: prefer classified system (including 'UND'), otherwise type name
    type_token = ''
    try:
        if classified_system:
            type_token = classified_system
        elif typ_clean:
            type_token = typ_clean
    except Exception:
        type_token = typ_clean or ''

    out = out.replace('<Type>', type_token)

    # Explicit tokens: <SYS> and <SIZE> are shortcuts for templates
    try:
        # <SYS> -> classified system code (e.g., SAD, UND)
        # Treat 'UND' as undefined and do not include it in primary names
        try:
            sys_val = classified_system if (classified_system and str(classified_system).upper() != 'UND') else ''
        except Exception:
            sys_val = ''
        out = out.replace('<SYS>', sys_val)
    except Exception:
        pass

    try:
        # <SIZE> -> prefer classified size token (e.g., DN150). If missing, fall back
        # to common size parameters on the type or instance (Size, Diameter, Dia, D).
        size_val = ''
        if classified_size:
            size_val = classified_size
        else:
            # heuristics: check type params then instance params for size-like values
            size_param_names = ['Size', 'Diameter', 'Dia', 'D', 'Duct Size', 'Neck Size']
            try:
                # first type parameters (first_type may be None)
                if first_type:
                    for sp in size_param_names:
                        try:
                            raw = first_type.get('params', {}).get(sp, '')
                            if raw:
                                mm = _extract_first_number_with_units(raw)
                                if mm:
                                    size_val = 'DN{}'.format(mm)
                                    break
                        except Exception:
                            continue
                # then instance params if still empty
                if not size_val:
                    for sp in size_param_names:
                        try:
                            raw = info.get('instance_params', {}).get(sp, '')
                            if raw:
                                mm = _extract_first_number_with_units(raw)
                                if mm:
                                    size_val = 'DN{}'.format(mm)
                                    break
                        except Exception:
                            continue
            except Exception:
                size_val = ''
        out = out.replace('<SIZE>', size_val)
    except Exception:
        pass

    # replace parameter placeholders
    for m in re.findall(r'<Param:([^>]+)>', out):
        # prefer classified_size for size params
        if classified_size and m.lower() in ['size', 'diameter', 'dia', 'd']:
            val = classified_size
        else:
            # prefer type parameter, fallback to instance parameter
            val = ''
            try:
                if first_type:
                    # exact name match then case-insensitive fallback
                    val = first_type.get('params', {}).get(m, '') or ''
                    if not val:
                        # try case-insensitive lookup
                        for k, v in first_type.get('params', {}).items():
                            try:
                                if k.lower() == m.lower():
                                    val = v or ''
                                    break
                            except Exception:
                                continue
            except Exception:
                val = ''
            if not val:
                try:
                    # instance params case-insensitive lookup
                    inst = info.get('instance_params', {}) or {}
                    val = inst.get(m, '') or ''
                    if not val:
                        for k, v in inst.items():
                            try:
                                if k.lower() == m.lower():
                                    val = v or ''
                                    break
                            except Exception:
                                continue
                except Exception:
                    val = ''
            # DN formatting for pipe size parameter names (heuristic)
            if m.lower() in ['size', 'diameter', 'dia', 'd'] and val:
                try:
                    mm = _extract_first_number_with_units(val)
                    if mm:
                        val = 'DN%d' % mm
                except Exception:
                    pass
        out = out.replace('<Param:%s>' % m, str(val))

    # Replace new subtype tokens (<Shape>, <Fitting>, <Modifier>) from classifier info
    try:
        shape_val = info.get('_classified_shape', '') or ''
        fitting_val = info.get('_classified_fitting', '') or ''
        modifier_val = info.get('_classified_modifier', '') or ''
        # ensure no None values
        out = out.replace('<Shape>', shape_val)
        out = out.replace('<Fitting>', fitting_val)
        out = out.replace('<Modifier>', modifier_val)
    except Exception:
        pass

    # If Type token wasn't used, clear it
    out = out.replace('<Type>', '')
    # clear any remaining Param tags
    out = re.sub(r'<Param:[^>]+>', '', out)

    # final cleanup: remove duplicate separators and stray tokens
    out = out.replace('--', '-')
    out = out.replace('- -', '-')
    out = re.sub(r'\s{2,}', ' ', out)
    out = out.strip('- ').strip()
    return out


def main():
    doc = revit.doc
    rules = load_rules(RULES_FILE)
    families = get_families(doc)
    instance_params_cache = _build_first_instance_params_cache(doc, families)

    templates = rules.get('TEMPLATES', {})

    results = []
    used_names = set()
    for f in families:
        info = gather_family_info(f, instance_params_cache)
        # determine template: category, Category:Family, Family, default
        rule = None
        cat_key = info['category']
        if cat_key in templates:
            rule = templates[cat_key]
        elif "%s:%s" % (cat_key, info['family_name']) in templates:
            rule = templates["%s:%s" % (cat_key, info['family_name'])]
        elif info['family_name'] in templates:
            rule = templates[info['family_name']]
        else:
            rule = templates.get('Default', '')

        suggestion = ''
        if rule:
            # --- Smart Classification Pass ---
            classified_name = classify_family(info, rules, used_names)
            if classified_name:
                info['_classified_name'] = classified_name # Store it for apply_template
                # Immediately reserve this classified name so subsequent families will try for
                # more specific variants (system/size) instead of the same generic name.
                try:
                    used_names.add((classified_name or '').strip().lower())
                except Exception:
                    pass
            suggestion = apply_template(rule, info, rules)
            # Prefer the smart/canonical suggestion produced by templates/classifier.
            # Only fall back to the current family name when no canonical suggestion was produced.
            try:
                if not suggestion:
                    company = (rules.get('COMPANY') or 'MHT')
                    # heuristic: pick Mechanical discipline for mechanical categories
                    cat = (info.get('category') or '').lower()
                    disc_key = 'Mechanical'
                    if 'plumb' in cat or 'pipe' in cat:
                        disc_key = 'Plumbing'
                    elif 'elect' in cat or 'light' in cat:
                        disc_key = 'Electrical'
                    disc_code = rules.get('DISCIPLINE', {}).get(disc_key, 'ME')
                    desired_prefix = "%s-%s" % (company, disc_code)
                    cur = info.get('family_name') or ''
                    if cur and isinstance(cur, str) and cur.startswith(desired_prefix):
                        suggestion = cur
            except Exception:
                pass
            # Final safety: never leave suggestion blank. As a last resort, use
            # <COMPANY>-<DISC> + cleaned current family name so CSVs always have a value.
            try:
                if not suggestion:
                    company = (rules.get('COMPANY') or 'MHT')
                    cat = (info.get('category') or '').lower()
                    disc_key = 'Mechanical'
                    if 'plumb' in cat or 'pipe' in cat:
                        disc_key = 'Plumbing'
                    elif 'elect' in cat or 'light' in cat:
                        disc_key = 'Electrical'
                    disc_code = rules.get('DISCIPLINE', {}).get(disc_key, 'ME')
                    desired_prefix = "%s-%s" % (company, disc_code)
                    cur = info.get('family_name') or ''
                    try:
                        base = remove_existing_prefixes(cur, rules) if rules else cur
                    except Exception:
                        base = cur
                    try:
                        base_clean = clean_name(base)
                    except Exception:
                        base_clean = (base or '').strip()
                    if base_clean:
                        suggestion = "%s-%s" % (desired_prefix, base_clean)
                    else:
                        safe = (cur or '').replace('"', '').strip()
                        if safe:
                            suggestion = "%s-%s" % (desired_prefix, safe)
                        else:
                            suggestion = "%s-UND" % desired_prefix
            except Exception:
                try:
                    suggestion = suggestion or (info.get('family_name') or '')
                except Exception:
                    suggestion = ''

            # Normalize and guard suggestion text: collapse duplicate hyphens and
            # if the template produced multiple consecutive separators (e.g. '---')
            # treat that as an indicator of missing tokens and fall back to
            # COMPANY-DISC + cleaned family name instead.
            try:
                import re
                orig = suggestion or ''
                # If template produced three or more consecutive hyphens, fallback
                if '---' in orig:
                    company = (rules.get('COMPANY') or 'MHT')
                    cat = (info.get('category') or '').lower()
                    disc_key = 'Mechanical'
                    if 'plumb' in cat or 'pipe' in cat:
                        disc_key = 'Plumbing'
                    elif 'elect' in cat or 'light' in cat:
                        disc_key = 'Electrical'
                    disc_code = rules.get('DISCIPLINE', {}).get(disc_key, 'ME')
                    desired_prefix = "%s-%s" % (company, disc_code)
                    cur = info.get('family_name') or ''
                    try:
                        base = remove_existing_prefixes(cur, rules) if rules else cur
                    except Exception:
                        base = cur
                    try:
                        base_clean = clean_name(base)
                    except Exception:
                        base_clean = (base or '').strip()
                    if base_clean:
                        suggestion = "%s-%s" % (desired_prefix, base_clean)
                    else:
                        safe = (cur or '').replace('"', '').strip()
                        suggestion = "%s-%s" % (desired_prefix, safe or 'UND')
                else:
                    # collapse multiple hyphens to single and normalize spaces
                    s = re.sub(r'-{2,}', '-', orig)
                    s = re.sub(r'\s*-\s*', '-', s)
                    s = s.strip('- ').strip()
                    suggestion = s
            except Exception:
                pass

            # Category-specific heuristics: for Duct Fittings, prefer the cleaned
            # current family name when the generated suggestion is clearly missing
            # discriminating tokens (many consecutive hyphens, only prefix+DN, or
            # missing classified tokens). This prevents outputs such as
            # 'MHT-ME-DF---DN23' or 'MHT-ME-DF' when the family name itself is
            # informative (e.g., 'Rectangular Union').
            try:
                cat_l = (info.get('category') or '').lower()
                if 'duct' in cat_l or 'duct fittings' in cat_l:
                    s = (suggestion or '').strip()
                    # consider it bad if it's exactly the DF prefix or contains triple hyphens
                    bad = False
                    if not s:
                        bad = True
                    else:
                        if s.upper() == 'MHT-ME-DF' or '---' in s:
                            bad = True
                        # pattern: prefix optionally followed only by DN numbers and numeric suffixes
                        import re as _re
                        # matches strings like MHT-ME-DF-DN23 or MHT-ME-DF--DN23 or MHT-ME-DF-DN23-DN23
                        if _re.match(r'^MHT-ME-DF(-DN\d+(-DN\d+)*)?(-\d{2})?$', s, flags=_re.I):
                            bad = True
                        # Also treat suggestions with very short descriptive parts (no letters) as bad
                        core = _re.sub(r'^MHT-ME-DF-?', '', s, flags=_re.I)
                        if core and not _re.search(r'[A-Za-z]', core):
                            bad = True

                    # if classified tokens missing, prefer family name too
                    cls_shape = info.get('_classified_shape')
                    cls_fit = info.get('_classified_fitting')
                    cls_size = info.get('_classified_size')
                    if bad or not cls_shape or not cls_fit or not cls_size:
                        try:
                            # For duct fallbacks keep the DF prefix (MHT-ME-DF-<FamilyClean>)
                            company = (rules.get('COMPANY') or 'MHT')
                            disc_code = rules.get('DISCIPLINE', {}).get('Mechanical', 'ME')
                            df_prefix = "%s-%s-DF" % (company, disc_code)
                            cur = info.get('family_name') or ''
                            try:
                                base = remove_existing_prefixes(cur, rules) if rules else cur
                            except Exception:
                                base = cur
                            try:
                                base_clean = clean_name(base)
                            except Exception:
                                base_clean = (base or '').strip()
                            if base_clean:
                                suggestion = "%s-%s" % (df_prefix, base_clean)
                            else:
                                suggestion = "%s-%s" % (df_prefix, (cur or '').strip() or 'UND')
                        except Exception:
                            pass
            except Exception:
                pass

            results.append({'family': info['family_name'], 'category': info['category'], 'suggested': suggestion, 'info': info})

    # detect duplicate suggested names (conflicts)
    counts = {}
    for r in results:
        s = r['suggested']
        if s:
            counts[s] = counts.get(s, 0) + 1
            used_names.add(s.lower())

    # Try hierarchical conflict resolution: for duplicates, attempt to append classified size (DN...) before numeric suffixing
    try:
        for name, cnt in list(counts.items()):
            if cnt > 1:
                # find all results with this suggested name
                dup_results = [r for r in results if r.get('suggested') == name]
                for r in dup_results:
                    try:
                        size = r.get('info', {}).get('_classified_size')
                        if size:
                            candidate = "%s-%s" % (name, size)
                            if candidate.lower() not in [x.lower() for x in counts.keys()] and candidate.lower() not in [rr.get('suggested','').lower() for rr in results]:
                                r['suggested'] = candidate
                                # update counts to avoid later conflict
                                counts[candidate] = counts.get(candidate, 0) + 1
                                counts[name] = counts.get(name, 1) - 1
                    except Exception:
                        pass
    except Exception:
        pass

    # auto-resolve duplicates by appending -01, -02, etc. when duplicates exist
    used = {}
    for r in results:
        s = r['suggested']
        if not s:
            r['conflict'] = False
            continue
        if counts.get(s, 0) > 1:
            # need to create unique suffix
            idx = used.get(s, 0) + 1
            used[s] = idx
            new_name = "%s-%02d" % (s, idx)
            r['suggested'] = new_name
            r['conflict'] = True
        else:
            r['conflict'] = False

    # print results
    out = ['Family,Category,SuggestedName,Conflict,ProblemReasons']
    for r in results:
        cur = r['family']
        cat = r['category']
        sug = r['suggested']
        conf = 'DUP' if r.get('conflict') else ''
        # compute problem reasons for CSV visibility
        reasons = []
        info = r.get('info', {}) or {}
        try:
            if not sug:
                reasons.append('no_suggestion')
        except Exception:
            pass
        try:
            cat_l = (cat or '').lower()
            # Air terminal specific checks
            if 'air' in cat_l or 'terminal' in cat_l:
                sysv = info.get('_classified_system')
                sizev = info.get('_classified_size')
                shapev = info.get('_classified_shape')
                if not sysv or (isinstance(sysv, str) and (sysv.upper() == 'UND' or sysv.strip() == '')):
                    reasons.append('missing_system')
                if not sizev:
                    reasons.append('missing_size')
                if not shapev:
                    reasons.append('missing_shape')
            # Duct fitting specific checks
            if 'duct' in cat_l:
                s = (sug or '').strip()
                if not s or s.upper() == 'MHT-ME-DF' or '---' in s:
                    reasons.append('incomplete_df_template')
                # missing classified tokens
                if not info.get('_classified_shape'):
                    reasons.append('missing_shape')
                if not info.get('_classified_fitting'):
                    reasons.append('missing_fitting')
                if not info.get('_classified_size'):
                    reasons.append('missing_size')
        except Exception:
            pass
        pr = ';'.join(reasons)
        out.append('"{}","{}","{}","{}","{}"'.format(cur, cat, sug, conf, pr))

    csv_text = '\n'.join(out)
    logger.info(csv_text)

    # write CSV next to script for review (disabled by default; enable via naming_rules.json AUTO_EXPORT: true)
    try:
        auto_export = bool(rules.get('AUTO_EXPORT', False))
    except Exception:
        auto_export = False

    if auto_export:
        try:
            from datetime import datetime
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            fname = 'family_name_suggestions_{}.csv'.format(ts)
            csv_path = os.path.join(SCRIPT_DIR, fname)
            # write with explicit encoding and newline handling
            # write CSV using cross-runtime helper to support IronPython
            try:
                ok = _write_text_file(csv_path, csv_text, encoding='utf-8')
                if not ok:
                    raise Exception('Failed to write CSV via helper')
            except Exception as _e:
                raise
            logger.info('Wrote suggestions to {}'.format(csv_path))
            # Also write a JSON debug dump for problematic rows to help rule tuning.
            try:
                debug_rows = []
                for r in results:
                    info = r.get('info', {}) or {}
                    is_problem = False
                    reasons = []
                    if not r.get('suggested'):
                        is_problem = True
                        reasons.append('no_suggestion')
                    try:
                        cat = (r.get('category') or '').lower()
                    except Exception:
                        cat = ''
                    # flag Air Terminal like items where key tokens are missing
                    if 'air' in cat or 'terminal' in cat or 'air terminal' in cat:
                        sysv = info.get('_classified_system')
                        sizev = info.get('_classified_size')
                        shapev = info.get('_classified_shape')
                        if not sysv or (isinstance(sysv, str) and (sysv.upper() == 'UND' or sysv.strip() == '')):
                            is_problem = True
                            reasons.append('missing_system')
                        if not sizev:
                            is_problem = True
                            reasons.append('missing_size')
                        if not shapev:
                            reasons.append('missing_shape')
                    if is_problem:
                        debug_rows.append({
                            'family': r.get('family'),
                            'category': r.get('category'),
                            'suggested': r.get('suggested'),
                            'reasons': reasons,
                            'classified_system': info.get('_classified_system'),
                            'classified_size': info.get('_classified_size'),
                            'classified_shape': info.get('_classified_shape'),
                            'classified_fitting': info.get('_classified_fitting'),
                            'classified_modifier': info.get('_classified_modifier'),
                            'type_params_first': (info.get('types') or [])[:1],
                            'instance_params': info.get('instance_params', {})
                        })
                if debug_rows:
                    try:
                        dbg_fname = 'family_name_suggestions_debug_{}.json'.format(ts)
                        dbg_path = os.path.join(SCRIPT_DIR, dbg_fname)
                        try:
                            dbg_text = json.dumps({'generated': ts, 'problems': debug_rows}, indent=2, ensure_ascii=False)
                            _write_text_file(dbg_path, dbg_text, encoding='utf-8')
                        except Exception:
                            raise
                        logger.info('Wrote debug JSON to {}'.format(dbg_path))
                    except Exception as _:
                        logger.warning('Failed to write debug JSON: {}'.format(_))
            except Exception:
                pass
            # Store csv_path in globals so UI can access it if user clicks "Open CSV" button
            try:
                globals()['_last_csv_path'] = csv_path
            except Exception:
                pass
        except Exception as e:
            try:
                # write detailed traceback to a file so the user can copy it
                import traceback
                from datetime import datetime
                ts_err = datetime.now().strftime('%Y%m%d_%H%M%S')
                err_fname = 'family_name_suggestions_export_error_{}.txt'.format(ts_err)
                err_path = os.path.join(SCRIPT_DIR, err_fname)
                try:
                    _write_text_file(err_path, 'Export error: {}\n\n{}'.format(e, traceback.format_exc()), encoding='utf-8')
                except Exception:
                    pass
                logger.warning('Failed to write CSV: {}. Wrote details to {}'.format(e, err_path))
                try:
                    # show a dialog so the user knows where to find the error details in Revit
                    try:
                        _show_error_dialog('Export failed. Detailed error written to:\n{}'.format(err_path), 'Export Failed')
                    except Exception:
                        pass
                except Exception:
                    pass
            except Exception:
                logger.warning('Failed to write CSV and also failed to write error details: {}'.format(e))
    else:
        logger.info('Auto export disabled (set AUTO_EXPORT: true in naming_rules.json to enable)')

    # Build a simple review UI (only when WinForms available)
    try:
        if DataGridView is None:
            raise Exception('WinForms not available in this environment')

        def apply_casing(s, mode):
            if not s:
                return s
            mode = (mode or '').upper()
            if mode == 'UPPER':
                return s.upper()
            if mode == 'TITLE':
                return '-'.join([w.title() for w in s.split('-')])
            return s

        class ReviewForm(Form):
            def __init__(self, rows, rules=None):
                self.Text = 'MHT Family Namer - Review Suggestions'
                self.Width = 1200
                self.Height = 800
                self.MinimumSize = Size(800, 600)
                self.StartPosition = FormStartPosition.CenterParent
                try:
                    self.WindowState = FormWindowState.Normal
                except Exception:
                    # fallback for environments where enum conversion fails
                    self.WindowState = 0
                # Enable resizing and ensure form responds to window size changes
                try:
                    self.AutoScaleDimensions = SizeF(6, 13)
                    self.AutoScaleMode = AutoScaleMode.Font
                except Exception:
                    pass

                # Professional dark theme colors
                self.bg_dark = Color.FromArgb(30, 30, 30)          # Main background (dark gray)
                self.bg_control = Color.FromArgb(45, 45, 45)       # Control background (lighter dark gray)
                self.bg_header = Color.FromArgb(20, 20, 20)        # Header background (darkest)
                self.fg_text = Color.FromArgb(220, 220, 220)       # Text color (light gray)
                self.accent = Color.FromArgb(0, 120, 212)          # Accent color (blue)
                self.accent_dark = Color.FromArgb(0, 90, 160)      # Darker accent

                # Set form colors
                self.BackColor = self.bg_dark
                self.ForeColor = self.fg_text

                self.dgv = DataGridView()
                self.dgv.Dock = DockStyle.Fill
                self.dgv.AllowUserToAddRows = False
                try:
                    # Try to set Fill mode; if enum unavailable, skip it
                    self.dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                except Exception:
                    # Fallback: Leave default (None) to avoid conversion errors
                    pass
                try:
                    self.dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
                except Exception:
                    pass
                self.dgv.AllowUserToResizeColumns = True
                self.dgv.AllowUserToResizeRows = False
                
                # Professional dark theme for DataGridView
                self.dgv.BackgroundColor = self.bg_dark
                self.dgv.ForeColor = self.fg_text
                self.dgv.GridColor = self.bg_control
                self.dgv.DefaultCellStyle.BackColor = self.bg_dark
                self.dgv.DefaultCellStyle.ForeColor = self.fg_text
                self.dgv.DefaultCellStyle.SelectionBackColor = self.accent
                self.dgv.DefaultCellStyle.SelectionForeColor = Color.White
                self.dgv.ColumnHeadersDefaultCellStyle.BackColor = self.bg_header
                self.dgv.ColumnHeadersDefaultCellStyle.ForeColor = self.fg_text
                self.dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = self.accent_dark
                self.dgv.EnableHeadersVisualStyles = False
                self.dgv.RowHeadersDefaultCellStyle.BackColor = self.bg_header
                self.dgv.RowHeadersDefaultCellStyle.ForeColor = self.fg_text
                # Attach a DataError handler to avoid the default modal error dialog
                try:
                    def _on_dgv_data_error(sender, evt):
                        try:
                            # suppress the modal exception dialog; log instead if logger available
                            try:
                                ex = getattr(evt, 'Exception', None)
                                script.get_logger().warning('DataGridView data error: {}'.format(ex))
                            except Exception:
                                pass
                            try:
                                evt.ThrowException = False
                            except Exception:
                                pass
                        except Exception:
                            pass
                    try:
                        self.dgv.DataError += _on_dgv_data_error
                    except Exception:
                        pass
                except Exception:
                    pass

                chk = DataGridViewCheckBoxColumn()
                chk.HeaderText = 'Apply'
                chk.Width = 50
                self.dgv.Columns.Add(chk)

                c1 = DataGridViewTextBoxColumn()
                c1.HeaderText = 'Family'
                c1.ReadOnly = True
                c1.Width = 300
                self.dgv.Columns.Add(c1)
                # hide the original Family column and keep Current Name visible
                try:
                    c1.Visible = False
                except Exception:
                    try:
                        # fallback: set on the grid if column object doesn't accept Visible
                        self.dgv.Columns[1].Visible = False
                    except Exception:
                        pass

                # show current family name explicitly (placed before category)
                c1b = DataGridViewTextBoxColumn()
                c1b.HeaderText = 'Current Name'
                c1b.ReadOnly = True
                c1b.Width = 300
                self.dgv.Columns.Add(c1b)

                c2 = DataGridViewTextBoxColumn()
                c2.HeaderText = 'Category'
                c2.ReadOnly = True
                c2.Width = 150
                self.dgv.Columns.Add(c2)

                # Suggested Name column: use a ComboBox column so users can pick alternatives
                try:
                    from System.Windows.Forms import DataGridViewComboBoxColumn
                    c3 = DataGridViewComboBoxColumn()
                    c3.HeaderText = 'Suggested Name'
                    c3.ReadOnly = False
                    c3.Width = 400
                    self.dgv.Columns.Add(c3)
                    use_combo = True
                except Exception:
                    # fallback to a text column if ComboBox column isn't available
                    c3 = DataGridViewTextBoxColumn()
                    c3.HeaderText = 'Suggested Name'
                    c3.ReadOnly = False
                    c3.Width = 400
                    self.dgv.Columns.Add(c3)
                    use_combo = False

                c4 = DataGridViewTextBoxColumn()
                c4.HeaderText = 'Conflict'
                c4.ReadOnly = True
                c4.Width = 80
                self.dgv.Columns.Add(c4)

                # store rows and rules for use in other instance methods
                self._rows = rows or []
                self._rules = rules or {}

                # prepare row data (don't add yet)
                row_data = []
                for r in rows:
                    fam = r.get('family', '')
                    cat = r.get('category', '')
                    sug = r.get('suggested', '')
                    conf = 'DUP' if r.get('conflict') else ''
                    # duplicate fam into a 'Current Name' column so it's visible before Category
                    # include original info dict so we can build alternative suggestions per-row
                    row_data.append((fam, fam, cat, sug, conf, r.get('conflict', False), r.get('info')))

                # controls
                # per-column filter panel (Family, Current Name, Category, Suggested)
                try:
                    from System.Windows.Forms import FlowLayoutPanel
                    self.filterPanel = FlowLayoutPanel()
                    try:
                        self.filterPanel.FlowDirection = FlowDirection.LeftToRight
                    except Exception:
                        # fallback: leave default or numeric value
                        self.filterPanel.FlowDirection = 0
                    self.filterPanel.Dock = DockStyle.Top
                    self.filterPanel.Height = 28

                    # Family column is hidden; skip creating a separate Family filter

                    self.filterCurrent = TextBox()
                    try:
                        setattr(self.filterCurrent, 'PlaceholderText', 'Filter Current Name')
                    except Exception:
                        pass
                    self.filterCurrent.Width = 200
                    self.filterCurrent.TextChanged += self.on_filter_changed
                    self.filterCurrent.BackColor = self.bg_control
                    self.filterCurrent.ForeColor = self.fg_text
                    self.filterPanel.Controls.Add(self.filterCurrent)

                    self.filterCategory = TextBox()
                    try:
                        setattr(self.filterCategory, 'PlaceholderText', 'Filter Category')
                    except Exception:
                        pass
                    self.filterCategory.Width = 150
                    self.filterCategory.TextChanged += self.on_filter_changed
                    self.filterCategory.BackColor = self.bg_control
                    self.filterCategory.ForeColor = self.fg_text
                    self.filterPanel.Controls.Add(self.filterCategory)

                    self.filterSuggested = TextBox()
                    try:
                        setattr(self.filterSuggested, 'PlaceholderText', 'Filter Suggested')
                    except Exception:
                        pass
                    self.filterSuggested.Width = 250
                    self.filterSuggested.TextChanged += self.on_filter_changed
                    self.filterSuggested.BackColor = self.bg_control
                    self.filterSuggested.ForeColor = self.fg_text
                    self.filterPanel.Controls.Add(self.filterSuggested)

                    # Set filter panel background for better visibility
                    self.filterPanel.BackColor = self.bg_header

                    # Exclude DUP checkbox: hide rows that contain DUP in current or suggested names
                    try:
                        from System.Windows.Forms import CheckBox
                        self.chkExcludeDup = CheckBox()
                        self.chkExcludeDup.Text = 'Exclude DUP'
                        self.chkExcludeDup.Checked = False
                        self.chkExcludeDup.AutoSize = True
                        self.chkExcludeDup.CheckedChanged += self.on_filter_changed
                        self.filterPanel.Controls.Add(self.chkExcludeDup)
                    except Exception:
                        self.chkExcludeDup = None
                except Exception:
                    self.filterPanel = None
                    self.filterFamily = None
                    self.filterCurrent = None
                    self.filterCategory = None
                    self.filterSuggested = None

                # search box (backwards compatibility)
                self.txtSearch = TextBox()
                try:
                    # some .NET versions support PlaceholderText
                    setattr(self.txtSearch, 'PlaceholderText', 'Search family/category/suggested')
                except Exception:
                    pass
                self.txtSearch.Dock = DockStyle.Bottom

                self.btnFilter = Button()
                self.btnFilter.Text = 'Filter'
                self.btnFilter.Dock = DockStyle.Bottom
                self.btnFilter.Height = 24
                self.btnFilter.Click += self.on_filter

                self.btnApply = Button()
                self.btnApply.Text = 'Apply Selected'
                self.btnApply.Dock = DockStyle.Bottom
                self.btnApply.Height = 30
                self.btnApply.Click += self.on_apply

                self.btnSelectAll = Button()
                self.btnSelectAll.Text = 'Select All'
                self.btnSelectAll.Dock = DockStyle.Bottom
                self.btnSelectAll.Height = 24
                self.btnSelectAll.Click += self.on_select_all

                self.btnClearAll = Button()
                self.btnClearAll.Text = 'Clear All'
                self.btnClearAll.Dock = DockStyle.Bottom
                self.btnClearAll.Height = 24
                self.btnClearAll.Click += self.on_clear_all

                self.btnExport = Button()
                self.btnExport.Text = 'Export CSV'
                self.btnExport.Dock = DockStyle.Bottom
                self.btnExport.Height = 30
                self.btnExport.Click += self.on_export

                self.btnOpenCSV = Button()
                self.btnOpenCSV.Text = 'Open Last CSV'
                self.btnOpenCSV.Dock = DockStyle.Bottom
                self.btnOpenCSV.Height = 30
                self.btnOpenCSV.Click += self.on_open_csv

                # Close All button: confirm and close all related tool windows
                self.btnCloseAll = Button()
                self.btnCloseAll.Text = 'Close All'
                self.btnCloseAll.Dock = DockStyle.Bottom
                self.btnCloseAll.Height = 30
                self.btnCloseAll.Click += self.on_close_all

                # status label
                self.lblStatus = Label()
                self.lblStatus.Text = ''
                self.lblStatus.Dock = DockStyle.Bottom
                self.lblStatus.ForeColor = self.fg_text
                self.lblStatus.BackColor = self.bg_header
                self.lblStatus.Height = 25
                try:
                    self.lblStatus.AutoSize = False
                except Exception:
                    pass

                # Timer to update status (shows if a modeless Type Renamer is open)
                try:
                    from System.Windows.Forms import Timer
                    self._status_timer = Timer()
                    self._status_timer.Interval = 1000
                    self._status_timer.Tick += self._on_status_tick
                    try:
                        self._status_timer.Start()
                    except Exception:
                        pass
                except Exception:
                    self._status_timer = None

                # Ensure timer is stopped when the form is closed to avoid Tick firing
                try:
                    def _on_closed(sender, evt):
                        try:
                            if getattr(self, '_status_timer', None) is not None:
                                try:
                                    self._status_timer.Stop()
                                except Exception:
                                    pass
                                try:
                                    # detach handler if possible
                                    self._status_timer.Tick -= self._on_status_tick
                                except Exception:
                                    pass
                                try:
                                    self._status_timer.Dispose()
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    try:
                        self.Closed += _on_closed
                    except Exception:
                        pass
                except Exception:
                    pass

                # now add rows with conflict coloring
                for fam, cur_name, cat, sug, conf_text, is_conflict, info_obj in row_data:
                    # columns: Apply, Family, Current Name, Category, Suggested Name, Conflict
                    # Insert row with an empty placeholder for the Suggested cell so per-cell
                    # ComboBox items can be added afterwards without triggering validation.
                    row_idx = self.dgv.Rows.Add(False, fam, cur_name, cat, '', conf_text)
                    # If this row is marked as conflict, color it
                    if is_conflict and Color is not None:
                        try:
                            for c in range(self.dgv.Columns.Count):
                                self.dgv.Rows[row_idx].Cells[c].Style.BackColor = Color.LightCoral
                        except Exception:
                            pass

                    # Populate alternatives into the Suggested Name combo cell (if available)
                    try:
                        suggested_cell = self.dgv.Rows[row_idx].Cells[4]
                        # Build alternatives list
                        def build_alternatives(info_local, current, suggested):
                            """Build a filtered list of alternatives that respect COMPANY-DISC prefix.

                            Rules:
                            - If the current name already starts with the desired company+discipline prefix, keep it as first option.
                            - Only include alternatives that start with the same company+discipline prefix (e.g., MHT-ME-...).
                            - Build canonical suggestions using the template where possible to ensure prefix correctness.
                            """
                            opts = []
                            try:
                                company = (rules.get('COMPANY') or 'MHT')
                                # choose discipline key heuristically based on category
                                cat = (info_local.get('category') or '').lower()
                                disc_key = 'Mechanical'
                                if 'plumb' in cat or 'pipe' in cat:
                                    disc_key = 'Plumbing'
                                elif 'elect' in cat or 'lighting' in cat:
                                    disc_key = 'Electrical'
                                elif 'annotation' in cat or 'tag' in cat:
                                    disc_key = 'Annotation'
                                disc_code = rules.get('DISCIPLINE', {}).get(disc_key, 'ME')
                                desired_prefix = "%s-%s" % (company, disc_code)

                                # helper to test prefix
                                def has_desired_prefix(s):
                                    try:
                                        return bool(s) and str(s).startswith(desired_prefix)
                                    except Exception:
                                        return False

                                # ensure suggested value is canonical and prefixed - recompute from template if possible
                                try:
                                    # find template for this category
                                    templates = rules.get('TEMPLATES', {})
                                    tpl = templates.get(info_local.get('category')) or templates.get('Default')
                                    if tpl:
                                        canonical = apply_template(tpl, info_local, rules)
                                    else:
                                        canonical = suggested
                                except Exception:
                                    canonical = suggested

                                # Prefer canonical first (if it matches desired prefix)
                                if canonical and has_desired_prefix(canonical) and canonical not in opts:
                                    opts.append(canonical)
                                # then prefer current if it matches desired prefix
                                if current and has_desired_prefix(current) and current not in opts:
                                    opts.append(current)

                                # family/generic based candidates (always prefix them)
                                fam_raw = info_local.get('_classified_name') or info_local.get('family_name') or ''
                                try:
                                    fam_stripped = remove_existing_prefixes(fam_raw, rules) if rules else fam_raw
                                except Exception:
                                    fam_stripped = fam_raw
                                fam_clean = clean_name(fam_stripped)
                                syscode = info_local.get('_classified_system') or ''
                                size = info_local.get('_classified_size') or ''
                                if fam_clean:
                                    # fam + sys
                                    if syscode:
                                        cand = "%s-%s-{}".format(syscode) % (desired_prefix, 'DF') if False else None
                                    # Instead of ad-hoc strings, build prefixed forms via template when possible
                                    try:
                                                if tpl:
                                                    # create an info copy and force family-based simple template
                                                    info_copy = dict(info_local)
                                                    # use family short as Family token
                                                    info_copy['family_name'] = fam_clean
                                                    cand = apply_template(tpl, info_copy, rules)
                                                    if cand and has_desired_prefix(cand) and cand not in opts:
                                                        # prefer these family-based canonicals after the main canonical
                                                        opts.append(cand)
                                    except Exception:
                                        pass

                                # type-based suggestion
                                try:
                                    types_list = info_local.get('types') or []
                                    first_type = types_list[0] if len(types_list) > 0 else None
                                    if first_type:
                                        typ_raw = first_type.get('type_name', '')
                                        try:
                                            typ_stripped = remove_existing_prefixes(typ_raw, rules) if rules else typ_raw
                                        except Exception:
                                            typ_stripped = typ_raw
                                        typ_clean = clean_name(typ_stripped)
                                        if typ_clean:
                                            # build via template where possible
                                            try:
                                                info_copy = dict(info_local)
                                                info_copy['types'] = [{'type_name': typ_clean, 'params': first_type.get('params', {})}]
                                                if tpl:
                                                    cand = apply_template(tpl, info_copy, rules)
                                                    if cand and has_desired_prefix(cand) and cand not in opts:
                                                        opts.append(cand)
                                                else:
                                                    cand = "%s-%s" % (desired_prefix, typ_clean)
                                                    if has_desired_prefix(cand) and cand not in opts:
                                                        opts.append(cand)
                                            except Exception:
                                                pass
                                except Exception:
                                    pass

                                # a simple auto-suffix option (ensure it's prefixed)
                                if canonical:
                                    try:
                                        suff = "%s-01" % canonical
                                        if has_desired_prefix(suff) and suff not in opts:
                                            opts.append(suff)
                                    except Exception:
                                        pass

                                # Finally, if no options yet and current doesn't match, allow current only if it already contains company code
                                if not opts and current and has_desired_prefix(current):
                                    opts.append(current)

                                # As a last resort, include canonical even if not prefixed (but user requested to hide these, so skip)
                            except Exception:
                                pass
                            # if no options were produced, emit debug info to help rule tuning
                            try:
                                if not opts:
                                    try:
                                        logger = getattr(script, 'get_logger')() if hasattr(script, 'get_logger') else None
                                    except Exception:
                                        logger = None
                                    try:
                                        dbg = {
                                            'family_name': info_local.get('family_name'),
                                            'category': info_local.get('category'),
                                            'current': current,
                                            'suggested': suggested,
                                            '_classified_system': info_local.get('_classified_system'),
                                            '_classified_shape': info_local.get('_classified_shape'),
                                            '_classified_size': info_local.get('_classified_size'),
                                            'type_params': (info_local.get('types') or [])[:1],
                                            'instance_params': info_local.get('instance_params', {})
                                        }
                                        try:
                                            if logger:
                                                logger.warning('No alternatives generated for family: {}'.format(dbg))
                                            else:
                                                print('No alternatives generated for family:', dbg)
                                        except Exception:
                                            try:
                                                print('No alternatives (logger failed):', dbg)
                                            except Exception:
                                                pass
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                            return opts

                        alternatives = build_alternatives(info_obj or {}, cur_name or '', sug or '')
                        # Ensure the precomputed suggestion from results appears first
                        try:
                            if sug and sug not in alternatives:
                                alternatives.insert(0, sug)
                        except Exception:
                            pass
                        # If the cell supports Items, add them; otherwise leave the text value
                        try:
                            if hasattr(suggested_cell, 'Items'):
                                # clear existing items if possible
                                try:
                                    suggested_cell.Items.Clear()
                                except Exception:
                                    pass
                                # Add alternatives with canonical first (alternatives already prioritise canonical)
                                try:
                                    for opt in alternatives:
                                        try:
                                            if not opt:
                                                continue
                                            suggested_cell.Items.Add(opt)
                                        except Exception:
                                            pass
                                except Exception:
                                    pass

                                # Fallback: if no alternatives were generated, ensure suggested value is in the cell
                                try:
                                    if (getattr(suggested_cell.Items, 'Count', 0) or 0) == 0 and sug:
                                        suggested_cell.Items.Add(sug)
                                except Exception:
                                    pass

                                # set selection: pick the first item (canonical preferred)
                                try:
                                    if getattr(suggested_cell.Items, 'Count', 0) > 0:
                                        suggested_cell.Value = suggested_cell.Items[0]
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    except Exception:
                        pass

                # casing controls
                self.lbl = Label()
                self.lbl.Text = 'Casing:'
                self.lbl.Dock = DockStyle.Bottom

                self.cbo = ComboBox()
                self.Controls.Add(self.btnClearAll)
                self.Controls.Add(self.btnSelectAll)
                self.cbo.Items.Add('UPPER')
                self.cbo.Items.Add('TITLE')
                self.cbo.Items.Add('ASIS')
                try:
                    default_casing = rules.get('DEFAULT_CASING', 'UPPER')
                except Exception:
                    default_casing = 'UPPER'
                if default_casing in ['UPPER', 'TITLE', 'ASIS']:
                    self.cbo.SelectedItem = default_casing
                else:
                    self.cbo.SelectedItem = 'UPPER'
                self.cbo.Dock = DockStyle.Bottom

                self.btnRefresh = Button()
                self.btnRefresh.Text = 'Refresh Casing'
                self.btnRefresh.Dock = DockStyle.Bottom
                self.btnRefresh.Height = 24
                self.btnRefresh.Click += self.on_refresh

                # add controls (order: bottom up)
                # add controls (order: bottom up)
                # Add controls (bottom up)
                self.Controls.Add(self.dgv)
                self.Controls.Add(self.btnExport)
                self.Controls.Add(self.btnOpenCSV)
                self.Controls.Add(self.btnApply)
                self.Controls.Add(self.btnCloseAll)
                self.Controls.Add(self.btnRefresh)
                self.Controls.Add(self.btnFilter)
                self.Controls.Add(self.txtSearch)
                self.Controls.Add(self.cbo)
                self.Controls.Add(self.lblStatus)
                self.Controls.Add(self.lbl)

                # Dark theme for all buttons and controls
                for ctrl in [self.btnExport, self.btnOpenCSV, self.btnApply, self.btnCloseAll, self.btnRefresh, 
                           self.btnFilter, self.btnSelectAll, self.btnClearAll]:
                    try:
                        ctrl.BackColor = self.bg_control
                        ctrl.ForeColor = self.fg_text
                        ctrl.FlatStyle = FlatStyle.Flat
                        # Set flat button appearance
                        ctrl.FlatAppearance.BorderColor = self.accent
                        ctrl.FlatAppearance.MouseDownBackColor = self.accent_dark
                        ctrl.FlatAppearance.MouseOverBackColor = self.accent
                    except Exception:
                        pass

                # Dark theme for TextBox
                try:
                    self.txtSearch.BackColor = self.bg_control
                    self.txtSearch.ForeColor = self.fg_text
                except Exception:
                    pass

                # Dark theme for ComboBox
                try:
                    self.cbo.BackColor = self.bg_control
                    self.cbo.ForeColor = self.fg_text
                except Exception:
                    pass
                # add filter panel above the grid if available
                try:
                    if self.filterPanel is not None:
                        self.Controls.Add(self.filterPanel)
                except Exception:
                    pass

            def on_export(self, sender, args):
                try:
                    # Use the manual export helper so UI-triggered exports match programmatic exports.
                    # Do not automatically open the file from the UI export unless rules request it.
                    try:
                        auto_open = bool(self._rules.get('AUTO_OPEN', False))
                    except Exception:
                        auto_open = False
                    csv_path, dbg_path = export_results_now(self._rows, rules=self._rules, script_dir=SCRIPT_DIR, open_file=auto_open)
                    # inform user where files were written
                    try:
                        msg = 'Exported CSV to: {}'.format(csv_path)
                        if dbg_path:
                            msg = msg + '\nDebug JSON: {}'.format(dbg_path)
                        # Always update status label; only show dialog when enabled.
                        try:
                            self.lblStatus.Text = msg
                        except Exception:
                            script.get_logger().info(msg)
                        _maybe_dialog(msg, 'Export Completed', MessageBoxButtons.OK, MessageBoxIcon.Information, default_result=None)
                    except Exception:
                        script.get_logger().info('Export completed')
                except Exception as e:
                    try:
                        import traceback
                        from datetime import datetime
                        ts_err = datetime.now().strftime('%Y%m%d_%H%M%S')
                        err_fname = 'family_name_suggestions_export_error_{}.txt'.format(ts_err)
                        err_path = os.path.join(SCRIPT_DIR, err_fname)
                        try:
                            _write_text_file(err_path, 'Export failed: {}\n\n{}'.format(e, traceback.format_exc()), encoding='utf-8')
                            script.get_logger().warning('Export failed: {}. Details written to {}'.format(e, err_path))
                        except Exception:
                            script.get_logger().warning('Export failed and failed to write error details: {}'.format(e))
                        try:
                            _show_error_dialog('Export failed. See details in:\n{}'.format(err_path), 'Export Failed')
                        except Exception:
                            pass
                    except Exception:
                        script.get_logger().warning('Export failed and failed to write error details: {}'.format(e))

            def on_open_csv(self, sender, args):
                """Open the last exported CSV file."""
                try:
                    csv_path = globals().get('_last_csv_path')
                    if not csv_path or not os.path.exists(csv_path):
                        _show_error_dialog('No CSV file found or file was moved/deleted.', 'Open CSV')
                        return
                    
                    # Open the file
                    if os.name == 'nt':
                        try:
                            os.startfile(csv_path)
                        except Exception:
                            try:
                                import webbrowser
                                webbrowser.open('file://' + os.path.realpath(csv_path))
                            except Exception:
                                _show_error_dialog('Could not open CSV file: {}'.format(csv_path), 'Open CSV')
                    else:
                        try:
                            import webbrowser
                            webbrowser.open('file://' + os.path.realpath(csv_path))
                        except Exception:
                            _show_error_dialog('Could not open CSV file: {}'.format(csv_path), 'Open CSV')
                except Exception as e:
                    _show_error_dialog('Error opening CSV: {}'.format(e), 'Open CSV')

            def update_status(self):
                try:
                    total = 0
                    selected = 0
                    for i in range(self.dgv.Rows.Count):
                        try:
                            if not hasattr(self.dgv.Rows[i], 'Visible') or self.dgv.Rows[i].Visible:
                                total += 1
                                if self.dgv.Rows[i].Cells[0].Value:
                                    selected += 1
                        except Exception:
                            pass
                    # if there are open modeless type renamer windows, show a visual cue
                    try:
                        open_forms = globals().get('_open_modeless_forms') or []
                        if open_forms:
                            self.lblStatus.Text = 'Total: %d    Selected: %d    [Type Renamer Open]' % (total, selected)
                        else:
                            self.lblStatus.Text = 'Total: %d    Selected: %d' % (total, selected)
                    except Exception:
                        self.lblStatus.Text = 'Total: %d    Selected: %d' % (total, selected)
                except Exception:
                    pass

            def _on_status_tick(self, sender, args):
                try:
                    # update status every tick to reflect open modeless windows
                    self.update_status()
                except Exception:
                    pass

            def on_filter(self, sender, args):
                try:
                    # if per-column filters present, use them; otherwise fallback to txtSearch
                    # Family filter removed; use current name filter only
                    f_cur = (self.filterCurrent.Text.strip().lower() if (self.filterCurrent and self.filterCurrent.Text) else '')
                    f_cat = (self.filterCategory.Text.strip().lower() if (self.filterCategory and self.filterCategory.Text) else '')
                    f_sug = (self.filterSuggested.Text.strip().lower() if (self.filterSuggested and self.filterSuggested.Text) else '')
                    q = ''
                    try:
                        q = self.txtSearch.Text.strip().lower() if self.txtSearch.Text else ''
                    except Exception:
                        q = ''

                    for i in range(self.dgv.Rows.Count):
                        try:
                            row = self.dgv.Rows[i]
                            # Family column is hidden; Current Name is at column index 2
                            fam = (row.Cells[1].Value or '').lower()
                            cur = (row.Cells[2].Value or '').lower()
                            cat = (row.Cells[3].Value or '').lower()
                            sug = (row.Cells[4].Value or '').lower()
                            # handle Exclude DUP checkbox
                            try:
                                if getattr(self, 'chkExcludeDup', None) and getattr(self.chkExcludeDup, 'Checked', False):
                                    if 'dup' in cur or 'dup' in sug:
                                        visible = False
                            except Exception:
                                pass
                            visible = True
                            if q and not (q in fam or q in cur or q in cat or q in sug):
                                visible = False
                            # no family filter: rely on current name filter instead
                            if f_cur and f_cur not in cur:
                                visible = False
                            if f_cur and f_cur not in cur:
                                visible = False
                            if f_cat and f_cat not in cat:
                                visible = False
                            if f_sug and f_sug not in sug:
                                visible = False

                            row.Visible = bool(visible)
                        except Exception:
                            pass
                    self.update_status()
                except Exception:
                    pass

            def on_filter_changed(self, sender, args):
                # TextChanged handler for per-column filters
                try:
                    self.on_filter(None, None)
                except Exception:
                    pass

            def on_remove_dup(self, sender, args):
                try:
                    rows_to_remove = []
                    for i in range(self.dgv.Rows.Count):
                        row = self.dgv.Rows[i]
                        if (row.Cells[5].Value or '') == 'DUP':
                            rows_to_remove.append(row)
                    for row in rows_to_remove:
                        self.dgv.Rows.Remove(row)
                    self.update_status()
                except Exception as e:
                    script.get_logger().warning('Failed to remove DUP rows: {}'.format(e))

            def on_apply(self, sender, args):
                # prevent concurrent apply operations
                try:
                    if globals().get('_mht_is_applying'):
                        _show_error_dialog('An apply operation is already in progress. Please wait until it finishes.', 'MHT Family Namer')
                        return
                except Exception:
                    pass

                # if there are open modeless type renamer windows, warn user
                try:
                    open_forms = globals().get('_open_modeless_forms') or []
                    if open_forms:
                        try:
                            res = _maybe_dialog(
                                'There are open Type Renamer windows. It is safer to close them before applying family renames. Continue anyway?',
                                'MHT Family Namer',
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Warning,
                                default_result=DialogResult.Yes
                            )
                            if res is not None and res != DialogResult.Yes:
                                return
                        except Exception:
                            # if messagebox fails, be conservative and stop
                            return
                except Exception:
                    pass

                # collect selected rows
                to_apply = []
                for i in range(self.dgv.Rows.Count):
                    cell = self.dgv.Rows[i].Cells[0].Value
                    if cell:
                        # Current Name is at column index 2 (Family column is hidden)
                        try:
                            fam_name = (self.dgv.Rows[i].Cells[2].Value or '').strip()
                        except Exception:
                            fam_name = (self.dgv.Rows[i].Cells[1].Value or '').strip()
                        # Suggested Name is at column index 4 (may be a ComboBox cell)
                        sug = ''
                        try:
                            sug_cell = self.dgv.Rows[i].Cells[4]
                            # prefer Value, fallback to FormattedValue or ToString
                            try:
                                val = getattr(sug_cell, 'Value', None)
                                if val is None:
                                    val = getattr(sug_cell, 'FormattedValue', None)
                                if val is None:
                                    # last resort: read the cell's DisplayedValue or call str
                                    try:
                                        val = str(sug_cell)
                                    except Exception:
                                        val = ''
                                sug = (val or '').strip()
                            except Exception:
                                try:
                                    sug = str(sug_cell.Value).strip()
                                except Exception:
                                    sug = ''
                        except Exception:
                            sug = ''
                        to_apply.append((fam_name, sug))

                if not to_apply:
                    script.get_logger().info('No rows selected to apply')
                    return

                # Handle exact duplicates before proceeding
                to_apply, duplicate_actions = handle_exact_duplicates(doc, to_apply, results)

                    # pre-check for name conflicts and ask user how to handle them
                try:
                    existing = _collect_existing_family_names(doc)
                    conflicts = []
                    for old, new in to_apply:
                        try:
                            if new and new.strip().lower() in existing and (old or '').strip().lower() != new.strip().lower():
                                conflicts.append({'old': old, 'new': new, 'action': 'ask'})
                        except Exception:
                            pass

                    auto_suffix = False
                    # If conflicts found, open a ConflictReviewForm to let user choose per-conflict action
                    if conflicts and Form is not None:
                        try:
                            # Build a compact list of strings for the dialog
                            class ConflictReviewForm(Form):
                                def __init__(self, items):
                                    self.Text = 'Conflicting Name Review'
                                    self.Width = 800
                                    self.Height = 400
                                    self.StartPosition = FormStartPosition.CenterParent
                                    self.dgv = DataGridView()
                                    self.dgv.Dock = DockStyle.Fill
                                    self.dgv.AllowUserToAddRows = False

                                    chk = DataGridViewCheckBoxColumn()
                                    chk.HeaderText = 'Apply'
                                    chk.Width = 50
                                    self.dgv.Columns.Add(chk)

                                    c_old = DataGridViewTextBoxColumn()
                                    c_old.HeaderText = 'Current Name'
                                    c_old.ReadOnly = True
                                    c_old.Width = 300
                                    self.dgv.Columns.Add(c_old)

                                    c_new = DataGridViewTextBoxColumn()
                                    c_new.HeaderText = 'Suggested Name'
                                    c_new.ReadOnly = True
                                    c_new.Width = 300
                                    self.dgv.Columns.Add(c_new)

                                    # Prefer a ComboBox column for safe action selection
                                    try:
                                        from System.Windows.Forms import DataGridViewComboBoxColumn
                                        c_action = DataGridViewComboBoxColumn()
                                        c_action.HeaderText = 'Action'
                                        c_action.Width = 160
                                        # Add items (strings) to the combo column
                                        try:
                                            c_action.Items.Add('Auto-suffix')
                                            c_action.Items.Add('Skip')
                                            c_action.Items.Add('Apply')
                                        except Exception:
                                            pass
                                        self.dgv.Columns.Add(c_action)
                                    except Exception:
                                        # fallback to text column if Combo not available
                                        c_action = DataGridViewTextBoxColumn()
                                        c_action.HeaderText = 'Action (Auto-suffix / Skip / Apply)'
                                        c_action.ReadOnly = False
                                        c_action.Width = 160
                                        self.dgv.Columns.Add(c_action)

                                    # small panel with Set All dropdown
                                    try:
                                        from System.Windows.Forms import FlowLayoutPanel, ComboBox as WinComboBox, Label as WinLabel
                                        panel = FlowLayoutPanel()
                                        panel.FlowDirection = 0
                                        panel.Dock = DockStyle.Top
                                        panel.Height = 30

                                        lbl = WinLabel()
                                        lbl.Text = 'Set all to:'
                                        panel.Controls.Add(lbl)

                                        self.set_all_combo = WinComboBox()
                                        try:
                                            self.set_all_combo.Items.Add('Auto-suffix')
                                            self.set_all_combo.Items.Add('Skip')
                                            self.set_all_combo.Items.Add('Apply')
                                        except Exception:
                                            pass
                                        try:
                                            # preselect from naming_rules if present
                                            rules_local = load_rules(RULES_FILE)
                                            default_conflict = rules_local.get('DEFAULT_CONFLICT_ACTION', 'Auto-suffix')
                                            if default_conflict in ['Auto-suffix', 'Skip', 'Apply']:
                                                self.set_all_combo.SelectedItem = default_conflict
                                            else:
                                                self.set_all_combo.SelectedItem = 'Auto-suffix'
                                        except Exception:
                                            try:
                                                self.set_all_combo.SelectedItem = 'Auto-suffix'
                                            except Exception:
                                                pass
                                        panel.Controls.Add(self.set_all_combo)

                                        self.btnSetAll = Button()
                                        self.btnSetAll.Text = 'Set All'
                                        self.btnSetAll.Height = 26
                                        self.btnSetAll.Click += self.on_set_all
                                        panel.Controls.Add(self.btnSetAll)

                                        # add the panel above the grid
                                        self.Controls.Add(panel)
                                    except Exception:
                                        self.set_all_combo = None

                                    for it in items:
                                        # default action is Auto-suffix
                                        idx = self.dgv.Rows.Add(True, it['old'], it['new'], 'Auto-suffix')

                                    self.btnOk = Button()
                                    self.btnOk.Text = 'OK'
                                    self.btnOk.Dock = DockStyle.Bottom
                                    self.btnOk.Height = 28
                                    self.btnOk.Click += self.on_ok

                                    self.btnCancel = Button()
                                    self.btnCancel.Text = 'Cancel'
                                    self.btnCancel.Dock = DockStyle.Bottom
                                    self.btnCancel.Height = 28
                                    self.btnCancel.Click += self.on_cancel

                                    self.Controls.Add(self.dgv)
                                    self.Controls.Add(self.btnOk)
                                    self.Controls.Add(self.btnCancel)

                                def on_ok(self, sender, args):
                                    try:
                                        # read actions back
                                        for i in range(self.dgv.Rows.Count):
                                            try:
                                                action = (self.dgv.Rows[i].Cells[3].Value or '').strip()
                                                if not action:
                                                    action = 'Auto-suffix'
                                                conflicts[i]['action'] = action
                                            except Exception:
                                                pass
                                    except Exception:
                                        pass
                                    try:
                                        # persist chosen default conflict action to naming_rules.json
                                        try:
                                            # load existing rules
                                            rls = load_rules(RULES_FILE) or {}
                                            # pick the most common action selected (simple heuristic)
                                            actions = []
                                            for i in range(self.dgv.Rows.Count):
                                                try:
                                                    a = (self.dgv.Rows[i].Cells[3].Value or '').strip()
                                                    if a:
                                                        actions.append(a)
                                                except Exception:
                                                    pass
                                            if actions:
                                                # choose the first action as default preference
                                                pref = actions[0]
                                                if pref in ['Auto-suffix', 'Skip', 'Apply']:
                                                    rls['DEFAULT_CONFLICT_ACTION'] = pref
                                                    try:
                                                        with open(RULES_FILE, 'w') as wf:
                                                            json.dump(rls, wf, indent=2)
                                                    except Exception:
                                                        pass
                                        except Exception:
                                            pass
                                        self.DialogResult = DialogResult.OK
                                    except Exception:
                                        pass
                                    try:
                                        self.Close()
                                    except Exception:
                                        pass

                                def on_cancel(self, sender, args):
                                    try:
                                        self.DialogResult = DialogResult.Cancel
                                    except Exception:
                                        pass
                                    try:
                                        self.Close()
                                    except Exception:
                                        pass

                                def on_set_all(self, sender, args):
                                    try:
                                        val = None
                                        try:
                                            val = self.set_all_combo.SelectedItem
                                        except Exception:
                                            try:
                                                val = (self.set_all_combo.Text or '')
                                            except Exception:
                                                val = None
                                        if not val:
                                            return
                                        for i in range(self.dgv.Rows.Count):
                                            try:
                                                self.dgv.Rows[i].Cells[3].Value = val
                                            except Exception:
                                                pass
                                    except Exception:
                                        pass

                            f = ConflictReviewForm(conflicts)
                            try:
                                dr = f.ShowDialog()
                            except Exception:
                                try:
                                    Application.Run(f)
                                    dr = None
                                except Exception:
                                    dr = None
                            # If dialog was cancelled, abort apply
                            try:
                                if dr == DialogResult.Cancel:
                                    return
                            except Exception:
                                pass
                        except Exception:
                            # fallback to global choice if dialog fails
                            res = _maybe_dialog(
                                'Detected %d potential name conflicts. Yes=Auto-suffix all, No=Skip conflicts, Cancel=Abort' % len(conflicts),
                                'MHT Family Namer - Conflicts',
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Warning,
                                default_result=DialogResult.No
                            )
                            if res is None:
                                # When dialogs are suppressed, default_result above should be returned.
                                # If the runtime returned None anyway, do the safest action: skip conflicts.
                                res = DialogResult.No
                            if res == DialogResult.Cancel:
                                return
                            auto_suffix = (res == DialogResult.Yes)
                    else:
                        auto_suffix = False
                except Exception:
                    auto_suffix = False

                # attempt rename inside transaction
                t = DB.Transaction(doc, 'MHT Family Namer - Apply Names')
                try:
                    # mark as applying so other UI actions can be blocked
                    globals()['_mht_is_applying'] = True
                    # disable apply button to avoid double-click / reentrancy
                    try:
                        # hide the form to avoid UI/WinForms race conditions while Revit performs renames
                        try:
                            self.Hide()
                        except Exception:
                            pass
                        self.btnApply.Enabled = False
                        self.btnExport.Enabled = False
                        self.btnRefresh.Enabled = False
                    except Exception:
                        pass
                    t.Start()
                    applied = 0
                    renamed_families = []
                    renamed_map = {}

                    # Build a lookup once: family name -> [Family]
                    fams_by_name = {}
                    try:
                        for x in DB.FilteredElementCollector(doc).OfClass(DB.Family):
                            try:
                                n = safe_get_name(x)
                                if n is None:
                                    continue
                                fams_by_name.setdefault(n, []).append(x)
                            except Exception:
                                continue
                    except Exception:
                        fams_by_name = {}
                    for fam_name, sug in to_apply:
                        # find family by name using safe_get_name
                        try:
                            fams = fams_by_name.get(fam_name, []) or []
                        except Exception:
                            fams = []
                        if not fams:
                            script.get_logger().warning('Family not found for rename row: {}'.format(fam_name))
                            continue

                        # Handle exact duplicate actions first
                        action = duplicate_actions.get(fam_name)
                        if action == "Delete":
                            if not is_family_in_use(doc, fam_name):
                                doc.Delete(fams[0].Id)
                                script.get_logger().info("Deleted duplicate family: {}".format(fam_name))
                            else:
                                script.get_logger().warning("Could not delete family '{}' because it is in use.".format(fam_name))
                            continue # Skip renaming
                        elif action == "Mark as Obsolete (X_ Prefix)":
                            try:
                                obsolete_name = "X_" + sug
                                fams[0].Name = obsolete_name
                                script.get_logger().info("Marked family as obsolete: {} -> {}".format(fam_name, obsolete_name))
                            except Exception as e:
                                script.get_logger().warning("Failed to mark family as obsolete '{}': {}".format(fam_name, e))
                            continue # Skip normal renaming

                        for fitem in fams:
                            try:
                                target_name = sug
                                # if auto_suffix and conflict exists, make a unique name
                                if auto_suffix:
                                    try:
                                        target_name = _make_unique(sug, existing)
                                    except Exception:
                                        target_name = sug
                                else:
                                    # if not auto-suffix and target conflicts, skip
                                    if sug and sug.strip().lower() in existing and (fam_name or '').strip().lower() != sug.strip().lower():
                                        script.get_logger().warning('Skipping rename of {} -> {} because name already exists'.format(fam_name, sug))
                                        continue
                                # Determine per-conflict action if present
                                try:
                                    action = 'Apply'
                                    for c in conflicts:
                                        try:
                                            if (c.get('old') or '').strip().lower() == (fam_name or '').strip().lower() and (c.get('new') or '').strip().lower() == (sug or '').strip().lower():
                                                action = (c.get('action') or 'Auto-suffix')
                                                break
                                        except Exception:
                                            pass
                                except Exception:
                                    action = 'Apply'

                                # Apply action logic
                                if action.lower() == 'skip':
                                    # do not rename this family
                                    script.get_logger().info('Skipping rename of {} -> {} (user choice)'.format(fam_name, sug))
                                    continue
                                if action.lower() == 'auto-suffix' or auto_suffix:
                                    try:
                                        target_name = _make_unique(sug, existing)
                                    except Exception:
                                        target_name = sug

                                # set the family name (Family.Name is writable)
                                fitem.Name = target_name
                                applied += 1
                                # remember renamed family (use the new name for filtering types)
                                renamed_families.append(target_name)
                                # map old -> new for type renamer
                                try:
                                    renamed_map[fam_name] = target_name
                                except Exception:
                                    pass
                                # ensure existing set includes the newly used name
                                try:
                                    existing.add((target_name or '').strip().lower())
                                except Exception:
                                    pass
                            except Exception as e:
                                script.get_logger().warning('Failed to rename {}: {}'.format(fam_name, e))
                    t.Commit()
                    script.get_logger().info('Applied {} family renames'.format(applied))

                    # Show an aggregated summary of what happened and offer to close all tool windows
                    try:
                        summary_lines = []
                        summary_lines.append('Family renames applied: %d' % applied)
                        # count skipped due to conflicts
                        skipped = 0
                        try:
                            for old, new in to_apply:
                                try:
                                    # if mapping doesn't contain the old family, it was skipped
                                    if (old not in renamed_map) and new:
                                        skipped += 1
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        summary_lines.append('Skipped (conflicts/choices): %d' % skipped)
                        summary_text = '\n'.join(summary_lines)
                        res = _maybe_dialog(
                            summary_text + '\n\nClose all tool windows now?',
                            'MHT Family Namer - Summary',
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information,
                            default_result=DialogResult.No
                        )
                        try:
                            if res == DialogResult.Yes:
                                # close modeless forms and review form
                                try:
                                    for f in list(globals().get('_open_modeless_forms') or []):
                                        try:
                                            f.Close()
                                        except Exception:
                                            pass
                                except Exception:
                                    pass
                                try:
                                    globals()['_open_modeless_forms'] = []
                                except Exception:
                                    pass
                                try:
                                    self.Close()
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    except Exception:
                        pass

                    # [OK] Ask user whether to rename types next (TypeNamer.md integration)
                    try:
                        result = _maybe_dialog(
                            "Family rename complete.\n\nDo you also want to rename the sub-types (family types)?",
                            "MHT Family Namer",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question,
                            default_result=DialogResult.No
                        )
                        if result == DialogResult.Yes:
                            script.get_logger().info('User chose to rename types.')
                            try:
                                # Hide this review form while the Type Renamer is open to avoid message-loop races
                                try:
                                    globals()['_mht_review_form_hidden_by_type'] = True
                                    try:
                                        self.Hide()
                                    except Exception:
                                        pass
                                except Exception:
                                    pass
                                show_type_renamer(doc, rules, include_families=renamed_families, name_map=renamed_map)
                            except Exception as e:
                                script.get_logger().warning('Failed to launch type renamer: {}'.format(e))
                        else:
                            _maybe_dialog(
                                "Family renaming complete.\n\nYou can close the tool.",
                                "MHT Family Namer",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information,
                                default_result=None
                            )
                    except Exception as e:
                        script.get_logger().warning('Prompt failed: {}'.format(e))
                except Exception as e:
                    try:
                        t.RollBack()
                    except Exception:
                        pass
                    script.get_logger().warning('Transaction failed: {}'.format(e))
                finally:
                    # clear applying flag and re-enable controls
                    try:
                        globals()['_mht_is_applying'] = False
                    except Exception:
                        pass
                    try:
                        # restore visibility and re-enable controls
                        try:
                            # Only show the review form again if it wasn't hidden by the type renamer
                            if not globals().get('_mht_review_form_hidden_by_type'):
                                self.Show()
                        except Exception:
                            pass
                        self.btnApply.Enabled = True
                        self.btnExport.Enabled = True
                        self.btnRefresh.Enabled = True
                    except Exception:
                        pass

            def on_close_all(self, sender, args):
                try:
                    res = _maybe_dialog(
                        'Close all open Family Renamer tool windows?',
                        'MHT Family Renamer',
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        default_result=DialogResult.No
                    )
                except Exception:
                    res = DialogResult.No
                try:
                    if res == DialogResult.Yes:
                        # attempt to close all modeless forms and this review form
                        try:
                            for f in list(globals().get('_open_modeless_forms') or []):
                                try:
                                    f.Close()
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        try:
                            globals()['_open_modeless_forms'] = []
                        except Exception:
                            pass
                        try:
                            # close this review form
                            self.Close()
                        except Exception:
                            pass
                except Exception:
                    pass

            def on_refresh(self, sender, args):
                # Refresh casing for the Suggested Name column (index 4). Safely update
                # combo cells by matching items case-insensitively when possible.
                mode = self.cbo.SelectedItem if self.cbo.SelectedItem else 'UPPER'
                for i in range(self.dgv.Rows.Count):
                    try:
                        cell = self.dgv.Rows[i].Cells[4]
                        cur = cell.Value
                        if cur is None:
                            continue
                        new = apply_casing(cur, mode)
                        try:
                            # If the cell exposes Items (ComboBox), try to find a case-insensitive
                            # match and assign that item. Otherwise assign the new string.
                            if hasattr(cell, 'Items'):
                                found = None
                                try:
                                    for itm in cell.Items:
                                        try:
                                            if str(itm).lower() == str(new).lower():
                                                found = itm
                                                break
                                        except Exception:
                                            pass
                                except Exception:
                                    found = None
                                if found is not None:
                                    cell.Value = found
                                else:
                                    # fallback: if Items has at least one entry, keep the first item
                                    try:
                                        if cell.Items and getattr(cell.Items, 'Count', 0) > 0:
                                            cell.Value = cell.Items[0]
                                        else:
                                            cell.Value = new
                                    except Exception:
                                        cell.Value = new
                            else:
                                cell.Value = new
                        except Exception:
                            pass
                    except Exception:
                        pass

            def on_select_all(self, sender, args):
                for i in range(self.dgv.Rows.Count):
                    try:
                        row = self.dgv.Rows[i]
                        if row.Visible:
                            row.Cells[0].Value = True
                    except Exception:
                        pass

            def on_clear_all(self, sender, args):
                for i in range(self.dgv.Rows.Count):
                    try:
                        row = self.dgv.Rows[i]
                        if row.Visible:
                            row.Cells[0].Value = False
                    except Exception:
                        pass

        form = ReviewForm(results, rules)
        try:
            # keep a global reference so other UI functions can hide/show it safely
            globals()['_mht_review_form'] = form
            globals()['_mht_review_form_hidden_by_type'] = False
        except Exception:
            pass
        try:
            try:
                Application.EnableVisualStyles()
            except Exception:
                pass
            # Show dialog (modal) to avoid starting a second message loop on the UI thread
            try:
                form.ShowDialog()
            except Exception:
                # fallback to Application.Run if ShowDialog not available
                Application.Run(form)
        except Exception:
            try:
                Application.Run(form)
            except Exception:
                pass
    except Exception as e:
        logger.warning('Failed to show review UI: {}'.format(e))


def is_family_in_use(doc, family_name):
    """Check if any instance of a family exists in the document."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance)
    for inst in collector:
        try:
            if inst.Symbol and inst.Symbol.Family:
                if inst.Symbol.Family.Name == family_name:
                    return True
        except Exception:
            continue
    return False




def show_type_renamer(doc, rules, include_families=None, name_map=None):
    """Open a second UI to rename all FamilySymbol (types)."""
    try:
        from System.Windows.Forms import Application, Form, DataGridView, DataGridViewCheckBoxColumn, DataGridViewTextBoxColumn, Button, DockStyle, FormStartPosition, Label, ComboBox
        from System.Drawing import Color
    except Exception:
        Application = None
        Form = object
        DataGridView = None
        DataGridViewCheckBoxColumn = None
        DataGridViewTextBoxColumn = None
        Button = None
        DockStyle = None
        FormStartPosition = None
        Label = None
        ComboBox = None
        Color = None


def handle_exact_duplicates(doc, to_apply, results):
    """Find and handle exact duplicate families before renaming.

    Returns (to_apply, actions) where actions is a dict mapping old_name -> action
    ("Keep", "Mark as Obsolete (X_ Prefix)", "Delete").
    """
    try:
        from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon
    except Exception:
        # UI not available; nothing to do
        return to_apply, {}

    potential_duplicates = {}
    for old_name, new_name in to_apply:
        if new_name not in potential_duplicates:
            potential_duplicates[new_name] = []
        potential_duplicates[new_name].append(old_name)

    exact_duplicates_groups = []
    for new_name, old_names in potential_duplicates.items():
        if len(old_names) > 1:
            # Gather infos for these families from results
            infos = [r['info'] for r in results if r['family'] in old_names]
            if not infos:
                continue
            first_info = infos[0]
            is_duplicate_group = True
            for other_info in infos[1:]:
                try:
                    if (first_info.get('category') != other_info.get('category') or
                        len(first_info.get('types', [])) != len(other_info.get('types', []))):
                        is_duplicate_group = False
                        break
                except Exception:
                    is_duplicate_group = False
                    break
            # further check: compute a lightweight fingerprint of family info (type names + params)
            try:
                def _family_fingerprint(info_obj):
                    try:
                        parts = []
                        parts.append((info_obj.get('category') or ''))
                        types = info_obj.get('types') or []
                        for t in sorted(types, key=lambda x: (x.get('type_name') or '')):
                            parts.append(t.get('type_name') or '')
                            params = t.get('params') or {}
                            for k in sorted(params.keys()):
                                parts.append('%s=%s' % (k, params.get(k) or ''))
                        inst = info_obj.get('instance_params') or {}
                        for k in sorted(inst.keys()):
                            parts.append('inst:%s=%s' % (k, inst.get(k) or ''))
                        import hashlib
                        h = hashlib.md5('\u001f'.join(parts).encode('utf-8')).hexdigest()
                        return h
                    except Exception:
                        return None
                fps = []
                for inf in infos:
                    try:
                        fps.append(_family_fingerprint(inf))
                    except Exception:
                        fps.append(None)
                identical_flag = (len(set(fps)) == 1)
            except Exception:
                identical_flag = False

            if is_duplicate_group:
                exact_duplicates_groups.append({'names': old_names, 'identical': identical_flag})

    if not exact_duplicates_groups:
        return to_apply, {}

    # Build a simple form to ask the user what to do with groups
    try:
        # Import UI elements lazily
        from System.Windows.Forms import Form, DataGridView, DataGridViewTextBoxColumn, DataGridViewComboBoxColumn, Button, DockStyle
        from System.Drawing import Color
    except Exception:
        return to_apply, {}

    class DuplicateHandlerForm(Form):
        def __init__(self, groups):
            self.Text = "Handle Exact Duplicate Families"
            self.Width = 800
            self.Height = 500
            self.StartPosition = FormStartPosition.CenterParent
            self.results = {}

            self.dgv = DataGridView()
            self.dgv.Dock = DockStyle.Fill
            self.dgv.AllowUserToAddRows = False
            self.dgv.ColumnHeadersHeightSizeMode = 3

            self.dgv.Columns.Add(DataGridViewTextBoxColumn())
            action_col = DataGridViewComboBoxColumn()
            action_col.HeaderText = "Action"
            try:
                action_col.Items.AddRange("Keep", "Mark as Obsolete (X_ Prefix)", "Delete")
            except Exception:
                pass
            action_col.Width = 200
            self.dgv.Columns.Add(action_col)
            self.dgv.Columns.Add(DataGridViewTextBoxColumn())

            for group in groups:
                try:
                    names = group.get('names') if isinstance(group, dict) else (group[0] if isinstance(group, (list, tuple)) else group)
                    identical = group.get('identical', False) if isinstance(group, dict) else False
                except Exception:
                    names = group
                    identical = False
                # Header row
                try:
                    row_idx = self.dgv.Rows.Add("--- Group for: %s ---" % (names[0] if names else 'Group'), None, None)
                    self.dgv.Rows[row_idx].ReadOnly = True
                    self.dgv.Rows[row_idx].DefaultCellStyle.BackColor = Color.Gray
                except Exception:
                    pass
                # Check usage in model and add rows
                usage = {name: is_family_in_use(doc, name) for name in names}
                kept_one = False
                for name in names:
                    is_used = usage.get(name, False)
                    status_text = "In Use" if is_used else "Not in Use"
                    # If the group was detected as identical and the family is not used, default to marking it obsolete
                    if identical and not is_used:
                        action = "Mark as Obsolete (X_ Prefix)"
                    else:
                        action = "Keep" if (is_used and not kept_one) else "Mark as Obsolete (X_ Prefix)"
                    if is_used:
                        kept_one = True
                    try:
                        row_idx = self.dgv.Rows.Add(name, action, status_text)
                        # If this family is in use, prevent deletion by removing 'Delete' from the combo cell
                        try:
                            cell = self.dgv.Rows[row_idx].Cells[1]
                            try:
                                # DataGridViewComboBoxCell exposes Items; remove 'Delete' when present
                                if hasattr(cell, 'Items'):
                                    try:
                                        cell.Items.Remove('Delete')
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                        except Exception:
                            pass
                    except Exception:
                        pass

            btn_ok = Button(Text="Confirm Actions", Dock=DockStyle.Bottom, Height=30)
            btn_ok.Click += self.on_ok
            self.Controls.Add(self.dgv)
            self.Controls.Add(btn_ok)

        def on_ok(self, sender, args):
            try:
                for row in self.dgv.Rows:
                    try:
                        if not row.IsNewRow and row.Cells[0].Value and "---" not in row.Cells[0].Value:
                            self.results[row.Cells[0].Value] = row.Cells[1].Value
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                self.DialogResult = DialogResult.OK
            except Exception:
                pass
            try:
                self.Close()
            except Exception:
                pass

    try:
        dup_form = DuplicateHandlerForm(exact_duplicates_groups)
        dr = dup_form.ShowDialog()
        if dr == DialogResult.OK:
            return to_apply, dup_form.results
    except Exception:
        pass

    return to_apply, {}

    # Gather type info
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
    results = []

    templates = rules.get('TEMPLATES', {})
    for sym in collector:
        try:
            fam = sym.Family
            fam_name = safe_get_name(fam)
            # if include_families provided, skip types not belonging to renamed families
            try:
                if include_families and fam_name not in include_families:
                    continue
            except Exception:
                pass
            cat_name = family_primary_category_name(fam)
            type_name = safe_get_name(sym)

            # Suggest name using family-based template
            rule = templates.get(cat_name, templates.get('Default', ''))
            if rule:
                info = gather_family_info(fam, None)
                # Override type name for current symbol
                info['types'] = [{'type_name': type_name, 'params': {}}]
                suggestion = apply_template(rule, info, rules)
                # apply name_map replacements (old->new) to suggestion so it prefers new family prefixes
                try:
                    if name_map:
                            import re as _re
                            for old, new in name_map.items():
                                try:
                                    # Replace only when the old family name appears at the start
                                    # of a token (start of string or after a delimiter like - _ or space).
                                    # This avoids accidental replacements in the middle of words.
                                    pattern = r'(^|(?<=[\-\_\s))' + _re.escape(old)
                                    # use a function replacement to preserve the leading delimiter if present
                                    def _repl(m):
                                        prefix = m.group(1) or ''
                                        return (prefix + new)
                                    suggestion = _re.sub(pattern, _repl, suggestion, flags=_re.I)
                                except Exception:
                                    try:
                                        # fallback: only replace if suggestion starts with old (case-insensitive)
                                        if suggestion and suggestion.lower().startswith(old.lower()):
                                            suggestion = new + suggestion[len(old):]
                                        else:
                                            suggestion = suggestion.replace(old, new)
                                    except Exception:
                                        pass
                except Exception:
                    pass
            else:
                suggestion = type_name
            results.append({'family': fam_name, 'type': type_name, 'category': cat_name, 'suggested': suggestion})
        except Exception:
            continue

    class TypeRenameForm(Form):
        def __init__(self, rows):
            self.Text = 'MHT Type Renamer - Review Type Names'
            self.Width = 1200
            self.Height = 800
            self.MinimumSize = Size(800, 600)
            self.StartPosition = FormStartPosition.CenterParent
            self.AutoScaleDimensions = SizeF(6, 13)
            self.AutoScaleMode = AutoScaleMode.Font

            # Professional dark theme colors matching ReviewForm
            self.bg_dark = Color.FromArgb(30, 30, 30)
            self.bg_control = Color.FromArgb(45, 45, 45)
            self.bg_header = Color.FromArgb(20, 20, 20)
            self.fg_text = Color.FromArgb(220, 220, 220)
            self.accent = Color.FromArgb(0, 120, 212)
            self.accent_dark = Color.FromArgb(0, 90, 160)

            # Dark theme for the type renamer form
            self.BackColor = self.bg_dark
            self.ForeColor = self.fg_text

            self.dgv = DataGridView()
            self.dgv.Dock = DockStyle.Fill
            self.dgv.AllowUserToAddRows = False
            try:
                self.dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            except Exception:
                # Fallback: Leave default to avoid conversion errors
                pass
            try:
                self.dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
            except Exception:
                pass
            self.dgv.AllowUserToResizeColumns = True
            self.dgv.AllowUserToResizeRows = False
            
            # Dark theme for DataGridView
            self.dgv.BackgroundColor = self.bg_dark
            self.dgv.ForeColor = self.fg_text
            self.dgv.GridColor = self.bg_control
            self.dgv.DefaultCellStyle.BackColor = self.bg_dark
            self.dgv.DefaultCellStyle.ForeColor = self.fg_text
            self.dgv.DefaultCellStyle.SelectionBackColor = self.accent
            self.dgv.DefaultCellStyle.SelectionForeColor = Color.White
            self.dgv.ColumnHeadersDefaultCellStyle.BackColor = self.bg_header
            self.dgv.ColumnHeadersDefaultCellStyle.ForeColor = self.fg_text
            self.dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = self.accent_dark
            self.dgv.EnableHeadersVisualStyles = False
            self.dgv.RowHeadersDefaultCellStyle.BackColor = self.bg_header
            self.dgv.RowHeadersDefaultCellStyle.ForeColor = self.fg_text

            chk = DataGridViewCheckBoxColumn()
            chk.HeaderText = 'Apply'
            chk.Width = 50
            self.dgv.Columns.Add(chk)

            c1 = DataGridViewTextBoxColumn()
            c1.HeaderText = 'Family'
            c1.ReadOnly = True
            c1.Width = 250
            self.dgv.Columns.Add(c1)

            c2 = DataGridViewTextBoxColumn()
            c2.HeaderText = 'Type'
            c2.ReadOnly = True
            c2.Width = 250
            self.dgv.Columns.Add(c2)

            c3 = DataGridViewTextBoxColumn()
            c3.HeaderText = 'Suggested Name'
            c3.ReadOnly = False
            c3.Width = 300
            self.dgv.Columns.Add(c3)

            for r in rows:
                row_idx = self.dgv.Rows.Add(False, r['family'], r['type'], r['suggested'])
                if Color is not None and r['type'] and r['type'].lower() == (r['suggested'] or '').lower():
                    try:
                        for c in range(self.dgv.Columns.Count):
                            self.dgv.Rows[row_idx].Cells[c].Style.BackColor = Color.LightYellow
                    except Exception:
                        pass

            # exclude DUP checkbox
            try:
                from System.Windows.Forms import CheckBox
                self.chkExcludeDup = CheckBox()
                self.chkExcludeDup.Text = 'Exclude DUP'
                self.chkExcludeDup.Checked = False
                self.chkExcludeDup.AutoSize = True
                self.chkExcludeDup.Dock = DockStyle.Bottom
                self.chkExcludeDup.ForeColor = SystemColors.ControlLightLight
                self.chkExcludeDup.CheckedChanged += self.on_exclude_dup_changed
                self.Controls.Add(self.chkExcludeDup)
            except Exception:
                self.chkExcludeDup = None

            # selection controls for types
            self.btnSelectAll = Button()
            self.btnSelectAll.Text = 'Select All'
            self.btnSelectAll.Dock = DockStyle.Bottom
            self.btnSelectAll.Height = 24
            self.btnSelectAll.Click += self.on_select_all

            self.btnClearAll = Button()
            self.btnClearAll.Text = 'Clear All'
            self.btnClearAll.Dock = DockStyle.Bottom
            self.btnClearAll.Height = 24
            self.btnClearAll.Click += self.on_clear_all

            self.btnApply = Button()
            self.btnApply.Text = 'Apply Selected'
            self.btnApply.Dock = DockStyle.Bottom
            self.btnApply.Height = 30
            self.btnApply.Click += self.on_apply

            # Close All button: confirm and close both type and review forms
            self.btnCloseAll = Button()
            self.btnCloseAll.Text = 'Close All'
            self.btnCloseAll.Dock = DockStyle.Bottom
            self.btnCloseAll.Height = 30
            self.btnCloseAll.Click += self.on_close_all

            self.Controls.Add(self.dgv)
            self.Controls.Add(self.btnApply)
            self.Controls.Add(self.btnCloseAll)
            self.Controls.Add(self.btnClearAll)
            self.Controls.Add(self.btnSelectAll)

            # Dark theme for buttons
            for ctrl in [self.btnApply, self.btnCloseAll, self.btnClearAll, self.btnSelectAll]:
                try:
                    ctrl.BackColor = self.bg_control
                    ctrl.ForeColor = self.fg_text
                    ctrl.FlatStyle = FlatStyle.Flat
                    ctrl.FlatAppearance.BorderColor = self.accent
                    ctrl.FlatAppearance.MouseDownBackColor = self.accent_dark
                    ctrl.FlatAppearance.MouseOverBackColor = self.accent
                except Exception:
                    pass

        def on_apply(self, sender, args):
            to_apply = []
            for i in range(self.dgv.Rows.Count):
                try:
                    if self.dgv.Rows[i].Cells[0].Value:
                        fam_name = self.dgv.Rows[i].Cells[1].Value
                        cur_type = self.dgv.Rows[i].Cells[2].Value
                        sug = self.dgv.Rows[i].Cells[3].Value
                        to_apply.append((fam_name, cur_type, sug))
                except Exception:
                    continue
            if not to_apply:
                script.get_logger().info('No types selected to rename')
                return
            # pre-check for name conflicts per family
            try:
                # build map of existing names per family
                fams = {}
                for fam_name, cur_type, sug in to_apply:
                    k = (fam_name or '').strip()
                    if k.lower() not in fams:
                        fams[k.lower()] = _collect_existing_type_names(doc, k)
                conflicts = []
                for fam_name, cur_type, sug in to_apply:
                    try:
                        if sug and sug.strip().lower() in fams.get((fam_name or '').strip().lower(), set()) and (cur_type or '').strip().lower() != sug.strip().lower():
                            conflicts.append((fam_name, cur_type, sug))
                    except Exception:
                        pass
                if conflicts:
                    try:
                        res = _maybe_dialog(
                            'Detected %d potential type name conflicts. Yes=Auto-suffix, No=Skip conflicts, Cancel=Abort' % len(conflicts),
                            'MHT Type Renamer - Conflicts',
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning,
                            default_result=DialogResult.No
                        )
                    except Exception:
                        res = DialogResult.No
                    if res is None:
                        res = DialogResult.No
                    if res == DialogResult.Cancel:
                        return
                    auto_suffix = (res == DialogResult.Yes)
                else:
                    auto_suffix = False
            except Exception:
                auto_suffix = False
            # prevent conflicting apply actions
            try:
                if globals().get('_mht_is_applying'):
                    _show_error_dialog('A family apply operation is in progress. Wait until it finishes before renaming types.', 'MHT Type Renamer')
                    return
            except Exception:
                pass

            t = DB.Transaction(doc, 'MHT Type Renamer - Apply Type Names')
            try:
                globals()['_mht_is_applying'] = True
                # hide the form to avoid UI/WinForms race conditions while Revit performs renames
                try:
                    self.Hide()
                except Exception:
                    pass
                # disable type apply button while applying
                try:
                    self.btnApply.Enabled = False
                except Exception:
                    pass
                t.Start()
                applied = 0
                skipped = 0
                rename_details = []  # Track details of what happened

                # First, build a mapping of all family types for efficient lookup
                family_types = {}
                collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
                for s in collector:
                    try:
                        fam = s.Family
                        fam_name = safe_get_name(fam)
                        type_name = safe_get_name(s)
                        if fam_name not in family_types:
                            family_types[fam_name] = []
                        family_types[fam_name].append((type_name, s))
                    except Exception:
                        continue

                # Process the renames
                for fam_name, type_name, sug in to_apply:
                    if fam_name in family_types:
                        matched = False
                        for current_type_name, symbol in family_types[fam_name]:
                            if current_type_name == type_name:
                                matched = True
                                target_name = sug
                                try:
                                    if auto_suffix:
                                        # ensure we have an existing set for this family
                                        ex = fams.get((fam_name or '').strip().lower(), set())
                                        target_name = _make_unique(sug, ex)
                                    else:
                                        # if conflict exists and not auto-suffix, skip
                                        ex = fams.get((fam_name or '').strip().lower(), set())
                                        if sug and sug.strip().lower() in ex and (type_name or '').strip().lower() != sug.strip().lower():
                                            reason = 'Name conflict'
                                            skipped += 1
                                            rename_details.append({
                                                'family': fam_name,
                                                'type': type_name,
                                                'suggested': sug,
                                                'status': 'Skipped',
                                                'reason': reason
                                            })
                                            script.get_logger().warning('Skipping type rename {} -> {} for family {} because name exists'.format(type_name, sug, fam_name))
                                            continue
                                except Exception:
                                    target_name = sug

                                try:
                                    symbol.Name = target_name
                                    applied += 1
                                    rename_details.append({
                                        'family': fam_name,
                                        'type': type_name,
                                        'suggested': target_name,
                                        'status': 'Renamed',
                                        'reason': ''
                                    })
                                except Exception as e:
                                    skipped += 1
                                    rename_details.append({
                                        'family': fam_name,
                                        'type': type_name,
                                        'suggested': target_name,
                                        'status': 'Failed',
                                        'reason': str(e)
                                    })
                        
                        if not matched:
                            skipped += 1
                            rename_details.append({
                                'family': fam_name,
                                'type': type_name,
                                'suggested': sug,
                                'status': 'Not Found',
                                'reason': 'Type not found in family'
                            })

                t.Commit()

                # Show detailed results dialog
                try:
                    summary = []
                    summary.append("Type Renaming Results:")
                    summary.append("------------------------")
                    summary.append("Total selected: {}".format(len(to_apply)))
                    summary.append("Successfully renamed: {}".format(applied))
                    summary.append("Skipped/Failed: {}".format(skipped))
                    summary.append("\nDetailed Results:")
                    summary.append("------------------------")
                    
                    for detail in rename_details:
                        status = detail['status']
                        if status == 'Renamed':
                            summary.append(u"[OK] {} : {} -> {}".format(detail['family'], detail['type'], detail['suggested']))
                        else:
                            reason = " ({})".format(detail['reason']) if detail['reason'] else ""
                            summary.append(u"[FAIL] {} : {} -> {} - {}{}".format(detail['family'], detail['type'], detail['suggested'], status, reason))

                    _maybe_dialog(
                        '\n'.join(summary),
                        'Type Renaming Results',
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        default_result=None
                    )
                except Exception as e:
                    script.get_logger().warning('Failed to show results dialog: {}'.format(e))

                script.get_logger().info('Type renaming complete. {} renamed, {} skipped'.format(applied, skipped))
            except Exception as e:
                try:
                    t.RollBack()
                except Exception:
                    pass
                script.get_logger().warning('Transaction failed: {}'.format(e))
            finally:
                try:
                    globals()['_mht_is_applying'] = False
                except Exception:
                    pass
                try:
                    # restore visibility and re-enable controls
                    try:
                        self.Show()
                    except Exception:
                        pass
                    self.btnApply.Enabled = True
                    try:
                        self.btnSelectAll.Enabled = True
                        self.btnClearAll.Enabled = True
                    except Exception:
                        pass
                except Exception:
                    pass

        def on_close_all(self, sender, args):
            try:
                res = _maybe_dialog(
                    'Close all open Family Renamer tool windows?',
                    'MHT Type Renamer',
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    default_result=DialogResult.No
                )
            except Exception:
                res = DialogResult.No
            try:
                if res == DialogResult.Yes:
                    # close any modeless forms
                    try:
                        for f in list(globals().get('_open_modeless_forms') or []):
                            try:
                                f.Close()
                            except Exception:
                                pass
                    except Exception:
                        pass
                    try:
                        globals()['_open_modeless_forms'] = []
                    except Exception:
                        pass
                    # close the review form if present
                    try:
                        rf = globals().get('_mht_review_form')
                        if rf:
                            try:
                                rf.Close()
                            except Exception:
                                pass
                    except Exception:
                        pass
                    # close this type renamer
                    try:
                        self.Close()
                    except Exception:
                        pass
            except Exception:
                pass

        def on_select_all(self, sender, args):
            for i in range(self.dgv.Rows.Count):
                row = self.dgv.Rows[i]
                try:
                    if row.Visible:
                        row.Cells[0].Value = True
                except Exception:
                    pass

        def on_clear_all(self, sender, args):
            for i in range(self.dgv.Rows.Count):
                row = self.dgv.Rows[i]
                try:
                    if row.Visible:
                        row.Cells[0].Value = False
                except Exception:
                    pass

        def on_exclude_dup_changed(self, sender, args):
            try:
                exclude = self.chkExcludeDup.Checked
                for i in range(self.dgv.Rows.Count):
                    try:
                        row = self.dgv.Rows[i]
                        current_name = (row.Cells[2].Value or '').lower()
                        suggested_name = (row.Cells[3].Value or '').lower()
                        
                        # Check if either name contains 'dup'
                        has_dup = 'dup' in current_name or 'dup' in suggested_name
                        row.Visible = not (exclude and has_dup)
                    except Exception:
                        pass
            except Exception:
                pass

    form = TypeRenameForm(results)
    try:
        # Enable visual styles if available
        try:
            Application.EnableVisualStyles()
        except Exception:
            pass

        # Show modelessly so both windows can remain open side-by-side. Keep a global
        # reference to the form to prevent it being garbage collected by IronPython.
        try:
            # global holder for modeless forms
            if not globals().get('_open_modeless_forms'):
                globals()['_open_modeless_forms'] = []
        except Exception:
            globals()['_open_modeless_forms'] = []

        globals()['_open_modeless_forms'].append(form)

        # If a Review form exists, hide it while the Type Renamer is open to avoid UI message-loop races
        try:
            review = globals().get('_mht_review_form')
            if review:
                try:
                    globals()['_mht_review_form_hidden_by_type'] = True
                    review.Hide()
                except Exception:
                    pass
        except Exception:
            pass

        # when the form closes, remove it from the global list
        try:
            def _on_closed(sender, evt):
                try:
                    globals()['_open_modeless_forms'].remove(sender)
                except Exception:
                    pass
                # When the type renamer closes, re-show the review form if it was hidden
                try:
                    review = globals().get('_mht_review_form')
                    if review and globals().get('_mht_review_form_hidden_by_type'):
                        try:
                            review.Show()
                        except Exception:
                            pass
                        try:
                            globals()['_mht_review_form_hidden_by_type'] = False
                        except Exception:
                            pass
                except Exception:
                    pass
            form.Closed += _on_closed
        except Exception:
            pass

        # Show modelessly. If Show isn't available, fallback to ShowDialog for safety.
        try:
            form.Show()
        except Exception:
            try:
                form.ShowDialog()
            except Exception:
                try:
                    Application.Run(form)
                except Exception:
                    pass
    except Exception:
        # final fallback: try Application.Run
        try:
            Application.Run(form)
        except Exception:
            pass


if __name__ == '__main__':
    main()
