# -*- coding: utf-8 -*-
__title__ = "GENERATE"
__helpurl__ = ""
__doc__ = """Version = 1.0
Date: 19.03.2026
Author: GM
_____________________________________________________________________
Description:
Use GENERATE to create new sheets with views from Excel data.
Creates parent views and sub-views for each zoning.
_____________________________________________________________________
How-to:

GENERATE Form:
-> Select Excel file
-> Drag Views from List to any SheetCard
-> You Can Drag Views between SheetCards as well.
-> Click on + Symbol to create a new SheetCard
-> Right-Click to remove views from SheetCards
-> Right-Click Views in list to duplicate them

Report Form:
-> Click on Cards in Report to open Sheets/Views 
_____________________________________________________________________
Features:
- Excel Import for sheet data
- Automatic zoning-based sub-view creation
- Dark theme UI
_____________________________________________________________________
"""


# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
from Autodesk.Revit.DB import *
from pyrevit import forms
import wpf, os, clr
import re
from System import Type, Activator
from System.Runtime.InteropServices import Marshal
from System.Reflection import BindingFlags
clr.AddReference("System.Windows.Forms")

# WPF Imports
clr.AddReference("System")
from System.Windows import Window, Visibility, HorizontalAlignment, VerticalAlignment, CornerRadius, Thickness
from System.Windows.Window import DragMove
from System.Windows.Controls import Orientation, CheckBox, DockPanel, Button,ComboBoxItem, TextBox, ListBoxItem, StackPanel, TextBlock, WrapPanel, Border, ScrollViewer
from System.Windows.Input import MouseButtonState, Keyboard, ModifierKeys
from System.Windows.Media import VisualTreeHelper, SolidColorBrush, Colors, SolidColorBrush, ColorConverter, Brushes
from System.Diagnostics.Process import Start
from System import Uri

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
PATH_SCRIPT = os.path.dirname(__file__)
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application
doc         = __revit__.ActiveUIDocument.Document
exit = False


# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝
#==================================================

def exitscript():
    output.print_md('**Error has occurred**')
    output.print_md(
        '**Please share the error message with me on [GitHub/EF-Tools:Issues](https://github.com/ErikFrits/EF-Tools/issues)**')
    print('\n Error Message:')
    import traceback, sys
    print(traceback.format_exc())
    sys.exit()

def rename_sheet(sheet, sheet_name, sheet_number):
    """Renames a ViewSheet with the specified SheetName and SheetNumber avoiding duplicates."""
    # Clear forbidden symbols
    forbidden_symbols = "\:{}[]|;<>?`~"
    sheet_name   = ''.join(c for c in sheet_name if c not in forbidden_symbols)
    sheet_number = ''.join(c for c in sheet_number if c not in forbidden_symbols)

    # Change SheetName
    sheet.Name = sheet_name

    # Change SheetNumber
    for i in range(1, 50):
        try:
            sheet.SheetNumber = sheet_number
            break
        except:
            sheet_number += '*'

def _com_get(com_obj, name):
    try:
        return getattr(com_obj, name)
    except Exception:
        return com_obj.GetType().InvokeMember(name, BindingFlags.GetProperty, None, com_obj, None)

def _com_call(com_obj, name, *args):
    try:
        member = getattr(com_obj, name)
        return member(*args)
    except Exception:
        return com_obj.GetType().InvokeMember(name, BindingFlags.InvokeMethod, None, com_obj, tuple(args))

def _com_set(com_obj, name, value):
    try:
        setattr(com_obj, name, value)
        return
    except Exception:
        return com_obj.GetType().InvokeMember(name, BindingFlags.SetProperty, None, com_obj, (value,))

def _com_item(com_obj, index):
    try:
        return com_obj[index]
    except Exception:
        pass
    try:
        return com_obj.Item[index]
    except Exception:
        pass
    try:
        return com_obj.Item(index)
    except Exception:
        return com_obj.GetType().InvokeMember("Item", BindingFlags.GetProperty, None, com_obj, (index,))

def _com_item2(com_obj, index1, index2):
    try:
        return com_obj[index1, index2]
    except Exception:
        pass
    try:
        return com_obj.Item[index1, index2]
    except Exception:
        pass
    try:
        return com_obj.Item(index1, index2)
    except Exception:
        return com_obj.GetType().InvokeMember("Item", BindingFlags.GetProperty, None, com_obj, (index1, index2))

def _release_com_object(com_obj):
    if com_obj is not None:
        try:
            Marshal.ReleaseComObject(com_obj)
        except:
            pass

def _open_excel_workbook(file_path):
    excel_type = Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise RuntimeError("Microsoft Excel is not installed or COM registration is unavailable.")

    excel_app = Activator.CreateInstance(excel_type)
    _com_set(excel_app, "Visible", False)
    _com_set(excel_app, "DisplayAlerts", False)
    workbooks = _com_get(excel_app, "Workbooks")
    workbook = _com_call(workbooks, "Open", file_path)
    _release_com_object(workbooks)
    return excel_app, workbook

def _close_excel_workbook(excel_app, workbook):
    try:
        if workbook is not None:
            _com_call(workbook, "Close", False)
    except Exception:
        pass

    try:
        if excel_app is not None:
            _com_call(excel_app, "Quit")
    except Exception:
        pass

    _release_com_object(workbook)
    _release_com_object(excel_app)

def _get_excel_worksheet(workbook, sheet_name):
    sheets = _com_get(workbook, "Sheets")
    worksheet = None
    try:
        sheet_name_text = str(sheet_name).strip()
        try:
            worksheet = _com_item(sheets, sheet_name_text)
        except Exception:
            worksheet = None

        if worksheet is None:
            sheet_count = int(_com_get(sheets, "Count"))
            target_name = sheet_name_text.lower()
            for index in range(1, sheet_count + 1):
                candidate = _com_item(sheets, index)
                candidate_name = str(_com_get(candidate, "Name")).strip()
                if candidate_name == sheet_name_text or candidate_name.lower() == target_name:
                    worksheet = candidate
                    break
                _release_com_object(candidate)
    finally:
        _release_com_object(sheets)
    return worksheet

def _get_excel_sheet_names(file_path):
    excel_app = None
    workbook = None
    sheets = None
    names = []

    try:
        excel_app, workbook = _open_excel_workbook(file_path)
        sheets = _com_get(workbook, "Sheets")
        sheet_count = int(_com_get(sheets, "Count"))
        for index in range(1, sheet_count + 1):
            sheet_obj = None
            try:
                sheet_obj = _com_item(sheets, index)
                names.append(str(_com_get(sheet_obj, "Name")))
            finally:
                _release_com_object(sheet_obj)
        return names
    finally:
        _release_com_object(sheets)
        _close_excel_workbook(excel_app, workbook)

def _get_matrix_value(values, row_index_1, col_index_1):
    """Safely read Value2 from Excel COM return shapes (1-based and 0-based SAFEARRAYs).

    Excel normally returns a 1-based System.Array[,] via COM, but IronPython's
    array subscript `values[r, c]` calls GetValue(r, c) using ABSOLUTE indices.
    For a 1-based array [1..N, 1..M] that is correct.  For a 0-based array
    (some COM marshalling paths return [0..N-1, 0..M-1]) the same call reads
    the wrong cell.  Using GetLowerBound lets us detect and compensate.
    """
    # Primary: use GetLowerBound to handle both 1-based and 0-based SAFEARRAYs.
    try:
        r_lb = values.GetLowerBound(0)
        c_lb = values.GetLowerBound(1)
        return values.GetValue(row_index_1 - 1 + r_lb, col_index_1 - 1 + c_lb)
    except Exception:
        pass

    # Fallback A: direct Python subscript (works for 1-based arrays where [r,c] == GetValue(r,c)).
    try:
        return values[row_index_1, col_index_1]
    except Exception:
        pass

    # Fallback B: explicit 0-indexed GetValue.
    try:
        return values.GetValue(row_index_1 - 1, col_index_1 - 1)
    except Exception:
        pass

    # Fallback C: treat as jagged (1-D of 1-D).
    try:
        row = values[row_index_1 - 1]
        return row[col_index_1 - 1]
    except Exception:
        pass

    # Last resort: single-cell range returns the scalar directly.
    if row_index_1 == 1 and col_index_1 == 1:
        return values

    return None

def _read_excel_data(file_path, sheet_name=None):
    excel_app = None
    workbook = None
    worksheet = None
    used_range = None
    rows = None
    cols = None

    try:
        excel_app, workbook = _open_excel_workbook(file_path)
        target_sheet = sheet_name
        if not target_sheet:
            sheets = _com_get(workbook, "Sheets")
            first_sheet = _com_item(sheets, 1)
            target_sheet = str(_com_get(first_sheet, "Name"))
            _release_com_object(first_sheet)
            _release_com_object(sheets)

        worksheet = _get_excel_worksheet(workbook, target_sheet)
        if worksheet is None:
            raise ValueError("Worksheet '{}' not found in Excel file.".format(target_sheet))

        used_range = _com_get(worksheet, "UsedRange")
        values = _com_get(used_range, "Value2")

        # Get dimensions
        rows = _com_get(used_range, "Rows")
        cols = _com_get(used_range, "Columns")
        row_count = int(_com_get(rows, "Count"))
        col_count = int(_com_get(cols, "Count"))

        data = []
        for r in range(1, min(row_count + 1, 1000)):  # Limit to 1000 rows
            row_data = []
            for c in range(1, min(col_count + 1, 10)):  # Limit to 10 columns
                value = _get_matrix_value(values, r, c)
                row_data.append(str(value).strip() if value is not None else "")
            data.append(row_data)

        return data

    finally:
        _release_com_object(rows)
        _release_com_object(cols)
        _release_com_object(used_range)
        _release_com_object(worksheet)
        _close_excel_workbook(excel_app, workbook)

def extract_zoning(sheet_number):
    """Extract zoning from sheet number, e.g., '521' from 'NWB-MHT-OPH-ME-212-05-1521'"""
    # Simple extraction: last 3 digits before last dash or something. Adjust as needed.
    parts = sheet_number.split('-')
    if len(parts) > 1:
        return parts[-1][:3]  # Assume last part, first 3 chars
    return "000"

def _to_text(value):
    if value is None:
        return ""
    return str(value).strip()

def _normalize_spaces(value_text):
    """Collapse repeated whitespace so parser/grouping is not broken by Excel spacing noise."""
    text = _to_text(value_text)
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()

def _short_sheet_number(value_text):
    """Convert long sheet numbers to the requested short form.

    Example:
    NWB-MHT-OPH-ME-212-GR-1511 -> 212-GR-1511
    """
    text = _normalize_spaces(value_text)
    if not text:
        return ""
    parts = [p for p in text.split("-") if p]
    if len(parts) >= 3:
        return "-".join(parts[-3:])
    return text

def _looks_like_sheet_number(value_text):
    text = _to_text(value_text)
    if not text:
        return False
    text_upper = text.upper()

    if _is_likely_scale_text(text):
        return False

    # Reject obvious view-title phrases that are often misread as numbers.
    if "LEVEL" in text_upper or "ZONE" in text_upper or "LAYOUT" in text_upper:
        return False

    # Sheet numbers are typically compact codes; long spaced phrases are likely titles.
    if text.count(" ") > 2:
        return False

    # Example: NWB-MHT-OPH-ME-212-GR-1511, but also shorter codes like GF-101, ME-01.
    if text.count("-") >= 2:
        return True

    # Single-dash codes like E-1, A-01 are also valid sheet numbers.
    if text.count("-") == 1:
        parts = text.split("-")
        # Both sides must be non-empty and at least one side must contain a digit.
        if all(parts) and any(ch.isdigit() for ch in text):
            return True

    # Also allow purely numeric sheet numbers such as 101, 2001, etc.
    if text.isdigit() and len(text) >= 2:
        return True

    # Fallback: alphanumeric code that has both letters and digits, compact (no long prose).
    has_letter = any(ch.isalpha() for ch in text)
    has_digit = any(ch.isdigit() for ch in text)
    return has_letter and has_digit and len(text) >= 3

def _find_header_layout(raw_rows):
    """Find header row and column indexes for sheet number/name/scale.

    Recognises all common header wordings:
      - Sheet Number, Document Number  → doc_col
      - Sheet Name, View Name, Title   → title_col
      - Scale                          → scale_col
    """
    scan_count = min(len(raw_rows), 40)
    for row_index in range(scan_count):
        row = raw_rows[row_index]
        doc_col = None
        title_col = None
        scale_col = None
        matched_header_count = 0
        for col_index, cell in enumerate(row):
            cell_text = _to_text(cell).upper()
            if not cell_text:
                continue
            # Sheet/document number column
            if "NUMBER" in cell_text:
                doc_col = col_index
                matched_header_count += 1
            # Sheet name / view name / title column (must NOT also be a number col)
            elif "NAME" in cell_text or "TITLE" in cell_text:
                title_col = col_index
                matched_header_count += 1
            # Scale column
            elif "SCALE" in cell_text:
                scale_col = col_index
                matched_header_count += 1

        if doc_col is not None and title_col is not None:
            return {
                "header_row": row_index,
                "doc_col": doc_col,
                "title_col": title_col,
                "scale_col": scale_col,
                "matched_header_count": matched_header_count,
            }

    # Fallback to column-order assumption when headers are missing or merged.
    return {
        "header_row": 0,
        "doc_col": 0,
        "title_col": 1,
        "scale_col": 2,
        "matched_header_count": 0,
    }

def _infer_layout_from_data(raw_rows, start_index, max_cols=10):
    """Infer doc/title/scale columns when headers are unreliable due formatting/merged cells."""
    row_slice = raw_rows[start_index:start_index + 120]
    if not row_slice:
        return {"doc_col": 0, "title_col": 1, "scale_col": 2}

    col_scores_doc = [0] * max_cols
    col_scores_zone = [0] * max_cols
    col_scores_nonempty = [0] * max_cols
    col_scores_scale = [0] * max_cols

    for row in row_slice:
        for c in range(max_cols):
            val = _to_text(row[c]) if c < len(row) else ""
            if not val:
                continue
            col_scores_nonempty[c] += 1
            if _looks_like_sheet_number(val):
                col_scores_doc[c] += 1
            upper_val = val.upper()
            if "ZONE" in upper_val:
                col_scores_zone[c] += 1
            if "PLAN" in upper_val and ":" in upper_val:
                col_scores_scale[c] += 1

    doc_col = max(range(max_cols), key=lambda i: col_scores_doc[i])

    # Prefer title columns to the right of doc column with zone-rich strings.
    title_candidates = [i for i in range(doc_col + 1, max_cols)]
    if not title_candidates:
        title_candidates = [i for i in range(max_cols) if i != doc_col]

    title_col = max(
        title_candidates,
        key=lambda i: (col_scores_zone[i], col_scores_nonempty[i], -col_scores_scale[i])
    ) if title_candidates else min(doc_col + 1, max_cols - 1)

    remaining = [i for i in range(max_cols) if i not in (doc_col, title_col)]
    scale_col = max(remaining, key=lambda i: col_scores_scale[i]) if remaining else None

    return {"doc_col": doc_col, "title_col": title_col, "scale_col": scale_col}

def _is_likely_scale_text(value_text):
    text = _to_text(value_text).upper()
    if not text:
        return False
    if "PLAN" in text and ":" in text:
        return True
    if ":" in text and any(ch.isdigit() for ch in text):
        return True
    return False

def _pick_title_and_scale_from_row(row, doc_col, title_col, scale_col):
    values = [_to_text(v) for v in row]

    c_title = values[title_col] if title_col < len(values) else ""
    c_scale = values[scale_col] if (scale_col is not None and scale_col < len(values)) else ""

    # Prefer explicit zone-bearing text for title.
    if not c_title or _is_likely_scale_text(c_title):
        zone_candidates = [v for v in values if "ZONE" in v.upper()]
        if zone_candidates:
            c_title = zone_candidates[0]

    # If still weak, pick a long descriptive text that is not doc number/scale.
    if not c_title or _is_likely_scale_text(c_title):
        for idx, v in enumerate(values):
            if idx == doc_col:
                continue
            if not v:
                continue
            if _looks_like_sheet_number(v):
                continue
            if _is_likely_scale_text(v):
                continue
            if len(v) >= 10:
                c_title = v
                break

    if not c_scale:
        for v in values:
            if _is_likely_scale_text(v):
                c_scale = v
                break

    return c_title, c_scale

def _parse_excel_rows(raw_rows):
    """Parse rows from template-like Excel sheets with section headers.

    Expected visible headers are DOCUMENT NUMBER | TITLE | SCALE, with optional
    section title rows (e.g. AIRSIDE) that have only first column populated.
    """
    if not raw_rows:
        return []

    header_layout = _find_header_layout(raw_rows)
    start_index = header_layout["header_row"] + 1
    inferred_layout = _infer_layout_from_data(raw_rows, start_index)

    doc_col = header_layout["doc_col"]
    title_col = header_layout["title_col"]
    scale_col = header_layout["scale_col"]

    # If header mapping looks weak, use inferred mapping.
    if header_layout.get("matched_header_count", 0) < 2 or title_col is None or doc_col == title_col:
        doc_col = inferred_layout["doc_col"]
        title_col = inferred_layout["title_col"]
        scale_col = inferred_layout["scale_col"]

    parsed = _parse_rows_with_layout(raw_rows, start_index, doc_col, title_col, scale_col)

    # If primary column layout found nothing, retry with the inferred layout.
    if not parsed:
        i_doc  = inferred_layout["doc_col"]
        i_title = inferred_layout["title_col"]
        i_scale = inferred_layout["scale_col"]
        if (i_doc, i_title) != (doc_col, title_col):
            parsed = _parse_rows_with_layout(raw_rows, start_index, i_doc, i_title, i_scale)

    # Last resort: try every column combination from row 0 upwards.
    if not parsed:
        parsed = _parse_rows_brute_force(raw_rows)

    return parsed


def _parse_rows_with_layout(raw_rows, start_index, doc_col, title_col, scale_col):
    """Core row-parsing loop for a given column layout."""
    parsed = []
    current_group = ""

    for row in raw_rows[start_index:]:
        c1 = _to_text(row[doc_col]) if doc_col < len(row) else ""
        c2, c3 = _pick_title_and_scale_from_row(row, doc_col, title_col, scale_col)

        # If doc/title are accidentally swapped by formatting, recover using heuristics.
        if c1 and c2 and not _looks_like_sheet_number(c1) and _looks_like_sheet_number(c2):
            c1, c2 = c2, c1

        # Skip empty rows.
        if not c1 and not c2 and not c3:
            continue

        # Skip known note rows.
        if "provided by" in c1.lower() or "provided by" in c2.lower():
            continue

        # Group header rows: document col has text but title/scale are empty.
        # Only treat as a section header if c1 is NOT a valid sheet number
        # (prevents valid sheet-number rows with an empty title from being silently discarded).
        if c1 and not c2 and not c3 and not _looks_like_sheet_number(c1):
            current_group = c1
            continue

        # Skip repeated column headers if present mid-sheet.
        c1_up = c1.upper()
        c2_up = c2.upper()
        if c1_up in ("DOCUMENT NUMBER", "SHEET NUMBER", "NUMBER") or \
           c2_up in ("TITLE", "SHEET NAME", "VIEW NAME", "SHEET NAME/VIEW NAME"):
            continue

        # Valid data rows need at least sheet number + title.
        if not c1 or not c2:
            continue
        if not _looks_like_sheet_number(c1):
            continue
        if _is_likely_scale_text(c2):
            continue

        parsed.append({
            "group": current_group,
            "sheet_number": c1,
            "sheet_name": c2,
            "scale": c3,
        })

    return parsed


def _parse_rows_brute_force(raw_rows):
    """Last-resort parser: scan every row for any column that looks like a sheet number."""
    parsed = []
    for row in raw_rows:
        values = [_to_text(v) for v in row]
        sheet_num = ""
        sheet_num_col = -1
        for idx, v in enumerate(values):
            if _looks_like_sheet_number(v):
                sheet_num = v
                sheet_num_col = idx
                break
        if not sheet_num:
            continue
        # Pick title: first non-empty non-sheet-number value in any other column.
        title = ""
        for idx, v in enumerate(values):
            if idx == sheet_num_col or not v:
                continue
            if _looks_like_sheet_number(v):
                continue
            if _is_likely_scale_text(v):
                continue
            title = v
            break
        if not title:
            continue
        parsed.append({
            "group": "",
            "sheet_number": sheet_num,
            "sheet_name": title,
            "scale": "",
        })
    return parsed

def _parent_title_from_view_title(view_title):
    title = _normalize_spaces(view_title)
    if not title:
        return ""

    # Robustly remove trailing zone token regardless of spacing around dashes.
    # Examples handled:
    #   "... - ZONE 511"
    #   "...-ZONE511"
    #   "...  -   ZONE   511"
    parent_title = re.sub(r"\s*-\s*ZONE\s*[A-Za-z0-9]+\s*$", "", title, flags=re.IGNORECASE)
    return _normalize_spaces(parent_title or title)

def _extract_level_hint_from_parent_title(parent_title):
    title = _to_text(parent_title).upper()

    if "LOWER GROUND" in title:
        return "LOWER GROUND"
    if "GROUND LEVEL" in title:
        return "GROUND LEVEL"

    level_match = re.search(r"\bLEVEL\s*(\d+)\b", title)
    if level_match:
        return "LEVEL {}".format(level_match.group(1))

    return ""

def _zone_from_title_or_sheet(view_title, sheet_number):
    title = _to_text(view_title)
    match = re.search(r"\bZONE\s*([A-Za-z0-9]+)\b", title, re.IGNORECASE)
    if match:
        return match.group(1)
    return extract_zoning(sheet_number)

def _build_excel_view_records(parsed_rows):
    records = []
    for row in parsed_rows:
        dependent_name = _normalize_spaces(row.get("sheet_name"))
        parent_name = _parent_title_from_view_title(dependent_name)
        level_hint = _extract_level_hint_from_parent_title(parent_name)
        zone_code = _zone_from_title_or_sheet(dependent_name, row.get("sheet_number", ""))
        sheet_number = _short_sheet_number(row.get("sheet_number"))

        records.append({
            "group": row.get("group", ""),
            "sheet_number": sheet_number,
            "sheet_name": dependent_name,
            "scale": _to_text(row.get("scale")),
            "parent_view_name": parent_name,
            "dependent_view_name": dependent_name,
            "zone_code": zone_code,
            "level_hint": level_hint,
            "parent_key": _normalize_spaces(parent_name).upper(),
        })

    return records

def _get_primary_view_id_safe(view_obj):
    try:
        primary_id = view_obj.GetPrimaryViewId()
        if primary_id and primary_id != ElementId.InvalidElementId and primary_id.IntegerValue != -1:
            return primary_id
    except Exception:
        pass
    return None

def _is_independent_view(view_obj):
    return _get_primary_view_id_safe(view_obj) is None

def _is_dependent_of_parent(view_obj, parent_id):
    primary_id = _get_primary_view_id_safe(view_obj)
    if primary_id is None or parent_id is None:
        return False
    return primary_id.IntegerValue == parent_id.IntegerValue

def _find_level_for_parent_name(levels, parent_name):
    if not levels:
        return None

    parent_upper = _to_text(parent_name).upper()

    # Prefer exact phrase hit by level name.
    for level in levels:
        level_name = _to_text(level.Name).upper()
        if level_name and level_name in parent_upper:
            return level

    # Common MEP naming heuristics.
    if "LOWER GROUND" in parent_upper:
        for level in levels:
            name = _to_text(level.Name).upper()
            if "LOWER" in name and "GROUND" in name:
                return level

    if "GROUND" in parent_upper:
        for level in levels:
            name = _to_text(level.Name).upper()
            if "GROUND" in name and "LOWER" not in name:
                return level

    level_match = re.search(r"\bLEVEL\s*(\d+)\b", parent_upper)
    if level_match:
        level_number = level_match.group(1)
        for level in levels:
            name = _to_text(level.Name).upper()
            if "LEVEL" in name and level_number in name:
                return level

    return levels[0]

def _find_level_for_view_record(levels, record):
    if not levels:
        return None

    level_hint = _to_text(record.get("level_hint")).upper()
    if level_hint:
        for level in levels:
            level_name = _to_text(level.Name).upper()
            if level_hint in level_name or level_name in level_hint:
                return level

    return _find_level_for_parent_name(levels, record.get("parent_view_name", ""))

def _set_view_name_safe(view_obj, target_name):
    clean_name = _to_text(target_name)
    forbidden_symbols = "\\:{}[]|;<>?`~"
    clean_name = "".join(ch for ch in clean_name if ch not in forbidden_symbols)
    clean_name = clean_name.replace("  ", " ").strip()
    if not clean_name:
        clean_name = "Generated View"

    candidate = clean_name
    for index in range(1, 60):
        try:
            view_obj.Name = candidate
            return candidate
        except Exception:
            candidate += "*"

    raise RuntimeError("Unable to set a unique view name for '{}'".format(clean_name))


def _get_default_floor_plan_view_family_type():
    """Return the first available FloorPlan view family type."""
    view_family_types = list(FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements())
    for view_family_type in view_family_types:
        if view_family_type.ViewFamily == ViewFamily.FloorPlan:
            return view_family_type
    return None


def _start_write_scope(document, scope_name):
    """Start Transaction or SubTransaction depending on document state."""
    if document is None:
        raise RuntimeError("No active Revit document.")
    if document.IsReadOnly:
        raise RuntimeError("Document is read-only. Cannot modify model.")

    if document.IsModifiable:
        sub_t = SubTransaction(document)
        status = sub_t.Start()
        if status != TransactionStatus.Started:
            raise RuntimeError("Could not start write scope (SubTransaction).")
        return sub_t

    t = Transaction(document, scope_name)
    status = t.Start()
    if status != TransactionStatus.Started:
        raise RuntimeError("Could not start write scope (Transaction).")
    return t


def _commit_write_scope(scope_obj):
    if scope_obj is not None and scope_obj.GetStatus() == TransactionStatus.Started:
        scope_obj.Commit()


def _rollback_write_scope(scope_obj):
    if scope_obj is not None and scope_obj.GetStatus() == TransactionStatus.Started:
        scope_obj.RollBack()


def _lookup_parameter_case_insensitive(element, parameter_name):
    """Lookup parameter by name with case-insensitive fallback."""
    if element is None:
        return None

    param = element.LookupParameter(parameter_name)
    if param:
        return param

    target = _to_text(parameter_name).lower()
    for candidate in element.Parameters:
        try:
            name = _to_text(candidate.Definition.Name).lower()
            if name == target:
                return candidate
        except Exception:
            continue
    return None


def _apply_sheet_parameters(sheet, parameters, level_name=""):
    """Apply dynamic sheet parameters to a ViewSheet."""
    if sheet is None or parameters is None:
        return

    try:
        for param_name, raw_value in parameters.items():
            value_text = _to_text(raw_value)
            if not value_text:
                continue

            param = _lookup_parameter_case_insensitive(sheet, param_name)
            if param is None or param.IsReadOnly:
                continue

            try:
                if param.StorageType == StorageType.String:
                    param.Set(value_text)
                elif param.StorageType == StorageType.Integer:
                    param.Set(int(value_text))
                elif param.StorageType == StorageType.Double:
                    param.Set(float(value_text.replace(",", ".")))
                # ElementId and other unsupported types are skipped.
            except Exception as set_ex:
                output.print_md("**Warning:** Could not set parameter '{}' with value '{}': {}".format(param_name, value_text, set_ex))

        # Keep LEVEL autofill from Excel when available and not manually specified.
        if level_name and not _to_text(parameters.get("LEVEL", "")):
            level_param = _lookup_parameter_case_insensitive(sheet, "LEVEL")
            if level_param and not level_param.IsReadOnly and level_param.StorageType == StorageType.String:
                level_param.Set(_to_text(level_name))
    except Exception as ex:
        output.print_md("**Warning:** Could not apply all sheet parameters: {}".format(ex))

# ╦ ╦╔═╗╦  ╔═╗╔═╗╦═╗  ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ╠═╣║╣ ║  ╠═╝║╣ ╠╦╝  ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╩ ╩╚═╝╩═╝╩  ╚═╝╩╚═  ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝
#==================================================

class ViewItemData:
    """Wrapper for view listbox items with selection state and tag data."""
    def __init__(self, display_text, tag_data=None, is_planned=False):
        self.display_text = display_text
        self.tag_data = tag_data  # Can be a View object or record dict
        self.is_planned = is_planned
        self.is_checked = True
    
    def __str__(self):
        return self.display_text

# ╔╦╗╔═╗╦╔╗╔  ╔═╗╔═╗╦═╗╔╦╗
# ║║║╠═╣║║║║  ╠╣ ║ ║╠╦╝║║║
# ╩ ╩╩ ╩╩╝╚╝  ╚  ╚═╝╩╚═╩ ╩
#==================================================

class MHT_SheetGenerator(Window):
    def __init__(self):
        # 🎨 Load XAML
        path_xaml_file = os.path.join(PATH_SCRIPT, 'SheetGeneratorUI.xaml')
        wpf.LoadComponent(self, path_xaml_file)
        self.load_logo()

        #⬇️ Populate the ListBox with views
        self.populate_title_blocks_combo()
        self.UI_combo_excel_sheet.Items.Clear()
        self.excel_rows = []
        self.generated_views_by_title = {}
        self.generated_views_by_sheet_number = {}
        self.records_for_generation = []
        self.views_generated = False
        self.auto_place_enabled = True
        self.run_requested = False
        self._loading_excel_sheet_list = False
        self._last_checkbox_index = -1
        self._updating_checkbox_state = False
        self._checkbox_click_snapshot = []
        self.Closed += self.UIe_window_closed

        #👀 Show Form
        self.UI_viewsListBox.Items.Clear()
        self.ShowDialog()

    # ╦ ╦╔═╗╦  ╔═╗╔═╗╦═╗  ╔╦╗╔═╗╔╦╗╦ ╦╔═╗╔╦╗╔═╗
    # ╠═╣║╣ ║  ╠═╝║╣ ╠╦╝  ║║║║╣  ║ ╠═╣║ ║ ║║╚═╗
    # ╩ ╩╚═╝╩═╝╩  ╚═╝╩╚═  ╩ ╩╚═╝ ╩ ╩ ╩╚═╝═╩╝╚═╝
    # ==================================================

    def load_logo(self):
        logo_path = os.path.join(PATH_SCRIPT, 'ef_logo.png')
        if os.path.exists(logo_path):
            from System.Windows.Media.Imaging import BitmapImage
            self.UI_ef_logo.Source = BitmapImage(Uri(logo_path))

    def duplicate_view(self, listBoxItem, view, duplicate_option):
        """Duplicate a selected view inside SheetGenerator Form."""

        write_scope = None

        try:
            write_scope = _start_write_scope(doc, "Duplicate View")

            # Duplicate View
            new_view_id = view.Duplicate(duplicate_option)
            new_view    = doc.GetElement(new_view_id)

            # Create a new ListBoxItem for the duplicated view
            view_name = '[{}] {}'.format(new_view.ViewType, new_view.Name)
            new_item = ListBoxItem(Content=view_name, Tag=new_view)

            # Insert Duplicated View after original
            index = self.UI_viewsListBox.Items.IndexOf(listBoxItem)
            self.UI_viewsListBox.Items.Insert(index + 1, new_item)

            # Refresh the search filter
            self.UIe_search_changed(None, None)
            _commit_write_scope(write_scope)

        except Exception as ex:
            print("Error duplicating view: {}".format(ex))
            _rollback_write_scope(write_scope)

    def FindAncestor(self, ancestorType, element):
        """Helper method to find ancestor of a specific type."""
        while element is not None and not isinstance(element, ancestorType):
            element = VisualTreeHelper.GetParent(element)
        return element

    def FindDescendant(self, root_element, target_type):
        """Depth-first search for the first visual descendant of a type."""
        if root_element is None:
            return None
        try:
            child_count = VisualTreeHelper.GetChildrenCount(root_element)
        except Exception:
            return None

        for i in range(child_count):
            child = VisualTreeHelper.GetChild(root_element, i)
            if isinstance(child, target_type):
                return child
            found = self.FindDescendant(child, target_type)
            if found is not None:
                return found
        return None

    def populate_views_listbox(self):
        """Keep the list focused on Excel plan/generated results only."""
        self.UI_viewsListBox.Items.Clear()

    def populate_planned_views_list(self, records):
        """Show planned parent/dependent structure from Excel, not generic project views."""
        self.UI_viewsListBox.Items.Clear()
        grouped = {}
        for record in records:
            grouped.setdefault(record["parent_key"], []).append(record)

        for parent_key in sorted(grouped.keys()):
            parent_records = grouped[parent_key]
            parent_name = parent_records[0]["parent_view_name"]
            parent_item_data = ViewItemData("[PARENT][PLANNED] {}".format(parent_name), tag_data=None, is_planned=True)
            self.UI_viewsListBox.Items.Add(parent_item_data)
            for record in parent_records:
                dep_item_data = ViewItemData("   [DEPENDENT][PLANNED] {}".format(record["dependent_view_name"]), tag_data=record, is_planned=True)
                self.UI_viewsListBox.Items.Add(dep_item_data)

    def populate_generated_views_list(self, hierarchy_rows):
        """Show actual created/reused parent and dependent results."""
        self.UI_viewsListBox.Items.Clear()
        for row in hierarchy_rows:
            item_data = ViewItemData(row.get("text", ""), tag_data=row.get("tag"), is_planned=False)
            self.UI_viewsListBox.Items.Add(item_data)

    def populate_title_blocks_combo(self):
        # Clear ComboBox
        self.UI_combo_title_blocks.Items.Clear()

        tb_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
        for tb_type in tb_types:
            key = '[{}] {}'.format(tb_type.FamilyName, Element.Name.GetValue(tb_type))

            text_block = TextBlock()
            text_block.Text = key

            combo_item = ComboBoxItem()
            combo_item.Content = text_block
            combo_item.Tag = tb_type

            self.UI_combo_title_blocks.Items.Add(combo_item)

        if self.UI_combo_title_blocks.Items.Count > 0:
            self.UI_combo_title_blocks.SelectedIndex = 0

    def populate_view_templates_combo(self):
        return

    def populate_global_view_templates_combo(self):
        return

    def _get_parameter_display_value(self, parameter):
        """Get parameter value as display text for UI defaults."""
        if parameter is None:
            return ""

        try:
            if parameter.StorageType == StorageType.String:
                return _to_text(parameter.AsString())
            value_string = parameter.AsValueString()
            if value_string:
                return _to_text(value_string)
            if parameter.StorageType == StorageType.Integer:
                return str(parameter.AsInteger())
            if parameter.StorageType == StorageType.Double:
                return str(parameter.AsDouble())
        except Exception:
            pass
        return ""

    def _collect_shared_sheet_parameters(self):
        """Collect writable shared parameters from sheet instances in the project."""
        parameter_map = {}
        sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())

        for sheet in sheets:
            for param in sheet.Parameters:
                try:
                    if param.IsReadOnly or not param.IsShared:
                        continue
                    if str(param.StorageType) == "None":
                        continue

                    name = _to_text(param.Definition.Name)
                    if not name:
                        continue

                    if name not in parameter_map:
                        parameter_map[name] = {
                            "default": self._get_parameter_display_value(param),
                            "storage": param.StorageType,
                        }
                except Exception:
                    continue

        # Sort for predictable UI ordering.
        return [(k, parameter_map[k]) for k in sorted(parameter_map.keys())]

    def populate_sheet_parameter_inputs(self):
        return

    def populate_scopebox_combo(self):
        return

    def update_parameters_display(self):
        return

    # ╔═╗╦═╗╔═╗╔═╗╔═╗╦═╗╔╦╗╦╔═╗╔═╗
    # ╠═╝╠╦╝║ ║╠═╝║╣ ╠╦╝ ║ ║║╣ ╚═╗
    # ╩  ╩╚═╚═╝╩  ╚═╝╩╚═ ╩ ╩╚═╝╚═╝
    # ==================================================

    @property
    def selected_title_block_type(self):
        combo_item = self.UI_combo_title_blocks.SelectedItem
        if combo_item:
            return combo_item.Tag
        return None

    @property
    def selected_view_family_type(self):
        return _get_default_floor_plan_view_family_type()

    @property
    def selected_global_view_template_id(self):
        return None

    @property
    def sheet_parameters(self):
        return {}

    # ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝
    # ==================================================

    def UIe_window_closed(self, sender, e):
        """Treat any non-run close path as cancel."""
        if not self.run_requested:
            self.run_requested = False

    def UIe_header_btn_close(self, sender, e):
        """Stop application by clicking on a <Close> button in the top right corner."""
        self.run_requested = False
        self.Close()

    def UIe_header_btn_maximize(self, sender, e):
        """Toggle between normal and maximized window states."""
        from System.Windows import WindowState
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def UIe_header_drag(self, sender, e):
        """Drag window by holding LeftButton on the header."""
        if e.LeftButton == MouseButtonState.Pressed:
            DragMove(self)

    def UIe_RequestNavigate(self, sender, e):
        """Forwarding for a Hyperlinks."""
        Start(e.Uri.AbsoluteUri)

    def UIe_search_changed(self, sender, e):
        """Filter items in the viewsListBox based on search input."""
        search_input = self.UI_search.Text.strip().lower()
        search_words = search_input.split() if search_input else []

        try:
            self.UI_viewsListBox.UpdateLayout()
        except Exception:
            pass

        def _set_row_visibility(list_item, is_visible):
            container = self.UI_viewsListBox.ItemContainerGenerator.ContainerFromItem(list_item)
            if container is not None:
                container.Visibility = Visibility.Visible if is_visible else Visibility.Collapsed

        # If empty search, show all items
        if not search_words:
            for item in self.UI_viewsListBox.Items:
                _set_row_visibility(item, True)
            return

        # Filter items based on search
        used_view_ids = []

        for item in self.UI_viewsListBox.Items:
            display_text = ""
            view = None
            if isinstance(item, ViewItemData):
                display_text = _to_text(item.display_text).lower()
                view = item.tag_data if isinstance(item.tag_data, View) else None
            elif isinstance(item, ListBoxItem):
                display_text = _to_text(item.Content).lower()
                view = item.Tag
            else:
                display_text = _to_text(item).lower()

            _set_row_visibility(item, all(word in display_text for word in search_words))

    def _iter_view_items(self):
        for item in self.UI_viewsListBox.Items:
            if isinstance(item, ViewItemData):
                yield item

    def _sync_checkbox_for_item(self, item_data):
        container = self.UI_viewsListBox.ItemContainerGenerator.ContainerFromItem(item_data)
        if container is None:
            return
        checkbox = self.FindDescendant(container, CheckBox)
        if checkbox is None:
            return
        desired = bool(item_data.is_checked)
        if checkbox.IsChecked != desired:
            checkbox.IsChecked = desired

    def _set_item_checked(self, item_data, is_checked):
        if not isinstance(item_data, ViewItemData):
            return
        item_data.is_checked = bool(is_checked)
        self._sync_checkbox_for_item(item_data)

    def _set_checked_for_items(self, items, is_checked):
        self._updating_checkbox_state = True
        try:
            for item_data in items:
                self._set_item_checked(item_data, is_checked)
        finally:
            self._updating_checkbox_state = False

    def UIe_select_all_views(self, sender, e):
        """Check all view rows currently shown in the list."""
        target_items = []
        for item_data in self._iter_view_items():
            container = self.UI_viewsListBox.ItemContainerGenerator.ContainerFromItem(item_data)
            if container is None or container.Visibility == Visibility.Visible:
                target_items.append(item_data)
        self._set_checked_for_items(target_items, True)

    def UIe_deselect_all_views(self, sender, e):
        """Uncheck all view rows currently shown in the list."""
        target_items = []
        for item_data in self._iter_view_items():
            container = self.UI_viewsListBox.ItemContainerGenerator.ContainerFromItem(item_data)
            if container is None or container.Visibility == Visibility.Visible:
                target_items.append(item_data)
        self._set_checked_for_items(target_items, False)

    def _apply_checkbox_modifier_selection(self, item_data, desired_state):
        index = self.UI_viewsListBox.Items.IndexOf(item_data)
        if index < 0:
            return

        modifiers = Keyboard.Modifiers
        is_shift = (modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
        is_alt = (modifiers & ModifierKeys.Alt) == ModifierKeys.Alt

        snapshot_targets = [x for x in self._checkbox_click_snapshot if isinstance(x, ViewItemData)]
        self._checkbox_click_snapshot = []

        # ALT: apply check/uncheck to currently selected rows, plus the clicked row.
        if is_alt:
            targets = []
            for selected_item in self.UI_viewsListBox.SelectedItems:
                if isinstance(selected_item, ViewItemData):
                    targets.append(selected_item)
            if item_data not in targets:
                targets.append(item_data)
            self._set_checked_for_items(targets, desired_state)
            self._last_checkbox_index = index
            return

        # Default bulk behavior: if multiple rows were selected before the checkbox click,
        # apply the same checked state to all of them.
        if len(snapshot_targets) > 1:
            targets = list(snapshot_targets)
            if item_data not in targets:
                targets.append(item_data)
            self._set_checked_for_items(targets, desired_state)
            self._last_checkbox_index = index
            return

        # SHIFT: apply range from last clicked checkbox index to current.
        if is_shift and self._last_checkbox_index >= 0:
            start_i = min(self._last_checkbox_index, index)
            end_i = max(self._last_checkbox_index, index)
            targets = []
            for i in range(start_i, end_i + 1):
                row_item = self.UI_viewsListBox.Items[i]
                if isinstance(row_item, ViewItemData):
                    targets.append(row_item)
            self._set_checked_for_items(targets, desired_state)
            self._last_checkbox_index = index
            return

        # Default single row behavior.
        self._set_checked_for_items([item_data], desired_state)
        self._last_checkbox_index = index

    def UIe_view_checkbox_preview_mouse_down(self, sender, e):
        """Snapshot selected rows before checkbox click can alter ListBox selection."""
        snapshot = []
        try:
            for selected_item in self.UI_viewsListBox.SelectedItems:
                if isinstance(selected_item, ViewItemData):
                    snapshot.append(selected_item)
        except Exception:
            snapshot = []

        listbox_item = self.FindAncestor(ListBoxItem, sender)
        if listbox_item and isinstance(listbox_item.DataContext, ViewItemData):
            clicked_item = listbox_item.DataContext
            if clicked_item not in snapshot:
                snapshot.append(clicked_item)

        self._checkbox_click_snapshot = snapshot

    def UIe_DuplicateView(self, sender, e):
        # Get the MenuItem
        menuItem = sender
        header = menuItem.Header

        # Get the ContextMenu
        contextMenu = menuItem.Parent
        # Get the PlacementTarget (ListBoxItem)
        listBoxItem = contextMenu.PlacementTarget
        # Get the view from the Tag
        view = listBoxItem.Tag
        if view is None:
            return

        # Determine the duplication option based on the MenuItem's Header
        if header == "Duplicate":                duplicate_option = ViewDuplicateOption.Duplicate
        elif header == "Duplicate As Detailed":  duplicate_option = ViewDuplicateOption.WithDetailing
        elif header == "Duplicate As Dependent": duplicate_option = ViewDuplicateOption.AsDependent
        else: return

        self.duplicate_view(listBoxItem, view, duplicate_option)

    def UIe_remove_item_on_right_click(self, sender, e):
        """Handle right-click to remove an item from the StackPanel and restore it in the ListBox."""
        # Get the clicked TextBlock
        text_block = sender

        # Find the parent StackPanel
        stack_panel = self.FindAncestor(StackPanel, text_block)
        if stack_panel:
            # Remove the TextBlock from the StackPanel
            stack_panel.Children.Remove(text_block)

        # Restore the corresponding ListBoxItem in the ListBox
        for item in self.UI_viewsListBox.Items:
            if item.Tag == text_block.Tag:  # Match by Tag
                item.Visibility = Visibility.Visible
                break

    def UIe_param_changed(self, sender, e):
        return

    def UIe_apply_scopebox_to_selected(self, sender, e):
        return

    def populate_excel_sheets_combo(self):
        """Populate worksheet names from selected Excel file."""
        self._loading_excel_sheet_list = True
        try:
            self.UI_combo_excel_sheet.Items.Clear()
            if not hasattr(self, 'excel_path') or not self.excel_path:
                return

            sheet_names = _get_excel_sheet_names(self.excel_path)
            for sheet_name in sheet_names:
                combo_item = ComboBoxItem()
                combo_item.Content = sheet_name
                combo_item.Tag = sheet_name
                self.UI_combo_excel_sheet.Items.Add(combo_item)

            if self.UI_combo_excel_sheet.Items.Count > 0:
                self.UI_combo_excel_sheet.SelectedIndex = 0
        finally:
            self._loading_excel_sheet_list = False

        # Load data for the initially selected worksheet once population is complete.
        if self.UI_combo_excel_sheet.Items.Count > 0:
            self.load_excel_data(silent_on_empty=True)

    def _show_preview_text(self, title, lines):
        """Render preview content inside the form (non-modal)."""
        if not isinstance(lines, list):
            lines = [_to_text(lines)]
        self.UI_preview_title.Text = _to_text(title)
        self.UI_preview_text.Text = "\n".join([_to_text(x) for x in lines])

    def UIe_clear_preview(self, sender, e):
        self.UI_preview_title.Text = "Preview"
        self.UI_preview_text.Text = "Preview output will appear here."

    def _get_checked_planned_records(self):
        """Return checked planned dependent records from the views list."""
        selected_records = []
        seen_keys = set()
        for item in self.UI_viewsListBox.Items:
            if not isinstance(item, ViewItemData):
                continue
            if not item.is_planned or not item.is_checked:
                continue
            record = item.tag_data
            if not isinstance(record, dict):
                continue
            rec_key = (_to_text(record.get("sheet_number")), _to_text(record.get("dependent_view_name")))
            if rec_key in seen_keys:
                continue
            seen_keys.add(rec_key)
            selected_records.append(record)
        return selected_records

    def UIe_excel_sheet_changed(self, sender, e):
        """Reload cards when worksheet selection changes."""
        if getattr(self, '_loading_excel_sheet_list', False):
            return
        if hasattr(self, 'excel_path') and self.excel_path:
            self.load_excel_data(silent_on_empty=True)

    def UIe_select_excel(self, sender, e):
        """Open file dialog to select an Excel file."""
        from System.Windows.Forms import OpenFileDialog, DialogResult
        dialog = OpenFileDialog()
        dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
        dialog.Title = "Select Excel File"
        if dialog.ShowDialog() == DialogResult.OK:
            self.excel_path = dialog.FileName
            self.UI_excel_path.Text = self.excel_path
            self.populate_excel_sheets_combo()

    def UIe_test_excel(self, sender, e):
        """Validate Excel read and print comprehensive diagnostics to pyRevit output."""
        if not hasattr(self, 'excel_path') or not self.excel_path:
            forms.alert("Select an Excel file first.", title="MHT Sheet Generator")
            return

        selected_sheet = None
        combo_item = self.UI_combo_excel_sheet.SelectedItem
        if combo_item:
            selected_sheet = combo_item.Tag

        try:
            from pyrevit import script as _script
            _out = _script.get_output()

            raw_data = _read_excel_data(self.excel_path, selected_sheet)
            row_count = len(raw_data)
            col_count = max([len(r) for r in raw_data]) if raw_data else 0

            _out.print_md("## Test Read: {}".format(selected_sheet or "(first sheet)"))
            _out.print_md("**Raw rows read:** {}  |  **Max cols per row:** {}".format(row_count, col_count))

            # Show first 5 raw rows
            _out.print_md("### First 5 raw rows:")
            for i, row in enumerate(raw_data[:5]):
                _out.print_md("`Row {}:` {}".format(i, " | ".join(str(v) for v in row)))

            # Show last 3 raw rows
            if row_count > 5:
                _out.print_md("### Last 3 raw rows:")
                for i, row in enumerate(raw_data[-3:], start=row_count - 3):
                    _out.print_md("`Row {}:` {}".format(i, " | ".join(str(v) for v in row)))

            # Header layout
            header_layout = _find_header_layout(raw_data)
            _out.print_md("### Header layout detected:")
            _out.print_md("`header_row={header_row}  doc_col={doc_col}  title_col={title_col}  scale_col={scale_col}  matched={matched_header_count}`".format(**header_layout))

            # Parse and show results
            parsed_data = _parse_excel_rows(raw_data)
            _out.print_md("### Parse result: **{} valid rows**".format(len(parsed_data)))
            for r in parsed_data[:8]:
                _out.print_md("- `{}` | `{}` | `{}`".format(r["sheet_number"], r["sheet_name"], r["scale"]))
            if len(parsed_data) > 8:
                _out.print_md("- *(+{} more)*".format(len(parsed_data) - 8))

            if not parsed_data:
                forms.alert(
                    "Excel read succeeded but no valid rows were parsed.\n"
                    "Raw rows: {}  Cols: {}\nHeader detected at row {}\n\n"
                    "Check pyRevit output for the raw row dump.".format(row_count, col_count, header_layout["header_row"]),
                    title="MHT Sheet Generator"
                )
            else:
                forms.alert(
                    "Excel read OK.  Valid rows: {}\n\nFirst row:\n{}  |  {}".format(
                        len(parsed_data),
                        parsed_data[0]["sheet_number"],
                        parsed_data[0]["sheet_name"]
                    ),
                    title="MHT Sheet Generator"
                )
        except Exception as ex:
            forms.alert("Excel test read failed:\n{}".format(ex), title="MHT Sheet Generator")

    def UIe_generate_views(self, sender, e):
        """Prepare and validate parent/dependent generation preview from checked rows."""
        if not self.excel_rows:
            forms.alert("Load Excel data first using Browse Excel.", title="MHT Sheet Generator")
            return

        selected_records = self._get_checked_planned_records()
        if not selected_records:
            self.views_generated = False
            self.records_for_generation = []
            self._show_preview_text(
                "View Preview",
                [
                    "No checked dependent views found.",
                    "Use the checkboxes in the Views list, then click Generate Views again."
                ]
            )
            return

        view_family_type = _get_default_floor_plan_view_family_type()
        if not view_family_type:
            forms.alert("No Floor Plan view family type was found in this project.", title="MHT Sheet Generator")
            return

        level_elements = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
        if not level_elements:
            forms.alert("No levels found in this project.", title="MHT Sheet Generator")
            return

        records_by_parent = {}
        for record in selected_records:
            records_by_parent.setdefault(record["parent_key"], []).append(record)

        # Validate each parent can resolve a level before enabling sheet generation.
        unresolved = []
        hierarchy_rows = []
        for parent_key in sorted(records_by_parent.keys()):
            parent_records = records_by_parent[parent_key]
            first_record = parent_records[0]
            parent_name = first_record["parent_view_name"]
            level_obj = _find_level_for_view_record(level_elements, first_record)
            if level_obj is None:
                unresolved.append(parent_name)

            hierarchy_rows.append({
                "text": "[PARENT][READY] {}".format(parent_name),
                "tag": None,
            })

            for record in parent_records:
                hierarchy_rows.append({
                    "text": "   [DEPENDENT][READY] {}".format(record["dependent_view_name"]),
                    "tag": record,
                })

        if unresolved:
            self.views_generated = False
            self.records_for_generation = []
            self._show_preview_text(
                "View Preview",
                ["Could not resolve levels for selected items:"] + ["- {}".format(x) for x in unresolved[:12]]
            )
            return

        self.records_for_generation = list(selected_records)
        self.views_generated = True
        self.populate_generated_views_list(hierarchy_rows)
        preview_lines = [
            "Selected dependent views: {}".format(len(selected_records)),
            "Parent views required: {}".format(len(records_by_parent)),
            "",
            "This is a preview only. No model views were created yet.",
            "Click Generate Sheets to create selected views and sheets."
        ]
        for row in hierarchy_rows:
            preview_lines.append(row.get("text", ""))
        self._show_preview_text("View Preview", preview_lines)

    def load_excel_data(self, silent_on_empty=False):
        """Load data from Excel and populate the planned views list."""
        try:
            selected_sheet = None
            combo_item = self.UI_combo_excel_sheet.SelectedItem
            if combo_item:
                selected_sheet = combo_item.Tag

            raw_data = _read_excel_data(self.excel_path, selected_sheet)
            parsed_data = _parse_excel_rows(raw_data)
            if not parsed_data:
                self.excel_rows = []
                self.records_for_generation = []
                self.views_generated = False
                self.generated_views_by_title = {}
                self.generated_views_by_sheet_number = {}
                self.UI_viewsListBox.Items.Clear()
                self.UIe_clear_preview(None, None)
                if not silent_on_empty:
                    forms.alert("No data found in the selected worksheet or the format is invalid.", title="MHT Sheet Generator")
                return

            self.excel_rows = _build_excel_view_records(parsed_data)
            self.records_for_generation = []
            self.views_generated = False
            self.generated_views_by_title = {}
            self.generated_views_by_sheet_number = {}
            self.UIe_clear_preview(None, None)

            self.populate_planned_views_list(self.excel_rows)
        except Exception as ex:
            forms.alert("Error loading Excel: {}".format(ex))

    def UIe_view_checkbox_checked(self, sender, e):
        """Handle checkbox checked event."""
        if self._updating_checkbox_state:
            return
        # Find the parent ListBoxItem and mark it as checked
        listbox_item = self.FindAncestor(ListBoxItem, sender)
        if listbox_item and isinstance(listbox_item.DataContext, ViewItemData):
            self._apply_checkbox_modifier_selection(listbox_item.DataContext, True)
            # Selection changed after preview: require user to refresh Generate Views preview.
            self.views_generated = False

    def UIe_view_checkbox_unchecked(self, sender, e):
        """Handle checkbox unchecked event."""
        if self._updating_checkbox_state:
            return
        # Find the parent ListBoxItem and mark it as unchecked
        listbox_item = self.FindAncestor(ListBoxItem, sender)
        if listbox_item and isinstance(listbox_item.DataContext, ViewItemData):
            self._apply_checkbox_modifier_selection(listbox_item.DataContext, False)
            # Selection changed after preview: require user to refresh Generate Views preview.
            self.views_generated = False

    def _get_sheet_rows_for_active_scope(self):
        """Return unique sheet rows for the current previewed selection or checked scope."""
        source_records = list(self.records_for_generation) if self.records_for_generation else self._get_checked_planned_records()
        unique_rows = []
        seen_sheet_numbers = set()
        for record in source_records:
            sheet_number = _to_text(record.get("sheet_number"))
            if not sheet_number or sheet_number in seen_sheet_numbers:
                continue
            seen_sheet_numbers.add(sheet_number)
            unique_rows.append({
                "sheet_number": sheet_number,
                "sheet_name": _to_text(record.get("sheet_name")),
            })
        return unique_rows

    def _build_sheet_generation_preview(self):
        """Build the exact sheet number/name preview based on current selected scope and rename rules."""
        forbidden_symbols = "\\:{}[]|;<>?`~"

        def _sanitize(value_text):
            value_text = _to_text(value_text)
            return ''.join(c for c in value_text if c not in forbidden_symbols)

        existing_numbers = set()
        existing_sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())
        for sheet in existing_sheets:
            existing_numbers.add(_to_text(sheet.SheetNumber))

        planned_rows = []
        for idx, row in enumerate(self._get_sheet_rows_for_active_scope(), 1):
            original_number = _to_text(row.get("sheet_number"))
            original_name = _to_text(row.get("sheet_name"))

            final_name = _sanitize(original_name)
            candidate_number = _sanitize(original_number)
            if not candidate_number:
                candidate_number = "UNSET-{}".format(idx)

            final_number = candidate_number
            duplicate_adjusted = False
            for _ in range(1, 50):
                if final_number not in existing_numbers:
                    break
                final_number += "*"
                duplicate_adjusted = True

            existing_numbers.add(final_number)

            notes = []
            if final_name != original_name:
                notes.append("name cleaned")
            if candidate_number != original_number:
                notes.append("number cleaned")
            if duplicate_adjusted:
                notes.append("duplicate -> '*' appended")

            planned_rows.append({
                "index": idx,
                "original_number": original_number,
                "original_name": original_name,
                "final_number": final_number,
                "final_name": final_name,
                "notes": ", ".join(notes),
            })

        return planned_rows

    def UIe_preview_sheets(self, sender, e):
        """Preview final sheet numbers and names before generating sheets."""
        rows = self._build_sheet_generation_preview()
        if not rows:
            self._show_preview_text("Sheet Preview", ["No selected sheets found to preview."])
            return

        lines = []
        adjusted_count = 0
        for row in rows:
            line = "{0:02d}. {1} | {2}".format(row["index"], row["final_number"], row["final_name"])
            if row["notes"]:
                line += "   [{}]".format(row["notes"])
                adjusted_count += 1
            lines.append(line)

        preview_text = "\n".join(lines[:40])
        if len(lines) > 40:
            preview_text += "\n... and {} more".format(len(lines) - 40)

        self._show_preview_text(
            "Sheet Preview",
            [
                "Planned sheets: {}".format(len(rows)),
                "Adjusted by rules: {}".format(adjusted_count),
                "",
                preview_text,
            ]
        )

    def UIe_btn_run(self, sender, e):
        if not self.views_generated:
            forms.alert("Generate views first, then generate sheets.", title="MHT Sheet Generator")
            return
        if not self.records_for_generation:
            forms.alert("No previewed selection found. Click Generate Views again.", title="MHT Sheet Generator")
            return
        self.auto_place_enabled = True
        self.run_requested = True
        self.Close()

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

from pyrevit import script
output = script.get_output()

#👀 Show form to the user
try:
    UI           = MHT_SheetGenerator()
    if not getattr(UI, "run_requested", False):
        exit = True
    title_block = UI.selected_title_block_type
    # Strictly use the previewed/selected generation scope from "Generate Views".
    excel_rows = list(UI.records_for_generation)
    auto_place_enabled = UI.auto_place_enabled
    generated_views_by_sheet_number = UI.generated_views_by_sheet_number
except SystemExit:
    exit = True
except:
    exitscript()

if exit:
    import sys
    sys.exit()


def _create_or_reuse_views_from_records(records):
    """Create/reuse parent and dependent views and return sheet-number mapping."""
    if not records:
        return {}, 0, 0
    view_family_type = _get_default_floor_plan_view_family_type()
    if view_family_type is None:
        raise RuntimeError("No Floor Plan view family type was found in this project.")

    levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
    if not levels:
        raise RuntimeError("No levels found in this project.")

    existing_views = [
        v for v in FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Views)
        .WhereElementIsNotElementType()
        .ToElements()
        if isinstance(v, View) and not v.IsTemplate
    ]
    existing_by_name = {_to_text(v.Name): v for v in existing_views}

    records_by_parent = {}
    for record in records:
        records_by_parent.setdefault(record.get("parent_key", ""), []).append(record)

    parent_views = {}
    views_by_sheet = {}
    generated_count = 0
    reused_count = 0

    for parent_key in sorted(records_by_parent.keys()):
        parent_records = records_by_parent[parent_key]
        first_record = parent_records[0]
        parent_name = _to_text(first_record.get("parent_view_name"))

        parent_view = parent_views.get(parent_key)
        if parent_view is None:
            reusable_parent = existing_by_name.get(parent_name)
            if reusable_parent is not None and not _is_independent_view(reusable_parent):
                reusable_parent = None

            if reusable_parent is None:
                target_level = _find_level_for_view_record(levels, first_record)
                if target_level is None:
                    raise RuntimeError("Could not resolve a level for parent view '{}'".format(parent_name))
                parent_view = ViewPlan.Create(doc, view_family_type.Id, target_level.Id)
                # Regenerate so Revit indexes the new element before we rename/use it.
                doc.Regenerate()
                final_parent_name = _set_view_name_safe(parent_view, parent_name)
                existing_by_name[final_parent_name] = parent_view
                generated_count += 1
            else:
                parent_view = reusable_parent
                reused_count += 1

            parent_views[parent_key] = parent_view

        # Regenerate once more before the first Duplicate call on this parent.
        doc.Regenerate()

        for record in parent_records:
            dependent_name = _to_text(record.get("dependent_view_name"))
            sheet_number = _to_text(record.get("sheet_number"))

            reusable_dependent = existing_by_name.get(dependent_name)
            if reusable_dependent is not None and not _is_dependent_of_parent(reusable_dependent, parent_view.Id):
                reusable_dependent = None

            dependent_view = None
            if reusable_dependent is None:
                try:
                    dependent_view_id = parent_view.Duplicate(ViewDuplicateOption.AsDependent)
                    if dependent_view_id is None or dependent_view_id.IntegerValue == -1:
                        raise RuntimeError("Duplicate returned an invalid ElementId.")
                    dependent_view = doc.GetElement(dependent_view_id)
                    if dependent_view is None:
                        raise RuntimeError("doc.GetElement returned None for the duplicated view.")
                    # Regenerate between duplicates so Revit registers each new element.
                    doc.Regenerate()
                    final_dependent_name = _set_view_name_safe(dependent_view, dependent_name)
                    existing_by_name[final_dependent_name] = dependent_view
                    generated_count += 1
                except Exception as dep_ex:
                    output.print_md("**Warning:** Could not create dependent '{}': {}".format(dependent_name, dep_ex))
                    continue
            else:
                dependent_view = reusable_dependent
                reused_count += 1

            if dependent_view is None:
                continue

            if sheet_number:
                if sheet_number not in views_by_sheet:
                    views_by_sheet[sheet_number] = []
                views_by_sheet[sheet_number].append(dependent_view.Id)

    return views_by_sheet, generated_count, reused_count

# Iterate through and create new sheets (if views were added)
report_data = []
write_scope = None

try:
    write_scope = _start_write_scope(doc, "MHT_Sheet Generator")

    # Build views from the validated Excel plan inside one controlled write scope.
    views_from_excel = {}
    generated_views = 0
    reused_views = 0
    if excel_rows:
        views_from_excel, generated_views, reused_views = _create_or_reuse_views_from_records(excel_rows)

    # Preserve any pre-existing mapping but prioritize freshly generated IDs.
    for key, ids in views_from_excel.items():
        generated_views_by_sheet_number[key] = list(ids)

    unique_sheet_rows = []
    seen_sheet_numbers = set()
    for record in excel_rows:
        sheet_number = _to_text(record.get("sheet_number"))
        if not sheet_number or sheet_number in seen_sheet_numbers:
            continue
        seen_sheet_numbers.add(sheet_number)
        unique_sheet_rows.append({
            "sheet_number": sheet_number,
            "sheet_name": _to_text(record.get("sheet_name")),
        })

    with forms.ProgressBar() as pb:
        for n, sheet_row in enumerate(unique_sheet_rows):
            pb.update_progress(n, len(unique_sheet_rows)) #Update Progress Bar

            views_to_place = []

            if auto_place_enabled:
                view_ids = generated_views_by_sheet_number.get(sheet_row["sheet_number"], [])
                for view_id in view_ids:
                    mapped_view = doc.GetElement(view_id)
                    if mapped_view is not None:
                        views_to_place.append(mapped_view)

            if not views_to_place:
                continue

            if title_block is None:
                raise RuntimeError("No title block type is available or selected.")

            # Create Sheet
            new_sheet = ViewSheet.Create(doc, title_block.Id)
            rename_sheet(new_sheet, sheet_row["sheet_name"], sheet_row["sheet_number"])

            # Place views on sheet
            x, y = 0.0, 0.0
            placed_views = []
            for index, view in enumerate(views_to_place):
                if Viewport.CanAddViewToSheet(doc, new_sheet.Id, view.Id):
                    # Simple vertical stack placement, with 2-column wrap.
                    if index > 0 and index % 10 == 0:
                        x += 1.0
                        y = 0.0
                    y += 0.8
                    Viewport.Create(doc, new_sheet.Id, view.Id, XYZ(x, y, 0))
                    placed_views.append(view)

            if placed_views:
                report_data.append((new_sheet, placed_views))

    _commit_write_scope(write_scope)

    if excel_rows:
        forms.alert(
            "Views processed successfully. Generated: {} | Reused: {}".format(generated_views, reused_views),
            title="MHT Sheet Generator"
        )
except Exception as ex:
    _rollback_write_scope(write_scope)
    forms.alert("Sheet generation failed:\n{}".format(ex), title="MHT Sheet Generator")
    exitscript()

# Show summary report without opening a blocking dialog.
if report_data:
    total_sheets = len(report_data)
    total_views = sum(len(vs) for _, vs in report_data)
    forms.alert(
        "Sheet generation completed.\nSheets created: {}\nViews placed: {}".format(total_sheets, total_views),
        title="MHT Sheet Generator"
    )
else:
    forms.alert(
        "No sheets were created. Check that cards have valid sheet data and views to place.",
        title="MHT Sheet Generator"
    )
