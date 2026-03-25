import os
import re
import sys
import clr
import traceback
from datetime import datetime

from System import Type, Activator
from System.Collections.Generic import List
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, DialogResult

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Color,
    CurveElement,
    CurveLoop,
    ElementId,
    ElementTypeGroup,
    ExternalDefinitionCreationOptions,
    FillPatternElement,
    FillPatternTarget,
    FilledRegion,
    FilledRegionType,
    FilteredElementCollector,
    GraphicsStyle,
    HorizontalTextAlignment,
    InstanceBinding,
    Line,
    OverrideGraphicSettings,
    SectionType,
    StorageType,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    ViewDrafting,
    ViewFamily,
    ViewFamilyType,
    ViewSchedule,
    XYZ,
)

try:
    from Autodesk.Revit.DB import GroupTypeId
except Exception:
    GroupTypeId = None

try:
    from Autodesk.Revit.DB import SpecTypeId
except Exception:
    SpecTypeId = None

try:
    from Autodesk.Revit.DB import TextElementBackground
except Exception:
    TextElementBackground = None

from rpw import revit
from pyrevit import forms
from pyrevit.forms import WPFWindow
from pyrevit.script import get_logger

if sys.version_info[0] >= 3:
    basestring = str
    unicode = str

COMMAND_PATH = globals().get("__commandpath__")
if not COMMAND_PATH:
    try:
        COMMAND_PATH = os.path.dirname(__file__)
    except Exception:
        COMMAND_PATH = os.getcwd()

logger = get_logger()
doc = revit.doc

LINK_PARAM_PATH = "MHT_Excel_Link_Path"
LINK_PARAM_SHEET = "MHT_Excel_Link_Sheet"
LINK_PARAM_TARGET_TYPE = "MHT_Excel_Link_TargetType"
LINK_PARAM_TARGET_NAME = "MHT_Excel_Link_TargetName"
LINK_PARAM_CATEGORY = "MHT_Excel_Link_Category"
LINK_PARAM_LAST_UPDATED = "MHT_Excel_Link_LastUpdated"

MAX_TABLE_ROWS = 400
MAX_TABLE_COLS = 120
MAX_TABLE_CELLS = 20000
MAX_CELL_TEXT_LENGTH = 500

EXCEL_WIDTH_SCALE = 0.12
EXCEL_HEIGHT_SCALE = 0.01
CELL_PADDING_X = 0.05
CELL_PADDING_Y = 0.04
MIN_TEXT_WIDTH = 0.10

XL_COLOR_INDEX_NONE = -4142
XL_LINESTYLE_NONE = -4142
XL_EDGE_LEFT = 7
XL_EDGE_TOP = 8
XL_EDGE_BOTTOM = 9
XL_EDGE_RIGHT = 10
XL_HALIGN_GENERAL = 1
XL_HALIGN_LEFT = -4131
XL_HALIGN_CENTER = -4108
XL_HALIGN_RIGHT = -4152
XL_HALIGN_CENTER_ACROSS = 7
XL_HALIGN_FILL = 5
XL_HALIGN_JUSTIFY = -4130
XL_HALIGN_DISTRIBUTED = -4117
XL_VALIGN_TOP = -4160
XL_VALIGN_CENTER = -4108
XL_VALIGN_BOTTOM = -4107

SCHEDULE_PARAM_GROUP = "MHT_Excel_Schedule"
SCHEDULE_PARAM_PREFIX = "MHT_Excel_Schedule_"
SCHEDULE_CATEGORY_CANDIDATES = (
    BuiltInCategory.OST_Rooms,
    BuiltInCategory.OST_MEPSpaces,
    BuiltInCategory.OST_Areas,
)


def _create_text_definition_options(name):
    if SpecTypeId is None:
        raise RuntimeError("SpecTypeId is required for Revit 2025 compatibility.")
    return ExternalDefinitionCreationOptions(name, SpecTypeId.String.Text)


def _insert_parameter_binding(definition, binding):
    bindings = doc.ParameterBindings
    if GroupTypeId is not None:
        try:
            if bindings.Insert(definition, binding, GroupTypeId.Data):
                return
            if bindings.ReInsert(definition, binding, GroupTypeId.Data):
                return
        except Exception:
            pass

    raise RuntimeError("Failed to bind shared parameter {} using GroupTypeId.Data".format(definition.Name))


def _get_project_info():
    return doc.ProjectInformation


def _ensure_shared_param_file(path):
    if not os.path.exists(path):
        with open(path, "w"):
            pass


def _ensure_project_info_param(name):
    project_info = _get_project_info()
    if project_info.LookupParameter(name):
        return

    app = doc.Application
    original_spf = app.SharedParametersFilename
    shared_param_path = os.path.join(COMMAND_PATH, "MHT_SharedParams.txt")
    _ensure_shared_param_file(shared_param_path)
    app.SharedParametersFilename = shared_param_path

    try:
        definition_file = app.OpenSharedParameterFile()
        if definition_file is None:
            forms.alert("Failed to open shared parameters file.", title="Excel Import")
            return

        group = definition_file.Groups.get_Item("MHT_Excel_Link")
        if group is None:
            group = definition_file.Groups.Create("MHT_Excel_Link")

        definition = group.Definitions.get_Item(name)
        if definition is None:
            options = _create_text_definition_options(name)
            definition = group.Definitions.Create(options)

        cat_set = app.Create.NewCategorySet()
        cat_set.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_ProjectInformation))
        binding = app.Create.NewInstanceBinding(cat_set)
        _insert_parameter_binding(definition, binding)
    finally:
        app.SharedParametersFilename = original_spf


def _ensure_project_info_params():
    for name in (
        LINK_PARAM_PATH,
        LINK_PARAM_SHEET,
        LINK_PARAM_TARGET_TYPE,
        LINK_PARAM_TARGET_NAME,
        LINK_PARAM_CATEGORY,
        LINK_PARAM_LAST_UPDATED,
    ):
        _ensure_project_info_param(name)


def _get_or_create_definition(group_name, definition_name):
    app = doc.Application
    original_spf = app.SharedParametersFilename
    shared_param_path = os.path.join(COMMAND_PATH, "MHT_SharedParams.txt")
    _ensure_shared_param_file(shared_param_path)
    app.SharedParametersFilename = shared_param_path

    try:
        definition_file = app.OpenSharedParameterFile()
        if definition_file is None:
            raise RuntimeError("Failed to open shared parameters file.")

        group = definition_file.Groups.get_Item(group_name)
        if group is None:
            group = definition_file.Groups.Create(group_name)

        definition = group.Definitions.get_Item(definition_name)
        if definition is None:
            definition = group.Definitions.Create(_create_text_definition_options(definition_name))
        return definition
    finally:
        app.SharedParametersFilename = original_spf


def _ensure_category_text_param(definition_name, category):
    definition = _get_or_create_definition(SCHEDULE_PARAM_GROUP, definition_name)
    cat_set = doc.Application.Create.NewCategorySet()
    cat_set.Insert(category)
    binding = doc.Application.Create.NewInstanceBinding(cat_set)
    try:
        _insert_parameter_binding(definition, binding)
    except Exception:
        # Parameter is usually already bound when importing again.
        pass


def _sanitize_parameter_token(text):
    token = re.sub(r"[^0-9A-Za-z]+", "_", str(text or "").strip())
    token = token.strip("_")
    return token or "Column"


def _make_schedule_param_name(header_text, index):
    token = _sanitize_parameter_token(header_text)
    base_name = "{}{:02d}_{}".format(SCHEDULE_PARAM_PREFIX, index, token)
    return base_name[:240]


def _get_category_from_builtin(built_in_category):
    try:
        return doc.Settings.Categories.get_Item(built_in_category)
    except Exception:
        return None


def _get_schedule_target_category(preferred_name=None):
    if preferred_name:
        for candidate in SCHEDULE_CATEGORY_CANDIDATES:
            category = _get_category_from_builtin(candidate)
            if category is not None and category.Name == preferred_name:
                return category

    for candidate in SCHEDULE_CATEGORY_CANDIDATES:
        category = _get_category_from_builtin(candidate)
        if category is not None:
            return category

    raise RuntimeError("No supported key-schedule category was found. Rooms, Spaces, or Areas must be available.")


def _get_schedulable_fields(schedule_view):
    fields = {}
    schedule_definition = schedule_view.Definition
    for schedulable_field in schedule_definition.GetSchedulableFields():
        parameter_id = schedulable_field.ParameterId
        if parameter_id is None or parameter_id == ElementId.InvalidElementId:
            continue
        parameter_element = doc.GetElement(parameter_id)
        if parameter_element is not None:
            fields[parameter_element.Name] = schedulable_field
    return fields


def _get_schedule_field_parameter_name(schedule_field):
    try:
        parameter_id = schedule_field.ParameterId
        if parameter_id and parameter_id != ElementId.InvalidElementId:
            parameter_element = doc.GetElement(parameter_id)
            if parameter_element is not None:
                return parameter_element.Name
    except Exception:
        pass

    try:
        return schedule_field.GetName()
    except Exception:
        return ""


def _get_or_create_key_schedule(view_name, category):
    for schedule_view in FilteredElementCollector(doc).OfClass(ViewSchedule):
        try:
            if (
                schedule_view.Name == view_name
                and schedule_view.Definition.IsKeySchedule
                and schedule_view.Definition.CategoryId == category.Id
            ):
                return schedule_view
        except Exception:
            pass

    schedule_view = ViewSchedule.CreateKeySchedule(doc, category.Id)
    try:
        schedule_view.Name = view_name
    except Exception:
        schedule_view.Name = "{}_{}".format(view_name, datetime.now().strftime("%Y%m%d_%H%M%S"))
    return schedule_view


def _prepare_schedule_data(table_model):
    if not table_model:
        return None

    visible_columns = []
    for col_index in range(table_model["cols"]):
        if table_model["col_widths"][col_index] <= 0:
            continue
        has_visible_cell = False
        for row_index in range(table_model["rows"]):
            if not table_model["cells"][row_index][col_index].get("hidden"):
                has_visible_cell = True
                break
        if has_visible_cell:
            visible_columns.append(col_index)

    visible_rows = []
    for row_index in range(table_model["rows"]):
        if table_model["row_heights"][row_index] <= 0:
            continue
        has_visible_cell = False
        for col_index in visible_columns:
            if not table_model["cells"][row_index][col_index].get("hidden"):
                has_visible_cell = True
                break
        if has_visible_cell:
            visible_rows.append(row_index)

    if not visible_columns or not visible_rows:
        return None

    header_row_index = visible_rows[0]
    headers = []
    column_widths = []
    for column_position, col_index in enumerate(visible_columns):
        header_text = _to_cell_text(table_model["cells"][header_row_index][col_index].get("text"))
        if not header_text:
            header_text = "Column {}".format(column_position + 1)
        headers.append(header_text)
        column_widths.append(max(0.25, table_model["col_widths"][col_index]))

    data_rows = []
    for row_index in visible_rows[1:]:
        row_values = []
        for col_index in visible_columns:
            row_values.append(_to_cell_text(table_model["cells"][row_index][col_index].get("text")))
        data_rows.append(row_values)

    return {
        "headers": headers,
        "rows": data_rows,
        "column_widths": column_widths,
    }


def _configure_key_schedule_fields(schedule_view, category, headers, column_widths):
    if not headers:
        raise RuntimeError("Schedule mode requires at least one header row.")

    schedule_definition = schedule_view.Definition
    try:
        schedule_view.KeyScheduleParameterName = headers[0]
    except Exception:
        pass

    existing_field_ids = list(schedule_definition.GetFieldOrder())
    for field_id in reversed(existing_field_ids):
        schedule_field = schedule_definition.GetField(field_id)
        parameter_name = _get_schedule_field_parameter_name(schedule_field)
        if parameter_name.startswith(SCHEDULE_PARAM_PREFIX):
            schedule_definition.RemoveField(field_id)

    for header_index, header_text in enumerate(headers[1:], 1):
        _ensure_category_text_param(_make_schedule_param_name(header_text, header_index), category)

    doc.Regenerate()
    schedulable_fields = _get_schedulable_fields(schedule_view)
    for header_index, header_text in enumerate(headers[1:], 1):
        parameter_name = _make_schedule_param_name(header_text, header_index)
        schedulable_field = schedulable_fields.get(parameter_name)
        if schedulable_field is None:
            raise RuntimeError("Could not schedule parameter {}.".format(parameter_name))
        schedule_field = schedule_definition.AddField(schedulable_field)
        schedule_field.ColumnHeading = header_text
        if header_index < len(column_widths):
            schedule_field.GridColumnWidth = max(0.25, column_widths[header_index])

    doc.Regenerate()
    try:
        first_field_id = list(schedule_definition.GetFieldOrder())[0]
        first_field = schedule_definition.GetField(first_field_id)
        first_field.ColumnHeading = headers[0]
        first_field.GridColumnWidth = max(0.25, column_widths[0])
    except Exception:
        pass


def _get_key_schedule_row_elements(schedule_view):
    try:
        elements = list(
            FilteredElementCollector(doc, schedule_view.Id)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        elements = []

    elements = [element for element in elements if element is not None and getattr(element, "Category", None) is not None]
    elements.sort(key=lambda element: element.Id.IntegerValue)
    return elements


def _set_parameter_text_value(element, parameter_name, value_text):
    parameter = None
    try:
        parameter = element.LookupParameter(parameter_name)
    except Exception:
        parameter = None

    if parameter is None or parameter.IsReadOnly:
        return False

    value_text = value_text or ""
    try:
        if parameter.StorageType == StorageType.String:
            parameter.Set(value_text)
            return True
        if parameter.StorageType == StorageType.Integer:
            parameter.Set(int(value_text) if value_text else 0)
            return True
        if parameter.StorageType == StorageType.Double:
            parameter.Set(float(value_text) if value_text else 0.0)
            return True
    except Exception:
        return False
    return False


def _get_schedule_field_parameter_names(schedule_view):
    parameter_names = []
    schedule_definition = schedule_view.Definition
    for field_id in schedule_definition.GetFieldOrder():
        schedule_field = schedule_definition.GetField(field_id)
        parameter_names.append(_get_schedule_field_parameter_name(schedule_field))
    return parameter_names


def _clear_schedule_body_rows(body_section):
    for row_number in range(body_section.LastRowNumber, body_section.FirstRowNumber - 1, -1):
        try:
            if body_section.CanRemoveRow(row_number):
                body_section.RemoveRow(row_number)
        except Exception:
            pass


def _build_key_schedule(schedule_view, schedule_data, category):
    headers = schedule_data.get("headers") or []
    rows = schedule_data.get("rows") or []
    column_widths = schedule_data.get("column_widths") or []

    _configure_key_schedule_fields(schedule_view, category, headers, column_widths)
    doc.Regenerate()

    table_data = schedule_view.GetTableData()
    body_section = table_data.GetSectionData(SectionType.Body)
    _clear_schedule_body_rows(body_section)

    for _ in rows:
        insert_row_number = body_section.LastRowNumber + 1 if body_section.NumberOfRows > 0 else body_section.FirstRowNumber
        if not body_section.CanInsertRow(insert_row_number):
            insert_row_number = body_section.FirstRowNumber
        body_section.InsertRow(insert_row_number)

    doc.Regenerate()
    body_section = schedule_view.GetTableData().GetSectionData(SectionType.Body)
    first_row_number = body_section.FirstRowNumber
    first_column_number = body_section.FirstColumnNumber

    parameter_names = _get_schedule_field_parameter_names(schedule_view)
    row_elements = _get_key_schedule_row_elements(schedule_view)
    if len(row_elements) < len(rows):
        raise RuntimeError(
            "Schedule rows were created, but only {} backing elements were found for {} Excel rows.".format(
                len(row_elements),
                len(rows),
            )
        )

    for row_offset, row_values in enumerate(rows):
        row_element = row_elements[row_offset]
        for col_offset, cell_text in enumerate(row_values):
            if col_offset >= len(parameter_names):
                continue
            value_text = cell_text
            if col_offset == 0 and not value_text:
                value_text = "Row {}".format(row_offset + 1)
            _set_parameter_text_value(row_element, parameter_names[col_offset], value_text)

    for col_offset, width_value in enumerate(column_widths):
        try:
            body_section.SetColumnWidth(first_column_number + col_offset, max(0.25, width_value))
        except Exception:
            pass


def _set_project_info_value(name, value):
    param = _get_project_info().LookupParameter(name)
    if param:
        param.Set(str(value) if value is not None else "")


def _get_project_info_value(name):
    param = _get_project_info().LookupParameter(name)
    return param.AsString() if param else None


def _release_com_object(com_object):
    if com_object is None:
        return
    try:
        Marshal.FinalReleaseComObject(com_object)
    except Exception:
        pass


def _com_set(obj, member_name, value):
    try:
        setattr(obj, member_name, value)
        return
    except Exception:
        pass

    try:
        obj.GetType().InvokeMember(
            member_name,
            BindingFlags.SetProperty,
            None,
            obj,
            (value,),
        )
    except Exception:
        pass


def _com_get(obj, member_name):
    try:
        return getattr(obj, member_name)
    except Exception:
        return obj.GetType().InvokeMember(
            member_name,
            BindingFlags.GetProperty,
            None,
            obj,
            None,
        )


def _com_call(obj, member_name, *args):
    try:
        member = getattr(obj, member_name)
        return member(*args)
    except Exception:
        try:
            return obj.GetType().InvokeMember(
                member_name,
                BindingFlags.InvokeMethod,
                None,
                obj,
                args,
            )
        except Exception as err:
            raise RuntimeError("COM call failed: {}({}) - {}".format(member_name, args, err))


def _com_item(collection_obj, index):
    try:
        return collection_obj[index]
    except Exception:
        pass

    try:
        return collection_obj.Item[index]
    except Exception:
        pass

    try:
        return collection_obj.Item(index)
    except Exception:
        pass

    return collection_obj.GetType().InvokeMember(
        "Item",
        BindingFlags.GetProperty,
        None,
        collection_obj,
        (index,),
    )


def _com_item2(collection_obj, index1, index2):
    try:
        return collection_obj[index1, index2]
    except Exception:
        pass

    try:
        return collection_obj.Item[index1, index2]
    except Exception:
        pass

    try:
        return collection_obj.Item(index1, index2)
    except Exception:
        pass

    return collection_obj.GetType().InvokeMember(
        "Item",
        BindingFlags.GetProperty,
        None,
        collection_obj,
        (index1, index2),
    )


def _normalize_excel_values(values):
    if values is None:
        return []

    if isinstance(values, tuple):
        if len(values) > 0 and isinstance(values[0], tuple):
            normalized = []
            for row in values:
                normalized.append(["" if cell is None else cell for cell in row])
            return normalized
        return [["" if cell is None else cell for cell in values]]

    if hasattr(values, "GetLowerBound") and hasattr(values, "GetUpperBound"):
        dimensions = int(values.Rank)
        if dimensions == 2:
            row_start = int(values.GetLowerBound(0))
            row_end = int(values.GetUpperBound(0))
            col_start = int(values.GetLowerBound(1))
            col_end = int(values.GetUpperBound(1))
            normalized = []
            for row_index in range(row_start, row_end + 1):
                row_data = []
                for col_index in range(col_start, col_end + 1):
                    cell_value = values.GetValue(row_index, col_index)
                    row_data.append("" if cell_value is None else cell_value)
                normalized.append(row_data)
            return normalized
        if dimensions == 1:
            row_start = int(values.GetLowerBound(0))
            row_end = int(values.GetUpperBound(0))
            row_data = []
            for row_index in range(row_start, row_end + 1):
                cell_value = values.GetValue(row_index)
                row_data.append("" if cell_value is None else cell_value)
            return [row_data]

    return [["" if values is None else values]]


def _to_cell_text(value):
    if value is None:
        return ""

    if isinstance(value, float):
        try:
            if abs(value - int(value)) < 0.0000001:
                value = int(value)
        except Exception:
            pass

    try:
        text = unicode(value)
    except Exception:
        text = str(value)

    text = text.replace("\r", " ").strip()
    if len(text) > MAX_CELL_TEXT_LENGTH:
        text = text[:MAX_CELL_TEXT_LENGTH - 3] + "..."
    return text


def _safe_float(value, default_value=0.0):
    try:
        return float(value)
    except Exception:
        return default_value


def _safe_int(value, default_value=0):
    try:
        return int(value)
    except Exception:
        return default_value


def _clamp255(value):
    try:
        value = int(round(float(value)))
    except Exception:
        value = 0
    return max(0, min(255, value))


def _rgb_tuple(r, g, b):
    return (_clamp255(r), _clamp255(g), _clamp255(b))


def _rgb_to_revit_color(rgb):
    if rgb is None:
        return None
    return Color(_clamp255(rgb[0]), _clamp255(rgb[1]), _clamp255(rgb[2]))


def _rgb_to_revit_int(rgb):
    if rgb is None:
        return 0
    return _clamp255(rgb[0]) + (_clamp255(rgb[1]) << 8) + (_clamp255(rgb[2]) << 16)


def _excel_color_to_rgb(color_value):
    if color_value is None:
        return None
    try:
        color_value = int(color_value)
    except Exception:
        return None
    if color_value < 0:
        return None
    return _rgb_tuple(color_value & 255, (color_value >> 8) & 255, (color_value >> 16) & 255)


def _is_number_like(value):
    return isinstance(value, (int, float))


def _normalize_horizontal_alignment(value, raw_value):
    try:
        value = int(value)
    except Exception:
        value = XL_HALIGN_GENERAL

    if value in (XL_HALIGN_CENTER, XL_HALIGN_CENTER_ACROSS, XL_HALIGN_DISTRIBUTED):
        return "center"
    if value == XL_HALIGN_RIGHT:
        return "right"
    if value in (XL_HALIGN_LEFT, XL_HALIGN_FILL, XL_HALIGN_JUSTIFY):
        return "left"
    if value == XL_HALIGN_GENERAL:
        return "right" if _is_number_like(raw_value) else "left"
    return "left"


def _normalize_vertical_alignment(value):
    try:
        value = int(value)
    except Exception:
        value = XL_VALIGN_TOP

    if value == XL_VALIGN_CENTER:
        return "middle"
    if value == XL_VALIGN_BOTTOM:
        return "bottom"
    return "top"


def _excel_border_weight_to_revit(weight_value, line_style):
    if line_style in (None, XL_LINESTYLE_NONE):
        return None

    try:
        weight_value = int(weight_value)
    except Exception:
        weight_value = 2

    if weight_value == 1:
        return 1
    if weight_value == 2:
        return 2
    if weight_value == 4:
        return 6
    if weight_value == -4138:
        return 4
    return 2


def _border_key(x1, y1, x2, y2):
    if (x1, y1) <= (x2, y2):
        return (round(x1, 6), round(y1, 6), round(x2, 6), round(y2, 6))
    return (round(x2, 6), round(y2, 6), round(x1, 6), round(y1, 6))


def _has_visible_border(border_data):
    return bool(border_data and border_data.get("visible"))


def _cell_has_visual_content(cell):
    if cell is None:
        return False
    if cell.get("hidden"):
        return False
    if cell.get("text"):
        return True
    if cell.get("has_fill"):
        return True
    borders = cell.get("borders") or {}
    for border_data in borders.values():
        if _has_visible_border(border_data):
            return True
    if cell.get("is_merge_root") and (cell.get("row_span", 1) > 1 or cell.get("col_span", 1) > 1):
        return True
    return False


def _get_value_from_matrix(values, row_index, col_index):
    if row_index < 0 or col_index < 0:
        return ""
    if row_index >= len(values):
        return ""
    row_data = values[row_index]
    if col_index >= len(row_data):
        return ""
    return row_data[col_index]


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
            workbook.Close(False)
    except Exception:
        pass

    try:
        if excel_app is not None:
            excel_app.Quit()
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

        if worksheet is None:
            raise RuntimeError("Worksheet not found: {}".format(sheet_name_text))

        return worksheet, sheets
    except Exception:
        _release_com_object(sheets)
        raise


def _read_excel_sheet_names(file_path):
    excel_app = None
    workbook = None
    worksheets = None
    try:
        excel_app, workbook = _open_excel_workbook(file_path)
        sheet_names = []
        worksheets = _com_get(workbook, "Worksheets")
        sheet_count = int(_com_get(worksheets, "Count"))
        for index in range(1, sheet_count + 1):
            worksheet = _com_item(worksheets, index)
            sheet_names.append(_com_get(worksheet, "Name"))
            _release_com_object(worksheet)
        return sheet_names
    finally:
        _release_com_object(worksheets)
        _close_excel_workbook(excel_app, workbook)


def _read_excel_borders(border_owner):
    result = {
        "left": {"visible": False, "weight": None, "color": None},
        "top": {"visible": False, "weight": None, "color": None},
        "right": {"visible": False, "weight": None, "color": None},
        "bottom": {"visible": False, "weight": None, "color": None},
    }

    borders = None
    try:
        borders = _com_get(border_owner, "Borders")
        for side_name, border_index in (("left", XL_EDGE_LEFT), ("top", XL_EDGE_TOP), ("right", XL_EDGE_RIGHT), ("bottom", XL_EDGE_BOTTOM)):
            border = None
            try:
                border = _com_call(borders, "Item", border_index)
                line_style = _com_get(border, "LineStyle")
                visible = line_style not in (None, XL_LINESTYLE_NONE)
                result[side_name] = {
                    "visible": visible,
                    "weight": _excel_border_weight_to_revit(_com_get(border, "Weight"), line_style),
                    "color": _excel_color_to_rgb(_com_get(border, "Color")) or _rgb_tuple(0, 0, 0),
                }
            except Exception:
                pass
            finally:
                _release_com_object(border)
    finally:
        _release_com_object(borders)

    return result


def _read_excel_table(file_path, sheet_name):
    excel_app = None
    workbook = None
    worksheet = None
    sheets = None
    used_range = None
    worksheet_rows = None
    worksheet_columns = None
    worksheet_cells = None

    try:
        excel_app, workbook = _open_excel_workbook(file_path)
        worksheet, sheets = _get_excel_worksheet(workbook, sheet_name)
        used_range = _com_get(worksheet, "UsedRange")

        start_row = _safe_int(_com_get(used_range, "Row"), 1)
        start_col = _safe_int(_com_get(used_range, "Column"), 1)
        used_rows = _com_get(used_range, "Rows")
        used_cols = _com_get(used_range, "Columns")
        row_count = _safe_int(_com_get(used_rows, "Count"), 0)
        col_count = _safe_int(_com_get(used_cols, "Count"), 0)
        values = _normalize_excel_values(_com_get(used_range, "Value2"))
        _release_com_object(used_rows)
        _release_com_object(used_cols)

        worksheet_rows = _com_get(worksheet, "Rows")
        worksheet_columns = _com_get(worksheet, "Columns")
        worksheet_cells = _com_get(worksheet, "Cells")

        row_heights = []
        row_hidden = []
        for row_index in range(row_count):
            row_obj = None
            try:
                row_obj = _com_item(worksheet_rows, start_row + row_index)
                hidden = bool(_com_get(row_obj, "Hidden"))
                height = 0.0 if hidden else _safe_float(_com_get(row_obj, "RowHeight"), 15.0) * EXCEL_HEIGHT_SCALE
                row_heights.append(height)
                row_hidden.append(hidden)
            finally:
                _release_com_object(row_obj)

        col_widths = []
        col_hidden = []
        for col_index in range(col_count):
            col_obj = None
            try:
                col_obj = _com_item(worksheet_columns, start_col + col_index)
                hidden = bool(_com_get(col_obj, "Hidden"))
                width = 0.0 if hidden else _safe_float(_com_get(col_obj, "ColumnWidth"), 8.43) * EXCEL_WIDTH_SCALE
                col_widths.append(width)
                col_hidden.append(hidden)
            finally:
                _release_com_object(col_obj)

        cells = []
        for row_index in range(row_count):
            row_cells = []
            for col_index in range(col_count):
                cell_obj = None
                font_obj = None
                interior_obj = None
                merge_area = None
                border_owner = None
                try:
                    cell_obj = _com_item2(worksheet_cells, start_row + row_index, start_col + col_index)
                    raw_value = _get_value_from_matrix(values, row_index, col_index)
                    text = _to_cell_text(raw_value)

                    font_obj = _com_get(cell_obj, "Font")
                    interior_obj = _com_get(cell_obj, "Interior")

                    is_merged = bool(_com_get(cell_obj, "MergeCells"))
                    merge_root = (row_index, col_index)
                    row_span = 1
                    col_span = 1
                    if is_merged:
                        merge_area = _com_get(cell_obj, "MergeArea")
                        merge_rows = _com_get(merge_area, "Rows")
                        merge_cols = _com_get(merge_area, "Columns")
                        root_row = _safe_int(_com_get(merge_area, "Row"), start_row) - start_row
                        root_col = _safe_int(_com_get(merge_area, "Column"), start_col) - start_col
                        row_span = _safe_int(_com_get(merge_rows, "Count"), 1)
                        col_span = _safe_int(_com_get(merge_cols, "Count"), 1)
                        merge_root = (root_row, root_col)
                        _release_com_object(merge_rows)
                        _release_com_object(merge_cols)

                    border_owner = merge_area if merge_area is not None else cell_obj
                    fill_color_index = _com_get(interior_obj, "ColorIndex")
                    has_fill = fill_color_index not in (None, XL_COLOR_INDEX_NONE)
                    row_cells.append({
                        "raw_value": raw_value,
                        "text": text,
                        "hidden": row_hidden[row_index] or col_hidden[col_index],
                        "font_name": str(_com_get(font_obj, "Name") or "Arial"),
                        "font_size": _safe_float(_com_get(font_obj, "Size"), 9.0),
                        "bold": bool(_com_get(font_obj, "Bold")),
                        "italic": bool(_com_get(font_obj, "Italic")),
                        "underline": _safe_int(_com_get(font_obj, "Underline"), 0) not in (0, -4142),
                        "text_color": _excel_color_to_rgb(_com_get(font_obj, "Color")) or _rgb_tuple(0, 0, 0),
                        "has_fill": has_fill,
                        "bg_color": _excel_color_to_rgb(_com_get(interior_obj, "Color")) if has_fill else None,
                        "wrap": bool(_com_get(cell_obj, "WrapText")),
                        "h_align": _normalize_horizontal_alignment(_com_get(cell_obj, "HorizontalAlignment"), raw_value),
                        "v_align": _normalize_vertical_alignment(_com_get(cell_obj, "VerticalAlignment")),
                        "merge_root": merge_root,
                        "is_merge_root": merge_root == (row_index, col_index),
                        "row_span": row_span,
                        "col_span": col_span,
                        "borders": _read_excel_borders(border_owner),
                    })
                finally:
                    _release_com_object(merge_area)
                    _release_com_object(interior_obj)
                    _release_com_object(font_obj)
                    _release_com_object(cell_obj)
            cells.append(row_cells)

        return {
            "rows": row_count,
            "cols": col_count,
            "row_heights": row_heights,
            "col_widths": col_widths,
            "cells": cells,
        }
    finally:
        _release_com_object(worksheet_cells)
        _release_com_object(worksheet_columns)
        _release_com_object(worksheet_rows)
        _release_com_object(used_range)
        _release_com_object(worksheet)
        _release_com_object(sheets)
        _close_excel_workbook(excel_app, workbook)


def _prepare_table_model(table_model):
    if not table_model:
        return None

    cells = table_model.get("cells") or []
    row_count = len(cells)
    col_count = len(cells[0]) if row_count else 0
    if row_count == 0 or col_count == 0:
        return None

    last_row = -1
    last_col = -1
    for row_index in range(row_count):
        row_has_content = False
        for col_index in range(col_count):
            cell = cells[row_index][col_index]
            if _cell_has_visual_content(cell):
                row_has_content = True
                candidate_col = col_index + max(0, cell.get("col_span", 1) - 1)
                if candidate_col > last_col:
                    last_col = candidate_col
                candidate_row = row_index + max(0, cell.get("row_span", 1) - 1)
                if candidate_row > last_row:
                    last_row = candidate_row
        if row_has_content and row_index > last_row:
            last_row = row_index

    if last_row < 0 or last_col < 0:
        return None

    trimmed_rows = last_row + 1
    trimmed_cols = last_col + 1

    if trimmed_rows > MAX_TABLE_ROWS or trimmed_cols > MAX_TABLE_COLS:
        raise RuntimeError(
            "Excel table is too large ({} rows x {} cols). Limit is {} rows x {} cols.".format(
                trimmed_rows,
                trimmed_cols,
                MAX_TABLE_ROWS,
                MAX_TABLE_COLS,
            )
        )

    if trimmed_rows * trimmed_cols > MAX_TABLE_CELLS:
        raise RuntimeError(
            "Excel table has too many cells ({}). Limit is {}.".format(
                trimmed_rows * trimmed_cols,
                MAX_TABLE_CELLS,
            )
        )

    prepared_cells = []
    for row_index in range(trimmed_rows):
        prepared_cells.append(table_model["cells"][row_index][:trimmed_cols])

    return {
        "rows": trimmed_rows,
        "cols": trimmed_cols,
        "row_heights": table_model["row_heights"][:trimmed_rows],
        "col_widths": table_model["col_widths"][:trimmed_cols],
        "cells": prepared_cells,
    }


def _get_drafting_view_type_id():
    for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if view_type.ViewFamily == ViewFamily.Drafting:
            return view_type.Id
    return None


def _get_text_note_type():
    default_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
    if default_type_id and default_type_id != ElementId.InvalidElementId:
        default_type = doc.GetElement(default_type_id)
        if default_type is not None:
            return default_type

    for text_type in FilteredElementCollector(doc).OfClass(TextNoteType):
        return text_type
    return None


def _get_filled_region_type_id():
    region_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.FilledRegionType)
    if region_type_id and region_type_id != ElementId.InvalidElementId:
        return region_type_id

    for region_type in FilteredElementCollector(doc).OfClass(FilledRegionType):
        return region_type.Id
    return None


def _get_solid_fill_pattern_id():
    try:
        pattern = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, "<Solid fill>")
        if pattern:
            return pattern.Id
    except Exception:
        pass
    return ElementId.InvalidElementId


def _get_invisible_line_style():
    for graphics_style in FilteredElementCollector(doc).OfClass(GraphicsStyle):
        try:
            if graphics_style.GraphicsStyleCategory.Id.IntegerValue == -2000064:
                return graphics_style
        except Exception:
            pass
    return None


def _get_or_create_drafting_view(view_name):
    for view in FilteredElementCollector(doc).OfClass(ViewDrafting):
        if view.Name == view_name:
            return view

    view_type_id = _get_drafting_view_type_id()
    if view_type_id is None:
        raise RuntimeError("Drafting view type not found.")
    new_view = ViewDrafting.Create(doc, view_type_id)
    new_view.Name = view_name
    return new_view


def _create_new_drafting_view(view_name):
    view_type_id = _get_drafting_view_type_id()
    if view_type_id is None:
        raise RuntimeError("Drafting view type not found.")

    created_view = ViewDrafting.Create(doc, view_type_id)
    try:
        created_view.Name = view_name
    except Exception:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        created_view.Name = "{}_{}".format(view_name, timestamp)
    return created_view


def _clear_view_contents(view):
    elements = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements()
    element_ids = []
    for element in elements:
        if isinstance(element, TextNote) or isinstance(element, CurveElement) or isinstance(element, FilledRegion):
            element_ids.append(element.Id)
    if element_ids:
        doc.Delete(List[ElementId](element_ids))


def _create_rectangle_loop(x1, y1, x2, y2):
    p1 = XYZ(x1, y1, 0)
    p2 = XYZ(x2, y1, 0)
    p3 = XYZ(x2, y2, 0)
    p4 = XYZ(x1, y2, 0)
    return CurveLoop.Create(List[Line]([
        Line.CreateBound(p1, p2),
        Line.CreateBound(p2, p3),
        Line.CreateBound(p3, p4),
        Line.CreateBound(p4, p1),
    ]))


def _apply_region_fill_override(view, region, fill_color, line_color=None, line_weight=None):
    solid_fill_id = _get_solid_fill_pattern_id()
    overrides = OverrideGraphicSettings()
    if solid_fill_id and solid_fill_id != ElementId.InvalidElementId and fill_color is not None:
        overrides.SetSurfaceForegroundPatternId(solid_fill_id)
        overrides.SetSurfaceForegroundPatternColor(fill_color)
        overrides.SetSurfaceBackgroundPatternId(solid_fill_id)
        overrides.SetSurfaceBackgroundPatternColor(fill_color)
    if line_color is not None:
        overrides.SetProjectionLineColor(line_color)
    if line_weight:
        overrides.SetProjectionLineWeight(line_weight)
    view.SetElementOverrides(region.Id, overrides)


def _set_int_parameter(element, bip, value):
    try:
        param = element.get_Parameter(bip)
        if param and not param.IsReadOnly:
            param.Set(int(value))
    except Exception:
        pass


def _set_double_parameter(element, bip, value):
    try:
        param = element.get_Parameter(bip)
        if param and not param.IsReadOnly:
            param.Set(float(value))
    except Exception:
        pass


def _set_string_parameter(element, bip, value):
    try:
        param = element.get_Parameter(bip)
        if param and not param.IsReadOnly:
            param.Set(str(value))
    except Exception:
        pass


def _points_to_feet(points_value):
    return float(points_value) / 864.0


def _estimate_text_top_offset(cell_height, font_size_points, vertical_alignment):
    font_height = max(_points_to_feet(font_size_points), 0.03)
    top_padding = min(max(CELL_PADDING_Y, font_height * 0.3), max(cell_height * 0.35, CELL_PADDING_Y))
    if vertical_alignment == "middle":
        return max(top_padding, (cell_height * 0.5) - (font_height * 0.6))
    if vertical_alignment == "bottom":
        return max(top_padding, cell_height - (font_height * 1.2))
    return top_padding


def _get_text_type_name(base_type_name, style_key, style_index):
    return "{}_MHTExcel_{:02d}_{:d}_{:d}_{:d}".format(
        base_type_name,
        style_index,
        int(style_key[1] * 100),
        1 if style_key[2] else 0,
        1 if style_key[3] else 0,
    )


def _get_or_create_text_type(base_text_type, cell_style, cache):
    style_key = (
        cell_style.get("font_name") or "Arial",
        round(_safe_float(cell_style.get("font_size"), 9.0), 2),
        bool(cell_style.get("bold")),
        bool(cell_style.get("italic")),
        bool(cell_style.get("underline")),
        tuple(cell_style.get("text_color") or _rgb_tuple(0, 0, 0)),
    )
    if style_key in cache:
        return cache[style_key]

    base_name = base_text_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    duplicate_name = _get_text_type_name(base_name, style_key, len(cache) + 1)
    try:
        new_type_id = base_text_type.Duplicate(duplicate_name)
    except Exception:
        duplicate_name = "{}_MHTExcel_{}".format(base_name, datetime.now().strftime("%H%M%S%f"))
        new_type_id = base_text_type.Duplicate(duplicate_name)

    text_type = doc.GetElement(new_type_id)
    _set_string_parameter(text_type, BuiltInParameter.TEXT_FONT, style_key[0])
    _set_double_parameter(text_type, BuiltInParameter.TEXT_SIZE, _points_to_feet(style_key[1]))
    _set_int_parameter(text_type, BuiltInParameter.TEXT_STYLE_BOLD, 1 if style_key[2] else 0)
    _set_int_parameter(text_type, BuiltInParameter.TEXT_STYLE_ITALIC, 1 if style_key[3] else 0)
    _set_int_parameter(text_type, BuiltInParameter.TEXT_STYLE_UNDERLINE, 1 if style_key[4] else 0)
    _set_int_parameter(text_type, BuiltInParameter.TEXT_COLOR, _rgb_to_revit_int(style_key[5]))
    if TextElementBackground is not None:
        try:
            _set_int_parameter(text_type, BuiltInParameter.TEXT_BACKGROUND, int(TextElementBackground.TBGR_TRANSPARENT))
        except Exception:
            _set_int_parameter(text_type, BuiltInParameter.TEXT_BACKGROUND, 1)
    else:
        _set_int_parameter(text_type, BuiltInParameter.TEXT_BACKGROUND, 1)

    cache[style_key] = text_type
    return text_type


def _build_layout_positions(widths, heights):
    x_positions = [0.0]
    for width in widths:
        x_positions.append(x_positions[-1] + max(0.0, width))

    y_positions = [0.0]
    for height in heights:
        y_positions.append(y_positions[-1] - max(0.0, height))

    return x_positions, y_positions


def _collect_border_segments(table_model, x_positions, y_positions):
    segment_map = {}
    cells = table_model["cells"]
    rows = table_model["rows"]
    cols = table_model["cols"]

    for row_index in range(rows):
        for col_index in range(cols):
            cell = cells[row_index][col_index]
            if not cell.get("is_merge_root"):
                continue
            if cell.get("hidden"):
                continue

            row_span = max(1, min(cell.get("row_span", 1), rows - row_index))
            col_span = max(1, min(cell.get("col_span", 1), cols - col_index))
            x1 = x_positions[col_index]
            x2 = x_positions[col_index + col_span]
            y1 = y_positions[row_index]
            y2 = y_positions[row_index + row_span]
            if abs(x2 - x1) < 0.000001 or abs(y2 - y1) < 0.000001:
                continue

            for side_name, endpoints in (
                ("top", (x1, y1, x2, y1)),
                ("bottom", (x1, y2, x2, y2)),
                ("left", (x1, y1, x1, y2)),
                ("right", (x2, y1, x2, y2)),
            ):
                border_data = (cell.get("borders") or {}).get(side_name)
                if not _has_visible_border(border_data):
                    continue
                key = _border_key(*endpoints)
                existing = segment_map.get(key)
                new_weight = border_data.get("weight") or 1
                if existing is None or new_weight >= (existing.get("weight") or 1):
                    segment_map[key] = {
                        "coords": endpoints,
                        "weight": new_weight,
                        "color": border_data.get("color") or _rgb_tuple(0, 0, 0),
                    }

    return segment_map


def _build_drafting_table(view, table_model):
    if not table_model:
        return

    rows = table_model["rows"]
    cols = table_model["cols"]
    if rows == 0 or cols == 0:
        return

    text_type = _get_text_note_type()
    if text_type is None:
        raise RuntimeError("TextNoteType not found.")

    region_type_id = _get_filled_region_type_id()
    invisible_line_style = _get_invisible_line_style()
    x_positions, y_positions = _build_layout_positions(table_model["col_widths"], table_model["row_heights"])
    text_type_cache = {}

    for row_index in range(rows):
        for col_index in range(cols):
            cell = table_model["cells"][row_index][col_index]
            if not cell.get("is_merge_root"):
                continue
            if cell.get("hidden"):
                continue

            row_span = max(1, min(cell.get("row_span", 1), rows - row_index))
            col_span = max(1, min(cell.get("col_span", 1), cols - col_index))
            x1 = x_positions[col_index]
            x2 = x_positions[col_index + col_span]
            y1 = y_positions[row_index]
            y2 = y_positions[row_index + row_span]
            cell_width = x2 - x1
            cell_height = y1 - y2
            if cell_width <= 0.000001 or cell_height <= 0.000001:
                continue

            if cell.get("has_fill") and region_type_id is not None:
                loop = _create_rectangle_loop(x1, y1, x2, y2)
                region = FilledRegion.Create(doc, region_type_id, view.Id, [loop])
                if invisible_line_style is not None:
                    try:
                        region.SetLineStyleId(invisible_line_style.Id)
                    except Exception:
                        pass
                _apply_region_fill_override(view, region, _rgb_to_revit_color(cell.get("bg_color") or _rgb_tuple(255, 255, 255)))

            if cell.get("text"):
                style_type = _get_or_create_text_type(text_type, cell, text_type_cache)
                note_options = TextNoteOptions(style_type.Id)
                if cell.get("h_align") == "center":
                    note_options.HorizontalAlignment = HorizontalTextAlignment.Center
                    anchor_x = x1 + (cell_width * 0.5)
                elif cell.get("h_align") == "right":
                    note_options.HorizontalAlignment = HorizontalTextAlignment.Right
                    anchor_x = x2 - CELL_PADDING_X
                else:
                    note_options.HorizontalAlignment = HorizontalTextAlignment.Left
                    anchor_x = x1 + CELL_PADDING_X

                top_offset = _estimate_text_top_offset(cell_height, cell.get("font_size", 9.0), cell.get("v_align"))
                anchor_y = y1 - top_offset
                text_point = XYZ(anchor_x, anchor_y, 0)
                text_width = max(MIN_TEXT_WIDTH, cell_width - (CELL_PADDING_X * 2.0))

                try:
                    if cell.get("wrap"):
                        TextNote.Create(doc, view.Id, text_point, text_width, cell.get("text"), note_options)
                    else:
                        TextNote.Create(doc, view.Id, text_point, cell.get("text"), note_options)
                except Exception:
                    TextNote.Create(doc, view.Id, text_point, text_width, cell.get("text"), note_options)

    border_segments = _collect_border_segments(table_model, x_positions, y_positions)
    for border_segment in border_segments.values():
        x1, y1, x2, y2 = border_segment["coords"]
        detail_line = doc.Create.NewDetailCurve(view, Line.CreateBound(XYZ(x1, y1, 0), XYZ(x2, y2, 0)))
        overrides = OverrideGraphicSettings()
        overrides.SetProjectionLineColor(_rgb_to_revit_color(border_segment["color"]))
        if border_segment.get("weight"):
            overrides.SetProjectionLineWeight(border_segment["weight"])
        view.SetElementOverrides(detail_line.Id, overrides)


def _activate_view(view):
    try:
        revit.uidoc.ActiveView = view
    except Exception as err:
        logger.info("Could not activate view {}: {}".format(view.Name, err))


class ExcelLinkWindow(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, "WPFWindow.xaml")
        self.rb_drafting.IsChecked = True
        self._update_notes_text()

    def _selected_target_type(self):
        if getattr(self, "rb_schedule", None) is not None and self.rb_schedule.IsChecked:
            return "Schedule"
        return "Drafting"

    def _update_notes_text(self):
        if self._selected_target_type() == "Schedule":
            self.tb_notes.Text = (
                "Schedule View mode creates a Revit key schedule. It imports headers, column order, column widths, and cell text, "
                "but Revit schedules cannot reproduce Excel fills, borders, merged cells, or font styling. "
                "Update rebuilds the stored key schedule from the saved Excel link."
            )
        else:
            self.tb_notes.Text = (
                "Drafting View mode rebuilds the sheet as formatted detail graphics, including merged cells, background fills, "
                "text styling, and border weights/colors where Revit drafting elements allow it. "
                "Update rebuilds from the saved Excel link."
            )

    def _selected_sheet_name(self):
        selected = self.cb_sheet.SelectedItem
        if selected is None:
            return ""
        return str(selected).strip()

    def _set_default_target_name(self):
        sheet_name = self._selected_sheet_name()
        if not sheet_name:
            return
        suffix = "_Schedule" if self._selected_target_type() == "Schedule" else ""
        self.tb_target_name.Text = "{}{}".format(sheet_name, suffix)

    def browse_click(self, sender, e):
        dialog = OpenFileDialog()
        dialog.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
        if dialog.ShowDialog() == DialogResult.OK:
            self.tb_excel_path.Text = dialog.FileName
            self._load_sheet_names(dialog.FileName)

    def _load_sheet_names(self, file_path):
        try:
            sheets = _read_excel_sheet_names(file_path)
            self.cb_sheet.ItemsSource = None
            self.cb_sheet.ItemsSource = sheets
            self.cb_sheet.Items.Refresh()
            if sheets:
                self.cb_sheet.SelectedIndex = 0
                self._set_default_target_name()
            else:
                forms.alert("No sheets were found in the selected workbook.", title="Excel Import")
        except Exception as err:
            forms.alert("Failed to read Excel sheets: {}".format(err), title="Excel Import")

    def sheet_changed(self, sender, e):
        self._set_default_target_name()

    def output_changed(self, sender, e):
        self._set_default_target_name()
        self._update_notes_text()

    def import_click(self, sender, e):
        self._run_import(update_only=False)

    def update_click(self, sender, e):
        self._run_import(update_only=True)

    def _run_import(self, update_only=False):
        try:
            output_view = None
            output_mode = ""
            if update_only:
                file_path = _get_project_info_value(LINK_PARAM_PATH)
                sheet_name = _get_project_info_value(LINK_PARAM_SHEET)
                target_type = _get_project_info_value(LINK_PARAM_TARGET_TYPE)
                target_name = _get_project_info_value(LINK_PARAM_TARGET_NAME)
                category_name = _get_project_info_value(LINK_PARAM_CATEGORY) or ""
            else:
                file_path = str(self.tb_excel_path.Text).strip()
                sheet_name = self._selected_sheet_name()
                target_type = self._selected_target_type()
                target_name = str(self.tb_target_name.Text).strip()
                category_name = ""

            if update_only and not file_path:
                forms.alert("No stored Excel link found. Run Import first.", title="Excel Import")
                return
            if not file_path or not os.path.exists(file_path):
                forms.alert("Select a valid Excel file.", title="Excel Import")
                return
            if not sheet_name:
                forms.alert("Select a sheet.", title="Excel Import")
                return
            if not target_name:
                forms.alert("Enter a target name.", title="Excel Import")
                return

            table_model = _read_excel_table(file_path, sheet_name)
            table_model = _prepare_table_model(table_model)
            if not table_model:
                forms.alert("No data or visible formatting found in the selected sheet.", title="Excel Import")
                return

            transaction_name = "Excel to Schedule View" if target_type == "Schedule" else "Excel to Drafting View"
            transaction = Transaction(doc, transaction_name)
            transaction.Start()
            try:
                if target_type == "Schedule":
                    category = _get_schedule_target_category(category_name)
                    schedule_data = _prepare_schedule_data(table_model)
                    if not schedule_data:
                        raise RuntimeError("Schedule mode requires at least one visible header row.")
                    view = _get_or_create_key_schedule(target_name, category)
                    _build_key_schedule(view, schedule_data, category)
                    category_name = category.Name
                    output_mode = "Schedule View"
                else:
                    view = _get_or_create_drafting_view(target_name)
                    _clear_view_contents(view)
                    _build_drafting_table(view, table_model)
                    output_mode = "Drafting View"
                transaction.Commit()
                output_view = view
            except Exception as first_error:
                try:
                    if target_type == "Schedule":
                        raise first_error
                    view = _create_new_drafting_view(target_name)
                    _build_drafting_table(view, table_model)
                    transaction.Commit()
                    output_view = view
                    output_mode = "Drafting View"
                except Exception:
                    transaction.RollBack()
                    raise first_error

            transaction = Transaction(doc, "Store Excel link")
            transaction.Start()
            _ensure_project_info_params()
            _set_project_info_value(LINK_PARAM_PATH, file_path)
            _set_project_info_value(LINK_PARAM_SHEET, sheet_name)
            _set_project_info_value(LINK_PARAM_TARGET_TYPE, target_type)
            _set_project_info_value(LINK_PARAM_TARGET_NAME, target_name)
            _set_project_info_value(LINK_PARAM_CATEGORY, category_name or "")
            _set_project_info_value(LINK_PARAM_LAST_UPDATED, datetime.now().strftime("%Y-%m-%d %H:%M"))
            transaction.Commit()

            if output_view is not None:
                _activate_view(output_view)
                forms.alert(
                    "Import complete. {} created/updated: {}".format(output_mode, output_view.Name),
                    title="Excel Import",
                )
            else:
                forms.alert("Import complete.", title="Excel Import")
        except Exception as err:
            logger.error(traceback.format_exc())
            forms.alert("Import failed: {}".format(err), title="Excel Import")

    def close_click(self, sender, e):
        self.Close()


ExcelLinkWindow().ShowDialog()
