import os
import clr
import traceback
from datetime import datetime

from System import Guid, Type, Activator
from System.Collections.Generic import List
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, DialogResult

from Autodesk.Revit.DB import (
    BuiltInCategory,
    CurveElement,
    ElementId,
    ExternalDefinitionCreationOptions,
    FilteredElementCollector,
    InstanceBinding,
    Line,
    StorageType,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    ViewDrafting,
    ViewFamily,
    ViewFamilyType,
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

from rpw import revit
from pyrevit import forms
from pyrevit.forms import WPFWindow
from pyrevit.script import get_logger

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
    shared_param_path = os.path.join(__commandpath__, "MHT_SharedParams.txt")
    _ensure_shared_param_file(shared_param_path)
    app.SharedParametersFilename = shared_param_path

    definition_file = app.OpenSharedParameterFile()
    if definition_file is None:
        app.SharedParametersFilename = original_spf
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


def _is_blank_cell(value):
    if value is None:
        return True
    if isinstance(value, basestring):
        return len(value.strip()) == 0
    return False


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

    text = text.replace("\r", " ").replace("\n", " ").strip()
    if len(text) > MAX_CELL_TEXT_LENGTH:
        text = text[:MAX_CELL_TEXT_LENGTH - 3] + "..."
    return text


def _prepare_table_data(data):
    if not data:
        return []

    row_count = len(data)
    col_count = 0
    for row in data:
        if len(row) > col_count:
            col_count = len(row)

    if row_count == 0 or col_count == 0:
        return []

    last_row = -1
    last_col = -1
    for row_index, row in enumerate(data):
        row_has_value = False
        for col_index, value in enumerate(row):
            if not _is_blank_cell(value):
                row_has_value = True
                if col_index > last_col:
                    last_col = col_index
        if row_has_value:
            last_row = row_index

    if last_row < 0 or last_col < 0:
        return []

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

    prepared = []
    for row_index in range(trimmed_rows):
        source_row = data[row_index]
        prepared_row = []
        for col_index in range(trimmed_cols):
            value = source_row[col_index] if col_index < len(source_row) else ""
            prepared_row.append(_to_cell_text(value))
        prepared.append(prepared_row)

    return prepared


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


def _read_excel_used_range(file_path, sheet_name):
    excel_app = None
    workbook = None
    worksheet = None
    sheets = None
    used_range = None
    data = []
    try:
        excel_app, workbook = _open_excel_workbook(file_path)
        sheets = _com_get(workbook, "Sheets")
        worksheet = None
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

        used_range = _com_get(worksheet, "UsedRange")
        values = _com_get(used_range, "Value2")
        data = _normalize_excel_values(values)
    finally:
        _release_com_object(sheets)
        _release_com_object(used_range)
        _release_com_object(worksheet)
        _close_excel_workbook(excel_app, workbook)

    return data


def _get_drafting_view_type_id():
    for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if view_type.ViewFamily == ViewFamily.Drafting:
            return view_type.Id
    return None


def _get_text_note_type():
    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):
        return tnt
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
        if isinstance(element, TextNote) or isinstance(element, CurveElement):
            element_ids.append(element.Id)
    if element_ids:
        doc.Delete(List[ElementId](element_ids))


def _build_drafting_table(view, data, excel_file=None, sheet_name=None):
    """
    Builds drafting table matching Excel structure as closely as possible.
    """

    if not data:
        return

    rows = len(data)
    cols = max(len(r) for r in data)

    text_type = _get_text_note_type()
    if text_type is None:
        raise RuntimeError("TextNoteType not found.")

    # ----------------------------------------
    # Try reading Excel formatting
    # ----------------------------------------
    col_widths = [1.5] * cols
    row_heights = [0.3] * rows

    if excel_file and sheet_name:
        try:
            excel_app, workbook = _open_excel_workbook(excel_file)
            sheets = _com_get(workbook, "Sheets")
            worksheet = _com_item(sheets, sheet_name)

            # Column widths
            for c in range(cols):
                col_obj = _com_get(worksheet, "Columns")
                col_item = _com_item(col_obj, c + 1)
                width = _com_get(col_item, "ColumnWidth")
                col_widths[c] = float(width) * 0.12  # scale factor
                _release_com_object(col_item)

            # Row heights
            for r in range(rows):
                row_obj = _com_get(worksheet, "Rows")
                row_item = _com_item(row_obj, r + 1)
                height = _com_get(row_item, "RowHeight")
                row_heights[r] = float(height) * 0.01  # scale factor
                _release_com_object(row_item)

            _release_com_object(worksheet)
            _release_com_object(sheets)
            _close_excel_workbook(excel_app, workbook)

        except Exception:
            pass

    origin = XYZ(0, 0, 0)

    # ----------------------------------------
    # Draw horizontal lines
    # ----------------------------------------
    y = origin.Y
    for r in range(rows + 1):
        start = XYZ(origin.X, y, 0)
        end = XYZ(origin.X + sum(col_widths), y, 0)
        doc.Create.NewDetailCurve(view, Line.CreateBound(start, end))
        if r < rows:
            y -= row_heights[r]

    # ----------------------------------------
    # Draw vertical lines
    # ----------------------------------------
    x = origin.X
    for c in range(cols + 1):
        start = XYZ(x, origin.Y, 0)
        end = XYZ(x, origin.Y - sum(row_heights), 0)
        doc.Create.NewDetailCurve(view, Line.CreateBound(start, end))
        if c < cols:
            x += col_widths[c]

    # ----------------------------------------
    # Add text
    # ----------------------------------------
    note_opts = TextNoteOptions(text_type.Id)

    y = origin.Y
    for r in range(rows):
        x = origin.X
        for c in range(cols):

            value = ""
            if c < len(data[r]) and data[r][c] is not None:
                value = str(data[r][c])

            text_point = XYZ(
                x + 0.05,
                y - row_heights[r] * 0.8,
                0
            )

            TextNote.Create(doc, view.Id, text_point, value, note_opts)

            x += col_widths[c]

        y -= row_heights[r]


def _activate_view(view):
    try:
        revit.uidoc.ActiveView = view
    except Exception as err:
        logger.info("Could not activate view {}: {}".format(view.Name, err))


def _get_category_by_name(name):
    for category in doc.Settings.Categories:
        if category.Name == name:
            return category
    return None


def _set_parameter_value(param, value):
    if param is None:
        return
    if value is None:
        value = ""
    try:
        if param.StorageType == StorageType.String:
            param.Set(str(value))
        elif param.StorageType == StorageType.Integer:
            param.Set(int(value))
        elif param.StorageType == StorageType.Double:
            param.Set(float(value))
        elif param.StorageType == StorageType.ElementId:
            param.Set(ElementId(int(value)))
    except Exception as err:
        logger.info("Failed to set parameter {}: {}".format(param.Definition.Name, err))


def _update_elements_from_excel(category, headers, data, id_index):
    for row in data[1:]:
        if id_index >= len(row):
            continue
        element_id_value = row[id_index]
        if element_id_value in (None, ""):
            continue
        try:
            element_id = ElementId(int(element_id_value))
        except Exception:
            continue
        element = doc.GetElement(element_id)
        if element is None or element.Category is None:
            continue
        if element.Category.Id != category.Id:
            continue
        for header_index, header in enumerate(headers):
            if header_index == id_index:
                continue
            if header_index >= len(row):
                continue
            param = element.LookupParameter(header)
            _set_parameter_value(param, row[header_index])


class ExcelLinkWindow(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, "WPFWindow.xaml")
        self.rb_drafting.IsChecked = True
        self._toggle_schedule_controls()

    def _toggle_schedule_controls(self):
        return

    def _selected_sheet_name(self):
        selected = self.cb_sheet.SelectedItem
        if selected is None:
            return ""
        return str(selected).strip()

    def _set_default_target_name(self):
        sheet_name = self._selected_sheet_name()
        if not sheet_name:
            return
        self.tb_target_name.Text = sheet_name

    def browse_click(self, sender, e):
        dialog = OpenFileDialog()
        dialog.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
        if dialog.ShowDialog() == DialogResult.OK:
            self.tb_excel_path.Text = dialog.FileName
            self._load_sheet_names(dialog.FileName)

    def output_changed(self, sender, e):
        self._toggle_schedule_controls()

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
                category_name = ""
            else:
                file_path = str(self.tb_excel_path.Text).strip()
                sheet_name = self._selected_sheet_name()
                target_type = "Drafting"
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

            data = _read_excel_used_range(file_path, sheet_name)
            data = _prepare_table_data(data)
            if not data:
                forms.alert("No data found in the selected sheet.", title="Excel Import")
                return

            transaction = Transaction(doc, "Excel to Drafting View")
            transaction.Start()
            try:
                view = _get_or_create_drafting_view(target_name)
                _clear_view_contents(view)
                _build_drafting_table(view, data, file_path, sheet_name)
                transaction.Commit()
                output_view = view
                output_mode = "Drafting View"
            except Exception as first_error:
                try:
                    view = _create_new_drafting_view(target_name)
                    _build_drafting_table(view, data, file_path, sheet_name)
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

