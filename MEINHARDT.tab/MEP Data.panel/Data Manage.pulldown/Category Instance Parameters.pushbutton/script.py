# coding: utf8
from __future__ import print_function

from collections import defaultdict

from Autodesk.Revit.DB import (
    BuiltInParameter,
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
)
from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow
from System.Windows.Controls import CheckBox
import System


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
config = script.get_config()

__title__ = "Category\nInstance Params"
__doc__ = "Batch update instance parameter values for multiple categories at once."


TARGET_CATEGORIES = [
    ("Spaces", BuiltInCategory.OST_MEPSpaces),
    ("Rooms", BuiltInCategory.OST_Rooms),
    ("HVAC Zones", BuiltInCategory.OST_HVAC_Zones),
    ("Ducts", BuiltInCategory.OST_DuctCurves),
    ("Duct Fittings", BuiltInCategory.OST_DuctFitting),
    ("Duct Accessories", BuiltInCategory.OST_DuctAccessory),
    ("Duct Insulations", BuiltInCategory.OST_DuctInsulations),
    ("Flex Ducts", BuiltInCategory.OST_FlexDuctCurves),
    ("Air Terminals", BuiltInCategory.OST_DuctTerminal),
    ("Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment),
    ("Pipes", BuiltInCategory.OST_PipeCurves),
    ("Pipe Fittings", BuiltInCategory.OST_PipeFitting),
    ("Pipe Accessories", BuiltInCategory.OST_PipeAccessory),
    ("Pipe Insulations", BuiltInCategory.OST_PipeInsulations),
    ("Flex Pipes", BuiltInCategory.OST_FlexPipeCurves),
    ("Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures),
    ("Sprinklers", BuiltInCategory.OST_Sprinklers),
    ("Cable Trays", BuiltInCategory.OST_CableTray),
    ("Cable Tray Fittings", BuiltInCategory.OST_CableTrayFitting),
    ("Conduits", BuiltInCategory.OST_Conduit),
    ("Conduit Fittings", BuiltInCategory.OST_ConduitFitting),
    ("Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment),
    ("Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures),
    ("Lighting Fixtures", BuiltInCategory.OST_LightingFixtures),
    ("Lighting Devices", BuiltInCategory.OST_LightingDevices),
    ("Data Devices", BuiltInCategory.OST_DataDevices),
    ("Communication Devices", BuiltInCategory.OST_CommunicationDevices),
    ("Fire Alarm Devices", BuiltInCategory.OST_FireAlarmDevices),
    ("Nurse Call Devices", BuiltInCategory.OST_NurseCallDevices),
    ("Security Devices", BuiltInCategory.OST_SecurityDevices),
    ("Telephone Devices", BuiltInCategory.OST_TelephoneDevices),
    ("Generic Models", BuiltInCategory.OST_GenericModel),
    ("Speciality Equipment", BuiltInCategory.OST_SpecialityEquipment),
    ("Casework", BuiltInCategory.OST_Casework),
    ("Furniture", BuiltInCategory.OST_Furniture),
    ("Furniture Systems", BuiltInCategory.OST_FurnitureSystems),
    ("Medical Equipment", BuiltInCategory.OST_MedicalEquipment),
]


def _storage_label(storage_type):
    if storage_type == StorageType.String:
        return "String"
    if storage_type == StorageType.Integer:
        return "Integer"
    if storage_type == StorageType.Double:
        return "Double"
    if storage_type == StorageType.ElementId:
        return "ElementId"
    return str(storage_type)


def _load_settings():
    return {
        "selected_categories": getattr(config, "selected_categories", []),
        "collect_scope": getattr(config, "collect_scope", "Active View Only"),
        "selected_level_id": getattr(config, "selected_level_id", -1),
        "duplicate_mode": getattr(config, "duplicate_mode", "Overwrite"),
    }


def _save_settings(settings):
    try:
        config.selected_categories = settings.get("selected_categories", [])
        config.collect_scope = settings.get("collect_scope", "Active View Only")
        config.selected_level_id = settings.get("selected_level_id", -1)
        config.duplicate_mode = settings.get("duplicate_mode", "Overwrite")
        script.save_config()
    except Exception:
        pass


def _read_parameter_value(param):
    if param is None or not param.HasValue:
        return None

    try:
        if param.StorageType == StorageType.String:
            return param.AsString()
        if param.StorageType == StorageType.Integer:
            return param.AsInteger()
        if param.StorageType == StorageType.Double:
            return param.AsDouble()
        if param.StorageType == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None:
                return None
            return eid.IntegerValue
    except Exception:
        return None

    return None


def _coerce_value(value_text, storage_type):
    text = "" if value_text is None else str(value_text).strip()

    if storage_type == StorageType.String:
        return text

    if storage_type == StorageType.Integer:
        if not text:
            raise ValueError("Integer value is required.")
        try:
            return int(text)
        except Exception:
            return int(float(text))

    if storage_type == StorageType.Double:
        if not text:
            raise ValueError("Double value is required.")
        return float(text)

    if storage_type == StorageType.ElementId:
        if not text:
            raise ValueError("ElementId integer value is required.")
        return ElementId(int(text))

    raise ValueError("Unsupported storage type.")


def _clear_parameter(param, storage_type):
    if storage_type == StorageType.String:
        param.Set("")
        return True
    if storage_type == StorageType.Integer:
        param.Set(0)
        return True
    if storage_type == StorageType.Double:
        param.Set(0.0)
        return True
    if storage_type == StorageType.ElementId:
        param.Set(ElementId.InvalidElementId)
        return True
    return False


def _set_parameter_value(param, coerced_value, storage_type, duplicate_mode):
    existing = _read_parameter_value(param)

    if storage_type == StorageType.String:
        if existing not in (None, ""):
            if duplicate_mode == "Skip Existing":
                return False, "skipped existing"
            if duplicate_mode == "Append (String only)":
                old_text = str(existing)
                new_text = str(coerced_value)
                success = param.Set(old_text + "; " + new_text)
                return bool(success), "appended" if success else "failed"

        success = param.Set(str(coerced_value))
        return bool(success), "written" if success else "failed"

    if storage_type == StorageType.Integer:
        success = param.Set(int(coerced_value))
        return bool(success), "written" if success else "failed"

    if storage_type == StorageType.Double:
        success = param.Set(float(coerced_value))
        return bool(success), "written" if success else "failed"

    if storage_type == StorageType.ElementId:
        success = param.Set(coerced_value)
        return bool(success), "written" if success else "failed"

    return False, "unsupported storage"


def _find_writable_instance_parameter(element, parameter_name, storage_type):
    if element is None or not parameter_name:
        return None

    target_name = parameter_name.strip()
    storage_none = getattr(StorageType, "None")

    try:
        for param in element.Parameters:
            if param is None or param.Definition is None:
                continue
            if param.Definition.Name != target_name:
                continue
            if param.IsReadOnly:
                continue
            if param.StorageType == storage_none:
                continue
            if param.StorageType != storage_type:
                continue
            return param
    except Exception:
        return None

    return None


def _collect_instance_parameter_map(element):
    parameters = {}
    storage_none = getattr(StorageType, "None")

    if element is None:
        return parameters

    try:
        for param in element.Parameters:
            if param is None or param.Definition is None:
                continue
            if param.IsReadOnly:
                continue
            if param.StorageType == storage_none:
                continue

            name = param.Definition.Name
            if not name:
                continue

            if name not in parameters:
                parameters[name] = param.StorageType
    except Exception:
        pass

    return parameters


def _get_element_level_id(element):
    if element is None:
        return None

    invalid_id = ElementId.InvalidElementId

    try:
        lvl_id = element.LevelId
        if lvl_id is not None and lvl_id != invalid_id and lvl_id.IntegerValue > 0:
            return lvl_id.IntegerValue
    except Exception:
        pass

    param_ids = [
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.RBS_START_LEVEL_PARAM,
    ]
    for bip in param_ids:
        try:
            p = element.get_Parameter(bip)
            if p is None:
                continue
            lvl_id = p.AsElementId()
            if lvl_id is not None and lvl_id != invalid_id and lvl_id.IntegerValue > 0:
                return lvl_id.IntegerValue
        except Exception:
            continue

    return None


def _iter_category_elements(category_ids, active_view_only=True, hard_limit=50000, level_id=None):
    count = 0
    seen_ids = set()

    for cat_id in category_ids:
        try:
            if active_view_only:
                collector = FilteredElementCollector(doc, doc.ActiveView.Id)
            else:
                collector = FilteredElementCollector(doc)

            elems = collector.OfCategoryId(ElementId(cat_id)).WhereElementIsNotElementType().ToElements()
        except Exception:
            continue

        for element in elems:
            if element is None:
                continue

            if level_id is not None and _get_element_level_id(element) != int(level_id):
                continue

            eid = element.Id.IntegerValue
            if eid in seen_ids:
                continue
            seen_ids.add(eid)
            yield element

            count += 1
            if count >= hard_limit:
                return


class CategoryInstanceParameterWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)

        self._ui_ready = False
        self.category_items = []
        self.level_items = []
        self.selected_elements = []
        self.common_param_map = {}
        self._settings = _load_settings()

        self._build_level_list()
        self._build_category_list()
        self._apply_saved_settings()
        self._update_scope_dependent_ui()
        self._refresh_element_summary()
        self._refresh_parameter_controls()
        self._ui_ready = True

    def _confirm(self, message, title="Confirm"):
        try:
            result = forms.alert(message, title=title, yes=True, no=True)
            if isinstance(result, bool):
                return result
            text = (str(result) if result is not None else "").strip().lower()
            return text in ("yes", "y", "true", "ok", "1")
        except Exception:
            return False

    def _build_category_list(self):
        selected = set(self._settings.get("selected_categories", []))
        self.category_items = []

        for name, bic in TARGET_CATEGORIES:
            cb = CheckBox()
            cb.Content = name
            cb.IsChecked = True if not selected else (int(bic) in selected)

            self.category_items.append({
                "name": name,
                "bic": int(bic),
                "cb": cb,
            })

        self._refresh_category_list("")

    def _apply_saved_settings(self):
        scope_name = self._settings.get("collect_scope", "Active View Only")
        if "Selected Level" in scope_name:
            self.cmbCollectScope.SelectedIndex = 1
        elif "Entire" in scope_name:
            self.cmbCollectScope.SelectedIndex = 2
        else:
            self.cmbCollectScope.SelectedIndex = 0

        saved_level_id = int(self._settings.get("selected_level_id", -1) or -1)
        if self.level_items:
            selected_idx = -1
            for idx, item in enumerate(self.level_items):
                if item["id"] == saved_level_id:
                    selected_idx = idx
                    break
            if selected_idx < 0:
                selected_idx = 0
            self.cmbSelectLevel.SelectedIndex = selected_idx

        mode_name = self._settings.get("duplicate_mode", "Overwrite")
        names = ["Overwrite", "Skip Existing", "Append (String only)"]
        if mode_name in names:
            self.cmbDuplicateMode.SelectedIndex = names.index(mode_name)
        else:
            self.cmbDuplicateMode.SelectedIndex = 0

    def _save_current_settings(self):
        settings = {
            "selected_categories": self._selected_category_ids(),
            "collect_scope": self._selected_scope(),
            "selected_level_id": self._selected_level_id(),
            "duplicate_mode": self._selected_duplicate_mode(),
        }
        _save_settings(settings)

    def _build_level_list(self):
        self.level_items = []
        self.cmbSelectLevel.Items.Clear()

        try:
            levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        except Exception:
            levels = []

        sortable = []
        for lvl in levels:
            try:
                sortable.append((float(lvl.Elevation), lvl))
            except Exception:
                sortable.append((0.0, lvl))

        for _, lvl in sorted(sortable, key=lambda x: (x[0], x[1].Name.lower() if hasattr(x[1], "Name") else "")):
            try:
                item = {"id": lvl.Id.IntegerValue, "name": lvl.Name}
            except Exception:
                continue
            self.level_items.append(item)
            self.cmbSelectLevel.Items.Add(item["name"])

    def _selected_level_id(self):
        idx = getattr(self.cmbSelectLevel, "SelectedIndex", -1)
        if idx is None or idx < 0 or idx >= len(self.level_items):
            return -1
        return int(self.level_items[idx]["id"])

    def _update_scope_dependent_ui(self):
        panel = getattr(self, "pnlSelectedLevel", None)
        if panel is None:
            return
        scope_name = self._selected_scope()
        show_level = "Selected Level" in scope_name
        panel.Visibility = System.Windows.Visibility.Visible if show_level else System.Windows.Visibility.Collapsed

    def _refresh_category_list(self, search_text):
        self.lstCategories.Items.Clear()

        query = (search_text or "").strip().lower()
        for item in self.category_items:
            if query and query not in item["name"].lower():
                continue
            self.lstCategories.Items.Add(item["cb"])

    def _selected_category_ids(self):
        ids = []
        for item in self.category_items:
            try:
                if item["cb"].IsChecked:
                    ids.append(item["bic"])
            except Exception:
                continue
        return ids

    def _selected_scope(self):
        try:
            selected_item = self.cmbCollectScope.SelectedItem
            if selected_item is None:
                return "Active View Only"
            text = str(selected_item.Content) if hasattr(selected_item, "Content") else str(selected_item)
            return text
        except Exception:
            return "Active View Only"

    def _selected_duplicate_mode(self):
        try:
            selected_item = self.cmbDuplicateMode.SelectedItem
            if selected_item is None:
                return "Overwrite"
            text = str(selected_item.Content) if hasattr(selected_item, "Content") else str(selected_item)
            return text
        except Exception:
            return "Overwrite"

    def _refresh_element_summary(self):
        if not self.selected_elements:
            self.txtElementSummary.Text = "No elements collected yet."
            return

        by_cat = defaultdict(int)
        for element in self.selected_elements:
            try:
                cat_name = element.Category.Name if element.Category else "<No Category>"
            except Exception:
                cat_name = "<No Category>"
            by_cat[cat_name] += 1

        lines = ["Collected elements: {0}".format(len(self.selected_elements))]
        for cat_name in sorted(by_cat.keys()):
            lines.append("- {0}: {1}".format(cat_name, by_cat[cat_name]))

        self.txtElementSummary.Text = "\n".join(lines)

    def _refresh_parameter_controls(self):
        self.lstCommonParams.Items.Clear()
        self.cmbParameter.Items.Clear()
        self.common_param_map = {}

        if not self.selected_elements:
            self.txtSelectedParameterInfo.Text = "No parameter selected."
            return

        param_maps = [_collect_instance_parameter_map(el) for el in self.selected_elements]
        if not param_maps:
            self.txtSelectedParameterInfo.Text = "No parameter selected."
            return

        common = dict(param_maps[0])
        for pmap in param_maps[1:]:
            next_common = {}
            for name, storage in common.items():
                if name in pmap and pmap[name] == storage:
                    next_common[name] = storage
            common = next_common
            if not common:
                break

        self.common_param_map = common

        for param_name in sorted(common.keys()):
            storage_label = _storage_label(common[param_name])
            self.lstCommonParams.Items.Add("{0} | {1}".format(param_name, storage_label))
            self.cmbParameter.Items.Add(param_name)

        if self.cmbParameter.Items.Count > 0:
            self.cmbParameter.SelectedIndex = 0
        else:
            self.txtSelectedParameterInfo.Text = "No common writable instance parameters found for this element set."

    def category_search_changed(self, sender, e):
        self._refresh_category_list(self.txtCategorySearch.Text)

    def select_all_categories_click(self, sender, e):
        for item in self.category_items:
            item["cb"].IsChecked = True

    def deselect_all_categories_click(self, sender, e):
        for item in self.category_items:
            item["cb"].IsChecked = False

    def collect_scope_changed(self, sender, e):
        if not getattr(self, "_ui_ready", False):
            return
        self._update_scope_dependent_ui()
        self._save_current_settings()

    def collect_elements_click(self, sender, e):
        category_ids = self._selected_category_ids()
        if not category_ids:
            forms.alert("Select at least one category first.")
            return

        scope_name = self._selected_scope()
        active_view_only = "Active" in scope_name
        selected_level_mode = "Selected Level" in scope_name
        selected_level_id = None

        if selected_level_mode:
            selected_level_id = self._selected_level_id()
            if selected_level_id < 0:
                forms.alert("Select a level first.")
                return

        if "Entire" in scope_name:
            proceed = self._confirm(
                "Collecting from Entire Model can take longer on large projects. Continue?",
                title="Large Collection Warning"
            )
            if not proceed:
                return

        elements = list(
            _iter_category_elements(
                category_ids,
                active_view_only=active_view_only,
                hard_limit=50000,
                level_id=selected_level_id,
            )
        )
        if len(elements) >= 50000:
            forms.alert("Element collection reached safety limit (50,000). Narrow categories or use Active View scope.")

        self.selected_elements = elements
        self._refresh_element_summary()
        self._refresh_parameter_controls()
        self._save_current_settings()

    def use_current_selection_click(self, sender, e):
        category_ids = set(self._selected_category_ids())
        if not category_ids:
            forms.alert("Select at least one category first.")
            return

        selected_ids = uidoc.Selection.GetElementIds()
        if not selected_ids:
            forms.alert("No elements are selected in Revit.")
            return

        elements = []
        for eid in selected_ids:
            element = doc.GetElement(eid)
            if element is None or element.Category is None:
                continue
            if element.Category.Id.IntegerValue not in category_ids:
                continue
            elements.append(element)

        self.selected_elements = elements
        self._refresh_element_summary()
        self._refresh_parameter_controls()
        self._save_current_settings()

    def parameter_selection_changed(self, sender, e):
        parameter_name = self.cmbParameter.SelectedItem
        if parameter_name is None:
            self.txtSelectedParameterInfo.Text = "No parameter selected."
            return

        storage_type = self.common_param_map.get(parameter_name)
        if storage_type is None:
            self.txtSelectedParameterInfo.Text = "No parameter selected."
            return

        self.txtSelectedParameterInfo.Text = (
            "Selected parameter: {0}\n"
            "Storage type: {1}\n"
            "Note: this tool writes instance parameters only."
        ).format(parameter_name, _storage_label(storage_type))

    def clear_value_changed(self, sender, e):
        clear_mode = bool(self.chkClearValue.IsChecked)
        self.txtValue.IsEnabled = not clear_mode

    def apply_update_click(self, sender, e):
        if not self.selected_elements:
            forms.alert("Collect or select elements before applying updates.")
            return

        parameter_name = self.cmbParameter.SelectedItem
        if parameter_name is None:
            forms.alert("Choose a target parameter first.")
            return

        storage_type = self.common_param_map.get(parameter_name)
        if storage_type is None:
            forms.alert("Selected parameter is not available.")
            return

        clear_mode = bool(self.chkClearValue.IsChecked)
        value_text = self.txtValue.Text if self.txtValue is not None else ""
        duplicate_mode = self._selected_duplicate_mode()

        coerced_value = None
        if not clear_mode:
            try:
                coerced_value = _coerce_value(value_text, storage_type)
            except Exception as ex:
                forms.alert("Invalid value for {0}: {1}".format(_storage_label(storage_type), ex))
                return

        element_count = len(self.selected_elements)
        if element_count > 2000:
            proceed = self._confirm(
                "You are about to update {0} elements. Continue?".format(element_count),
                title="Bulk Update Confirmation"
            )
            if not proceed:
                return

        written = 0
        skipped = 0
        failed = 0
        fail_samples = []

        tx = Transaction(doc, "Batch Update Category Instance Parameters")
        tx.Start()
        try:
            for element in self.selected_elements:
                param = _find_writable_instance_parameter(element, parameter_name, storage_type)
                if param is None:
                    skipped += 1
                    continue

                try:
                    if clear_mode:
                        ok = _clear_parameter(param, storage_type)
                        if ok:
                            written += 1
                        else:
                            failed += 1
                            if len(fail_samples) < 10:
                                fail_samples.append("Element {0}: clear failed".format(element.Id.IntegerValue))
                        continue

                    ok, message = _set_parameter_value(param, coerced_value, storage_type, duplicate_mode)
                    if ok:
                        written += 1
                    elif message == "skipped existing":
                        skipped += 1
                    else:
                        failed += 1
                        if len(fail_samples) < 10:
                            fail_samples.append("Element {0}: {1}".format(element.Id.IntegerValue, message))
                except Exception as ex:
                    failed += 1
                    if len(fail_samples) < 10:
                        fail_samples.append("Element {0}: {1}".format(element.Id.IntegerValue, ex))

            tx.Commit()
        except Exception as ex:
            tx.RollBack()
            forms.alert("Update failed and was rolled back.\n\n{0}".format(ex))
            return

        summary_lines = [
            "Batch update complete.",
            "Parameter: {0}".format(parameter_name),
            "Elements processed: {0}".format(element_count),
            "Written: {0}".format(written),
            "Skipped: {0}".format(skipped),
            "Failed: {0}".format(failed),
        ]

        if fail_samples:
            summary_lines.append("")
            summary_lines.append("Sample failures:")
            summary_lines.extend(["- {0}".format(item) for item in fail_samples])

        forms.alert("\n".join(summary_lines), title="Category Instance Parameters")
        self._save_current_settings()

    def close_click(self, sender, e):
        self._save_current_settings()
        self.Close()


try:
    window = CategoryInstanceParameterWindow("WPFWindow.xaml")
    window.ShowDialog()
except Exception as ex:
    forms.alert("Unable to launch Category Instance Parameters tool.\n\n{0}".format(ex))
