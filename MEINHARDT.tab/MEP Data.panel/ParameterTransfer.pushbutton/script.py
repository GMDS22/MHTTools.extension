# coding: utf8
from __future__ import print_function

import os
import tempfile

from Autodesk.Revit.DB import (
    BuiltInParameter,
    CategorySet,
    ElementId,
    ExternalDefinitionCreationOptions,
    InstanceBinding,
    StorageType,
)

# BuiltInParameterGroup was removed in Revit 2022; import defensively.
try:
    from Autodesk.Revit.DB import BuiltInParameterGroup
except ImportError:
    BuiltInParameterGroup = None
from pyrevit import HOST_APP, forms, revit, script
from pyrevit.forms import WPFWindow

try:
    from Autodesk.Revit.DB import GroupTypeId
except Exception:
    GroupTypeId = None

try:
    from Autodesk.Revit.DB import SpecTypeId
except Exception:
    SpecTypeId = None

try:
    from Autodesk.Revit.DB import ParameterType
except Exception:
    ParameterType = None

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

__title__ = "Parameter Transfer"
__doc__ = "Copy or move parameter values from one parameter to another on selected elements."


def _get_element_type(element):
    try:
        type_id = element.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            return doc.GetElement(type_id)
    except Exception:
        return None
    return None


def _get_type_name(element_type):
    if element_type is None:
        return ""

    try:
        p = element_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p and p.AsString():
            return p.AsString()
    except Exception:
        pass

    try:
        p = element_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.AsString():
            return p.AsString()
    except Exception:
        pass

    try:
        return element_type.Name or ""
    except Exception:
        pass

    try:
        return str(element_type.Id.IntegerValue)
    except Exception:
        return ""


def _get_instance_type_name(element):
    try:
        p = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if p:
            # AsValueString is usually the displayed type name.
            v = p.AsValueString()
            if v:
                return v
            v = p.AsString()
            if v:
                return v
            try:
                tid = p.AsElementId()
                if tid and tid != ElementId.InvalidElementId:
                    t = doc.GetElement(tid)
                    if t is not None:
                        return _get_type_name(t)
            except Exception:
                pass
    except Exception:
        pass
    return ""


def _get_family_name(element, element_type):
    try:
        p = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
        if p:
            v = p.AsValueString()
            if v:
                return v
            v = p.AsString()
            if v:
                return v
    except Exception:
        pass

    try:
        p = element.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
        if p:
            v = p.AsString()
            if v:
                return v
    except Exception:
        pass

    try:
        if element_type is not None and hasattr(element_type, "FamilyName"):
            return element_type.FamilyName or ""
    except Exception:
        pass

    return ""


def _value_to_string(value):
    if value is None:
        return ""
    if isinstance(value, ElementId):
        if value == ElementId.InvalidElementId:
            return ""
        ref_element = doc.GetElement(value)
        if ref_element is not None:
            try:
                return ref_element.Name or str(value.IntegerValue)
            except Exception:
                return str(value.IntegerValue)
        try:
            return str(value.IntegerValue)
        except Exception:
            return str(value)
    try:
        return str(value)
    except Exception:
        return ""


def _string_data_type():
    if HOST_APP.is_older_than(2023):
        return ParameterType.Text
    return SpecTypeId.String.Text


def _safe_data_type_key(parameter):
    definition = parameter.Definition
    try:
        return str(definition.GetDataType())
    except Exception:
        try:
            return str(definition.ParameterType)
        except Exception:
            return str(parameter.StorageType)


def _build_options_from_template_element(element):
    """Return source/destination options from one template element.

    Using a single template element avoids empty dropdowns when selection
    contains mixed categories that do not share identical parameter sets.
    """
    options = {}

    # Names covered by our virtual info fields — suppress real BuiltIn duplicates.
    _VIRTUAL_NAMES = {"Type Name", "Family Name"}

    def add_option(scope, parameter, display_suffix):
        if parameter is None or parameter.Definition is None:
            return
        if parameter.StorageType == StorageType.None:
            return

        name = parameter.Definition.Name
        # Skip real params whose concept is already cleanly exposed via virtual fields.
        if name in _VIRTUAL_NAMES:
            return

        storage = parameter.StorageType
        dtype = _safe_data_type_key(parameter)
        key = "{}|{}|{}|{}".format(scope, name, storage, dtype)
        if key in options:
            return

        options[key] = {
            "key": key,
            "name": name,
            "scope": scope,
            "display": "{} {}".format(name, display_suffix),
            "storage": storage,
            "dtype": dtype,
            "readonly": parameter.IsReadOnly,
            "virtual": False,
            "writable_all": not parameter.IsReadOnly,
        }

    for parameter in element.Parameters:
        add_option("instance", parameter, "[Instance]")

    element_type = _get_element_type(element)
    if element_type is not None:
        for parameter in element_type.Parameters:
            add_option("type", parameter, "[Type]")

    virtual_fields = [
        ("virtual|Family Name|string", "Family Name"),
        ("virtual|Type Name|string",   "Type Name"),
        ("virtual|Category|string",    "Category"),
    ]
    for key, label in virtual_fields:
        options[key] = {
            "key": key,
            "name": label,
            "scope": "virtual",
            # No suffix clutter — these are now the single canonical source.
            "display": label,
            "storage": StorageType.String,
            "dtype": "virtual:string",
            "readonly": True,
            "virtual": True,
            "writable_all": False,
        }

    by_display = {}
    for item in sorted(options.values(), key=lambda x: x["display"]):
        by_display[item["display"]] = item
    return by_display


def _get_param_on_element(element, meta):
    def pick_parameter(container):
        if container is None:
            return None
        try:
            candidates = list(container.GetParameters(meta["name"]))
        except Exception:
            candidates = []

        # Match writable+storage first to avoid duplicate-name readonly mismatches.
        for p in candidates:
            try:
                if p and p.StorageType == meta["storage"] and not p.IsReadOnly:
                    return p
            except Exception:
                continue

        # Then match storage regardless of readonly.
        for p in candidates:
            try:
                if p and p.StorageType == meta["storage"]:
                    return p
            except Exception:
                continue

        # Fallback to LookupParameter for built-ins that may not appear in GetParameters.
        try:
            p = container.LookupParameter(meta["name"])
            if p:
                return p
        except Exception:
            pass

        if candidates:
            return candidates[0]
        return None

    if meta["scope"] == "instance":
        return pick_parameter(element)
    if meta["scope"] == "type":
        element_type = _get_element_type(element)
        return pick_parameter(element_type)
    return None


def _get_option_raw_dtype(meta, element):
    if meta.get("virtual"):
        return None
    parameter = _get_param_on_element(element, meta)
    if parameter is None:
        return None
    try:
        return parameter.Definition.GetDataType()
    except Exception:
        try:
            return parameter.Definition.ParameterType
        except Exception:
            return None


def _read_option_value(element, meta):
    if meta.get("virtual"):
        element_type = _get_element_type(element)

        if meta["name"] == "Type Name":
            # element_type.Name is the most direct and reliable source for all
            # categories: FilledRegion, Wall, Floor, Door, Pipe, Duct, etc.
            if element_type is not None:
                try:
                    v = element_type.Name
                    if v:
                        return v
                except Exception:
                    pass
            # Fallbacks for unusual element types without a proper type object.
            v = _get_instance_type_name(element) or _get_type_name(element_type)
            if v:
                return v
            try:
                return element.Name or ""
            except Exception:
                pass
            return ""

        if meta["name"] == "Family Name":
            # FamilySymbol.FamilyName works for all loadable families.
            if element_type is not None:
                try:
                    v = element_type.FamilyName
                    if v:
                        return v
                except Exception:
                    pass
            # Fallback covers system families (walls, floors) and MEP elements.
            v = _get_family_name(element, element_type)
            if v:
                return v
            # Last resort: use category name (system families have no family name).
            try:
                return element.Category.Name if element.Category else ""
            except Exception:
                return ""

        if meta["name"] == "Category":
            try:
                return element.Category.Name if element.Category else ""
            except Exception:
                return ""
        return ""

    parameter = _get_param_on_element(element, meta)
    if parameter is None:
        return None
    return _get_param_value(parameter, meta["storage"])


def _write_option_value(element, meta, value):
    parameter = _get_param_on_element(element, meta)
    if parameter is None or parameter.IsReadOnly:
        return False
    _set_param_value(parameter, value, meta["storage"])
    return True


def _clear_option_value(element, meta):
    parameter = _get_param_on_element(element, meta)
    if parameter is None or parameter.IsReadOnly:
        return False
    _clear_param(parameter, meta["storage"])
    return True


def _set_param_value(parameter, value, storage_type):
    if storage_type == StorageType.String:
        parameter.Set(value if value is not None else "")
    elif storage_type == StorageType.Integer:
        parameter.Set(int(value) if value is not None else 0)
    elif storage_type == StorageType.Double:
        parameter.Set(float(value) if value is not None else 0.0)
    elif storage_type == StorageType.ElementId:
        parameter.Set(value if value is not None else ElementId.InvalidElementId)


def _get_param_value(parameter, storage_type):
    if storage_type == StorageType.String:
        return parameter.AsString()
    if storage_type == StorageType.Integer:
        return parameter.AsInteger()
    if storage_type == StorageType.Double:
        return parameter.AsDouble()
    if storage_type == StorageType.ElementId:
        return parameter.AsElementId()
    return None


def _clear_param(parameter, storage_type):
    _set_param_value(parameter, None, storage_type)


def _ensure_shared_parameter_file():
    app = HOST_APP.app
    sp_file = app.OpenSharedParameterFile()
    if sp_file:
        return sp_file

    temp_dir = tempfile.gettempdir()
    temp_file_path = os.path.join(temp_dir, "mhttools_shared_parameters.txt")
    if not os.path.exists(temp_file_path):
        with open(temp_file_path, "w"):
            pass
    app.SharedParametersFilename = temp_file_path
    return app.OpenSharedParameterFile()


def _default_group_id():
    if GroupTypeId is not None:
        return GroupTypeId.Data
    if BuiltInParameterGroup is not None:
        return BuiltInParameterGroup.PG_DATA
    raise RuntimeError("Neither GroupTypeId nor BuiltInParameterGroup is available in this Revit version.")


def _create_destination_parameter(name, source_meta, elements):
    sp_file = _ensure_shared_parameter_file()
    if sp_file is None:
        raise RuntimeError("Unable to access shared parameter file.")

    definition_group = sp_file.Groups.get_Item("MHTTools")
    if definition_group is None:
        definition_group = sp_file.Groups.Create("MHTTools")

    try:
        existing = definition_group.Definitions.get_Item(name)
    except Exception:
        existing = None
    if existing is not None:
        return existing

    if HOST_APP.is_older_than(2023):
        options = ExternalDefinitionCreationOptions(name, source_meta["raw_dtype"])
    else:
        options = ExternalDefinitionCreationOptions(name, source_meta["raw_dtype"])

    definition = definition_group.Definitions.Create(options)

    app = HOST_APP.app
    category_set = app.Create.NewCategorySet()  # type: CategorySet
    for element in elements:
        if element.Category and element.Category.AllowsBoundParameters:
            category_set.Insert(element.Category)

    if category_set.IsEmpty:
        raise RuntimeError("Selected element categories cannot host project parameters.")

    binding = app.Create.NewInstanceBinding(category_set)  # type: InstanceBinding
    with revit.Transaction("Create destination parameter"):
        binding_map = doc.ParameterBindings
        if binding_map[definition]:
            binding_map.ReInsert(definition, binding, _default_group_id())
        else:
            binding_map.Insert(definition, binding, _default_group_id())
    return definition


class ParameterTransferWindow(WPFWindow):
    def __init__(self, xaml_path, elements):
        WPFWindow.__init__(self, xaml_path)
        self.elements = elements
        self.template_element = elements[0]
        self.common_map = _build_options_from_template_element(self.template_element)
        self.source_names = sorted(self.common_map.keys())

        if not self.source_names:
            forms.alert("No common transferable parameters found on the selected elements.", exitscript=True)

        self.source_combo.ItemsSource = self.source_names
        self.source_combo.SelectedIndex = 0
        self._refresh_destinations()

    def _source_meta(self):
        source_name = self.source_combo.SelectedItem
        if source_name is None:
            return None
        meta = dict(self.common_map[source_name])
        meta["raw_dtype"] = _get_option_raw_dtype(meta, self.template_element)
        return meta

    def _refresh_destinations(self):
        source_name = self.source_combo.SelectedItem
        if source_name is None:
            self.destination_combo.ItemsSource = []
            return

        source_meta = self.common_map[source_name]
        force_string = bool(self.convert_to_string_chk.IsChecked)
        destinations = []
        for name, meta in self.common_map.items():
            if name == source_name:
                continue
            # Virtual (read-only info) fields are never valid destinations
            if meta.get("virtual"):
                continue
            if force_string:
                # Only String storage can receive a string value
                if meta["storage"] != StorageType.String:
                    continue
            else:
                # Match on StorageType only; dtype sub-types vary between
                # GetDataType() (ForgeTypeId) and legacy ParameterType so
                # comparing them would hide perfectly valid destinations.
                if meta["storage"] != source_meta["storage"]:
                    continue
            destinations.append(name)

        self.destination_combo.ItemsSource = sorted(destinations)
        if destinations:
            self.destination_combo.SelectedIndex = 0

    # noinspection PyUnusedLocal
    def source_changed(self, sender, e):
        self._refresh_destinations()

    # noinspection PyUnusedLocal
    def convert_to_string_changed(self, sender, e):
        self._refresh_destinations()

    # noinspection PyUnusedLocal
    def create_param_click(self, sender, e):
        new_name = (self.new_param_name.Text or "").strip()
        if not new_name:
            forms.alert("Please enter a destination parameter name.")
            return

        source_meta = self._source_meta()
        if source_meta is None:
            forms.alert("Unable to resolve source parameter type for creation.")
            return

        if bool(self.convert_to_string_chk.IsChecked) or source_meta.get("virtual"):
            source_meta["raw_dtype"] = _string_data_type()

        if source_meta.get("raw_dtype") is None:
            forms.alert("Unable to resolve source parameter type for creation.")
            return

        try:
            _create_destination_parameter(new_name, source_meta, self.elements)
        except Exception as create_error:
            forms.alert("Could not create parameter:\n{}".format(create_error))
            return

        self.common_map = _build_options_from_template_element(self.template_element)
        self.source_names = sorted(self.common_map.keys())
        self.source_combo.ItemsSource = self.source_names

        current_source = self.source_combo.SelectedItem
        if current_source not in self.source_names:
            self.source_combo.SelectedIndex = 0

        self._refresh_destinations()
        destination_names = list(self.destination_combo.ItemsSource)
        preferred = "{} [Instance]".format(new_name)
        if preferred in destination_names:
            self.destination_combo.SelectedItem = preferred
        elif new_name in destination_names:
            self.destination_combo.SelectedItem = new_name
        forms.alert("Destination parameter created and added to selection categories.")

    def _transfer_values(self, source_name, destination_name, move_values=False, force_string=False):
        source_meta = self.common_map[source_name]
        destination_meta = self.common_map[destination_name]

        copied = 0
        moved = 0
        skipped = 0
        missing_source = 0
        missing_destination = 0
        readonly_destination = 0
        empty_source_value = 0

        with revit.Transaction("Parameter Transfer"):
            for element in self.elements:
                if not source_meta.get("virtual") and _get_param_on_element(element, source_meta) is None:
                    missing_source += 1
                    skipped += 1
                    continue

                dest_param = _get_param_on_element(element, destination_meta)
                if dest_param is None:
                    missing_destination += 1
                    skipped += 1
                    continue
                if dest_param.IsReadOnly:
                    readonly_destination += 1
                    skipped += 1
                    continue

                value = _read_option_value(element, source_meta)
                if force_string:
                    value = _value_to_string(value)
                if isinstance(value, str):
                    value = value.strip()
                if value in (None, ""):
                    empty_source_value += 1

                if _write_option_value(element, destination_meta, value):
                    copied += 1
                else:
                    skipped += 1
                    continue

                if move_values and not source_meta.get("virtual"):
                    if _clear_option_value(element, source_meta):
                        moved += 1

        if copied == 0:
            forms.alert("Nothing was copied. Check that the destination parameter exists and is writable on the selected elements.", exitscript=False)
        else:
            forms.alert("{} element(s) updated.".format(copied), exitscript=False)

    # noinspection PyUnusedLocal
    def apply_click(self, sender, e):
        source_name = self.source_combo.SelectedItem
        destination_name = self.destination_combo.SelectedItem
        if source_name is None:
            forms.alert("Please select a source parameter.")
            return
        if destination_name is None:
            if self.create_dest_chk.IsChecked and (self.new_param_name.Text or "").strip():
                self.create_param_click(sender, e)
                destination_name = self.destination_combo.SelectedItem
            if destination_name is None:
                forms.alert("Please select or create a destination parameter.")
                return

        move_values = bool(self.move_values_chk.IsChecked)
        force_string = bool(self.convert_to_string_chk.IsChecked)
        self._transfer_values(
            source_name,
            destination_name,
            move_values=move_values,
            force_string=force_string,
        )
        self.Close()

    # noinspection PyUnusedLocal
    def cancel_click(self, sender, e):
        self.Close()


def main():
    selected_ids = list(uidoc.Selection.GetElementIds())
    if not selected_ids:
        forms.alert("Select elements first, then run this tool.", exitscript=True)

    elements = [doc.GetElement(eid) for eid in selected_ids]
    elements = [e for e in elements if e is not None]

    if not elements:
        forms.alert("No valid selected elements found.", exitscript=True)

    xaml_path = os.path.join(os.path.dirname(__file__), "ParameterTransfer.xaml")
    ParameterTransferWindow(xaml_path, elements).ShowDialog()


if __name__ == "__main__":
    main()
