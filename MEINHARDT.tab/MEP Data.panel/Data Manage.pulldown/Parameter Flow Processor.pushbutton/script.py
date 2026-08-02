# coding: utf8
from __future__ import print_function

import re
import time

import clr

from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
    TransactionGroup,
)
from pyrevit import forms, revit, script
from pyrevit.forms import WPFWindow
import System

try:
    clr.AddReference("System.Windows.Forms")
    from System.Windows.Forms import Application as WinFormsApplication
except Exception:
    WinFormsApplication = None


doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

__title__ = "Parameter Flow\nProcessor"
__doc__ = "Filter active-view elements by parameter values and process parameters with verified writes."

MAX_APPLY_LIMIT = 50000


try:
    _TEXT_TYPE = unicode
except NameError:
    _TEXT_TYPE = str


def _safe_str(value):
    try:
        if value is None:
            return ""
        return str(value)
    except Exception:
        try:
            return _TEXT_TYPE(value)
        except Exception:
            return ""


def _is_valid_element(element):
    if element is None:
        return False
    try:
        return bool(getattr(element, "IsValidObject", True))
    except Exception:
        return False


def _safe_element_id(element):
    if not _is_valid_element(element):
        return "<invalid>"
    try:
        return str(element.Id.IntegerValue)
    except Exception:
        return "<unknown>"


def _iter_element_params(element):
    try:
        for param in element.Parameters:
            yield param
    except Exception:
        return


def _find_param_on_element(element, param_name):
    if not _is_valid_element(element) or not param_name:
        return None

    try:
        p = element.LookupParameter(param_name)
        if p:
            return p
    except Exception:
        pass

    target = _safe_str(param_name).lower()
    for p in _iter_element_params(element):
        try:
            name = p.Definition.Name
            if name and _safe_str(name).lower() == target:
                return p
        except Exception:
            continue
    return None


def _read_parameter_value(param):
    if param is None:
        return None

    try:
        if param.StorageType == StorageType.String:
            raw = param.AsString()
            return "" if raw is None else raw
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
        pass

    try:
        val = param.AsValueString()
        if val is not None:
            return val
    except Exception:
        pass

    return None


def _coerce_readback_value(param, storage_type):
    if param is None:
        return None

    try:
        if storage_type == StorageType.String:
            raw = param.AsString()
            if raw is not None:
                return raw
            disp = param.AsValueString()
            if disp is not None:
                return disp
            return ""
        if storage_type == StorageType.Integer:
            return param.AsInteger()
        if storage_type == StorageType.Double:
            return param.AsDouble()
        if storage_type == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None:
                return None
            return eid.IntegerValue
    except Exception:
        pass

    return _read_parameter_value(param)


def _values_match(left_value, right_value, storage_type):
    if storage_type == StorageType.String:
        return ("" if left_value is None else str(left_value)) == ("" if right_value is None else str(right_value))

    if storage_type == StorageType.Integer:
        try:
            return int(left_value) == int(right_value)
        except Exception:
            return False

    if storage_type == StorageType.Double:
        try:
            return abs(float(left_value) - float(right_value)) <= 1e-6
        except Exception:
            return False

    if storage_type == StorageType.ElementId:
        try:
            return int(left_value) == int(right_value)
        except Exception:
            return False

    return _safe_str(left_value) == _safe_str(right_value)


def _display_parameter_value(param):
    if param is None:
        return "<missing>"

    try:
        st = param.StorageType
    except Exception:
        st = None

    try:
        if st == StorageType.String:
            raw = param.AsString()
            if raw is None:
                return "<empty>"
            text = _safe_str(raw)
            return text if text.strip() else "<empty>"
        if st == StorageType.Integer:
            vs = param.AsValueString()
            if vs:
                return _safe_str(vs)
            return _safe_str(param.AsInteger())
        if st == StorageType.Double:
            vs = param.AsValueString()
            if vs:
                return _safe_str(vs)
            return _safe_str(param.AsDouble())
        if st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None or eid == ElementId.InvalidElementId:
                return "<empty>"
            return _safe_str(eid.IntegerValue)
    except Exception:
        pass

    return "<empty>"


def _collect_current_view_elements(document):
    try:
        return list(FilteredElementCollector(document, document.ActiveView.Id).WhereElementIsNotElementType())
    except Exception:
        return []


def _collect_common_maps(elements):
    readable_union = {}
    writable_union = {}
    storage_none = getattr(StorageType, "None")

    for element in elements:
        if not _is_valid_element(element):
            continue

        for param in _iter_element_params(element):
            try:
                if param is None or param.Definition is None:
                    continue
                name = param.Definition.Name
                if not name:
                    continue
                st = param.StorageType
                if st == storage_none:
                    continue
                if name not in readable_union:
                    readable_union[name] = st
                if not param.IsReadOnly and name not in writable_union:
                    writable_union[name] = st
            except Exception:
                continue

    return readable_union, writable_union


def _param_value_key_display(param):
    if param is None:
        return ("<missing>", "<missing>")

    try:
        if not param.HasValue:
            return ("<empty>", "<empty>")
    except Exception:
        pass

    try:
        st = param.StorageType
    except Exception:
        st = None

    try:
        if st == StorageType.String:
            raw = param.AsString()
            raw = "" if raw is None else _safe_str(raw)
            return (("s", raw), raw if raw.strip() else "<empty>")
        if st == StorageType.Integer:
            val = param.AsInteger()
            disp = param.AsValueString()
            if not disp:
                disp = _safe_str(val)
            return (("i", _safe_str(val)), _safe_str(disp))
        if st == StorageType.Double:
            val = param.AsDouble()
            disp = param.AsValueString()
            if not disp:
                disp = _safe_str(val)
            return (("d", _safe_str(val)), _safe_str(disp))
        if st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid is None or eid == ElementId.InvalidElementId:
                return (("eid", "-1"), "<none>")
            return (("eid", _safe_str(eid.IntegerValue)), _safe_str(eid.IntegerValue))
    except Exception:
        pass

    text = _safe_str(_read_parameter_value(param))
    return (("u", text), text if text else "<empty>")


def _coerce_value(value_text, storage_type):
    text = "" if value_text is None else _safe_str(value_text).strip()

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


def _set_parameter_value(param, coerced_value, storage_type):
    if storage_type == StorageType.String:
        return bool(param.Set(_safe_str(coerced_value)))
    if storage_type == StorageType.Integer:
        return bool(param.Set(int(coerced_value)))
    if storage_type == StorageType.Double:
        return bool(param.Set(float(coerced_value)))
    if storage_type == StorageType.ElementId:
        return bool(param.Set(coerced_value))
    return False


def _is_effective_value_set(param, storage_type):
    val = _read_parameter_value(param)
    if storage_type == StorageType.String:
        return val not in (None, "")
    if storage_type == StorageType.ElementId:
        try:
            return int(val) != ElementId.InvalidElementId.IntegerValue
        except Exception:
            return False
    return val is not None


def _format_estimate(seconds_value):
    try:
        secs = float(seconds_value)
    except Exception:
        secs = 0.0
    return "~{0:.1f}s ({1} ms)".format(secs, int(round(secs * 1000.0)))


def _estimate_seconds_for_apply(count, source_mode):
    if count <= 0:
        return 0.5
    per_item = 0.026 if source_mode else 0.017
    return max(0.5, 0.5 + (float(count) * per_item))


def _selected_combo_text(combo):
    try:
        item = combo.SelectedItem
        if item is None:
            return ""
        return _safe_str(item.Content) if hasattr(item, "Content") else _safe_str(item)
    except Exception:
        return ""


def _parse_dictionary_entries(text_value):
    raw = "" if text_value is None else _safe_str(text_value)
    mapping = {}
    for line in raw.splitlines():
        item = line.strip()
        if not item or item.startswith("#"):
            continue
        sep = "=>" if "=>" in item else ("=" if "=" in item else None)
        if sep is None:
            continue
        left, right = item.split(sep, 1)
        key = left.strip().lower()
        val = right.strip()
        if key:
            mapping[key] = val
    return mapping


def _transform_value(raw_value, transform_name, dict_text):
    text = "" if raw_value is None else _safe_str(raw_value)

    if transform_name == "Keep Original":
        return raw_value
    if transform_name == "Uppercase":
        return text.upper()
    if transform_name == "Lowercase":
        return text.lower()
    if transform_name == "Trim Spaces":
        return text.strip()
    if transform_name == "Digits Only":
        return re.sub(r"\D", "", text)
    if transform_name == "Dictionary Translate":
        mapping = _parse_dictionary_entries(dict_text)
        if not mapping:
            return raw_value
        key = text.strip().lower()
        return mapping.get(key, raw_value)

    return raw_value


class ProcessingMonitorWindow(WPFWindow):
    def __init__(self, xaml_file_name, owner_window):
        WPFWindow.__init__(self, xaml_file_name)
        self.owner_window = owner_window
        self._row_ids = []
        self._max_rows = 5000

    def set_status(self, text):
        try:
            self.txtMonitorStatus.Text = _safe_str(text)
        except Exception:
            pass

    def append_row(self, element_id, status, category_name, rule_text, message):
        eid_text = "-" if element_id is None else _safe_str(element_id)
        cat_text = _safe_str(category_name) if category_name else "-"
        rule_part = _safe_str(rule_text) if rule_text else "-"
        msg_part = _safe_str(message) if message else ""
        label = "ID {0} | {1} | Cat: {2} | Step: {3} | {4}".format(
            eid_text,
            _safe_str(status).upper(),
            cat_text,
            rule_part,
            msg_part,
        )

        self.lstMonitorItems.Items.Add(label)
        try:
            self._row_ids.append(int(element_id) if element_id is not None else None)
        except Exception:
            self._row_ids.append(None)

        while self.lstMonitorItems.Items.Count > self._max_rows:
            self.lstMonitorItems.Items.RemoveAt(0)
            if self._row_ids:
                del self._row_ids[0]

        try:
            idx = self.lstMonitorItems.Items.Count - 1
            if idx >= 0:
                self.lstMonitorItems.ScrollIntoView(self.lstMonitorItems.Items[idx])
        except Exception:
            pass

    def _get_selected_element_id(self):
        try:
            idx = int(self.lstMonitorItems.SelectedIndex)
        except Exception:
            idx = -1
        if idx < 0 or idx >= len(self._row_ids):
            return None
        return self._row_ids[idx]

    def _center_selected(self):
        eid = self._get_selected_element_id()
        if eid is None:
            self.set_status("Select a processed row with a valid ID first.")
            return
        if self.owner_window is None:
            self.set_status("Monitor owner is unavailable.")
            return
        ok, msg = self.owner_window.focus_element_from_monitor(eid)
        self.set_status(msg)
        if not ok:
            return

    def monitor_selection_changed(self, sender, e):
        self._center_selected()

    def center_selected_click(self, sender, e):
        self._center_selected()

    def clear_monitor_click(self, sender, e):
        self.lstMonitorItems.Items.Clear()
        self._row_ids = []
        self.set_status("Monitor list cleared.")

    def close_monitor_click(self, sender, e):
        self.Hide()


class ViewFilterParameterProcessorWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)

        self._ui_ready = False
        self._is_busy = False
        self._cancel_requested = False
        self._cancel_mode = None
        self._pending_focus_element_id = None
        self._monitor_window = None
        self._last_filter_map = {}

        self._common_readable_map = {}
        self._common_writable_map = {}

        self.all_view_elements = []
        self.current_pool_elements = []
        self.filtered_elements = []
        self.written_element_ids = set()

        self._initialize_from_active_view()

    def _initialize_from_active_view(self):
        self.all_view_elements = _collect_current_view_elements(doc)
        self.current_pool_elements = list(self.all_view_elements)
        self.filtered_elements = list(self.current_pool_elements)
        self._rebuild_parameter_maps()
        self._refresh_filter_parameter_combo()
        self._refresh_target_parameter_combo()
        self._refresh_source_parameter_combo()
        self._refresh_pool_summary()
        self._set_filter_summary("No filter applied yet.")
        self._set_run_detail("Current item: -")

    def _set_filter_summary(self, text):
        self.txtFilterSummary.Text = _safe_str(text)

    def _set_run_status(self, text):
        self.txtRunStatus.Text = _safe_str(text)

    def _set_run_detail(self, text):
        try:
            self.txtRunDetail.Text = _safe_str(text)
        except Exception:
            pass

    def _refresh_pool_summary(self):
        view_count = len(self.all_view_elements)
        pool_count = len(self.current_pool_elements)
        filtered_count = len(self.filtered_elements)
        written_count = len(self.written_element_ids)

        self.txtViewSummary.Text = (
            "Active view: {0} | All elements: {1} | Current pool: {2} | Current filtered set: {3} | Written in session: {4}"
        ).format(_safe_str(doc.ActiveView.Name), view_count, pool_count, filtered_count, written_count)

        self.txtPoolSummary.Text = "Pool: {0} | Filtered: {1} | Written this session: {2}".format(
            pool_count, filtered_count, written_count
        )

    def _refresh_counts(self, processed, written, skipped, failed):
        self.txtRunCounts.Text = "Processed: {0} | Written: {1} | Skipped: {2} | Failed: {3}".format(
            processed, written, skipped, failed
        )

    def _ensure_monitor_window(self):
        if self._monitor_window is not None:
            return self._monitor_window
        try:
            self._monitor_window = ProcessingMonitorWindow("ProcessingMonitorWindow.xaml", self)
        except Exception as ex:
            forms.alert("Unable to open processing monitor.\n\n{0}".format(ex), title="Parameter Flow Processor")
            self._monitor_window = None
        return self._monitor_window

    def _show_monitor_window(self):
        monitor = self._ensure_monitor_window()
        if monitor is None:
            return None
        try:
            monitor.Show()
            monitor.Activate()
        except Exception:
            pass
        return monitor

    def _monitor_set_status(self, text):
        if self._monitor_window is None:
            return
        try:
            self._monitor_window.set_status(text)
        except Exception:
            pass

    def _monitor_add_row(self, element, status, rule_text, message):
        if self._monitor_window is None:
            return
        eid = None
        category_name = ""
        try:
            if element is not None:
                eid = element.Id.IntegerValue
        except Exception:
            eid = None
        try:
            if element is not None and element.Category is not None:
                category_name = element.Category.Name
        except Exception:
            category_name = ""

        try:
            self._monitor_window.append_row(eid, status, category_name, rule_text, message)
        except Exception:
            pass

    def _drain_pending_focus_request(self):
        if self._pending_focus_element_id is None:
            return
        eid = self._pending_focus_element_id
        self._pending_focus_element_id = None
        ok, msg = self._focus_element_now(eid)
        self._monitor_set_status(msg)
        if ok:
            self._set_run_detail("Current item: focused element {0}".format(eid))

    def _focus_element_now(self, element_id):
        try:
            eid_int = int(element_id)
        except Exception:
            return False, "Invalid element ID: {0}".format(element_id)

        element = None
        try:
            element = doc.GetElement(ElementId(eid_int))
        except Exception:
            element = None

        if not _is_valid_element(element):
            return False, "Element {0} does not exist or is invalid.".format(eid_int)

        try:
            # Use ShowElements directly to avoid runtime issues with generic-list interop.
            uidoc.ShowElements(ElementId(eid_int))
            try:
                uidoc.RefreshActiveView()
            except Exception:
                pass
        except Exception as ex:
            return False, "Could not focus element {0}: {1}".format(eid_int, ex)

        return True, "Focused element {0} in active view.".format(eid_int)

    def focus_element_from_monitor(self, element_id):
        if self._is_busy:
            self._pending_focus_element_id = element_id
            return True, "Element {0} queued. It will be centered at the next safe checkpoint.".format(element_id)
        return self._focus_element_now(element_id)

    def _rebuild_parameter_maps(self):
        readable, writable = _collect_common_maps(self.current_pool_elements)
        self._common_readable_map = readable
        self._common_writable_map = writable

    def _refresh_filter_parameter_combo(self):
        self.cmbFilterParameter.Items.Clear()
        self.cmbFilterParameter.Items.Add("<Category>")
        names = sorted(self._common_readable_map.keys())
        for name in names:
            self.cmbFilterParameter.Items.Add(name)
        if self.cmbFilterParameter.Items.Count > 0:
            self.cmbFilterParameter.SelectedItem = "<Category>"

    def _refresh_target_parameter_combo(self):
        self.cmbTargetParameter.Items.Clear()
        names = sorted(self._common_writable_map.keys())
        for name in names:
            self.cmbTargetParameter.Items.Add(name)
        if self.cmbTargetParameter.Items.Count > 0:
            self.cmbTargetParameter.SelectedIndex = 0

    def _refresh_source_parameter_combo(self):
        self.cmbSourceParameter.Items.Clear()
        names = sorted(self._common_readable_map.keys())
        for name in names:
            self.cmbSourceParameter.Items.Add(name)
        if self.cmbSourceParameter.Items.Count > 0:
            self.cmbSourceParameter.SelectedIndex = 0

    def _collect_value_buckets(self, param_name):
        buckets = {}
        for element in self.current_pool_elements:
            if not _is_valid_element(element):
                continue
            if param_name == "<Category>":
                cat_name = "<No Category>"
                try:
                    if element.Category is not None and element.Category.Name:
                        cat_name = _safe_str(element.Category.Name)
                except Exception:
                    pass
                key = ("cat", cat_name)
                disp = cat_name
            else:
                param = _find_param_on_element(element, param_name)
                key, disp = _param_value_key_display(param)
            if key not in buckets:
                buckets[key] = {"display": disp, "count": 0}
            buckets[key]["count"] += 1

        return buckets

    def _load_value_list_for_selected_parameter(self):
        self.lstFilterValues.Items.Clear()
        self._last_filter_map = {}

        param_name = _selected_combo_text(self.cmbFilterParameter)
        if not param_name:
            self._set_filter_summary("Choose a filter parameter first.")
            return

        buckets = self._collect_value_buckets(param_name)
        keys = sorted(buckets.keys(), key=lambda k: (_safe_str(buckets[k].get("display", "")), _safe_str(k)))

        for key in keys:
            disp = _safe_str(buckets[key].get("display", "<empty>"))
            count = buckets[key].get("count", 0)
            label = "{0} ({1})".format(disp, count)
            self.lstFilterValues.Items.Add(label)
            self._last_filter_map[label] = key

        self._set_filter_summary("Loaded {0} unique values for parameter '{1}'.".format(len(keys), param_name))

    def _selected_value_keys(self):
        keys = set()
        try:
            for item in self.lstFilterValues.SelectedItems:
                label = _safe_str(item)
                if label in self._last_filter_map:
                    keys.add(self._last_filter_map[label])
        except Exception:
            pass
        return keys

    def _apply_filter(self, keep_checked):
        param_name = _selected_combo_text(self.cmbFilterParameter)
        if not param_name:
            forms.alert("Choose a filter parameter first.")
            return

        selected_keys = self._selected_value_keys()
        if not selected_keys:
            forms.alert("Select one or more values from the list first.")
            return

        out_elements = []
        for element in self.current_pool_elements:
            if not _is_valid_element(element):
                continue
            if param_name == "<Category>":
                cat_name = "<No Category>"
                try:
                    if element.Category is not None and element.Category.Name:
                        cat_name = _safe_str(element.Category.Name)
                except Exception:
                    pass
                key = ("cat", cat_name)
            else:
                param = _find_param_on_element(element, param_name)
                key, _ = _param_value_key_display(param)
            is_match = key in selected_keys
            if (keep_checked and is_match) or ((not keep_checked) and (not is_match)):
                out_elements.append(element)

        self.filtered_elements = out_elements
        mode_text = "Keep Checked" if keep_checked else "Unselect Checked"
        self._set_filter_summary(
            "Filter mode: {0} | Parameter: {1} | Result elements: {2}".format(mode_text, param_name, len(out_elements))
        )
        self._refresh_pool_summary()

    def _is_source_mode(self):
        return _selected_combo_text(self.cmbProcessMode) == "Copy Source Parameter"

    def _update_mode_ui(self):
        pnl_direct = getattr(self, "pnlDirectMode", None)
        pnl_source = getattr(self, "pnlSourceMode", None)
        if pnl_direct is None or pnl_source is None:
            return
        source_mode = self._is_source_mode()
        pnl_direct.Visibility = System.Windows.Visibility.Collapsed if source_mode else System.Windows.Visibility.Visible
        pnl_source.Visibility = System.Windows.Visibility.Visible if source_mode else System.Windows.Visibility.Collapsed

    def _confirm_long_run(self, count, source_mode):
        estimate = _estimate_seconds_for_apply(count, source_mode)
        text = (
            "This run will process {0} elements.\n"
            "Estimated run time: {1}\n\n"
            "Continue?"
        ).format(count, _format_estimate(estimate))
        return forms.alert(text, title="Confirm Processing", yes=True, no=True), estimate

    def _preflight_test(self, target_name, target_storage, source_name, transform_name, dict_text):
        sample = None
        for element in self.filtered_elements:
            if _is_valid_element(element):
                sample = element
                break

        if sample is None:
            return True

        tx = Transaction(doc, "View Filter Processor Preflight")
        try:
            tx.Start()

            target_param = _find_param_on_element(sample, target_name)
            if target_param is None or target_param.IsReadOnly:
                raise ValueError("Sample target parameter is missing or read-only.")

            if source_name:
                source_param = _find_param_on_element(sample, source_name)
                if source_param is None:
                    raise ValueError("Sample source parameter is missing.")
                raw_source = _read_parameter_value(source_param)
                transformed = _transform_value(raw_source, transform_name, dict_text)
                coerced = _coerce_value(transformed, target_storage) if target_storage != StorageType.String else _safe_str(transformed)
            else:
                coerced = _coerce_value(self.txtDirectValue.Text, target_storage)

            ok = _set_parameter_value(target_param, coerced, target_storage)
            if not ok:
                raise ValueError("Preflight write returned false.")

            tx.RollBack()
            return True
        except Exception as ex:
            try:
                tx.RollBack()
            except Exception:
                pass
            forms.alert("Preflight test failed:\n{0}".format(ex), title="Parameter Flow Processor")
            return False

    def _pump_ui(self):
        try:
            if WinFormsApplication is not None:
                WinFormsApplication.DoEvents()
        except Exception:
            pass

    def _set_running(self, running):
        self._is_busy = bool(running)
        if running:
            self._cancel_requested = False
            self._cancel_mode = None
            self._pending_focus_element_id = None

    def _request_cancel(self):
        if not self._is_busy:
            forms.alert("No process is currently running.")
            return
        choice = None
        try:
            choice = forms.CommandSwitchWindow.show(
                [
                    "Rollback all changes from this run",
                    "Keep already processed data",
                ],
                message="Choose cancel behavior for the running process:",
            )
        except Exception:
            choice = None

        if not choice:
            answer = forms.alert(
                "Cancel running process?\n\nYes = Rollback all changes from this run.\nNo = Keep already committed chunk changes and stop now.",
                title="Cancel Processing",
                yes=True,
                no=True,
            )
            choice = "Rollback all changes from this run" if answer else "Keep already processed data"

        self._cancel_mode = "rollback" if "Rollback" in _safe_str(choice) else "keep"
        self._cancel_requested = True
        self._set_run_status("Cancel requested ({0}). Waiting for safe stop point...".format(self._cancel_mode))
        self._monitor_set_status("Cancel requested: {0}. Finishing current safe step...".format(self._cancel_mode))

    def _apply_processing(self):
        if not self.filtered_elements:
            forms.alert("No filtered elements to process. Apply a filter first.")
            return

        target_name = _selected_combo_text(self.cmbTargetParameter)
        if not target_name:
            forms.alert("Choose a target parameter first.")
            return

        target_storage = self._common_writable_map.get(target_name)
        if target_storage is None:
            forms.alert("Selected target parameter is not writable for the current pool.")
            return

        source_mode = self._is_source_mode()
        source_name = _selected_combo_text(self.cmbSourceParameter) if source_mode else ""
        transform_name = _selected_combo_text(self.cmbTransform)
        dict_text = self.txtDictionary.Text if hasattr(self, "txtDictionary") else ""
        skip_existing = bool(self.chkSkipExisting.IsChecked) if hasattr(self, "chkSkipExisting") else False

        if source_mode and not source_name:
            forms.alert("Choose a source parameter for source processing mode.")
            return

        if len(self.filtered_elements) > MAX_APPLY_LIMIT:
            forms.alert(
                "Safety limit exceeded: {0:,} filtered elements. Limit is {1:,}. Narrow your filter and retry.".format(
                    len(self.filtered_elements), MAX_APPLY_LIMIT
                )
            )
            return

        proceed, estimate_seconds = self._confirm_long_run(len(self.filtered_elements), source_mode)
        if not proceed:
            return

        if not self._preflight_test(target_name, target_storage, source_name, transform_name, dict_text):
            return

        self._set_running(True)
        self._show_monitor_window()
        self.prgRun.Minimum = 0
        self.prgRun.Maximum = max(1, len(self.filtered_elements))
        self.prgRun.Value = 0

        written = 0
        skipped = 0
        failed = 0
        processed = 0
        sample_fails = []
        sample_writes = []
        cancelled = False
        written_ids_this_run = set()

        self._set_run_status("Running...")
        self._set_run_detail("Current item: preparing...")
        self._refresh_counts(0, 0, 0, 0)
        self._monitor_set_status("Run started. Preparing first chunk...")

        chunk_size = 220 if source_mode else 360
        tg = TransactionGroup(doc, "Parameter Flow Processor")
        tg.Start()

        try:
            total = len(self.filtered_elements)
            i = 0
            while i < total:
                if self._cancel_requested:
                    cancelled = True
                    break

                chunk = self.filtered_elements[i:i + chunk_size]
                tx = Transaction(doc, "View Filter Processor Chunk")
                tx.Start()
                chunk_cancelled = False

                try:
                    for element in chunk:
                        if self._cancel_requested:
                            cancelled = True
                            chunk_cancelled = True
                            break

                        self._pump_ui()
                        self._drain_pending_focus_request()

                        current_id = _safe_element_id(element)
                        self._set_run_detail("Current item: Element {0}".format(current_id))

                        if not _is_valid_element(element):
                            skipped += 1
                            processed += 1
                            self._monitor_add_row(element, "SKIP", target_name, "Invalid element")
                            self._refresh_counts(processed, written, skipped, failed)
                            continue

                        target_param = _find_param_on_element(element, target_name)
                        if target_param is None or target_param.IsReadOnly:
                            skipped += 1
                            processed += 1
                            self._monitor_add_row(element, "SKIP", target_name, "Target missing or read-only")
                            self._refresh_counts(processed, written, skipped, failed)
                            continue

                        if skip_existing and _is_effective_value_set(target_param, target_storage):
                            skipped += 1
                            processed += 1
                            self._monitor_add_row(element, "SKIP", target_name, "Skipped existing value")
                            self._refresh_counts(processed, written, skipped, failed)
                            continue

                        try:
                            if source_mode:
                                source_param = _find_param_on_element(element, source_name)
                                if source_param is None:
                                    skipped += 1
                                    processed += 1
                                    self._monitor_add_row(element, "SKIP", "{0} -> {1}".format(source_name, target_name), "Source parameter missing")
                                    self._refresh_counts(processed, written, skipped, failed)
                                    continue
                                raw_source = _read_parameter_value(source_param)
                                transformed = _transform_value(raw_source, transform_name, dict_text)
                                coerced = _safe_str(transformed) if target_storage == StorageType.String else _coerce_value(transformed, target_storage)
                            else:
                                coerced = _coerce_value(self.txtDirectValue.Text, target_storage)

                            ok = _set_parameter_value(target_param, coerced, target_storage)
                            if not ok:
                                raise ValueError("Set() returned false")

                            readback = _coerce_readback_value(target_param, target_storage)
                            expected = coerced.IntegerValue if target_storage == StorageType.ElementId else coerced
                            if not _values_match(readback, expected, target_storage):
                                raise ValueError("Verification mismatch. Expected {0}, got {1}.".format(expected, _display_parameter_value(target_param)))

                            written += 1
                            try:
                                written_ids_this_run.add(element.Id.IntegerValue)
                            except Exception:
                                pass
                            if len(sample_writes) < 8:
                                sample_writes.append("Element {0}: {1} -> {2}".format(_safe_element_id(element), target_name, _display_parameter_value(target_param)))
                            step_name = "{0} -> {1}".format(source_name, target_name) if source_mode else target_name
                            self._monitor_add_row(element, "SUCCESS", step_name, "Value applied and verified")
                        except Exception as row_ex:
                            failed += 1
                            if len(sample_fails) < 8:
                                sample_fails.append("Element {0}: {1}".format(_safe_element_id(element), row_ex))
                            step_name = "{0} -> {1}".format(source_name, target_name) if source_mode else target_name
                            self._monitor_add_row(element, "FAIL", step_name, _safe_str(row_ex))

                        processed += 1
                        self.prgRun.Value = processed
                        self._set_run_status("Running: {0}/{1}".format(processed, total))
                        self._monitor_set_status(
                            "Running {0}/{1} | Written: {2} | Skipped: {3} | Failed: {4}".format(
                                processed, total, written, skipped, failed
                            )
                        )
                        self._refresh_counts(processed, written, skipped, failed)

                    if chunk_cancelled:
                        tx.RollBack()
                    else:
                        tx.Commit()
                        self._drain_pending_focus_request()
                except Exception:
                    tx.RollBack()
                    raise

                if cancelled:
                    break
                i += chunk_size

            if cancelled:
                if self._cancel_mode == "rollback":
                    tg.RollBack()
                    self._set_run_status("Cancelled and rolled back. No changes kept.")
                    self._monitor_set_status("Cancelled and rolled back at {0}/{1}.".format(processed, total))
                    written_ids_this_run = set()
                else:
                    tg.Assimilate()
                    self._set_run_status("Cancelled and kept committed chunk changes.")
                    self._monitor_set_status("Cancelled and kept committed changes at {0}/{1}.".format(processed, total))
            else:
                tg.Assimilate()
                self._set_run_status("Completed successfully.")
                self._monitor_set_status("Completed successfully: processed {0} items.".format(processed))
        except Exception as ex:
            try:
                tg.RollBack()
            except Exception:
                pass
            self._set_run_status("Failed. Rolled back changes.")
            self._monitor_set_status("Failed and rolled back: {0}".format(ex))
            forms.alert("Processing failed and was rolled back.\n\n{0}".format(ex), title="Parameter Flow Processor")
            self._set_running(False)
            return

        self._set_running(False)
        self._set_run_detail("Current item: -")

        for eid in written_ids_this_run:
            self.written_element_ids.add(eid)

        elapsed_text = _format_estimate(estimate_seconds)
        summary = [
            "Run finished.",
            "Mode: {0}".format(_selected_combo_text(self.cmbProcessMode)),
            "Target parameter: {0}".format(target_name),
            "Estimated duration: {0}".format(elapsed_text),
            "Processed: {0}".format(processed),
            "Written: {0}".format(written),
            "Skipped: {0}".format(skipped),
            "Failed: {0}".format(failed),
        ]

        if sample_writes:
            summary.append("")
            summary.append("Sample written values:")
            summary.extend(["- {0}".format(x) for x in sample_writes])

        if sample_fails:
            summary.append("")
            summary.append("Sample failures:")
            summary.extend(["- {0}".format(x) for x in sample_fails])

        forms.alert("\n".join(summary), title="Parameter Flow Processor")

        self._refresh_pool_summary()

    def window_loaded(self, sender, e):
        try:
            try:
                if self.cmbProcessMode.SelectedIndex < 0 and self.cmbProcessMode.Items.Count > 0:
                    self.cmbProcessMode.SelectedIndex = 0
            except Exception:
                pass
            self._update_mode_ui()
            self._load_value_list_for_selected_parameter()
            self._ui_ready = True
        except Exception:
            pass

    def reload_active_view_click(self, sender, e):
        if self._is_busy:
            forms.alert("Cannot reload while processing.")
            return
        self._initialize_from_active_view()
        self._load_value_list_for_selected_parameter()

    def use_all_view_click(self, sender, e):
        if self._is_busy:
            forms.alert("Cannot change pool while processing.")
            return
        self.current_pool_elements = list(self.all_view_elements)
        self.filtered_elements = list(self.current_pool_elements)
        self._rebuild_parameter_maps()
        self._refresh_filter_parameter_combo()
        self._refresh_target_parameter_combo()
        self._refresh_source_parameter_combo()
        self._load_value_list_for_selected_parameter()
        self._set_filter_summary("Pool reset to all active-view elements.")
        self._refresh_pool_summary()

    def use_unprocessed_click(self, sender, e):
        if self._is_busy:
            forms.alert("Cannot change pool while processing.")
            return

        out = []
        for element in self.all_view_elements:
            if not _is_valid_element(element):
                continue
            try:
                if element.Id.IntegerValue in self.written_element_ids:
                    continue
            except Exception:
                pass
            out.append(element)

        self.current_pool_elements = out
        self.filtered_elements = list(out)
        self._rebuild_parameter_maps()
        self._refresh_filter_parameter_combo()
        self._refresh_target_parameter_combo()
        self._refresh_source_parameter_combo()
        self._load_value_list_for_selected_parameter()
        self._set_filter_summary("Pool set to unprocessed active-view elements.")
        self._refresh_pool_summary()

    def clear_filter_click(self, sender, e):
        if self._is_busy:
            forms.alert("Cannot clear filter while processing.")
            return
        self.filtered_elements = list(self.current_pool_elements)
        self._set_filter_summary("Filter cleared. Current filtered set equals current pool.")
        self._refresh_pool_summary()

    def filter_search_changed(self, sender, e):
        search = _safe_str(self.txtFilterSearch.Text).strip().lower()
        self.cmbFilterParameter.Items.Clear()
        for name in sorted(self._common_readable_map.keys()):
            if search and search not in _safe_str(name).lower():
                continue
            self.cmbFilterParameter.Items.Add(name)
        if self.cmbFilterParameter.Items.Count > 0:
            self.cmbFilterParameter.SelectedIndex = 0

    def filter_parameter_changed(self, sender, e):
        self._load_value_list_for_selected_parameter()

    def load_filter_values_click(self, sender, e):
        self._load_value_list_for_selected_parameter()

    def keep_checked_click(self, sender, e):
        if self._is_busy:
            forms.alert("Cannot apply filter while processing.")
            return
        self._apply_filter(True)

    def unselect_checked_click(self, sender, e):
        if self._is_busy:
            forms.alert("Cannot apply filter while processing.")
            return
        self._apply_filter(False)

    def process_mode_changed(self, sender, e):
        if not getattr(self, "_ui_ready", False):
            return
        self._update_mode_ui()

    def process_filtered_click(self, sender, e):
        if self._is_busy:
            forms.alert("A process is already running.")
            return
        self._apply_processing()

    def open_monitor_click(self, sender, e):
        monitor = self._show_monitor_window()
        if monitor is None:
            return
        if not self._is_busy:
            self._monitor_set_status("Monitor ready. Start a processing run to stream live item statuses.")

    def cancel_run_click(self, sender, e):
        self._request_cancel()

    def close_click(self, sender, e):
        if self._is_busy:
            forms.alert("Cancel or wait for the current run to finish before closing.")
            return
        try:
            if self._monitor_window is not None:
                self._monitor_window.Close()
        except Exception:
            pass
        self.Close()


try:
    window = ViewFilterParameterProcessorWindow("WPFWindow.xaml")
    window.ShowDialog()
except Exception as ex:
    forms.alert("Unable to launch Parameter Flow Processor.\n\n{0}".format(ex))
