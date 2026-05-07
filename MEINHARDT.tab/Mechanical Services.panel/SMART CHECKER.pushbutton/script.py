# -*- coding: utf-8 -*-
__title__ = "SMART CHECKER"
__doc__ = """Version = 2.0
Date: 2026-03-26
Author: GM
Description:
Smart duct QA checker with a WPF panel for scope, level filtering, editable
parameter rules, and selectable results with auto-navigation.
How-to:
1. Choose scope and optional level filter.
2. Edit or add the parameter rules to validate.
3. Run the check.
4. Click results to auto-show the failing duct in the current or 3D view.
"""

import clr
import re

clr.AddReference("System")
clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Windows.Forms")

import System
from System.Windows.Forms import Control
from System.Collections.Generic import List
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import DB, forms, revit, script


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None
logger = script.get_logger()
output = script.get_output()

NUMERIC_RE = re.compile(r"[-+]?\d*\.?\d+")

SHOW_PASSED_RESULTS = False
AUTO_SELECT_FAILING = True
AUTO_SHOW_SELECTED_RESULT = True


class ScopeChoice(object):
    def __init__(self, value, display):
        self.Value = value
        self.Display = display

    def __str__(self):
        return self.Display


class CheckRule(object):
    def __init__(self, name, aliases_text, mode, allowed_values_text="", min_value_text="", is_enabled=True):
        self.Name = name
        self.AliasesText = aliases_text
        self.Mode = mode
        self.AllowedValuesText = allowed_values_text
        self.MinValueText = min_value_text
        self.IsEnabled = is_enabled
        self.MatchedParameter = ""
        self.ResultSummary = "Not run"
        self.Display = ""
        self.refresh_display()

    def refresh_display(self):
        status = "On" if self.IsEnabled else "Off"
        match_text = self.MatchedParameter or "-"
        self.Display = "[{}] {} | {} | Match: {} | {}".format(
            status,
            self.Name,
            self.Mode,
            match_text,
            self.ResultSummary,
        )


class ResultRow(object):
    def __init__(self, duct, level_name, values, issues):
        self.Duct = duct
        self.LevelName = level_name
        self.Values = values
        self.Issues = issues
        self.Status = "PASS" if not issues else "FAIL"
        duct_name = _safe_element_name(duct)
        type_name = _safe_type_name(duct)
        issue_text = "OK" if not issues else "; ".join(issues)
        self.Display = "{} | Id {} | {} | {} | {}".format(
            self.Status,
            duct.Id.IntegerValue,
            level_name,
            duct_name or type_name,
            issue_text,
        )
        self.DetailText = self._build_detail_text(type_name)

    def _build_detail_text(self, type_name):
        lines = [
            "Status: {}".format(self.Status),
            "Element Id: {}".format(self.Duct.Id.IntegerValue),
            "Type: {}".format(type_name or "<unnamed>"),
            "Level: {}".format(self.LevelName),
        ]
        for key in sorted(self.Values.keys()):
            lines.append("{}: {}".format(key, self.Values[key]))
        if self.Issues:
            lines.append("Issues:")
            for issue in self.Issues:
                lines.append("- {}".format(issue))
        else:
            lines.append("Issues: None")
        return "\n".join(lines)


class DuctPickFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return bool(
                element
                and element.Category
                and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_DuctCurves)
            )
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return True


def _normalize_name(text):
    text = (text or "").strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _split_csv_text(text):
    values = []
    for token in re.split(r"[,\n;]+", text or ""):
        cleaned = token.strip()
        if cleaned:
            values.append(cleaned)
    return values


def _safe_param_name(param):
    try:
        definition = getattr(param, "Definition", None)
        if definition:
            return definition.Name or ""
    except Exception:
        pass
    return ""


def _safe_param_text(param):
    if param is None:
        return ""
    try:
        value = param.AsValueString()
        if value not in (None, ""):
            return value
    except Exception:
        pass
    try:
        value = param.AsString()
        if value not in (None, ""):
            return value
    except Exception:
        pass
    try:
        if param.StorageType == DB.StorageType.Integer:
            return str(param.AsInteger())
        if param.StorageType == DB.StorageType.Double:
            return str(param.AsDouble())
        if param.StorageType == DB.StorageType.ElementId:
            element_id = param.AsElementId()
            if element_id and element_id != DB.ElementId.InvalidElementId:
                return str(element_id.IntegerValue)
    except Exception:
        pass
    return ""


def _safe_param_number(param):
    if param is None:
        return None
    try:
        if param.StorageType == DB.StorageType.Double:
            return float(param.AsDouble())
        if param.StorageType == DB.StorageType.Integer:
            return float(param.AsInteger())
    except Exception:
        pass
    text = _safe_param_text(param)
    if not text:
        return None
    match = NUMERIC_RE.search(text.replace(",", ""))
    if not match:
        return None
    try:
        return float(match.group(0))
    except Exception:
        return None


def _safe_parameters(element):
    try:
        return list(element.Parameters)
    except Exception:
        return []


def _match_score(param_name, aliases):
    normalized = _normalize_name(param_name)
    if not normalized:
        return 0
    padded_name = " {} ".format(normalized)
    best = 0
    for alias in aliases:
        normalized_alias = _normalize_name(alias)
        if not normalized_alias:
            continue
        padded_alias = " {} ".format(normalized_alias)
        if normalized == normalized_alias:
            best = max(best, 100)
        elif normalized.startswith(normalized_alias + " ") or normalized.endswith(" " + normalized_alias):
            best = max(best, 85)
        elif padded_alias in padded_name:
            best = max(best, 70)
    return best


def _aliases_for_rule(rule):
    aliases = _split_csv_text(rule.AliasesText)
    if rule.Name and rule.Name not in aliases:
        aliases.append(rule.Name)
    return aliases


def _safe_element_name(element):
    try:
        name = getattr(element, "Name", None)
        if name:
            return name
    except Exception:
        pass
    return ""


def _safe_type_name(element):
    try:
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            return getattr(element_type, "Name", None) or ""
    except Exception:
        pass
    return _safe_element_name(element)


def _safe_duct_level(duct):
    try:
        level = getattr(duct, "ReferenceLevel", None)
        if level:
            return level
    except Exception:
        pass
    for built_in in (DB.BuiltInParameter.RBS_START_LEVEL_PARAM, DB.BuiltInParameter.FAMILY_LEVEL_PARAM):
        try:
            parameter = duct.get_Parameter(built_in)
            if parameter:
                level_id = parameter.AsElementId()
                if level_id and level_id != DB.ElementId.InvalidElementId:
                    level = doc.GetElement(level_id)
                    if level:
                        return level
        except Exception:
            continue
    return None


def _safe_duct_level_name(duct):
    level = _safe_duct_level(duct)
    if level:
        return getattr(level, "Name", None) or "No Level"
    return "No Level"


def _collect_selected_ducts():
    ducts = []
    for element_id in uidoc.Selection.GetElementIds():
        element = doc.GetElement(element_id)
        if element and element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_DuctCurves):
            ducts.append(element)
    return ducts


def _collect_view_ducts(view):
    if not view or getattr(view, "IsTemplate", False):
        return []
    try:
        return list(
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(DB.BuiltInCategory.OST_DuctCurves)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []


def _collect_model_ducts():
    return list(
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_DuctCurves)
        .WhereElementIsNotElementType()
    )


def _first_non_template_3d_view():
    try:
        views = DB.FilteredElementCollector(doc).OfClass(DB.View3D).ToElements()
    except Exception:
        views = []
    preferred = None
    fallback = None
    for view in views:
        try:
            if view.IsTemplate:
                continue
            if fallback is None:
                fallback = view
            if (view.Name or "") == "{3D}":
                preferred = view
                break
        except Exception:
            continue
    return preferred or fallback


class SmartCheckerWindow(forms.WPFWindow):
    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)

        self._panel_ratio = (0.95, 1.15, 1.2)
        self._is_custom_maximized = False
        self._restore_bounds = None

        self._selected_ducts = []
        self._view_ducts = []
        self._model_ducts = []
        self._rules = []
        self._results = []

        self._apply_screen_constraints()
        self._init_controls()
        self._load_context()
        self._update_window_buttons()

    def _init_controls(self):
        self.UI_scope.ItemsSource = []
        self.UI_scope.DisplayMemberPath = "Display"

        self.UI_level_filter.ItemsSource = ["All Levels"]
        self.UI_level_filter.SelectedIndex = 0

        self.UI_rule_mode.ItemsSource = ["Required", "Positive Number", "Minimum Number", "Allowed Values"]
        self.UI_rule_mode.SelectedIndex = 0

        self._reset_rules_to_defaults()
        self._clear_rule_editor()
        self._refresh_rule_list()
        self._refresh_results([])
        self._set_status("Choose scope and rules, then run the check.")

    def _update_scope_controls(self):
        scope_value = self._selected_scope_value()
        show_level = scope_value == "selected_level"
        show_pick = scope_value == "selected_ducts"

        level_visibility = System.Windows.Visibility.Visible if show_level else System.Windows.Visibility.Collapsed
        pick_visibility = System.Windows.Visibility.Visible if show_pick else System.Windows.Visibility.Collapsed

        try:
            self.UI_level_label.Visibility = level_visibility
            self.UI_level_filter.Visibility = level_visibility
        except Exception:
            pass

        try:
            self.UI_pick_button.Visibility = pick_visibility
        except Exception:
            pass

    def _build_default_rules(self):
        return [
            CheckRule("Pa", "pa, pressure drop, pressure loss, pressure drop per length, pressure loss per length, friction loss", "Positive Number"),
            CheckRule("System Name", "system name, system abbreviation", "Required"),
            CheckRule("Size", "size, calculated size", "Required"),
            CheckRule("Flow", "flow, air flow", "Required"),
            CheckRule("Velocity", "velocity, air velocity", "Required"),
        ]

    def _reset_rules_to_defaults(self):
        self._rules = self._build_default_rules()

    def _load_context(self, preferred_scope=None):
        self._selected_ducts = _collect_selected_ducts()
        self._view_ducts = _collect_view_ducts(doc.ActiveView)
        self._model_ducts = _collect_model_ducts()

        scope_choices = [
            ScopeChoice("current_view", "Current View ({})".format(len(self._view_ducts))),
            ScopeChoice("selected_level", "Selected Level ({})".format(len(self._model_ducts))),
            ScopeChoice("selected_ducts", "Selected Ducts ({})".format(len(self._selected_ducts))),
        ]
        current_value = preferred_scope or self._selected_scope_value()
        self.UI_scope.ItemsSource = scope_choices
        selected_index = 0
        for index, choice in enumerate(scope_choices):
            if choice.Value == current_value:
                selected_index = index
                break
        self.UI_scope.SelectedIndex = selected_index

        levels = sorted({_safe_duct_level_name(duct) for duct in self._model_ducts})
        current_level = self.UI_level_filter.SelectedItem if self.UI_level_filter else "All Levels"
        level_items = ["All Levels"] + levels
        self.UI_level_filter.ItemsSource = level_items
        try:
            if current_level in level_items:
                self.UI_level_filter.SelectedItem = current_level
            else:
                self.UI_level_filter.SelectedIndex = 0
        except Exception:
            self.UI_level_filter.SelectedIndex = 0

        self._update_scope_controls()
        self._update_counts()
        self._append_log("Context refreshed: {} selected, {} in active view, {} in model.".format(
            len(self._selected_ducts), len(self._view_ducts), len(self._model_ducts)
        ))

    def _selected_scope_value(self):
        choice = getattr(self.UI_scope, "SelectedItem", None)
        return getattr(choice, "Value", None) or "current_view"

    def _smart_scope_preview(self):
        scope_value = self._selected_scope_value()
        if scope_value == "selected_level":
            selected_level = self.UI_level_filter.SelectedItem if self.UI_level_filter else "All Levels"
            return "selected level -> {}".format(selected_level or "All Levels")
        if scope_value == "selected_ducts":
            return "selected ducts -> {}".format(len(self._selected_ducts))
        return "current view -> {}".format(len(self._view_ducts))

    def _scope_ducts(self):
        scope_value = self._selected_scope_value()
        if scope_value == "selected_ducts":
            return list(self._selected_ducts), "Selected Ducts"
        if scope_value == "selected_level":
            return list(self._model_ducts), "Selected Level"
        return list(self._view_ducts), "Current View"

    def _filtered_scope_ducts(self):
        ducts, scope_label = self._scope_ducts()
        selected_level = self.UI_level_filter.SelectedItem if self.UI_level_filter else "All Levels"
        if self._selected_scope_value() == "selected_level" and selected_level and selected_level != "All Levels":
            ducts = [duct for duct in ducts if _safe_duct_level_name(duct) == selected_level]
            scope_label = "{} | Level: {}".format(scope_label, selected_level)
        return ducts, scope_label

    def _refresh_rule_list(self):
        for rule in self._rules:
            rule.refresh_display()
        self.UI_rule_list.ItemsSource = None
        self.UI_rule_list.ItemsSource = self._rules
        self.UI_rule_list.Items.Refresh()
        self._update_counts()

    def _clear_rule_editor(self):
        self.UI_rule_name.Text = ""
        self.UI_rule_aliases.Text = ""
        self.UI_rule_mode.SelectedIndex = 0
        self.UI_rule_allowed_values.Text = ""
        self.UI_rule_min_value.Text = ""
        self.UI_rule_enabled.IsChecked = True
        self.UI_rule_list.SelectedItem = None
        self._update_rule_editor_state()

    def _load_rule_into_editor(self, rule):
        if not rule:
            return
        self.UI_rule_name.Text = rule.Name
        self.UI_rule_aliases.Text = rule.AliasesText
        self.UI_rule_mode.SelectedItem = rule.Mode
        self.UI_rule_allowed_values.Text = rule.AllowedValuesText
        self.UI_rule_min_value.Text = rule.MinValueText
        self.UI_rule_enabled.IsChecked = bool(rule.IsEnabled)
        self._update_rule_editor_state()

    def _update_rule_editor_state(self):
        mode = self.UI_rule_mode.SelectedItem or "Required"
        self.UI_rule_allowed_values.IsEnabled = mode == "Allowed Values"
        self.UI_rule_min_value.IsEnabled = mode == "Minimum Number"

    def _collect_rule_from_editor(self):
        name = (self.UI_rule_name.Text or "").strip()
        aliases = (self.UI_rule_aliases.Text or "").strip()
        mode = self.UI_rule_mode.SelectedItem or "Required"
        allowed_values = (self.UI_rule_allowed_values.Text or "").strip()
        min_value = (self.UI_rule_min_value.Text or "").strip()
        is_enabled = bool(self.UI_rule_enabled.IsChecked)
        if not name:
            forms.alert("Enter a rule name first.", title=__title__)
            return None
        if mode == "Allowed Values" and not allowed_values:
            forms.alert("Enter at least one allowed value for this rule.", title=__title__)
            return None
        if mode == "Minimum Number":
            try:
                float(min_value)
            except Exception:
                forms.alert("Enter a valid minimum number.", title=__title__)
                return None
        return CheckRule(name, aliases, mode, allowed_values, min_value, is_enabled)

    def _append_log(self, message):
        if not message:
            return
        try:
            current_text = self.UI_log.Text or ""
            self.UI_log.Text = current_text + ("\n" if current_text else "") + message
            self.UI_log.ScrollToEnd()
        except Exception:
            pass

    def _set_status(self, message):
        try:
            self.UI_status_bar.Text = message
        except Exception:
            pass

    def _update_counts(self):
        try:
            scope_ducts, _ = self._filtered_scope_ducts()
            enabled_rules = len([rule for rule in self._rules if rule.IsEnabled])
            result_count = len(self._results)
            failing_count = len([row for row in self._results if row.Issues])
            self.UI_header_mode.Text = "Scope ducts: {} | Rules: {} | Results: {}".format(len(scope_ducts), enabled_rules, result_count)
            self.UI_rule_count.Text = "({} rules, {} enabled)".format(len(self._rules), enabled_rules)
            self.UI_result_count.Text = "({} results, {} failing)".format(result_count, failing_count)
            self.UI_info_summary.Text = "Scope preview: {}".format(self._smart_scope_preview())
        except Exception:
            pass

    def _find_best_parameter_name(self, ducts, rule):
        aliases = _aliases_for_rule(rule)
        scores = {}
        for duct in ducts:
            for param in _safe_parameters(duct):
                param_name = _safe_param_name(param)
                score = _match_score(param_name, aliases)
                if score > 0:
                    scores[param_name] = scores.get(param_name, 0) + score
        return max(scores.items(), key=lambda item: item[1])[0] if scores else ""

    def _find_param_for_rule(self, duct, rule, preferred_name):
        params = _safe_parameters(duct)
        if preferred_name:
            for param in params:
                if _safe_param_name(param) == preferred_name:
                    return param
        aliases = _aliases_for_rule(rule)
        best_param = None
        best_score = 0
        for param in params:
            score = _match_score(_safe_param_name(param), aliases)
            if score > best_score:
                best_param = param
                best_score = score
        return best_param

    def _evaluate_rule(self, param, rule):
        if param is None:
            return False, "Missing parameter", "<missing>"
        text_value = _safe_param_text(param)
        if rule.Mode == "Required":
            return (True, "OK", text_value) if text_value.strip() else (False, "Blank value", text_value or "<blank>")
        if rule.Mode == "Positive Number":
            number_value = _safe_param_number(param)
            if number_value is None:
                return False, "Unreadable number", text_value or "<blank>"
            return (True, "OK", text_value or str(number_value)) if number_value > 0 else (False, "Non-positive number", text_value or str(number_value))
        if rule.Mode == "Minimum Number":
            number_value = _safe_param_number(param)
            if number_value is None:
                return False, "Unreadable number", text_value or "<blank>"
            threshold = float(rule.MinValueText)
            return (True, "OK", text_value or str(number_value)) if number_value >= threshold else (False, "Below minimum {}".format(threshold), text_value or str(number_value))
        if rule.Mode == "Allowed Values":
            actual = _normalize_name(text_value)
            allowed = [_normalize_name(value) for value in _split_csv_text(rule.AllowedValuesText)]
            return (True, "OK", text_value) if actual and actual in allowed else (False, "Value not in allowed list", text_value or "<blank>")
        return True, "OK", text_value

    def _run_check(self):
        ducts, scope_label = self._filtered_scope_ducts()
        if not ducts:
            forms.alert("No ducts are available in the chosen scope and level filter.", title=__title__)
            return
        enabled_rules = [rule for rule in self._rules if rule.IsEnabled]
        if not enabled_rules:
            forms.alert("Enable at least one rule before running the check.", title=__title__)
            return
        matched_names = {}
        for rule in enabled_rules:
            matched_names[rule.Name] = self._find_best_parameter_name(ducts, rule)
            rule.MatchedParameter = matched_names[rule.Name] or "Not detected"
            rule.ResultSummary = "Running"

        results = []
        failing_elements = []
        passed_count = 0
        for duct in sorted(ducts, key=lambda item: item.Id.IntegerValue):
            values = {"Scope": scope_label, "Type": _safe_type_name(duct) or "<unnamed>"}
            issues = []
            for rule in enabled_rules:
                param = self._find_param_for_rule(duct, rule, matched_names.get(rule.Name, ""))
                passed, message, display_value = self._evaluate_rule(param, rule)
                values[rule.Name] = display_value
                if not passed:
                    issues.append("{}: {}".format(rule.Name, message))
            row = ResultRow(duct, _safe_duct_level_name(duct), values, issues)
            if row.Issues:
                failing_elements.append(duct)
            else:
                passed_count += 1
            if row.Issues or SHOW_PASSED_RESULTS:
                results.append(row)

        failing_count = len(failing_elements)
        for rule in enabled_rules:
            failures_for_rule = 0
            for row in results:
                for issue in row.Issues:
                    if issue.startswith(rule.Name + ":"):
                        failures_for_rule += 1
            rule.ResultSummary = "{} fail".format(failures_for_rule)

        self._refresh_rule_list()
        self._refresh_results(results)
        if failing_elements and AUTO_SELECT_FAILING:
            self._select_elements(failing_elements)
        self._append_log("Run complete on {} ducts. {} failing, {} passing. Scope: {}.".format(len(ducts), failing_count, passed_count, scope_label))
        self._set_status("Run complete: {} failing ducts.".format(failing_count))
        self._update_counts()
        self._write_output_report(scope_label, ducts, enabled_rules, results, failing_count, passed_count)

    def _write_output_report(self, scope_label, ducts, rules, results, failing_count, passed_count):
        output.print_md("# {}".format(__title__))
        output.print_md("**Scope:** {}".format(scope_label))
        output.print_md("**Checked ducts:** {}".format(len(ducts)))
        output.print_md("**Failing:** {} | **Passing:** {}".format(failing_count, passed_count))
        output.print_md("## Rule mapping")
        output.print_table([[rule.Name, rule.Mode, rule.MatchedParameter or "Not detected"] for rule in rules], columns=["Rule", "Mode", "Matched Parameter"])
        if results:
            output.print_md("## Results")
            output.print_table([[output.linkify(row.Duct.Id), row.Status, row.LevelName, row.Values.get("Type", ""), "; ".join(row.Issues) if row.Issues else "OK"] for row in results], columns=["Duct", "Status", "Level", "Type", "Issues"])

    def _refresh_results(self, results):
        self._results = results
        self.UI_results_list.ItemsSource = None
        self.UI_results_list.ItemsSource = self._results
        self.UI_results_list.Items.Refresh()
        self.UI_result_details.Text = "Select a result to inspect details."
        self.UI_feedback.Text = "No results yet." if not results else "Select a result to review and auto-show it."
        self._update_counts()

    def _select_elements(self, elements):
        net_ids = List[DB.ElementId]()
        for element in elements:
            net_ids.Add(element.Id)
        uidoc.Selection.SetElementIds(net_ids)

    def _show_result(self, row):
        if not row:
            return
        self._select_elements([row.Duct])
        if self._show_in_view(row.Duct, uidoc.ActiveView):
            return
        view3d = _first_non_template_3d_view()
        if view3d and self._show_in_view(row.Duct, view3d):
            return
        self._append_log("Auto navigation could not show element {}.".format(row.Duct.Id.IntegerValue))

    def _show_in_view(self, element, target_view):
        if target_view is None:
            return False
        try:
            if uidoc.ActiveView.Id != target_view.Id:
                uidoc.ActiveView = target_view
            uidoc.ShowElements(element)
            return True
        except Exception:
            return False

    def _get_current_screen_work_area(self):
        try:
            cursor = Control.MousePosition
            screen = System.Windows.Forms.Screen.FromPoint(cursor)
            if screen is not None:
                return screen.WorkingArea
        except Exception:
            pass
        try:
            return System.Windows.SystemParameters.WorkArea
        except Exception:
            return None

    def _apply_screen_constraints(self):
        work_area = self._get_current_screen_work_area()
        if work_area is None:
            return
        try:
            self.MaxHeight = work_area.Height
            self.MaxWidth = work_area.Width
        except Exception:
            pass

    def _store_restore_bounds(self):
        if self._is_custom_maximized:
            return
        try:
            self._restore_bounds = {"Left": self.Left, "Top": self.Top, "Width": self.Width, "Height": self.Height}
        except Exception:
            self._restore_bounds = None

    def _restore_window_bounds(self):
        if not self._restore_bounds:
            return
        try:
            self.Left = self._restore_bounds["Left"]
            self.Top = self._restore_bounds["Top"]
            self.Width = self._restore_bounds["Width"]
            self.Height = self._restore_bounds["Height"]
        except Exception:
            pass

    def _maximize_to_current_monitor(self):
        self._store_restore_bounds()
        work_area = self._get_current_screen_work_area()
        if work_area is None:
            return
        try:
            self.WindowState = System.Windows.WindowState.Normal
            self.Left = work_area.Left
            self.Top = work_area.Top
            self.Width = work_area.Width
            self.Height = work_area.Height
            self._is_custom_maximized = True
        except Exception:
            pass

    def _restore_from_custom_maximize(self):
        self._is_custom_maximized = False
        self._restore_window_bounds()

    def _capture_panel_ratio(self):
        try:
            widths = [float(self.UI_col_panel_1.ActualWidth), float(self.UI_col_panel_2.ActualWidth), float(self.UI_col_panel_3.ActualWidth)]
            total = sum(width for width in widths if width > 0)
            if total > 0:
                self._panel_ratio = tuple(width / total for width in widths)
        except Exception:
            pass

    def _apply_panel_ratio(self):
        try:
            ratio1, ratio2, ratio3 = self._panel_ratio
            total = ratio1 + ratio2 + ratio3
            if total <= 0:
                return
            self.UI_col_panel_1.Width = System.Windows.GridLength(ratio1 / total, System.Windows.GridUnitType.Star)
            self.UI_col_panel_2.Width = System.Windows.GridLength(ratio2 / total, System.Windows.GridUnitType.Star)
            self.UI_col_panel_3.Width = System.Windows.GridLength(ratio3 / total, System.Windows.GridUnitType.Star)
        except Exception:
            pass

    def _update_window_buttons(self):
        try:
            self.UI_btn_maximize.Content = "o" if self._is_custom_maximized else "[]"
        except Exception:
            pass

    def _walk_visual_tree(self, root):
        if root is None:
            return
        yield root
        try:
            child_count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(root)
        except Exception:
            child_count = 0
        for index in range(child_count):
            try:
                child = System.Windows.Media.VisualTreeHelper.GetChild(root, index)
            except Exception:
                continue
            for descendant in self._walk_visual_tree(child):
                yield descendant

    def _apply_combo_theme(self):
        combo_boxes = [self.UI_scope, self.UI_level_filter, self.UI_rule_mode]
        for combo in combo_boxes:
            if combo is None:
                continue
            try:
                combo.ApplyTemplate()
                combo.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(238, 242, 214))
                combo.Foreground = System.Windows.Media.Brushes.Black
                combo.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 94, 168))
            except Exception:
                pass
            try:
                editable = combo.Template.FindName("PART_EditableTextBox", combo)
                if editable is not None:
                    editable.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(238, 242, 214))
                    editable.Foreground = System.Windows.Media.Brushes.Black
                    editable.CaretBrush = System.Windows.Media.Brushes.Black
                    editable.BorderThickness = System.Windows.Thickness(0)
                    editable.IsReadOnly = True
            except Exception:
                pass
            for child in self._walk_visual_tree(combo):
                try:
                    if isinstance(child, System.Windows.Controls.TextBox):
                        child.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(238, 242, 214))
                        child.Foreground = System.Windows.Media.Brushes.Black
                        child.CaretBrush = System.Windows.Media.Brushes.Black
                        child.BorderThickness = System.Windows.Thickness(0)
                        child.IsReadOnly = True
                except Exception:
                    continue

    def _is_source_from_button(self, source):
        current = source
        while current is not None:
            try:
                if isinstance(current, System.Windows.Controls.Button):
                    return True
            except Exception:
                pass
            try:
                current = System.Windows.Media.VisualTreeHelper.GetParent(current)
            except Exception:
                current = None
        return False

    def header_drag(self, sender, event_args):
        try:
            if self._is_source_from_button(getattr(event_args, "OriginalSource", None)):
                return
        except Exception:
            pass
        try:
            click_count = getattr(event_args, "ClickCount", 1)
        except Exception:
            click_count = 1
        if click_count >= 2:
            self.button_maximize(sender, event_args)
            return
        try:
            self.DragMove()
        except Exception:
            pass

    def window_loaded(self, sender, event_args):
        self._apply_panel_ratio()
        self._apply_combo_theme()
        self._update_rule_editor_state()

    def button_minimize(self, sender, event_args):
        try:
            self.WindowState = System.Windows.WindowState.Minimized
        except Exception:
            pass

    def button_maximize(self, sender, event_args):
        self._capture_panel_ratio()
        if self._is_custom_maximized or self.WindowState == System.Windows.WindowState.Maximized:
            self.WindowState = System.Windows.WindowState.Normal
            self._restore_from_custom_maximize()
        else:
            self._maximize_to_current_monitor()
        try:
            self.UpdateLayout()
        except Exception:
            pass
        self._apply_panel_ratio()
        self._update_window_buttons()

    def button_close(self, sender, event_args):
        try:
            self.Close()
        except Exception:
            pass

    def panel_splitter_drag_completed(self, sender, event_args):
        self._capture_panel_ratio()

    def scope_changed(self, sender, event_args):
        self._update_scope_controls()
        self._update_counts()

    def level_changed(self, sender, event_args):
        self._update_counts()

    def rule_mode_changed(self, sender, event_args):
        self._update_rule_editor_state()

    def rule_enabled_click(self, sender, event_args):
        rule = getattr(sender, "Tag", None)
        if rule:
            rule.IsEnabled = bool(getattr(sender, "IsChecked", False))
            rule.refresh_display()
            self.UI_rule_list.Items.Refresh()
            self._update_counts()

    def rule_selection_changed(self, sender, event_args):
        rule = self.UI_rule_list.SelectedItem
        if rule:
            self._load_rule_into_editor(rule)

    def result_selection_changed(self, sender, event_args):
        row = self.UI_results_list.SelectedItem
        if row:
            self.UI_result_details.Text = row.DetailText
            self.UI_feedback.Text = row.Display
            if AUTO_SHOW_SELECTED_RESULT:
                self._show_result(row)

    def button_new_rule(self, sender, event_args):
        self._clear_rule_editor()
        self._set_status("Rule editor cleared.")

    def button_save_rule(self, sender, event_args):
        rule = self._collect_rule_from_editor()
        if rule is None:
            return
        existing = self.UI_rule_list.SelectedItem
        if existing and existing in self._rules:
            self._rules[self._rules.index(existing)] = rule
            self._append_log("Updated rule '{}'".format(rule.Name))
        else:
            self._rules.append(rule)
            self._append_log("Added rule '{}'".format(rule.Name))
        self._refresh_rule_list()
        self._clear_rule_editor()

    def button_remove_rule(self, sender, event_args):
        rule = self.UI_rule_list.SelectedItem
        if not rule:
            forms.alert("Select a rule to remove.", title=__title__)
            return
        self._rules.remove(rule)
        self._refresh_rule_list()
        self._clear_rule_editor()
        self._append_log("Removed rule '{}'".format(rule.Name))

    def button_reset_rules(self, sender, event_args):
        self._reset_rules_to_defaults()
        self._refresh_rule_list()
        self._clear_rule_editor()
        self._append_log("Rules reset to default duct checks.")

    def button_refresh_context(self, sender, event_args):
        self._load_context()
        self._set_status("Context refreshed.")

    def button_pick_from_model(self, sender, event_args):
        picked_ducts = []
        try:
            self.Hide()
        except Exception:
            pass

        try:
            references = uidoc.Selection.PickObjects(
                ObjectType.Element,
                DuctPickFilter(),
                "Select ducts for SMART DUCT CHECKER",
            )
            if references:
                element_ids = List[DB.ElementId]()
                for reference in references:
                    element = doc.GetElement(reference.ElementId)
                    if element and element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_DuctCurves):
                        picked_ducts.append(element)
                        element_ids.Add(element.Id)
                if picked_ducts:
                    uidoc.Selection.SetElementIds(element_ids)
        except OperationCanceledException:
            self._append_log("Model pick cancelled.")
        except Exception as exc:
            self._append_log("Model pick failed: {}".format(exc))
        finally:
            try:
                self.Show()
                self.Activate()
            except Exception:
                pass

        if picked_ducts:
            self._load_context(preferred_scope="selected_ducts")
            self._set_status("Picked {} ducts from model.".format(len(picked_ducts)))
            self._append_log("Picked {} ducts from model and switched scope to Selected Ducts.".format(len(picked_ducts)))

    def button_run_check(self, sender, event_args):
        self._run_check()

    def button_clear_results(self, sender, event_args):
        self._refresh_results([])
        self.UI_log.Text = "Ready. Choose scope and rules, then run the check."
        self._set_status("Results cleared.")

    def button_show_selected(self, sender, event_args):
        row = self.UI_results_list.SelectedItem
        if not row:
            forms.alert("Select a result first.", title=__title__)
            return
        self._show_result(row)

    def button_reset_layout(self, sender, event_args):
        self._panel_ratio = (0.95, 1.15, 1.2)
        self._apply_panel_ratio()
        self._set_status("Panel layout reset.")


if __name__ == "__main__":
    if not doc:
        forms.alert("Open a Revit model first.", title=__title__)
        script.exit()
    window = SmartCheckerWindow("SmartChecker.xaml")
    window.ShowDialog()