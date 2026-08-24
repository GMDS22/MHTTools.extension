# coding: utf8
from __future__ import print_function

import os
import traceback

from pyrevit import forms, revit, script
from Autodesk.Revit import DB


doc = revit.doc
logger = script.get_logger()
output = script.get_output()

TOOL_TITLE = "NWB Family Shared Cleaner"
NWB_PREFIX = "NWB_"
CONVERTER_SCRIPT = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "MEP Manage.panel",
        "Family Tools.stack",
        "09 Family Convention Converter.pushbutton",
        "script.py",
    )
)


def _safe_str(value):
    try:
        if value is None:
            return ""
        return str(value)
    except Exception:
        return ""


def _alert(message, exitscript=False):
    forms.alert(message, title=TOOL_TITLE, exitscript=exitscript)


def _load_converter_module():
    if not os.path.isfile(CONVERTER_SCRIPT):
        raise RuntimeError("Family Convention Converter script was not found: {0}".format(CONVERTER_SCRIPT))

    namespace = {
        "__name__": "nwb_family_shared_cleaner_support",
        "__file__": CONVERTER_SCRIPT,
    }
    with open(CONVERTER_SCRIPT, "r") as stream:
        source = stream.read()
    exec(compile(source, CONVERTER_SCRIPT, "exec"), namespace)
    return namespace


def _read_guid_text(value):
    if value is None:
        return ""
    try:
        if callable(value):
            value = value()
    except Exception:
        pass
    text = _safe_str(value).strip()
    return text.lower()


def _guid_text_from_object(obj):
    if obj is None:
        return ""
    for attr_name in ("GUID", "GuidValue", "Guid"):
        try:
            value = getattr(obj, attr_name)
        except Exception:
            value = None
        text = _read_guid_text(value)
        if text:
            return text
    return ""


def _family_parameter_guid_text(family_param):
    text = _guid_text_from_object(family_param)
    if text:
        return text
    try:
        definition = family_param.Definition
    except Exception:
        definition = None
    return _guid_text_from_object(definition)


def _definition_guid_text(definition):
    return _guid_text_from_object(definition)


def _nwb_parameter_name(family_param):
    try:
        name = family_param.Definition.Name
    except Exception:
        name = ""
    return _safe_str(name).strip()


class FamilyChoice(object):
    def __init__(self, family, converter):
        self.family = family
        self.name = converter["safe_name"](family)

    @property
    def Name(self):
        return self.name

    def __str__(self):
        return self.name


class FamilyCleanPlan(object):
    def __init__(self, family, family_name):
        self.family = family
        self.family_name = family_name
        self.replace_ops = []
        self.kept = []
        self.skipped = []
        self.preview_error = None
        self.is_editable = True

    @property
    def total_changes(self):
        return len(self.replace_ops)


def _get_selected_families(converter):
    selected = {}
    try:
        selected_ids = revit.get_selection().element_ids
    except Exception:
        try:
            selected_ids = revit.uidoc.Selection.GetElementIds()
        except Exception:
            selected_ids = []

    for element_id in selected_ids:
        try:
            element = doc.GetElement(element_id)
        except Exception:
            continue

        family = None
        if isinstance(element, DB.Family):
            family = element
        elif isinstance(element, DB.FamilySymbol):
            family = element.Family
        elif isinstance(element, DB.FamilyInstance):
            family = element.Symbol.Family
        if family is None:
            continue

        try:
            selected[family.Id.IntegerValue] = family
        except Exception:
            continue

    return list(selected.values())


def _choose_families(converter):
    selected = _get_selected_families(converter)
    if selected:
        return selected

    candidates = []
    for family in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if converter["is_editable_family"](family):
            candidates.append(FamilyChoice(family, converter))

    if not candidates:
        return []

    picked = forms.SelectFromList.show(
        sorted(candidates, key=lambda item: item.name.lower()),
        title="Select families to clean NWB parameters",
        multiselect=True,
        button_name="Clean Selected Families",
    )
    if not picked:
        return []
    return [item.family for item in picked]


def _replace_reason(family_param, target_definition):
    if not bool(getattr(family_param, "IsShared", False)):
        return "Replace non-shared family parameter with approved shared NWB definition"

    current_guid = _family_parameter_guid_text(family_param)
    target_guid = _definition_guid_text(target_definition)
    if current_guid and target_guid and current_guid != target_guid:
        return "Replace wrong-GUID shared parameter with approved shared NWB definition"
    return ""


def _build_family_plan(converter, shared_definition_file, family):
    family_name = converter["safe_name"](family)
    plan = FamilyCleanPlan(family, family_name)
    if not converter["is_editable_family"](family):
        plan.is_editable = False
        plan.skipped.append("Family is not editable.")
        return plan

    family_doc = None
    try:
        family_doc = doc.EditFamily(family)
        found_nwb = False
        for family_param in converter["iter_family_parameters"](family_doc):
            param_name = _nwb_parameter_name(family_param)
            if not param_name or not param_name.upper().startswith(NWB_PREFIX):
                continue
            found_nwb = True

            if converter["is_builtin_family_parameter"](family_param):
                plan.skipped.append("{0}: built-in parameter".format(param_name))
                continue
            if converter["has_formula"](family_param):
                plan.skipped.append("{0}: formula-driven parameter".format(param_name))
                continue
            if converter["is_reporting"](family_param):
                plan.skipped.append("{0}: reporting parameter".format(param_name))
                continue

            target_definition = converter["find_shared_definition"](shared_definition_file, param_name)
            if target_definition is None:
                plan.skipped.append("{0}: no matching shared definition was found in the active shared parameter file".format(param_name))
                continue

            current_is_shared = bool(getattr(family_param, "IsShared", False))
            current_guid = _family_parameter_guid_text(family_param)
            target_guid = _definition_guid_text(target_definition)

            if current_is_shared:
                if not current_guid or not target_guid:
                    plan.skipped.append("{0}: shared parameter GUID could not be verified safely".format(param_name))
                    continue
                if current_guid == target_guid:
                    plan.kept.append(param_name)
                    continue

            reason = _replace_reason(family_param, target_definition)
            if not reason:
                plan.kept.append(param_name)
                continue

            plan.replace_ops.append(
                {
                    "old_name": param_name,
                    "new_name": param_name,
                    "definition_group": "Meinhardt",
                    "target_group": None,
                    "reason": reason,
                    "current_guid": current_guid,
                    "target_guid": target_guid,
                    "is_instance": bool(getattr(family_param, "IsInstance", False)),
                }
            )

        if not found_nwb:
            plan.skipped.append("No NWB_ family parameters were found.")
    except Exception as exc:
        plan.preview_error = _safe_str(exc)
        logger.error(traceback.format_exc())
    finally:
        try:
            if family_doc is not None:
                family_doc.Close(False)
        except Exception:
            pass

    return plan


def _print_preview(plans):
    output.close_others()
    output.print_md("# {0}".format(TOOL_TITLE))

    total_replace = 0
    total_kept = 0
    total_skipped = 0

    for plan in plans:
        output.print_md("## {0}".format(plan.family_name or "<Unnamed Family>"))
        if not plan.is_editable:
            output.print_md("- Not editable")
            continue
        if plan.preview_error:
            output.print_md("- Preview failed: {0}".format(plan.preview_error))
            continue

        total_replace += len(plan.replace_ops)
        total_kept += len(plan.kept)
        total_skipped += len(plan.skipped)

        if plan.replace_ops:
            output.print_md("- Replacement candidates:")
            for op in plan.replace_ops:
                guid_text = "{0} -> {1}".format(op.get("current_guid", "?") or "?", op.get("target_guid", "?") or "?")
                output.print_md("  - `{0}`: {1} (`{2}`)".format(op.get("old_name", ""), op.get("reason", ""), guid_text))
        else:
            output.print_md("- Replacement candidates: none")

        if plan.kept:
            output.print_md("- Already correct shared NWB parameters: {0}".format(", ".join(sorted(plan.kept))))

        if plan.skipped:
            output.print_md("- Skipped:")
            for item in plan.skipped[:12]:
                output.print_md("  - {0}".format(item))
            if len(plan.skipped) > 12:
                output.print_md("  - ... and {0} more".format(len(plan.skipped) - 12))

    output.print_md("---")
    output.print_md(
        "**Summary:** {0} families reviewed, {1} replacement candidate(s), {2} already-correct shared NWB parameter(s), {3} skipped item(s).".format(
            len(plans),
            total_replace,
            total_kept,
            total_skipped,
        )
    )


def _apply_plans(converter, shared_definition_file, plans):
    load_options = converter["FamilyLoadOptions"](False)
    summary = {
        "families_reloaded": 0,
        "params_replaced": 0,
        "values_restored": 0,
        "errors": [],
    }

    for plan in plans:
        if not plan.is_editable or plan.preview_error or not plan.replace_ops:
            continue

        family_doc = None
        try:
            family_doc = doc.EditFamily(plan.family)
            transaction = DB.Transaction(family_doc, TOOL_TITLE)
            transaction.Start()
            try:
                for op in plan.replace_ops:
                    sub = DB.SubTransaction(family_doc)
                    sub.Start()
                    try:
                        family_param = family_doc.FamilyManager.get_Parameter(op["old_name"])
                        if family_param is None:
                            raise Exception("Parameter not found")
                        restored = converter["replace_shared_parameter"](family_doc, family_param, op, shared_definition_file)
                        summary["params_replaced"] += 1
                        summary["values_restored"] += restored
                        sub.Commit()
                    except Exception as exc:
                        try:
                            sub.RollBack()
                        except Exception:
                            pass
                        summary["errors"].append(
                            "{0} | {1}: {2}".format(plan.family_name, op.get("old_name", ""), _safe_str(exc))
                        )
                transaction.Commit()
            except Exception:
                transaction.RollBack()
                raise

            family_doc.LoadFamily(doc, load_options)
            summary["families_reloaded"] += 1
        except Exception as exc:
            summary["errors"].append("{0}: {1}".format(plan.family_name, _safe_str(exc)))
            logger.error(traceback.format_exc())
        finally:
            try:
                if family_doc is not None:
                    family_doc.Close(False)
            except Exception:
                pass

    return summary


def main():
    if doc.IsFamilyDocument:
        _alert("Run this tool from a project document. It opens selected loadable families, replaces family-side NWB parameters, and reloads the families.", exitscript=True)
        return

    try:
        converter = _load_converter_module()
    except Exception as exc:
        _alert(_safe_str(exc), exitscript=True)
        return

    shared_definition_file = converter["get_shared_definition_file"]()
    if shared_definition_file is None:
        _alert("No shared parameter file is currently configured in Revit. Set the active shared parameter file first, then run the cleaner again.", exitscript=True)
        return

    families = _choose_families(converter)
    if not families:
        _alert("No families selected.")
        return

    plans = [_build_family_plan(converter, shared_definition_file, family) for family in families]
    _print_preview(plans)

    replace_candidates = sum(len(plan.replace_ops) for plan in plans)
    if replace_candidates <= 0:
        _alert("No safe NWB family-parameter replacements were found. Review the output preview for skipped items and already-correct shared parameters.")
        return

    proceed = forms.alert(
        "Preview written to the output window. Apply {0} NWB family-parameter replacement(s) across {1} selected family/families now?\n\nThis only cleans family parameters. Project parameter bindings are not removed by this tool.".format(
            replace_candidates,
            len(families),
        ),
        title=TOOL_TITLE,
        yes=True,
        no=True,
    )
    if not proceed:
        return

    summary = _apply_plans(converter, shared_definition_file, plans)
    output.print_md("## Apply Summary")
    output.print_md("- Families reloaded: {0}".format(summary["families_reloaded"]))
    output.print_md("- NWB family parameters replaced: {0}".format(summary["params_replaced"]))
    output.print_md("- Family type values restored: {0}".format(summary["values_restored"]))
    output.print_md("- Errors / warnings: {0}".format(len(summary["errors"])))
    if summary["errors"]:
        for item in summary["errors"]:
            output.print_md("  - {0}".format(item))
        _alert("Completed with {0} warning(s). Review the output window before cleaning project schedules.".format(len(summary["errors"])))
    else:
        _alert("NWB family shared-parameter cleanup completed successfully.")


if __name__ == "__main__":
    main()