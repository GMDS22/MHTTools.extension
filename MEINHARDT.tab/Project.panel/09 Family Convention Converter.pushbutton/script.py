# -*- coding: utf-8 -*-
"""
MHT Family Convention Converter

Safely converts selected loadable families to the company naming convention by:
- renaming the family in the project document
- renaming eligible non-shared family parameters in the family document
- optionally replacing shared family parameters using explicit mappings
  from converter_rules.json with best-effort type-value restoration

This is intentionally conservative. Built-in, formula-driven, and unsupported
parameters are skipped and reported instead of being modified.
"""
from __future__ import print_function

import os
import re
import json
import traceback

from pyrevit import revit, script, forms, HOST_APP
from Autodesk.Revit import DB

logger = script.get_logger()
output = script.get_output()

SCRIPT_DIR = os.path.dirname(__file__)
RULES_FILE = os.path.join(SCRIPT_DIR, 'converter_rules.json')
TOOL_NAME = 'MHT Family Convention Converter'


def _alert(msg, exitscript=False):
    forms.alert(msg, title=TOOL_NAME, exitscript=exitscript)


class FamilyLoadOptions(DB.IFamilyLoadOptions):
    def __init__(self, overwrite_parameter_values):
        self.overwrite_parameter_values = bool(overwrite_parameter_values)

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = self.overwrite_parameter_values
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = self.overwrite_parameter_values
        return True


class FamilyChoice(object):
    def __init__(self, family):
        self.family = family
        self.name = safe_name(family)

    @property
    def Name(self):
        return self.name

    def __str__(self):
        return self.name


class ChangePlan(object):
    def __init__(self, family):
        self.family = family
        self.family_name = safe_name(family)
        self.target_family_name = self.family_name
        self.rename_ops = []
        self.shared_ops = []
        self.skipped = []
        self.errors = []
        self.is_editable = True
        self.preview_error = None

    @property
    def total_changes(self):
        return len(self.rename_ops) + len(self.shared_ops) + (1 if self.family_name != self.target_family_name else 0)



def safe_name(element):
    try:
        return element.Name or ''
    except Exception:
        try:
            return DB.Element.Name.__get__(element) or ''
        except Exception:
            return ''



def load_rules():
    try:
        with open(RULES_FILE, 'r') as fp:
            return json.load(fp)
    except Exception as exc:
        raise Exception('Failed to load converter rules from {}: {}'.format(RULES_FILE, exc))



def make_unique_name(base_name, existing_names, current_name=None):
    existing = set()
    for item in existing_names or []:
        try:
            existing.add((item or '').strip().lower())
        except Exception:
            pass

    base = (base_name or '').strip()
    if not base:
        base = 'MHT_NAME'

    current_key = (current_name or '').strip().lower()
    if base.lower() == current_key or base.lower() not in existing:
        return base

    idx = 1
    while True:
        candidate = '{}_{:02d}'.format(base, idx)
        if candidate.lower() == current_key or candidate.lower() not in existing:
            return candidate
        idx += 1



def normalize_name(name, cfg):
    original = name or ''
    result = original.strip()
    if not result:
        return result

    exact_map = cfg.get('exact_map', {}) or {}
    if original in exact_map:
        result = exact_map[original]
    elif result in exact_map:
        result = exact_map[result]

    strip_prefixes = cfg.get('strip_prefixes', []) or []
    changed = True
    while changed:
        changed = False
        lowered = result.lower()
        for prefix in strip_prefixes:
            if not prefix:
                continue
            if lowered.startswith(prefix.lower()):
                result = result[len(prefix):]
                changed = True
                break

    if cfg.get('replace_leading_abbreviation_with_prefix'):
        pattern = cfg.get('leading_abbreviation_pattern') or r'^(?:[A-Za-z]{2,6}_)+ '
        pattern = pattern.rstrip()
        try:
            result = re.sub(pattern, '', result)
        except Exception:
            pass

    for rule in cfg.get('regex_replacements', []) or []:
        pattern = rule.get('pattern')
        replace = rule.get('replace', '')
        if not pattern:
            continue
        try:
            result = re.sub(pattern, replace, result)
        except Exception:
            continue

    result = result.strip('_ ').replace(' ', '_')
    result = re.sub(r'_{2,}', '_', result)

    prefix = cfg.get('prefix', '') or ''
    suffix = cfg.get('suffix', '') or ''
    if prefix and not result.lower().startswith(prefix.lower()):
        result = prefix + result
    if suffix and not result.lower().endswith(suffix.lower()):
        result = result + suffix

    if cfg.get('uppercase'):
        result = result.upper()
    elif cfg.get('lowercase'):
        result = result.lower()

    result = result.strip('_ ')
    result = re.sub(r'_{2,}', '_', result)
    return result or original



def is_editable_family(family):
    try:
        if family.IsInPlace:
            return False
    except Exception:
        pass
    try:
        return bool(family.IsEditable)
    except Exception:
        return True



def iter_family_parameters(family_doc):
    params = []
    try:
        iterator = family_doc.FamilyManager.Parameters.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            params.append(iterator.Current)
    except Exception:
        try:
            for item in family_doc.FamilyManager.Parameters:
                params.append(item)
        except Exception:
            pass
    return params



def iter_family_types(family_doc):
    types = []
    try:
        iterator = family_doc.FamilyManager.Types.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            types.append(iterator.Current)
    except Exception:
        try:
            for item in family_doc.FamilyManager.Types:
                types.append(item)
        except Exception:
            pass
    return types



def get_selected_families(doc):
    families = {}
    try:
        selected_ids = revit.get_selection().element_ids
    except Exception:
        try:
            selected_ids = revit.uidoc.Selection.GetElementIds()
        except Exception:
            selected_ids = []

    for elid in selected_ids:
        try:
            element = doc.GetElement(elid)
        except Exception:
            continue
        family = None
        if isinstance(element, DB.Family):
            family = element
        elif isinstance(element, DB.FamilySymbol):
            family = element.Family
        elif isinstance(element, DB.FamilyInstance):
            family = element.Symbol.Family
        if family is not None:
            try:
                families[family.Id.IntegerValue] = family
            except Exception:
                pass
    return families.values()



def choose_families(doc):
    selected = list(get_selected_families(doc))
    if selected:
        return selected

    candidates = []
    for fam in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if is_editable_family(fam):
            candidates.append(FamilyChoice(fam))

    if not candidates:
        return []

    picked = forms.SelectFromList.show(
        sorted(candidates, key=lambda x: x.name.lower()),
        title='Select families to convert',
        multiselect=True,
        button_name='Convert Selected Families'
    )
    if not picked:
        return []
    return [item.family for item in picked]



def is_builtin_family_parameter(family_param):
    try:
        return family_param.Definition.BuiltInParameter != DB.BuiltInParameter.INVALID
    except Exception:
        return False



def has_formula(family_param):
    try:
        return bool(family_param.IsDeterminedByFormula)
    except Exception:
        return False



def is_reporting(family_param):
    try:
        return bool(family_param.IsReporting)
    except Exception:
        return False



def rename_family_parameter(family_doc, family_param, new_name):
    manager = family_doc.FamilyManager
    try:
        manager.RenameParameter(family_param, new_name)
        return
    except Exception:
        pass
    manager.Rename(family_param, new_name)



def find_shared_mapping(shared_rules, old_name):
    for mapping in shared_rules:
        if (mapping.get('old_name') or '').strip().lower() == (old_name or '').strip().lower():
            return mapping
    return None



def get_existing_family_names(doc):
    names = []
    try:
        for fam in DB.FilteredElementCollector(doc).OfClass(DB.Family):
            names.append(safe_name(fam))
    except Exception:
        pass
    return names



def get_existing_parameter_names(family_doc):
    return [p.Definition.Name for p in iter_family_parameters(family_doc)]



def resolve_target_group(group_name, fallback_group):
    if not group_name:
        return fallback_group
    try:
        return getattr(DB.BuiltInParameterGroup, group_name)
    except Exception:
        return fallback_group



def get_shared_definition_file():
    try:
        return HOST_APP.app.OpenSharedParameterFile()
    except Exception:
        return None



def find_shared_definition(definition_file, target_name):
    if not definition_file:
        return None
    try:
        for group in definition_file.Groups:
            for definition in group.Definitions:
                if definition.Name == target_name:
                    return definition
    except Exception:
        return None
    return None



def ensure_shared_definition(definition_file, target_name, source_param, definition_group_name):
    if not definition_file:
        raise Exception('No shared parameter file is currently configured in Revit.')

    existing = find_shared_definition(definition_file, target_name)
    if existing is not None:
        return existing

    group_name = definition_group_name or 'Meinhardt'
    group = None
    try:
        group = definition_file.Groups[group_name]
    except Exception:
        group = None
    if group is None:
        group = definition_file.Groups.Create(group_name)

    try:
        datatype = source_param.Definition.GetDataType()
        options = DB.ExternalDefinitionCreationOptions(target_name, datatype)
    except Exception:
        datatype = source_param.Definition.ParameterType
        options = DB.ExternalDefinitionCreationOptions(target_name, datatype)
    return group.Definitions.Create(options)



def get_storage_type(family_param):
    try:
        return family_param.StorageType
    except Exception:
        return None



def capture_parameter_values(family_doc, family_param):
    values = []
    manager = family_doc.FamilyManager
    original_type = None
    try:
        original_type = manager.CurrentType
    except Exception:
        original_type = None

    try:
        for family_type in iter_family_types(family_doc):
            manager.CurrentType = family_type
            storage = get_storage_type(family_param)
            try:
                has_value = family_type.HasValue(family_param)
            except Exception:
                has_value = True
            if not has_value:
                values.append((family_type, False, None))
                continue

            if storage == DB.StorageType.String:
                value = family_type.AsString(family_param)
            elif storage == DB.StorageType.Double:
                value = family_type.AsDouble(family_param)
            elif storage == DB.StorageType.Integer:
                value = family_type.AsInteger(family_param)
            elif storage == DB.StorageType.ElementId:
                value = family_type.AsElementId(family_param)
            else:
                value = None
            values.append((family_type, True, value))
    finally:
        try:
            manager.CurrentType = original_type
        except Exception:
            pass
    return values



def restore_parameter_values(family_doc, family_param, values):
    manager = family_doc.FamilyManager
    original_type = None
    restored = 0
    try:
        original_type = manager.CurrentType
    except Exception:
        original_type = None

    try:
        for family_type, has_value, value in values:
            if not has_value:
                continue
            try:
                manager.CurrentType = family_type
                if value is None and get_storage_type(family_param) == DB.StorageType.String:
                    continue
                manager.Set(family_param, value)
                restored += 1
            except Exception:
                logger.warning('Failed to restore value for type {} and parameter {}.'.format(safe_name(family_type), family_param.Definition.Name))
    finally:
        try:
            manager.CurrentType = original_type
        except Exception:
            pass
    return restored



def build_plan_for_family(doc, family, rules, family_name_targets):
    plan = ChangePlan(family)
    if not is_editable_family(family):
        plan.is_editable = False
        plan.skipped.append('Family is not editable (likely in-place or protected).')
        return plan

    plan.target_family_name = make_unique_name(
        normalize_name(plan.family_name, rules.get('family_rules', {})),
        family_name_targets,
        current_name=plan.family_name
    )
    family_name_targets.append(plan.target_family_name)

    shared_rules = rules.get('shared_parameter_replacements', []) or []
    skip_names = set([(x or '').strip().lower() for x in (rules.get('skip_parameter_names', []) or [])])

    family_doc = None
    try:
        family_doc = doc.EditFamily(family)
        existing_param_names = get_existing_parameter_names(family_doc)
        reserved_names = list(existing_param_names)

        for family_param in iter_family_parameters(family_doc):
            param_name = family_param.Definition.Name
            low_name = param_name.strip().lower()

            if is_builtin_family_parameter(family_param):
                plan.skipped.append('Skipped built-in parameter: {}'.format(param_name))
                continue
            if low_name in skip_names:
                plan.skipped.append('Skipped by configuration: {}'.format(param_name))
                continue
            if has_formula(family_param):
                plan.skipped.append('Skipped formula-driven parameter: {}'.format(param_name))
                continue
            if is_reporting(family_param):
                plan.skipped.append('Skipped reporting parameter: {}'.format(param_name))
                continue

            if family_param.IsShared:
                mapping = find_shared_mapping(shared_rules, param_name)
                if mapping:
                    target_name = (mapping.get('new_name') or '').strip()
                    if target_name and target_name.lower() != low_name:
                        conflict = False
                        for reserved in reserved_names:
                            if reserved.lower() == target_name.lower() and reserved.lower() != low_name:
                                conflict = True
                                break
                        if conflict:
                            plan.skipped.append('Shared parameter target already exists: {} -> {}'.format(param_name, target_name))
                        else:
                            plan.shared_ops.append({
                                'old_name': param_name,
                                'new_name': target_name,
                                'definition_group': mapping.get('definition_group') or 'Meinhardt',
                                'target_group': mapping.get('target_group'),
                            })
                            reserved_names.append(target_name)
                continue

            target_name = normalize_name(param_name, rules.get('parameter_rules', {}))
            target_name = make_unique_name(target_name, reserved_names, current_name=param_name)
            if target_name != param_name:
                plan.rename_ops.append({
                    'old_name': param_name,
                    'new_name': target_name,
                })
                reserved_names.append(target_name)
    except Exception as exc:
        plan.preview_error = str(exc)
        plan.errors.append(traceback.format_exc())
    finally:
        try:
            if family_doc is not None:
                family_doc.Close(False)
        except Exception:
            pass

    return plan



def print_preview(plans):
    output.close_others()
    output.print_md('# {}'.format(TOOL_NAME))
    output.print_md('## Preview')
    output.print_md('Rules file: {}'.format(RULES_FILE.replace('\\', '/')))

    total_families = 0
    total_param_renames = 0
    total_shared_ops = 0
    total_family_renames = 0

    for plan in plans:
        if not plan.is_editable:
            output.print_md('### {}  '.format(plan.family_name))
            output.print_md('- Not editable')
            continue
        if plan.preview_error:
            output.print_md('### {}  '.format(plan.family_name))
            output.print_md('- Preview failed: {}'.format(plan.preview_error))
            continue

        total_families += 1
        total_param_renames += len(plan.rename_ops)
        total_shared_ops += len(plan.shared_ops)
        if plan.target_family_name != plan.family_name:
            total_family_renames += 1

        output.print_md('### {}'.format(plan.family_name))
        if plan.target_family_name != plan.family_name:
            output.print_md('- Family: `{}` → `{}`'.format(plan.family_name, plan.target_family_name))
        else:
            output.print_md('- Family: no rename needed')

        if plan.rename_ops:
            output.print_md('- Non-shared parameter renames:')
            for op in plan.rename_ops:
                output.print_md('  - `{}` → `{}`'.format(op['old_name'], op['new_name']))
        else:
            output.print_md('- Non-shared parameter renames: none')

        if plan.shared_ops:
            output.print_md('- Shared parameter replacements:')
            for op in plan.shared_ops:
                output.print_md('  - `{}` → `{}`'.format(op['old_name'], op['new_name']))
        else:
            output.print_md('- Shared parameter replacements: none')

        if plan.skipped:
            output.print_md('- Skipped items:')
            for msg in plan.skipped[:12]:
                output.print_md('  - {}'.format(msg))
            if len(plan.skipped) > 12:
                output.print_md('  - ... and {} more'.format(len(plan.skipped) - 12))

    output.print_md('---')
    output.print_md('**Summary:** {} editable families, {} family renames, {} non-shared parameter renames, {} shared replacement candidates.'.format(
        total_families, total_family_renames, total_param_renames, total_shared_ops
    ))



def replace_shared_parameter(family_doc, family_param, op, definition_file):
    target_group = resolve_target_group(op.get('target_group'), family_param.Definition.ParameterGroup)
    snapshot = capture_parameter_values(family_doc, family_param)
    target_definition = ensure_shared_definition(
        definition_file,
        op.get('new_name'),
        family_param,
        op.get('definition_group')
    )
    family_doc.FamilyManager.RemoveParameter(family_param)
    new_param = family_doc.FamilyManager.AddParameter(target_definition, target_group, family_param.IsInstance)
    restored_count = restore_parameter_values(family_doc, new_param, snapshot)
    return restored_count



def apply_plan(doc, plans, rules):
    load_options = FamilyLoadOptions(rules.get('reload_overwrite_parameter_values', False))
    shared_definition_file = None
    if any(plan.shared_ops for plan in plans):
        shared_definition_file = get_shared_definition_file()
        if shared_definition_file is None:
            _alert('Shared-parameter mappings were found, but no shared parameter file is configured in Revit. Shared replacements will be skipped.')

    project_renames = []
    summary = {
        'families_renamed': 0,
        'params_renamed': 0,
        'shared_replaced': 0,
        'shared_values_restored': 0,
        'errors': []
    }

    for plan in plans:
        if not plan.is_editable or plan.preview_error:
            continue
        if plan.total_changes == 0:
            continue

        family_doc = None
        try:
            family_doc = doc.EditFamily(plan.family)
            trans = DB.Transaction(family_doc, TOOL_NAME)
            trans.Start()
            try:
                for op in plan.rename_ops:
                    family_param = family_doc.FamilyManager.get_Parameter(op['old_name'])
                    if family_param is None:
                        summary['errors'].append('Parameter not found: {} in {}'.format(op['old_name'], plan.family_name))
                        continue
                    rename_family_parameter(family_doc, family_param, op['new_name'])
                    summary['params_renamed'] += 1

                for op in plan.shared_ops:
                    if shared_definition_file is None:
                        summary['errors'].append('Shared replacement skipped (no shared parameter file): {} in {}'.format(op['old_name'], plan.family_name))
                        continue

                    sub = DB.SubTransaction(family_doc)
                    sub.Start()
                    try:
                        family_param = family_doc.FamilyManager.get_Parameter(op['old_name'])
                        if family_param is None:
                            raise Exception('Parameter not found')
                        restored = replace_shared_parameter(family_doc, family_param, op, shared_definition_file)
                        summary['shared_replaced'] += 1
                        summary['shared_values_restored'] += restored
                        sub.Commit()
                    except Exception as exc:
                        try:
                            sub.RollBack()
                        except Exception:
                            pass
                        summary['errors'].append('Shared replacement failed for {} in {}: {}'.format(op['old_name'], plan.family_name, exc))

                trans.Commit()
            except Exception:
                trans.RollBack()
                raise

            family_doc.LoadFamily(doc, load_options)
            if plan.target_family_name != plan.family_name:
                project_renames.append((plan.family, plan.family_name, plan.target_family_name))
        except Exception as exc:
            summary['errors'].append('Family conversion failed for {}: {}'.format(plan.family_name, exc))
            logger.error(traceback.format_exc())
        finally:
            try:
                if family_doc is not None:
                    family_doc.Close(False)
            except Exception:
                pass

    if project_renames:
        t = DB.Transaction(doc, '{} - Rename Families'.format(TOOL_NAME))
        t.Start()
        try:
            for family, old_name, new_name in project_renames:
                try:
                    family.Name = new_name
                    summary['families_renamed'] += 1
                except Exception as exc:
                    summary['errors'].append('Family rename failed for {} -> {}: {}'.format(old_name, new_name, exc))
            t.Commit()
        except Exception:
            t.RollBack()
            raise

    return summary



def main():
    doc = revit.doc
    if doc.IsFamilyDocument:
        _alert('Run this tool from a project document. It opens selected loadable families for editing and reloads them after conversion.', exitscript=True)
        return

    try:
        rules = load_rules()
    except Exception as exc:
        _alert(str(exc), exitscript=True)
        return

    families = choose_families(doc)
    if not families:
        _alert('No families selected.')
        return

    family_name_targets = get_existing_family_names(doc)
    plans = []
    for family in families:
        plans.append(build_plan_for_family(doc, family, rules, family_name_targets))

    print_preview(plans)

    editable_plans = [p for p in plans if p.is_editable and not p.preview_error and p.total_changes > 0]
    if not editable_plans:
        _alert('No applicable changes were found. Review converter_rules.json if you expected renames.')
        return

    proceed = forms.alert(
        'Preview written to the output window. Apply {} family conversions now?'.format(len(editable_plans)),
        title=TOOL_NAME,
        yes=True,
        no=True
    )
    if not proceed:
        return

    summary = apply_plan(doc, plans, rules)
    output.print_md('## Apply Summary')
    output.print_md('- Families renamed: {}'.format(summary['families_renamed']))
    output.print_md('- Non-shared parameters renamed: {}'.format(summary['params_renamed']))
    output.print_md('- Shared parameters replaced: {}'.format(summary['shared_replaced']))
    output.print_md('- Shared type values restored: {}'.format(summary['shared_values_restored']))
    if summary['errors']:
        output.print_md('- Errors / warnings: {}'.format(len(summary['errors'])))
        for item in summary['errors']:
            output.print_md('  - {}'.format(item))
        _alert('Completed with {} warning(s). Check the output window for details.'.format(len(summary['errors'])))
    else:
        _alert('Conversion completed successfully.')


if __name__ == '__main__':
    main()
