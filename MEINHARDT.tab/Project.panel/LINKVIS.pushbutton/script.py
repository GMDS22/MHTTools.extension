# pylint: disable=import-error,invalid-name,broad-except
import clr
clr.AddReference('System')

from pyrevit import revit, DB
from pyrevit import forms, script
from pyrevit.framework import List


def _pick_views():
    # Prefer pyRevit's built-in picker; allow selecting from current selection.
    views = forms.select_views(
        title='Select Views',
        multiple=True,
        use_selection=True
    )
    return views or []


def _pick_link_types(doc):
    link_types = DB.FilteredElementCollector(doc) \
        .OfClass(DB.RevitLinkType) \
        .ToElements()

    if not link_types:
        return []

    name_map = {}
    for lt in link_types:
        try:
            name = lt.Name
        except Exception:
            continue
        # De-dupe by appending id if names collide
        key = name
        if key in name_map:
            key = '{} ({})'.format(name, lt.Id.IntegerValue)
        name_map[key] = lt

    picked = forms.SelectFromList.show(
        sorted(name_map.keys()),
        multiselect=True,
        title='Select Revit Links',
        button_name='Select'
    )

    if not picked:
        return []

    return [name_map[x] for x in picked if x in name_map]


def _resolve_targets(doc, views, template_policy):
    """Return a list of views/templates to edit based on policy.

    template_policy:
      - 'Apply changes to View Template'
      - 'Detach templates (set None)'
    """
    targets = []
    templates_seen = set()

    for v in views:
        if not v:
            continue

        # If user selected a template directly, treat it as a target.
        try:
            if getattr(v, 'IsTemplate', False):
                if v.Id.IntegerValue not in templates_seen:
                    templates_seen.add(v.Id.IntegerValue)
                    targets.append(v)
                continue
        except Exception:
            pass

        vtid = None
        try:
            vtid = v.ViewTemplateId
        except Exception:
            vtid = DB.ElementId.InvalidElementId

        if vtid and vtid != DB.ElementId.InvalidElementId:
            if template_policy == 'Apply changes to View Template':
                vt = doc.GetElement(vtid)
                if vt and vt.Id.IntegerValue not in templates_seen:
                    templates_seen.add(vt.Id.IntegerValue)
                    targets.append(vt)
                continue

            if template_policy == 'Detach templates (set None)':
                try:
                    v.ViewTemplateId = DB.ElementId.InvalidElementId
                except Exception:
                    # If we can't detach, skip modifying this view
                    continue

        targets.append(v)

    return targets


def main():
    views = _pick_views()
    if not views:
        script.exit()

    # If any views have templates, ask what to do.
    templated = []
    for v in views:
        try:
            if getattr(v, 'IsTemplate', False):
                continue
            if v.ViewTemplateId and v.ViewTemplateId != DB.ElementId.InvalidElementId:
                templated.append(v)
        except Exception:
            continue

    template_policy = None
    if templated:
        template_policy = forms.CommandSwitchWindow.show(
            [
                'Apply changes to View Template',
                'Detach templates (set None)'
            ],
            message='Some selected views have view templates. How should LINKVIS handle them?'
        )
        if not template_policy:
            script.exit()
    else:
        # No templated views; policy doesn't matter
        template_policy = 'Apply changes to View Template'

    action = forms.CommandSwitchWindow.show(
        ['Hide', 'Unhide'],
        message='Hide or Unhide selected Revit links?'
    )
    if not action:
        script.exit()

    link_types = _pick_link_types(revit.doc)
    if not link_types:
        script.exit()

    targets = _resolve_targets(revit.doc, views, template_policy)
    if not targets:
        forms.alert('No editable target views/templates found.')
        return

    changed = 0
    skipped = 0

    with revit.Transaction('LINKVIS - {}'.format(action)):
        for v in targets:
            for lt in link_types:
                try:
                    if not lt.CanBeHidden(v):
                        skipped += 1
                        continue

                    if action == 'Hide':
                        if not lt.IsHidden(v):
                            v.HideElements(List[DB.ElementId]([lt.Id]))
                            changed += 1
                    else:
                        if lt.IsHidden(v):
                            v.UnhideElements(List[DB.ElementId]([lt.Id]))
                            changed += 1
                except Exception:
                    skipped += 1
                    continue

    forms.alert('Done. Changed: {}   Skipped: {}'.format(changed, skipped))


if __name__ == '__main__':
    main()
