# pylint: disable=import-error,invalid-name,broad-except
from pyrevit import revit, DB
from pyrevit import forms


def _safe_get_view_level(view):
    try:
        if hasattr(view, 'GenLevel'):
            return view.GenLevel
    except Exception:
        return None


def _find_link_doc_for_type(host_doc, link_type_id):
    try:
        insts = DB.FilteredElementCollector(host_doc) \
            .OfClass(DB.RevitLinkInstance) \
            .ToElements()
        for inst in insts:
            try:
                if inst.GetTypeId() == link_type_id:
                    ldoc = inst.GetLinkDocument()
                    if ldoc:
                        return ldoc
            except Exception:
                continue
    except Exception:
        pass
    return None


def _best_match_link_level(link_doc, host_level):
    if not link_doc or not host_level:
        return None

    try:
        link_levels = DB.FilteredElementCollector(link_doc) \
            .OfClass(DB.Level) \
            .ToElements()
        if not link_levels:
            return None

        host_name = None
        host_elev = None
        try:
            host_name = host_level.Name
        except Exception:
            host_name = None
        try:
            host_elev = host_level.Elevation
        except Exception:
            host_elev = None

        if host_name:
            for lvl in link_levels:
                try:
                    if lvl.Name == host_name:
                        return lvl
                except Exception:
                    continue

        if host_elev is not None:
            best = None
            best_abs = None
            for lvl in link_levels:
                try:
                    d = abs(lvl.Elevation - host_elev)
                    if best_abs is None or d < best_abs:
                        best_abs = d
                        best = lvl
                except Exception:
                    continue
            return best

    except Exception:
        return None

    return None


def _choose_linked_view_for_level(link_doc, src_link_view, dst_host_level):
    if not link_doc or not src_link_view or not dst_host_level:
        return None

    dst_link_level = _best_match_link_level(link_doc, dst_host_level)
    if not dst_link_level:
        return None

    src_viewtype = None
    try:
        src_viewtype = src_link_view.ViewType
    except Exception:
        src_viewtype = None

    src_template_id = None
    try:
        src_template_id = src_link_view.ViewTemplateId
    except Exception:
        src_template_id = None

    src_name = None
    try:
        src_name = src_link_view.Name
    except Exception:
        src_name = None

    # Try name swap using level names
    desired_name = None
    try:
        src_level = src_link_view.GenLevel if hasattr(src_link_view, 'GenLevel') else None
        src_level_name = src_level.Name if src_level else None
        dst_level_name = dst_link_level.Name
        if src_name and src_level_name and dst_level_name and (src_level_name in src_name):
            desired_name = src_name.replace(src_level_name, dst_level_name)
    except Exception:
        desired_name = None

    candidates = []
    try:
        for v in DB.FilteredElementCollector(link_doc).OfClass(DB.View).ToElements():
            try:
                if v.IsTemplate:
                    continue
                if src_viewtype is not None and v.ViewType != src_viewtype:
                    continue
                if not hasattr(v, 'GenLevel'):
                    continue
                if not v.GenLevel or v.GenLevel.Id != dst_link_level.Id:
                    continue
                candidates.append(v)
            except Exception:
                continue
    except Exception:
        return None

    if not candidates:
        return None

    if desired_name:
        for v in candidates:
            try:
                if v.Name == desired_name:
                    return v
            except Exception:
                continue

    if src_template_id and src_template_id != DB.ElementId.InvalidElementId:
        for v in candidates:
            try:
                if v.ViewTemplateId == src_template_id:
                    return v
            except Exception:
                continue

    return candidates[0]


def _collect_link_views(link_doc):
    views = []
    try:
        for v in DB.FilteredElementCollector(link_doc).OfClass(DB.View).ToElements():
            try:
                if v.IsTemplate:
                    continue
                views.append(v)
            except Exception:
                continue
    except Exception:
        return []
    return views


def _find_link_view_by_name(link_doc, view_name):
    if not link_doc or not view_name:
        return None
    try:
        for v in _collect_link_views(link_doc):
            try:
                if v.Name == view_name:
                    return v
            except Exception:
                continue
    except Exception:
        return None
    return None


def main():
    view = revit.active_view
    if view is None:
        forms.alert('No active view.')
        return

    link_vis_enum = getattr(DB, 'LinkVisibility', None)
    if not link_vis_enum:
        forms.alert('This Revit version/API does not expose LinkVisibility.')
        return

    options = [
        'All links in active view',
        'Pick links in view',
    ]
    scope = forms.CommandSwitchWindow.show(
        options,
        message='Select which links to modify:',
    )
    if not scope:
        return

    mode_options = [
        'By Host View',
        'By Linked View (match level)',
        'Custom (keep existing custom settings)',
    ]
    mode = forms.CommandSwitchWindow.show(
        mode_options,
        message='Select link display mode:',
    )
    if not mode:
        return

    linked_view_strategy = None
    linked_view_name = None
    if mode == 'By Linked View (match level)':
        linked_view_strategy = forms.CommandSwitchWindow.show(
            [
                'Match level (current behavior)',
                'Use host view name',
                'Pick linked view by name',
            ],
            message='Linked view selection strategy:',
        )
        if not linked_view_strategy:
            return

    halftone = forms.alert(
        'Set Halftone ON for these links?',
        yes=True,
        no=True,
    )
    if halftone is None:
        return

    # Build target link types
    target_link_types = []
    if scope == 'All links in active view':
        # Use link instances visible in this view
        try:
            insts = DB.FilteredElementCollector(revit.doc, view.Id) \
                .OfClass(DB.RevitLinkInstance) \
                .ToElements()
            type_ids = set()
            for inst in insts:
                try:
                    type_ids.add(inst.GetTypeId())
                except Exception:
                    continue
            for tid in type_ids:
                lt = revit.doc.GetElement(tid)
                if lt:
                    target_link_types.append(lt)
        except Exception:
            target_link_types = []
    else:
        # Pick in view
        from pyrevit import UI

        class _LinkPickFilter(UI.Selection.ISelectionFilter):
            def AllowElement(self, el):
                try:
                    return isinstance(el, DB.RevitLinkInstance)
                except Exception:
                    return False

            def AllowReference(self, ref, pt):
                return False

        try:
            refs = revit.uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element,
                _LinkPickFilter(),
                'Pick Revit links'
            )
        except Exception:
            return

        type_ids = set()
        for r in refs:
            try:
                inst = revit.doc.GetElement(r.ElementId)
                if inst:
                    type_ids.add(inst.GetTypeId())
            except Exception:
                continue
        for tid in type_ids:
            lt = revit.doc.GetElement(tid)
            if lt:
                target_link_types.append(lt)

    if not target_link_types:
        forms.alert('No links found.')
        return

    link_docs = {}
    for lt in target_link_types:
        try:
            ldoc = _find_link_doc_for_type(revit.doc, lt.Id)
            if ldoc:
                link_docs[lt.Id] = ldoc
        except Exception:
            continue

    if mode == 'By Linked View (match level)' and linked_view_strategy in (
        'Use host view name',
        'Pick linked view by name',
    ):
        if not link_docs:
            forms.alert('No loaded link documents found.')
            return

        all_names = []
        for ldoc in link_docs.values():
            names = []
            for v in _collect_link_views(ldoc):
                try:
                    names.append(v.Name)
                except Exception:
                    continue
            all_names.append(set(names))

        common_names = set.intersection(*all_names) if all_names else set()
        union_names = set().union(*all_names) if all_names else set()

        if linked_view_strategy == 'Use host view name':
            linked_view_name = view.Name
        else:
            if common_names:
                pick_from = sorted(common_names)
                pick_hint = 'Select a view name found in all linked files:'
            else:
                pick_from = sorted(union_names)
                pick_hint = ('Select a view name (not found in every link; '
                             'missing views will be skipped):')

            if not pick_from:
                forms.alert('No view names found in linked documents.')
                return

            linked_view_name = forms.SelectFromList.show(
                pick_from,
                multiselect=False,
                title='Select Linked View Name',
                button_name='Select',
                prompt=pick_hint,
            )
            if not linked_view_name:
                return

    # Apply
    dst_level = _safe_get_view_level(view)

    missing_link_views = []
    with revit.Transaction('Apply Link Display Settings'):
        for lt in target_link_types:
            try:
                ls = view.GetLinkOverrides(lt.Id)
                if not ls:
                    continue

                if mode == 'By Host View':
                    ls.LinkVisibilityType = link_vis_enum.ByHostView

                elif mode == 'By Linked View (match level)':
                    ls.LinkVisibilityType = link_vis_enum.ByLinkedView
                    if dst_level:
                        link_doc = _find_link_doc_for_type(revit.doc, lt.Id)
                        if link_doc:
                            # Choose a linked view based on current linked view if possible,
                            # otherwise just pick the first plan at the matching level.
                            src_link_view = None
                            try:
                                if ls.LinkedViewId and ls.LinkedViewId != DB.ElementId.InvalidElementId:
                                    src_link_view = link_doc.GetElement(ls.LinkedViewId)
                            except Exception:
                                src_link_view = None

                            if linked_view_strategy == 'Use host view name' or \
                               linked_view_strategy == 'Pick linked view by name':
                                chosen = _find_link_view_by_name(link_doc, linked_view_name)
                            elif src_link_view:
                                chosen = _choose_linked_view_for_level(
                                    link_doc, src_link_view, dst_level)
                            else:
                                chosen = None

                            if chosen:
                                ls.LinkedViewId = chosen.Id
                            elif linked_view_strategy in (
                                'Use host view name',
                                'Pick linked view by name',
                            ):
                                try:
                                    missing_link_views.append(lt.Name)
                                except Exception:
                                    missing_link_views.append('<Unknown Link>')

                else:
                    # Custom: keep whatever custom settings exist; no forced mode flip
                    pass

                # Halftone is not guaranteed on all API versions; set best-effort.
                try:
                    if hasattr(ls, 'Halftone'):
                        ls.Halftone = bool(halftone)
                except Exception:
                    pass

                view.SetLinkOverrides(lt.Id, ls)
            except Exception:
                continue

    if missing_link_views:
        forms.alert(
            'Done. Missing linked view "{}" in:\n{}'.format(
                linked_view_name,
                '\n'.join(sorted(set(missing_link_views)))
            )
        )
    else:
        forms.alert('Done.')


if __name__ == '__main__':
    main()
