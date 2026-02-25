# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals

__title__ = 'Callout Riser\nRenamer'
__author__ = 'Gino Moreno'
__doc__ = (
    'Rename selected callout views based on the reference level plus all '
    'Riser values found in Generic Model elements inside each callout view.'
)

from pyrevit import revit, DB, forms, script

logger = script.get_logger()

try:
    _TEXT_TYPE = unicode  # noqa: F821 (IronPython)
except Exception:
    _TEXT_TYPE = str


class CalloutRenameItem(object):
    def __init__(self, view, current_name, new_name, ref_level, riser_values):
        self.View = view
        self.ViewId = view.Id
        self.Selected = True
        self.CurrentName = current_name
        self.NewName = new_name
        self.ReferenceLevel = ref_level
        self.RiserValues = riser_values


class CalloutRiserRenamerWindow(forms.WPFWindow):
    def __init__(self, xaml_path, items):
        forms.WPFWindow.__init__(self, xaml_path)
        self.Title = 'Callout Riser Renamer'
        self.ApplyChanges = False
        self._items = list(items or [])
        self.CalloutGrid.ItemsSource = self._items
        try:
            self.CountText.Text = str(len(self._items))
        except Exception:
            pass

    def apply_click(self, sender, args):
        self.ApplyChanges = True
        self.Close()

    def close_click(self, sender, args):
        self.ApplyChanges = False
        self.Close()

    def select_all_click(self, sender, args):
        for item in self._items:
            item.Selected = True
        try:
            self.CalloutGrid.Items.Refresh()
        except Exception:
            pass

    def clear_all_click(self, sender, args):
        for item in self._items:
            item.Selected = False
        try:
            self.CalloutGrid.Items.Refresh()
        except Exception:
            pass


def _safe_str(value):
    try:
        if value is None:
            return ''
        return _TEXT_TYPE(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return ''


def _get_reference_level_name(view, doc):
    try:
        level = getattr(view, 'GenLevel', None)
        if level:
            return _safe_str(level.Name)
    except Exception:
        pass

    for bip_name in (
        'VIEW_REFERENCE_LEVEL',
        'VIEWER_LEVEL',
        'VIEWER_ASSOC_LEVEL',
        'VIEW_GENLEVEL'
    ):
        try:
            bip = getattr(DB.BuiltInParameter, bip_name, None)
        except Exception:
            bip = None
        if not bip:
            continue
        try:
            p = view.get_Parameter(bip)
            if not p:
                continue
            if p.StorageType == DB.StorageType.ElementId:
                lvl = doc.GetElement(p.AsElementId())
                if lvl:
                    return _safe_str(lvl.Name)
            val = p.AsString() or p.AsValueString()
            if val:
                return _safe_str(val)
        except Exception:
            pass

    return '<No Level>'


def _get_param_value(doc, element, param_name):
    try:
        p = element.LookupParameter(param_name)
    except Exception:
        p = None

    if not p:
        try:
            for prm in element.Parameters:
                try:
                    if prm.Definition and prm.Definition.Name and prm.Definition.Name.lower() == param_name.lower():
                        p = prm
                        break
                except Exception:
                    pass
        except Exception:
            p = None

    if not p or not p.HasValue:
        return None

    try:
        if p.StorageType == DB.StorageType.String:
            return _safe_str(p.AsString())
        if p.StorageType == DB.StorageType.Double:
            return _safe_str(p.AsValueString() or p.AsDouble())
        if p.StorageType == DB.StorageType.Integer:
            return _safe_str(p.AsValueString() or p.AsInteger())
        if p.StorageType == DB.StorageType.ElementId:
            elem = doc.GetElement(p.AsElementId())
            return _safe_str(elem.Name if elem else None)
    except Exception:
        pass

    try:
        return _safe_str(p.AsValueString() or p.AsString())
    except Exception:
        return None


def _collect_riser_values(doc, view):
    values = []
    try:
        collector = (DB.FilteredElementCollector(doc, view.Id)
                     .OfCategory(DB.BuiltInCategory.OST_GenericModel)
                     .WhereElementIsNotElementType())
        for elem in collector:
            val = _get_param_value(doc, elem, 'Riser')
            if val:
                values.append(val)
    except Exception as ex:
        logger.debug('Failed collecting riser values: %s', ex)

    # unique, preserve order
    seen = set()
    unique_vals = []
    for v in values:
        if v in seen:
            continue
        seen.add(v)
        unique_vals.append(v)
    return unique_vals


def _is_valid_view(view):
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass

    try:
        invalid_types = [DB.ViewType.Schedule, DB.ViewType.DrawingSheet]
        proj_browser = getattr(DB.ViewType, 'ProjectBrowser', None)
        if proj_browser is not None:
            invalid_types.append(proj_browser)
        if view.ViewType in tuple(invalid_types):
            return False
    except Exception:
        pass

    return True


def _try_get_view_from_element(doc, element):
    if isinstance(element, DB.View):
        return element

    # Try to discover a referenced view through element id parameters
    try:
        for prm in element.Parameters:
            try:
                if prm.StorageType != DB.StorageType.ElementId:
                    continue
                vid = prm.AsElementId()
                if not vid or vid == DB.ElementId.InvalidElementId:
                    continue
                v = doc.GetElement(vid)
                if isinstance(v, DB.View):
                    return v
            except Exception:
                pass
    except Exception:
        pass

    return None


def _get_selected_views(uidoc, doc):
    views = []
    try:
        sel_ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        sel_ids = []

    for eid in sel_ids:
        elem = doc.GetElement(eid)
        if not elem:
            continue
        view = _try_get_view_from_element(doc, elem)
        if view and _is_valid_view(view):
            views.append(view)

    # de-dup by id
    uniq = []
    seen = set()
    for v in views:
        if v.Id.IntegerValue in seen:
            continue
        seen.add(v.Id.IntegerValue)
        uniq.append(v)
    return uniq


def _apply_callout_name(view, new_name):
    updated = False
    if not new_name:
        return False

    try:
        if view.Name != new_name:
            view.Name = new_name
            updated = True
    except Exception:
        pass

    for pname in ('Callout', 'Callout Name', 'CalloutName', 'Callout View Name'):
        try:
            p = view.LookupParameter(pname)
            if p and not p.IsReadOnly:
                p.Set(new_name)
                updated = True
        except Exception:
            pass

    return updated


def main():
    doc = revit.doc
    uidoc = revit.uidoc

    views = _get_selected_views(uidoc, doc)
    if not views:
        forms.alert(
            'Select one or more callout views (or callout elements) before running this tool.',
            title='Callout Riser Renamer'
        )
        return

    items = []
    for view in views:
        ref_level = _get_reference_level_name(view, doc)
        risers = _collect_riser_values(doc, view)
        riser_text = ', '.join(risers) if risers else 'No Riser values'
        if risers:
            new_name = u'{} - {}'.format(ref_level, ', '.join(risers))
        else:
            new_name = u'{}'.format(ref_level)
        items.append(CalloutRenameItem(view, _safe_str(view.Name), new_name, ref_level, riser_text))

    xaml_path = script.get_bundle_file('CalloutRiserRenamer.xaml')
    window = CalloutRiserRenamerWindow(xaml_path, items)
    window.ShowDialog()

    if not window.ApplyChanges:
        return

    to_rename = [i for i in items if i.Selected]
    if not to_rename:
        forms.alert('No callouts selected for rename.', title='Callout Riser Renamer')
        return

    renamed = 0
    failed = []

    with revit.Transaction('Rename callouts from Riser parameter'):
        for item in to_rename:
            try:
                if _apply_callout_name(item.View, item.NewName):
                    renamed += 1
            except Exception as ex:
                failed.append((item.CurrentName, _safe_str(ex)))

    if failed:
        msg = 'Renamed: {}\nFailed: {}\n\n'.format(renamed, len(failed))
        msg += '\n'.join(['{} -> {}'.format(name, err) for name, err in failed])
        forms.alert(msg, title='Callout Riser Renamer')
    else:
        forms.alert('Renamed {} callout view(s).'.format(renamed), title='Callout Riser Renamer')


if __name__ == '__main__':
    main()
