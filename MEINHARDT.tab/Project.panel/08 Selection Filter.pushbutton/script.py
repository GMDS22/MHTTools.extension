# -*- coding: utf-8 -*-
"""Meinhardt - Selection Filter

Filters the *current selection* based on a chosen parameter and one or more values.
Supports keep/remove behavior (similar to selection filtering workflows).

IronPython-safe (no f-strings, no Python 3-only syntax).
"""
from __future__ import print_function

from pyrevit import revit, forms, script
from Autodesk.Revit import DB

try:
    from Autodesk.Revit import UI
    from Autodesk.Revit.UI.Selection import ObjectType
except Exception:
    UI = None
    ObjectType = None

try:
    from System.Collections.Generic import List
except Exception:
    List = None

# WinForms UI
try:
    from System.Windows.Forms import (
        Application, Form, SplitContainer, ListBox, CheckedListBox, Button, Label,
        DockStyle, SelectionMode, FlatStyle, CheckState, TextBox, ComboBox, Panel,
        TreeView, TreeNode
    )
    from System.Drawing import Size, Color, Point
except Exception:
    Application = None
    Form = object
    SplitContainer = None
    ListBox = None
    CheckedListBox = None
    Button = None
    Label = None
    TextBox = None
    ComboBox = None
    Panel = None
    TreeView = None
    TreeNode = None
    DockStyle = None
    SelectionMode = None
    FlatStyle = None
    CheckState = None
    Size = None
    Color = None
    Point = None

logger = script.get_logger()

# IronPython has `unicode`; CPython doesn't.
try:
    import __builtin__ as _builtins
except Exception:
    import builtins as _builtins
try:
    unicode = _builtins.unicode
except Exception:
    unicode = str


def _safe_str(x):
    try:
        if x is None:
            return ''
        return str(x)
    except Exception:
        try:
            return unicode(x)
        except Exception:
            return ''


def _get_selected_element_ids(uidoc):
    try:
        ids = uidoc.Selection.GetElementIds()
        return list(ids) if ids else []
    except Exception:
        return []


def _iter_element_params(el):
    try:
        for p in el.Parameters:
            yield p
    except Exception:
        return


def _find_param_on_element(el, param_name):
    if not param_name:
        return None

    # Fast path
    try:
        p = el.LookupParameter(param_name)
        if p:
            return p
    except Exception:
        pass

    # Case-insensitive fallback across all parameters
    try:
        target = param_name.lower()
        for p in _iter_element_params(el):
            try:
                n = p.Definition.Name
                if n and n.lower() == target:
                    return p
            except Exception:
                continue
    except Exception:
        pass

    return None


def _collect_elements_for_scope(doc, uidoc, scope_label):
    scope = _safe_str(scope_label)
    if scope == 'Current Selection':
        ids = _get_selected_element_ids(uidoc)
        els = []
        for eid in ids:
            try:
                el = doc.GetElement(eid)
                if el is not None:
                    els.append(el)
            except Exception:
                continue
        return els

    if scope == 'Current View':
        try:
            viewid = doc.ActiveView.Id
            return list(DB.FilteredElementCollector(doc, viewid).WhereElementIsNotElementType())
        except Exception:
            return []

    if scope == 'Entire Project':
        try:
            return list(DB.FilteredElementCollector(doc).WhereElementIsNotElementType())
        except Exception:
            return []

    return []


def _collect_param_names_from_bindings(doc):
    names = set()
    try:
        bmap = doc.ParameterBindings
        it = bmap.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            try:
                d = it.Key
                if d is None:
                    continue
                n = getattr(d, 'Name', None)
                if n:
                    names.add(_safe_str(n))
            except Exception:
                continue
    except Exception:
        pass
    return names


def _collect_param_names_from_elements(elements, max_scan=500):
    names = set()
    if not elements:
        return names
    count = 0
    for el in elements:
        if count >= max_scan:
            break
        count += 1
        for p in _iter_element_params(el):
            try:
                n = p.Definition.Name
                if n:
                    names.add(_safe_str(n))
            except Exception:
                continue
    return names


def _param_value_key_display(param, doc):
    """Return (key, display) for grouping values.

    key must be hashable and stable; display is user-facing.
    """
    if param is None:
        return ('<missing>', '<missing>')

    try:
        if not param.HasValue:
            return ('<empty>', '<empty>')
    except Exception:
        # Some params throw on HasValue
        pass

    try:
        st = param.StorageType
    except Exception:
        st = None

    # Prefer formatted display string when available
    try:
        disp = param.AsValueString()
        if disp:
            disp = _safe_str(disp)
        else:
            disp = ''
    except Exception:
        disp = ''

    try:
        if st == DB.StorageType.String:
            v = param.AsString()
            v = _safe_str(v)
            if not disp:
                disp = v
            key = ('s', v)
            return (key, disp if disp else '<empty>')

        if st == DB.StorageType.Integer:
            v = None
            try:
                v = int(param.AsInteger())
            except Exception:
                v = _safe_str(param.AsInteger())
            if not disp:
                disp = _safe_str(v)
            key = ('i', _safe_str(v))
            return (key, disp if disp else '<empty>')

        if st == DB.StorageType.Double:
            v = None
            try:
                v = float(param.AsDouble())
            except Exception:
                v = _safe_str(param.AsDouble())
            if not disp:
                disp = _safe_str(v)
            key = ('d', _safe_str(v))
            return (key, disp if disp else '<empty>')

        if st == DB.StorageType.ElementId:
            try:
                eid = param.AsElementId()
            except Exception:
                eid = None

            if eid is None or eid == DB.ElementId.InvalidElementId:
                return (('eid', '-1'), '<none>')

            # Try resolve to element name
            try:
                el = doc.GetElement(eid)
                if el is not None:
                    try:
                        n = el.Name
                    except Exception:
                        n = _safe_str(eid.IntegerValue)
                    if n:
                        if not disp:
                            disp = _safe_str(n)
                        return (('eid', _safe_str(eid.IntegerValue)), disp)
            except Exception:
                pass

            if not disp:
                disp = _safe_str(eid.IntegerValue)
            return (('eid', _safe_str(eid.IntegerValue)), disp)

    except Exception:
        pass

    # Unknown storage type
    try:
        if not disp:
            disp = _safe_str(param)
    except Exception:
        disp = '<value>'
    return (('u', disp), disp)


def main():
    doc = revit.doc
    uidoc = revit.uidoc

    # Default scope: if nothing is selected, fall back to Current View
    initial_scope = 'Current Selection'
    try:
        if not _get_selected_element_ids(uidoc):
            initial_scope = 'Current View'
    except Exception:
        initial_scope = 'Current View'

    elements = _collect_elements_for_scope(doc, uidoc, initial_scope)
    if not elements:
        forms.alert('No elements found in the chosen scope.', exitscript=True)

    # Collect parameter names from bindings + a sample of elements
    param_names = set()
    try:
        param_names |= _collect_param_names_from_bindings(doc)
    except Exception:
        pass
    try:
        param_names |= _collect_param_names_from_elements(elements, max_scan=500)
    except Exception:
        pass

    if not param_names:
        forms.alert('Could not read parameters from the chosen scope.', exitscript=True)

    if Application is None or TreeView is None or TreeNode is None:
        forms.alert('WinForms UI not available in this environment.', exitscript=True)

    class SelectionFilterForm(Form):
        def __init__(self, doc, uidoc, elements, param_names, initial_scope):
            self._doc = doc
            self._uidoc = uidoc
            self._elements = elements
            self._param_names = sorted(list(param_names))
            self._all_param_names = list(self._param_names)
            self._param_search_text = ''
            self._scope = _safe_str(initial_scope)

            # param_name -> set(value_key)
            self._checked_values_by_param = {}
            self._suppress_tree_events = False

            # Cache per parameter
            self._param_buckets_cache = {}   # param_name -> buckets
            self._param_elkey_cache = {}     # param_name -> { elementIdInt: value_key }

            self.Text = 'GM Selection Filter'
            try:
                self.MinimumSize = Size(900, 600)
            except Exception:
                pass

            # Meinhardt theme (matches Family Renamer)
            try:
                self.bg_dark = Color.FromArgb(30, 30, 30)
                self.bg_control = Color.FromArgb(45, 45, 45)
                self.bg_header = Color.FromArgb(20, 20, 20)
                self.fg_text = Color.FromArgb(220, 220, 220)
                # Two accent shades
                self.accent = Color.FromArgb(0, 120, 212)
                self.accent_dark = Color.FromArgb(0, 90, 160)
                self.BackColor = self.bg_dark
                self.ForeColor = self.fg_text
            except Exception:
                self.bg_dark = None
                self.bg_control = None
                self.bg_header = None
                self.fg_text = None
                self.accent = None
                self.accent_dark = None

            self._lblTop = Label()
            self._lblTop.Text = 'Expand a parameter and check value(s). Use Mode to Keep or Unselect matching elements.'
            self._lblTop.Dock = DockStyle.Top
            try:
                self._lblTop.Height = 30
            except Exception:
                pass
            try:
                if self.bg_dark is not None:
                    self._lblTop.BackColor = self.bg_dark
                if self.fg_text is not None:
                    self._lblTop.ForeColor = self.fg_text
            except Exception:
                pass

            self._split = SplitContainer()
            self._split.Dock = DockStyle.Fill
            try:
                self._split.SplitterDistance = 650
            except Exception:
                pass

            # Top bar: scope selector
            self._topbar = Panel()
            self._topbar.Dock = DockStyle.Top
            try:
                self._topbar.Height = 32
            except Exception:
                pass
            try:
                if self.bg_header is not None:
                    self._topbar.BackColor = self.bg_header
            except Exception:
                pass

            self._lblScope = Label()
            self._lblScope.Text = 'Scope:'
            try:
                self._lblScope.Location = Point(10, 8)
                self._lblScope.Size = Size(50, 18)
            except Exception:
                pass
            try:
                if self.bg_header is not None:
                    self._lblScope.BackColor = self.bg_header
                if self.fg_text is not None:
                    self._lblScope.ForeColor = self.fg_text
            except Exception:
                pass

            self._cmbScope = ComboBox()
            try:
                self._cmbScope.DropDownStyle = 2  # DropDownList
            except Exception:
                pass
            try:
                self._cmbScope.Location = Point(65, 6)
                self._cmbScope.Size = Size(180, 22)
            except Exception:
                pass
            try:
                if self.bg_control is not None:
                    self._cmbScope.BackColor = self.bg_control
                if self.fg_text is not None:
                    self._cmbScope.ForeColor = self.fg_text
            except Exception:
                pass
            try:
                self._cmbScope.Items.Add('Current Selection')
                self._cmbScope.Items.Add('Current View')
                self._cmbScope.Items.Add('Entire Project')
            except Exception:
                pass
            try:
                # Set initial scope selection
                idx = 0
                if self._scope == 'Current View':
                    idx = 1
                elif self._scope == 'Entire Project':
                    idx = 2
                self._cmbScope.SelectedIndex = idx
            except Exception:
                pass
            try:
                self._cmbScope.SelectedIndexChanged += self._on_scope_changed
            except Exception:
                pass

            try:
                self._topbar.Controls.Add(self._cmbScope)
                self._topbar.Controls.Add(self._lblScope)
            except Exception:
                pass

            self._lblParams = Label()
            self._lblParams.Text = 'Parameters (expand to see values)'
            self._lblParams.Dock = DockStyle.Top
            try:
                self._lblParams.Height = 20
            except Exception:
                pass
            try:
                if self.bg_header is not None:
                    self._lblParams.BackColor = self.bg_header
                if self.fg_text is not None:
                    self._lblParams.ForeColor = self.fg_text
            except Exception:
                pass

            self._lblParamSearch = Label()
            self._lblParamSearch.Text = 'Search:'
            self._lblParamSearch.Dock = DockStyle.Top
            try:
                self._lblParamSearch.Height = 18
            except Exception:
                pass
            try:
                if self.bg_header is not None:
                    self._lblParamSearch.BackColor = self.bg_header
                if self.fg_text is not None:
                    self._lblParamSearch.ForeColor = self.fg_text
            except Exception:
                pass

            self._txtParamSearch = TextBox()
            self._txtParamSearch.Dock = DockStyle.Top
            try:
                self._txtParamSearch.Height = 24
            except Exception:
                pass
            try:
                if self.bg_control is not None:
                    self._txtParamSearch.BackColor = self.bg_control
                if self.fg_text is not None:
                    self._txtParamSearch.ForeColor = self.fg_text
            except Exception:
                pass
            try:
                self._txtParamSearch.TextChanged += self._on_param_search_changed
            except Exception:
                pass

            self._tree = TreeView()
            self._tree.Dock = DockStyle.Fill
            try:
                self._tree.CheckBoxes = True
            except Exception:
                pass
            try:
                if self.bg_control is not None:
                    self._tree.BackColor = self.bg_control
                if self.fg_text is not None:
                    self._tree.ForeColor = self.fg_text
            except Exception:
                pass
            try:
                self._tree.BeforeExpand += self._on_tree_before_expand
            except Exception:
                pass
            try:
                self._tree.AfterCheck += self._on_tree_after_check
            except Exception:
                pass

            self._rebuild_tree('')

            left_panel = self._split.Panel1
            left_panel.Controls.Add(self._tree)
            left_panel.Controls.Add(self._lblParams)
            left_panel.Controls.Add(self._txtParamSearch)
            left_panel.Controls.Add(self._lblParamSearch)

            # Right panel: mode + apply
            self._lblMode = Label()
            self._lblMode.Text = 'Mode:'
            self._lblMode.Dock = DockStyle.Top
            try:
                self._lblMode.Height = 18
            except Exception:
                pass
            try:
                if self.bg_header is not None:
                    self._lblMode.BackColor = self.bg_header
                if self.fg_text is not None:
                    self._lblMode.ForeColor = self.fg_text
            except Exception:
                pass

            self._cmbMode = ComboBox()
            self._cmbMode.Dock = DockStyle.Top
            try:
                self._cmbMode.DropDownStyle = 2  # DropDownList
            except Exception:
                pass
            try:
                if self.bg_control is not None:
                    self._cmbMode.BackColor = self.bg_control
                if self.fg_text is not None:
                    self._cmbMode.ForeColor = self.fg_text
            except Exception:
                pass
            try:
                self._cmbMode.Items.Add('Keep Checked')
                self._cmbMode.Items.Add('Unselect Checked')
                self._cmbMode.SelectedIndex = 0
            except Exception:
                pass

            self._btnFilter = Button()
            self._btnFilter.Text = 'Apply Filter'
            self._btnFilter.Dock = DockStyle.Bottom
            try:
                self._btnFilter.Height = 35
            except Exception:
                pass
            try:
                if self.accent is not None:
                    self._btnFilter.BackColor = self.accent
                if Color is not None:
                    self._btnFilter.ForeColor = Color.White
                if FlatStyle is not None:
                    self._btnFilter.FlatStyle = FlatStyle.Flat
                try:
                    self._btnFilter.FlatAppearance.BorderColor = self.accent_dark
                    self._btnFilter.FlatAppearance.MouseOverBackColor = self.accent
                    self._btnFilter.FlatAppearance.MouseDownBackColor = self.accent_dark
                except Exception:
                    pass
            except Exception:
                pass
            self._btnFilter.Click += self._on_filter

            try:
                self._btnFilter.Enabled = True
            except Exception:
                pass

            right_panel = self._split.Panel2
            right_panel.Controls.Add(self._btnFilter)
            right_panel.Controls.Add(self._cmbMode)
            right_panel.Controls.Add(self._lblMode)

            self.Controls.Add(self._split)
            self.Controls.Add(self._topbar)
            self.Controls.Add(self._lblTop)

        def _rebuild_tree(self, search_text):
            self._param_search_text = _safe_str(search_text)
            needle = self._param_search_text.lower().strip()

            try:
                self._tree.Nodes.Clear()
            except Exception:
                return

            for n in self._all_param_names:
                pname = _safe_str(n)
                if needle and needle not in pname.lower():
                    continue
                try:
                    pnode = TreeNode(pname)
                    pnode.Tag = ('param', pname)
                    # Add placeholder child for lazy load
                    pnode.Nodes.Add(TreeNode('...'))
                    self._tree.Nodes.Add(pnode)
                except Exception:
                    continue

        def _on_param_search_changed(self, sender, args):
            try:
                self._rebuild_tree(self._txtParamSearch.Text)
            except Exception:
                pass

        def _on_scope_changed(self, sender, args):
            try:
                scope = _safe_str(self._cmbScope.SelectedItem)
            except Exception:
                scope = self._scope

            if not scope:
                return

            # Validate selection scope
            if scope == 'Current Selection':
                try:
                    if not _get_selected_element_ids(self._uidoc):
                        forms.alert('Nothing selected. Select elements or choose Current View/Entire Project.', exitscript=False)
                        # Revert
                        try:
                            if self._scope == 'Current View':
                                self._cmbScope.SelectedIndex = 1
                            elif self._scope == 'Entire Project':
                                self._cmbScope.SelectedIndex = 2
                            else:
                                self._cmbScope.SelectedIndex = 0
                        except Exception:
                            pass
                        return
                except Exception:
                    pass

            els = _collect_elements_for_scope(self._doc, self._uidoc, scope)
            if not els:
                forms.alert('No elements found in the chosen scope.', exitscript=False)
                return

            self._scope = scope
            self._elements = els
            self._param_buckets_cache = {}
            self._param_elkey_cache = {}

            # Refresh parameter names (bindings + sample)
            new_names = set()
            try:
                new_names |= _collect_param_names_from_bindings(self._doc)
            except Exception:
                pass
            try:
                new_names |= _collect_param_names_from_elements(self._elements, max_scan=500)
            except Exception:
                pass
            self._param_names = sorted(list(new_names))
            self._all_param_names = list(self._param_names)

            # Drop checked values for params that no longer exist
            try:
                self._checked_values_by_param = dict([(p, s) for p, s in self._checked_values_by_param.items() if p in new_names])
            except Exception:
                self._checked_values_by_param = {}

            try:
                self._rebuild_tree(self._txtParamSearch.Text)
            except Exception:
                self._rebuild_tree('')

        def _is_placeholder_node(self, node):
            try:
                if node is None:
                    return False
                if node.Nodes is None:
                    return False
                if node.Nodes.Count != 1:
                    return False
                child = node.Nodes[0]
                return _safe_str(getattr(child, 'Text', '')) == '...'
            except Exception:
                return False

        def _ensure_param_values_loaded(self, pnode):
            try:
                tag = getattr(pnode, 'Tag', None)
                if not tag or len(tag) < 2:
                    return
                if tag[0] != 'param':
                    return
                pname = _safe_str(tag[1])
            except Exception:
                return

            if not self._is_placeholder_node(pnode):
                return

            try:
                pnode.Nodes.Clear()
            except Exception:
                return

            try:
                buckets, _ = self._build_values_for_param(pname)
            except Exception:
                buckets = {}

            keys_sorted = []
            try:
                keys_sorted = sorted(buckets.keys(), key=lambda k: (_safe_str(buckets[k].get('disp', '')), _safe_str(k)))
            except Exception:
                keys_sorted = list(buckets.keys())

            checked_keys = set()
            try:
                checked_keys = set(self._checked_values_by_param.get(pname, set()))
            except Exception:
                checked_keys = set()

            for k in keys_sorted:
                try:
                    disp = buckets[k].get('disp', '')
                    cnt = buckets[k].get('count', 0)
                    val_text = _safe_str(disp) if disp else '<empty>'
                    label = u"{} ({})".format(val_text, cnt)
                except Exception:
                    label = _safe_str(k)

                try:
                    vnode = TreeNode(label)
                    vnode.Tag = ('value', pname, k)
                    try:
                        vnode.Checked = (k in checked_keys)
                    except Exception:
                        pass
                    pnode.Nodes.Add(vnode)
                except Exception:
                    continue

        def _on_tree_before_expand(self, sender, args):
            try:
                node = args.Node
            except Exception:
                node = None
            if node is None:
                return
            self._ensure_param_values_loaded(node)

        def _update_parent_checked_state(self, pnode):
            try:
                if pnode is None:
                    return
                if pnode.Nodes is None or pnode.Nodes.Count == 0:
                    return
                total = 0
                checked = 0
                for child in pnode.Nodes:
                    total += 1
                    try:
                        if child.Checked:
                            checked += 1
                    except Exception:
                        pass
                if total == 0:
                    return
                if checked == total:
                    pnode.Checked = True
                elif checked == 0:
                    pnode.Checked = False
                else:
                    # No tri-state; leave as-is
                    pass
            except Exception:
                pass

        def _on_tree_after_check(self, sender, args):
            if self._suppress_tree_events:
                return

            try:
                node = args.Node
            except Exception:
                node = None
            if node is None:
                return

            try:
                tag = getattr(node, 'Tag', None)
            except Exception:
                tag = None

            if not tag:
                return

            # Parameter node toggles all its values
            try:
                if tag[0] == 'param':
                    pname = _safe_str(tag[1])
                    self._ensure_param_values_loaded(node)
                    self._suppress_tree_events = True
                    try:
                        if node.Checked:
                            # Check all
                            keys = set()
                            for child in node.Nodes:
                                try:
                                    ctag = getattr(child, 'Tag', None)
                                    if ctag and ctag[0] == 'value':
                                        keys.add(ctag[2])
                                        child.Checked = True
                                except Exception:
                                    continue
                            self._checked_values_by_param[pname] = keys
                        else:
                            # Uncheck all
                            for child in node.Nodes:
                                try:
                                    child.Checked = False
                                except Exception:
                                    pass
                            if pname in self._checked_values_by_param:
                                del self._checked_values_by_param[pname]
                    finally:
                        self._suppress_tree_events = False
                    return
            except Exception:
                self._suppress_tree_events = False
                return

            # Value node toggles its own key
            try:
                if tag[0] == 'value':
                    pname = _safe_str(tag[1])
                    vkey = tag[2]
                    if pname not in self._checked_values_by_param:
                        self._checked_values_by_param[pname] = set()
                    if node.Checked:
                        self._checked_values_by_param[pname].add(vkey)
                    else:
                        try:
                            if vkey in self._checked_values_by_param[pname]:
                                self._checked_values_by_param[pname].remove(vkey)
                        except Exception:
                            pass
                        try:
                            if len(self._checked_values_by_param[pname]) == 0:
                                del self._checked_values_by_param[pname]
                        except Exception:
                            pass
                    try:
                        self._update_parent_checked_state(node.Parent)
                    except Exception:
                        pass
            except Exception:
                pass

        def _build_values_for_param(self, param_name):
            if not param_name:
                return {}, {}

            if param_name in self._param_buckets_cache and param_name in self._param_elkey_cache:
                return self._param_buckets_cache[param_name], self._param_elkey_cache[param_name]

            buckets = {}  # key -> {'disp': str, 'count': int}
            el_key = {}

            for el in self._elements:
                try:
                    p = _find_param_on_element(el, param_name)
                    key, disp = _param_value_key_display(p, self._doc)
                except Exception:
                    key, disp = ('<error>', '<error>')

                try:
                    el_key[el.Id.IntegerValue] = key
                except Exception:
                    pass

                if key not in buckets:
                    buckets[key] = {'disp': disp, 'count': 0}
                buckets[key]['count'] += 1

            self._param_buckets_cache[param_name] = buckets
            self._param_elkey_cache[param_name] = el_key
            return buckets, el_key

        def _on_filter(self, sender, args):
            criteria_by_param = {}
            try:
                criteria_by_param = dict([(p, set(v)) for p, v in self._checked_values_by_param.items() if v])
            except Exception:
                criteria_by_param = {}

            if not criteria_by_param:
                return

            mode = 'Keep Checked'
            try:
                mode = _safe_str(self._cmbMode.SelectedItem)
            except Exception:
                mode = 'Keep Checked'

            out_ids = []
            for el in self._elements:
                try:
                    elid_int = el.Id.IntegerValue
                    matches_all = True
                    for pname, selected_keys in criteria_by_param.items():
                        # Prefer cached element->key mapping
                        vkey = None
                        try:
                            elmap = self._param_elkey_cache.get(pname)
                            if elmap is not None:
                                vkey = elmap.get(elid_int)
                        except Exception:
                            vkey = None

                        if vkey is None:
                            try:
                                p = _find_param_on_element(el, pname)
                                vkey, _ = _param_value_key_display(p, self._doc)
                            except Exception:
                                vkey = '<error>'

                        # AND across parameters, OR within each parameter's selected values
                        if vkey not in selected_keys:
                            matches_all = False
                            break

                    keep_el = matches_all
                    if mode == 'Unselect Checked':
                        keep_el = not matches_all

                    if keep_el:
                        out_ids.append(el.Id)
                except Exception:
                    continue

            try:
                if List is None:
                    self._uidoc.Selection.SetElementIds(out_ids)
                else:
                    self._uidoc.Selection.SetElementIds(List[DB.ElementId](out_ids))
            except Exception:
                pass

            try:
                self.Close()
            except Exception:
                pass

    f = SelectionFilterForm(doc, uidoc, elements, param_names, initial_scope)
    try:
        f.ShowDialog()
    except Exception:
        # Fallback to modeless if needed
        try:
            f.Show()
        except Exception:
            pass


if __name__ == '__main__':
    main()
