# -*- coding: utf-8 -*-
from __future__ import print_function

__title__ = 'Search Tool'
__author__ = 'GM'
__doc__ = 'Search and run any loaded pyRevit tool by keyword.'

import os

from pyrevit import forms


def _safe_str(value):
    try:
        if value is None:
            return ''
        return str(value)
    except Exception:
        return ''


def _get_known_pyrevit_tab_titles():
    tab_titles = set()
    appdata_dir = os.environ.get('APPDATA')
    if not appdata_dir:
        return tab_titles

    exts_dir = os.path.join(appdata_dir, 'pyReVit', 'Extensions')
    if not os.path.isdir(exts_dir):
        exts_dir = os.path.join(appdata_dir, 'pyRevit', 'Extensions')
    if not os.path.isdir(exts_dir):
        return tab_titles

    def _add_tab_title(tab_title_raw):
        tab_title_raw = (tab_title_raw or '').strip()
        if not tab_title_raw:
            return
        tab_titles.add(tab_title_raw)
        tab_titles.add(tab_title_raw.replace('_', ' '))


    def _add_tabs_in_dir(parent_dir):
        try:
            for child in os.listdir(parent_dir):
                if child.lower().endswith('.tab'):
                    _add_tab_title(child[:-4])
        except Exception:
            pass

    # Tabs can exist directly under Extensions or inside .extension folders
    _add_tabs_in_dir(exts_dir)
    try:
        for entry in os.listdir(exts_dir):
            full = os.path.join(exts_dir, entry)
            if os.path.isdir(full) and entry.lower().endswith('.extension'):
                _add_tabs_in_dir(full)
    except Exception:
        pass

    return tab_titles


def _extract_tooltip_text(tooltip_obj):
    # Tooltip can be a string, or an Autodesk.Windows.RibbonToolTip-like object.
    if tooltip_obj is None:
        return ''

    if isinstance(tooltip_obj, str):
        return tooltip_obj

    parts = []
    for attr in ['Title', 'Content', 'Text', 'Description']:
        try:
            v = getattr(tooltip_obj, attr, None)
            sv = _safe_str(v).strip()
            if sv:
                parts.append(sv)
        except Exception:
            continue

    return ' | '.join([p for p in parts if p])


def _iter_children(item):
    # Many ribbon containers expose `Items`.
    try:
        children = getattr(item, 'Items', None)
        if children is not None:
            for c in children:
                yield c
    except Exception:
        pass


def _collect_ribbon_tools():
    try:
        from Autodesk.Windows import ComponentManager
    except Exception:
        return []

    ribbon = getattr(ComponentManager, 'Ribbon', None)
    if ribbon is None:
        return []

    known_tabs = _get_known_pyrevit_tab_titles()

    tools = []
    seen_ids = set()

    for tab in getattr(ribbon, 'Tabs', []) or []:
        tab_title = _safe_str(getattr(tab, 'Title', None)).strip()
        tab_title_alt = tab_title.replace('_', ' ')
        if known_tabs and (tab_title not in known_tabs and tab_title_alt not in known_tabs):
            continue

        for panel in getattr(tab, 'Panels', []) or []:
            psource = getattr(panel, 'Source', None)
            panel_title = _safe_str(getattr(psource, 'Title', None)).strip()
            items = getattr(psource, 'Items', None)
            if items is None:
                continue

            stack = list(items)
            while stack:
                item = stack.pop(0)
                if item is None:
                    continue

                try:
                    item_id = _safe_str(getattr(item, 'Id', None))
                    if item_id and item_id in seen_ids:
                        continue
                    if item_id:
                        seen_ids.add(item_id)
                except Exception:
                    pass

                # Recurse into containers
                for child in _iter_children(item):
                    stack.append(child)

                # Detect executable buttons
                try:
                    text = _safe_str(getattr(item, 'Text', None)).strip()
                    if not text:
                        continue
                    is_visible = True
                    try:
                        is_visible = bool(getattr(item, 'IsVisible', True))
                    except Exception:
                        pass
                    if not is_visible:
                        continue

                    cmd = getattr(item, 'CommandHandler', None)
                    if cmd is None:
                        continue
                except Exception:
                    continue

                tooltip_txt = _extract_tooltip_text(getattr(item, 'ToolTip', None))
                tools.append((tab_title, panel_title, text, tooltip_txt, item))

    # stable sort
    tools.sort(key=lambda r: (r[0].lower(), r[1].lower(), r[2].lower()))
    return tools


class _ToolItem(object):
    def __init__(self, tab_title, panel_title, tool_title, tooltip, ribbon_button):
        self.tab_title = tab_title
        self.panel_title = panel_title
        self.tool_title = tool_title
        self.tooltip = tooltip
        self.ribbon_button = ribbon_button
        if tooltip:
            self.display = '{}  —  {} > {}  —  {}'.format(tool_title, tab_title, panel_title, tooltip)
        else:
            self.display = '{}  —  {} > {}'.format(tool_title, tab_title, panel_title)


def _run_ribbon_button(ribbon_button):
    try:
        cmd = getattr(ribbon_button, 'CommandHandler', None)
        if cmd is None:
            forms.alert('Selected item does not have a command handler.', title='Search Tool')
            return
        try:
            can = True
            try:
                can = bool(cmd.CanExecute(None))
            except Exception:
                pass
            if not can:
                forms.alert('Selected tool is currently disabled.', title='Search Tool')
                return
        except Exception:
            pass

        cmd.Execute(None)
    except Exception as ex:
        forms.alert('Failed to run tool.\n\n{}'.format(ex), title='Search Tool')


def main():
    tools = _collect_ribbon_tools()
    if not tools:
        forms.alert('No pyRevit tools found to search.\n\nTip: make sure pyRevit ribbon is loaded.', title='Search Tool')
        return

    items = [_ToolItem(t, p, n, tt, btn) for (t, p, n, tt, btn) in tools]
    picked = forms.SelectFromList.show(
        items,
        name_attr='display',
        multiselect=False,
        title='Search Tool (type to filter)'
    )
    if not picked:
        return

    _run_ribbon_button(picked.ribbon_button)


if __name__ == '__main__':
    main()
