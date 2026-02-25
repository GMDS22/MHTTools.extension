# -*- coding: utf-8 -*-
from __future__ import print_function

__title__ = 'Search Tool'
__author__ = 'GMoreno'
__doc__ = 'Search and run any loaded pyRevit tool by keyword.'

import os

from pyrevit import forms
from pyrevit import script
from pyrevit.loader import sessionmgr
from pyrevit.coreutils.ribbon import load_bitmapimage


logger = script.get_logger()

# =============================================================================
# GLOBAL CACHE - persists across tool invocations for speed
# =============================================================================
_CACHED_TOOLS = None


def _safe_str(value):
    try:
        if value is None:
            return ''
        return str(value)
    except Exception:
        return ''


def _split_extension_ui(script_path):
    """Extract tab/panel names from a bundle script path."""
    if not script_path:
        return '', ''
    try:
        parts = os.path.normpath(script_path).split(os.sep)
    except Exception:
        return '', ''

    tab_title = ''
    panel_title = ''
    for p in parts:
        pl = p.lower()
        if pl.endswith('.tab') and not tab_title:
            tab_title = p[:-4]
        elif pl.endswith('.panel') and not panel_title:
            panel_title = p[:-6]
    return tab_title, panel_title


def _pick_icon_for_bundle(script_path):
    """Load icon for a tool bundle."""
    try:
        if not script_path:
            return None
        bundle_dir = os.path.dirname(script_path)
        for name in ('icon.dark.png', 'icon.png'):
            path = os.path.join(bundle_dir, name)
            if os.path.isfile(path):
                return load_bitmapimage(path)
    except Exception:
        pass
    return None


def _collect_pyrevit_commands():
    """Collect all available pyRevit commands."""
    cmds = []
    try:
        for cmd in sessionmgr.find_all_available_commands(use_current_context=False, cache=True):
            if cmd and cmd.name:
                cmds.append(cmd)
    except Exception as ex:
        logger.debug('Failed collecting pyRevit commands: %s', ex)
    return cmds


def _get_cached_tools(force_refresh=False):
    """Get tool items from cache or build them."""
    global _CACHED_TOOLS
    if _CACHED_TOOLS is None or force_refresh:
        cmds = _collect_pyrevit_commands()
        _CACHED_TOOLS = [_ToolItem(c) for c in cmds]
    return _CACHED_TOOLS


class _ToolItem(object):
    """Wrapper for a pyRevit command with display properties."""
    def __init__(self, pyrvt_cmd):
        self.cmd = pyrvt_cmd
        self.tool_title = _safe_str(getattr(pyrvt_cmd, 'name', None)).strip()
        self.script_path = _safe_str(getattr(pyrvt_cmd, 'script', None)).strip()
        self.extension = _safe_str(getattr(pyrvt_cmd, 'extension', None)).strip()

        tab_title, panel_title = _split_extension_ui(self.script_path)
        self.tab_title = tab_title
        self.panel_title = panel_title

        self.location = u'{} > {}'.format(self.tab_title, self.panel_title) if (self.tab_title or self.panel_title) else self.extension

        # Load icon
        self.icon = _pick_icon_for_bundle(self.script_path)

        # Build search string (lowercase for fast matching)
        self._search = u'{} {} {} {}'.format(
            self.tool_title,
            self.tab_title,
            self.panel_title,
            self.extension or ''
        ).lower()

    def matches(self, q):
        if not q:
            return True
        return q in self._search


class SearchToolWindow(forms.WPFWindow):
    """Modal search window for pyRevit tools."""
    def __init__(self, xaml_path, tools):
        forms.WPFWindow.__init__(self, xaml_path)
        self.Title = 'Search Tool by GM'
        self.selected_tool = None

        self._all_tools = list(tools or [])
        self._filtered_tools = list(self._all_tools)

        self.ToolsList.ItemsSource = self._filtered_tools
        self.CountText.Text = str(len(self._filtered_tools))

        try:
            self.StatusText.Text = '{} tools. Type to filter, double-click or press Enter to run.'.format(len(self._all_tools))
        except Exception:
            pass

        try:
            if self._filtered_tools:
                self.ToolsList.SelectedIndex = 0
            self.SearchBox.Focus()
        except Exception:
            pass

    def _apply_filter(self):
        try:
            q = (self.SearchBox.Text or '').strip().lower()
        except Exception:
            q = ''

        if not q:
            self._filtered_tools = list(self._all_tools)
        else:
            self._filtered_tools = [t for t in self._all_tools if t.matches(q)]

        self.ToolsList.ItemsSource = self._filtered_tools
        self.CountText.Text = str(len(self._filtered_tools))
        try:
            if self._filtered_tools:
                self.ToolsList.SelectedIndex = 0
        except Exception:
            pass

    def search_changed(self, sender, args):
        self._apply_filter()

    def run_selected(self, sender, args):
        try:
            picked = self.ToolsList.SelectedItem
        except Exception:
            picked = None

        if not picked:
            return

        self.selected_tool = picked
        self.Close()

    def close_click(self, sender, args):
        self.selected_tool = None
        self.Close()


def main():
    # Get tools from cache (instant after first load)
    tools = _get_cached_tools()
    if not tools:
        forms.alert(
            'No pyRevit tools found.\n\nMake sure pyRevit extensions are loaded.',
            title='Search Tool'
        )
        return

    xaml_path = script.get_bundle_file('SearchTool.xaml')

    # Show modal window
    window = SearchToolWindow(xaml_path, tools)
    window.ShowDialog()

    # Run selected tool after window closes
    selected = window.selected_tool
    if not selected:
        return

    try:
        sessionmgr.execute_command_cls(selected.cmd.extcmd_type, exec_from_ui=True)
    except Exception as ex:
        forms.alert('Failed to run {}.\n\n{}'.format(selected.tool_title, ex), title='Search Tool')


if __name__ == '__main__':
    main()
