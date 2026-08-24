# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import shutil
import stat
import time

from pyrevit import forms, script


TOOL_TITLE = 'Free Space'
SECONDS_PER_DAY = 86400.0

logger = script.get_logger()
output = script.get_output()


def _format_bytes(size_bytes):
    units = ['B', 'KB', 'MB', 'GB', 'TB']
    size = float(max(size_bytes, 0))
    for unit in units:
        if size < 1024.0 or unit == units[-1]:
            if unit == 'B':
                return '{0:.0f} {1}'.format(size, unit)
            return '{0:.2f} {1}'.format(size, unit)
        size /= 1024.0
    return '0 B'


def _safe_listdir(path):
    try:
        return os.listdir(path)
    except Exception as ex:
        logger.debug('Could not list directory {0}: {1}'.format(path, ex))
        return []


def _collect_path_info(path):
    total = 0
    latest_mtime = 0.0

    if not path or not os.path.exists(path):
        return total, latest_mtime

    if os.path.isfile(path):
        try:
            return os.path.getsize(path), os.path.getmtime(path)
        except Exception:
            return 0, 0.0

    try:
        latest_mtime = os.path.getmtime(path)
    except Exception:
        latest_mtime = 0.0

    for root, _, files in os.walk(path):
        try:
            root_mtime = os.path.getmtime(root)
            if root_mtime > latest_mtime:
                latest_mtime = root_mtime
        except Exception:
            pass

        for filename in files:
            full_path = os.path.join(root, filename)
            try:
                total += os.path.getsize(full_path)
            except Exception:
                continue
            try:
                file_mtime = os.path.getmtime(full_path)
                if file_mtime > latest_mtime:
                    latest_mtime = file_mtime
            except Exception:
                continue

    return total, latest_mtime


def _path_age_seconds(path, latest_mtime=None):
    if latest_mtime is None:
        _, latest_mtime = _collect_path_info(path)
    if latest_mtime <= 0.0:
        return 0.0
    return max(0.0, time.time() - latest_mtime)


def _remove_readonly(func, path, excinfo):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        pass


def _delete_path(path):
    if os.path.isdir(path) and not os.path.islink(path):
        shutil.rmtree(path, onerror=_remove_readonly)
    else:
        os.chmod(path, stat.S_IWRITE)
        os.remove(path)


def _local_appdata():
    return os.environ.get('LOCALAPPDATA', '')


def _temp_dir():
    temp_dir = os.environ.get('TEMP', '')
    if temp_dir:
        return temp_dir
    local_appdata = _local_appdata()
    if local_appdata:
        return os.path.join(local_appdata, 'Temp')
    return ''


def _revit_root():
    local_appdata = _local_appdata()
    if not local_appdata:
        return ''
    return os.path.join(local_appdata, 'Autodesk', 'Revit')


def _iter_revit_cache_paths(child_folder_name):
    cache_paths = []
    revit_root = _revit_root()
    if not revit_root or not os.path.isdir(revit_root):
        return cache_paths
    for folder_name in _safe_listdir(revit_root):
        if not folder_name.startswith('Autodesk Revit '):
            continue
        full_path = os.path.join(revit_root, folder_name, child_folder_name)
        if os.path.isdir(full_path):
            cache_paths.append((folder_name, full_path))
    return cache_paths


def _build_safe_profiles():
    profiles = []
    temp_path = _temp_dir()
    if temp_path:
        profiles.append({
            'label': 'Windows Temp',
            'path': temp_path,
            'age_days': 1,
        })

    for revit_name, journal_path in _iter_revit_cache_paths('Journals'):
        profiles.append({
            'label': '{0} Journals'.format(revit_name),
            'path': journal_path,
            'age_days': 7,
        })

    for revit_name, cef_cache_path in _iter_revit_cache_paths('CefCache'):
        profiles.append({
            'label': '{0} CefCache'.format(revit_name),
            'path': cef_cache_path,
            'age_days': 7,
        })

    return profiles


def _build_deep_profiles():
    profiles = []
    for revit_name, collab_path in _iter_revit_cache_paths('CollaborationCache'):
        profiles.append({
            'label': '{0} CollaborationCache'.format(revit_name),
            'path': collab_path,
        })

    pac_cache = os.path.join(_revit_root(), 'PacCache') if _revit_root() else ''
    if pac_cache and os.path.isdir(pac_cache):
        profiles.append({
            'label': 'Revit PacCache',
            'path': pac_cache,
        })

    return profiles


def _scan_profile(profile):
    eligible_items = []
    total_bytes = 0
    cutoff_seconds = profile['age_days'] * SECONDS_PER_DAY
    path = profile['path']
    if not path or not os.path.isdir(path):
        return eligible_items, total_bytes

    for name in _safe_listdir(path):
        full_path = os.path.join(path, name)
        size_bytes, latest_mtime = _collect_path_info(full_path)
        if size_bytes <= 0:
            continue
        age_seconds = _path_age_seconds(full_path, latest_mtime)
        if age_seconds < cutoff_seconds:
            continue
        eligible_items.append({
            'path': full_path,
            'size_bytes': size_bytes,
            'age_seconds': age_seconds,
        })
        total_bytes += size_bytes
    return eligible_items, total_bytes


def _scan_safe_cleanup():
    results = []
    for profile in _build_safe_profiles():
        eligible_items, total_bytes = _scan_profile(profile)
        results.append({
            'label': profile['label'],
            'path': profile['path'],
            'age_days': profile['age_days'],
            'eligible_items': eligible_items,
            'eligible_bytes': total_bytes,
        })
    return results


def _scan_deep_cleanup_sizes():
    results = []
    for profile in _build_deep_profiles():
        size_bytes, _ = _collect_path_info(profile['path'])
        results.append({
            'label': profile['label'],
            'path': profile['path'],
            'size_bytes': size_bytes,
        })
    return results


def _delete_candidates(scan_results):
    deleted_bytes = 0
    deleted_count = 0
    skipped = []

    for result in scan_results:
        for item in result['eligible_items']:
            target_path = item['path']
            try:
                if not os.path.exists(target_path):
                    continue
                _delete_path(target_path)
                deleted_bytes += item['size_bytes']
                deleted_count += 1
            except Exception as ex:
                skipped.append((target_path, str(ex)))
                logger.debug('Skipped cleanup target {0}: {1}'.format(target_path, ex))

    return deleted_bytes, deleted_count, skipped


def _render_summary_lines(scan_results):
    lines = []
    for result in scan_results:
        if result['eligible_bytes'] <= 0:
            continue
        lines.append('- {0}: {1} older than {2} day(s)'.format(
            result['label'],
            _format_bytes(result['eligible_bytes']),
            result['age_days']
        ))
    return lines


def _render_deep_lines(deep_results):
    lines = []
    for result in deep_results:
        if result['size_bytes'] <= 0:
            continue
        lines.append('- {0}: {1}'.format(result['label'], _format_bytes(result['size_bytes'])))
    return lines


def _print_report(scan_results, deep_results, deleted_bytes, deleted_count, skipped):
    output.print_md('### {0}'.format(TOOL_TITLE))
    output.print_md('Deleted **{0}** across **{1}** item(s).'.format(
        _format_bytes(deleted_bytes),
        deleted_count
    ))

    if skipped:
        output.print_md('Skipped **{0}** locked or protected item(s).'.format(len(skipped)))

    output.print_md('#### Safe cleanup targets')
    for line in _render_summary_lines(scan_results):
        output.print_md(line)

    deep_lines = _render_deep_lines(deep_results)
    if deep_lines:
        output.print_md('#### Deeper cleanup after closing Revit')
        for line in deep_lines:
            output.print_md(line)


def main():
    scan_results = _scan_safe_cleanup()
    deep_results = _scan_deep_cleanup_sizes()

    total_eligible_bytes = sum(result['eligible_bytes'] for result in scan_results)
    summary_lines = _render_summary_lines(scan_results)
    deep_lines = _render_deep_lines(deep_results)

    if total_eligible_bytes <= 0:
        message_lines = [
            'No old Temp or Revit journal/cache items are currently eligible for safe live cleanup.',
        ]
        if deep_lines:
            message_lines.append('')
            message_lines.append('Close Revit first before clearing these larger caches:')
            message_lines.extend(deep_lines)
        forms.alert('\n'.join(message_lines), title=TOOL_TITLE)
        return

    prompt_lines = [
        'This button can safely free about {0} right now.'.format(_format_bytes(total_eligible_bytes)),
        '',
        'Targets:',
    ]
    prompt_lines.extend(summary_lines)

    if deep_lines:
        prompt_lines.append('')
        prompt_lines.append('Not touched while Revit is open:')
        prompt_lines.extend(deep_lines)

    prompt_lines.append('')
    prompt_lines.append('Proceed with cleanup now?')

    should_continue = forms.alert(
        '\n'.join(prompt_lines),
        title=TOOL_TITLE,
        yes=True,
        no=True
    )
    if not should_continue:
        return

    deleted_bytes, deleted_count, skipped = _delete_candidates(scan_results)
    _print_report(scan_results, deep_results, deleted_bytes, deleted_count, skipped)

    result_lines = [
        'Freed {0} from {1} item(s).'.format(_format_bytes(deleted_bytes), deleted_count),
    ]
    if skipped:
        result_lines.append('Skipped {0} locked or protected item(s).'.format(len(skipped)))
    if deleted_bytes <= 0 and skipped:
        result_lines.append('Most likely those files are still in use. Close Revit and retry for deeper cleanup.')
    if deep_lines:
        result_lines.append('')
        result_lines.append('For larger recovery, close Revit and then clear CollaborationCache if appropriate.')
    forms.alert('\n'.join(result_lines), title=TOOL_TITLE)


if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        logger.exception('Free Temp Space failed.')
        forms.alert('Cleanup failed safely:\n{0}'.format(ex), title=TOOL_TITLE)