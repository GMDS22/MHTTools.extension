import os
import tempfile
import webbrowser
from pathlib import Path

try:
    from html import escape as html_escape
except Exception:
    import cgi
    html_escape = cgi.escape

from pyrevit import script

# Get the extension directory
extension_dir = Path(__file__).parent.parent.parent.parent
doc_html_path = extension_dir / "MeinhardtTabTools.html"
doc_md_path = extension_dir / "MeinhardtTabTools.md"


def _open_in_browser(path_obj):
    """Open a local file in the default browser."""
    webbrowser.open(path_obj.resolve().as_uri(), new=2)

# Open the document
if doc_html_path.exists():
    try:
        _open_in_browser(doc_html_path)
    except Exception:
        os.system('start "" "{}"'.format(doc_html_path.resolve().as_uri()))
elif doc_md_path.exists():
    try:
        with doc_md_path.open('r') as md_file:
            md_content = md_file.read()

        html_content = (
            '<!doctype html><html><head><meta charset="utf-8">'
            '<title>Meinhardt Tools Description</title>'
            '<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:1000px;margin:24px auto;padding:0 16px;line-height:1.5;}pre{white-space:pre-wrap;word-wrap:break-word;}</style>'
            '</head><body><pre>{}</pre></body></html>'
        ).format(html_escape(md_content))

        temp_dir = Path(tempfile.gettempdir())
        temp_html_path = temp_dir / 'MHTTools_MeinhardtTabTools_preview.html'
        with temp_html_path.open('w') as html_file:
            html_file.write(html_content)

        _open_in_browser(temp_html_path)
    except Exception:
        _open_in_browser(doc_md_path)
else:
    script.get_logger().error("Tools description document not found.")