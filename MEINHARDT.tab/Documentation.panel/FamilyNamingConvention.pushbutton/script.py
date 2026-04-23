import subprocess
import sys
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
doc_md_path = extension_dir / "FamilyNamingConvention.md"


def _open_in_browser(path_obj):
    uri = path_obj.resolve().as_uri()
    if sys.platform.startswith("win"):
        try:
            subprocess.Popen(["rundll32", "url.dll,FileProtocolHandler", uri], shell=False)
            return
        except Exception:
            pass
        try:
            subprocess.Popen('start "" "{}"'.format(uri), shell=True)
            return
        except Exception:
            pass
    webbrowser.open(uri, new=2)


# Open the document
if doc_md_path.exists():
    # Convert markdown to HTML so it renders properly in a browser
    try:
        with doc_md_path.open("r") as f:
            md_content = f.read()
        html_content = (
            '<!doctype html><html><head><meta charset="utf-8">'
            '<title>Family Naming Convention</title>'
            '<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:1000px;margin:24px auto;padding:0 16px;line-height:1.5;}pre{white-space:pre-wrap;word-wrap:break-word;}</style>'
            '</head><body><pre>{}</pre></body></html>'
        ).format(html_escape(md_content))
        temp_path = Path(tempfile.gettempdir()) / "MHTTools_FamilyNamingConvention.html"
        with temp_path.open("w") as f:
            f.write(html_content)
        _open_in_browser(temp_path)
    except Exception:
        _open_in_browser(doc_md_path)
else:
    script.get_logger().error("Family naming convention document not found.")