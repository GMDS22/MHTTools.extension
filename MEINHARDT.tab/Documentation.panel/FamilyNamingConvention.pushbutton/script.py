import webbrowser
from pathlib import Path

from pyrevit import script

# Get the extension directory
extension_dir = Path(__file__).parent.parent.parent.parent
doc_path = extension_dir / "FamilyNamingConvention.md"

# Open the document
if doc_path.exists():
    webbrowser.open(doc_path.resolve().as_uri())
else:
    script.get_logger().error("Family naming convention document not found.")