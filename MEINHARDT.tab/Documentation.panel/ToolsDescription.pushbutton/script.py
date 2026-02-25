import os
from pathlib import Path

from pyrevit import script

# Get the extension directory
extension_dir = Path(__file__).parent.parent.parent.parent
doc_path = extension_dir / "MeinhardtTabTools.md"

# Open the document
if doc_path.exists():
    os.startfile(str(doc_path))
else:
    script.get_logger().error("Tools description document not found.")