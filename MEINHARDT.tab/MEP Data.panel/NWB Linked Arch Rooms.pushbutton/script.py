# coding: utf8
from __future__ import print_function

import os

from pyrevit import forms


SOURCE_SCRIPT = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "Data Manage.pulldown",
        "NWB Linked Arch Rooms Export.pushbutton",
        "script.py",
    )
)


def run():
    if not os.path.exists(SOURCE_SCRIPT):
        forms.alert("Source exporter script was not found:\n{0}".format(SOURCE_SCRIPT), title="Arch Rooms Export")
        return

    scope = {"__file__": SOURCE_SCRIPT, "__name__": "__main__"}
    with open(SOURCE_SCRIPT, "rb") as stream:
        code = compile(stream.read(), SOURCE_SCRIPT, "exec")
    exec(code, scope, scope)


if __name__ == "__main__":
    run()