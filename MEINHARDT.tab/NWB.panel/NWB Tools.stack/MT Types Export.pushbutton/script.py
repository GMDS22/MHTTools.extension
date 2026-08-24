# coding: utf8
from __future__ import print_function

import os


def _run_target_tool():
    target = os.path.normpath(
        os.path.join(os.path.dirname(__file__), "..", "..", "NWB MT12 Schedule Export.source", "script.py")
    )
    namespace = {"__name__": "__main__", "__file__": target}
    with open(target, "r") as stream:
        source = stream.read()
    exec(compile(source, target, "exec"), namespace)


if __name__ == "__main__":
    _run_target_tool()