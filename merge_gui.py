"""Compatibility entry point for :mod:`mddm.merge_gui`."""

import importlib
import runpy
import sys

if __name__ == "__main__":
    runpy.run_module("mddm.merge_gui", run_name="__main__")
else:
    sys.modules[__name__] = importlib.import_module("mddm.merge_gui")
