"""Compatibility entry point for :mod:`mddm.multi_copy_merger`."""

import importlib
import runpy
import sys

if __name__ == "__main__":
    runpy.run_module("mddm.multi_copy_merger", run_name="__main__")
else:
    sys.modules[__name__] = importlib.import_module("mddm.multi_copy_merger")
