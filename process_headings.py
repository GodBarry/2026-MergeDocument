"""Compatibility entry point for :mod:`mddm.process_headings`."""

import importlib
import runpy
import sys

if __name__ == "__main__":
    runpy.run_module("mddm.process_headings", run_name="__main__")
else:
    sys.modules[__name__] = importlib.import_module("mddm.process_headings")
