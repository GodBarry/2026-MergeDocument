"""Compatibility entry point for :mod:`mddm.mathtype_converter`."""

import importlib
import runpy
import sys

if __name__ == "__main__":
    runpy.run_module("mddm.mathtype_converter", run_name="__main__")
else:
    sys.modules[__name__] = importlib.import_module("mddm.mathtype_converter")
