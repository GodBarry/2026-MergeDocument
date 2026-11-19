"""Compatibility entry point for :mod:`mddm.convert_docx_formulas`."""

import importlib
import runpy
import sys

if __name__ == "__main__":
    runpy.run_module("mddm.convert_docx_formulas", run_name="__main__")
else:
    sys.modules[__name__] = importlib.import_module("mddm.convert_docx_formulas")
