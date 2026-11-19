"""Compatibility entry point for :mod:`mddm.figure_numbering`."""

import importlib
import runpy
import sys

if __name__ == "__main__":
    runpy.run_module("mddm.figure_numbering", run_name="__main__")
else:
    sys.modules[__name__] = importlib.import_module("mddm.figure_numbering")
