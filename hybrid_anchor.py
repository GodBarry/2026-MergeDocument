"""Compatibility entry point for :mod:`mddm.hybrid_anchor`."""

import importlib
import runpy
import sys

if __name__ == "__main__":
    runpy.run_module("mddm.hybrid_anchor", run_name="__main__")
else:
    sys.modules[__name__] = importlib.import_module("mddm.hybrid_anchor")
