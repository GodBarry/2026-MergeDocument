"""Compatibility entry point for :mod:`mddm.logging_config`."""

import importlib
import runpy
import sys

if __name__ == "__main__":
    runpy.run_module("mddm.logging_config", run_name="__main__")
else:
    sys.modules[__name__] = importlib.import_module("mddm.logging_config")
