"""Workbook Change Assurance Benchmark (WCAB)."""

from .build import CASE_IDS, build_all
from .validate import validate_all

__all__ = ["CASE_IDS", "build_all", "validate_all"]

__version__ = "0.1.1"
