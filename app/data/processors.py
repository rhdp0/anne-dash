"""Data processing utilities for consultório dashboards."""
from __future__ import annotations

import re
from typing import Any, Iterable

from unidecode import unidecode

import numpy as np
import pandas as pd


__all__ = [
    "normalize_column_name",
    "format_consultorio_label",
    "first_nonempty",
    "to_number",
    "normalize_plano_value",
]


def normalize_column_name(value: Any) -> str:
    """Normalize column names by removing accents and collapsing whitespace."""
    text = str(value).strip().lower()
    replacements = {
        "á": "a",
        "ã": "a",
        "â": "a",
        "é": "e",
        "ê": "e",
        "í": "i",
        "î": "i",
        "ó": "o",
        "õ": "o",
        "ô": "o",
        "ú": "u",
        "ü": "u",
        "ç": "c",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"\s+", " ", text)
    return text


def to_number(value: Any):
    """Convert arbitrary values into numeric types, preserving missing data."""
    if pd.isna(value):
        return np.nan
    text = str(value)
    text = re.sub(r"[^\d,.-]", "", text)
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text and "." not in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except Exception:  # pragma: no cover - defensive branch for unexpected inputs
        return pd.NA


def first_nonempty(series: Iterable[Any]) -> str:
    """Return the first non-empty textual value from an iterable."""
    for value in series:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def format_consultorio_label(name: Any) -> str:
    """Standardize consultório labels used across different sheets."""
    label = str(name).strip()
    label = re.sub(r"(?i)^produtividade\s*[:\-]*", "", label).strip()
    label = re.sub(r"(?i)consult[óo]rio", "Consultório", label)
    label = re.sub(r"\s+", " ", label).strip(" -_:")
    return label or str(name).strip()


def normalize_plano_value(value: Any, *, remove_accents: bool = True) -> str:
    """Normalize plano/convênio names for consistent grouping."""

    if pd.isna(value):
        return ""

    text = str(value).strip()
    if remove_accents:
        text = unidecode(text)
    return text.upper()
