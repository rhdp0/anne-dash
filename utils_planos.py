# -*- coding: utf-8 -*-
from typing import Optional

def strip_accents(s: str) -> str:
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def classify_planos(text) -> Optional[str]:
    """Classify free-text PLANOS into canonical categories for the board.

    Categories: 'JAYME', 'MISTO', 'PROPRIO', 'OUTROS'
    Rules:
      - If contains only 'JAYME' -> 'JAYME'
      - If contains 'JAYME' and any other plan -> 'MISTO'
      - If contains 'PROPRIO'/'PRï¿½"PRIO' or 'PARTICULAR' -> 'PROPRIO'
      - Else -> 'OUTROS'
    Supports separators: ',', ';', '/', '|', '+', '&', ' e '.
    """
    try:
        import pandas as pd  # type: ignore
    except Exception:  # pragma: no cover
        pd = None

    if text is None or (pd is not None and isinstance(text, float) and pd.isna(text)):
        return None
    raw = str(text).strip()
    if not raw:
        return None
    norm = strip_accents(raw).upper()
    import re
    parts = re.split(r"\s*(?:,|;|/|\||\+|&|\be\b)\s*", norm)
    parts = [p.strip() for p in parts if p and p.strip()]
    if not parts:
        return None
    has_jayme = any("JAYME" in p for p in parts)
    has_outro = any("JAYME" not in p for p in parts)
    has_proprio = any("PROPRIO" in p or "PARTICULAR" in p for p in parts)
    if has_proprio and not has_jayme and not has_outro:
        return "PROPRIO"
    if has_jayme and has_outro:
        return "MISTO"
    if has_jayme:
        return "JAYME"
    if has_proprio:
        return "PROPRIO"
    return "OUTROS"


