"""Data access layer for the dashboard application."""
from .facade import ConsultorioDataFacade
from .loader import (
    detect_header_and_parse,
    load_excel_workbook,
    load_medicos_from_workbook,
    load_produtividade_from_workbook,
    tidy_agenda_from_workbook,
)
from .processors import (
    first_nonempty,
    format_consultorio_label,
    normalize_column_name,
    normalize_plano_value,
    to_number,
)

__all__ = [
    "ConsultorioDataFacade",
    "detect_header_and_parse",
    "load_excel_workbook",
    "load_medicos_from_workbook",
    "load_produtividade_from_workbook",
    "tidy_agenda_from_workbook",
    "first_nonempty",
    "format_consultorio_label",
    "normalize_column_name",
    "normalize_plano_value",
    "to_number",
]
