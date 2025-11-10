"""High-level interface for accessing consultório datasets."""
from __future__ import annotations

from typing import Any, Dict, Iterable, List, Mapping, Sequence

import pandas as pd

from .loader import (
    load_excel_workbook,
    load_medicos_from_workbook,
    load_produtividade_from_workbook,
    tidy_agenda_from_workbook,
)


class ConsultorioDataFacade:
    """Facade that centralises data access patterns used by the dashboard."""

    def load_workbook(self, source: Any) -> pd.ExcelFile:
        """Load an Excel workbook from any supported source."""
        if isinstance(source, pd.ExcelFile):
            return source
        return load_excel_workbook(source)

    def load_dataset(self, source: Any) -> Dict[str, pd.DataFrame]:
        """Return the key datasets extracted from the Excel workbook."""
        workbook = self.load_workbook(source)
        agenda = tidy_agenda_from_workbook(workbook)
        produtividade = load_produtividade_from_workbook(workbook)
        medicos = load_medicos_from_workbook(workbook)
        return {
            "agenda": agenda,
            "produtividade": produtividade,
            "medicos": medicos,
        }

    def filter_by_date(
        self,
        df: pd.DataFrame,
        *,
        start=None,
        end=None,
        date_column: str = "Data",
        allowed_values: Iterable[Any] | None = None,
    ) -> pd.DataFrame:
        """Filter a DataFrame by a date column or by explicit values."""
        if df.empty:
            return df.copy()
        if date_column not in df.columns:
            return df.copy()

        series = df[date_column]
        mask = pd.Series(True, index=df.index)

        if allowed_values is not None:
            allowed = [str(value) for value in allowed_values]
            mask &= series.astype(str).isin(allowed)
            return df[mask].copy()

        series_dt = pd.to_datetime(series, errors="coerce")
        if start is not None:
            mask &= series_dt >= pd.to_datetime(start)
        if end is not None:
            mask &= series_dt <= pd.to_datetime(end)
        return df[mask].copy()

    def group_metrics(
        self,
        df: pd.DataFrame,
        group_by: Sequence[str],
        aggregations: Mapping[str, Any],
    ) -> pd.DataFrame:
        """Aggregate metrics using a consistent group-by strategy."""
        if df.empty:
            ordered_columns: List[str] = list(dict.fromkeys([*group_by, *aggregations.keys()]))
            return pd.DataFrame(columns=ordered_columns)
        return df.groupby(list(group_by), as_index=False).agg(aggregations)


__all__ = ["ConsultorioDataFacade"]
