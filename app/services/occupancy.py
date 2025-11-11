"""Services responsible for computing occupancy KPIs and derived tables."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, MutableMapping, Optional, Sequence

import pandas as pd


@dataclass
class OccupancyKPIs:
    """Bundle of high level occupancy indicators for the dashboard."""

    total_salas: int = 0
    total_slots: int = 0
    slots_ocupados: int = 0
    slots_livres: int = 0
    taxa_ocupacao: float = 0.0
    medicos_distintos: int = 0

    def as_dict(self) -> MutableMapping[str, float]:
        return {
            "total_salas": self.total_salas,
            "total_slots": self.total_slots,
            "slots_ocupados": self.slots_ocupados,
            "slots_livres": self.slots_livres,
            "taxa_ocupacao": self.taxa_ocupacao,
            "medicos_distintos": self.medicos_distintos,
        }


class OccupancyAnalyzer:
    """Compute occupancy KPI metrics, rankings and helper tables."""

    def __init__(
        self,
        base_df: Optional[pd.DataFrame] = None,
        filtered_df: Optional[pd.DataFrame] = None,
        *,
        selected_salas: Optional[Iterable[str]] = None,
        selected_dias: Optional[Iterable[str]] = None,
        selected_turnos: Optional[Iterable[str]] = None,
        selected_medicos: Optional[Iterable[str]] = None,
        ranking_df: Optional[pd.DataFrame] = None,
    ) -> None:
        self.base_df = base_df.copy() if base_df is not None else pd.DataFrame()
        self.filtered_df = filtered_df.copy() if filtered_df is not None else pd.DataFrame()
        self.selected_salas = list(selected_salas or [])
        self.selected_dias = list(selected_dias or [])
        self.selected_turnos = list(selected_turnos or [])
        self.selected_medicos = list(selected_medicos or [])
        self.ranking_df = ranking_df.copy() if ranking_df is not None else pd.DataFrame()

    def get_kpi_summary(self) -> OccupancyKPIs:
        """Return the main occupancy indicators for the selected filters."""

        if self.base_df.empty or "Ocupado" not in self.base_df:
            return OccupancyKPIs(total_salas=len(self.selected_salas))

        total_slots = len(self.base_df)
        occupied = int(self.base_df["Ocupado"].fillna(False).sum())
        slots_livres = max(total_slots - occupied, 0)
        taxa_ocupacao = (occupied / total_slots * 100) if total_slots else 0.0
        medicos_distintos = 0
        if "Médico" in self.base_df.columns:
            medicos_distintos = self.base_df.loc[
                self.base_df["Ocupado"].fillna(False), "Médico"
            ].nunique()

        total_salas = (
            len(set(self.selected_salas))
            if self.selected_salas
            else self.base_df.get("Sala", pd.Series(dtype=object)).nunique()
        )

        return OccupancyKPIs(
            total_salas=total_salas,
            total_slots=total_slots,
            slots_ocupados=occupied,
            slots_livres=slots_livres,
            taxa_ocupacao=taxa_ocupacao,
            medicos_distintos=medicos_distintos,
        )

    def build_summary_metadata(self) -> MutableMapping[str, object]:
        """Produce the dictionary used on the executive summary card."""

        kpis = self.get_kpi_summary()
        dias = ", ".join(self.selected_dias) if self.selected_dias else "Todos"
        turnos = ", ".join(self.selected_turnos) if self.selected_turnos else "Todos"

        summary = {
            "Consultórios selecionados": kpis.total_salas,
            "Slots analisados": kpis.total_slots,
            "Slots livres": kpis.slots_livres,
            "Slots ocupados": kpis.slots_ocupados,
            "Taxa de ocupação": f"{kpis.taxa_ocupacao:.1f}%",
            "Médicos distintos": kpis.medicos_distintos,
            "Dias filtrados": dias,
            "Turnos filtrados": turnos,
        }

        if self.selected_medicos:
            summary["Médicos no filtro"] = len(self.selected_medicos)

        if not self.ranking_df.empty and "Receita" in self.ranking_df.columns:
            total_receita = self.ranking_df["Receita"].sum()
            if total_receita > 0:
                summary["Receita total (produtividade)"] = total_receita

        return summary

    def build_timeseries(self, group_by: Sequence[str]) -> pd.DataFrame:
        """Aggregate occupancy percentage grouped by the given columns."""

        if not group_by:
            raise ValueError("group_by must contain at least one column")

        if self.base_df.empty or "Ocupado" not in self.base_df:
            columns = list(group_by) + ["Taxa de Ocupação (%)", "Slots Ocupados", "Total Slots"]
            return pd.DataFrame(columns=columns)

        group_cols = list(group_by)
        grouped = (
            self.base_df.groupby(group_cols)["Ocupado"].agg(["sum", "count"]).reset_index()
        )
        grouped.rename(columns={"sum": "Slots Ocupados", "count": "Total Slots"}, inplace=True)
        grouped["Taxa de Ocupação (%)"] = (
            grouped["Slots Ocupados"] / grouped["Total Slots"] * 100
        ).fillna(0).round(1)
        grouped["Slots Ocupados"] = grouped["Slots Ocupados"].astype(int)
        grouped["Total Slots"] = grouped["Total Slots"].astype(int)
        return grouped

    def top_medicos_por_turnos(self, limit: int = 15) -> pd.DataFrame:
        """Return the top doctors by number of occupied slots."""

        if self.filtered_df.empty or "Ocupado" not in self.filtered_df:
            return pd.DataFrame(columns=["Médico", "Turnos Utilizados"])

        base = self.filtered_df[self.filtered_df["Ocupado"].fillna(False)]
        if base.empty or "Médico" not in base:
            return pd.DataFrame(columns=["Médico", "Turnos Utilizados"])

        top = (
            base.groupby("Médico")
            .size()
            .reset_index(name="Turnos Utilizados")
            .sort_values("Turnos Utilizados", ascending=False)
        )
        return top.head(limit)

    @staticmethod
    def compute_basic_metrics(df: pd.DataFrame) -> OccupancyKPIs:
        """Compute the same KPI bundle for an arbitrary dataframe."""

        analyzer = OccupancyAnalyzer(base_df=df)
        return analyzer.get_kpi_summary()
