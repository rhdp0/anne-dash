"""PDF builder utilities for exporting the dashboard report."""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from io import BytesIO
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
from fpdf import FPDF


PDF_PRIMARY_COLOR = (27, 59, 95)
PDF_ACCENT_COLOR = (76, 137, 198)
PDF_TEXT_COLOR = (20, 33, 61)
PDF_MUTED_COLOR = (95, 108, 133)
PDF_SOFT_BACKGROUND = (244, 247, 251)
PDF_WHITE = (255, 255, 255)

PDF_MARGIN = 18
PDF_SECTION_GAP = 8
PDF_CARD_GAP = 6
PDF_CARD_HEIGHT = 26
PDF_CARD_PADDING = 4

PDF_TITLE_FONT = ("Helvetica", "B", 28)
PDF_SUBTITLE_FONT = ("Helvetica", "", 12)
PDF_SECTION_TITLE_FONT = ("Helvetica", "B", 16)
PDF_SECTION_SUBTITLE_FONT = ("Helvetica", "", 11)
PDF_BODY_FONT = ("Helvetica", "", 10)
PDF_SUBSECTION_FONT = ("Helvetica", "B", 12)
PDF_KPI_VALUE_FONT = ("Helvetica", "B", 16)
PDF_KPI_LABEL_FONT = ("Helvetica", "", 9)


class DashboardPDF(FPDF):
    """Styled PDF with consistent footer for the dashboard report."""

    def __init__(self, data_source: str, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.data_source = data_source
        self.generated_at = datetime.now()

    def footer(self):  # pragma: no cover - UI rendering helper
        self.set_y(-15)
        family, style, _ = PDF_BODY_FONT
        self.set_font(family, style, 8)
        self.set_text_color(*PDF_MUTED_COLOR)
        footer_text = (
            f"Fonte: {self.data_source} | "
            f"Gerado em {self.generated_at.strftime('%d/%m/%Y %H:%M')} | "
            f"Página {self.page_no()}"
        )
        self.multi_cell(0, 4, _sanitize_pdf_text(footer_text), align="C")


def _sanitize_pdf_text(text: str) -> str:
    """Remove acentuação incompatível e caracteres fora do conjunto latin-1."""
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)

    normalized = unicodedata.normalize("NFKD", text)
    cleaned = "".join(ch for ch in normalized if not unicodedata.combining(ch))

    substitutions = {
        "•": "-",
        "–": "-",
        "—": "-",
        "“": '"',
        "”": '"',
        "’": "'",
        "´": "'",
        "`": "'",
        "ª": "a",
        "º": "o",
    }
    for old, new in substitutions.items():
        cleaned = cleaned.replace(old, new)

    cleaned = cleaned.replace("\xa0", " ")

    lines: List[str] = []
    for line in cleaned.splitlines():
        collapsed = re.sub(r"\s+", " ", line).strip()
        lines.append(collapsed)
    cleaned = "\n".join(lines).strip()

    cleaned = cleaned.encode("latin-1", "ignore").decode("latin-1")
    return cleaned


class DashboardPDFBuilder:
    """High-level helper to assemble the dashboard PDF report."""

    def __init__(
        self,
        *,
        data_source: str,
        summary_metrics: Optional[Dict[str, object]] = None,
        ranking_df: Optional[pd.DataFrame] = None,
        med_df: Optional[pd.DataFrame] = None,
        agenda_df: Optional[pd.DataFrame] = None,
        ranking_limits: Optional[Dict[str, int]] = None,
    ) -> None:
        self.data_source = data_source or "Origem não informada"
        self.summary_metrics = summary_metrics or {}
        self.ranking_df = ranking_df.copy() if ranking_df is not None else pd.DataFrame()
        self.med_df = med_df.copy() if med_df is not None else pd.DataFrame()
        self.agenda_df = agenda_df.copy() if agenda_df is not None else pd.DataFrame()
        self.ranking_limits = ranking_limits or {}

        self.pdf: Optional[DashboardPDF] = None
        self.effective_width: float = 0.0
        self.sections_index: List[Tuple[str, int]] = []

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def build(self) -> bytes:
        """Generate the PDF report and return the resulting bytes."""

        self.pdf = DashboardPDF(data_source=_sanitize_pdf_text(str(self.data_source)))
        self.pdf.set_margins(PDF_MARGIN, PDF_MARGIN, PDF_MARGIN)
        self.pdf.set_auto_page_break(auto=True, margin=PDF_MARGIN)
        self.pdf.alias_nb_pages()

        self.effective_width = self.pdf.w - self.pdf.l_margin - self.pdf.r_margin
        self.sections_index = []

        self._draw_cover_page()

        self.pdf.add_page()
        self._set_body_font()

        self._render_summary_section()
        self._render_ranking_section()
        self._render_med_info_section()
        self._render_agenda_section()
        self._render_toc()

        output = self.pdf.output(dest="S")
        if isinstance(output, str):
            output_bytes = output.encode("latin-1")
        else:
            output_bytes = bytes(output)

        buffer = BytesIO()
        buffer.write(output_bytes)
        buffer.seek(0)
        return buffer.getvalue()

    # ------------------------------------------------------------------
    # Rendering helpers
    # ------------------------------------------------------------------
    def _set_body_font(self) -> None:
        family, style, size = PDF_BODY_FONT
        self.pdf.set_font(family, style, size)
        self.pdf.set_text_color(*PDF_TEXT_COLOR)

    def _draw_cover_page(self) -> None:
        pdf = self.pdf
        pdf.add_page()

        pdf.set_fill_color(*PDF_PRIMARY_COLOR)
        pdf.rect(0, 0, pdf.w, pdf.h * 0.45, "F")
        pdf.set_fill_color(*PDF_SOFT_BACKGROUND)
        pdf.rect(0, pdf.h * 0.45, pdf.w, pdf.h * 0.55, "F")

        pdf.set_xy(pdf.l_margin, 40)
        family, style, size = PDF_TITLE_FONT
        pdf.set_font(family, style, size)
        pdf.set_text_color(255, 255, 255)
        pdf.multi_cell(
            self.effective_width,
            12,
            _sanitize_pdf_text("Relatório Completo"),
        )

        family, style, size = PDF_SUBTITLE_FONT
        pdf.set_font(family, style, size)
        pdf.multi_cell(
            self.effective_width,
            8,
            _sanitize_pdf_text("Dashboard de Ocupação dos Consultórios"),
        )

        block_x = pdf.l_margin
        block_y = pdf.get_y() + 10
        block_w = self.effective_width
        block_h = 60
        pdf.set_fill_color(*PDF_WHITE)
        pdf.set_draw_color(*PDF_ACCENT_COLOR)
        pdf.set_line_width(0.4)
        pdf.rect(block_x, block_y, block_w, block_h, "DF")

        pdf.set_xy(block_x + PDF_CARD_PADDING, block_y + PDF_CARD_PADDING)
        family, style, size = PDF_SECTION_SUBTITLE_FONT
        pdf.set_font(family, style, size)
        pdf.set_text_color(*PDF_PRIMARY_COLOR)
        pdf.cell(0, 6, _sanitize_pdf_text("Sobre este relatório"), ln=1)

        self._set_body_font()
        pdf.set_x(block_x + PDF_CARD_PADDING)
        about_lines = [
            "Panorama executivo com indicadores de produtividade e agenda.",
            f"Fonte dos dados: {self.data_source}.",
            "Geração automática via Dashboard Consultórios.",
        ]
        for line in about_lines:
            pdf.multi_cell(
                block_w - 2 * PDF_CARD_PADDING,
                6,
                _sanitize_pdf_text(line),
            )
        pdf.set_y(block_y + block_h + 12)

        self._set_body_font()

    def _draw_section_header(self, title: str, subtitle: Optional[str] = None) -> None:
        pdf = self.pdf
        if pdf.get_y() < PDF_MARGIN:
            pdf.set_y(PDF_MARGIN)
        start_y = pdf.get_y()
        pdf.set_fill_color(*PDF_ACCENT_COLOR)
        pdf.rect(pdf.l_margin, start_y, 4, 12, "F")
        pdf.set_xy(pdf.l_margin + 8, start_y)

        family, style, size = PDF_SECTION_TITLE_FONT
        pdf.set_font(family, style, size)
        pdf.set_text_color(*PDF_PRIMARY_COLOR)
        pdf.cell(0, 10, _sanitize_pdf_text(title), ln=1)

        if subtitle:
            family, style, size = PDF_SECTION_SUBTITLE_FONT
            pdf.set_font(family, style, size)
            pdf.set_text_color(*PDF_MUTED_COLOR)
            pdf.multi_cell(0, 6, _sanitize_pdf_text(subtitle))

        pdf.ln(2)
        self._set_body_font()

    def _draw_subsection_header(self, title: str) -> None:
        family, style, size = PDF_SUBSECTION_FONT
        self.pdf.set_font(family, style, size)
        self.pdf.set_text_color(*PDF_ACCENT_COLOR)
        self.pdf.cell(0, 8, _sanitize_pdf_text(title), ln=1)
        self._set_body_font()

    def _write_body_line(self, text: str, height: float = 6) -> None:
        sanitized = _sanitize_pdf_text(text)
        if not sanitized:
            self.pdf.ln(height)
            return

        self.pdf.set_x(self.pdf.l_margin)
        width = self.effective_width or (self.pdf.w - self.pdf.l_margin - self.pdf.r_margin)
        if width <= 0:
            self.pdf.ln(height)
            return

        self.pdf.multi_cell(width, height, sanitized)

    def _draw_kpi_cards(self, metrics: Dict[str, object]) -> None:
        if not metrics:
            return

        pdf = self.pdf
        cards_per_row = 2
        card_width = (self.effective_width - PDF_CARD_GAP * (cards_per_row - 1)) / cards_per_row
        items = list(metrics.items())
        for idx, (label, value) in enumerate(items):
            if pdf.get_y() + PDF_CARD_HEIGHT > pdf.page_break_trigger:
                pdf.add_page()
                self._set_body_font()

            column = idx % cards_per_row
            x = pdf.l_margin + column * (card_width + PDF_CARD_GAP)
            y = pdf.get_y()

            pdf.set_fill_color(*PDF_SOFT_BACKGROUND)
            pdf.set_draw_color(*PDF_ACCENT_COLOR)
            pdf.set_line_width(0.3)
            pdf.rect(x, y, card_width, PDF_CARD_HEIGHT, "DF")

            inner_x = x + PDF_CARD_PADDING
            inner_width = card_width - 2 * PDF_CARD_PADDING

            family, style, size = PDF_KPI_LABEL_FONT
            pdf.set_font(family, style, size)
            pdf.set_text_color(*PDF_MUTED_COLOR)
            pdf.set_xy(inner_x, y + 4)
            pdf.multi_cell(inner_width, 5, _sanitize_pdf_text(str(label)))

            family, style, size = PDF_KPI_VALUE_FONT
            pdf.set_font(family, style, size)
            pdf.set_text_color(*PDF_PRIMARY_COLOR)
            pdf.set_xy(inner_x, y + PDF_CARD_HEIGHT / 2)
            pdf.cell(inner_width, 6, _sanitize_pdf_text(str(value)))

            if column == cards_per_row - 1 or idx == len(items) - 1:
                pdf.set_y(y + PDF_CARD_HEIGHT + PDF_CARD_GAP)
            else:
                pdf.set_xy(x + card_width + PDF_CARD_GAP, y)

        pdf.ln(2)
        self._set_body_font()

    def _draw_table(
        self,
        columns: List[Tuple[str, float]],
        data: Iterable[Dict[str, object]] | pd.DataFrame,
        *,
        header_height: float = 7,
        line_height: float = 6,
    ) -> None:
        if not columns:
            return

        pdf = self.pdf
        if pdf is None:
            return

        if isinstance(data, pd.DataFrame):
            records: List[Dict[str, object]] = data.to_dict(orient="records")
        else:
            records = list(data)

        if not records:
            return

        total_spec = sum(width for _, width in columns) or float(len(columns))
        effective_width = self.effective_width or (
            pdf.w - pdf.l_margin - pdf.r_margin
        )
        scale = effective_width / total_spec
        col_widths = [width * scale for _, width in columns]

        width_sum = sum(col_widths)
        if width_sum > effective_width:
            shrink_factor = effective_width / width_sum
            col_widths = [w * shrink_factor for w in col_widths]
        elif width_sum < effective_width and col_widths:
            col_widths[-1] += effective_width - width_sum

        header_labels = [label for label, _ in columns]

        def _draw_header() -> None:
            if pdf.get_y() + header_height > pdf.page_break_trigger:
                pdf.add_page()
                self._set_body_font()

            pdf.set_x(pdf.l_margin)
            pdf.set_fill_color(*PDF_SOFT_BACKGROUND)
            pdf.set_draw_color(*PDF_ACCENT_COLOR)
            pdf.set_line_width(0.2)
            family, _, size = PDF_BODY_FONT
            pdf.set_font(family, "B", size)
            pdf.set_text_color(*PDF_PRIMARY_COLOR)

            for label, width in zip(header_labels, col_widths):
                pdf.multi_cell(
                    width,
                    header_height,
                    _sanitize_pdf_text(str(label)),
                    border=1,
                    align="L",
                    fill=True,
                    new_x="RIGHT",
                    new_y="TOP",
                    max_line_height=header_height,
                )

            pdf.ln(header_height)
            self._set_body_font()
            pdf.set_text_color(*PDF_TEXT_COLOR)

        def _prepare_cell_text(value: object) -> str:
            if value is None:
                return "—"
            try:
                if isinstance(value, float) and pd.isna(value):
                    return "—"
            except TypeError:
                pass
            return _sanitize_pdf_text(str(value)) or "—"

        _draw_header()

        for record in records:
            row_texts = [_prepare_cell_text(record.get(label)) for label in header_labels]

            line_counts: List[int] = []
            for text, width in zip(row_texts, col_widths):
                lines = pdf.multi_cell(
                    width,
                    line_height,
                    text,
                    split_only=True,
                )
                line_counts.append(max(1, len(lines)))

            row_height = max(line_counts) * line_height
            if pdf.get_y() + row_height > pdf.page_break_trigger:
                pdf.add_page()
                self._set_body_font()
                _draw_header()

            start_y = pdf.get_y()
            x = pdf.l_margin
            for text, width in zip(row_texts, col_widths):
                pdf.set_xy(x, start_y)
                pdf.multi_cell(
                    width,
                    line_height,
                    text,
                    border=1,
                    align="L",
                    fill=False,
                    new_x="RIGHT",
                    new_y="TOP",
                    max_line_height=line_height,
                )
                x += width

            pdf.set_xy(pdf.l_margin, start_y + row_height)

        pdf.ln(2)
        self._set_body_font()

    def _safe_int(self, value):
        if value is None:
            return None
        if isinstance(value, str) and not value.strip():
            return None
        try:
            if pd.isna(value):
                return None
        except TypeError:
            pass
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return None

    # ------------------------------------------------------------------
    # Section rendering
    # ------------------------------------------------------------------
    def _start_section(self, title: str, subtitle: Optional[str] = None) -> None:
        if self.pdf.get_y() > PDF_MARGIN:
            self.pdf.ln(PDF_SECTION_GAP)
        self.sections_index.append((title, self.pdf.page_no()))
        self._draw_section_header(title, subtitle)

    def _render_summary_section(self) -> None:
        if not self.summary_metrics:
            return
        self._start_section(
            "Resumo Executivo",
            "Indicadores principais para acompanhamento rápido do desempenho.",
        )
        self._draw_kpi_cards(self.summary_metrics)

    def _render_ranking_section(self) -> None:
        ranking_df = self.ranking_df
        if ranking_df is None or ranking_df.empty:
            return

        limits_cfg = self.ranking_limits or {}

        def _get_limit(key: str, default: int = 10) -> int:
            try:
                value = int(limits_cfg.get(key, default))
                return value if value > 0 else default
            except (TypeError, ValueError):
                return default

        limit_total = _get_limit("total", 10)
        limit_exames = _get_limit("exames", limit_total)
        limit_cirurgias = _get_limit("cirurgias", limit_total)
        limit_receita = _get_limit("receita", limit_total)

        def _prepare_ranking(df_source: pd.DataFrame, order: Iterable[Tuple[str, bool]]) -> pd.DataFrame:
            sort_cols: List[str] = []
            ascending: List[bool] = []
            for col, asc in order:
                if col in df_source.columns:
                    sort_cols.append(col)
                    ascending.append(asc)
            if sort_cols:
                sorted_df = df_source.sort_values(sort_cols, ascending=ascending)
            else:
                sorted_df = df_source.copy()
            sorted_df = sorted_df.reset_index(drop=True)
            sorted_df.insert(0, "Rank", range(1, len(sorted_df) + 1))
            return sorted_df

        ranking_total_pdf = _prepare_ranking(
            ranking_df,
            [
                ("Total Procedimentos", False),
                ("Cirurgias Solicitadas", False),
                ("Exames Solicitados", False),
                ("Profissional", True),
                ("Consultório", True),
            ],
        ).head(min(limit_total, len(ranking_df)))

        ranking_exames_pdf = _prepare_ranking(
            ranking_df,
            [
                ("Exames Solicitados", False),
                ("Cirurgias Solicitadas", False),
                ("Total Procedimentos", False),
                ("Profissional", True),
                ("Consultório", True),
            ],
        ).head(min(limit_exames, len(ranking_df)))

        ranking_cirurgias_pdf = _prepare_ranking(
            ranking_df,
            [
                ("Cirurgias Solicitadas", False),
                ("Exames Solicitados", False),
                ("Total Procedimentos", False),
                ("Profissional", True),
                ("Consultório", True),
            ],
        ).head(min(limit_cirurgias, len(ranking_df)))

        ranking_receita_pdf = (
            _prepare_ranking(
                ranking_df,
                [
                    ("Receita", False),
                    ("Total Procedimentos", False),
                    ("Profissional", True),
                    ("Consultório", True),
                ],
            ).head(min(limit_receita, len(ranking_df)))
            if "Receita" in ranking_df.columns
            else pd.DataFrame()
        )

        ranking_table_columns: List[Tuple[str, float]] = [
            ("Rank", 1.0),
            ("Profissional", 2.8),
            ("Consultório", 2.0),
            ("Especialidade", 2.2),
            ("Totais", 3.0),
            ("Receita", 1.8),
        ]

        def _build_ranking_table(dataset: pd.DataFrame) -> pd.DataFrame:
            if dataset.empty:
                return pd.DataFrame(columns=[label for label, _ in ranking_table_columns])

            registros: List[Dict[str, object]] = []
            for _, row in dataset.iterrows():
                rank_val = self._safe_int(row.get("Rank"))
                prof = row.get("Profissional") or "Não informado"
                consultorio = row.get("Consultório") or "Não informado"
                especialidade = row.get("Especialidade") or "Não informada"
                total = self._safe_int(row.get("Total Procedimentos"))
                exames = self._safe_int(row.get("Exames Solicitados"))
                cirurgias = self._safe_int(row.get("Cirurgias Solicitadas"))
                receita = row.get("Receita") if "Receita" in row else None

                totais_partes: List[str] = []
                if total is not None:
                    totais_partes.append(f"Total: {total}")
                if exames is not None:
                    totais_partes.append(f"Exames: {exames}")
                if cirurgias is not None:
                    totais_partes.append(f"Cirurgias: {cirurgias}")

                registros.append(
                    {
                        "Rank": str(rank_val) if rank_val is not None else "—",
                        "Profissional": prof,
                        "Consultório": consultorio,
                        "Especialidade": especialidade,
                        "Totais": " | ".join(totais_partes) if totais_partes else "—",
                        "Receita": self._format_currency_value(receita),
                    }
                )

            return pd.DataFrame(registros, columns=[label for label, _ in ranking_table_columns])

        def _write_ranking_section(title: str, dataset: pd.DataFrame, limit_used: int) -> None:
            if dataset.empty:
                return
            tabela = _build_ranking_table(dataset)
            if tabela.empty:
                return
            self._draw_subsection_header(f"{title} (top {limit_used})")
            self._draw_table(ranking_table_columns, tabela)
            self.pdf.ln(PDF_SECTION_GAP / 2)

        self._start_section(
            "Rankings de Produtividade",
            "Análise dos profissionais com maior volume de solicitações.",
        )
        _write_ranking_section(
            "Top profissionais por produtividade",
            ranking_total_pdf,
            min(limit_total, len(ranking_df)),
        )
        _write_ranking_section(
            "Top solicitantes de exames",
            ranking_exames_pdf,
            min(limit_exames, len(ranking_df)),
        )
        _write_ranking_section(
            "Top solicitantes de cirurgias",
            ranking_cirurgias_pdf,
            min(limit_cirurgias, len(ranking_df)),
        )
        if "Receita" in ranking_df.columns:
            _write_ranking_section(
                "Top profissionais por receita",
                ranking_receita_pdf,
                min(limit_receita, len(ranking_df)),
            )

    def _render_med_info_section(self) -> None:
        med_df = self.med_df
        if med_df is None or med_df.empty:
            return

        med_pdf = med_df.copy()
        if "Valor Aluguel" in med_pdf.columns:
            med_pdf["Valor Aluguel"] = pd.to_numeric(
                med_pdf["Valor Aluguel"], errors="coerce"
            )

        self._start_section(
            "Planos, Aluguel e Profissionais",
            "Composição de planos, profissionais ativos e valores praticados.",
        )

        total_profissionais = (
            med_pdf["Médico"].nunique() if "Médico" in med_pdf.columns else len(med_pdf)
        )
        self._write_body_line(f"Profissionais analisados: {total_profissionais}")

        if "Planos" in med_pdf.columns:
            self._draw_subsection_header("Distribuição por planos")
            planos = med_pdf.copy()
            planos["Planos"] = planos["Planos"].fillna("Nao informado").astype(str).str.strip()
            if "Médico" in planos.columns:
                planos_grouped = (
                    planos.groupby("Planos", observed=False)["Médico"].nunique().reset_index(name="Profissionais")
                )
            else:
                planos_grouped = planos["Planos"].value_counts().reset_index()
                planos_grouped.columns = ["Planos", "Profissionais"]
            planos_grouped = planos_grouped.sort_values("Profissionais", ascending=False)
            for _, row in planos_grouped.head(5).iterrows():
                plano_nome = row.get("Planos", "Nao informado")
                qtd = self._safe_int(row.get("Profissionais", 0)) or 0
                self._write_body_line(f"- {plano_nome}: {qtd} profissionais", height=5)

        if "Consultório" in med_pdf.columns:
            self._draw_subsection_header("Totais por consultório")
            consult = med_pdf.copy()
            consult["Consultório"] = consult["Consultório"].fillna("Nao informado").astype(str).str.strip()
            consult_totais = consult.groupby("Consultório", observed=False)
            consult_resumo = consult_totais["Médico"].nunique().reset_index(name="Profissionais")
            if "Valor Aluguel" in consult.columns:
                consult_resumo["Valor total aluguel"] = consult_totais["Valor Aluguel"].sum(
                    min_count=1
                )
            if "Valor total aluguel" in consult_resumo.columns:
                consult_resumo = consult_resumo.sort_values(
                    ["Valor total aluguel", "Profissionais"],
                    ascending=[False, False],
                    na_position="last",
                )
            else:
                consult_resumo = consult_resumo.sort_values(
                    "Profissionais", ascending=False, na_position="last"
                )
            for _, row in consult_resumo.head(5).iterrows():
                texto = (
                    f"- {row.get('Consultório', 'Nao informado')}: "
                    f"{int(row.get('Profissionais', 0))} profissionais"
                )
                if (
                    "Valor total aluguel" in consult_resumo.columns
                    and pd.notna(row.get("Valor total aluguel"))
                ):
                    texto += (
                        f" | Valor total: {self._format_currency(row['Valor total aluguel'])}"
                    )
                self._write_body_line(texto, height=5)

            if "Planos" in consult.columns and "Médico" in consult.columns:
                self._draw_subsection_header("Convênios ativos por consultório")
                consult_planos_pdf = consult.copy()
                consult_planos_pdf["Planos"] = (
                    consult_planos_pdf["Planos"].fillna("Nao informado").astype(str).str.strip()
                )
                consult_planos_pdf = (
                    consult_planos_pdf.groupby(["Consultório", "Planos"], observed=False)["Médico"].nunique().reset_index(name="Profissionais")
                )
                consult_planos_pdf = consult_planos_pdf[
                    consult_planos_pdf["Profissionais"].gt(0)
                ]
                if not consult_planos_pdf.empty:
                    consult_planos_pdf = consult_planos_pdf.sort_values(
                        ["Consultório", "Profissionais", "Planos"],
                        ascending=[True, False, True],
                    )
                    for consultorio_nome, grupo in consult_planos_pdf.groupby(
                        "Consultório", observed=False
                    ):
                        grupo_top = grupo.head(5)
                        convenios_txt: List[str] = []
                        for _, plano_row in grupo_top.iterrows():
                            qtd = self._safe_int(plano_row.get("Profissionais", 0)) or 0
                            plano_nome = plano_row.get("Planos", "Nao informado") or "Nao informado"
                            sufixo = "profissional" if qtd == 1 else "profissionais"
                            convenios_txt.append(f"{plano_nome}: {qtd} {sufixo}")
                        resumo_conv = "; ".join(convenios_txt) if convenios_txt else "Nenhum convênio informado"
                        self._write_body_line(
                            f"- {consultorio_nome}: {resumo_conv}", height=5
                        )

        if "Valor Aluguel" in med_pdf.columns:
            valores = med_pdf["Valor Aluguel"].dropna()
            if not valores.empty:
                self._draw_subsection_header("Valores de aluguel")
                media = valores.mean()
                minimo = valores.min()
                maximo = valores.max()
                self._write_body_line(
                    f"- Média: {self._format_currency(media)}", height=5
                )
                self._write_body_line(
                    f"- Mínimo: {self._format_currency(minimo)}", height=5
                )
                self._write_body_line(
                    f"- Máximo: {self._format_currency(maximo)}", height=5
                )

    def _render_agenda_section(self) -> None:
        self._start_section(
            "Agenda Filtrada",
            "Recorte dos agendamentos conforme filtros aplicados no dashboard.",
        )
        agenda_df = self.agenda_df
        if agenda_df is None or agenda_df.empty:
            self._write_body_line("Nenhum agendamento encontrado para os filtros atuais.")
            return

        agenda_view = agenda_df.copy()
        sort_cols = [c for c in ["Sala", "Dia", "Turno"] if c in agenda_view.columns]
        if sort_cols:
            agenda_view = agenda_view.sort_values(sort_cols)

        total_registros = len(agenda_view)
        summary_lines = [f"Total de agendamentos: {total_registros}"]
        if "Sala" in agenda_view.columns:
            salas_distintas = agenda_view["Sala"].dropna().nunique()
            summary_lines.append(f"Salas distintas: {salas_distintas}")
        if "Médico" in agenda_view.columns:
            medicos_distintos = agenda_view["Médico"].dropna().nunique()
            summary_lines.append(f"Profissionais distintos: {medicos_distintos}")

        for resumo in summary_lines:
            self._write_body_line(resumo, height=5)

        self.pdf.ln(2)

        agenda_expected = ["Sala", "Dia", "Turno", "Médico"]
        for col in agenda_expected:
            if col not in agenda_view.columns:
                agenda_view[col] = "—"

        agenda_view = agenda_view[agenda_expected]

        for col in agenda_view.columns:
            series = agenda_view[col]
            if pd.api.types.is_categorical_dtype(series):
                if "—" not in series.cat.categories:
                    agenda_view[col] = series.cat.add_categories(["—"])

        agenda_view = agenda_view.fillna("—")

        def _format_dia(value: object) -> object:
            if isinstance(value, (datetime, pd.Timestamp)):
                return value.strftime("%d/%m/%Y")
            return value

        if "Dia" in agenda_view.columns:
            agenda_view["Dia"] = agenda_view["Dia"].apply(_format_dia)

        agenda_table_columns: List[Tuple[str, float]] = [
            ("Sala", 1.4),
            ("Dia", 1.8),
            ("Turno", 1.4),
            ("Médico", 3.4),
        ]

        chunk_size = 30
        for start in range(0, total_registros, chunk_size):
            bloco = agenda_view.iloc[start : start + chunk_size]
            if bloco.empty:
                continue
            titulo = "Agenda detalhada" if start == 0 else f"Agenda detalhada (registros {start + 1}-{start + len(bloco)})"
            self._draw_subsection_header(titulo)
            self._draw_table(agenda_table_columns, bloco)

    def _render_toc(self) -> None:
        if not self.sections_index:
            return

        self.pdf.add_page()
        self._draw_section_header("Sumário", "Referência rápida das seções geradas.")
        family, style, size = PDF_BODY_FONT
        self.pdf.set_font(family, style, size)
        self.pdf.set_text_color(*PDF_TEXT_COLOR)
        for title, page_number in self.sections_index:
            self.pdf.set_x(self.pdf.l_margin)
            self.pdf.cell(
                self.effective_width - 20,
                6,
                _sanitize_pdf_text(title),
            )
            self.pdf.cell(20, 6, str(page_number), align="R", ln=1)

    def _format_currency(self, value: float) -> str:
        return (
            f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    def _format_currency_value(self, value) -> str:
        if value is None:
            return "—"
        try:
            if pd.isna(value):
                return "—"
        except TypeError:
            pass
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return "—"
        formatted = f"R$ {numeric:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return formatted
