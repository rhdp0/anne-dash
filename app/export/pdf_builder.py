"""PDF builder utilities for exporting the dashboard report."""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from io import BytesIO
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Union

import pandas as pd
from fpdf import FPDF
import plotly.graph_objects as go
import plotly.io as pio


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

PX_PER_MM = 96 / 25.4


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
        overview_timeseries: Optional[Dict[str, pd.DataFrame]] = None,
        top_medicos_turnos: Optional[pd.DataFrame] = None,
        ranking_limits: Optional[Dict[str, int]] = None,
        consultorios_data: Optional[Dict[str, Dict[str, object]]] = None,
        overview_figures: Optional[
            Union[
                Dict[str, go.Figure],
                Sequence[Tuple[str, go.Figure]],
            ]
        ] = None,
        ranking_figures: Optional[
            Union[
                Dict[str, go.Figure],
                Sequence[Tuple[str, go.Figure]],
            ]
        ] = None,
        consultorio_figures: Optional[
            Dict[str, Union[Dict[str, go.Figure], Sequence[Tuple[str, go.Figure]]]]
        ] = None,
        planos_figures: Optional[
            Union[
                Dict[str, go.Figure],
                Sequence[Tuple[str, go.Figure]],
            ]
        ] = None,
    ) -> None:
        self.data_source = data_source or "Origem não informada"
        self.summary_metrics = summary_metrics or {}
        self.ranking_df = ranking_df.copy() if ranking_df is not None else pd.DataFrame()
        self.med_df = med_df.copy() if med_df is not None else pd.DataFrame()
        self.agenda_df = agenda_df.copy() if agenda_df is not None else pd.DataFrame()
        self.overview_timeseries: Dict[str, pd.DataFrame] = {}
        if overview_timeseries:
            for key, value in overview_timeseries.items():
                if isinstance(value, pd.DataFrame):
                    self.overview_timeseries[key] = value.copy()
        self.top_medicos_turnos = (
            top_medicos_turnos.copy()
            if top_medicos_turnos is not None
            else pd.DataFrame()
        )
        self.ranking_limits = ranking_limits or {}
        self.consultorios_data: Dict[str, Dict[str, object]] = {}
        if consultorios_data:
            for key, value in consultorios_data.items():
                if not isinstance(value, dict):
                    continue
                entry: Dict[str, object] = {}
                for sub_key, sub_value in value.items():
                    if isinstance(sub_value, pd.DataFrame):
                        entry[sub_key] = sub_value.copy()
                    else:
                        entry[sub_key] = sub_value
                self.consultorios_data[str(key)] = entry

        self.pdf: Optional[DashboardPDF] = None
        self.effective_width: float = 0.0
        self.sections_index: List[Tuple[str, int]] = []
        self.overview_figures = self._normalize_figures(overview_figures)
        self.ranking_figures = self._normalize_figures(ranking_figures)
        self.planos_figures = self._normalize_figures(planos_figures)
        self.consultorio_figures: Dict[str, List[Tuple[str, go.Figure]]] = {}
        if consultorio_figures:
            for consultorio, payload in consultorio_figures.items():
                figures = self._normalize_figures(payload)
                if figures:
                    self.consultorio_figures[str(consultorio)] = figures

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def build(self) -> bytes:
        """Generate the PDF report and return the resulting bytes."""

        self.pdf = DashboardPDF(
            data_source=_sanitize_pdf_text(str(self.data_source)),
            orientation="P",
            unit="mm",
            format="A4",
        )
        self.pdf.set_margins(PDF_MARGIN, PDF_MARGIN, PDF_MARGIN)
        self.pdf.set_auto_page_break(auto=True, margin=PDF_MARGIN)
        self.pdf.alias_nb_pages()

        self.effective_width = self.pdf.w - self.pdf.l_margin - self.pdf.r_margin
        self.sections_index = []

        self._draw_cover_page()

        self.pdf.add_page()
        self._set_body_font()

        self._render_overview_section()
        self._render_summary_section()
        self._render_ranking_section()
        self._render_consultorios_section()
        self._render_med_info_section()
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

        pdf.set_xy(pdf.l_margin, PDF_MARGIN)
        family, style, size = PDF_SUBTITLE_FONT
        pdf.set_font(family, style, size)
        pdf.set_text_color(*PDF_MUTED_COLOR)
        pdf.multi_cell(
            self.effective_width,
            6,
            _sanitize_pdf_text("Dashboard de Ocupação dos Consultórios"),
        )

        pdf.ln(4)
        family, style, size = PDF_TITLE_FONT
        pdf.set_font(family, style, size)
        pdf.set_text_color(*PDF_PRIMARY_COLOR)
        pdf.multi_cell(
            self.effective_width,
            14,
            _sanitize_pdf_text("Relatório Completo"),
        )

        pdf.ln(2)
        pdf.set_draw_color(*PDF_ACCENT_COLOR)
        pdf.set_line_width(0.6)
        current_y = pdf.get_y()
        pdf.line(
            pdf.l_margin,
            current_y,
            pdf.w - pdf.r_margin,
            current_y,
        )
        pdf.ln(6)

        family, style, size = PDF_SECTION_SUBTITLE_FONT
        pdf.set_font(family, style, size)
        pdf.set_text_color(*PDF_PRIMARY_COLOR)
        pdf.cell(0, 6, _sanitize_pdf_text("Sobre este relatório"), ln=1)

        self._set_body_font()
        about_lines = [
            "Panorama executivo com indicadores de produtividade dos consultórios.",
            f"Fonte dos dados: {self.data_source}.",
            "Geração automática via Dashboard Consultórios.",
        ]
        for line in about_lines:
            self._write_body_line(line, height=5)

        pdf.ln(4)
        self._set_body_font()

    def _draw_section_header(self, title: str, subtitle: Optional[str] = None) -> None:
        pdf = self.pdf
        if pdf.get_y() < PDF_MARGIN:
            pdf.set_y(PDF_MARGIN)
        start_y = pdf.get_y()
        family, style, size = PDF_SECTION_TITLE_FONT
        pdf.set_font(family, style, size)
        pdf.set_text_color(*PDF_PRIMARY_COLOR)
        pdf.set_xy(pdf.l_margin, start_y)
        pdf.cell(0, 9, _sanitize_pdf_text(title), ln=1)

        if subtitle:
            family, style, size = PDF_SECTION_SUBTITLE_FONT
            pdf.set_font(family, style, size)
            pdf.set_text_color(*PDF_MUTED_COLOR)
            pdf.multi_cell(0, 5, _sanitize_pdf_text(subtitle))

        pdf.ln(1)
        self._set_body_font()

    def _draw_subsection_header(self, title: str) -> None:
        family, style, size = PDF_SUBSECTION_FONT
        self.pdf.set_font(family, style, size)
        self.pdf.set_text_color(*PDF_ACCENT_COLOR)
        self.pdf.cell(0, 8, _sanitize_pdf_text(title), ln=1)
        self._set_body_font()

    def _write_body_line(
        self, text: str, height: float = 6, *, indent: float = 0.0
    ) -> None:
        sanitized = _sanitize_pdf_text(text)
        if not sanitized:
            self.pdf.ln(height)
            return

        width = self.effective_width or (
            self.pdf.w - self.pdf.l_margin - self.pdf.r_margin
        )
        if indent:
            width -= indent
        if width <= 0:
            self.pdf.ln(height)
            return

        self.pdf.set_x(self.pdf.l_margin + indent)
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

            label_text = _sanitize_pdf_text(str(label))
            family, style, base_label_size = PDF_KPI_LABEL_FONT
            label_size = float(base_label_size)
            pdf.set_font(family, style, label_size)
            while (
                pdf.get_string_width(label_text) > inner_width and label_size > 6
            ):
                label_size -= 0.5
                pdf.set_font(family, style, label_size)
            pdf.set_text_color(*PDF_MUTED_COLOR)
            pdf.set_xy(inner_x, y + 4)
            pdf.multi_cell(inner_width, 4.5, label_text)

            value_text = _sanitize_pdf_text(str(value))
            family, style, base_value_size = PDF_KPI_VALUE_FONT
            value_size = float(base_value_size)
            pdf.set_font(family, style, value_size)
            while (
                pdf.get_string_width(value_text) > inner_width and value_size > 9
            ):
                value_size -= 0.5
                pdf.set_font(family, style, value_size)
            pdf.set_text_color(*PDF_PRIMARY_COLOR)
            pdf.set_xy(inner_x, y + PDF_CARD_HEIGHT / 2)
            pdf.cell(inner_width, 6, value_text)

            if column == cards_per_row - 1 or idx == len(items) - 1:
                pdf.set_y(y + PDF_CARD_HEIGHT + PDF_CARD_GAP)
            else:
                pdf.set_xy(x + card_width + PDF_CARD_GAP, y)

        pdf.ln(2)
        self._set_body_font()

    def _normalize_figures(
        self,
        figures: Optional[
            Union[
                Dict[str, go.Figure],
                Sequence[Tuple[str, go.Figure]],
            ]
        ],
    ) -> List[Tuple[str, go.Figure]]:
        normalized: List[Tuple[str, go.Figure]] = []
        if not figures:
            return normalized

        if isinstance(figures, dict):
            items = figures.items()
        else:
            items = figures

        for label, figure in items:
            if isinstance(figure, go.Figure):
                caption = str(label).strip() if label is not None else "Gráfico"
                normalized.append((caption, figure))
        return normalized

    def _figure_to_image(
        self, figure: go.Figure
    ) -> Optional[Tuple[BytesIO, float, float]]:  # pragma: no cover - rendering helper
        if self.pdf is None:
            return None

        target_width_mm = max(self.effective_width, 1.0)
        layout_width = getattr(figure.layout, "width", None)
        layout_height = getattr(figure.layout, "height", None)
        if layout_width and layout_height and layout_width > 0:
            aspect_ratio = float(layout_height) / float(layout_width)
        else:
            aspect_ratio = 0.6
        aspect_ratio = max(aspect_ratio, 0.1)

        max_height_mm = max(
            self.pdf.h - self.pdf.t_margin - self.pdf.b_margin - 10,
            40,
        )
        chart_height_mm = target_width_mm * aspect_ratio
        if chart_height_mm > max_height_mm and chart_height_mm > 0:
            scale = max_height_mm / chart_height_mm
            chart_height_mm = max_height_mm
            target_width_mm *= scale

        width_px = max(1, int(round(target_width_mm * PX_PER_MM)))
        height_px = max(1, int(round(chart_height_mm * PX_PER_MM)))

        try:
            image_bytes = pio.to_image(
                figure,
                format="png",
                width=width_px,
                height=height_px,
                scale=1,
            )
        except Exception:
            return None

        return BytesIO(image_bytes), target_width_mm, chart_height_mm

    def _draw_chart(
        self, figure: go.Figure, caption: Optional[str] = None
    ) -> None:  # pragma: no cover - rendering helper
        if self.pdf is None:
            return
        image_payload = self._figure_to_image(figure)
        if not image_payload:
            return

        image_stream, width_mm, height_mm = image_payload
        caption_text = _sanitize_pdf_text(caption or "")
        caption_height = 5 if caption_text else 0
        required_height = height_mm + caption_height + 2
        if self.pdf.get_y() + required_height > self.pdf.page_break_trigger:
            self.pdf.add_page()
            self._set_body_font()

        x_offset = self.pdf.l_margin
        if width_mm < self.effective_width:
            x_offset += (self.effective_width - width_mm) / 2

        start_y = self.pdf.get_y()
        self.pdf.image(image_stream, x=x_offset, y=start_y, w=width_mm)
        self.pdf.set_y(start_y + height_mm + 1.5)

        if caption_text:
            family, style, size = PDF_BODY_FONT
            self.pdf.set_font(family, "I", max(size - 1, 8))
            self.pdf.set_text_color(*PDF_MUTED_COLOR)
            self.pdf.set_x(self.pdf.l_margin)
            self.pdf.multi_cell(self.effective_width, 4.5, caption_text, align="C")
            self.pdf.ln(1)
            self._set_body_font()

    def _draw_figures_group(self, figures: Sequence[Tuple[str, go.Figure]]) -> None:
        for caption, figure in figures:
            if isinstance(figure, go.Figure):
                self._draw_chart(figure, caption)

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

        empty_placeholder = _sanitize_pdf_text("—") or "-"

        def _prepare_cell_text(value: object) -> str:
            if value is None:
                return empty_placeholder
            try:
                if isinstance(value, float) and pd.isna(value):
                    return empty_placeholder
            except TypeError:
                pass

            sanitized_text = _sanitize_pdf_text(str(value))
            return sanitized_text or empty_placeholder

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

    def _render_overview_section(self) -> None:
        overview_map = self.overview_timeseries or {}
        by_sala = overview_map.get("por_sala")
        by_dia = overview_map.get("por_dia")
        by_turno = overview_map.get("por_turno")
        top_medicos = self.top_medicos_turnos

        datasets = [by_sala, by_dia, by_turno, top_medicos]
        if all(df is None or df.empty for df in datasets):
            return

        self._start_section(
            "Panorama de Ocupação",
            "Distribuição agregada dos slots ocupados para orientar a leitura do relatório.",
        )
        self._write_body_line(
            "Os quadros a seguir espelham os dados utilizados nos gráficos de visão geral do dashboard.",
            height=5,
        )
        self.pdf.ln(1)

        def _render_timeseries_table(
            df_source: Optional[pd.DataFrame],
            *,
            title: str,
            primary_label: str,
            rename_map: Optional[Dict[str, str]] = None,
        ) -> None:
            self._draw_subsection_header(title)
            if df_source is None or df_source.empty:
                self._write_body_line("Sem dados disponíveis para este agrupamento.", height=5)
                return

            working = df_source.copy()
            if rename_map:
                working = working.rename(columns=rename_map)

            if "Taxa de Ocupação (%)" in working.columns:
                working = working.sort_values("Taxa de Ocupação (%)", ascending=False)
                working["Taxa de Ocupação (%)"] = working["Taxa de Ocupação (%)"].apply(
                    lambda value: f"{float(value):.1f}%" if pd.notna(value) else "—"
                )

            def _format_slot(value: object) -> object:
                cleaned = self._safe_int(value)
                return cleaned if cleaned is not None else "—"

            for col in ["Slots Ocupados", "Total Slots"]:
                if col in working.columns:
                    working[col] = working[col].apply(_format_slot)

            columns_order = [
                primary_label,
                "Taxa de Ocupação (%)",
                "Slots Ocupados",
                "Total Slots",
            ]
            available_cols = [col for col in columns_order if col in working.columns]
            if not available_cols:
                self._write_body_line("Estrutura de dados inesperada para esta tabela.", height=5)
                return

            limit = 15
            total_registros = len(working)
            working = working[available_cols]
            display = working.head(limit)
            if total_registros > limit:
                self._write_body_line(
                    f"Listando os {limit} primeiros registros de um total de {total_registros} ordenados por maior ocupação.",
                    height=5,
                )

            columns_spec: List[Tuple[str, float]] = []
            for label in available_cols:
                if label == primary_label:
                    columns_spec.append((label, 2.6))
                elif label == "Taxa de Ocupação (%)":
                    columns_spec.append((label, 1.6))
                else:
                    columns_spec.append((label, 1.2))

            self._draw_table(columns_spec, display)

        _render_timeseries_table(
            by_sala,
            title="Ocupação por consultório",
            primary_label="Consultório",
            rename_map={"Sala": "Consultório"},
        )
        _render_timeseries_table(
            by_dia,
            title="Ocupação por dia da semana",
            primary_label="Dia",
        )
        _render_timeseries_table(
            by_turno,
            title="Ocupação por turno",
            primary_label="Turno",
        )

        if self.overview_figures:
            self.pdf.ln(2)
            self._draw_figures_group(self.overview_figures)

        self._draw_subsection_header("Top médicos por turnos utilizados")
        if top_medicos is None or top_medicos.empty:
            self._write_body_line(
                "Sem registros de profissionais ocupando turnos nos filtros atuais.",
                height=5,
            )
        else:
            med_table = top_medicos.copy()
            med_table["Turnos Utilizados"] = pd.to_numeric(
                med_table.get("Turnos Utilizados"), errors="coerce"
            ).fillna(0).astype(int)
            total_medicos = len(med_table)
            self._write_body_line(
                f"Total de profissionais com turnos ocupados: {total_medicos}.",
                height=5,
            )
            med_columns: List[Tuple[str, float]] = [
                ("Médico", 3.2),
                ("Turnos Utilizados", 1.4),
            ]
            self._draw_table(med_columns, med_table)

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

        if self.ranking_figures:
            self.pdf.ln(2)
            self._draw_figures_group(self.ranking_figures)

    def _render_consultorios_section(self) -> None:
        consultorios = self.consultorios_data or {}
        if not consultorios:
            return

        self._start_section(
            "Consultórios selecionados",
            "Detalhamento dos consultórios filtrados com produtividade e agenda resumida.",
        )

        sorted_items = sorted(consultorios.items(), key=lambda item: str(item[0]))

        for consultorio, payload in sorted_items:
            titulo = str(consultorio).strip() or "Consultório não informado"
            self._draw_subsection_header(titulo)

            metrics_bundle: Dict[str, object] = {}
            if isinstance(payload, dict):
                raw_metrics = payload.get("metrics") or payload.get("kpis")
                if isinstance(raw_metrics, dict):
                    metrics_bundle = raw_metrics

            has_content = False
            if metrics_bundle:
                self._draw_kpi_cards(metrics_bundle)
                has_content = True

            top_data = pd.DataFrame()
            if isinstance(payload, dict) and "top_profissionais" in payload:
                source = payload.get("top_profissionais")
                if isinstance(source, pd.DataFrame):
                    top_data = source.copy()
                elif isinstance(source, list):
                    top_data = pd.DataFrame(source)
                elif isinstance(source, dict):
                    top_data = pd.DataFrame([source])

            if not top_data.empty:
                top_data = top_data.head(8).copy()

                def _format_int(value: object) -> object:
                    cleaned = self._safe_int(value)
                    return cleaned if cleaned is not None else "—"

                for column in ["Procedimentos", "Exames", "Cirurgias"]:
                    if column in top_data.columns:
                        top_data[column] = top_data[column].apply(_format_int)

                if "Receita" in top_data.columns:
                    top_data["Receita"] = top_data["Receita"].apply(
                        self._format_currency_value
                    )

                table_columns: List[Tuple[str, float]] = [
                    ("Profissional", 2.8),
                    ("Especialidade", 2.4),
                ]
                optional_specs = [
                    ("Procedimentos", 1.2),
                    ("Exames", 1.0),
                    ("Cirurgias", 1.0),
                    ("Receita", 1.4),
                ]
                for label, width in optional_specs:
                    if label in top_data.columns:
                        table_columns.append((label, width))

                ordered_columns = [label for label, _ in table_columns]
                top_data = top_data[[col for col in ordered_columns if col in top_data.columns]]

                if table_columns and not top_data.empty:
                    self._write_body_line("Profissionais em destaque", height=5)
                    self._draw_table(table_columns, top_data)
                    has_content = True

            agenda_data = pd.DataFrame()
            if isinstance(payload, dict) and "agenda_resumo" in payload:
                source = payload.get("agenda_resumo")
                if isinstance(source, pd.DataFrame):
                    agenda_data = source.copy()
                elif isinstance(source, list):
                    agenda_data = pd.DataFrame(source)
                elif isinstance(source, dict):
                    agenda_data = pd.DataFrame([source])

            if not agenda_data.empty:
                agenda_data = agenda_data.head(12).copy()

                def _format_numeric(value: object) -> object:
                    cleaned = self._safe_int(value)
                    return cleaned if cleaned is not None else "—"

                for column in ["Slots Ocupados", "Total Slots", "Médicos Ativos"]:
                    if column in agenda_data.columns:
                        agenda_data[column] = agenda_data[column].apply(_format_numeric)

                agenda_columns: List[Tuple[str, float]] = []
                for label, width in [
                    ("Dia", 1.8),
                    ("Turno", 1.2),
                    ("Slots Ocupados", 1.3),
                    ("Total Slots", 1.3),
                    ("Médicos Ativos", 1.6),
                ]:
                    if label in agenda_data.columns:
                        agenda_columns.append((label, width))

                if agenda_columns:
                    self._write_body_line("Agenda resumida", height=5)
                    self._draw_table(agenda_columns, agenda_data)
                    has_content = True

            figuras_consultorio = (
                self.consultorio_figures.get(titulo)
                or self.consultorio_figures.get(str(consultorio))
            )
            if figuras_consultorio:
                if not has_content:
                    self._write_body_line("Visualizações do consultório", height=5)
                self._draw_figures_group(figuras_consultorio)
                has_content = True

            if not has_content:
                self._write_body_line(
                    "Sem dados consolidados disponíveis para este consultório.",
                    height=5,
                )

            self.pdf.ln(PDF_SECTION_GAP / 2)

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
                            plano_nome = (
                                plano_row.get("Planos", "Nao informado") or "Nao informado"
                            )
                            sufixo = "profissional" if qtd == 1 else "profissionais"
                            convenios_txt.append(f"{plano_nome}: {qtd} {sufixo}")

                        self._write_body_line(f"- {consultorio_nome}:", height=5)
                        if convenios_txt:
                            for item in convenios_txt:
                                self._write_body_line(f"• {item}", height=5, indent=5)
                        else:
                            self._write_body_line(
                                "• Nenhum convênio informado", height=5, indent=5
                            )
                        self.pdf.ln(1)

        if self.planos_figures:
            self.pdf.ln(1)
            self._draw_figures_group(self.planos_figures)

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
