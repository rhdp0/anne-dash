from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from contextlib import contextmanager
import numpy as np

import pandas as pd
import streamlit as st
import plotly.express as px

from app.data import (
    ConsultorioDataFacade,
    first_nonempty,
    format_consultorio_label,
    normalize_column_name,
)
from app.export.pdf_builder import DashboardPDFBuilder
from app.services import OccupancyAnalyzer


st.set_page_config(page_title="Dashboard Consultórios", layout="wide")

# --- Corporate styling ---
st.markdown("""
<style>
:root {
    --primary-color: #1b3b5f;
    --accent-color: #4c89c6;
    --bg-soft: #f4f7fb;
    --bg-card: #ffffff;
    --text-color: #14213d;
    --muted-color: #5f6c85;
}

.stApp {
    background-color: var(--bg-soft);
    color: var(--text-color);
    font-family: "Segoe UI", "Inter", sans-serif;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
}

section[data-testid="stSidebar"] {
    background-color: #ffffff;
    border-right: 1px solid rgba(27, 59, 95, 0.08);
}

section[data-testid="stSidebar"] > div {
    padding: 1.5rem 1rem;
    display: flex;
    flex-direction: column;
    min-height: 100vh;
    gap: 1.25rem;
}

.sidebar-download-container {
    margin-top: auto;
    padding-top: 1.5rem;
}

.sidebar-download-container .stDownloadButton button {
    width: 100%;
    justify-content: flex-start;
}

h1, h2, h3, h4, h5, h6 {
    color: var(--text-color);
    font-weight: 600;
}

.section-title {
    display: flex;
    align-items: center;
    gap: 0.65rem;
    padding: 0;
    margin: 0 0 1rem;
    font-size: clamp(1.4rem, 1.1rem + 1vw, 1.9rem);
    color: var(--primary-color);
}

.section-subtitle {
    margin-top: -0.5rem;
    margin-bottom: 1.5rem;
    color: var(--muted-color);
    font-size: 0.95rem;
}

.section-card {
    background-color: var(--bg-card);
    border-radius: 1rem;
    padding: 2rem;
    margin: 2.5rem 0;
    border: 1px solid rgba(27, 59, 95, 0.08);
    box-shadow: 0 18px 35px -20px rgba(20, 33, 61, 0.5);
}

.section-card > *:last-child {
    margin-bottom: 0;
}

div[data-testid="stMetricValue"] {
    color: var(--primary-color);
    font-weight: 700;
}

div[data-testid="stMetricLabel"] {
    color: var(--muted-color);
    text-transform: uppercase;
    font-size: 0.75rem;
    letter-spacing: 0.08em;
}

div[data-testid="stMetricDelta"] {
    color: var(--accent-color);
    font-weight: 500;
}

div[data-testid="stTabs"] button {
    border-radius: 999px;
    padding: 0.45rem 1.1rem;
    border: none;
    background-color: transparent;
    color: var(--muted-color);
    font-weight: 500;
}

div[data-testid="stTabs"] button[aria-selected="true"] {
    background-color: rgba(76, 137, 198, 0.12);
    color: var(--primary-color);
}

.stDownloadButton button,
.stButton button {
    background: linear-gradient(135deg, var(--primary-color), var(--accent-color));
    color: #ffffff;
    border: none;
    border-radius: 0.75rem;
    padding: 0.6rem 1.4rem;
    font-weight: 600;
}

.stDownloadButton button:hover,
.stButton button:hover {
    opacity: 0.92;
}

a {
    color: var(--accent-color);
}
</style>
""", unsafe_allow_html=True)

@contextmanager
def section_block(title: str, description: Optional[str] = None, anchor: Optional[str] = None):
    wrapper = st.container()
    if anchor:
        wrapper.markdown(f'<div id="{anchor}"></div>', unsafe_allow_html=True)
    wrapper.markdown('<div class="section-card">', unsafe_allow_html=True)
    wrapper.markdown(f'<h2 class="section-title">{title}</h2>', unsafe_allow_html=True)
    if description:
        wrapper.markdown(f'<p class="section-subtitle">{description}</p>', unsafe_allow_html=True)
    content = wrapper.container()
    try:
        with content:
            yield content
    finally:
        wrapper.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div id="topo"></div>', unsafe_allow_html=True)
st.title("🏥 Dashboard de Ocupação dos Consultórios")
st.caption("Lendo somente as abas **CONSULTÓRIO** (ignorando 'OCUPAÇÃO DAS SALAS'). Integra automaticamente TODAS as abas **MÉDICOS** (ex.: 'MÉDICOS 1', 'MÉDICOS 2', 'MÉDICOS 3').")

DEFAULT_PATH = Path("/mnt/data/ESCALA DOS CONSULTORIOS DEFINITIVO.xlsx")

# ---------- Sidebar: Upload ----------
st.sidebar.header("📂 Fonte de Dados")
uploaded = st.sidebar.file_uploader("Envie o Excel (.xlsx)", type=["xlsx"], key="main_xlsx")

data_facade = ConsultorioDataFacade()

excel = None
if uploaded is not None:
    try:
        excel = data_facade.load_workbook(uploaded)
        fonte = "Upload do usuário"
    except ValueError as exc:
        st.error(str(exc))
elif DEFAULT_PATH.exists():
    try:
        excel = data_facade.load_workbook(DEFAULT_PATH)
        fonte = f"Arquivo padrão: {DEFAULT_PATH.name}"
    except ValueError as exc:
        st.error(str(exc))
else:
    st.error("Nenhum arquivo encontrado. Envie um Excel com as abas de CONSULTÓRIO.")
    st.stop()

if excel is None:
    st.stop()

st.sidebar.success(f"Usando dados de: {fonte}")
# A navegação por seções será configurada após os filtros.

# ---------- Carregamento de dados ----------
datasets = data_facade.load_dataset(excel)
df = datasets.get("agenda", pd.DataFrame())
produtividade_df = datasets.get("produtividade", pd.DataFrame())
med_df = datasets.get("medicos", pd.DataFrame())

if df.empty:
    st.error("Não foram encontrados dados nas abas 'CONSULTÓRIO'.")
    st.stop()

# ---------- Utilitários ----------
def format_currency_value(value) -> str:
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

# ---------- Filtros ----------
st.sidebar.header("🔎 Filtros")
salas = sorted(df["Sala"].dropna().unique().tolist())
dias = [d for d in ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado"] if d in df["Dia"].astype(str).unique()]
turnos = sorted(df["Turno"].dropna().unique().tolist())
medicos = sorted([m for m in df["Médico"].dropna().unique().tolist() if m])

sel_salas = st.sidebar.multiselect("Consultório(s)", salas, default=salas)
sel_dias = st.sidebar.multiselect("Dia(s)", dias, default=dias)
sel_turnos = st.sidebar.multiselect("Turno(s)", turnos, default=turnos)
sel_medicos = st.sidebar.multiselect("Médico(s)", medicos, default=[], help="Deixe vazio para não filtrar por médico.")

filtered_df = data_facade.filter_by_date(
    df,
    allowed_values=sel_dias,
    date_column="Dia",
)

# Configuração das seções disponíveis no dashboard
section_labels = (
    "📊 Visão Geral",
    "🏆 Ranking",
    "🔍 Consultórios",
    "💼 Planos & Aluguel",
    "📋 Agenda",
)
selected_section = st.sidebar.radio(
    "Ir para a seção",
    section_labels,
    index=0,
    key="selected_section",
)

sidebar_pdf_container = st.sidebar.container()

# Base para KPIs (NÃO filtra por médico)
mask_base = (
    filtered_df["Sala"].isin(sel_salas)
    & filtered_df["Turno"].isin(sel_turnos)
)
fdf_base = filtered_df[mask_base].copy()

# Aplicar filtro de médico apenas onde fizer sentido
if sel_medicos:
    mask_medico = filtered_df["Médico"].isin(sel_medicos)
else:
    mask_medico = pd.Series(True, index=filtered_df.index)
fdf = filtered_df[mask_base & mask_medico].copy()

ranking_prod_total = pd.DataFrame()
receita_por_medico = pd.DataFrame()
receita_por_consultorio = pd.DataFrame()
produtividade_base = pd.DataFrame()
if not produtividade_df.empty:
    produtividade_base = produtividade_df.copy()
    produtividade_base["Especialidade"] = (
        produtividade_base["Especialidade"].fillna("").astype(str).str.strip()
    )
    produtividade_base.loc[
        produtividade_base["Especialidade"].eq(""), "Especialidade"
    ] = "Não informada"

    agg_map = {
        "Exames Solicitados": "sum",
        "Cirurgias Solicitadas": "sum",
    }
    if "CRM" in produtividade_base.columns:
        agg_map["CRM"] = first_nonempty
    if "Receita" in produtividade_base.columns:
        agg_map["Receita"] = "sum"

    ranking_prod_total = data_facade.group_metrics(
        produtividade_base,
        ["Consultório", "Especialidade", "Profissional"],
        agg_map,
    )

    if "CRM" not in ranking_prod_total.columns:
        ranking_prod_total["CRM"] = ""

    if "Receita" not in ranking_prod_total.columns:
        ranking_prod_total["Receita"] = 0.0

    ranking_prod_total["Receita"] = pd.to_numeric(
        ranking_prod_total["Receita"], errors="coerce"
    ).fillna(0.0)

    for col in ["Exames Solicitados", "Cirurgias Solicitadas"]:
        ranking_prod_total[col] = pd.to_numeric(ranking_prod_total[col], errors="coerce").fillna(0)

    ranking_prod_total["Total Procedimentos"] = (
        ranking_prod_total["Exames Solicitados"] + ranking_prod_total["Cirurgias Solicitadas"]
    )

    ranking_prod_total = ranking_prod_total[
        (ranking_prod_total["Total Procedimentos"] > 0)
        | (ranking_prod_total["Receita"] > 0)
    ]

    for col in ["Exames Solicitados", "Cirurgias Solicitadas", "Total Procedimentos"]:
        ranking_prod_total[col] = ranking_prod_total[col].round().astype(int)

    ranking_prod_total["Receita"] = ranking_prod_total["Receita"].astype(float)

    ranking_prod_total["SalaNorm"] = ranking_prod_total["Consultório"].apply(normalize_column_name)
    ranking_prod_total["Etiqueta"] = ranking_prod_total.apply(
        lambda r: (
            f"{r['Profissional']} - {r['Especialidade']} ({r['Consultório']})"
            if r.get("Especialidade") and r.get("Especialidade") != "Não informada"
            else f"{r['Profissional']} ({r['Consultório']})"
        ),
        axis=1,
    )

    receita_por_medico = (
        ranking_prod_total.groupby("Profissional", as_index=False)["Receita"].sum()
        .rename(columns={"Receita": "Receita Total"})
        .sort_values("Receita Total", ascending=False)
    )

    receita_por_consultorio = (
        ranking_prod_total.groupby("Consultório", as_index=False)["Receita"].sum()
        .rename(columns={"Receita": "Receita Total"})
        .sort_values("Receita Total", ascending=False)
    )

occupancy_service = OccupancyAnalyzer(
    base_df=fdf_base,
    filtered_df=fdf,
    selected_salas=sel_salas,
    selected_dias=sel_dias,
    selected_turnos=sel_turnos,
    selected_medicos=sel_medicos,
    ranking_df=ranking_prod_total,
)
overview_timeseries = {
    "por_sala": occupancy_service.build_timeseries(["Sala"]),
    "por_dia": occupancy_service.build_timeseries(["Dia"]),
    "por_turno": occupancy_service.build_timeseries(["Turno"]),
}

heatmap_dimension = None
for candidate in ["Horário", "Horario", "Turno"]:
    if candidate in fdf.columns:
        heatmap_dimension = candidate
        break

heatmap_source = pd.DataFrame()
if heatmap_dimension is not None and not fdf.empty and "Ocupado" in fdf.columns:
    heatmap_source = (
        fdf.groupby(["Dia", heatmap_dimension], observed=False)["Ocupado"]
        .agg(
            **{
                "Slots Ocupados": lambda s: int(s.fillna(False).astype(bool).sum()),
                "Total Slots": "size",
            }
        )
        .reset_index()
    )
    heatmap_source["Taxa de Ocupação (%)"] = (
        heatmap_source["Slots Ocupados"] / heatmap_source["Total Slots"] * 100
    ).round(1)
    heatmap_source["Slots Ocupados"] = (
        heatmap_source["Slots Ocupados"].fillna(0).astype(int)
    )
    heatmap_source["Total Slots"] = heatmap_source["Total Slots"].fillna(0).astype(int)
top_medicos_turnos = occupancy_service.top_medicos_por_turnos(15)
kpis = occupancy_service.get_kpi_summary()
summary_metrics = occupancy_service.build_summary_metadata()
if "Receita total (produtividade)" in summary_metrics:
    summary_metrics["Receita total (produtividade)"] = format_currency_value(
        summary_metrics["Receita total (produtividade)"]
    )

overview_pdf_figures: List[Tuple[str, object]] = []
ranking_pdf_figures: List[Tuple[str, object]] = []
planos_pdf_figures: List[Tuple[str, object]] = []
consultorio_pdf_figures: Dict[str, List[Tuple[str, object]]] = {}

if selected_section == "📊 Visão Geral":
    with section_block(
        "📊 Visão Geral",
        description="Resumo executivo dos consultórios e turnos filtrados.",
        anchor="visao-geral",
    ) as sec:
        c1, c2, c3, c4 = sec.columns(4)
        c1.metric("Consultórios selecionados", kpis.total_salas)
        c2.metric("Slots (dia × turno × sala)", kpis.total_slots)
        c3.metric("Slots livres", kpis.slots_livres)
        c4.metric("Ocupados", kpis.slots_ocupados)

        kc1, kc2 = sec.columns(2)
        kc1.metric("Taxa de ocupação", f"{kpis.taxa_ocupacao:.1f}%")
        kc2.metric("Médicos distintos", kpis.medicos_distintos)

        colA, colB = sec.columns(2)
        by_sala = overview_timeseries.get("por_sala", pd.DataFrame())
        fig1 = px.bar(
            by_sala,
            x="Sala",
            y="Taxa de Ocupação (%)",
            title="Ocupação por consultório",
            text="Taxa de Ocupação (%)",
        )
        fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig1.update_yaxes(range=[0, 100])
        colA.plotly_chart(fig1, use_container_width=True)
        overview_pdf_figures.append(("Ocupação por consultório", fig1))

        by_dia = overview_timeseries.get("por_dia", pd.DataFrame())
        fig2 = px.bar(
            by_dia,
            x="Dia",
            y="Taxa de Ocupação (%)",
            title="Ocupação por dia da semana",
            text="Taxa de Ocupação (%)",
        )
        fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig2.update_yaxes(range=[0, 100])
        colB.plotly_chart(fig2, use_container_width=True)
        overview_pdf_figures.append(("Ocupação por dia da semana", fig2))

        colC, colD = sec.columns(2)
        by_turno = overview_timeseries.get("por_turno", pd.DataFrame())
        fig3 = px.bar(
            by_turno,
            x="Turno",
            y="Taxa de Ocupação (%)",
            title="Ocupação por turno",
            text="Taxa de Ocupação (%)",
        )
        fig3.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig3.update_yaxes(range=[0, 100])
        colC.plotly_chart(fig3, use_container_width=True)
        overview_pdf_figures.append(("Ocupação por turno", fig3))

        top_med = top_medicos_turnos
        if not top_med.empty:
            fig4 = px.bar(
                top_med,
                x="Turnos Utilizados",
                y="Médico",
                orientation="h",
                title="Top médicos por turnos utilizados",
                text="Turnos Utilizados",
            )
            fig4.update_traces(textposition="outside")
            colD.plotly_chart(fig4, use_container_width=True)
            overview_pdf_figures.append(("Top médicos por turnos utilizados", fig4))
        else:
            colD.info("Sem médicos ocupando slots nos filtros atuais.")

        st.markdown("##### Ocupação por dia × turno/horário")
        if heatmap_dimension is None:
            st.info("Inclua uma coluna de turno ou horário na agenda para gerar o mapa de ocupação.")
        elif heatmap_source.empty:
            st.info("Sem dados suficientes para montar a tabela dinâmica nos filtros atuais.")
        else:
            metric_options = ["Taxa de Ocupação (%)", "Slots Ocupados", "Total Slots"]
            heatmap_metric = st.radio(
                "Métrica exibida",
                metric_options,
                horizontal=True,
                key="heatmap_metric_overview",
            )

            heatmap_pivot = heatmap_source.pivot(
                index="Dia", columns=heatmap_dimension, values=heatmap_metric
            )
            if "Dia" in fdf.columns and pd.api.types.is_categorical_dtype(fdf["Dia"]):
                ordered_days = [day for day in fdf["Dia"].cat.categories if day in heatmap_pivot.index]
                heatmap_pivot = heatmap_pivot.loc[ordered_days]

            is_percentage = "Taxa" in heatmap_metric
            color_scale = [[0, "#e3f2fd"], [0.35, "#bbdefb"], [0.7, "#64b5f6"], [1, "#1b3b5f"]]
            fig_heatmap = px.imshow(
                heatmap_pivot,
                text_auto=".1f" if is_percentage else True,
                aspect="auto",
                color_continuous_scale=color_scale,
                zmin=0,
                zmax=100 if is_percentage else None,
                labels={"color": heatmap_metric, "x": heatmap_dimension, "y": "Dia"},
                title="Mapa de ocupação dos slots",
            )
            fig_heatmap.update_layout(
                margin=dict(t=60, r=20, l=20, b=20), coloraxis_colorbar=dict(title="Ocupação")
            )
            fig_heatmap.update_traces(hovertemplate="Dia: %{y}<br>" + f"{heatmap_dimension}: %{x}<br>" + "Valor: %{z}<extra></extra>")
            st.plotly_chart(fig_heatmap, use_container_width=True)

            st.caption(
                "Células mais claras destacam horários ociosos ou com menos slots ocupados conforme a métrica escolhida."
            )


if selected_section == "🏆 Ranking":
    with section_block(
        "🏆 Ranking de produtividade dos médicos",
        description="Comparativo completo dos profissionais considerando solicitações, cirurgias, exames e receita registrada.",
        anchor="ranking",
    ) as sec:
        if ranking_prod_total.empty:
            sec.info("Sem dados nas abas de produtividade para gerar o ranking geral.")
        else:
            receita_total = ranking_prod_total["Receita"].sum()
            col_receita_total, col_receita_medico, col_receita_consult = sec.columns(3)
            col_receita_total.metric(
                "Receita total registrada",
                format_currency_value(receita_total) if receita_total else "—",
            )

            if not receita_por_medico.empty:
                top_medico_receita = receita_por_medico.iloc[0]
                col_receita_medico.metric(
                    "Maior receita por médico",
                    format_currency_value(top_medico_receita["Receita Total"]),
                    top_medico_receita["Profissional"],
                )
            else:
                col_receita_medico.metric("Maior receita por médico", "—", "Sem dados")

            if not receita_por_consultorio.empty:
                top_consult_receita = receita_por_consultorio.iloc[0]
                col_receita_consult.metric(
                    "Maior receita por consultório",
                    format_currency_value(top_consult_receita["Receita Total"]),
                    top_consult_receita["Consultório"],
                )
            else:
                col_receita_consult.metric("Maior receita por consultório", "—", "Sem dados")

            ranking_total = ranking_prod_total.sort_values(
                [
                    "Total Procedimentos",
                    "Cirurgias Solicitadas",
                    "Exames Solicitados",
                    "Profissional",
                    "Consultório",
                ],
                ascending=[False, False, False, True, True],
            ).reset_index(drop=True)
            ranking_total.insert(0, "Rank", range(1, len(ranking_total) + 1))

            ranking_exames = ranking_prod_total.sort_values(
                [
                    "Exames Solicitados",
                    "Cirurgias Solicitadas",
                    "Total Procedimentos",
                    "Profissional",
                    "Consultório",
                ],
                ascending=[False, False, False, True, True],
            ).reset_index(drop=True)
            ranking_exames.insert(0, "Rank", range(1, len(ranking_exames) + 1))

            ranking_cirurgias = ranking_prod_total.sort_values(
                [
                    "Cirurgias Solicitadas",
                    "Exames Solicitados",
                    "Total Procedimentos",
                    "Profissional",
                    "Consultório",
                ],
                ascending=[False, False, False, True, True],
            ).reset_index(drop=True)
            ranking_cirurgias.insert(0, "Rank", range(1, len(ranking_cirurgias) + 1))

            ranking_receita = ranking_prod_total.sort_values(
                ["Receita", "Total Procedimentos", "Profissional", "Consultório"],
                ascending=[False, False, True, True],
            ).reset_index(drop=True)
            ranking_receita.insert(0, "Rank", range(1, len(ranking_receita) + 1))

            if ranking_total.empty:
                sec.info("Sem registros de produtividade para os filtros atuais.")
            else:
                max_slider = max(1, len(ranking_total))
                top_n_default = min(max_slider, 10)
                top_n = sec.slider(
                    "Quantidade de profissionais no ranking",
                    min_value=1,
                    max_value=max_slider,
                    value=top_n_default,
                    key="ranking_produtividade_top",
                )

                top_total = ranking_total.head(top_n)
                top_exames = ranking_exames.head(top_n)
                top_cirurgias = ranking_cirurgias.head(top_n)
                top_receita = ranking_receita.head(top_n)

                tab_total, tab_exames, tab_cirurgias, tab_receita = sec.tabs(
                    ["Produtividade Geral", "Top Exames", "Top Cirurgias", "Top Receita"]
                )

                def _render_highlights(container, dataset):
                    destaques = dataset.head(3).to_dict("records")
                    if not destaques:
                        container.info("Sem registros para os filtros atuais.")
                        return
                    destaque_cols = container.columns(len(destaques))
                    for col, row in zip(destaque_cols, destaques):
                        total = int(row.get("Total Procedimentos", 0))
                        exames = int(row.get("Exames Solicitados", 0))
                        cirurgias = int(row.get("Cirurgias Solicitadas", 0))
                        receita_valor = float(row.get("Receita", 0) or 0)
                        profissional = row.get("Profissional", "")
                        especialidade = row.get("Especialidade", "")
                        consultorio = row.get("Consultório", "")
                        rank = row.get("Rank", "-")

                        titulo = f"{rank}º {profissional}" if profissional else f"{rank}º Profissional"
                        if especialidade and especialidade != "Não informada":
                            titulo = f"{titulo} - {especialidade}"
                        if consultorio:
                            titulo = f"{titulo} ({consultorio})"

                        metric_value = f"{total} Solicitações"
                        delta_parts = [f"Exames: {exames}", f"Cirurgias: {cirurgias}"]
                        if receita_valor:
                            metric_value = format_currency_value(receita_valor)
                            delta_parts.insert(0, f"Solicitações: {total}")
                            delta_parts.append(f"Receita: {format_currency_value(receita_valor)}")
                        col.metric(titulo, metric_value, " • ".join(delta_parts))

                def _render_chart(container, dataset, value_col, title, label_col="Etiqueta", is_currency=False):
                    if dataset.empty:
                        container.info("Sem registros para os filtros atuais.")
                        return

                    display_df = dataset.copy()
                    display_df[value_col] = pd.to_numeric(display_df[value_col], errors="coerce").fillna(0)
                    if is_currency:
                        display_df["__text"] = display_df[value_col].apply(format_currency_value)
                    else:
                        display_df[value_col] = display_df[value_col].round().astype(int)
                        display_df["__text"] = display_df[value_col]

                    fig = px.bar(
                        display_df,
                        x=value_col,
                        y=label_col,
                        orientation="h",
                        color=value_col,
                        color_continuous_scale="Blues",
                        title=title,
                        text="__text",
                    )
                    fig.update_layout(coloraxis_showscale=False)
                    if is_currency:
                        fig.update_traces(
                            texttemplate="%{text}",
                            textposition="outside",
                            customdata=display_df[["Rank", "Consultório", "Especialidade", "Total Procedimentos"]],
                            hovertemplate=(
                                "%{customdata[0]}º %{y}<br>"
                                "Consultório: %{customdata[1]}<br>"
                                "Especialidade: %{customdata[2]}<br>"
                                "Receita: %{text}<br>"
                                "Total de procedimentos: %{customdata[3]}<extra></extra>"
                            ),
                        )
                    else:
                        fig.update_traces(
                            texttemplate="%{text}",
                            textposition="outside",
                            customdata=display_df[["Rank", "Consultório", "Especialidade", "Exames Solicitados", "Cirurgias Solicitadas"]],
                            hovertemplate=(
                                "%{customdata[0]}º %{y}<br>"
                                "Consultório: %{customdata[1]}<br>"
                                "Especialidade: %{customdata[2]}<br>"
                                "Exames solicitados: %{customdata[3]}<br>"
                                "Cirurgias solicitadas: %{customdata[4]}<extra></extra>"
                            ),
                        )
                    fig.update_yaxes(
                        categoryorder="array",
                        categoryarray=display_df[label_col].tolist()[::-1],
                    )
                    container.plotly_chart(fig, use_container_width=True)
                    ranking_pdf_figures.append((title, fig))

                with tab_total:
                    _render_highlights(tab_total, top_total)
                    if not top_total.empty:
                        total_display = top_total.copy()
                        total_display["Total Solicitações"] = total_display["Total Procedimentos"]
                        _render_chart(
                            tab_total,
                            total_display,
                            "Total Solicitações",
                            "Top profissionais por produtividade",
                        )

                with tab_exames:
                    _render_highlights(tab_exames, top_exames)
                    if not top_exames.empty:
                        _render_chart(
                            tab_exames,
                            top_exames,
                            "Exames Solicitados",
                            "Top profissionais por exames solicitados",
                        )

                with tab_cirurgias:
                    _render_highlights(tab_cirurgias, top_cirurgias)
                    if not top_cirurgias.empty:
                        _render_chart(
                            tab_cirurgias,
                            top_cirurgias,
                            "Cirurgias Solicitadas",
                            "Top profissionais por cirurgias solicitadas",
                        )

                with tab_receita:
                    _render_highlights(tab_receita, top_receita)
                    if not top_receita.empty:
                        _render_chart(
                            tab_receita,
                            top_receita,
                            "Receita",
                            "Top profissionais por receita",
                            is_currency=True,
                        )

                if not receita_por_consultorio.empty or not receita_por_medico.empty:
                    sec.markdown("#### Distribuição de receita consolidada")
                    graf_receita_consult, graf_receita_medico = sec.columns(2)

                    if not receita_por_consultorio.empty:
                        consult_display = receita_por_consultorio.head(15).copy()
                        consult_display["Receita Formatada"] = consult_display["Receita Total"].apply(
                            format_currency_value
                        )
                        fig_receita_consult = px.bar(
                            consult_display,
                            x="Receita Total",
                            y="Consultório",
                            orientation="h",
                            title="Top consultórios por receita",
                            text="Receita Formatada",
                        )
                        fig_receita_consult.update_traces(textposition="outside")
                        fig_receita_consult.update_yaxes(
                            categoryorder="array",
                            categoryarray=consult_display["Consultório"].tolist()[::-1],
                        )
                        graf_receita_consult.plotly_chart(fig_receita_consult, use_container_width=True)
                        ranking_pdf_figures.append(("Top consultórios por receita", fig_receita_consult))
                    else:
                        graf_receita_consult.info("Sem dados de receita por consultório.")

                    if not receita_por_medico.empty:
                        med_display = receita_por_medico.head(15).copy()
                        med_display["Receita Formatada"] = med_display["Receita Total"].apply(
                            format_currency_value
                        )
                        fig_receita_medico = px.bar(
                            med_display,
                            x="Receita Total",
                            y="Profissional",
                            orientation="h",
                            title="Top médicos por receita consolidada",
                            text="Receita Formatada",
                        )
                        fig_receita_medico.update_traces(textposition="outside")
                        fig_receita_medico.update_yaxes(
                            categoryorder="array",
                            categoryarray=med_display["Profissional"].tolist()[::-1],
                        )
                        graf_receita_medico.plotly_chart(fig_receita_medico, use_container_width=True)
                        ranking_pdf_figures.append(("Top médicos por receita consolidada", fig_receita_medico))
                    else:
                        graf_receita_medico.info("Sem dados de receita por médico consolidada.")
if selected_section == "🔍 Consultórios":
    # ---------- Visão individual por consultório ----------
    with section_block(
        "🔍 Indicadores individuais por consultório",
        description="Análise aprofundada das salas selecionadas com destaque de produtividade.",
        anchor="consultorio",
    ):

        salas_disponiveis = sorted(df["Sala"].dropna().unique().tolist())
        if not salas_disponiveis:
            st.info("Não há consultórios disponíveis para detalhar.")
        else:
            sala_detalhe = st.selectbox("Escolha um consultório para detalhar", salas_disponiveis, key="detalhe_sala")
            sala_label_pdf = format_consultorio_label(sala_detalhe)

            mask_sala_base = ((df["Sala"] == sala_detalhe)
                              & df["Dia"].astype(str).isin(sel_dias)
                              & df["Turno"].isin(sel_turnos))
            mask_sala = mask_sala_base & mask_medico

            detalhe_base = df[mask_sala_base].copy()
            detalhe_df = df[mask_sala].copy()

            if detalhe_base.empty:
                st.info("Sem dados para o consultório selecionado com os filtros atuais de dia/turno.")
            else:
                kpis_consultorio = OccupancyAnalyzer.compute_basic_metrics(detalhe_base)

                ic1, ic2, ic3, ic4 = st.columns(4)
                ic1.metric("Consultório", sala_detalhe)
                ic2.metric("Slots do consultório", kpis_consultorio.total_slots)
                ic3.metric("Slots livres", kpis_consultorio.slots_livres)
                ic4.metric("Ocupados", kpis_consultorio.slots_ocupados)

                ic5, ic6 = st.columns(2)
                ic5.metric(
                    "Taxa de ocupação do consultório",
                    f"{kpis_consultorio.taxa_ocupacao:.1f}%",
                )
                ic6.metric(
                    "Médicos distintos no consultório",
                    kpis_consultorio.medicos_distintos,
                )

                ranking_ind_total = pd.DataFrame()
                ranking_ind_exames = pd.DataFrame()
                ranking_ind_cirurgias = pd.DataFrame()
                empty_ind_cols = [
                    "Profissional",
                    "Especialidade",
                    "Consultório",
                    "CRM",
                    "Total Procedimentos",
                    "Exames Solicitados",
                    "Cirurgias Solicitadas",
                    "Receita",
                    "EtiquetaLocal",
                    "Rank",
                ]
                top_total_ind = pd.DataFrame(columns=empty_ind_cols)
                top_exames_ind = pd.DataFrame(columns=empty_ind_cols)
                top_cirurgias_ind = pd.DataFrame(columns=empty_ind_cols)
                top_receita_ind = pd.DataFrame(columns=empty_ind_cols)
                sala_norm = normalize_column_name(sala_detalhe)
                if ranking_prod_total.empty:
                    st.info("Sem dados de produtividade carregados para detalhar este consultório.")
                else:
                    ranking_ind_base = ranking_prod_total[ranking_prod_total["SalaNorm"] == sala_norm].copy()
                    if ranking_ind_base.empty:
                        st.info("Sem registros de produtividade para o consultório selecionado.")
                    else:
                        receita_total_consultorio = ranking_ind_base["Receita"].sum()
                        receita_media_profissional = 0.0
                        if not ranking_ind_base.empty:
                            receita_media_profissional = (
                                ranking_ind_base.groupby("Profissional")["Receita"].sum().mean()
                            ) or 0.0
                        ic7, ic8 = st.columns(2)
                        ic7.metric(
                            "Receita total no consultório",
                            format_currency_value(receita_total_consultorio),
                        )
                        ic8.metric(
                            "Receita média por médico",
                            format_currency_value(receita_media_profissional)
                            if not pd.isna(receita_media_profissional)
                            else "—",
                        )

                        ranking_ind_total = ranking_ind_base.sort_values(
                            ["Total Procedimentos", "Cirurgias Solicitadas", "Exames Solicitados", "Profissional"],
                            ascending=[False, False, False, True],
                        ).reset_index(drop=True)
                        ranking_ind_total.insert(0, "Rank", range(1, len(ranking_ind_total) + 1))
                        ranking_ind_total["EtiquetaLocal"] = ranking_ind_total.apply(
                            lambda r: f"{r['Profissional']} - {r['Especialidade']}"
                            if r.get("Especialidade") and r.get("Especialidade") != "Não informada"
                            else r.get("Profissional", ""),
                            axis=1,
                        )

                        ranking_ind_exames = ranking_ind_base.sort_values(
                            ["Exames Solicitados", "Cirurgias Solicitadas", "Total Procedimentos", "Profissional"],
                            ascending=[False, False, False, True],
                        ).reset_index(drop=True)
                        ranking_ind_exames.insert(0, "Rank", range(1, len(ranking_ind_exames) + 1))
                        ranking_ind_exames["EtiquetaLocal"] = ranking_ind_exames.apply(
                            lambda r: f"{r['Profissional']} - {r['Especialidade']}"
                            if r.get("Especialidade") and r.get("Especialidade") != "Não informada"
                            else r.get("Profissional", ""),
                            axis=1,
                        )

                        ranking_ind_cirurgias = ranking_ind_base.sort_values(
                            ["Cirurgias Solicitadas", "Exames Solicitados", "Total Procedimentos", "Profissional"],
                            ascending=[False, False, False, True],
                        ).reset_index(drop=True)
                        ranking_ind_cirurgias.insert(0, "Rank", range(1, len(ranking_ind_cirurgias) + 1))
                        ranking_ind_cirurgias["EtiquetaLocal"] = ranking_ind_cirurgias.apply(
                            lambda r: f"{r['Profissional']} - {r['Especialidade']}"
                            if r.get("Especialidade") and r.get("Especialidade") != "Não informada"
                            else r.get("Profissional", ""),
                            axis=1,
                        )

                        ranking_ind_receita = ranking_ind_base.sort_values(
                            ["Receita", "Total Procedimentos", "Profissional"],
                            ascending=[False, False, True],
                        ).reset_index(drop=True)
                        ranking_ind_receita.insert(0, "Rank", range(1, len(ranking_ind_receita) + 1))
                        ranking_ind_receita["EtiquetaLocal"] = ranking_ind_receita.apply(
                            lambda r: f"{r['Profissional']} - {r['Especialidade']}"
                            if r.get("Especialidade") and r.get("Especialidade") != "Não informada"
                            else r.get("Profissional", ""),
                            axis=1,
                        )

                        top_n_ind_default = min(len(ranking_ind_total), 10) if len(ranking_ind_total) else 1
                        top_n_ind = st.slider(
                            "Quantidade de profissionais no ranking do consultório",
                            min_value=1,
                            max_value=len(ranking_ind_total),
                            value=top_n_ind_default,
                            key=f"ranking_ind_top_{sala_norm}",
                        )

                        top_total_ind = ranking_ind_total.head(top_n_ind)
                        top_exames_ind = ranking_ind_exames.head(top_n_ind)
                        top_cirurgias_ind = ranking_ind_cirurgias.head(top_n_ind)
                        top_receita_ind = ranking_ind_receita.head(top_n_ind)

                        st.markdown("#### Destaques de produtividade no consultório")
                        tabs_ind = st.tabs(["Produtividade Geral", "Top Exames", "Top Cirurgias", "Top Receita"])

                        def _render_ind_highlights(dataset: pd.DataFrame) -> None:
                            destaques = dataset.head(3).to_dict("records")
                            if not destaques:
                                st.info("Sem registros para os filtros atuais.")
                                return
                            destaque_cols_ind = st.columns(len(destaques))
                            for col, row in zip(destaque_cols_ind, destaques):
                                total = int(row.get("Total Procedimentos", 0))
                                exames = int(row.get("Exames Solicitados", 0))
                                cirurgias = int(row.get("Cirurgias Solicitadas", 0))
                                receita_valor = float(row.get("Receita", 0) or 0)
                                profissional = row.get("Profissional", "")
                                especialidade = row.get("Especialidade", "")
                                crm = row.get("CRM", "")
                                rank = row.get("Rank", "-")

                                titulo_local = f"{rank}º {profissional}" if profissional else f"{rank}º Profissional"
                                if especialidade and especialidade != "Não informada":
                                    titulo_local = f"{titulo_local} - {especialidade}"

                                delta_parts = [f"Exames: {exames}", f"Cirurgias: {cirurgias}"]
                                if receita_valor:
                                    delta_parts.append(f"Receita: {format_currency_value(receita_valor)}")
                                if crm:
                                    delta_parts.insert(0, f"CRM {crm}")

                                metric_value = f"{total} Solicitações"
                                if receita_valor:
                                    metric_value = format_currency_value(receita_valor)
                                    delta_parts.insert(0, f"Solicitações: {total}")

                                col.metric(
                                    titulo_local,
                                    metric_value,
                                    " • ".join(delta_parts),
                                )

                        def _render_ind_chart(
                            dataset: pd.DataFrame,
                            value_col: str,
                            title: str,
                            is_currency: bool = False,
                        ) -> None:
                            if dataset.empty:
                                st.info("Sem registros para os filtros atuais.")
                                return

                            display_df = dataset.copy()
                            display_df[value_col] = pd.to_numeric(display_df[value_col], errors="coerce").fillna(0)
                            if is_currency:
                                display_df["__text"] = display_df[value_col].apply(format_currency_value)
                            else:
                                display_df[value_col] = display_df[value_col].round().astype(int)
                                display_df["__text"] = display_df[value_col]
                            fig = px.bar(
                                display_df,
                                x=value_col,
                                y="EtiquetaLocal",
                                orientation="h",
                                title=title,
                                text="__text",
                            )
                            if is_currency:
                                fig.update_traces(
                                    textposition="outside",
                                    customdata=display_df[["Rank", "Total Procedimentos"]],
                                    hovertemplate=(
                                        "%{customdata[0]}º %{y}<br>"
                                        "Receita: %{text}<br>"
                                        "Total de procedimentos: %{customdata[1]}<extra></extra>"
                                    ),
                                )
                            else:
                                fig.update_traces(
                                    textposition="outside",
                                    customdata=display_df[["Rank", "Exames Solicitados", "Cirurgias Solicitadas"]],
                                    hovertemplate=(
                                        "%{customdata[0]}º %{y}<br>"
                                        "Exames solicitados: %{customdata[1]}<br>"
                                        "Cirurgias solicitadas: %{customdata[2]}<extra></extra>"
                                    ),
                                )
                            fig.update_yaxes(
                                categoryorder="array",
                                categoryarray=display_df["EtiquetaLocal"].tolist()[::-1],
                            )
                            st.plotly_chart(fig, use_container_width=True)
                            consultorio_pdf_figures.setdefault(sala_label_pdf, []).append((title, fig))

                        with tabs_ind[0]:
                            _render_ind_highlights(top_total_ind)
                            _render_ind_chart(
                                top_total_ind.assign(**{"Total Solicitações": top_total_ind["Total Procedimentos"]}),
                                "Total Solicitações",
                                f"Produtividade no consultório {sala_detalhe}",
                            )

                        with tabs_ind[1]:
                            _render_ind_highlights(top_exames_ind)
                            _render_ind_chart(
                                top_exames_ind,
                                "Exames Solicitados",
                                f"Exames solicitados no consultório {sala_detalhe}",
                            )

                        with tabs_ind[2]:
                            _render_ind_highlights(top_cirurgias_ind)
                            _render_ind_chart(
                                top_cirurgias_ind,
                                "Cirurgias Solicitadas",
                                f"Cirurgias solicitadas no consultório {sala_detalhe}",
                            )

                        with tabs_ind[3]:
                            _render_ind_highlights(top_receita_ind)
                            _render_ind_chart(
                                top_receita_ind,
                                "Receita",
                                f"Receita no consultório {sala_detalhe}",
                                is_currency=True,
                            )

                graf1, graf2 = st.columns(2)
                with graf1:
                    by_dia_ind = detalhe_base.groupby("Dia")["Ocupado"].mean().reset_index()
                    by_dia_ind["Taxa de Ocupação (%)"] = (by_dia_ind["Ocupado"] * 100).round(1)
                    fig_ind_dia = px.bar(by_dia_ind, x="Dia", y="Taxa de Ocupação (%)",
                                         title=f"Ocupação por dia - {sala_detalhe}", text="Taxa de Ocupação (%)")
                    fig_ind_dia.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                    fig_ind_dia.update_yaxes(range=[0, 100])
                    st.plotly_chart(fig_ind_dia, use_container_width=True)
                    consultorio_pdf_figures.setdefault(sala_label_pdf, []).append(
                        (
                            fig_ind_dia.layout.title.text
                            or f"Ocupação por dia - {sala_detalhe}",
                            fig_ind_dia,
                        )
                    )

                with graf2:
                    by_turno_ind = detalhe_base.groupby("Turno")["Ocupado"].mean().reset_index()
                    by_turno_ind["Taxa de Ocupação (%)"] = (by_turno_ind["Ocupado"] * 100).round(1)
                    fig_ind_turno = px.bar(by_turno_ind, x="Turno", y="Taxa de Ocupação (%)",
                                           title=f"Ocupação por turno - {sala_detalhe}", text="Taxa de Ocupação (%)")
                    fig_ind_turno.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                    fig_ind_turno.update_yaxes(range=[0, 100])
                    st.plotly_chart(fig_ind_turno, use_container_width=True)
                    consultorio_pdf_figures.setdefault(sala_label_pdf, []).append(
                        (
                            fig_ind_turno.layout.title.text
                            or f"Ocupação por turno - {sala_detalhe}",
                            fig_ind_turno,
                        )
                    )

                top_med_ind = (
                    top_total_ind if not top_total_ind.empty else pd.DataFrame(columns=["EtiquetaLocal", "Total Procedimentos"])
                )
                if not top_med_ind.empty:
                    top_med_ind_display = top_med_ind.copy()
                    top_med_ind_display["Total Solicitações"] = top_med_ind_display["Total Procedimentos"]
                    fig_top_ind = px.bar(
                        top_med_ind_display,
                        x="Total Solicitações",
                        y="EtiquetaLocal",
                        orientation="h",
                        title=f"Produtividade no consultório {sala_detalhe}",
                        text="Total Solicitações",
                    )
                    fig_top_ind.update_traces(
                        textposition="outside",
                        customdata=top_med_ind_display[["Rank", "Exames Solicitados", "Cirurgias Solicitadas"]],
                        hovertemplate=(
                            "%{customdata[0]}º %{y}<br>"
                            "Exames solicitados: %{customdata[1]}<br>"
                            "Cirurgias solicitadas: %{customdata[2]}<extra></extra>"
                        ),
                    )
                    fig_top_ind.update_yaxes(
                        categoryorder="array",
                        categoryarray=top_med_ind_display["EtiquetaLocal"].tolist()[::-1],
                    )
                    st.plotly_chart(
                        fig_top_ind,
                        use_container_width=True,
                        key=f"consultorio_prod_{sala_norm}",
                    )
                    consultorio_pdf_figures.setdefault(sala_label_pdf, []).append(
                        (
                            fig_top_ind.layout.title.text
                            or f"Produtividade no consultório {sala_detalhe}",
                            fig_top_ind,
                        )
                    )

                if not top_receita_ind.empty and top_receita_ind["Receita"].sum() > 0:
                    top_receita_display = top_receita_ind.copy()
                    top_receita_display["Receita Formatada"] = top_receita_display["Receita"].apply(
                        format_currency_value
                    )
                    fig_top_receita = px.bar(
                        top_receita_display,
                        x="Receita",
                        y="EtiquetaLocal",
                        orientation="h",
                        title=f"Receita no consultório {sala_detalhe}",
                        text="Receita Formatada",
                    )
                    fig_top_receita.update_traces(
                        textposition="outside",
                        customdata=top_receita_display[["Rank", "Total Procedimentos"]],
                        hovertemplate=(
                            "%{customdata[0]}º %{y}<br>"
                            "Receita: %{text}<br>"
                            "Total de procedimentos: %{customdata[1]}<extra></extra>"
                        ),
                    )
                    fig_top_receita.update_yaxes(
                        categoryorder="array",
                        categoryarray=top_receita_display["EtiquetaLocal"].tolist()[::-1],
                    )
                    st.plotly_chart(
                        fig_top_receita,
                        use_container_width=True,
                        key=f"consultorio_receita_{sala_norm}",
                    )
                    consultorio_pdf_figures.setdefault(sala_label_pdf, []).append(
                        (
                            fig_top_receita.layout.title.text
                            or f"Receita no consultório {sala_detalhe}",
                            fig_top_receita,
                        )
                    )

# ---------- Integração das abas MÉDICOS (1, 2, 3...) ----------

med_enriched = pd.DataFrame()
consultorio_medico_agg = pd.DataFrame()
consultorio_totais = pd.DataFrame()
consultorio_planos = pd.DataFrame()
medicos_warning = None

if med_df.empty:
    medicos_warning = "Não foram encontradas abas de **MÉDICOS** no arquivo. Os indicadores de plano/aluguel ficarão ocultos."
else:
    # Enriquecer com turnos utilizados
    usos = fdf_base.groupby("Médico").size().reset_index(name="Turnos Utilizados")
    med_enriched = med_df.merge(usos, on="Médico", how="left")

    if "Consultório" in med_enriched.columns and "Médico" in med_enriched.columns:
        med_consult = med_enriched.copy()
        med_consult["Consultório"] = med_consult["Consultório"].apply(
            lambda v: v if pd.isna(v) else format_consultorio_label(v)
        )
        med_consult = med_consult.dropna(subset=["Consultório"])
        med_consult = med_consult[med_consult["Consultório"].astype(str).str.strip() != ""]

        if (
            sel_salas
            and "Consultório" in med_consult.columns
            and len(sel_salas) != len(salas)
        ):
            med_consult = med_consult[med_consult["Consultório"].isin(sel_salas)]
        if sel_medicos:
            med_consult = med_consult[med_consult["Médico"].isin(sel_medicos)]

        if not med_consult.empty:
            def _sum_ignore_missing(series: pd.Series):
                non_null = series.dropna()
                if non_null.empty:
                    return np.nan
                return non_null.sum()

            group_cols = ["Consultório", "Médico"]
            agg_dict = {"Profissionais": ("Médico", "nunique")}
            if "Valor Aluguel" in med_consult.columns:
                agg_dict["Valor Aluguel Total"] = ("Valor Aluguel", _sum_ignore_missing)
            consultorio_medico_agg = (
                med_consult.groupby(group_cols)
                .agg(**agg_dict)
                .reset_index()
            )

            if "Planos" in med_consult.columns:
                consultorio_planos = med_consult.copy()
                consultorio_planos["Planos"] = (
                    consultorio_planos["Planos"].fillna("Não informado").astype(str).str.strip()
                )
                consultorio_planos = (
                    consultorio_planos.groupby(["Consultório", "Planos"])["Médico"]
                    .nunique()
                    .reset_index(name="Profissionais")
                )

            agg_totais = {"Profissionais": ("Médico", "nunique")}
            if "Valor Aluguel" in med_consult.columns:
                agg_totais["Valor Aluguel Total"] = ("Valor Aluguel", _sum_ignore_missing)
            consultorio_totais = (
                med_consult.groupby("Consultório")
                .agg(**agg_totais)
                .reset_index()
            )
            if "Valor Aluguel Total" in consultorio_totais.columns:
                soma_geral_aluguel = consultorio_totais["Valor Aluguel Total"].sum(
                    min_count=1
                )
                if pd.notna(soma_geral_aluguel) and soma_geral_aluguel != 0:
                    consultorio_totais["Participação do Aluguel (%)"] = (
                        consultorio_totais["Valor Aluguel Total"]
                        / soma_geral_aluguel
                        * 100
                    )
                else:
                    consultorio_totais["Participação do Aluguel (%)"] = pd.NA
            if "Valor Aluguel Total" in consultorio_totais.columns:
                consultorio_totais = consultorio_totais.sort_values(
                    ["Valor Aluguel Total", "Profissionais"],
                    ascending=[False, False],
                    na_position="last",
                )

            if "Valor Aluguel Total" in consultorio_medico_agg.columns:
                share_map = consultorio_totais.set_index("Consultório").get(
                    "Participação do Aluguel (%)"
                )
                if share_map is not None:
                    consultorio_medico_agg["Participação do Aluguel (%)"] = (
                        consultorio_medico_agg["Consultório"].map(share_map)
                    )

ranking_para_pdf = ranking_prod_total.copy()
if not ranking_para_pdf.empty:
    if sel_salas:
        ranking_para_pdf = ranking_para_pdf[ranking_para_pdf["Consultório"].isin(sel_salas)]
    if sel_medicos:
        ranking_para_pdf = ranking_para_pdf[ranking_para_pdf["Profissional"].isin(sel_medicos)]

consultorios_pdf_data: Dict[str, Dict[str, object]] = {}
consultorio_alvo = sel_salas or salas
if consultorio_alvo:
    ranking_para_pdf_norm = ranking_para_pdf.copy()
    if not ranking_para_pdf_norm.empty and "Consultório" in ranking_para_pdf_norm.columns:
        ranking_para_pdf_norm["_SalaLabel"] = ranking_para_pdf_norm["Consultório"].apply(
            format_consultorio_label
        )

    def _normalize_sala_column(df_source: pd.DataFrame) -> pd.Series:
        if df_source is None or df_source.empty or "Sala" not in df_source.columns:
            return pd.Series(dtype=str)
        return df_source["Sala"].apply(
            lambda value: format_consultorio_label(value) if pd.notna(value) else ""
        )

    fdf_base_norm = _normalize_sala_column(fdf_base)
    fdf_norm = _normalize_sala_column(fdf)

    for sala_nome in consultorio_alvo:
        sala_label = format_consultorio_label(sala_nome)
        entry: Dict[str, object] = {}
        metrics_map: Dict[str, object] = {}

        if not fdf_base.empty and not fdf_base_norm.empty:
            sala_base_df = fdf_base.loc[fdf_base_norm == sala_label]
            if not sala_base_df.empty:
                sala_kpi = OccupancyAnalyzer.compute_basic_metrics(sala_base_df)
                metrics_map["Taxa de ocupação"] = f"{sala_kpi.taxa_ocupacao:.1f}%"
                metrics_map["Slots ocupados"] = sala_kpi.slots_ocupados
                metrics_map["Slots livres"] = sala_kpi.slots_livres
                metrics_map["Total de slots"] = sala_kpi.total_slots
                if sala_kpi.medicos_distintos:
                    metrics_map["Médicos distintos"] = sala_kpi.medicos_distintos

        sala_ranking = pd.DataFrame()
        if not ranking_para_pdf_norm.empty and "_SalaLabel" in ranking_para_pdf_norm.columns:
            sala_ranking = ranking_para_pdf_norm.loc[
                ranking_para_pdf_norm["_SalaLabel"] == sala_label
            ]

        if not sala_ranking.empty:
            total_proced = (
                pd.to_numeric(sala_ranking.get("Total Procedimentos"), errors="coerce")
                .fillna(0)
                .sum()
            )
            if total_proced > 0:
                metrics_map["Total procedimentos"] = int(total_proced)

            if "Receita" in sala_ranking.columns:
                receita_total = (
                    pd.to_numeric(sala_ranking["Receita"], errors="coerce")
                    .fillna(0)
                    .sum()
                )
                if receita_total > 0:
                    metrics_map["Receita total"] = format_currency_value(receita_total)

            sort_columns = []
            ascending_flags = []
            if "Total Procedimentos" in sala_ranking.columns:
                sort_columns.append("Total Procedimentos")
                ascending_flags.append(False)
            if "Receita" in sala_ranking.columns:
                sort_columns.append("Receita")
                ascending_flags.append(False)
            if sort_columns:
                sala_sorted = sala_ranking.sort_values(
                    sort_columns,
                    ascending=ascending_flags,
                    na_position="last",
                )
            else:
                sala_sorted = sala_ranking.copy()

            top_records = []
            for _, row in sala_sorted.head(8).iterrows():
                registro = {
                    "Profissional": row.get("Profissional") or "Não informado",
                    "Especialidade": row.get("Especialidade") or "Não informada",
                }
                col_map = {
                    "Total Procedimentos": "Procedimentos",
                    "Exames Solicitados": "Exames",
                    "Cirurgias Solicitadas": "Cirurgias",
                    "Receita": "Receita",
                }
                for origem, destino in col_map.items():
                    if origem in sala_sorted.columns:
                        registro[destino] = row.get(origem)
                top_records.append(registro)

            if top_records:
                entry["top_profissionais"] = pd.DataFrame(top_records)

        if metrics_map:
            entry["metrics"] = metrics_map

        if not fdf.empty and not fdf_norm.empty:
            sala_agenda_df = fdf.loc[fdf_norm == sala_label]
            if not sala_agenda_df.empty:
                group_columns = [
                    col for col in ["Dia", "Turno"] if col in sala_agenda_df.columns
                ]
                if group_columns:
                    agenda_work = sala_agenda_df.copy()

                    agg_spec = {"Total Slots": ("Sala", "size")}
                    if "Ocupado" in agenda_work.columns:
                        agg_spec["Slots Ocupados"] = (
                            "Ocupado",
                            lambda s: int(s.fillna(False).astype(bool).sum()),
                        )
                    if "Médico" in agenda_work.columns:
                        agg_spec["Médicos Ativos"] = ("Médico", pd.Series.nunique)

                    agenda_summary = (
                        agenda_work.groupby(group_columns, observed=False)
                        .agg(**agg_spec)
                        .reset_index()
                    )

                    agg_columns = [
                        column
                        for column in agg_spec.keys()
                        if column in agenda_summary.columns
                    ]
                    for column in agg_columns:
                        agenda_summary[column] = (
                            pd.to_numeric(agenda_summary[column], errors="coerce")
                            .fillna(0)
                            .astype("Int64")
                        )

                    if "Dia" in agenda_summary.columns:
                        agenda_summary["Dia"] = agenda_summary["Dia"].apply(
                            lambda value: value.strftime("%d/%m/%Y")
                            if isinstance(value, (pd.Timestamp, datetime))
                            else str(value)
                        )

                    agenda_summary = agenda_summary.sort_values(group_columns)
                    entry["agenda_resumo"] = agenda_summary.head(12)

        if entry:
            consultorios_pdf_data[sala_label] = entry

pdf_builder = DashboardPDFBuilder(
    data_source=fonte,
    summary_metrics=summary_metrics,
    ranking_df=ranking_para_pdf,
    med_df=med_enriched if not med_df.empty else pd.DataFrame(),
    agenda_df=fdf,
    overview_timeseries=overview_timeseries,
    top_medicos_turnos=top_medicos_turnos,
    consultorios_data=consultorios_pdf_data,
    ranking_limits={
        "total": st.session_state.get("ranking_produtividade_top", 10),
        "exames": st.session_state.get("ranking_produtividade_top", 10),
        "cirurgias": st.session_state.get("ranking_produtividade_top", 10),
        "receita": st.session_state.get("ranking_produtividade_top", 10),
    },
    overview_figures=overview_pdf_figures,
    ranking_figures=ranking_pdf_figures,
    consultorio_figures=consultorio_pdf_figures,
    planos_figures=planos_pdf_figures,
)
pdf_bytes = pdf_builder.build()

with sidebar_pdf_container:
    st.markdown('<div class="sidebar-download-container">', unsafe_allow_html=True)
    st.download_button(
        "📄 Baixar relatório completo (PDF)",
        data=pdf_bytes,
        file_name="dashboard_consultorios.pdf",
        mime="application/pdf",
    )
    st.markdown("</div>", unsafe_allow_html=True)

if selected_section == "💼 Planos & Aluguel":
    if medicos_warning:
        st.warning(medicos_warning)
    else:
        with section_block(
            "💼 Indicador: PLANOS × Aluguel × Profissionais",
            description="Integra convênios, valores de aluguel e atuação por consultório para orientar decisões comerciais.",
            anchor="planos",
        ):

            # KPIs deste bloco
            tot_prof = med_enriched["Médico"].nunique()
            categorias_planos = med_enriched["Planos"].nunique() if "Planos" in med_enriched.columns else 0
            total_consultorios = (
                med_enriched["Consultório"].nunique()
                if "Consultório" in med_enriched.columns
                else 0
            )
            valor_total_aluguel = (
                med_enriched["Valor Aluguel"].sum(min_count=1)
                if "Valor Aluguel" in med_enriched.columns
                else pd.NA
            )
            cpa, cpb, cpc, cpd, cpe = st.columns(5)
            cpa.metric("Profissionais (total)", tot_prof)
            cpb.metric("Consultórios (total)", total_consultorios if total_consultorios else "—")
            cpc.metric("Categorias em PLANOS", categorias_planos)
            if "Valor Aluguel" in med_enriched.columns:
                media_valor = med_enriched["Valor Aluguel"].dropna().mean()
                valor_total_formatado = (
                    f"{valor_total_aluguel:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    if pd.notna(valor_total_aluguel)
                    else "—"
                )
                cpd.metric("Valor total de aluguel (R$)", valor_total_formatado)
                media_formatada = (
                    f"{media_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    if pd.notna(media_valor)
                    else "—"
                )
                cpe.metric("Valor médio de aluguel (R$)", media_formatada)
            else:
                cpd.metric("Valor total de aluguel (R$)", "—")
                cpe.metric("Valor médio de aluguel (R$)", "—")

            st.markdown("#### Receita ao longo do tempo (produtividade)")
            if produtividade_base.empty or "Receita" not in produtividade_base.columns:
                st.info(
                    "Sem dados de produtividade com valores de receita para gerar a linha do tempo financeira."
                )
            else:
                receita_timeline = produtividade_base.copy()
                receita_timeline["Receita"] = pd.to_numeric(
                    receita_timeline["Receita"], errors="coerce"
                ).fillna(0)

                if sel_salas and "Consultório" in receita_timeline.columns:
                    receita_timeline = receita_timeline[
                        receita_timeline["Consultório"].isin(sel_salas)
                    ]

                if "Data" in receita_timeline.columns:
                    receita_timeline["Data"] = pd.to_datetime(
                        receita_timeline["Data"], errors="coerce"
                    )
                    receita_timeline["PeríodoData"] = (
                        receita_timeline["Data"].dt.to_period("M").dt.to_timestamp()
                    )
                elif "Período" in receita_timeline.columns:
                    receita_timeline["PeríodoData"] = pd.to_datetime(
                        receita_timeline["Período"], errors="coerce"
                    ).dt.to_period("M").dt.to_timestamp()
                else:
                    receita_timeline["PeríodoData"] = pd.NaT

                receita_timeline = receita_timeline.dropna(subset=["PeríodoData"])
                if receita_timeline.empty:
                    st.info(
                        "Inclua datas ou períodos nas abas de produtividade para visualizar a evolução da receita."
                    )
                else:
                    receita_timeline["Ano"] = receita_timeline["PeríodoData"].dt.year
                    anos_disponiveis = (
                        receita_timeline["Ano"].dropna().unique().astype(int).tolist()
                    )
                    anos_disponiveis.sort()
                    default_anos = [
                        ano for ano in [2025, 2026] if ano in anos_disponiveis
                    ] or anos_disponiveis

                    anos_selecionados = st.multiselect(
                        "Ano(s) para a linha do tempo",
                        anos_disponiveis,
                        default=default_anos,
                        help="Filtra a linha do tempo de receita para 2025 ou 2026.",
                    )

                    if anos_selecionados:
                        receita_timeline = receita_timeline[
                            receita_timeline["Ano"].isin(anos_selecionados)
                        ]

                    if receita_timeline.empty:
                        st.info("Nenhum registro de receita para os anos selecionados.")
                    else:
                        group_cols = ["PeríodoData"]
                        if "Consultório" in receita_timeline.columns:
                            group_cols.append("Consultório")

                        receita_agg = (
                            receita_timeline.groupby(group_cols)["Receita"]
                            .sum()
                            .reset_index()
                        )
                        receita_agg = receita_agg.sort_values("PeríodoData")
                        receita_agg["Período"] = receita_agg["PeríodoData"].dt.strftime(
                            "%b/%Y"
                        )

                        if "Consultório" in receita_agg.columns:
                            fig_receita = px.bar(
                                receita_agg,
                                x="Período",
                                y="Receita",
                                color="Consultório",
                                barmode="stack",
                                title="Receita por período (soma da produtividade)",
                                labels={
                                    "Receita": "Receita (R$)",
                                    "Período": "Mês/Ano",
                                },
                            )
                        else:
                            fig_receita = px.line(
                                receita_agg,
                                x="Período",
                                y="Receita",
                                markers=True,
                                title="Receita por período (soma da produtividade)",
                                labels={
                                    "Receita": "Receita (R$)",
                                    "Período": "Mês/Ano",
                                },
                            )

                        fig_receita.update_yaxes(tickprefix="R$ ")
                        st.plotly_chart(fig_receita, use_container_width=True)
                        planos_pdf_figures.append(
                            ("Receita por período (produtividade)", fig_receita)
                        )

            g1, g2 = st.columns(2)
            with g1:
                if "Planos" in med_enriched.columns:
                    cont = med_enriched.groupby("Planos")["Médico"].nunique().reset_index(name="Profissionais")
                    fig7 = px.bar(cont, x="Planos", y="Profissionais", title="Profissionais por PLANOS", text="Profissionais")
                    fig7.update_traces(textposition="outside")
                    st.plotly_chart(fig7, use_container_width=True)
                    planos_pdf_figures.append(("Profissionais por PLANOS", fig7))
                else:
                    st.info("Coluna PLANOS não encontrada.")

            with g2:
                if "Valor Aluguel" in med_enriched.columns and "Planos" in med_enriched.columns:
                    avgv = med_enriched.groupby("Planos")["Valor Aluguel"].mean().reset_index(name="Valor médio (R$)")
                    avgv["Valor médio (R$)"] = avgv["Valor médio (R$)"].round(2)
                    fig8 = px.bar(avgv, x="Planos", y="Valor médio (R$)", title="Valor médio de aluguel por PLANOS", text="Valor médio (R$)")
                    fig8.update_traces(texttemplate="R$ %{y:.2f}", textposition="outside")
                    st.plotly_chart(fig8, use_container_width=True)
                    planos_pdf_figures.append(("Valor médio de aluguel por PLANOS", fig8))
                else:
                    st.info("Inclua as colunas PLANOS e Valor Aluguel.")

            if "Valor Aluguel" in med_enriched.columns:
                st.markdown("##### Distribuição de profissionais por faixa de aluguel × PLANOS")
                bins = [0,500,1000,1500,2000,3000,9999999]
                labels = ["até 500","501–1000","1001–1500","1501–2000","2001–3000","3000+"]
                med_enriched["Faixa Aluguel"] = pd.cut(med_enriched["Valor Aluguel"], bins=bins, labels=labels, include_lowest=True)
                dist = med_enriched.groupby(["Planos","Faixa Aluguel"])["Médico"].nunique().reset_index(name="Profissionais")
                fig9 = px.bar(dist, x="Faixa Aluguel", y="Profissionais", color="Planos", barmode="group",
                              title="Profissionais por faixa de aluguel × PLANOS", text="Profissionais")
                fig9.update_traces(textposition="outside")
                st.plotly_chart(fig9, use_container_width=True)
                planos_pdf_figures.append(("Profissionais por faixa de aluguel × PLANOS", fig9))

            if "Especialidade" in med_enriched.columns and "Valor Aluguel" in med_enriched.columns:
                esp_avg = med_enriched.groupby("Especialidade")["Valor Aluguel"].mean().reset_index(name="Valor médio (R$)").sort_values("Valor médio (R$)", ascending=False)
                fig10 = px.bar(
                    esp_avg,
                    x="Valor médio (R$)",
                    y="Especialidade",
                    orientation="h",
                    title="Valor médio de aluguel por especialidade",
                    text="Valor médio (R$)",
                )
                fig10.update_traces(texttemplate="R$ %{x:.2f}", textposition="outside")
                st.plotly_chart(fig10, use_container_width=True)
                planos_pdf_figures.append(("Valor médio de aluguel por especialidade", fig10))
            else:
                st.info("Inclua 'Especialidade' e 'Valor Aluguel'.")

            if "Planos" in med_enriched.columns and "Especialidade" in med_enriched.columns:
                plano_esp = med_enriched.groupby(["Especialidade", "Planos"])["Médico"].nunique().reset_index(name="Profissionais")
                fig11 = px.bar(
                    plano_esp,
                    x="Especialidade",
                    y="Profissionais",
                    color="Planos",
                    barmode="group",
                    title="Profissionais por especialidade × PLANOS",
                    text="Profissionais",
                )
                fig11.update_traces(textposition="outside")
                st.plotly_chart(fig11, use_container_width=True)
                planos_pdf_figures.append(("Profissionais por especialidade × PLANOS", fig11))
            else:
                st.info("Inclua 'Especialidade' e 'PLANOS'.")

            st.markdown("##### Indicadores por consultório")

            if not consultorio_planos.empty:
                st.markdown("###### Convênios por consultório")
                gp1, gp2 = st.columns((2, 1))
                with gp1:
                    planos_ord = consultorio_planos.sort_values(
                        ["Consultório", "Planos"],
                        ascending=[True, True],
                    )
                    consultorios_ordenados = planos_ord["Consultório"].unique().tolist()
                    if not consultorios_ordenados:
                        st.info("Nenhum consultório disponível para montar os gráficos de convênios.")
                    else:
                        tab_labels = []
                        for nome in consultorios_ordenados:
                            if pd.isna(nome) or not str(nome).strip():
                                tab_labels.append("Não informado")
                            else:
                                tab_labels.append(str(nome))
                        tabs = st.tabs(tab_labels)
                        for tab, consultorio_nome, display_nome in zip(tabs, consultorios_ordenados, tab_labels):
                            with tab:
                                if pd.isna(consultorio_nome):
                                    dados_cons = planos_ord[planos_ord["Consultório"].isna()]
                                else:
                                    dados_cons = planos_ord[planos_ord["Consultório"] == consultorio_nome]
                                fig_cons_planos = px.bar(
                                    dados_cons,
                                    x="Planos",
                                    y="Profissionais",
                                    color="Planos",
                                    title=f"Convênios atendidos no {display_nome}",
                                    text="Profissionais",
                                )
                                fig_cons_planos.update_traces(textposition="outside")
                                fig_cons_planos.update_layout(
                                    xaxis_title="Convênio",
                                    yaxis_title="Profissionais",
                                    showlegend=False,
                                )
                                st.plotly_chart(fig_cons_planos, use_container_width=True)
                                planos_pdf_figures.append(
                                    (
                                        fig_cons_planos.layout.title.text
                                        or f"Convênios atendidos no {display_nome}",
                                        fig_cons_planos,
                                    )
                                )
                with gp2:
                    pivot_planos = (
                        consultorio_planos.pivot_table(
                            index="Consultório",
                            columns="Planos",
                            values="Profissionais",
                            aggfunc="sum",
                            fill_value=0,
                        )
                        .astype(int)
                        .reset_index()
                    )
                    pivot_planos = pivot_planos.sort_values("Consultório")
                    st.dataframe(pivot_planos, use_container_width=True)
            else:
                st.info(
                    "Inclua 'Consultório', 'Planos' e 'Médico' para visualizar a distribuição de convênios por consultório."
                )

            gc1, gc2 = st.columns(2)
            with gc1:
                if not consultorio_totais.empty and "Valor Aluguel Total" in consultorio_totais.columns:
                    consultorio_valores = consultorio_totais.dropna(subset=["Valor Aluguel Total"])
                    if not consultorio_valores.empty:
                        if "Participação do Aluguel (%)" in consultorio_valores.columns:
                            consultorio_valores = consultorio_valores.copy()
                            consultorio_valores["Participação do Aluguel (%)"] = consultorio_valores[
                                "Participação do Aluguel (%)"
                            ].round(1)
                        fig_cons_valor = px.bar(
                            consultorio_valores,
                            x="Consultório",
                            y="Valor Aluguel Total",
                            title="Valor total de aluguel por consultório",
                            text="Valor Aluguel Total",
                        )
                        customdata_cols: List[str] = []
                        if "Participação do Aluguel (%)" in consultorio_valores.columns:
                            customdata_cols.append("Participação do Aluguel (%)")
                        fig_cons_valor.update_traces(
                            texttemplate=(
                                "R$ %{y:,.2f}<br>%{customdata[0]:.1f}%"
                                if customdata_cols
                                else "R$ %{y:,.2f}"
                            ),
                            textposition="outside",
                            customdata=consultorio_valores[customdata_cols]
                            if customdata_cols
                            else None,
                        )
                        fig_cons_valor.update_layout(xaxis_title="Consultório", yaxis_title="Valor total (R$)")
                        st.plotly_chart(fig_cons_valor, use_container_width=True)
                        planos_pdf_figures.append(("Valor total de aluguel por consultório", fig_cons_valor))
                    else:
                        st.info("Nenhum valor de aluguel informado para os consultórios listados.")
                else:
                    st.info("Inclua 'Consultório' e 'Valor Aluguel' para visualizar os totais.")
            with gc2:
                if not consultorio_totais.empty:
                    fig_cons_prof = px.bar(
                        consultorio_totais,
                        x="Consultório",
                        y="Profissionais",
                        title="Profissionais por consultório",
                        text="Profissionais",
                    )
                    fig_cons_prof.update_traces(textposition="outside")
                    st.plotly_chart(fig_cons_prof, use_container_width=True)
                    planos_pdf_figures.append(("Profissionais por consultório", fig_cons_prof))
                else:
                    st.info("Inclua 'Consultório' para visualizar a distribuição de profissionais.")

            if not consultorio_medico_agg.empty:
                st.markdown("##### Tabela por consultório × médico")
                display_cols = [
                    c
                    for c in [
                        "Consultório",
                        "Médico",
                        "Profissionais",
                        "Valor Aluguel Total",
                        "Participação do Aluguel (%)",
                    ]
                    if c in consultorio_medico_agg.columns
                ]
                st.dataframe(
                    consultorio_medico_agg[display_cols].sort_values(
                        ["Consultório", "Médico"], na_position="last"
                    ),
                    use_container_width=True,
                )

            g5, g6 = st.columns(2)
            with g5:
                if "Sala Exclusiva" in med_enriched.columns or "Sala Dividida" in med_enriched.columns:
                    ts = med_enriched.copy()
                    ts["Tipo de Sala"] = None
                    if "Sala Exclusiva" in ts.columns:
                        ts.loc[ts["Sala Exclusiva"].eq("Sim"), "Tipo de Sala"] = "Exclusiva"
                    if "Sala Dividida" in ts.columns:
                        ts.loc[ts["Sala Dividida"].eq("Sim"), "Tipo de Sala"] = ts["Tipo de Sala"].fillna("Dividida")
                    ts = ts.dropna(subset=["Tipo de Sala"])
                    if not ts.empty:
                        dist_ts = ts.groupby("Tipo de Sala")["Médico"].nunique().reset_index(name="Profissionais")
                        fig12 = px.bar(dist_ts, x="Tipo de Sala", y="Profissionais", title="Profissionais por tipo de sala", text="Profissionais")
                        fig12.update_traces(textposition="outside")
                        st.plotly_chart(fig12, use_container_width=True)
                        planos_pdf_figures.append(("Profissionais por tipo de sala", fig12))
                    else:
                        st.info("Sem marcações de sala exclusiva/dividida para analisar.")
                else:
                    st.info("Inclua colunas 'Sala Exclusiva' e/ou 'Sala Dividida'.")

            st.markdown("##### Tabela (Médico × CRM × Especialidade × PLANOS × Valor × Tipo de Sala × Turnos)")
            cols_show = [
                c
                for c in [
                    "Médico",
                    "Consultório",
                    "CRM",
                    "Especialidade",
                    "Planos",
                    "Valor Aluguel",
                    "Sala Exclusiva",
                    "Sala Dividida",
                    "Turnos Utilizados",
                ]
                if c in med_enriched.columns
            ]
            sort_cols = [
                c
                for c in ["Planos", "Consultório", "Especialidade", "Valor Aluguel", "Médico"]
                if c in med_enriched.columns
            ]
            st.dataframe(
                med_enriched[cols_show].sort_values(sort_cols, na_position="last") if sort_cols else med_enriched[cols_show],
                use_container_width=True,
            )

if selected_section == "📋 Agenda":
    # ---------- Detalhamento ----------
    with section_block(
        "📋 Agenda Detalhada (Tabela)",
        description="Visualize, exporte e compartilhe a agenda filtrada em diferentes formatos.",
        anchor="agenda",
    ) as sec:
        sec.dataframe(
            fdf.sort_values(["Sala", "Dia", "Turno"]).reset_index(drop=True)[
                ["Sala", "Dia", "Turno", "Médico"]
            ],
            use_container_width=True,
        )

        csv = fdf.to_csv(index=False).encode("utf-8-sig")
        sec.download_button(
            "⬇️ Baixar dados filtrados (CSV)",
            data=csv,
            file_name="agenda_filtrada.csv",
            mime="text/csv",
        )

        sec.markdown(
            '<div style="text-align: right; margin-top: 2rem;"><a href="#topo">⬆️ Voltar ao topo</a></div>',
            unsafe_allow_html=True,
        )
