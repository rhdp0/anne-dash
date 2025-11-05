import re
from io import BytesIO
from pathlib import Path
from datetime import datetime
import unicodedata

import pandas as pd
import streamlit as st
import plotly.express as px
from fpdf import FPDF

st.set_page_config(page_title="Dashboard Consultórios", layout="wide")

# --- Corporate styling ---
st.markdown("""
<style>
.block-container {padding-top: 1.5rem;}
div[data-testid="stMetricValue"] {color:#0F4C81;}
h1, h2, h3 { color:#1f2a44; }
section[data-testid="stSidebar"] {background-color:#f5f7fb}
.section-title {
    display: block;
    background-color: #eef3fb;
    border-left: 6px solid #0F4C81;
    padding: 0.75rem 1rem;
    margin: 2rem 0 1rem;
    border-radius: 0.5rem;
    font-size: clamp(1.35rem, 1.2rem + 1vw, 1.75rem);
    line-height: 1.4;
    color: #0F1A33;
}
.section-title strong {
    color: inherit;
}
.section-card {
    background-color: #ffffff;
    border: 1px solid rgba(15, 26, 51, 0.08);
    border-radius: 0.75rem;
    padding: 1.5rem;
    margin: 2rem 0;
    box-shadow: 0 10px 30px rgba(15, 26, 51, 0.08);
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div id="topo"></div>', unsafe_allow_html=True)
st.title("🏥 Dashboard de Ocupação dos Consultórios")
st.caption("Lendo somente as abas **CONSULTÓRIO** (ignorando 'OCUPAÇÃO DAS SALAS'). Integra automaticamente TODAS as abas **MÉDICOS** (ex.: 'MÉDICOS 1', 'MÉDICOS 2', 'MÉDICOS 3').")

DEFAULT_PATH = Path("/mnt/data/ESCALA DOS CONSULTORIOS DEFINITIVO.xlsx")

# ---------- Sidebar: Upload ----------
st.sidebar.header("📂 Fonte de Dados")
uploaded = st.sidebar.file_uploader("Envie o Excel (.xlsx)", type=["xlsx"], key="main_xlsx")

def load_excel(file_like):
    try:
        return pd.ExcelFile(file_like)
    except Exception as e:
        st.error(f"Não foi possível abrir o arquivo: {e}")
        return None

excel = None
if uploaded is not None:
    excel = load_excel(uploaded)
    fonte = "Upload do usuário"
elif DEFAULT_PATH.exists():
    excel = load_excel(DEFAULT_PATH)
    fonte = f"Arquivo padrão: {DEFAULT_PATH.name}"
else:
    st.error("Nenhum arquivo encontrado. Envie um Excel com as abas de CONSULTÓRIO.")
    st.stop()

st.sidebar.success(f"Usando dados de: {fonte}")
st.sidebar.markdown("### Navegação")
st.sidebar.markdown(
    """
    <a href="#ranking">🏆 Ranking</a><br>
    <a href="#consultorio">🔍 Consultórios</a><br>
    <a href="#planos">💼 Planos &amp; Aluguel</a><br>
    <a href="#agenda">📋 Agenda</a>
    """,
    unsafe_allow_html=True,
)

# ---------- Utilitários ----------
def _normalize_col(col):
    c = str(col).strip().lower()
    c = (c
         .replace("á","a").replace("ã","a").replace("â","a")
         .replace("é","e").replace("ê","e")
         .replace("í","i").replace("î","i")
         .replace("ó","o").replace("õ","o").replace("ô","o")
         .replace("ú","u").replace("ü","u")
         .replace("ç","c"))
    c = re.sub(r"\s+", " ", c)
    return c

def _to_number(x):
    import numpy as np, re as _re
    if pd.isna(x):
        return np.nan
    txt = str(x)
    txt = _re.sub(r"[^\d,.-]", "", txt)
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt and "." not in txt:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except:
        return pd.NA

def _first_nonempty(series):
    for val in series:
        if pd.isna(val):
            continue
        text = str(val).strip()
        if text:
            return text
    return ""

def _format_consultorio_label(name):
    label = str(name).strip()
    label = re.sub(r"(?i)^produtividade\s*[:\-]*", "", label).strip()
    label = re.sub(r"(?i)consult[óo]rio", "Consultório", label)
    label = re.sub(r"\s+", " ", label).strip(" -_:")
    return label or str(name).strip()


def _sanitize_pdf_text(text: str) -> str:
    """Remove acentuação incompatível com fontes padrão do FPDF."""
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def build_pdf_report(summary_metrics, ranking_df, med_df, agenda_df) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    def _safe_int(value):
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

    pdf.set_font("Helvetica", "B", 16)
    pdf.multi_cell(0, 10, _sanitize_pdf_text("Relatorio Completo - Dashboard de Consultorios"))
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, _sanitize_pdf_text(f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"), ln=1)
    pdf.ln(4)

    if summary_metrics:
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 8, _sanitize_pdf_text("Resumo dos principais indicadores"), ln=1)
        pdf.set_font("Helvetica", "", 11)
        for key, value in summary_metrics.items():
            pdf.multi_cell(0, 6, _sanitize_pdf_text(f"{key}: {value}"))
        pdf.ln(2)

    if ranking_df is not None and not ranking_df.empty:
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 8, _sanitize_pdf_text("Top 10 profissionais por produtividade"), ln=1)
        pdf.set_font("Helvetica", "", 11)
        sort_cols = []
        ascending = []
        for col, asc in [
            ("Total Procedimentos", False),
            ("Cirurgias Solicitadas", False),
            ("Exames Solicitados", False),
            ("Profissional", True),
            ("Consultório", True),
        ]:
            if col in ranking_df.columns:
                sort_cols.append(col)
                ascending.append(asc)
        if sort_cols:
            ranking_sorted = ranking_df.sort_values(sort_cols, ascending=ascending)
        else:
            ranking_sorted = ranking_df
        top_ranking = ranking_sorted.head(10).reset_index(drop=True)
        for idx, row in top_ranking.iterrows():
            prof = row.get("Profissional", "")
            especialidade = row.get("Especialidade", "")
            consultorio = row.get("Consultório", "")
            crm = row.get("CRM", "")
            total = _safe_int(row.get("Total Procedimentos"))
            exames = _safe_int(row.get("Exames Solicitados"))
            cirurgias = _safe_int(row.get("Cirurgias Solicitadas"))

            total_txt = f"Total: {total}" if total is not None else ""
            detalhes = []
            if consultorio:
                detalhes.append(f"Consultorio: {consultorio}")
            if crm and str(crm).strip():
                detalhes.append(f"CRM: {crm}")
            if exames is not None:
                detalhes.append(f"Exames: {exames}")
            if cirurgias is not None:
                detalhes.append(f"Cirurgias: {cirurgias}")
            if total_txt:
                detalhes.insert(0, total_txt)

            titulo = f"{idx + 1}. {prof}" if prof else f"{idx + 1}. Profissional"
            if especialidade and especialidade != "Não informada":
                titulo = f"{titulo} - {especialidade}"

            pdf.multi_cell(0, 6, _sanitize_pdf_text(titulo))
            if detalhes:
                pdf.multi_cell(0, 5, _sanitize_pdf_text(" • ".join(detalhes)))
            pdf.ln(1)
        pdf.ln(2)

    if med_df is not None and not med_df.empty:
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 8, _sanitize_pdf_text("Planos, aluguel e profissionais"), ln=1)
        pdf.set_font("Helvetica", "", 11)
        total_profissionais = (
            med_df["Médico"].nunique() if "Médico" in med_df.columns else len(med_df)
        )
        pdf.multi_cell(0, 6, _sanitize_pdf_text(f"Profissionais analisados: {total_profissionais}"))

        if "Planos" in med_df.columns:
            planos = med_df.copy()
            planos["Planos"] = planos["Planos"].fillna("Nao informado").astype(str).str.strip()
            if "Médico" in planos.columns:
                planos_grouped = planos.groupby("Planos")["Médico"].nunique().reset_index(name="Profissionais")
            else:
                planos_grouped = planos["Planos"].value_counts().reset_index()
                planos_grouped.columns = ["Planos", "Profissionais"]
            planos_grouped = planos_grouped.sort_values("Profissionais", ascending=False)
            pdf.multi_cell(0, 6, _sanitize_pdf_text("Distribuicao por PLANOS:"))
            for _, row in planos_grouped.head(5).iterrows():
                plano_nome = row.get("Planos", "Nao informado")
                qtd = _safe_int(row.get("Profissionais", 0)) or 0
                pdf.multi_cell(0, 5, _sanitize_pdf_text(f"- {plano_nome}: {qtd} profissionais"))

        if "Valor Aluguel" in med_df.columns:
            valores = pd.to_numeric(med_df["Valor Aluguel"], errors="coerce").dropna()
            if not valores.empty:
                media = valores.mean()
                minimo = valores.min()
                maximo = valores.max()
                format_currency = lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                pdf.multi_cell(0, 6, _sanitize_pdf_text("Valores de aluguel (considerando dados disponiveis):"))
                pdf.multi_cell(0, 5, _sanitize_pdf_text(f"- Media: {format_currency(media)}"))
                pdf.multi_cell(0, 5, _sanitize_pdf_text(f"- Minimo: {format_currency(minimo)}"))
                pdf.multi_cell(0, 5, _sanitize_pdf_text(f"- Maximo: {format_currency(maximo)}"))
        pdf.ln(2)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, _sanitize_pdf_text("Agenda filtrada"), ln=1)
    pdf.set_font("Helvetica", "", 11)
    if agenda_df is None or agenda_df.empty:
        pdf.multi_cell(0, 6, _sanitize_pdf_text("Nenhum agendamento encontrado para os filtros atuais."))
    else:
        agenda_cols = [c for c in ["Sala", "Dia", "Turno", "Médico"] if c in agenda_df.columns]
        agenda_view = agenda_df.copy()
        if agenda_cols:
            agenda_view = agenda_view[agenda_cols]
        sort_cols = [c for c in ["Sala", "Dia", "Turno"] if c in agenda_view.columns]
        if sort_cols:
            agenda_view = agenda_view.sort_values(sort_cols)
        pdf.multi_cell(0, 6, _sanitize_pdf_text("Primeiros 30 registros:"))
        for _, row in agenda_view.head(30).iterrows():
            linha = " | ".join(str(row.get(col, "")) for col in agenda_cols)
            pdf.multi_cell(0, 5, _sanitize_pdf_text(linha))

    output = pdf.output(dest="S")
    if isinstance(output, str):
        output_bytes = output.encode("latin-1")
    else:
        output_bytes = bytes(output)
    buffer = BytesIO()
    buffer.write(output_bytes)
    buffer.seek(0)
    return buffer.getvalue()

def detect_header_and_parse(excel, sheet_name):
    for header in [0,1,2,3,4]:
        try:
            df = excel.parse(sheet_name, header=header)
        except Exception:
            continue
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if df.empty:
            continue

        cols_norm = [_normalize_col(c) for c in df.columns]
        col_dia = None; col_manha=None; col_tarde=None

        for i, cn in enumerate(cols_norm):
            if col_dia is None:
                if "dia" in cn or any(d in cn for d in ["segunda","terca","terça","quarta","quinta","sexta","sabado","sábado"]):
                    col_dia = df.columns[i]
            if any(k in cn for k in ["manha","manhã"]): col_manha = df.columns[i]
            if "tarde" in cn: col_tarde = df.columns[i]

        # fallback: primeira coluna contém dias
        if col_dia is None and len(df.columns) >= 1:
            first_col = df.columns[0]
            sample = df[first_col].astype(str).str.lower()
            if sample.str.contains("segunda|terca|terça|quarta|quinta|sexta|sabado|sábado").any():
                col_dia = first_col

        if col_dia is not None and (col_manha is not None or col_tarde is not None):
            use_cols = [c for c in [col_dia, col_manha, col_tarde] if c is not None]
            df = df[use_cols].copy()
            rename = {col_dia:"Dia"}
            if col_manha is not None: rename[col_manha]="Manhã"
            if col_tarde is not None: rename[col_tarde]="Tarde"
            df = df.rename(columns=rename)
            df["Dia"] = df["Dia"].astype(str).str.strip()
            df = df[df["Dia"].str.len()>0]
            return df
    return None

def tidy_from_sheets(excel):
    frames = []
    for sheet in excel.sheet_names:
        s_norm = _normalize_col(sheet)
        if ("consult" in s_norm) and ("ocupa" not in s_norm):
            df = detect_header_and_parse(excel, sheet)
            if df is None or df.empty:
                continue
            df["Dia"] = (df["Dia"].astype(str).str.strip()
                         .str.replace("terca","terça", case=False)
                         .str.replace("sabado","sábado", case=False)
                         .str.capitalize())
            df.insert(0, "Sala", sheet.strip())
            long = df.melt(id_vars=["Sala","Dia"], value_vars=[c for c in ["Manhã","Tarde"] if c in df.columns],
                           var_name="Turno", value_name="Médico")
            long["Médico"] = long["Médico"].astype(str).replace({"nan":"","None":""}).str.strip()
            frames.append(long)
    if not frames:
        return pd.DataFrame(columns=["Sala","Dia","Turno","Médico"])
    full = pd.concat(frames, ignore_index=True)
    full["Dia"] = pd.Categorical(full["Dia"], categories=["Segunda","Terça","Quarta","Quinta","Sexta","Sábado"], ordered=True)
    full["Ocupado"] = full["Médico"].str.len() > 0
    return full

def load_produtividade_from_excel(excel: pd.ExcelFile) -> pd.DataFrame:
    frames = []
    for sheet in excel.sheet_names:
        s_norm = _normalize_col(sheet)
        if "consultorios" in s_norm:
            continue
        if "produtiv" not in s_norm or "consult" not in s_norm:
            continue
        for header in range(0, 6):
            try:
                dfp = excel.parse(sheet, header=header)
            except Exception:
                continue
            dfp = dfp.dropna(how="all").dropna(axis=1, how="all")
            if dfp.empty:
                continue

            rename = {}
            for col in dfp.columns:
                norm = _normalize_col(col)
                if "nome" in norm and "consult" not in norm:
                    rename[col] = "Profissional"
                elif norm == "crm" or "crm" in norm:
                    rename[col] = "CRM"
                elif "especial" in norm:
                    rename[col] = "Especialidade"
                elif "exame" in norm:
                    rename[col] = "Exames Solicitados"
                elif "cirurg" in norm:
                    rename[col] = "Cirurgias Solicitadas"
                elif "consult" in norm and "produtiv" not in norm:
                    rename[col] = "Consultório"

            dfp = dfp.rename(columns=rename)

            if "Profissional" not in dfp.columns:
                continue
            if "Exames Solicitados" not in dfp.columns and "Cirurgias Solicitadas" not in dfp.columns:
                continue

            keep = [c for c in ["Profissional", "CRM", "Especialidade", "Exames Solicitados", "Cirurgias Solicitadas", "Consultório"] if c in dfp.columns]
            dfp = dfp[keep].copy()

            dfp["Profissional"] = dfp["Profissional"].astype(str).str.strip()
            dfp = dfp[dfp["Profissional"].str.len() > 0]
            dfp = dfp[dfp["Profissional"].apply(lambda x: _normalize_col(x) not in {"total", "totais", "subtotal"})]

            if "Consultório" not in dfp.columns:
                dfp["Consultório"] = _format_consultorio_label(sheet)
            else:
                dfp["Consultório"] = (
                    dfp["Consultório"].astype(str).str.strip().replace(
                        r"(?i)^(nan|none|null|na|n/a|sem\s*informac[aã]o|sem\s*dados?)$",
                        "",
                        regex=True,
                    )
                )
                dfp.loc[dfp["Consultório"].eq(""), "Consultório"] = _format_consultorio_label(sheet)
                dfp["Consultório"] = dfp["Consultório"].fillna(_format_consultorio_label(sheet))

            if "Especialidade" in dfp.columns:
                dfp["Especialidade"] = dfp["Especialidade"].astype(str).str.strip()
            else:
                dfp["Especialidade"] = ""

            if "CRM" in dfp.columns:
                dfp["CRM"] = dfp["CRM"].astype(str).str.strip()
            else:
                dfp["CRM"] = ""

            for col in ["Exames Solicitados", "Cirurgias Solicitadas"]:
                if col in dfp.columns:
                    dfp[col] = dfp[col].apply(
                        lambda value: 0
                        if (pd.isna(value) or str(value).strip() == "")
                        else _to_number(value)
                    )
                else:
                    dfp[col] = 0
                dfp[col] = pd.to_numeric(dfp[col], errors="coerce").fillna(0)

            dfp["Consultório"] = dfp["Consultório"].apply(_format_consultorio_label)
            dfp["_SalaNorm"] = dfp["Consultório"].apply(_normalize_col)

            frames.append(dfp)
            break
    if not frames:
        return pd.DataFrame(columns=["Profissional", "CRM", "Especialidade", "Exames Solicitados", "Cirurgias Solicitadas", "Consultório", "_SalaNorm"])
    return pd.concat(frames, ignore_index=True)

df = tidy_from_sheets(excel)
if df.empty:
    st.error("Não foram encontrados dados nas abas 'CONSULTÓRIO'.")
    st.stop()

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

# Base para KPIs (NÃO filtra por médico)
mask_base = (df["Sala"].isin(sel_salas) & df["Dia"].astype(str).isin(sel_dias) & df["Turno"].isin(sel_turnos))
fdf_base = df[mask_base].copy()

# Aplicar filtro de médico apenas onde fizer sentido
if sel_medicos:
    mask_medico = df["Médico"].isin(sel_medicos)
else:
    mask_medico = pd.Series(True, index=df.index)
fdf = df[mask_base & mask_medico].copy()

produtividade_df = load_produtividade_from_excel(excel)
ranking_prod_total = pd.DataFrame()
if not produtividade_df.empty:
    base_prod = produtividade_df.copy()
    base_prod["Especialidade"] = base_prod["Especialidade"].fillna("").astype(str).str.strip()
    base_prod.loc[base_prod["Especialidade"].eq(""), "Especialidade"] = "Não informada"

    agg_map = {
        "Exames Solicitados": "sum",
        "Cirurgias Solicitadas": "sum",
    }
    if "CRM" in base_prod.columns:
        agg_map["CRM"] = _first_nonempty

    ranking_prod_total = (
        base_prod.groupby(["Consultório", "Especialidade", "Profissional"], as_index=False)
        .agg(agg_map)
    )

    if "CRM" not in ranking_prod_total.columns:
        ranking_prod_total["CRM"] = ""

    for col in ["Exames Solicitados", "Cirurgias Solicitadas"]:
        ranking_prod_total[col] = pd.to_numeric(ranking_prod_total[col], errors="coerce").fillna(0)

    ranking_prod_total["Total Procedimentos"] = (
        ranking_prod_total["Exames Solicitados"] + ranking_prod_total["Cirurgias Solicitadas"]
    )

    ranking_prod_total = ranking_prod_total[ranking_prod_total["Total Procedimentos"] > 0]

    for col in ["Exames Solicitados", "Cirurgias Solicitadas", "Total Procedimentos"]:
        ranking_prod_total[col] = ranking_prod_total[col].round().astype(int)

    ranking_prod_total["SalaNorm"] = ranking_prod_total["Consultório"].apply(_normalize_col)
    ranking_prod_total["Etiqueta"] = ranking_prod_total.apply(
        lambda r: (
            f"{r['Profissional']} - {r['Especialidade']} ({r['Consultório']})"
            if r.get("Especialidade") and r.get("Especialidade") != "Não informada"
            else f"{r['Profissional']} ({r['Consultório']})"
        ),
        axis=1,
    )

# ---------- KPIs ----------
total_salas = len(set(sel_salas))
total_slots = len(fdf_base)
ocupados = int(fdf_base["Ocupado"].sum())
tx_ocup = (ocupados / total_slots * 100) if total_slots > 0 else 0
slots_livres = max(total_slots - ocupados, 0)
medicos_distintos = fdf_base.loc[fdf_base["Ocupado"], "Médico"].nunique()

summary_metrics = {
    "Consultórios selecionados": total_salas,
    "Slots analisados": total_slots,
    "Slots livres": slots_livres,
    "Slots ocupados": ocupados,
    "Taxa de ocupação": f"{tx_ocup:.1f}%",
    "Médicos distintos": medicos_distintos,
    "Dias filtrados": ", ".join(sel_dias) if sel_dias else "Todos",
    "Turnos filtrados": ", ".join(sel_turnos) if sel_turnos else "Todos",
}
if sel_medicos:
    summary_metrics["Médicos no filtro"] = len(sel_medicos)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Consultórios selecionados", total_salas)
c2.metric("Slots (dia x turno x sala)", total_slots)
c3.metric("Slots livres", slots_livres)
c4.metric("Ocupados", ocupados)

kc1, kc2 = st.columns(2)
kc1.metric("Taxa de ocupação", f"{tx_ocup:.1f}%")
kc2.metric("Médicos distintos (no filtro de sala/dia/turno)", medicos_distintos)

# ---------- Gráficos de ocupação (sem heatmap) com porcentagens nas barras ----------
colA, colB = st.columns(2)
with colA:
    by_sala = fdf_base.groupby("Sala")["Ocupado"].mean().reset_index()
    by_sala["Taxa de Ocupação (%)"] = (by_sala["Ocupado"]*100).round(1)
    fig1 = px.bar(by_sala, x="Sala", y="Taxa de Ocupação (%)", title="Ocupação por Consultório (%)", text="Taxa de Ocupação (%)")
    fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig1.update_yaxes(range=[0,100])
    st.plotly_chart(fig1, use_container_width=True)

with colB:
    by_dia = fdf_base.groupby("Dia")["Ocupado"].mean().reset_index()
    by_dia["Taxa de Ocupação (%)"] = (by_dia["Ocupado"]*100).round(1)
    fig2 = px.bar(by_dia, x="Dia", y="Taxa de Ocupação (%)", title="Ocupação por Dia da Semana (%)", text="Taxa de Ocupação (%)")
    fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig2.update_yaxes(range=[0,100])
    st.plotly_chart(fig2, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    by_turno = fdf_base.groupby("Turno")["Ocupado"].mean().reset_index()
    by_turno["Taxa de Ocupação (%)"] = (by_turno["Ocupado"]*100).round(1)
    fig3 = px.bar(by_turno, x="Turno", y="Taxa de Ocupação (%)", title="Ocupação por Turno (%)", text="Taxa de Ocupação (%)")
    fig3.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig3.update_yaxes(range=[0,100])
    st.plotly_chart(fig3, use_container_width=True)

with colD:
    top_med = (fdf[fdf["Ocupado"]]
               .groupby("Médico")
               .size()
               .reset_index(name="Turnos Utilizados")
               .sort_values("Turnos Utilizados", ascending=False)
               .head(15))
    if not top_med.empty:
        fig4 = px.bar(top_med, x="Turnos Utilizados", y="Médico", orientation="h", title="Top Médicos por Nº de Turnos", text="Turnos Utilizados")
        fig4.update_traces(textposition="outside")
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Sem médicos ocupando slots nos filtros atuais.")

# ---------- Ranking de produtividade dos médicos ----------
st.markdown('<div id="ranking"></div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">🏆 Ranking de produtividade dos médicos</h2>', unsafe_allow_html=True)

    if ranking_prod_total.empty:
        st.info("Sem dados nas abas de produtividade para gerar o ranking geral.")
    else:
        ranking = ranking_prod_total.copy()
        ranking = ranking.sort_values(
            ["Total Procedimentos", "Cirurgias Solicitadas", "Exames Solicitados", "Profissional", "Consultório"],
            ascending=[False, False, False, True, True],
        ).reset_index(drop=True)
        ranking.insert(0, "Rank", range(1, len(ranking) + 1))

        top_n_default = min(len(ranking), 10) if len(ranking) else 1
        top_n = st.slider(
            "Quantidade de profissionais no ranking",
            min_value=1,
            max_value=len(ranking),
            value=top_n_default,
            key="ranking_produtividade_top",
        )
        top_view = ranking.head(top_n)

        destaque_registros = top_view.head(3).to_dict("records")
        if destaque_registros:
            destaque_cols = st.columns(len(destaque_registros))
            for col, row in zip(destaque_cols, destaque_registros):
                total = int(row.get("Total Procedimentos", 0))
                exames = int(row.get("Exames Solicitados", 0))
                cirurgias = int(row.get("Cirurgias Solicitadas", 0))
                profissional = row.get("Profissional", "")
                especialidade = row.get("Especialidade", "")
                consultorio = row.get("Consultório", "")
                crm = row.get("CRM", "")
                rank = row.get("Rank", "-")

                titulo = f"{rank}º {profissional}" if profissional else f"{rank}º Profissional"
                if especialidade and especialidade != "Não informada":
                    titulo = f"{titulo} - {especialidade}"

                info_parts = []
                if consultorio:
                    info_parts.append(consultorio)
                if crm:
                    info_parts.append(f"CRM {crm}")
                info_parts.append(f"Exames: {exames}")
                info_parts.append(f"Cirurgias: {cirurgias}")

                col.metric(
                    titulo,
                    f"{total} Solicitações",
                    " • ".join(info_parts),
                )

        if not top_view.empty:
            top_view_display = top_view.copy()
            top_view_display["Total Solicitações"] = top_view_display["Total Procedimentos"]
            fig_rank = px.bar(
                top_view_display,
                x="Total Solicitações",
                y="Etiqueta",
                orientation="h",
                color="Total Solicitações",
                color_continuous_scale="Blues",
                title="Top profissionais por produtividade",
                text="Total Solicitações",
            )
            fig_rank.update_layout(coloraxis_showscale=False)
            fig_rank.update_traces(
                texttemplate="%{text}",
                textposition="outside",
                customdata=top_view_display[["Rank", "Consultório", "Especialidade", "Exames Solicitados", "Cirurgias Solicitadas"]],
                hovertemplate=(
                    "%{customdata[0]}º %{y}<br>"
                    "Consultório: %{customdata[1]}<br>"
                    "Especialidade: %{customdata[2]}<br>"
                    "Exames solicitados: %{customdata[3]}<br>"
                    "Cirurgias solicitadas: %{customdata[4]}<extra></extra>"
                ),
            )
            fig_rank.update_yaxes(
                categoryorder="array",
                categoryarray=top_view_display["Etiqueta"].tolist()[::-1],
            )
            st.plotly_chart(fig_rank, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ---------- Visão individual por consultório ----------
st.markdown('<div id="consultorio"></div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-title">🔍 Indicadores individuais por consultório</h2>', unsafe_allow_html=True)

    salas_disponiveis = sorted(df["Sala"].dropna().unique().tolist())
    if not salas_disponiveis:
        st.info("Não há consultórios disponíveis para detalhar.")
    else:
        sala_detalhe = st.selectbox("Escolha um consultório para detalhar", salas_disponiveis, key="detalhe_sala")

        mask_sala_base = ((df["Sala"] == sala_detalhe)
                          & df["Dia"].astype(str).isin(sel_dias)
                          & df["Turno"].isin(sel_turnos))
        mask_sala = mask_sala_base & mask_medico

        detalhe_base = df[mask_sala_base].copy()
        detalhe_df = df[mask_sala].copy()

        if detalhe_base.empty:
            st.info("Sem dados para o consultório selecionado com os filtros atuais de dia/turno.")
        else:
            slots_totais = len(detalhe_base)
            ocupados_ind = int(detalhe_base["Ocupado"].sum())
            livres_ind = max(slots_totais - ocupados_ind, 0)
            taxa_ind = (ocupados_ind / slots_totais * 100) if slots_totais > 0 else 0
            medicos_ind = detalhe_base.loc[detalhe_base["Ocupado"], "Médico"].nunique()

            ic1, ic2, ic3, ic4 = st.columns(4)
            ic1.metric("Consultório", sala_detalhe)
            ic2.metric("Slots do consultório", slots_totais)
            ic3.metric("Slots livres", livres_ind)
            ic4.metric("Ocupados", ocupados_ind)

            ic5, ic6 = st.columns(2)
            ic5.metric("Taxa de ocupação do consultório", f"{taxa_ind:.1f}%")
            ic6.metric("Médicos distintos no consultório", medicos_ind)

            ranking_ind = pd.DataFrame()
            sala_norm = _normalize_col(sala_detalhe)
            if ranking_prod_total.empty:
                st.info("Sem dados de produtividade carregados para detalhar este consultório.")
            else:
                ranking_ind = ranking_prod_total[ranking_prod_total["SalaNorm"] == sala_norm].copy()
                if ranking_ind.empty:
                    st.info("Sem registros de produtividade para o consultório selecionado.")
                else:
                    ranking_ind = ranking_ind.sort_values(
                        ["Total Procedimentos", "Cirurgias Solicitadas", "Exames Solicitados", "Profissional"],
                        ascending=[False, False, False, True],
                    ).reset_index(drop=True)
                    ranking_ind.insert(0, "Rank", range(1, len(ranking_ind) + 1))
                    ranking_ind["EtiquetaLocal"] = ranking_ind.apply(
                        lambda r: f"{r['Profissional']} - {r['Especialidade']}"
                        if r.get("Especialidade") and r.get("Especialidade") != "Não informada"
                        else r.get("Profissional", ""),
                        axis=1,
                    )

                    destaque_ind = ranking_ind.head(3).to_dict("records")
                    if destaque_ind:
                        st.markdown("#### Destaques de produtividade no consultório")
                        destaque_cols_ind = st.columns(len(destaque_ind))
                        for col, row in zip(destaque_cols_ind, destaque_ind):
                            total = int(row.get("Total Procedimentos", 0))
                            exames = int(row.get("Exames Solicitados", 0))
                            cirurgias = int(row.get("Cirurgias Solicitadas", 0))
                            profissional = row.get("Profissional", "")
                            especialidade = row.get("Especialidade", "")
                            crm = row.get("CRM", "")
                            rank = row.get("Rank", "-")

                            titulo_local = f"{rank}º {profissional}" if profissional else f"{rank}º Profissional"
                            if especialidade and especialidade != "Não informada":
                                titulo_local = f"{titulo_local} - {especialidade}"

                            delta_parts = [f"Exames: {exames}", f"Cirurgias: {cirurgias}"]
                            if crm:
                                delta_parts.insert(0, f"CRM {crm}")

                            col.metric(
                                titulo_local,
                                f"{total} Solicitações",
                                " • ".join(delta_parts),
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

            with graf2:
                by_turno_ind = detalhe_base.groupby("Turno")["Ocupado"].mean().reset_index()
                by_turno_ind["Taxa de Ocupação (%)"] = (by_turno_ind["Ocupado"] * 100).round(1)
                fig_ind_turno = px.bar(by_turno_ind, x="Turno", y="Taxa de Ocupação (%)",
                                       title=f"Ocupação por turno - {sala_detalhe}", text="Taxa de Ocupação (%)")
                fig_ind_turno.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                fig_ind_turno.update_yaxes(range=[0, 100])
                st.plotly_chart(fig_ind_turno, use_container_width=True)

            top_med_ind = ranking_ind.head(10) if not ranking_ind.empty else pd.DataFrame(columns=["EtiquetaLocal", "Total Procedimentos"])
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
                st.plotly_chart(fig_top_ind, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ---------- Integração das abas MÉDICOS (1, 2, 3...) ----------
def load_medicos_from_excel(excel: pd.ExcelFile):
    frames = []
    for s in excel.sheet_names:
        sn = _normalize_col(s)
        if "medic" in sn:  # captura "médicos", "medicos"
            try:
                dfm = excel.parse(s, header=0)
            except Exception:
                continue
            if dfm is None or dfm.empty:
                continue
            # normaliza colunas
            norm = {c:_normalize_col(c) for c in dfm.columns}
            dfm.columns = [norm[c] for c in dfm.columns]
            rename = {}
            for c in dfm.columns:
                if "nome" in c or "medico" in c: rename[c]="Médico"
                if c=="crm" or "crm" in c: rename[c]="CRM"
                if "especial" in c: rename[c]="Especialidade"
                if "planos" in c or c=="plano": rename[c]="Planos"
                if "valor" in c or "aluguel" in c or "negoci" in c: rename[c]="Valor Aluguel"
                if "exclus" in c: rename[c]="Sala Exclusiva"
                if "divid" in c: rename[c]="Sala Dividida"
            dfm = dfm.rename(columns=rename)
            keep = [c for c in ["Médico","CRM","Especialidade","Planos","Sala Exclusiva","Sala Dividida","Valor Aluguel"] if c in dfm.columns]
            if not keep:
                continue
            dfm = dfm[keep].copy()
            frames.append(dfm)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)
    # normalizações finais
    if "Médico" in out.columns: out["Médico"] = out["Médico"].astype(str).str.strip()
    if "Planos" in out.columns: out["Planos"] = out["Planos"].astype(str).str.strip()
    if "Valor Aluguel" in out.columns: out["Valor Aluguel"] = out["Valor Aluguel"].apply(_to_number)
    for c in ["Sala Exclusiva","Sala Dividida"]:
        if c in out.columns:
            out[c] = out[c].astype(str).str.strip().str.upper().replace({"X":"Sim","":""})
    return out

med_df = load_medicos_from_excel(excel)

st.markdown('<div id="planos"></div>', unsafe_allow_html=True)

med_enriched = pd.DataFrame()

if med_df.empty:
    st.warning("Não foram encontradas abas de **MÉDICOS** no arquivo. Os indicadores de plano/aluguel ficarão ocultos.")
else:
    # Enriquecer com turnos utilizados
    usos = fdf_base.groupby("Médico").size().reset_index(name="Turnos Utilizados")
    med_enriched = med_df.merge(usos, on="Médico", how="left")

    with st.container():
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<h2 class="section-title">💼 Indicador: PLANOS × Aluguel × Profissionais</h2>', unsafe_allow_html=True)

        # KPIs deste bloco
        tot_prof = med_enriched["Médico"].nunique()
        categorias_planos = med_enriched["Planos"].nunique() if "Planos" in med_enriched.columns else 0
        cpa, cpb, cpc = st.columns(3)
        cpa.metric("Profissionais (total)", tot_prof)
        cpb.metric("Categorias em PLANOS", categorias_planos)
        if "Valor Aluguel" in med_enriched.columns:
            media_valor = med_enriched["Valor Aluguel"].dropna().mean()
            cpc.metric("Valor médio de aluguel (R$)", f"{media_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X","."))
        else:
            cpc.metric("Valor médio de aluguel (R$)", "—")

        g1, g2 = st.columns(2)
        with g1:
            if "Planos" in med_enriched.columns:
                cont = med_enriched.groupby("Planos")["Médico"].nunique().reset_index(name="Profissionais")
                fig7 = px.bar(cont, x="Planos", y="Profissionais", title="Profissionais por PLANOS", text="Profissionais")
                fig7.update_traces(textposition="outside")
                st.plotly_chart(fig7, use_container_width=True)
            else:
                st.info("Coluna PLANOS não encontrada.")

        with g2:
            if "Valor Aluguel" in med_enriched.columns and "Planos" in med_enriched.columns:
                avgv = med_enriched.groupby("Planos")["Valor Aluguel"].mean().reset_index(name="Valor médio (R$)")
                avgv["Valor médio (R$)"] = avgv["Valor médio (R$)"].round(2)
                fig8 = px.bar(avgv, x="Planos", y="Valor médio (R$)", title="Valor médio de aluguel por PLANOS", text="Valor médio (R$)")
                fig8.update_traces(texttemplate="R$ %{y:.2f}", textposition="outside")
                st.plotly_chart(fig8, use_container_width=True)
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

        g3, g4 = st.columns(2)
        with g3:
            if "Especialidade" in med_enriched.columns and "Valor Aluguel" in med_enriched.columns:
                esp_avg = med_enriched.groupby("Especialidade")["Valor Aluguel"].mean().reset_index(name="Valor médio (R$)").sort_values("Valor médio (R$)", ascending=False)
                fig10 = px.bar(esp_avg, x="Valor médio (R$)", y="Especialidade", orientation="h", title="Valor médio de aluguel por especialidade", text="Valor médio (R$)")
                fig10.update_traces(texttemplate="R$ %{x:.2f}", textposition="outside")
                st.plotly_chart(fig10, use_container_width=True)
            else:
                st.info("Inclua 'Especialidade' e 'Valor Aluguel'.")
        with g4:
            if "Planos" in med_enriched.columns and "Especialidade" in med_enriched.columns:
                plano_esp = med_enriched.groupby(["Especialidade","Planos"])["Médico"].nunique().reset_index(name="Profissionais")
                fig11 = px.bar(plano_esp, x="Especialidade", y="Profissionais", color="Planos", barmode="group",
                               title="Profissionais por especialidade × PLANOS", text="Profissionais")
                fig11.update_traces(textposition="outside")
                st.plotly_chart(fig11, use_container_width=True)
            else:
                st.info("Inclua 'Especialidade' e 'PLANOS'.")

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
                else:
                    st.info("Sem marcações de sala exclusiva/dividida para analisar.")
            else:
                st.info("Inclua colunas 'Sala Exclusiva' e/ou 'Sala Dividida'.")

        st.markdown("##### Tabela (Médico × CRM × Especialidade × PLANOS × Valor × Tipo de Sala × Turnos)")
        cols_show = [c for c in ["Médico","CRM","Especialidade","Planos","Valor Aluguel","Sala Exclusiva","Sala Dividida","Turnos Utilizados"] if c in med_enriched.columns]
        st.dataframe(med_enriched[cols_show].sort_values(["Planos","Especialidade","Valor Aluguel","Médico"], na_position="last"), use_container_width=True)

        st.markdown('</div>', unsafe_allow_html=True)

# ---------- Detalhamento ----------
st.markdown('<div id="agenda"></div>', unsafe_allow_html=True)
st.markdown('<h2 class="section-title">📋 Agenda Detalhada (Tabela)</h2>', unsafe_allow_html=True)
st.dataframe(
    fdf.sort_values(["Sala","Dia","Turno"]).reset_index(drop=True)[["Sala","Dia","Turno","Médico"]],
    use_container_width=True
)

ranking_para_pdf = ranking_prod_total.copy()
if not ranking_para_pdf.empty:
    if sel_salas:
        ranking_para_pdf = ranking_para_pdf[ranking_para_pdf["Consultório"].isin(sel_salas)]
    if sel_medicos:
        ranking_para_pdf = ranking_para_pdf[ranking_para_pdf["Profissional"].isin(sel_medicos)]

pdf_bytes = build_pdf_report(
    summary_metrics,
    ranking_para_pdf,
    med_enriched if not med_df.empty else pd.DataFrame(),
    fdf,
)

csv = fdf.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "📄 Baixar relatório completo (PDF)",
    data=pdf_bytes,
    file_name="dashboard_consultorios.pdf",
    mime="application/pdf",
)
st.download_button("⬇️ Baixar dados filtrados (CSV)", data=csv, file_name="agenda_filtrada.csv", mime="text/csv")

st.markdown(
    '<div style="text-align: right; margin-top: 2rem;"><a href="#topo">⬆️ Voltar ao topo</a></div>',
    unsafe_allow_html=True,
)
