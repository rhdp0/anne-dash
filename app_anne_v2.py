
# streamlit_app.py
import io
import re
import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
from streamlit import config as st_config

st.set_page_config(page_title="Dashboard Consultórios", layout="wide")

THEME_OPTIONS = {"Escuro": "dark", "Claro": "light"}


def apply_theme(base_theme: str):
    """Adjust Streamlit and Plotly appearance according to the selected theme."""
    base = base_theme if base_theme in ("dark", "light") else "dark"
    # Update Streamlit base theme and align Plotly templates for better contrast
    st_config.set_option("theme.base", base)
    px.defaults.template = "plotly_dark" if base == "dark" else "plotly_white"


if "theme_base" not in st.session_state:
    st.session_state["theme_base"] = "dark"

apply_theme(st.session_state["theme_base"])

st.title("🏥 Dashboard de Consultórios — Ocupação, Médicos e Produtividade")

st.markdown(
    """
    Faça upload da planilha Excel (com abas como **OCUPAÇÃO DAS SALAS 1/2/3**, **MÉDICOS 1/2/3/4**, **PRODUTIVIDADE...**).
    O app vai consolidar os dados, criar indicadores e gráficos com filtros.
    """
)

uploaded = st.file_uploader("📤 Envie o arquivo .xlsx", type=["xlsx"])

# ---------------------------
# Helpers
# ---------------------------
DAY_ORDER = [
    "Segunda",
    "Terça",
    "Quarta",
    "Quinta",
    "Sexta",
    "Sábado",
    "Domingo",
]
DAY_INDEX = {day: idx for idx, day in enumerate(DAY_ORDER)}
DAY_ALIASES = {
    "SEGUNDA": "Segunda",
    "SEGUNDA-FEIRA": "Segunda",
    "SEGUNDA FEIRA": "Segunda",
    "TERCA": "Terça",
    "TERCA-FEIRA": "Terça",
    "TERCA FEIRA": "Terça",
    "QUARTA": "Quarta",
    "QUARTA-FEIRA": "Quarta",
    "QUARTA FEIRA": "Quarta",
    "QUINTA": "Quinta",
    "QUINTA-FEIRA": "Quinta",
    "QUINTA FEIRA": "Quinta",
    "SEXTA": "Sexta",
    "SEXTA-FEIRA": "Sexta",
    "SEXTA FEIRA": "Sexta",
    "SABADO": "Sábado",
    "SABADO-FEIRA": "Sábado",
    "SABADO FEIRA": "Sábado",
    "DOMINGO": "Domingo",
}
TURNS = ["MANHÃ", "TARDE"]
OCC_PREFIX = "OCUPAÇÃO DAS SALAS"
MED_PREFIX = "MÉDICOS"
PROD_PREFIX = "PRODUTIVIDADE"
IGNORAR_PALAVRAS_DEFAULT = ["alugada", "soube"]

def normalize_cols(df: pd.DataFrame):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def dedupe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure column labels are unique by keeping the first occurrence.
    This avoids pandas concat/groupby errors with duplicate column names.
    """
    if df is None or df.empty:
        return df
    # Keep first occurrence of duplicate columns
    return df.loc[:, ~pd.Index(df.columns).duplicated()]

def strip_accents(s: str) -> str:
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def canonical_day(label):
    if label is None or (isinstance(label, float) and pd.isna(label)):
        return None
    label_norm = strip_accents(str(label)).upper()
    for alias, canonical in DAY_ALIASES.items():
        if alias in label_norm:
            return canonical
    return None

def clean_text(s):
    if pd.isna(s):
        return None
    s2 = str(s).strip()
    return s2 if s2 else None

def first_non_null(seq):
    for x in seq:
        if pd.notna(x):
            return x
    return None

def try_number(x):
    try:
        return float(str(x).replace(",", ".").strip())
    except:
        return None

# ---------------------------
# Parsing
# ---------------------------
def parse_doctors(xls: pd.ExcelFile):
    doctors_frames = []
    for sheet in xls.sheet_names:
        if sheet.upper().startswith(MED_PREFIX):
            df = pd.read_excel(xls, sheet_name=sheet)
            if df is None or len(df) == 0:
                continue
            df = normalize_cols(df)
            # Remove duplicate columns before any renaming
            df = dedupe_columns(df)
            # Padroniza colunas principais se existirem
            colmap = {}
            for c in df.columns:
                cl = strip_accents(c).upper()
                if "NOME" == cl:
                    colmap[c] = "NOME"
                elif cl in ("CRM", "CRM "):
                    colmap[c] = "CRM"
                elif "ESPECIALIDADE" in cl:
                    colmap[c] = "ESPECIALIDADE"
                elif "CONSULTORIO" in cl:
                    colmap[c] = "CONSULTORIO"
                elif "CADASTRO" in cl:
                    colmap[c] = "CADASTRO"
                elif "PLANO" in cl or "CONVENIO" in cl:
                    colmap[c] = "TIPO_PLANO"
                elif "NEGOCIACAO" in cl or "NEGOCIA" in cl:
                    colmap[c] = "NEGOCIACAO"
            df = df.rename(columns=colmap)
            # If renaming created duplicate standardized names, keep the first
            df = dedupe_columns(df)
            df["SHEET_ORIGEM"] = sheet
            doctors_frames.append(df)
    if doctors_frames:
        doctors = pd.concat(doctors_frames, ignore_index=True)
        # Tipos
        if "TIPO_PLANO" not in doctors.columns:
            doctors["TIPO_PLANO"] = np.nan  # se não existir na planilha
        return doctors
    return pd.DataFrame()

def parse_occupancy(xls: pd.ExcelFile, ignorar_keywords=None):
    """Transforma as abas de ocupação em formato longo: uma linha por sala/dia/turno."""
    if ignorar_keywords is None:
        ignorar_keywords = IGNORAR_PALAVRAS_DEFAULT
    
    occ_rows = []
    for sheet in xls.sheet_names:
        if sheet.upper().startswith(OCC_PREFIX):
            df_raw = pd.read_excel(xls, sheet_name=sheet, header=0)
            if df_raw is None or len(df_raw) == 0:
                continue
            df = df_raw.copy()
            df = normalize_cols(df)

            # Primeira linha costuma conter rótulos de turno (MANHÃ/TARDE) por coluna
            # A primeira coluna geralmente é o identificador (SALA)
            if df.shape[0] < 2 or df.shape[1] < 3:
                continue

            # Turnos por coluna (a linha 0 costuma ter os rótulos de turnos)
            turnos_por_col = {}
            for c in df.columns:
                val = clean_text(df.iloc[0][c]) if 0 in df.index else None
                if val:
                    val_up = strip_accents(val).upper()
                    if any(t in val_up for t in ["MANH", "TARD"]):
                        if "MAN" in val_up:
                            turnos_por_col[c] = "MANHÃ"
                        elif "TARD" in val_up:
                            turnos_por_col[c] = "TARDE"
                        else:
                            turnos_por_col[c] = None
                    else:
                        turnos_por_col[c] = None
                else:
                    turnos_por_col[c] = None

            # Dias por coluna: se um cabeçalho é "Unnamed", herda o último dia nomeado
            dias_por_col = {}
            last_day = None
            for c in df.columns:
                header = strip_accents(str(c)).upper()
                # Algumas versões trazem "Unnamed: X" para TARDE; usamos o último dia nomeado
                is_unnamed = header.startswith("UNNAMED")
                if not is_unnamed:
                    matched_day = canonical_day(header)
                    if matched_day:
                        last_day = matched_day
                dias_por_col[c] = last_day

            # Tenta descobrir a coluna da SALA (onde há "SALA x")
            # Normalmente é a primeira coluna
            col_sala = df.columns[0]
            # Para cada linha (a partir da linha 1), processa
            for idx in range(1, len(df)):
                row = df.iloc[idx]
                sala_raw = clean_text(row[col_sala])
                if not sala_raw or "SALA" not in strip_accents(sala_raw).upper():
                    # ignora linhas sem SALA
                    continue
                sala = sala_raw

                for c in df.columns[1:]:
                    dia = dias_por_col.get(c)
                    turno = turnos_por_col.get(c)
                    if not dia or not turno:
                        continue
                    val = clean_text(row[c])

                    # Classificar status
                    status = "disponível"
                    medico_texto = None
                    if val:
                        vlow = strip_accents(val).lower()
                        if any(kw in vlow for kw in ignorar_keywords):
                            status = "ignorar"
                        else:
                            status = "ocupado"
                            medico_texto = val

                    occ_rows.append({
                        "SHEET_ORIGEM": sheet,
                        "CONSULTORIO": re.sub(r"[^0-9,]+", "", sheet).strip() or None,
                        "SALA": sala,
                        "DIA": dia,
                        "TURNO": turno,
                        "STATUS": status,
                        "MEDICO_RAW": medico_texto
                    })
    occ = pd.DataFrame(occ_rows)
    # Limpa consultório (ex.: "1, 2" -> múltiplos). Se vier vazio, define como "1"
    if not occ.empty:
        occ["CONSULTORIO"] = occ["CONSULTORIO"].replace("", np.nan)
        occ["CONSULTORIO"] = occ["CONSULTORIO"].fillna("1")
    return occ

def parse_productivity(xls: pd.ExcelFile):
    frames = []
    for sheet in xls.sheet_names:
        if sheet.upper().startswith(PROD_PREFIX):
            df = pd.read_excel(xls, sheet_name=sheet)
            if df is None or len(df) == 0:
                continue
            df = normalize_cols(df)
            df["SHEET_ORIGEM"] = sheet

            # Identifica colunas de interesse por palavras-chave
            col_consulta = [c for c in df.columns if "CONSULT" in strip_accents(c).upper()]
            col_exame = [c for c in df.columns if "EXAME" in strip_accents(c).upper()]
            col_cirur = [c for c in df.columns if "CIRUR" in strip_accents(c).upper() or "CIRURG" in strip_accents(c).upper()]

            # Tenta identificar consultórios citados na planilha
            df_long = df.copy()
            # Mantém apenas numéricos nas colunas alvo quando possível
            for cc in col_consulta + col_exame + col_cirur:
                df_long[cc] = pd.to_numeric(df_long[cc], errors="coerce")

            # Agrega por sheet (se não houver chaves explícitas)
            agg = {}
            if col_consulta: agg["CONSULTAS"] = df_long[col_consulta].sum(axis=1)
            else: agg["CONSULTAS"] = 0
            if col_exame: agg["EXAMES"] = df_long[col_exame].sum(axis=1)
            else: agg["EXAMES"] = 0
            if col_cirur: agg["CIRURGIAS"] = df_long[col_cirur].sum(axis=1)
            else: agg["CIRURGIAS"] = 0
            df_out = pd.DataFrame(agg)
            df_out["SHEET_ORIGEM"] = sheet
            frames.append(df_out)

    if frames:
        prod = pd.concat(frames, ignore_index=True)
        # Total geral por sheet (como fallback)
        prod = prod.groupby("SHEET_ORIGEM", as_index=False)[["CONSULTAS","EXAMES","CIRURGIAS"]].sum()
        return prod
    return pd.DataFrame(columns=["SHEET_ORIGEM","CONSULTAS","EXAMES","CIRURGIAS"])

# ---------------------------
# App main
# ---------------------------
if uploaded is None:
    st.info("Envie o arquivo Excel para começar.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded)
except Exception as e:
    st.error(f"Não foi possível abrir o arquivo: {e}")
    st.stop()

# Parâmetros avançados
with st.expander("⚙️ Opções avançadas de parsing"):
    ignorar_kw = st.text_input(
        "Palavras-chave para marcar horários/salas a **ignorar** na taxa de ocupação (separadas por vírgula).",
        value=", ".join(IGNORAR_PALAVRAS_DEFAULT)
    )
    ignorar_keywords = [w.strip().lower() for w in ignorar_kw.split(",") if w.strip()]

# Parse
doctors = parse_doctors(xls)
occ = parse_occupancy(xls, ignorar_keywords=ignorar_keywords)
prod = parse_productivity(xls)

# Mostra status de ingestão
c1, c2, c3 = st.columns(3)
with c1:
    st.metric("Abas de Médicos", f"{sum(1 for s in xls.sheet_names if s.upper().startswith(MED_PREFIX))}")
with c2:
    st.metric("Abas de Ocupação", f"{sum(1 for s in xls.sheet_names if s.upper().startswith(OCC_PREFIX))}")
with c3:
    st.metric("Abas de Produtividade", f"{sum(1 for s in xls.sheet_names if s.upper().startswith(PROD_PREFIX))}")

st.divider()

# ---------------------------
# Filtros globais
# ---------------------------
st.sidebar.header("Aparência")
theme_label = st.sidebar.radio(
    "Tema do dashboard",
    list(THEME_OPTIONS.keys()),
    index=list(THEME_OPTIONS.values()).index(st.session_state["theme_base"]),
)
selected_theme = THEME_OPTIONS[theme_label]
if selected_theme != st.session_state["theme_base"]:
    st.session_state["theme_base"] = selected_theme
    apply_theme(selected_theme)
    st.rerun()

st.sidebar.divider()
st.sidebar.header("Filtros")
consultorios_disp = sorted(set((occ["CONSULTORIO"].dropna().unique().tolist() if not occ.empty else []) +
                               (doctors["CONSULTORIO"].dropna().unique().tolist() if "CONSULTORIO" in doctors.columns else [])))
consultorio_sel = st.sidebar.multiselect("Consultório", consultorios_disp, default=consultorios_disp)

especialidades_disp = sorted(doctors["ESPECIALIDADE"].dropna().unique().tolist()) if "ESPECIALIDADE" in doctors.columns else []
especialidade_sel = st.sidebar.multiselect("Especialidade", especialidades_disp, default=especialidades_disp)

tipos_plano_disp = sorted(doctors["TIPO_PLANO"].dropna().astype(str).unique().tolist()) if "TIPO_PLANO" in doctors.columns else []
tipo_plano_sel = st.sidebar.multiselect("Tipo de plano", tipos_plano_disp, default=tipos_plano_disp)

# Aplicação de filtros nos datasets
occ_f = occ.copy()
if consultorio_sel and not occ_f.empty:
    occ_f = occ_f[occ_f["CONSULTORIO"].isin(consultorio_sel)]

doctors_f = doctors.copy()
if "CONSULTORIO" in doctors_f.columns and consultorio_sel:
    doctors_f = doctors_f[doctors_f["CONSULTORIO"].astype(str).isin(consultorio_sel)]
if "ESPECIALIDADE" in doctors_f.columns and especialidade_sel:
    doctors_f = doctors_f[doctors_f["ESPECIALIDADE"].isin(especialidade_sel)]
if "TIPO_PLANO" in doctors_f.columns and tipo_plano_sel:
    doctors_f = doctors_f[doctors_f["TIPO_PLANO"].astype(str).isin(tipo_plano_sel)]

# ---------------------------
# KPIs topo (Visão Geral)
# ---------------------------
st.subheader("📊 Visão Geral")

# Taxa de ocupação
if not occ_f.empty:
    total_slots = len(occ_f[occ_f["STATUS"] != "ignorar"])
    ocupados = len(occ_f[occ_f["STATUS"] == "ocupado"])
    disponiveis = len(occ_f[occ_f["STATUS"] == "disponível"])
    taxa = (ocupados / total_slots * 100) if total_slots else 0.0
else:
    taxa, total_slots, ocupados, disponiveis = 0.0, 0, 0, 0

# Médicos
total_medicos = doctors_f["CRM"].nunique() if "CRM" in doctors_f.columns else doctors_f.shape[0]
pct_parceiros = 0.0
pct_nao_estrateg = 0.0
if "TIPO_PLANO" in doctors_f.columns and doctors_f.shape[0] > 0:
    parceiros = doctors_f["TIPO_PLANO"].astype(str).str.upper().str.contains("JAYME|MISTO", regex=True, na=False).sum()
    nao_estr = doctors_f["TIPO_PLANO"].astype(str).str.upper().str.contains("PROPRIO|PRÓPRIO", regex=True, na=False).sum()
    pct_parceiros = parceiros / max(1, doctors_f.shape[0]) * 100
    pct_nao_estrateg = nao_estr / max(1, doctors_f.shape[0]) * 100

# Produtividade
if not prod.empty:
    prod_tot = prod[["CONSULTAS","EXAMES","CIRURGIAS"]].sum()
    total_consultas, total_exames, total_cirurgias = int(prod_tot.get("CONSULTAS",0)), int(prod_tot.get("EXAMES",0)), int(prod_tot.get("CIRURGIAS",0))
else:
    total_consultas = total_exames = total_cirurgias = 0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Taxa de ocupação", f"{taxa:.1f}%")
k2.metric("Slots ocupados", f"{ocupados}/{total_slots}")
k3.metric("Médicos (únicos)", f"{int(total_medicos)}")
k4.metric("% Médicos parceiros (Jayme/Misto)", f"{pct_parceiros:.0f}%")
k5.metric("Consultas / Exames / Cirurgias", f"{total_consultas} / {total_exames} / {total_cirurgias}")

st.divider()

# ---------------------------
# Seção 1: Ocupação das Salas
# ---------------------------
st.header("📅 Ocupação das Salas")

if occ_f.empty:
    st.warning("Não foi possível identificar dados de **Ocupação** nas abas enviadas.")
else:
    occ_valid = occ_f[occ_f["STATUS"] != "ignorar"].copy()
    if occ_valid.empty:
        st.info("Todos os registros de ocupação foram marcados como 'ignorar'.")
    else:
        occ_valid["OCUPADO"] = (
            occ_valid["STATUS"].astype(str).str.lower().eq("ocupado").astype(float)
        )

        def sort_dia(dia):
            if pd.isna(dia):
                return len(DAY_ORDER)
            canonical = canonical_day(dia)
            if canonical:
                return DAY_INDEX.get(canonical, len(DAY_ORDER))
            return len(DAY_ORDER)

        dias_presentes = sorted(occ_valid["DIA"].dropna().unique(), key=sort_dia)

        if dias_presentes:
            st.subheader("Visão diária (sala x turno)")
            tabs = st.tabs([str(dia).title() for dia in dias_presentes])
            for tab, dia in zip(tabs, dias_presentes):
                with tab:
                    df_dia = occ_valid[occ_valid["DIA"] == dia]
                    df_slots = (
                        df_dia.groupby(["SALA", "TURNO"], as_index=False)
                        .agg(
                            {
                                "STATUS": lambda vals: first_non_null(vals),
                                "MEDICO_RAW": lambda vals: first_non_null(vals),
                            }
                        )
                    )

                    if df_slots.empty:
                        st.info("Sem dados suficientes para este dia.")
                    else:
                        def build_display(row):
                            status = str(row["STATUS"]).lower()
                            if status == "ocupado":
                                medico = row.get("MEDICO_RAW")
                                if pd.isna(medico):
                                    return "Ocupado"
                                text = str(medico).strip()
                                return text if text else "Ocupado"
                            # Slots livres ficam sem texto, apenas com destaque visual
                            return ""

                        df_slots["DISPLAY"] = df_slots.apply(build_display, axis=1)

                        # Mantém a ordem original das salas e garante colunas por turno
                        salas_order = (
                            df_slots["SALA"]
                            .drop_duplicates()
                            .tolist()
                        )
                        display_table = (
                            df_slots.pivot(index="SALA", columns="TURNO", values="DISPLAY")
                            .reindex(index=salas_order)
                        )
                        display_table = display_table.reindex(columns=TURNS, fill_value="")

                        status_table = (
                            df_slots.pivot(index="SALA", columns="TURNO", values="STATUS")
                            .reindex(index=salas_order)
                        )
                        status_table = status_table.reindex(columns=TURNS)

                        def highlight_free(row):
                            statuses = status_table.loc[row.name]
                            styles = []
                            for col in row.index:
                                if isinstance(statuses.get(col), str) and statuses.get(col).lower() == "disponível":
                                    styles.append("background-color: #d4edda; color: #155724; font-weight: 600;")
                                else:
                                    styles.append("")
                            return styles

                        styled = display_table.fillna("").style.apply(highlight_free, axis=1)

                        st.dataframe(
                            styled,
                            use_container_width=True,
                        )
                        st.caption(
                            "Nome do médico responsável por sala e turno; células livres destacadas em verde."
                        )

        # Barras por sala (visão consolidada)
        df_bar = (
            occ_valid.groupby(["SALA"], as_index=False)["OCUPADO"].mean()
            .rename(columns={"OCUPADO": "OCUPACAO_%"})
        )
        df_bar["OCUPACAO_%"] = df_bar["OCUPACAO_%"] * 100
        df_bar = df_bar.sort_values("OCUPACAO_%", ascending=False)
        if not df_bar.empty:
            fig2 = px.bar(
                df_bar,
                x="OCUPACAO_%",
                y="SALA",
                orientation="h",
                title="Taxa média de ocupação por sala (%)",
            )
            fig2.update_layout(height=450, margin=dict(l=20, r=20, t=50, b=20))
            st.plotly_chart(fig2, use_container_width=True)

    # Stacked status por consultório
    df_stack = (
        occ_f.assign(
            STATUS2=occ_f["STATUS"].replace(
                {"disponível": "Disponível", "ocupado": "Ocupado", "ignorar": "Ignorar"}
            )
        )
        .groupby(["CONSULTORIO", "STATUS2"], as_index=False)
        .size()
    )
    if not df_stack.empty:
        fig3 = px.bar(
            df_stack,
            x="CONSULTORIO",
            y="size",
            color="STATUS2",
            title="Distribuição de status por consultório",
            barmode="stack",
        )
        st.plotly_chart(fig3, use_container_width=True)

    with st.expander("🔎 Tabela detalhada (Ocupação)"):
        st.dataframe(occ_f, use_container_width=True, height=350)

st.divider()

# ---------------------------
# Seção 2: Médicos e Especialidades
# ---------------------------
st.header("👩‍⚕️ Médicos e Especialidades")

if doctors_f.empty:
    st.warning("Não foi possível identificar dados de **Médicos** nas abas enviadas.")
else:
    col_a, col_b = st.columns(2)

    # Distribuição de médicos por especialidade
    if "ESPECIALIDADE" in doctors_f.columns:
        dist_esp = doctors_f.groupby("ESPECIALIDADE", as_index=False).size().sort_values("size", ascending=False)
        with col_a:
            fig4 = px.treemap(dist_esp, path=["ESPECIALIDADE"], values="size", title="Distribuição por especialidade")
            st.plotly_chart(fig4, use_container_width=True)

    # Distribuição por tipo de plano
    if "TIPO_PLANO" in doctors_f.columns:
        dist_plano = doctors_f["TIPO_PLANO"].fillna("Não informado").value_counts().reset_index()
        dist_plano.columns = ["TIPO_PLANO", "QTD"]
        with col_b:
            fig5 = px.pie(dist_plano, names="TIPO_PLANO", values="QTD", title="Tipos de plano (convênio)")
            st.plotly_chart(fig5, use_container_width=True)

    # Ranking médicos por produtividade (se existir prod detalhado no futuro)
    # Como fallback, mostramos contagem por consultório/origem
    if "CONSULTORIO" in doctors_f.columns:
        dist_cons = doctors_f["CONSULTORIO"].astype(str).value_counts().reset_index()
        dist_cons.columns = ["CONSULTORIO", "QTD_MEDICOS"]
        fig6 = px.bar(dist_cons, x="CONSULTORIO", y="QTD_MEDICOS", title="Médicos por consultório")
        st.plotly_chart(fig6, use_container_width=True)

    # Tabela com sinalizadores
    df_tab = doctors_f.copy()
    if "TIPO_PLANO" in df_tab.columns:
        tipo_up = df_tab["TIPO_PLANO"].astype(str).str.upper()
        df_tab["SINAL"] = np.where(tipo_up.str.contains("PROPRIO|PRÓPRIO", regex=True, na=False), "🔴 Não estratégico",
                            np.where(tipo_up.str.contains("JAYME|MISTO", regex=True, na=False), "🟢 Parceiro", "🟡 Neutro"))
    with st.expander("📋 Tabela de médicos (com sinalizadores)"):
        st.dataframe(df_tab, use_container_width=True, height=360)

st.divider()

# ---------------------------
# Seção 3: Produtividade e Planos
# ---------------------------
st.header("📈 Produtividade")

if prod.empty:
    st.info("Não identifiquei abas de PRODUTIVIDADE. Quando existirem, este painel mostrará comparativos.")
else:
    cpa, cpb = st.columns(2)
    with cpa:
        fig7 = px.bar(prod.melt(id_vars="SHEET_ORIGEM", value_vars=["CONSULTAS","EXAMES","CIRURGIAS"],
                                var_name="Tipo", value_name="Quantidade"),
                      x="SHEET_ORIGEM", y="Quantidade", color="Tipo", barmode="group",
                      title="Consultas, Exames e Cirurgias por aba de produtividade")
        st.plotly_chart(fig7, use_container_width=True)
    with cpb:
        tot = prod[["CONSULTAS","EXAMES","CIRURGIAS"]].sum().reset_index()
        tot.columns = ["Tipo","Total"]
        fig8 = px.pie(tot, names="Tipo", values="Total", title="Participação por tipo")
        st.plotly_chart(fig8, use_container_width=True)

    # Placeholder para scatter Negociação x Produtividade (se a planilha de médicos tiver NEGOCIAÇÃO + quando houver produtividade por médico)
    if "NEGOCIACAO" in doctors_f.columns:
        # Sem produtividade por médico, usamos proxy de 1 para desenhar e não quebrar — o usuário poderá evoluir este bloco quando houver dados.
        tmp = doctors_f.copy()
        tmp["NEGOCIACAO_NUM"] = pd.to_numeric(tmp["NEGOCIACAO"], errors="coerce")
        tmp = tmp.dropna(subset=["NEGOCIACAO_NUM"])
        if not tmp.empty:
            tmp["PRODUTIVIDADE_PROXY"] = 1  # placeholder
            fig_sc = px.scatter(tmp, x="NEGOCIACAO_NUM", y="PRODUTIVIDADE_PROXY",
                                hover_data=["NOME","ESPECIALIDADE","TIPO_PLANO"],
                                title="Negociação (R$) × Produtividade (proxy) — ajuste quando tiver dados por médico")
            st.plotly_chart(fig_sc, use_container_width=True)

st.caption("💡 Dica: ajuste as palavras-chave de *ignorar* na seção de Opções Avançadas para excluir salas/slots alugados da taxa de ocupação.")
