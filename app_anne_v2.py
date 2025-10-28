
import datetime as dt
import re
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Indicadores Clínica - Anne (v2)", page_icon="🩺", layout="wide")
st.title("🩺 Dashboard – Clínica (v2)")
st.caption("Adaptado para planilhas onde **cada dia possui subcolunas _MANHÃ_ e _TARDE_** nas abas de ocupação.")

uploaded = st.sidebar.file_uploader("Envie o arquivo Excel da Anne (.xlsx)", type=["xlsx"])

# ---------------- Utils ----------------
def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols = (pd.Index(df.columns.astype(str))
            .str.strip()
            .str.lower()
            .str.normalize('NFKD')
            .str.encode('ascii', errors='ignore')
            .str.decode('utf-8')
            .str.replace(' ', '_', regex=False))
    df.columns = cols
    return df

def read_multiheader(xlpath: str, sheet: str):
    try:
        df = pd.read_excel(xlpath, sheet_name=sheet, header=[0,1])
        # remove unnamed dup cols
        mask = ~df.columns.to_frame().apply(lambda col: col.isna().all()).values
        df = df.loc[:, mask]
        return df
    except Exception:
        # fallback single header
        try:
            df = pd.read_excel(xlpath, sheet_name=sheet)
            return df
        except Exception:
            return pd.DataFrame()

def read_single(xl, sheet):
    try:
        df = pd.read_excel(xl, sheet_name=sheet)
        # drop unnamed trailing
        df = df.loc[:, ~pd.Index(df.columns.astype(str)).str.contains("^Unnamed")]
        return df
    except Exception:
        return pd.DataFrame()

def long_from_occupancy(df_mh: pd.DataFrame, consultorio_label: str) -> pd.DataFrame:
    """
    df_mh: DataFrame with MultiIndex columns where level0 are days and level1 are MANHÃ/TARDE.
    First column usually holds 'SALA' identifiers.
    """
    if df_mh is None or df_mh.empty:
        return pd.DataFrame(columns=['data','consultorio','sala','dia_semana','turno','status','responsavel'])
    # Identify sala column (first col level0 is a digit like '1','2','3' or any non-day)
    # Build a tidy frame
    # Normalize column tuples to strings
    cols = df_mh.columns
    # room col assumed first
    sala_col = cols[0]
    sala_series = df_mh[sala_col].astype(str).str.strip()
    # Days expected: SEGUNDA..SEXTA (allow SABADO/DOMINGO if present)
    dias_validos = ["SEGUNDA","TERÇA","TERCA","QUARTA","QUINTA","SEXTA","SABADO","SÁBADO","DOMINGO"]
    stacks = []
    for (dia, turno) in cols[1:]:
        dia_up = str(dia).upper()
        turno_up = str(turno).upper()
        if dia_up not in dias_validos:
            continue
        # Extract column
        col_series = df_mh[(dia, turno)]
        # Coerce text
        txt = col_series.astype(str).str.strip()
        # Define status rules
        # ocupado = non-empty and not placeholder keywords like DISPONIVEL, LIVRE, VAGO
        mask_na = col_series.isna() | txt.eq("") | txt.eq("nan")
        mask_disp = txt.str.contains(r"DISPONIVEL|DISPONÍVEL|LIVRE|VAGO", case=False, na=False)
        mask_alugada = txt.str.contains(r"ALUGADA|ALUGADO|ALUG.", case=False, na=False)
        status = np.where(mask_alugada, "alugada",
                 np.where(~mask_na & ~mask_disp, "ocupado",
                 "disponivel"))
        stacks.append(pd.DataFrame({
            "sala": sala_series,
            "dia_semana": dia_up.replace("TERÇA","TERCA").replace("SÁBADO","SABADO"),
            "turno": turno_up,
            "responsavel": np.where(mask_na, None, txt.replace({"nan": None})),
            "status": status,
            "consultorio": str(consultorio_label),
        }))
    if not stacks:
        return pd.DataFrame(columns=['data','consultorio','sala','dia_semana','turno','status','responsavel'])
    out = pd.concat(stacks, ignore_index=True)
    # Map dia_semana to weekday index (Mon=0...)
    map_idx = {"SEGUNDA":0,"TERCA":1,"QUARTA":2,"QUINTA":3,"SEXTA":4,"SABADO":5,"DOMINGO":6}
    out["weekday"] = out["dia_semana"].map(map_idx)
    # Build pseudo dates for current week so series plots work
    today = dt.date.today()
    start_mon = today - dt.timedelta(days=today.weekday())
    out["data"] = out["weekday"].apply(lambda w: start_mon + dt.timedelta(days=int(w)) if pd.notna(w) else today)
    return out

def parse_ocupacao_all(xlpath: str):
    xls = pd.ExcelFile(xlpath)
    occ_frames = []
    for sheet in xls.sheet_names:
        if "OCUPAÇÃO DAS SALAS" in sheet.upper():
            df_mh = read_multiheader(xlpath, sheet)
            # consultório label: take first token digit from sheet name
            cons = re.findall(r"(\d+)", sheet)
            consultorio_label = cons[0] if cons else sheet
            occ_frames.append(long_from_occupancy(df_mh, consultorio_label))
    if occ_frames:
        return pd.concat(occ_frames, ignore_index=True)
    return pd.DataFrame(columns=['data','consultorio','sala','dia_semana','turno','status','responsavel'])

def parse_medicos_all(xlpath: str):
    xls = pd.ExcelFile(xlpath)
    med_frames = []
    for sheet in xls.sheet_names:
        if sheet.upper().startswith("MÉDICOS") or sheet.upper().startswith("MEDICOS"):
            df = read_single(xlpath, sheet)
            df = normalize_cols(df)
            rename = {"nome":"medico","crm":"crm","especialidade":"especialidade",
                      "planos":"planos","quais_planos_atende":"quais_planos_atende",
                      "consultorio":"consultorio","consultório":"consultorio","cadastro":"cadastro"}
            df = df.rename(columns={k:v for k,v in rename.items() if k in df.columns})
            keep = [c for c in ["medico","crm","especialidade","planos","quais_planos_atende","consultorio","cadastro"] if c in df.columns]
            med_frames.append(df[keep])
    if med_frames:
        md = pd.concat(med_frames, ignore_index=True).dropna(how="all")
        md["convenio_exclusivo"] = False
        md["negociacao"] = np.nan
        return md
    return pd.DataFrame(columns=["medico","crm","especialidade","planos","quais_planos_atende","consultorio","cadastro","convenio_exclusivo","negociacao"])

def parse_produtividade(xlpath: str):
    xls = pd.ExcelFile(xlpath)
    sheet = next((s for s in xls.sheet_names if "PRODUTIVIDADE" in s.upper()), None)
    if sheet is None:
        return pd.DataFrame(columns=['data','consultorio','medico','tipo','quantidade'])
    df = read_single(xlpath, sheet)
    df = normalize_cols(df)
    df = df.rename(columns={
        "medicos":"medico","médicos":"medico",
        "consultas_marcadas":"consultas_marcadas",
        "exames_solicitados":"exames_solicitados","exames_solicitados_":"exames_solicitados",
        "cirurgias_solicitadas":"cirurgias_solicitadas","cirurgias_solicitadas_":"cirurgias_solicitadas",
    })
    for c in ["consultas_marcadas","exames_solicitados","cirurgias_solicitadas"]:
        if c not in df.columns: df[c]=0
    m = df.melt(id_vars=[c for c in ["medico"] if c in df.columns],
                value_vars=["consultas_marcadas","exames_solicitados","cirurgias_solicitadas"],
                var_name="metric", value_name="quantidade")
    m["tipo"] = m["metric"].map({"consultas_marcadas":"marcacao","exames_solicitados":"exame","cirurgias_solicitadas":"cirurgia"})
    today = dt.date.today()
    m["data"] = dt.date(today.year, today.month, 1)
    m["consultorio"] = np.nan
    return m[["data","consultorio","medico","tipo","quantidade"]]

def classify_parceria(vol):
    if vol >= 50: return "Parceiro (Alto Volume)"
    if vol >= 20: return "Parceiro Potencial"
    return "Parceiro"

# ---------------- Main ----------------
if uploaded is None:
    st.info("Envie o arquivo .xlsx da Anne para ver os indicadores.")
    st.stop()

xlpath = uploaded
df_occ = parse_ocupacao_all(xlpath)
df_med = parse_medicos_all(xlpath)
df_prod = parse_produtividade(xlpath)

# KPIs de ocupação (por consultório e turno)
st.subheader("🏥 Ocupação de salas (sem contar **alugada**)")
if df_occ.empty:
    st.warning("Não foi possível ler as abas de ocupação com cabeçalho duplo (dias e MANHÃ/TARDE).")
else:
    df_occ_use = df_occ[~df_occ["status"].eq("alugada")].copy()
    # métricas
    occ_summary = (df_occ_use
                   .assign(ocupado=lambda d: np.where(d["status"].eq("ocupado"),1,0),
                           possiveis=1)  # disponível + ocupado contam como possíveis
                   .groupby(["consultorio","dia_semana","turno"], as_index=False)[["ocupado","possiveis"]].sum())
    occ_summary["taxa_ocupacao_%"] = np.where(occ_summary["possiveis"]>0, 100*occ_summary["ocupado"]/occ_summary["possiveis"], np.nan)
    c1,c2 = st.columns([2,1])
    fig = px.bar(occ_summary, x="dia_semana", y="taxa_ocupacao_%", color="turno", barmode="group",
                 facet_col="consultorio", facet_col_wrap=3, title="Taxa de ocupação por dia e turno (%)")
    c1.plotly_chart(fig, use_container_width=True)
    st.dataframe(occ_summary.sort_values(["consultorio","dia_semana","turno"]), use_container_width=True, height=420)

st.divider()

# Produtividade / marcações / exames / cirurgias (a partir da aba agregada)
st.subheader("📊 Produtividade (marcações, exames, cirurgias)")
if df_prod.empty:
    st.info("Preencha a aba de produtividade (ex.: 'PRODUTIVIDADE CONSULTÓRIOS 1, 2').")
else:
    by_tipo = df_prod.groupby("tipo", as_index=False)["quantidade"].sum()
    fig2 = px.bar(by_tipo, x="tipo", y="quantidade", title="Totais por tipo")
    st.plotly_chart(fig2, use_container_width=True)
    if "medico" in df_prod.columns:
        top_med = (df_prod.groupby("medico", as_index=False)["quantidade"].sum()
                   .sort_values("quantidade", ascending=False).head(10))
        fig3 = px.bar(top_med, x="medico", y="quantidade", title="Top 10 médicos (volume total)")
        st.plotly_chart(fig3, use_container_width=True)
    st.dataframe(df_prod, use_container_width=True, height=360)

st.divider()

# Médicos e parceria (placeholder)
st.subheader("🏷️ Médicos & Parceria (regra inicial)")
if df_med.empty:
    st.info("Preencha as abas 'MÉDICOS' para habilitar esta seção.")
else:
    vols = (df_prod.groupby("medico", as_index=False)["quantidade"].sum()
            if not df_prod.empty and "medico" in df_prod.columns else pd.DataFrame(columns=["medico","quantidade"]))
    md = df_med.merge(vols.rename(columns={"quantidade":"volume_total"}), on="medico", how="left")
    md["volume_total"] = md["volume_total"].fillna(0).astype(int)
    if "convenio_exclusivo" not in md.columns: md["convenio_exclusivo"] = False
    md["classificacao"] = np.where(md["convenio_exclusivo"], "Não interessante",
                                   md["volume_total"].apply(classify_parceria))
    st.dataframe(md.sort_values(["classificacao","volume_total"], ascending=[True, False]), use_container_width=True, height=420)
    pie = md.groupby("classificacao", as_index=False)["medico"].count().rename(columns={"medico":"qtd"})
    if not pie.empty:
        st.plotly_chart(px.pie(pie, names="classificacao", values="qtd", title="Distribuição de parceria"), use_container_width=True)

st.caption("Regras de classificação e suposições de leitura podem ser ajustadas conforme Anne validar os dados reais.")
