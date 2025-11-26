"""Excel reading helpers for consultório datasets."""
from __future__ import annotations

from typing import Any, Dict, List, Optional

import pandas as pd

from .processors import format_consultorio_label, normalize_column_name, to_number


__all__ = [
    "load_excel_workbook",
    "detect_header_and_parse",
    "tidy_agenda_from_workbook",
    "load_produtividade_from_workbook",
    "load_medicos_from_workbook",
]


def load_excel_workbook(source: Any) -> pd.ExcelFile:
    """Return an :class:`~pandas.ExcelFile` from a file-like or path."""
    try:
        return pd.ExcelFile(source)
    except Exception as exc:  # pragma: no cover - streamlit surfaces the error message
        raise ValueError("Não foi possível abrir o arquivo de Excel") from exc


def detect_header_and_parse(excel: pd.ExcelFile, sheet_name: str) -> Optional[pd.DataFrame]:
    """Heuristically detect headers for agenda sheets and parse them."""
    for header in range(0, 5):
        try:
            df = excel.parse(sheet_name, header=header)
        except Exception:
            continue
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if df.empty:
            continue

        cols_norm = [normalize_column_name(col) for col in df.columns]
        col_dia = None
        col_manha = None
        col_tarde = None

        for idx, col_norm in enumerate(cols_norm):
            if col_dia is None:
                if "dia" in col_norm or any(
                    day in col_norm
                    for day in [
                        "segunda",
                        "terca",
                        "terça",
                        "quarta",
                        "quinta",
                        "sexta",
                        "sabado",
                        "sábado",
                    ]
                ):
                    col_dia = df.columns[idx]
            if any(key in col_norm for key in ["manha", "manhã"]):
                col_manha = df.columns[idx]
            if "tarde" in col_norm:
                col_tarde = df.columns[idx]

        if col_dia is None and len(df.columns) >= 1:
            first_col = df.columns[0]
            sample = df[first_col].astype(str).str.lower()
            if sample.str.contains("segunda|terca|terça|quarta|quinta|sexta|sabado|sábado").any():
                col_dia = first_col

        if col_dia is not None and (col_manha is not None or col_tarde is not None):
            use_cols = [col for col in [col_dia, col_manha, col_tarde] if col is not None]
            df = df[use_cols].copy()
            rename = {col_dia: "Dia"}
            if col_manha is not None:
                rename[col_manha] = "Manhã"
            if col_tarde is not None:
                rename[col_tarde] = "Tarde"
            df = df.rename(columns=rename)
            df["Dia"] = df["Dia"].astype(str).str.strip()
            df = df[df["Dia"].str.len() > 0]
            return df
    return None


def tidy_agenda_from_workbook(excel: pd.ExcelFile) -> pd.DataFrame:
    """Return the agenda in a tidy format by stacking relevant sheets."""
    frames: List[pd.DataFrame] = []
    for sheet in excel.sheet_names:
        normalized = normalize_column_name(sheet)
        if "consult" in normalized and "ocupa" not in normalized:
            df = detect_header_and_parse(excel, sheet)
            if df is None or df.empty:
                continue
            df["Dia"] = (
                df["Dia"]
                .astype(str)
                .str.strip()
                .str.replace("terca", "terça", case=False)
                .str.replace("sabado", "sábado", case=False)
                .str.capitalize()
            )
            df.insert(0, "Sala", sheet.strip())
            long_df = df.melt(
                id_vars=["Sala", "Dia"],
                value_vars=[col for col in ["Manhã", "Tarde"] if col in df.columns],
                var_name="Turno",
                value_name="Médico",
            )
            long_df["Médico"] = (
                long_df["Médico"].astype(str).replace({"nan": "", "None": ""}).str.strip()
            )
            frames.append(long_df)
    if not frames:
        return pd.DataFrame(columns=["Sala", "Dia", "Turno", "Médico"])
    agenda = pd.concat(frames, ignore_index=True)
    agenda["Dia"] = pd.Categorical(
        agenda["Dia"],
        categories=["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"],
        ordered=True,
    )
    agenda["Ocupado"] = agenda["Médico"].str.len() > 0
    return agenda


def load_produtividade_from_workbook(excel: pd.ExcelFile) -> pd.DataFrame:
    """Load productivity data from the Excel workbook."""
    frames: List[pd.DataFrame] = []
    for sheet in excel.sheet_names:
        normalized_sheet = normalize_column_name(sheet)
        if "consultorios" in normalized_sheet:
            continue
        is_prod_sheet = "produtiv" in normalized_sheet and "consult" in normalized_sheet
        is_contas_sheet = "conta" in normalized_sheet and "medic" in normalized_sheet
        if not (is_prod_sheet or is_contas_sheet):
            continue
        for header in range(0, 6):
            try:
                dfp = excel.parse(sheet, header=header)
            except Exception:
                continue
            dfp = dfp.dropna(how="all").dropna(axis=1, how="all")
            if dfp.empty:
                continue

            rename: Dict[Any, str] = {}
            for col in dfp.columns:
                norm = normalize_column_name(col)
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
                elif (
                    "receita" in norm
                    or ("valor" in norm and "aluguel" not in norm and "total" in norm)
                    or "fatur" in norm
                ):
                    rename[col] = "Receita"
                elif any(key in norm for key in ["data", "period", "período", "mes", "mês", "compet"]):
                    rename[col] = "Data"
                elif "consult" in norm and "produtiv" not in norm:
                    rename[col] = "Consultório"

            dfp = dfp.rename(columns=rename)

            if "Profissional" not in dfp.columns:
                continue
            if (
                "Exames Solicitados" not in dfp.columns
                and "Cirurgias Solicitadas" not in dfp.columns
            ):
                continue

            keep = [
                col
                for col in [
                    "Profissional",
                    "CRM",
                    "Especialidade",
                    "Exames Solicitados",
                    "Cirurgias Solicitadas",
                    "Receita",
                    "Consultório",
                    "Data",
                ]
                if col in dfp.columns
            ]
            dfp = dfp[keep].copy()

            dfp["Profissional"] = dfp["Profissional"].astype(str).str.strip()
            dfp = dfp[dfp["Profissional"].str.len() > 0]
            dfp = dfp[
                dfp["Profissional"].apply(
                    lambda value: normalize_column_name(value)
                    not in {"total", "totais", "subtotal"}
                )
            ]

            if "Consultório" not in dfp.columns:
                dfp["Consultório"] = format_consultorio_label(sheet)
            else:
                dfp["Consultório"] = (
                    dfp["Consultório"].astype(str)
                    .str.strip()
                    .replace(
                        r"(?i)^(nan|none|null|na|n/a|sem\s*informac[aã]o|sem\s*dados?)$",
                        "",
                        regex=True,
                    )
                )
                dfp.loc[dfp["Consultório"].eq(""), "Consultório"] = format_consultorio_label(sheet)
                dfp["Consultório"] = dfp["Consultório"].fillna(
                    format_consultorio_label(sheet)
                )

            if "Especialidade" in dfp.columns:
                dfp["Especialidade"] = dfp["Especialidade"].astype(str).str.strip()
            else:
                dfp["Especialidade"] = ""

            if "CRM" in dfp.columns:
                dfp["CRM"] = dfp["CRM"].astype(str).str.strip()

            if "Data" in dfp.columns:
                dfp["Data"] = pd.to_datetime(dfp["Data"], errors="coerce")
                dfp["Período"] = dfp["Data"].dt.to_period("M").astype(str)
                dfp.loc[dfp["Data"].isna(), "Período"] = ""
            else:
                dfp["Período"] = ""

            for col in ["Exames Solicitados", "Cirurgias Solicitadas", "Receita"]:
                if col in dfp.columns:
                    dfp[col] = pd.to_numeric(dfp[col], errors="coerce").fillna(0)

            dfp["Consultório"] = dfp["Consultório"].apply(format_consultorio_label)
            dfp["_SalaNorm"] = dfp["Consultório"].apply(normalize_column_name)
            frames.append(dfp)
    if not frames:
        return pd.DataFrame(
            columns=[
                "Profissional",
                "CRM",
                "Especialidade",
                "Exames Solicitados",
                "Cirurgias Solicitadas",
                "Receita",
                "Consultório",
            ]
        )
    produtividade_df = pd.concat(frames, ignore_index=True)
    numeric_defaults = {
        "Receita": 0.0,
        "Exames Solicitados": 0,
        "Cirurgias Solicitadas": 0,
    }
    for column, default in numeric_defaults.items():
        if column in produtividade_df.columns:
            produtividade_df[column] = (
                pd.to_numeric(produtividade_df[column], errors="coerce").fillna(default)
            )
        else:
            produtividade_df[column] = default
    return produtividade_df


def load_medicos_from_workbook(excel: pd.ExcelFile) -> pd.DataFrame:
    """Load the MÉDICOS tabs and normalize their structure."""
    frames: List[pd.DataFrame] = []
    for sheet_name in excel.sheet_names:
        normalized_sheet = normalize_column_name(sheet_name)
        if "medic" not in normalized_sheet:
            continue
        try:
            dfm = excel.parse(sheet_name, header=0)
        except Exception:
            continue
        if dfm is None or dfm.empty:
            continue
        normalized_cols = {col: normalize_column_name(col) for col in dfm.columns}
        dfm.columns = [normalized_cols[col] for col in dfm.columns]
        rename: Dict[str, str] = {}
        consultorio_candidates: List[str] = []

        def is_consultorio_candidate(column: str) -> bool:
            col_norm = normalize_column_name(column)
            if not col_norm:
                return False
            if any(
                keyword in col_norm
                for keyword in ["exclus", "divid", "plan", "valor", "crm", "turno"]
            ):
                return False
            return any(keyword in col_norm for keyword in ["consult", "sala", "unid"])

        for column in dfm.columns:
            if "nome" in column or "medico" in column:
                rename[column] = "Médico"
            if column == "crm" or "crm" in column:
                rename[column] = "CRM"
            if "especial" in column:
                rename[column] = "Especialidade"
            if "planos" in column or column == "plano":
                rename[column] = "Planos"
            if "valor" in column or "aluguel" in column or "negoci" in column:
                rename[column] = "Valor Aluguel"
            if "exclus" in column:
                rename[column] = "Sala Exclusiva"
            if "divid" in column:
                rename[column] = "Sala Dividida"
            if is_consultorio_candidate(column):
                consultorio_candidates.append(column)
                if "consult" in column:
                    rename[column] = "Consultório"

        dfm = dfm.rename(columns=rename)
        if "Consultório" not in dfm.columns and consultorio_candidates:
            candidate = next((cand for cand in consultorio_candidates if cand in dfm.columns), None)
            if candidate is not None:
                dfm = dfm.rename(columns={candidate: "Consultório"})
        if "Consultório" not in dfm.columns:
            dfm["Consultório"] = format_consultorio_label(sheet_name)

        keep = [
            column
            for column in [
                "Médico",
                "CRM",
                "Especialidade",
                "Planos",
                "Sala Exclusiva",
                "Sala Dividida",
                "Consultório",
                "Valor Aluguel",
            ]
            if column in dfm.columns
        ]
        if not keep:
            continue
        frames.append(dfm[keep].copy())
    if not frames:
        return pd.DataFrame()
    result = pd.concat(frames, ignore_index=True)
    if "Médico" in result.columns:
        result["Médico"] = result["Médico"].astype(str).str.strip()
    if "CRM" in result.columns:
        result["CRM"] = result["CRM"].astype(str).str.strip()
    if "Especialidade" in result.columns:
        result["Especialidade"] = result["Especialidade"].astype(str).str.strip()
    if "Consultório" in result.columns:
        result["Consultório"] = result["Consultório"].apply(format_consultorio_label)
    if "Valor Aluguel" in result.columns:
        result["Valor Aluguel"] = result["Valor Aluguel"].apply(to_number)
    return result
