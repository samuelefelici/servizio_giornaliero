import io, csv
from pathlib import Path
import pandas as pd
from dateutil import parser
from .constants import (
    ABSENCE_CODES, RESIDENZA_RENAME, DEFAULT_SORT,
    HEADER_PROBE, EXPECTED_COLUMNS
)
from .utils import find_header_row, coerce_time, clean_spaces, clean_column_names
# ...............................

def read_input_excel(file) -> tuple[pd.DataFrame, dict]:
    df_raw, origine = _read_excel_robusto(file)

    # 🔹 PULIZIA PRIMA
    df_raw = clean_spaces(df_raw)

    date_str, day_str = parse_date_and_day(df_raw)
    hdr_row = find_header_row(df_raw, HEADER_PROBE)
    if hdr_row is None:
        raise ValueError(f"Intestazione '{HEADER_PROBE}' non trovata.")

    # Prendiamo tutte le colonne dopo l'header, poi selezioniamo per nome
    df = df_raw.iloc[hdr_row:, :].copy()
    df.columns = clean_column_names(df.iloc[0].tolist())
    df = df.iloc[1:].reset_index(drop=True)

    # 🔹 PULIZIA DOPO: trim su tutte le celle stringa delle colonne utili
    keep_cols = [c for c in EXPECTED_COLUMNS if c in df.columns]
    if not keep_cols:
        raise ValueError(
            "Nessuna delle colonne attese è presente. "
            "Controlla che l'intestazione contenga almeno una tra: "
            + ", ".join(EXPECTED_COLUMNS)
        )
    df = df[keep_cols].copy()
    df = clean_spaces(df)

    # tipizzazioni / orari
    if "Matricola" in df.columns:
        df["Matricola"] = df["Matricola"].astype(str).str.strip()
    if "Inizio" in df.columns:
        df["Inizio"] = coerce_time(df["Inizio"])
    if "Fine" in df.columns:
        df["Fine"] = coerce_time(df["Fine"])

    meta = {"data": date_str, "giorno": day_str, "origine": origine}
    return df, meta

def transform_dataframe(df: pd.DataFrame, config_dir: Path) -> pd.DataFrame:
    """
    Applica:
      1) filtro matricole/turni da omettere (da CSV in config_dir)
      2) rinomina Residenza
      3) sostituzione sigle di assenza -> 'Assente' direttamente in Turno
      4) ordinamento per DEFAULT_SORT (+ 'Cognome e Nome' come tie-breaker)
    """
    matricole_omit, turni_omit = load_config_tables(config_dir)

    # 1) Filtri
    if "Matricola" in df.columns and len(matricole_omit):
        df = df[~df["Matricola"].isin(matricole_omit)].copy()
    if "Turno" in df.columns and len(turni_omit):
        df = df[~df["Turno"].isin(turni_omit)].copy()

    # 2) Rinomina residenze
    if "Residenza" in df.columns:
        df["Residenza"] = df["Residenza"].replace(RESIDENZA_RENAME)

    # 3) Turno -> 'Assente' se sigla di assenza
    if "Turno" in df.columns:
        df["Turno"] = (
            df["Turno"].astype(str).str.strip()
            .apply(lambda x: "Assente" if x in ABSENCE_CODES else x)
        )

    # 4) Ordinamento
    sort_cols = [c for c in DEFAULT_SORT if c in df.columns]
    by = sort_cols.copy()
    if "Cognome e Nome" in df.columns:
        by.append("Cognome e Nome")

    df_sorted = df.sort_values(by=by, kind="mergesort").reset_index(drop=True) if by else df.reset_index(drop=True)
    return df_sorted

