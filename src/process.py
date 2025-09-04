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
