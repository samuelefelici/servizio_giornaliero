import pandas as pd
from dateutil import parser
from .constants import (
    ABSENCE_CODES, RESIDENZA_RENAME, DEFAULT_SORT,
    HEADER_PROBE, EXPECTED_COLUMNS
)
from .utils import find_header_row, coerce_time, clean_spaces
from pathlib import Path

def load_config_tables(config_dir: Path):
    m_omit = pd.read_csv(config_dir / "matricole_da_omettere.csv", dtype=str, keep_default_na=False)
    t_omit = pd.read_csv(config_dir / "turni_attivita_da_omettere.csv", dtype=str, keep_default_na=False)
    matricole = set(m_omit["Matricola"].astype(str).str.strip())
    turni = set(t_omit["Turno"].astype(str).str.strip())
    return matricole, turni

def parse_date_and_day(df_raw: pd.DataFrame) -> tuple[str,str]:
    # Prima riga: data (es. 05/09/2025); seconda riga: giorno (es. Venerdi)
    date_str = str(df_raw.iloc[0,0]).strip() if df_raw.shape[1]>0 else ""
    day_str  = str(df_raw.iloc[1,0]).strip() if df_raw.shape[1]>0 and len(df_raw)>1 else ""
    try:
        # normalizza formato data (se possibile)
        dt = parser.parse(date_str, dayfirst=True).date()
        date_str = dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return date_str, day_str

def read_input_excel(file) -> tuple[pd.DataFrame, dict]:
    # header=None, per trovare la riga giusta
    df_raw = pd.read_excel(file, header=None, engine=None)  # openpyxl o xlrd a seconda estensione
    date_str, day_str = parse_date_and_day(df_raw)
    hdr_row = find_header_row(df_raw, HEADER_PROBE)
    if hdr_row is None:
        raise ValueError(f"Intestazione '{HEADER_PROBE}' non trovata.")
    df = df_raw.iloc[hdr_row:, :len(EXPECTED_COLUMNS)].copy()
    df.columns = df.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)

    # normalizza nomi colonna
    df = df[[c for c in EXPECTED_COLUMNS if c in df.columns]].copy()
    df = clean_spaces(df)

    # tipologie base
    # Matricola in stringa per uniformare
    if "Matricola" in df.columns:
        df["Matricola"] = df["Matricola"].astype(str).str.strip()

    # orari
    if "Inizio" in df.columns:
        df["Inizio"] = coerce_time(df["Inizio"])
    if "Fine" in df.columns:
        df["Fine"]   = coerce_time(df["Fine"])

    meta = {"data": date_str, "giorno": day_str}
    return df, meta

def transform_dataframe(df: pd.DataFrame, config_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    matricole_omit, turni_omit = load_config_tables(config_dir)

    # 1) Filtri: matricole da omettere
    if "Matricola" in df.columns:
        df = df[~df["Matricola"].isin(matricole_omit)].copy()

    # 2) Filtri: turni/attività da omettere
    if "Turno" in df.columns:
        df = df[~df["Turno"].isin(turni_omit)].copy()

    # 3) Rinomina residenze
    if "Residenza" in df.columns:
        df["Residenza"] = df["Residenza"].replace(RESIDENZA_RENAME)

    # 4) Sostituzione sigle di assenza -> "Assente" in una colonna STATO
    df["Stato"] = df["Turno"].apply(lambda x: "Assente" if str(x).strip() in ABSENCE_CODES else "")

    # 5) Ordinamento (robusto sui campi presenti)
    sort_cols = [c for c in DEFAULT_SORT if c in df.columns]
    df_sorted = df.sort_values(by=sort_cols + ["Cognome e Nome"] if "Cognome e Nome" in df.columns else sort_cols,
                               kind="mergesort").reset_index(drop=True)

    # 6) Riepiloghi tipo COUNTIF sulle sigle (come nel foglio Riepiloghi)
    #    Facciamo una tabellina: Sigla, Conteggio
    riepilogo = (
        df_sorted["Turno"]
        .astype(str).str.strip()
        .pipe(lambda s: s[s.isin(ABSENCE_CODES)])
        .value_counts(dropna=False)
        .rename_axis("Sigla")
        .reset_index(name="Conteggio")
        .sort_values(by="Sigla")
        .reset_index(drop=True)
    )

    return df_sorted, riepilogo
