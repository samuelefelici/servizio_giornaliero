import io
from pathlib import Path
import pandas as pd
from dateutil import parser
from .constants import (
    ABSENCE_CODES, RESIDENZA_RENAME, DEFAULT_SORT,
    HEADER_PROBE, EXPECTED_COLUMNS
)
from .utils import find_header_row, coerce_time, clean_spaces


# --- Config tables: gestisci assenza dei CSV senza errori
def load_config_tables(config_dir: Path):
    m_path = config_dir / "matricole_da_omettere.csv"
    t_path = config_dir / "turni_attivita_da_omettere.csv"

    if m_path.exists():
        m_omit = pd.read_csv(m_path, dtype=str, keep_default_na=False)
        matricole = set(m_omit.get("Matricola", pd.Series(dtype=str)).astype(str).str.strip())
    else:
        matricole = set()

    if t_path.exists():
        t_omit = pd.read_csv(t_path, dtype=str, keep_default_na=False)
        turni = set(t_omit.get("Turno", pd.Series(dtype=str)).astype(str).str.strip())
    else:
        turni = set()

    return matricole, turni


def parse_date_and_day(df_raw: pd.DataFrame) -> tuple[str, str]:
    # Prima riga: data (es. 05/09/2025); seconda riga: giorno (es. Venerdi)
    date_str = str(df_raw.iloc[0, 0]).strip() if df_raw.shape[1] > 0 else ""
    day_str = str(df_raw.iloc[1, 0]).strip() if df_raw.shape[1] > 0 and len(df_raw) > 1 else ""
    try:
        dt = parser.parse(date_str, dayfirst=True).date()
        date_str = dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return date_str, day_str


def _pick_engine(file_obj) -> str | None:
    """Sceglie l'engine in base all'estensione (se disponibile)."""
    name = getattr(file_obj, "name", None)
    if not name:
        return None
    ext = Path(name).suffix.lower()
    if ext == ".xls":
        return "xlrd"
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return "openpyxl"
    return None


def read_input_excel(file) -> tuple[pd.DataFrame, dict]:
    """
    Legge il file Excel (supporta UploadedFile di Streamlit).
    - cerca automaticamente la riga di intestazione con HEADER_PROBE
    - normalizza colonne e orari
    """
    engine = _pick_engine(file)

    # Prova principale con engine scelto; in caso di fallimento, fallback lasciando scegliere a pandas.
    try:
        df_raw = pd.read_excel(file, header=None, engine=engine)
    except Exception:
        # Se è uno stream (es. UploadedFile), riportiamo il puntatore all'inizio
        if hasattr(file, "seek"):
            try:
                file.seek(0)
            except Exception:
                pass
        df_raw = pd.read_excel(file, header=None)

    date_str, day_str = parse_date_and_day(df_raw)
    hdr_row = find_header_row(df_raw, HEADER_PROBE)
    if hdr_row is None:
        raise ValueError(f"Intestazione '{HEADER_PROBE}' non trovata.")

    # Taglia dal header in poi, e limita al numero atteso di colonne (se il file ne ha di più)
    df = df_raw.iloc[hdr_row:, : len(EXPECTED_COLUMNS)].copy()
    df.columns = df.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)

    # Tieni solo le colonne previste che esistono davvero
    keep_cols = [c for c in EXPECTED_COLUMNS if c in df.columns]
    df = df[keep_cols].copy()
    df = clean_spaces(df)

    # Tipizzazioni base
    if "Matricola" in df.columns:
        df["Matricola"] = df["Matricola"].astype(str).str.strip()

    # Orari
    if "Inizio" in df.columns:
        df["Inizio"] = coerce_time(df["Inizio"])
    if "Fine" in df.columns:
        df["Fine"] = coerce_time(df["Fine"])

    meta = {"data": date_str, "giorno": day_str}
    return df, meta


def transform_dataframe(df: pd.DataFrame, config_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    matricole_omit, turni_omit = load_config_tables(config_dir)

    # 1) Filtri: matricole da omettere
    if "Matricola" in df.columns and len(matricole_omit):
        df = df[~df["Matricola"].isin(matricole_omit)].copy()

    # 2) Filtri: turni/attività da omettere
    if "Turno" in df.columns and len(turni_omit):
        df = df[~df["Turno"].isin(turni_omit)].copy()

    # 3) Rinomina residenze
    if "Residenza" in df.columns:
        df["Residenza"] = df["Residenza"].replace(RESIDENZA_RENAME)

    # 4) Stato = "Assente" se Turno in ABSENCE_CODES
    if "Turno" in df.columns:
        df["Stato"] = df["Turno"].astype(str).str.strip().apply(
            lambda x: "Assente" if x in ABSENCE_CODES else ""
        )
    else:
        df["Stato"] = ""

    # 5) Ordinamento (robusto)
    sort_cols = [c for c in DEFAULT_SORT if c in df.columns]
    by = sort_cols.copy()
    if "Cognome e Nome" in df.columns:
        by.append("Cognome e Nome")

    if by:
        df_sorted = df.sort_values(by=by, kind="mergesort").reset_index(drop=True)
    else:
        df_sorted = df.reset_index(drop=True)

    # 6) Riepiloghi (conteggio delle sigle di assenza)
    if "Turno" in df_sorted.columns:
        riepilogo = (
            df_sorted["Turno"].astype(str).str.strip()
            .pipe(lambda s: s[s.isin(ABSENCE_CODES)])
            .value_counts(dropna=False)
            .rename_axis("Sigla")
            .reset_index(name="Conteggio")
            .sort_values(by="Sigla")
            .reset_index(drop=True)
        )
    else:
        riepilogo = pd.DataFrame(columns=["Sigla", "Conteggio"])

    return df_sorted, riepilogo
