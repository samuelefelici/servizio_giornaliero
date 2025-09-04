import io
from pathlib import Path
import pandas as pd
from dateutil import parser
from .constants import (
    ABSENCE_CODES, RESIDENZA_RENAME, DEFAULT_SORT,
    HEADER_PROBE, EXPECTED_COLUMNS
)
from .utils import find_header_row, coerce_time, clean_spaces

# --- helper: sniff dei bytes per scegliere l'engine
def _pick_engine_from_bytes(b: bytes) -> str | None:
    # XLSX (ZIP): inizia con PK\x03\x04
    if len(b) >= 4 and b[:4] == b"PK\x03\x04":
        return "openpyxl"
    # XLS (OLE CF): D0 CF 11 E0 A1 B1 1A E1
    if len(b) >= 8 and b[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
        return "xlrd"
    # XML 2003 o HTML "mascherati": non gestiamo qui, proveremo gli engine standard
    return None

def _read_excel_robusto(file) -> pd.DataFrame:
    """
    Legge qualsiasi UploadedFile/stream/percorso provando:
    - engine dedotto dai bytes se disponibile
    - openpyxl -> xlrd (o viceversa) come fallback
    """
    # 1) se è un path stringa, usiamo il percorso direttamente
    if isinstance(file, (str, Path)):
        ext = Path(file).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        try:
            return pd.read_excel(file, header=None, engine=engine)
        except Exception:
            # fallback inverso
            alt = "openpyxl" if engine == "xlrd" else "xlrd"
            return pd.read_excel(file, header=None, engine=alt)

    # 2) se è un oggetto tipo Streamlit UploadedFile / file-like: prendo i bytes
    raw_bytes = None
    if hasattr(file, "getvalue"):         # Streamlit UploadedFile
        raw_bytes = file.getvalue()
    elif hasattr(file, "read"):            # file-like
        raw_bytes = file.read()
        try:
            file.seek(0)
        except Exception:
            pass
    else:
        raise ValueError("Oggetto file non supportato.")

    bio = io.BytesIO(raw_bytes)
    guessed = _pick_engine_from_bytes(raw_bytes)

    # Prova 1: engine “sniffato” (se c'è)
    if guessed:
        try:
            bio.seek(0)
            return pd.read_excel(bio, header=None, engine=guessed)
        except Exception:
            pass

    # Prova 2: openpyxl
    try:
        bio.seek(0)
        return pd.read_excel(bio, header=None, engine="openpyxl")
    except Exception:
        pass

    # Prova 3: xlrd (per .xls)
    try:
        bio.seek(0)
        return pd.read_excel(bio, header=None, engine="xlrd")
    except Exception as e:
        raise ValueError(
            "Formato Excel non riconosciuto: prova a salvare il file come .xlsx oppure "
            "verifica che non sia un HTML/XML camuffato."
        ) from e


def parse_date_and_day(df_raw: pd.DataFrame) -> tuple[str, str]:
    date_str = str(df_raw.iloc[0, 0]).strip() if df_raw.shape[1] > 0 else ""
    day_str  = str(df_raw.iloc[1, 0]).strip() if df_raw.shape[1] > 0 and len(df_raw) > 1 else ""
    try:
        dt = parser.parse(date_str, dayfirst=True).date()
        date_str = dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return date_str, day_str

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

def read_input_excel(file) -> tuple[pd.DataFrame, dict]:
    # <-- QUI il cambio: usiamo il lettore robusto che sceglie l'engine
    df_raw = _read_excel_robusto(file)

    date_str, day_str = parse_date_and_day(df_raw)
    hdr_row = find_header_row(df_raw, HEADER_PROBE)
    if hdr_row is None:
        raise ValueError(f"Intestazione '{HEADER_PROBE}' non trovata.")

    # Taglia dal header in poi, limitando al numero atteso di colonne
    df = df_raw.iloc[hdr_row:, : len(EXPECTED_COLUMNS)].copy()
    df.columns = df.iloc[0].tolist()
    df = df.iloc[1:].reset_index(drop=True)

    keep_cols = [c for c in EXPECTED_COLUMNS if c in df.columns]
    df = df[keep_cols].copy()
    df = clean_spaces(df)

    if "Matricola" in df.columns:
        df["Matricola"] = df["Matricola"].astype(str).str.strip()

    if "Inizio" in df.columns:
        df["Inizio"] = coerce_time(df["Inizio"])
    if "Fine" in df.columns:
        df["Fine"] = coerce_time(df["Fine"])

    meta = {"data": date_str, "giorno": day_str}
    return df, meta

def transform_dataframe(df: pd.DataFrame, config_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    matricole_omit, turni_omit = load_config_tables(config_dir)

    if "Matricola" in df.columns and len(matricole_omit):
        df = df[~df["Matricola"].isin(matricole_omit)].copy()

    if "Turno" in df.columns and len(turni_omit):
        df = df[~df["Turno"].isin(turni_omit)].copy()

    if "Residenza" in df.columns:
        df["Residenza"] = df["Residenza"].replace(RESIDENZA_RENAME)

    if "Turno" in df.columns:
        df["Stato"] = df["Turno"].astype(str).str.strip().apply(lambda x: "Assente" if x in ABSENCE_CODES else "")
    else:
        df["Stato"] = ""

    sort_cols = [c for c in DEFAULT_SORT if c in df.columns]
    by = sort_cols.copy()
    if "Cognome e Nome" in df.columns:
        by.append("Cognome e Nome")

    df_sorted = df.sort_values(by=by, kind="mergesort").reset_index(drop=True) if by else df.reset_index(drop=True)

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
