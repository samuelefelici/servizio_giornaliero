import io, csv
from pathlib import Path
import pandas as pd
from dateutil import parser
from .constants import (
    ABSENCE_CODES, RESIDENZA_RENAME, DEFAULT_SORT,
    HEADER_PROBE, EXPECTED_COLUMNS
)
from .utils import find_header_row, coerce_time, clean_spaces, clean_column_names

# ----------------------- Sniffer/lettura robusta -----------------------

def _is_zip(b: bytes) -> bool:
    # XLSX/OOXML zip magic
    return len(b) >= 4 and b[:4] == b"PK\x03\x04"

def _is_ole(b: bytes) -> bool:
    # XLS OLE2/CFB magic
    return len(b) >= 8 and b[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

def _is_html(b: bytes) -> bool:
    head = b[:1024].lstrip().lower()
    return head.startswith(b"<!doctype html") or head.startswith(b"<html")

def _is_xml_spreadsheetml(b: bytes) -> bool:
    head = b[:2048].lstrip().lower()
    return head.startswith(b"<?xml") and b"<workbook" in head and b"spreadsheetml" in head

def _is_probably_text(b: bytes) -> bool:
    # molti “.xls” legacy sono TSV/CSV o UTF-16
    if b.startswith(b"\xff\xfe") or b.startswith(b"\xfe\xff"):
        return True  # UTF-16 BOM
    if b[:2000].count(b"\x00") > 50:
        return True  # probabile UTF-16 senza BOM
    txt = b[:4096].decode("utf-8", errors="ignore")
    return any(sep in txt for sep in ["\t", ";", ",", "|"])

def _read_text_table(raw: bytes) -> pd.DataFrame:
    """
    Legge file testuali (TSV/CSV anche ‘sporchi’) con righe a larghezza variabile:
    usa csv.reader e fa padding a destra.
    """
    # decodifica robusta
    txt = None
    for enc in ("utf-16", "utf-8-sig", "cp1252", "utf-8", "latin-1"):
        try:
            txt = raw.decode(enc)
            break
        except Exception:
            continue
    if txt is None:
        txt = raw.decode("utf-8", errors="replace")

    candidate_seps = ["\t", ";", ",", "|"]
    best_df = None
    for sep in candidate_seps:
        try:
            rows = list(csv.reader(io.StringIO(txt), delimiter=sep))
            if not rows:
                continue
            width = max(len(r) for r in rows)
            if width < 2:
                continue
            # pad a destra
            for r in rows:
                if len(r) < width:
                    r.extend([""] * (width - len(r)))
            df = pd.DataFrame(rows)
            # euristica: accetta tabelle con almeno 5x5
            if df.shape[1] >= 5 and df.shape[0] >= 5:
                best_df = df
                break
            if best_df is None:
                best_df = df
        except Exception:
            continue

    if best_df is None:
        # fallback con separatore whitespace
        df = pd.read_csv(io.StringIO(txt), sep=r"\s+", header=None, dtype=str, keep_default_na=False, engine="python")
        best_df = df
    return best_df

def _read_html_table(raw: bytes) -> pd.DataFrame:
    # richiede lxml nei requirements
    txt = raw.decode("utf-8", errors="ignore")
    tables = pd.read_html(io.StringIO(txt), header=None, flavor="lxml")
    # prendi la più grande
    return max(tables, key=lambda d: d.shape[0] * d.shape[1])

def _read_excel_robusto(file) -> tuple[pd.DataFrame, str]:
    """
    Ritorna (df_raw, origine) dove origine ∈ {xls-ole, xlsx-zip, text, html, xml}.
    - path stringa/Path → usa estensione + fallback su sniff
    - stream (UploadedFile) → sniff sui magic bytes
    """
    # 1) path
    if isinstance(file, (str, Path)):
        p = Path(file)
        ext = p.suffix.lower()
        if ext == ".xls":
            try:
                return pd.read_excel(p, header=None, engine="xlrd"), "xls-ole"
            except Exception:
                raw = p.read_bytes()
                if _is_html(raw):
                    return _read_html_table(raw), "html"
                if _is_probably_text(raw):
                    return _read_text_table(raw), "text"
                raise
        else:
            # assume xlsx
            return pd.read_excel(p, header=None, engine="openpyxl"), "xlsx-zip"

    # 2) stream
    if hasattr(file, "getvalue"):
        raw = file.getvalue()
    elif hasattr(file, "read"):
        raw = file.read()
        try:
            file.seek(0)
        except Exception:
            pass
    else:
        raise ValueError("Oggetto file non supportato (né path né file-like).")

    # 3) sniff
    if _is_ole(raw):
        return pd.read_excel(io.BytesIO(raw), header=None, engine="xlrd"), "xls-ole"
    if _is_zip(raw):
        return pd.read_excel(io.BytesIO(raw), header=None, engine="openpyxl"), "xlsx-zip"
    if _is_html(raw):
        return _read_html_table(raw), "html"
    if _is_xml_spreadsheetml(raw):
        try:
            return _read_html_table(raw), "html"
        except Exception:
            return _read_text_table(raw), "xml"
    if _is_probably_text(raw):
        return _read_text_table(raw), "text"

    # 4) tentativi finali
    for eng, tag in (("openpyxl", "xlsx-zip"), ("xlrd", "xls-ole")):
        try:
            return pd.read_excel(io.BytesIO(raw), header=None, engine=eng), tag
        except Exception:
            continue

    raise ValueError("Formato Excel non riconosciuto. Prova a risalvare come .xlsx o esporta CSV/TSV.")

# ----------------------- Parsing base -----------------------

def parse_date_and_day(df_raw: pd.DataFrame) -> tuple[str, str]:
    date_str = str(df_raw.iloc[0, 0]).strip() if df_raw.shape[1] > 0 else ""
    day_str  = str(df_raw.iloc[1, 0]).strip() if df_raw.shape[1] > 0 and len(df_raw) > 1 else ""
    try:
        dt = parser.parse(date_str, dayfirst=True).date()
        date_str = dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return date_str, day_str

def read_input_excel(file) -> tuple[pd.DataFrame, dict]:
    """
    Legge il file, pulisce spazi/NBSP, trova l'intestazione,
    seleziona le colonne attese, rimuove righe vuote, normalizza orari.
    """
    df_raw, origine = _read_excel_robusto(file)

    # Pulizia preliminare (trim su stringhe e NBSP)
    df_raw = clean_spaces(df_raw)

    # Meta (data e giorno) dalle prime righe
    date_str, day_str = parse_date_and_day(df_raw)

    # Trova la riga di intestazione (es. "Cognome e Nome")
    hdr_row = find_header_row(df_raw, HEADER_PROBE)
    if hdr_row is None:
        raise ValueError(f"Intestazione '{HEADER_PROBE}' non trovata.")

    # Prendi tutte le colonne dall'header in poi
    df = df_raw.iloc[hdr_row:, :].copy()
    # Nomi colonna ripuliti
    df.columns = clean_column_names(df.iloc[0].tolist())
    # Rimuove la riga header duplicata
    df = df.iloc[1:].reset_index(drop=True)

    # Tieni solo le colonne attese che esistono davvero
    keep_cols = [c for c in EXPECTED_COLUMNS if c in df.columns]
    if not keep_cols:
        raise ValueError(
            "Nessuna delle colonne attese è presente. "
            "Controlla che l'intestazione contenga almeno una tra: "
            + ", ".join(EXPECTED_COLUMNS)
        )
    df = df[keep_cols].copy()

    # Pulizia finale su celle stringa
    df = clean_spaces(df)

    # ✅ Elimina righe completamente vuote sulle colonne utili
    def _row_is_empty(row) -> bool:
        for v in row:
            if pd.isna(v):
                continue
            if isinstance(v, str) and v.strip() == "":
                continue
            return False
        return True

    df = df[~df.apply(_row_is_empty, axis=1)].reset_index(drop=True)

    # Tipizzazioni / orari
    if "Matricola" in df.columns:
        df["Matricola"] = df["Matricola"].astype(str).str.strip()
    if "Inizio" in df.columns:
        df["Inizio"] = coerce_time(df["Inizio"])
    if "Fine" in df.columns:
        df["Fine"] = coerce_time(df["Fine"])

    meta = {"data": date_str, "giorno": day_str, "origine": origine}
    return df, meta


# ----------------------- Trasformazione (senza riepilogo) -----------------------

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

def transform_dataframe(df: pd.DataFrame, config_dir: Path) -> pd.DataFrame:
    """
    1) filtro matricole/turni da omettere
    2) rinomina Residenza
    3) sigle assenza -> 'Assente' in Turno
    4) ordinamento per DEFAULT_SORT (+ 'Cognome e Nome')
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

    # 3) Turno -> 'Assente' se sigla tra le assenze
    if "Turno" in df.columns:
        df["Turno"] = (
            df["Turno"]
            .astype(str).str.strip()
            .apply(lambda x: "Assente" if x in ABSENCE_CODES else x)
        )

    # 4) Ordinamento
    sort_cols = [c for c in DEFAULT_SORT if c in df.columns]
    by = sort_cols.copy()
    if "Cognome e Nome" in df.columns:
        by.append("Cognome e Nome")
    df_sorted = df.sort_values(by=by, kind="mergesort").reset_index(drop=True) if by else df.reset_index(drop=True)

    return df_sorted

__all__ = ["read_input_excel", "transform_dataframe"]
