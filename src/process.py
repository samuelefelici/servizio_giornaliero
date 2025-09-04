import csv
import io
from pathlib import Path
import pandas as pd
from dateutil import parser
from .constants import (
    ABSENCE_CODES, RESIDENZA_RENAME, DEFAULT_SORT,
    HEADER_PROBE, EXPECTED_COLUMNS
)
from .utils import find_header_row, coerce_time, clean_spaces


# ----------------------- Sniffer di formato -----------------------

def _is_zip(b: bytes) -> bool:
    # XLSX/OOXML zip magic
    return len(b) >= 4 and b[:4] == b"PK\x03\x04"

def _is_ole(b: bytes) -> bool:
    # XLS OLE2/CFB magic
    return len(b) >= 8 and b[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

def _is_probably_text(b: bytes) -> bool:
    # Molti "xls" legacy sono in realtà TSV/CSV (anche UTF-16)
    # Heuristics: presenza di tab/comma/semicolon in prime righe o BOM UTF-16
    if b.startswith(b"\xff\xfe") or b.startswith(b"\xfe\xff"):
        return True  # UTF-16 testo
    # se contiene molti NUL, potrebbe essere UTF-16 senza BOM
    if b[:2000].count(b"\x00") > 50:
        return True
    txt = b[:4096].decode("utf-8", errors="ignore")
    seps = ["\t", ";", ",", "|"]
    return any(s in txt for s in seps)

def _is_html(b: bytes) -> bool:
    # Excel "Salva come pagina web" con estensione .xls
    head = b[:1024].lstrip().lower()
    return head.startswith(b"<!doctype html") or head.startswith(b"<html")

def _is_xml_spreadsheetml(b: bytes) -> bool:
    # Excel 2003 XML (SpreadsheetML). Facoltativo, qui lo trattiamo come testo/HTML.
    head = b[:2048].lstrip().lower()
    return head.startswith(b"<?xml") and b"<workbook" in head and b"spreadsheetml" in head


def _guess_text_delimiter(text: str) -> str:
    # Sceglie il separatore più probabile tra tab, ;, , , |
    lines = [ln for ln in text.splitlines()[:10] if ln.strip()]
    cand = ["\t", ";", ",", "|"]
    counts = {c: 0 for c in cand}
    for ln in lines:
        for c in cand:
            counts[c] += ln.count(c)
    # default: tab
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else "\t"


def _read_text_table(raw: bytes) -> pd.DataFrame:
    """
    Legge file testuali (TSV/CSV anche sporchi, UTF-16/UTF-8/CP1252).
    Usa csv.reader per gestire righe con numero di colonne variabile,
    poi pad a destra per uniformare la larghezza.
    """
    # 1) Decodifica robusta
    txt = None
    for enc in ("utf-16", "utf-8-sig", "cp1252", "utf-8", "latin-1"):
        try:
            txt = raw.decode(enc)
            break
        except Exception:
            continue
    if txt is None:
        txt = raw.decode("utf-8", errors="replace")

    # 2) Candidati separatori (tab è il default più probabile)
    candidate_seps = ["\t", ";", ",", "|"]

    # 3) Prova i separatori in ordine, scegli il primo che produce una tabella "decente"
    best_df = None
    for sep in candidate_seps:
        try:
            rows = list(csv.reader(io.StringIO(txt), delimiter=sep))
            if not rows:
                continue
            width = max(len(r) for r in rows)
            # Se width è 1, probabilmente sep sbagliato → prova il prossimo
            if width < 2:
                continue
            # Pad a destra
            for r in rows:
                if len(r) < width:
                    r.extend([""] * (width - len(r)))
            df = pd.DataFrame(rows)
            # euristica: accettiamo se almeno 5 colonne e almeno 5 righe
            if df.shape[1] >= 5 and df.shape[0] >= 5:
                best_df = df
                break
            # altrimenti conserva come possibile fallback "meno buono"
            if best_df is None:
                best_df = df
        except Exception:
            continue

    if best_df is None:
        # Estremo fallback: whitespace variabile
        # (meno preciso, ma meglio che niente)
        df = pd.read_csv(io.StringIO(txt), sep=r"\s+", header=None, dtype=str, keep_default_na=False, engine="python")
        best_df = df

    return best_df


def _read_html_table(raw: bytes) -> pd.DataFrame:
    # Richiede lxml nei requirements
    txt = raw.decode("utf-8", errors="ignore")
    tables = pd.read_html(io.StringIO(txt), header=None, flavor="lxml")
    # Prendiamo la prima tabella "significativa"
    df = max(tables, key=lambda d: d.shape[1] * d.shape[0])
    return df


def _read_excel_robusto(file) -> tuple[pd.DataFrame, str]:
    """
    Ritorna (df_raw, origine), dove origine ∈ {"xls-ole","xlsx-zip","text","html","xml","unknown"}
    Supporta:
      - path stringa/Path
      - stream/UploadedFile
    """
    # 1) Se è un path, usiamo estensione per engine
    if isinstance(file, (str, Path)):
        ext = Path(file).suffix.lower()
        if ext == ".xls":
            try:
                return pd.read_excel(file, header=None, engine="xlrd"), "xls-ole"
            except Exception:
                # potrebbe essere HTML/TSV travestito
                raw = Path(file).read_bytes()
                if _is_html(raw):
                    return _read_html_table(raw), "html"
                if _is_probably_text(raw):
                    return _read_text_table(raw), "text"
                raise
        else:
            # assume xlsx/ooxml
            return pd.read_excel(file, header=None, engine="openpyxl"), "xlsx-zip"

    # 2) Se è uno stream (Streamlit UploadedFile o simili): leggi i bytes
    if hasattr(file, "getvalue"):
        raw = file.getvalue()
    elif hasattr(file, "read"):
        raw = file.read()
        try:
            file.seek(0)
        except Exception:
            pass
    else:
        raise ValueError("Oggetto file non supportato.")

    # 3) Sniff
    if _is_ole(raw):
        bio = io.BytesIO(raw)
        return pd.read_excel(bio, header=None, engine="xlrd"), "xls-ole"

    if _is_zip(raw):
        bio = io.BytesIO(raw)
        return pd.read_excel(bio, header=None, engine="openpyxl"), "xlsx-zip"

    if _is_html(raw):
        return _read_html_table(raw), "html"

    if _is_xml_spreadsheetml(raw):
        # Molti SpreadsheetML sono apribili anche con read_html; se serve, si può
        # implementare un parser XML dedicato. Per ora usiamo fallback testuale.
        try:
            return _read_html_table(raw), "html"
        except Exception:
            return _read_text_table(raw), "xml"

    if _is_probably_text(raw):
        return _read_text_table(raw), "text"

    # Ultimo tentativo: prova entrambi gli engine in caso di file borderline
    for eng, tag in (("openpyxl", "xlsx-zip"), ("xlrd", "xls-ole")):
        try:
            bio = io.BytesIO(raw)
            return pd.read_excel(bio, header=None, engine=eng), tag
        except Exception:
            continue

    raise ValueError(
        "Formato Excel non riconosciuto. Se il file proviene da un gestionale legacy, "
        "prova a risalvarlo come .xlsx oppure esportalo come CSV/TSV."
    )


# ----------------------- Logica applicativa -----------------------

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
    date_str = str(df_raw.iloc[0, 0]).strip() if df_raw.shape[1] > 0 else ""
    day_str  = str(df_raw.iloc[1, 0]).strip() if df_raw.shape[1] > 0 and len(df_raw) > 1 else ""
    try:
        dt = parser.parse(date_str, dayfirst=True).date()
        date_str = dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return date_str, day_str


def read_input_excel(file) -> tuple[pd.DataFrame, dict]:
    # Usa il lettore robusto che sceglie la strategia corretta
    df_raw, origine = _read_excel_robusto(file)

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

    meta = {"data": date_str, "giorno": day_str, "origine": origine}
    return df, meta


def transform_dataframe(df: pd.DataFrame, config_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    matricole_omit, turni_omit = load_config_tables(config_dir)

    # 1) Matricole da omettere
    if "Matricola" in df.columns and len(matricole_omit):
        df = df[~df["Matricola"].isin(matricole_omit)].copy()

    # 2) Turni/attività da omettere
    if "Turno" in df.columns and len(turni_omit):
        df = df[~df["Turno"].isin(turni_omit)].copy()

    # 3) Rinomina residenze
    if "Residenza" in df.columns:
        df["Residenza"] = df["Residenza"].replace(RESIDENZA_RENAME)

    # 4) Stato = "Assente" se Turno ∈ ABSENCE_CODES
    if "Turno" in df.columns:
        df["Stato"] = df["Turno"].astype(str).str.strip().apply(
            lambda x: "Assente" if x in ABSENCE_CODES else ""
        )
    else:
        df["Stato"] = ""

    # 5) Ordinamento
    sort_cols = [c for c in DEFAULT_SORT if c in df.columns]
    by = sort_cols.copy()
    if "Cognome e Nome" in df.columns:
        by.append("Cognome e Nome")

    df_sorted = df.sort_values(by=by, kind="mergesort").reset_index(drop=True) if by else df.reset_index(drop=True)

    # 6) Riepilogo sigle assenza
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
