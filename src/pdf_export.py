# src/pdf_export.py
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
import pandas as pd
from pathlib import Path
import re

from .constants import DISPLAY_ORDER  # Matricola, Cognome e Nome, Turno, Inizio, Fine, Indennità e note

REST_CODES = {"R", "RR"}  # turni di riposo: mai in grassetto

# ---------- helper prefissi deposito/turno ----------
def _res_to_prefix(res: str) -> str | None:
    """Mappa il testo della residenza a un prefisso 'atteso' (J e JU sono famigliari)."""
    if not isinstance(res, str):
        return None
    r = res.upper().replace("_", " ").strip()
    if "JESI URBANO" in r or r == "JU": return "JU"
    if "JESI" in r or r == "J":         return "J"
    if "MARINA" in r or r == "M":       return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r == "C": return "C"
    if "OSIMO" in r or r == "O":        return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r == "F":    return "F"
    if "POLVERIGI" in r or r == "P":    return "P"
    if "OSTRA" in r or r == "D":        return "D"
    if "BELVED" in r or r == "B" or "DEPBELVE" in r:     return "B"
    if "ANCONA" in r or r == "A":       return "A"
    return None

def _turno_bucket(turno: str) -> str | None:
    """
    Riduce il codice turno a un 'bucket' per confronto:
      - 'JU…' -> 'JU'
      - 'J…'  (non JU) -> 'J'
      - altri -> prima lettera (M, C, O, F, P, D, A, B…)
    Esclude 'Assente' e i riposi R/RR.
    """
    if not isinstance(turno, str):
        turno = str(turno)
    s = turno.strip()
    if not s:
        return None
    up_full = s.upper()
    if up_full == "ASSENTE" or up_full in REST_CODES:
        return None
    m = re.match(r"[A-Za-z]+", s)
    if not m:
        return None
    up = m.group(0).upper()
    if up.startswith("JU"): return "JU"
    if up.startswith("J"):  return "J"
    return up[0]

def _accepted_prefixes_for_res(prefix: str | None) -> set[str]:
    """Prefissi considerati 'di casa' per la residenza (J e JU sono equivalenti)."""
    if prefix in ("J", "JU"):
        return {"J", "JU"}
    return {prefix} if prefix else set()

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    """
    True se la riga è in 'trasferta' (turno non appartiene ai prefissi accettati).
    Considera J e JU equivalenti; esclude 'Assente' e R/RR.
    """
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)

    def _is_trasferta(row) -> bool:
        turn = str(row["Turno"]).strip().upper()
        if turn in REST_CODES or turn == "ASSENTE":
            return False
        rp = _res_to_prefix(row["Residenza"])
        tb = _turno_bucket(row["Turno"])
        if rp is None or tb is None:
            return False
        return tb not in _accepted_prefixes_for_res(rp)

    return df.apply(_is_trasferta, axis=1)

# ---------- shaping tabella ----------
def _collapse_repeats(gdf: pd.DataFrame,
                      key_cols=("Cognome e Nome", "Matricola"),
                      collapse_cols=("Cognome e Nome", "Matricola")) -> pd.DataFrame:
    """Sui record consecutivi della stessa persona azzera i campi ripetitivi."""
    missing = [c for c in key_cols if c not in gdf.columns]
    if missing:
        return gdf
    g = gdf.copy()
    dup_mask = (g[list(key_cols)] == g[list(key_cols)].shift(1)).all(axis=1)
    for c in collapse_cols:
        if c in g.columns:
            g.loc[dup_mask, c] = ""
    return g

def _table_data_for(df: pd.DataFrame):
    """Applica DISPLAY_ORDER e rimuove 'Residenza' se presente."""
    cols = [c for c in DISPLAY_ORDER if c in df.columns]
    dfp = df.copy()
    if cols:
        dfp = dfp[cols]
    dfp = dfp.drop(columns=["Residenza"], errors="ignore")
    header = list(dfp.columns)  # manteniamo "Indennità e note"
    rows = dfp.fillna("").values.tolist()
    return [header] + rows

# ---------- build PDF ----------
def build_pdf(path_out: Path, df: pd.DataFrame, meta: dict,
              logo_path: Path | None = None, title: str = "Servizio Giornaliero",
              inner_sort: str = "nome"):
    """
    Crea un PDF raggruppato per Residenza.
    - Titolo compatto in alto
    - Per ogni Residenza: titolo gruppo e tabella (senza colonna Residenza)
    - inner_sort: "nome" (A→Z + Inizio come tie-breaker) oppure "inizio"
    - Trasferte: celle Turno/Inizio/Fine in grassetto.
    """
    doc = SimpleDocTemplate(
        str(path_out), pagesize=A4,
        rightMargin=10*mm, leftMargin=10*mm,
        topMargin=12*mm, bottomMargin=12*mm
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleTight", parent=styles["Title"], spaceAfter=0, spaceBefore=0)
    group_style = ParagraphStyle("GroupTitle", parent=styles["Heading2"],
                                 fontSize=14, leading=16, spaceBefore=6, spaceAfter=2)

    elems = []

    # Intestazione
    header_text = f"{title} — {meta.get('data','')} — {meta.get('giorno','')}"
    elems.append(Paragraph(header_text, title_style))

    if logo_path and logo_path.exists():
        elems.append(Spacer(1, 2*mm))
        img = Image(str(logo_path))
        img._restrictSize(40*mm, 15*mm)
        elems.append(img)

    # Grouping per Residenza
    if "Residenza" not in df.columns:
        blocks = [("TUTTI", df)]
    else:
        df = df.sort_values(by=["Residenza"]).reset_index(drop=True)
        blocks = [(str(res), g.copy()) for res, g in df.groupby("Residenza", sort=True)]

    first_block = True
    for res_name, gdf in blocks:
        # Ordinamento interno
        if inner_sort == "inizio" and "Inizio" in gdf.columns:
            by = ["Inizio"] + (["Cognome e Nome"] if "Cognome e Nome" in gdf.columns else [])
            gdf = gdf.sort_values(by=by).reset_index(drop=True)
        else:
            by = []
            if "Cognome e Nome" in gdf.columns:
                by.append("Cognome e Nome")
            if "Inizio" in gdf.columns:
                by.append("Inizio")
            if by:
                gdf = gdf.sort_values(by=by).reset_index(drop=True)
            gdf = _collapse_repeats(
                gdf,
                key_cols=("Cognome e Nome", "Matricola"),
                collapse_cols=("Cognome e Nome", "Matricola")
            )

        # Maschera trasferte (sull’ordine definitivo del gruppo)
        trasferte = _trasferta_mask(gdf)

        if not first_block:
            elems.append(Spacer(1, 3*mm))
        first_block = False

        # Titolo gruppo
        elems.append(Paragraph(res_name, group_style))

        # Tabella dati
        data = _table_data_for(gdf)
        header = data[0]

        # Indici dinamici per colonne interessate
        col_idx = {name: header.index(name) for name in ["Turno", "Inizio", "Fine"] if name in header}
        idx_inizio = col_idx.get("Inizio")
        idx_fine   = col_idx.get("Fine")

        tbl = Table(data, repeatRows=1)

        # Stile base
        base_style = [
            ("GRID",       (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN",     (0, 0), (-1, -1), "TOP"),
        ]
        # Allineamento Inizio/Fine al centro se presenti
        if idx_inizio is not None:
            base_style.append(("ALIGN", (idx_inizio, 1), (idx_inizio, -1), "CENTER"))
        if idx_fine is not None:
            base_style.append(("ALIGN", (idx_fine, 1), (idx_fine, -1), "CENTER"))

        # Grassetto per trasferte solo su Turno/Inizio/Fine
        for i, is_tr in enumerate(trasferte.tolist(), start=1):  # +1 per saltare l'header
            if not is_tr:
                continue
            for cidx in col_idx.values():
                base_style.append(("FONTNAME", (cidx, i), (cidx, i), "Helvetica-Bold"))

        tbl.setStyle(TableStyle(base_style))
        elems.append(tbl)

    doc.build(elems)
