# src/pdf_export.py
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
import pandas as pd
from pathlib import Path
import re

from .constants import DISPLAY_ORDER

# ---------- helper prefissi deposito/turno ----------
def _res_to_prefix(res: str) -> str | None:
    if not isinstance(res, str): return None
    r = res.upper().replace("_", " ").strip()
    if "JESI URBANO" in r or r == "JU": return "JU"
    if "JESI" in r or r == "J":       return "J"
    if "MARINA" in r or r == "M":     return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r == "C": return "C"
    if "OSIMO" in r or r == "O":      return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r == "F":    return "F"
    if "POLVERIGI" in r or r == "P":  return "P"
    if "OSTRA" in r or r == "D":      return "D"
    if "BELVED" in r or r == "B" or "DEPBELVE" in r:     return "B"
    if "ANCONA" in r or r == "A":     return "A"
    return None

def _turno_bucket(turno: str) -> str | None:
    if not isinstance(turno, str): turno = str(turno)
    s = turno.strip()
    if not s or s.upper() == "ASSENTE": return None
    m = re.match(r"[A-Za-z]+", s)
    if not m: return None
    up = m.group(0).upper()
    if up.startswith("JU"): return "JU"
    if up.startswith("J"):  return "J"
    return up[0]

def _accepted_prefixes_for_res(prefix: str | None) -> set[str]:
    if prefix in ("J", "JU"): return {"J", "JU"}
    return {prefix} if prefix else set()

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)
    expected = df["Residenza"].apply(_res_to_prefix)
    actual   = df["Turno"].astype(str).apply(_turno_bucket)
    ok = expected.apply(_accepted_prefixes_for_res)
    return expected.notna() & actual.notna() & (~ok.apply(lambda s: actual.iloc[s.name] in s))

# ---------- shaping tabella ----------
def _collapse_repeats(gdf: pd.DataFrame,
                      key_cols=("Cognome e Nome", "Matricola"),
                      collapse_cols=("Cognome e Nome", "Matricola")) -> pd.DataFrame:
    missing = [c for c in key_cols if c not in gdf.columns]
    if missing: return gdf
    g = gdf.copy()
    dup_mask = (g[list(key_cols)] == g[list(key_cols)].shift(1)).all(axis=1)
    for c in collapse_cols:
        if c in g.columns:
            g.loc[dup_mask, c] = ""
    return g

def _table_data_for(df: pd.DataFrame):
    cols = [c for c in DISPLAY_ORDER if c in df.columns]
    dfp = df.copy()
    if cols: dfp = dfp[cols]
    dfp = dfp.drop(columns=["Residenza"], errors="ignore")
    header = list(dfp.columns)
    rows = dfp.fillna("").values.tolist()
    return [header] + rows

def _col_widths(header: list[str], avail_width: float) -> list[float]:
    """
    Usa tutta la larghezza disponibile del foglio.
    Tutte le colonne hanno larghezza fissa tranne 'Indennità e note' che prende il resto.
    Se manca 'Indennità e note', si ripartisce proporzionalmente.
    """
    # fissi (puoi ritoccarli)
    FIXED = {
        "Matricola": 22*mm,
        "Cognome e Nome": 62*mm,
        "Turno": 24*mm,
        "Inizio": 18*mm,
        "Fine": 18*mm,
    }
    NOTE_NAMES = {"Indennità e note", "Note"}

    fixed_sum = sum(FIXED.get(h, 0) for h in header if h not in NOTE_NAMES)
    has_note = any(h in NOTE_NAMES for h in header)

    widths = []
    if has_note:
        min_note = 30*mm
        if fixed_sum > max(0, avail_width - min_note) and fixed_sum > 0:
            # scala i fissi per lasciare almeno min_note alle note
            scale = (avail_width - min_note) / fixed_sum
            for h in header:
                if h in NOTE_NAMES:
                    widths.append(min_note)  # placeholder; aggiustiamo dopo
                else:
                    widths.append(FIXED.get(h, 18*mm) * scale)
            # ricalcola resto per la colonna Note (tutta la parte rimanente)
            used = sum(w for h, w in zip(header, widths) if h not in NOTE_NAMES)
            note_w = max(min_note, avail_width - used)
            widths = [note_w if h in NOTE_NAMES else w for h, w in zip(header, widths)]
        else:
            # fissi come da mappa, resto alla colonna Note
            used = sum(FIXED.get(h, 18*mm) for h in header if h not in NOTE_NAMES)
            note_w = max(min_note, avail_width - used)
            for h in header:
                widths.append(note_w if h in NOTE_NAMES else FIXED.get(h, 18*mm))
    else:
        # nessuna colonna note: ripartizione proporzionale
        if fixed_sum == 0:
            widths = [avail_width/len(header)]*len(header)
        else:
            for h in header:
                w = FIXED.get(h, 18*mm) / fixed_sum * avail_width
                widths.append(w)
    return widths

# ---------- build PDF ----------
def build_pdf(path_out: Path, df: pd.DataFrame, meta: dict,
              logo_path: Path | None = None, title: str = "Servizio Giornaliero",
              inner_sort: str = "nome"):
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
            if "Cognome e Nome" in gdf.columns: by.append("Cognome e Nome")
            if "Inizio" in gdf.columns:         by.append("Inizio")
            if by: gdf = gdf.sort_values(by=by).reset_index(drop=True)
            gdf = _collapse_repeats(gdf, key_cols=("Cognome e Nome", "Matricola"),
                                    collapse_cols=("Cognome e Nome", "Matricola"))

        trasferte = _trasferta_mask(gdf)

        if not first_block:
            elems.append(Spacer(1, 3*mm))
        first_block = False

        # Titolo gruppo
        elems.append(Paragraph(res_name, group_style))

        # Tabella dati
        data = _table_data_for(gdf)
        header = data[0]

        # col widths: usa tutta la riga; solo "Indennità e note" si adatta
        avail_width = A4[0] - doc.leftMargin - doc.rightMargin
        col_widths = _col_widths(header, avail_width)

        # Indici dinamici per colonne interessate
        col_idx = {name: header.index(name) for name in ["Turno", "Inizio", "Fine"] if name in header}
        idx_inizio = col_idx.get("Inizio")
        idx_fine   = col_idx.get("Fine")

        tbl = Table(data, repeatRows=1, colWidths=col_widths)

        # Stile base
        base_style = [
            ("GRID",       (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN",     (0, 0), (-1, -1), "TOP"),
        ]
        if idx_inizio is not None:
            base_style.append(("ALIGN", (idx_inizio, 1), (idx_inizio, -1), "CENTER"))
        if idx_fine is not None:
            base_style.append(("ALIGN", (idx_fine, 1), (idx_fine, -1), "CENTER"))

        # Grassetto per trasferte solo su Turno/Inizio/Fine
        for i, is_tr in enumerate(trasferte.tolist(), start=1):  # +1 per saltare l'header
            if not is_tr: continue
            for cidx in col_idx.values():
                base_style.append(("FONTNAME", (cidx, i), (cidx, i), "Helvetica-Bold"))

        tbl.setStyle(TableStyle(base_style))
        elems.append(tbl)

    doc.build(elems)
