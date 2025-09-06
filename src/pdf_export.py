# src/pdf_export.py
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER
from xml.sax.saxutils import escape
from datetime import datetime
from pathlib import Path
import pandas as pd
import re
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.pdfbase.pdfmetrics import stringWidth


# TZ Europe/Rome
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    _TZ_ROMA = ZoneInfo("Europe/Rome")
except Exception:  # pragma: no cover
    import pytz
    _TZ_ROMA = pytz.timezone("Europe/Rome")

from .constants import DISPLAY_ORDER

# ======================= Costanti & regole =======================
REST_CODES = {"R", "RR"}  # riposo

# Eccezioni GLOBALI: NON evidenziare mai questi turni (match esatto)
GLOBAL_EXC_PATTERNS = (
    r"^IAST$",  # esattamente IAST
    r"^N$",     # esattamente N
)

# Eccezioni specifiche per ANCONA (prefissi)
EXC_ANCONA_PREFIXES = (
    "D1R1", "D1R2", "D1R5", "D2R1", "D2R2", "D2R3", "D2R6",
    "NP", "ASC", "V5",
    "LU", "MA", "ME", "GI", "VE", "SA", "DO",
)

# Eccezioni per gli altri depositi (prefissi)
EXC_OTHER_PREFIXES = ("IAST", "N")

# Dimensioni massime logo (regolabili)
MAX_LOGO_W = 60 * mm
MAX_LOGO_H = 22.5 * mm

# ======================= Utility evidenziazione =======================
def _res_to_prefix(res: str) -> str | None:
    if not isinstance(res, str):
        return None
    r = res.upper().replace("_", " ").strip()
    if "JESI URBANO" in r or r == "JU": return "JU"
    if "JESI" in r or r == "J":        return "J"
    if "MARINA" in r or r == "M":      return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r == "C": return "C"
    if "OSIMO" in r or r == "O":       return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r == "F":    return "F"
    if "POLVERIGI" in r or r == "P":   return "P"
    if "OSTRA" in r or r == "D":       return "D"
    if "BELVED" in r or r == "B" or "DEPBELVE" in r:     return "B"
    if "ANCONA" in r or r == "A":      return "A"
    return None

def _norm_turno(s) -> str:
    return str(s or "").upper().strip().replace(".", "").replace(" ", "")

def _starts_with_any(raw_turno: str, prefixes: tuple[str, ...]) -> bool:
    t = _norm_turno(raw_turno)
    return any(t.startswith(p) for p in prefixes)

def _match_any(s: str, patterns: tuple[str, ...]) -> bool:
    return any(re.match(p, s) for p in patterns)

def _accepted_prefixes_for_res(prefix: str | None) -> set[str]:
    if prefix in {"J", "JU"}:
        return {"J", "JU"}
    return {prefix} if prefix else set()

def _should_highlight_turno(residenza, turno) -> bool:
    """
    Decide se evidenziare (grassetto) Turno/Inizio/Fine per la riga.
    """
    t = _norm_turno(turno)
    if not t or t in {"ASSENTE", *REST_CODES}:
        return False

    # Eccezioni globali (match esatto)
    if _match_any(t, GLOBAL_EXC_PATTERNS):
        return False

    rp = _res_to_prefix(residenza)

    # Eccezioni per deposito
    if rp == "A":  # ANCONA
        if _starts_with_any(t, EXC_ANCONA_PREFIXES):
            return False
    else:
        if _starts_with_any(t, EXC_OTHER_PREFIXES):
            return False

    accepted = _accepted_prefixes_for_res(rp)

    if t.startswith("JU"):
        b = "JU"
    elif t.startswith("J"):
        b = "J"
    else:
        b = t[0] if t else None

    if not rp or not b:
        return False
    return b not in accepted

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)
    return df.apply(lambda r: _should_highlight_turno(r.get("Residenza"), r.get("Turno")), axis=1)

# ======================= Header helper =======================
def _header_table(title_para: Paragraph,
                  logo_path: Path | None,
                  page_w: float,
                  export_text: str,
                  small_style: ParagraphStyle) -> Table:
    """
    Riga con titolo (sx) e colonna destra con:
    [logo centrato]
    [nota allineata a destra]
    La larghezza della colonna destra viene calcolata in base alla nota,
    così il testo arriva esattamente al margine destro pagina.
    """
    # misura la larghezza del testo (pt)
    font_name = getattr(small_style, "fontName", "Helvetica")
    font_size = getattr(small_style, "fontSize", 8)
    text_w    = stringWidth(export_text, font_name, font_size)

    # colonna destra larga quanto serve (nota + un piccolo margine)
    right_w = max(MAX_LOGO_W, text_w + 4*mm)

    # logo
    if logo_path and logo_path.exists():
        img = Image(str(logo_path))
        img._restrictSize(MAX_LOGO_W, MAX_LOGO_H)
    else:
        img = Spacer(MAX_LOGO_W, MAX_LOGO_H)

    # nota (Paragraph) allineata a destra
    small_para = Paragraph(escape(export_text), small_style)

    # tabella della colonna destra: logo centrato, nota a destra
    right_col = Table([[img],
                       [small_para]],
                      colWidths=[right_w])
    right_col.setStyle(TableStyle([
        ("ALIGN",  (0,0), (0,0), "CENTER"),  # logo centrato
        ("ALIGN",  (0,1), (0,1), "RIGHT"),   # nota a destra
        ("TOPPADDING",   (0,1), (0,1), 6),   # <-- più staccata dal logo (6pt)
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("BOTTOMPADDING",(0,0), (-1,-1), 0),
    ]))

    # riga header: titolo a sinistra, colonna destra (logo+nota) a destra
    tbl = Table([[title_para, right_col]],
                colWidths=[page_w - right_w, right_w])
    tbl.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING",   (0,0), (-1,-1), 0),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
    ]))
    return tbl


# ======================= Shaping tabella =======================
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

def _table_data_for(df: pd.DataFrame, para_style: ParagraphStyle):
    """
    Applica DISPLAY_ORDER, rimuove 'Residenza' e converte la colonna Note in Paragraph
    per il word-wrapping. Sostituisce '*' con <br/> (a capo forzato).
    """
    cols = [c for c in DISPLAY_ORDER if c in df.columns]
    dfp = df.copy()
    if cols:
        dfp = dfp[cols]
    dfp = dfp.drop(columns=["Residenza"], errors="ignore")

    header = list(dfp.columns)
    rows = dfp.fillna("").values.tolist()

    if "Indennità e note" in header:
        idx_note = header.index("Indennità e note")
        for r in rows:
            txt = str(r[idx_note])
            txt = escape(txt).replace("*", "<br/>")  # escape prima, poi <br/> per '*'
            r[idx_note] = Paragraph(txt, para_style)

    return [header] + rows

def _calc_col_widths(page_width: float) -> list[float]:
    """
    Calcola larghezze colonne (in punti) usando tutto lo spazio disponibile.
    Ordine: Matricola, Cognome e Nome, Turno, Inizio, Fine, Indennità e note
    """
    w_matricola = 24 * mm
    w_nome      = 60 * mm
    w_turno     = 22 * mm
    w_inizio    = 18 * mm
    w_fine      = 18 * mm
    used = w_matricola + w_nome + w_turno + w_inizio + w_fine
    w_note = max(30 * mm, page_width - used)  # il resto alla colonna Note (min 30mm)
    return [w_matricola, w_nome, w_turno, w_inizio, w_fine, w_note]

# ======================= Build PDF =======================
def build_pdf(path_out: Path, df: pd.DataFrame, meta: dict,
              logo_path: Path | None = None, title: str = "Servizio Giornaliero",
              inner_sort: str = "nome", exported_at: datetime | None = None):
    """
    - Titolo a sinistra + logo a destra con nota in piccolo sotto il logo
    - Larghezza piena, Note che assorbe lo spazio residuo, wrapping con '*' -> newline.
    - Grassetto su Turno/Inizio/Fine per righe in trasferta (con eccezioni).
    """
    # Margini/doc
    right = 10 * mm
    left  = 10 * mm
    top   = 12 * mm
    bottom= 12 * mm

    doc = SimpleDocTemplate(str(path_out), pagesize=A4,
                            rightMargin=right, leftMargin=left,
                            topMargin=top, bottomMargin=bottom)
    page_w = A4[0] - left - right  # larghezza utile

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleTight",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=20,
        textColor=colors.red,
        spaceAfter=6,
        spaceBefore=0,
    )
    group_style = ParagraphStyle(
        "GroupTitle",
        parent=styles["Heading2"],
        fontSize=14,
        leading=16,
        spaceBefore=6,
        spaceAfter=2,
        alignment=TA_CENTER,
    )
    note_style = ParagraphStyle(
        "NoteBody",
        parent=styles["BodyText"],
        leading=12,
        wordWrap="CJK",  # wrapping aggressivo per parole lunghe
    )
    small_note_style = ParagraphStyle(
        "SmallExportNote",
        parent=styles["Normal"],
        fontSize=8,
        textColor=colors.grey,
        spaceBefore=1,
        spaceAfter=0,
        alignment=TA_RIGHT,
    )

    elems: list = []

    # Testo "esportato il ... alle ..." (Europe/Rome)
    if exported_at is None:
        exported_at = datetime.now(_TZ_ROMA)
    export_text = exported_at.strftime("servizio esportato il %d/%m/%Y alle %H:%M")

    # Header: titolo + (logo sopra, export_text sotto)
    header_text = f"Servizio Giornaliero: {meta.get('giorno','')} {meta.get('data','')}"
    title_para  = Paragraph(header_text, title_style)
    elems.append(_header_table(title_para, logo_path, page_w, export_text, small_note_style))

    # Raggruppamento per Residenza
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
                collapse_cols=("Cognome e Nome", "Matricola"),
            )

        trasferte = _trasferta_mask(gdf)

        if not first_block:
            elems.append(Spacer(1, 3 * mm))
        first_block = False

        # Titolo gruppo (centrato)
        elems.append(Paragraph(res_name, group_style))

        # Tabella dati
        data = _table_data_for(gdf, note_style)
        header = data[0]

        # Indici colonne interessate (dinamici)
        col_idx = {name: header.index(name) for name in ["Turno", "Inizio", "Fine"] if name in header}
        idx_inizio = col_idx.get("Inizio")
        idx_fine = col_idx.get("Fine")

        col_widths = _calc_col_widths(page_w)
        tbl = Table(data, repeatRows=1, colWidths=col_widths)

        base_style = [
            ("GRID",       (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN",     (0, 0), (-1, -1), "TOP"),
        ]
        # Allineamento centro per Inizio/Fine
        if idx_inizio is not None:
            base_style.append(("ALIGN", (idx_inizio, 1), (idx_inizio, -1), "CENTER"))
        if idx_fine is not None:
            base_style.append(("ALIGN", (idx_fine, 1), (idx_fine, -1), "CENTER"))

        # Grassetto per trasferte (solo Turno/Inizio/Fine)
        for i, is_tr in enumerate(trasferte.tolist(), start=1):  # +1 per saltare header
            if not is_tr:
                continue
            for cidx in col_idx.values():
                base_style.append(("FONTNAME", (cidx, i), (cidx, i), "Helvetica-Bold"))

        tbl.setStyle(TableStyle(base_style))
        elems.append(tbl)

    # Build (niente canvas personalizzati, nessun foglio extra)
    doc.build(elems)
