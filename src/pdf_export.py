# src/pdf_export.py
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfgen import canvas as rl_canvas
from xml.sax.saxutils import escape
from datetime import datetime
import pandas as pd
from pathlib import Path
import re

from datetime import datetime
try:
    from zoneinfo import ZoneInfo         # Python 3.9+
    _TZ_ROMA = ZoneInfo("Europe/Rome")
except Exception:
    import pytz                           # fallback
    _TZ_ROMA = pytz.timezone("Europe/Rome")

from .constants import DISPLAY_ORDER

REST_CODES = {"R", "RR"}

# Eccezioni GLOBALI: valgono per tutti i depositi
GLOBAL_EXC_PATTERNS = (
    r"^IAST$",   # esattamente IAST
    r"^N$",      # esattamente N
)

EXC_ANCONA_PREFIXES = (
    "D1R1","D1R2","D1R5","D2R1","D2R2","D2R3","D2R6",
    "NP","ASC","V5",
    "LU","MA","ME","GI","VE","SA","DO",
)
EXC_OTHER_PREFIXES = ("IAST","N")

# -------- utility evidenziazione (allineate ad app.py) --------
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
    t = _norm_turno(turno)
    if not t or t in {"ASSENTE", *REST_CODES}:
        return False
    # eccezioni globali
    if _match_any(t, GLOBAL_EXC_PATTERNS):
        return False

    rp = _res_to_prefix(residenza)

    # eccezioni per deposito
    if rp == "A":  # ANCONA
        if _starts_with_any(t, EXC_ANCONA_PREFIXES):
            return False
    else:
        if _starts_with_any(t, EXC_OTHER_PREFIXES):
            return False

    # regola standard J/JU equivalenti
    if t.startswith("JU"):
        b = "JU"
    elif t.startswith("J"):
        b = "J"
    else:
        b = t[0] if t else None

    if not rp or not b:
        return False
    return b not in _accepted_prefixes_for_res(rp)

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)
    return df.apply(lambda r: _should_highlight_turno(r.get("Residenza"), r.get("Turno")), axis=1)

# -------- header & footer helpers --------
def _header_table(title_para: Paragraph, logo_path: Path | None, page_w: float) -> Table:
    """Riga con titolo (sx) e logo (dx) che occupa tutta la larghezza."""
    max_logo_w, max_logo_h = 60*mm, 22.5*mm
    if logo_path and logo_path.exists():
        img = Image(str(logo_path))
        img._restrictSize(max_logo_w, max_logo_h)
    else:
        img = Spacer(max_logo_w, max_logo_h)

    tbl = Table([[title_para, img]],
                colWidths=[page_w - max_logo_w, max_logo_w])
    tbl.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",  (1,0), (1,0),  "RIGHT"),
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING",   (0,0), (-1,-1), 0),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
    ]))
    return tbl

class _LastPageFooterCanvas(rl_canvas.Canvas):
    """Disegna un testo in piccolo in basso a destra SOLO sull'ultima pagina."""
    def __init__(self, *args, exported_text: str = "", **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        self._exported_text = exported_text

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        # salva anche lo stato della pagina corrente
        self._saved_page_states.append(dict(self.__dict__))
        for i, state in enumerate(self._saved_page_states):
            self.__dict__.update(state)
            if i == len(self._saved_page_states) - 1:
                self._draw_footer()
            super().showPage()
        super().save()

    def _draw_footer(canvas, doc, footer_text: str):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        # posizionamento in basso a destra
        w = canvas.stringWidth(footer_text, "Helvetica", 8)
        x = doc.pagesize[0] - doc.rightMargin - w
        y = doc.bottomMargin * 0.6
        canvas.drawString(x, y, footer_text)
        canvas.restoreState()

# -------- shaping tabella --------
def _collapse_repeats(gdf: pd.DataFrame,
                      key_cols=("Cognome e Nome", "Matricola"),
                      collapse_cols=("Cognome e Nome", "Matricola")) -> pd.DataFrame:
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
    cols = [c for c in DISPLAY_ORDER if c in df.columns]
    dfp = df.copy()
    if cols:
        dfp = dfp[cols]
    dfp = dfp.drop(columns=["Residenza"], errors="ignore")

    header = list(dfp.columns)
    rows = dfp.fillna("").values.tolist()

    # Note con a capo automatico + a capo forzato su "*"
    if "Indennità e note" in header:
        idx_note = header.index("Indennità e note")
        for r in rows:
            txt = str(r[idx_note])
            txt = escape(txt).replace("*", "<br/>")
            r[idx_note] = Paragraph(txt, para_style)

    return [header] + rows

def _calc_col_widths(page_width_mm: float) -> list:
    w_matricola = 24 * mm
    w_nome      = 60 * mm
    w_turno     = 22 * mm
    w_inizio    = 18 * mm
    w_fine      = 18 * mm
    used = w_matricola + w_nome + w_turno + w_inizio + w_fine
    w_note = max(30 * mm, page_width_mm - used)
    return [w_matricola, w_nome, w_turno, w_inizio, w_fine, w_note]

# -------- build PDF --------
def build_pdf(path_out: Path, df: pd.DataFrame, meta: dict,
              logo_path: Path | None = None, title: str = "Servizio Giornaliero",
              inner_sort: str = "nome", exported_at: datetime | None = None):
    """
    - Titolo a sinistra + logo a destra sulla stessa riga (header del documento)
    - Footer in ultima pagina: 'servizio esportato il DATA alle ORA.'
    - Larghezza piena, Note che assorbe lo spazio residuo, wrapping con '*' -> newline.
    """
    right=10*mm; left=10*mm; top=12*mm; bottom=12*mm
    doc = SimpleDocTemplate(str(path_out), pagesize=A4,
                            rightMargin=right, leftMargin=left,
                            topMargin=top, bottomMargin=bottom)
    page_w = A4[0] - left - right

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleTight", parent=styles["Title"],
        fontName="Helvetica-Bold", fontSize=18, leading=20,
        textColor=colors.red, spaceAfter=6, spaceBefore=0
    )
    group_style = ParagraphStyle(
        "GroupTitle", parent=styles["Heading2"],
        fontSize=14, leading=16, spaceBefore=6, spaceAfter=2, alignment=TA_CENTER
    )
    note_style  = ParagraphStyle("NoteBody", parent=styles["BodyText"],
                                 leading=12, wordWrap="CJK")

    elems = []

    # Header: titolo + logo sulla stessa riga
    header_text = f"Servizio Giornaliero: {meta.get('giorno','')} {meta.get('data','')}"
    title_para  = Paragraph(header_text, title_style)
    elems.append(_header_table(title_para, logo_path, page_w))

    # Raggruppamento per Residenza
    if "Residenza" not in df.columns:
        blocks = [("TUTTI", df)]
    else:
        df = df.sort_values(by=["Residenza"]).reset_index(drop=True)
        blocks = [(str(res), g.copy()) for res, g in df.groupby("Residenza", sort=True)]

    first_block = True
    for res_name, gdf in blocks:
        if inner_sort == "inizio" and "Inizio" in gdf.columns:
            by = ["Inizio"] + (["Cognome e Nome"] if "Cognome e Nome" in gdf.columns else [])
            gdf = gdf.sort_values(by=by).reset_index(drop=True)
        else:
            by = []
            if "Cognome e Nome" in gdf.columns: by.append("Cognome e Nome")
            if "Inizio" in gdf.columns:         by.append("Inizio")
            if by:
                gdf = gdf.sort_values(by=by).reset_index(drop=True)
            gdf = _collapse_repeats(
                gdf, key_cols=("Cognome e Nome", "Matricola"),
                collapse_cols=("Cognome e Nome", "Matricola")
            )

        trasferte = _trasferta_mask(gdf)

        if not first_block:
            elems.append(Spacer(1, 3*mm))
        first_block = False

        elems.append(Paragraph(res_name, group_style))

        data   = _table_data_for(gdf, note_style)
        header = data[0]
        col_idx   = {name: header.index(name) for name in ["Turno","Inizio","Fine"] if name in header}
        idx_inizio = col_idx.get("Inizio")
        idx_fine   = col_idx.get("Fine")

        col_widths = _calc_col_widths(page_w)
        tbl = Table(data, repeatRows=1, colWidths=col_widths)

        base_style = [
            ("GRID",       (0,0), (-1,-1), 0.25, colors.grey),
            ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
            ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
            ("VALIGN",     (0,0), (-1,-1), "TOP"),
        ]
        if idx_inizio is not None:
            base_style.append(("ALIGN", (idx_inizio,1), (idx_inizio,-1), "CENTER"))
        if idx_fine is not None:
            base_style.append(("ALIGN", (idx_fine,1), (idx_fine,-1), "CENTER"))

        for i, is_tr in enumerate(trasferte.tolist(), start=1):
            if not is_tr:
                continue
            for cidx in col_idx.values():
                base_style.append(("FONTNAME", (cidx, i), (cidx, i), "Helvetica-Bold"))

        tbl.setStyle(TableStyle(base_style))
        elems.append(tbl)

    # Footer (solo ultima pagina)
    if exported_at is None:
        exported_at = datetime.now()
    footer_text = exported_at.strftime("servizio esportato il %d/%m/%Y alle %H:%M")

    def _canvas_factory(*args, **kwargs):
        return _LastPageFooterCanvas(*args, exported_text=footer_text, **kwargs)

    doc.build(elems, canvasmaker=_canvas_factory)
