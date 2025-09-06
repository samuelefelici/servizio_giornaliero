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
from pathlib import Path
import pandas as pd
import re

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
def _header_table(title_para: Paragraph, logo_path: Path | None, page_w: float) -> Table:
    if logo_path and logo_path.exists():
        img = Image(str(logo_path))
        img._restrictSize(MAX_LOGO_W, MAX_LOGO_H)
    else:
        img = Spacer(MAX_LOGO_W, MAX_LOGO_H)

    tbl = Table([[title_para, img]], colWidths=[page_w - MAX_LOGO_W, MAX_LOGO_W])
    tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",  (1, 0), (1, 0),  "RIGHT"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
    ]))
    return tbl

# ======================= Shaping tabella =======================
def _collapse_repeats(gdf: pd.DataFrame,
                      key_cols=("Cognome e Nome", "Matricola"),
                      collapse_cols=("Cognome e Nome", "Matricola")) -> pdDataFrame:
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

    if "Indennità e note" in header:
        idx_note = header.index("Indennità e note")
        for r in rows:
            txt = str(r[idx_note])
            txt = escape(txt).replace("*", "<br/>")
            r[idx_note] = Paragraph(txt, para_style)

    return [header] + rows

def _calc_col_widths(page_width: float) -> list[float]:
    w_matricola = 24 * mm
    w_nome      = 60 * mm
    w_turno     = 22 * mm
    w_inizio    = 18 * mm
    w_fine      = 18 * mm
    used = w_matricola + w_nome + w_turno + w_inizio + w_fine
    w_note = max(30 * mm, page_width - used)
    return [w_matricola, w_nome, w_turno, w_inizio, w_fine, w_note]

# ======================= Story builder =======================
def _make_story(df: pd.DataFrame, meta: dict, logo_path: Path | None,
                page_w: float, styles, inner_sort: str) -> list:
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
    note_style = ParagraphStyle("NoteBody", parent=styles["BodyText"], leading=12, wordWrap="CJK")

    elems: list = []
    header_text = f"Servizio Giornaliero: {meta.get('giorno','')} {meta.get('data','')}"
    title_para = Paragraph(header_text, title_style)
    elems.append(_header_table(title_para, logo_path, page_w))

    # Grouping per Residenza
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

        elems.append(Paragraph(res_name, group_style))

        data = _table_data_for(gdf, note_style)
        header = data[0]

        col_idx = {name: header.index(name) for name in ["Turno", "Inizio", "Fine"] if name in header}
        idx_inizio = col_idx.get("Inizio")
        idx_fine   = col_idx.get("Fine")

        col_widths = _calc_col_widths(page_w)
        tbl = Table(data, repeatRows=1, colWidths=col_widths)

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

        for i, is_tr in enumerate(trasferte.tolist(), start=1):
            if not is_tr:
                continue
            for cidx in col_idx.values():
                base_style.append(("FONTNAME", (cidx, i), (cidx, i), "Helvetica-Bold"))

        tbl.setStyle(TableStyle(base_style))
        elems.append(tbl)

    return elems

# ======================= Canvas (ultima pagina, nessun foglio extra) =======================
class _FooterCanvas(rl_canvas.Canvas):
    """
    Canvas che disegna il footer SOLO sull'ultima pagina, senza introdurre pagine extra.
    Nota: usiamo _startPage() in showPage() (NON super().showPage()) come da ricetta ReportLab.
    """
    def __init__(self, *args, footer_text: str = "", right_margin: float = 10*mm,
                 bottom_margin: float = 12*mm, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        self._footer_text = footer_text
        self._right_margin = right_margin
        self._bottom_margin = bottom_margin

    def showPage(self):
        # salva lo stato della pagina corrente e avvia la successiva SENZA scriverla subito
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()  # <- fondamentale per NON creare una pagina extra

    def save(self):
        # aggiungi anche l'ultima pagina
        self._saved_page_states.append(dict(self.__dict__))

        for i, state in enumerate(self._saved_page_states):
            self.__dict__.update(state)
            is_last = (i == len(self._saved_page_states) - 1)

            if is_last and self._footer_text:
                self.saveState()
                self.setFont("Helvetica", 8)
                w = self.stringWidth(self._footer_text, "Helvetica", 8)
                x = self._pagesize[0] - self._right_margin - w
                y = self._bottom_margin * 0.6
                self.drawString(x, y, self._footer_text)
                self.restoreState()

            rl_canvas.Canvas.showPage(self)  # scrivi la pagina
        rl_canvas.Canvas.save(self)

# ======================= Build PDF =======================
def build_pdf(path_out: Path, df: pd.DataFrame, meta: dict,
              logo_path: Path | None = None, title: str = "Servizio Giornaliero",
              inner_sort: str = "nome", exported_at: datetime | None = None):

    right = 10 * mm
    left  = 10 * mm
    top   = 12 * mm
    bottom= 12 * mm

    doc = SimpleDocTemplate(str(path_out), pagesize=A4,
                            rightMargin=right, leftMargin=left,
                            topMargin=top, bottomMargin=bottom)
    styles = getSampleStyleSheet()
    page_w = A4[0] - left - right

    story = _make_story(df, meta, logo_path, page_w, styles, inner_sort)

    # Footer (ultima pagina) — orario Europe/Rome
    if exported_at is None:
        exported_at = datetime.now(_TZ_ROMA)
    footer_text = exported_at.strftime("servizio esportato il %d/%m/%Y alle %H:%M")

    def _canvas_factory(*args, **kwargs):
        return _FooterCanvas(*args,
                             footer_text=footer_text,
                             right_margin=right,
                             bottom_margin=bottom,
                             **kwargs)

    doc.build(story, canvasmaker=_canvas_factory)
