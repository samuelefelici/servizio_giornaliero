# src/pdf_export.py
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, Flowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.pdfbase.pdfmetrics import stringWidth
from xml.sax.saxutils import escape
from datetime import datetime, timezone
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

GLOBAL_EXC_PATTERNS = (r"^IAST$", r"^N$",)

EXC_ANCONA_PREFIXES = (
    "D1R1", "D1R2", "D1R5", "D2R1", "D2R2", "D2R3", "D2R6",
    "NP", "ASC", "V5", "LU", "MA", "ME", "GI", "VE", "SA", "DO",
)
EXC_OTHER_PREFIXES = ("IAST", "N")

MAX_LOGO_W = 60 * mm
MAX_LOGO_H = 22.5 * mm

ARROW_MARK = "↳"  # segnale usato nella colonna nome per “secondo turno”
BLUE = colors.HexColor("#0b5ed7")  # colore testo per righe aggiunte (come anteprima)


# ---------- Freccia vettoriale (no font) ----------
class CornerArrow(Flowable):
    """
    Disegna un’icona “angolo + freccia verso destra” (tipo └→),
    ancorata al margine destro della cella.
    """

    def __init__(
        self,
        cell_width: float,
        size: float = 3.5 * mm,
        stroke: float = 1.0,
        shift_left_mm: float = 0.0,
        stroke_color=colors.black,
    ):
        super().__init__()
        self.width = float(cell_width)          # la Table userà questa width
        self.height = float(size * 1.6)         # altezza sufficiente per la punta
        self._s = float(size)
        self._stroke = float(stroke)
        self._shift = float(shift_left_mm)      # offset verso sinistra
        self._color = stroke_color

    def draw(self):
        c = self.canv
        s = self._s
        # ancoraggio vicino al bordo destro della cella
        x0 = self.width - (s * 4.2 + self._shift)
        y0 = self.height * 0.55

        c.saveState()
        c.setLineWidth(self._stroke)
        c.setStrokeColor(self._color)

        # segmento verticale
        c.line(x0, y0 - s, x0, y0)
        # segmento orizzontale
        x1 = x0 + s * 3.2
        c.line(x0, y0 - s, x1, y0 - s)
        # punta freccia
        c.line(x1, y0 - s, x1 - s * 0.9, y0 - s + s * 0.55)
        c.line(x1, y0 - s, x1 - s * 0.9, y0 - s - s * 0.55)

        c.restoreState()


# ======================= Utility orario =======================
def _as_rome(dt: datetime | None) -> datetime:
    """Rende dt in Europe/Rome. Se dt è naive, lo assume UTC."""
    if dt is None:
        return datetime.now(_TZ_ROMA)
    if dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None:
        return dt.replace(tzinfo=timezone.utc).astimezone(_TZ_ROMA)
    return dt.astimezone(_TZ_ROMA)


# ======================= Utility evidenziazione =======================
def _res_to_prefix(res: str) -> str | None:
    if not isinstance(res, str):
        return None
    r = res.upper().replace("_", " ").strip()
    if "JESI URBANO" in r or r == "JU":
        return "JU"
    if "JESI" in r or r == "J":
        return "J"
    if "MARINA" in r or r == "M":
        return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r == "C":
        return "C"
    if "OSIMO" in r or r == "O":
        return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r == "F":
        return "F"
    if "POLVERIGI" in r or r == "P":
        return "P"
    if "OSTRA" in r or r == "D":
        return "D"
    if "BELVED" in r or r == "B" or "DEPBELVE" in r:
        return "B"
    if "ANCONA" in r or r == "A":
        return "A"
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
    if _match_any(t, GLOBAL_EXC_PATTERNS):
        return False

    rp = _res_to_prefix(residenza)
    if rp == "A":
        if _starts_with_any(t, EXC_ANCONA_PREFIXES):
            return False
    else:
        if _starts_with_any(t, EXC_OTHER_PREFIXES):
            return False

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
    return df.apply(
        lambda r: _should_highlight_turno(r.get("Residenza"), r.get("Turno")),
        axis=1,
    )


# ======================= Header helper =======================
def _header_table(
    title_para: Paragraph,
    logo_path: Path | None,
    page_w: float,
    export_text: str,
    small_style: ParagraphStyle,
) -> Table:
    font_name = getattr(small_style, "fontName", "Helvetica")
    font_size = getattr(small_style, "fontSize", 8)
    text_w = stringWidth(export_text, font_name, font_size)

    right_w = max(MAX_LOGO_W, text_w + 4 * mm)

    if logo_path and logo_path.exists():
        img = Image(str(logo_path))
        img._restrictSize(MAX_LOGO_W, MAX_LOGO_H)
    else:
        img = Spacer(MAX_LOGO_W, MAX_LOGO_H)

    small_para = Paragraph(escape(export_text), small_style)

    right_col = Table([[img], [small_para]], colWidths=[right_w])
    right_col.setStyle(
        TableStyle([
            ("ALIGN", (0, 0), (0, 0), "CENTER"),
            ("ALIGN", (0, 1), (0, 1), "RIGHT"),
            ("TOPPADDING", (0, 1), (0, 1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ])
    )

    tbl = Table([[title_para, right_col]], colWidths=[page_w - right_w, right_w])
    tbl.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ])
    )
    return tbl


# ======================= Shaping tabella =======================
def _collapse_repeats(
    gdf: pd.DataFrame,
    key_cols=("Cognome e Nome", "Matricola"),
    collapse_cols=("Cognome e Nome", "Matricola"),
) -> pd.DataFrame:
    """
    Sui record consecutivi della stessa persona azzera i campi ripetitivi,
    ma sotto il Nome mostra il simbolo ↳ per indicare un secondo turno.
    """
    missing = [c for c in key_cols if c not in gdf.columns]
    if missing:
        return gdf
    g = gdf.copy()
    dup_mask = (g[list(key_cols)] == g[list(key_cols)].shift(1)).all(axis=1)
    if "Matricola" in g.columns:
        g.loc[dup_mask, "Matricola"] = ""
    if "Cognome e Nome" in g.columns:
        g.loc[dup_mask, "Cognome e Nome"] = ARROW_MARK
    return g


def _table_data_for(
    df: pd.DataFrame,
    para_style: ParagraphStyle,
    para_style_blue: ParagraphStyle | None = None,
    rows_blue: set[int] | None = None,
):
    """
    Converte df in dati per Table:
    - applica DISPLAY_ORDER, rimuove 'Residenza'
    - trasforma la colonna Note in Paragraph (wrapping, '*' -> <br/>)
    - se la riga è in rows_blue e para_style_blue è fornito, la colonna Note usa lo stile blu
    """
    rows_blue = rows_blue or set()
    cols = [c for c in DISPLAY_ORDER if c in df.columns]
    dfp = df.copy()
    if cols:
        dfp = dfp[cols]
    dfp = dfp.drop(columns=["Residenza"], errors="ignore")

    header = list(dfp.columns)
    rows = dfp.fillna("").values.tolist()

    idx_note = header.index("Indennità e note") if "Indennità e note" in header else None
    table_rows = [header]
    for ridx, r in enumerate(rows, start=1):  # 1-based (per allinearsi agli indici della Table)
        if idx_note is not None:
            txt = str(r[idx_note])
            txt = escape(txt).replace("*", "<br/>")
            # Stile blu sulle righe segnate
            style_to_use = para_style_blue if (para_style_blue and ridx in rows_blue) else para_style
            r[idx_note] = Paragraph(txt, style_to_use)
        table_rows.append(r)

    return table_rows


def _calc_col_widths(page_width: float) -> list[float]:
    w_matricola = 24 * mm
    w_nome = 60 * mm
    w_turno = 22 * mm
    w_inizio = 18 * mm
    w_fine = 18 * mm
    used = w_matricola + w_nome + w_turno + w_inizio + w_fine
    w_note = max(30 * mm, page_width - used)
    return [w_matricola, w_nome, w_turno, w_inizio, w_fine, w_note]


# ======================= Build PDF =======================
def build_pdf(
    path_out: Path,
    df: pd.DataFrame,
    meta: dict,
    logo_path: Path | None = None,
    title: str = "Servizio Giornaliero",
    inner_sort: str = "nome",
    exported_at: datetime | None = None,
):
    # ================== Auto-nome file: SG_ggmmaaaa.pdf ==================
    oggi = datetime.now(_TZ_ROMA)
    file_name = f"SG_{oggi.strftime('%d%m%Y')}.pdf"
    # Se path_out è una cartella, crea il file dentro di essa
    if path_out.is_dir():
        path_out = path_out / file_name
    else:
        # Se path_out è un file, sostituisci il nome
        path_out = path_out.with_name(file_name)

    """
    - Titolo a sinistra + logo a destra con nota in piccolo sotto il logo
    - Wrapping note con '*' -> newline forzato
    - Grassetto su Turno/Inizio/Fine per righe in trasferta (con eccezioni)
    - Righe aggiunte (_added=True): testo blu su tutta la riga (incluse note e freccia)
    """
    # Margini/doc
    right = 10 * mm
    left = 10 * mm
    top = 12 * mm
    bottom = 12 * mm

    doc = SimpleDocTemplate(
        str(path_out),
        pagesize=A4,
        rightMargin=right,
        leftMargin=left,
        topMargin=top,
        bottomMargin=bottom,
    )
    page_w = A4[0] - left - right

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
        wordWrap="CJK",
    )
    # Stile note blu (per righe aggiunte)
    note_style_blue = ParagraphStyle(
        "NoteBodyBlue",
        parent=note_style,
        textColor=BLUE,
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

    # Orario Europe/Rome
    exported_at = _as_rome(exported_at)
    export_text = exported_at.strftime("servizio esportato il %d/%m/%Y alle %H:%M")

    header_text = f"Servizio Giornaliero: {meta.get('giorno', '')} {meta.get('data', '')}"
    title_para = Paragraph(header_text, title_style)
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
        added_mask = (
            gdf["_added"].fillna(False)
            if "_added" in gdf.columns
            else pd.Series(False, index=gdf.index)
        )

        if not first_block:
            elems.append(Spacer(1, 3 * mm))
        first_block = False

        elems.append(Paragraph(res_name, group_style))

        # Indici delle righe aggiunte (1-based, esclusa header)
        added_rows = {i for i, v in enumerate(added_mask.tolist(), start=1) if v}

        # Tabella dati (nota: per le righe blu, la colonna Note usa lo stile blu)
        data = _table_data_for(gdf, note_style, para_style_blue=note_style_blue, rows_blue=added_rows)
        header = data[0]

        col_idx = {name: header.index(name) for name in ["Turno", "Inizio", "Fine"] if name in header}
        idx_inizio = col_idx.get("Inizio")
        idx_fine = col_idx.get("Fine")
        idx_nome = header.index("Cognome e Nome") if "Cognome e Nome" in header else None

        col_widths = _calc_col_widths(page_w)

        # Sostituisco i segnaposto ARROW_MARK con frecce vettoriali (blu se riga aggiunta)
        if idx_nome is not None:
            for ridx, row in enumerate(data[1:], start=1):
                if str(row[idx_nome]).strip() == ARROW_MARK:
                    stroke_col = BLUE if ridx in added_rows else colors.black
                    row[idx_nome] = CornerArrow(
                        cell_width=col_widths[idx_nome],
                        size=1.5 * mm,
                        stroke=0.8,
                        shift_left_mm=2 * mm,
                        stroke_color=stroke_col,
                    )

        tbl = Table(data, repeatRows=1, colWidths=col_widths)

        base_style = [
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
        if idx_inizio is not None:
            base_style.append(("ALIGN", (idx_inizio, 1), (idx_inizio, -1), "CENTER"))
        if idx_fine is not None:
            base_style.append(("ALIGN", (idx_fine, 1), (idx_fine, -1), "CENTER"))

        # Grassetto per trasferte (solo Nome/Turno/Inizio/Fine)
        for i, is_tr in enumerate(trasferte.tolist(), start=1):
            if is_tr:
                for cidx in col_idx.values():
                    base_style.append(("FONTNAME", (cidx, i), (cidx, i), "Helvetica-Bold"))
                    base_style.append(("TEXTCOLOR", (cidx, i), (cidx, i), colors.HexColor("#0b5ed7")))
                    # Colonna precedente (Cognome e Nome)
                    if cidx > 0:
                        base_style.append(("FONTNAME", (cidx-1, i), (cidx-1, i), "Helvetica-Bold"))
                        base_style.append(("TEXTCOLOR", (cidx-1, i), (cidx-1, i), colors.HexColor("#0b5ed7")))

        # Righe aggiunte: testo blu sull'intera riga (plus grassetto già presente)
        for i in added_rows:
            base_style.append(("TEXTCOLOR", (0, i), (-1, i), BLUE))
            base_style.append(("FONTNAME", (0, i), (-1, i), "Helvetica-Bold"))

        tbl.setStyle(TableStyle(base_style))
        elems.append(tbl)

    doc.build(elems)
    return path_out
