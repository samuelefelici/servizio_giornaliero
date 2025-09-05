from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
import pandas as pd
from pathlib import Path
from .constants import DISPLAY_ORDER
import re
from .constants import DISPLAY_ORDER


# colonne da stampare (senza Residenza)
_PRINT_COLS = ["Cognome e Nome", "Matricola", "Turno", "Inizio", "Fine", "Indennità e note"]

def _res_to_prefix(res: str) -> str | None:
    if not isinstance(res, str): return None
    r = res.upper().replace("_", " ")
    if "JESI URBANO" in r or r.strip() == "JU": return "JU"
    if "JESI" in r or r.strip() == "J": return "J"
    if "MARINA" in r or r.strip() == "M": return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r.strip() == "C": return "C"
    if "OSIMO" in r or r.strip() == "O": return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r.strip() == "F": return "F"
    if "POLVERIGI" in r or r.strip() == "P": return "P"
    if "OSTRA" in r or r.strip() == "D": return "D"
    if "BELVED" in r or r.strip() == "B" or "DEPBELVE" in r: return "B"
    if "ANCONA" in r or r.strip() == "A": return "A"
    return None

def _turno_prefix(turno: str) -> str | None:
    if not isinstance(turno, str): turno = str(turno)
    m = re.match(r"[A-Za-z]+", turno.strip())
    return m.group(0).upper() if m else None

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)
    expected = df["Residenza"].apply(_res_to_prefix)
    actual = df["Turno"].astype(str).apply(_turno_prefix)
    return expected.notna() & actual.notna() & (expected != actual)


def _table_data_for(df: pd.DataFrame):
    # Applica l’ordine standard e ignora "Residenza" se presente
    cols = [c for c in DISPLAY_ORDER if c in df.columns]
    dfp = df.copy()
    if cols:
        dfp = dfp[cols]
    dfp = dfp.drop(columns=["Residenza"], errors="ignore")

    # NIENTE rename -> teniamo "Indennità e note"
    header = list(dfp.columns)
    rows = dfp.fillna("").values.tolist()
    return [header] + rows


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

def build_pdf(path_out: Path, df: pd.DataFrame, meta: dict,
              logo_path: Path|None=None, title: str="Servizio Giornaliero",
              inner_sort: str = "nome"):
    """
    Crea un PDF raggruppato per Residenza.
    - Titolo pagina senza spazio dopo
    - Per ogni Residenza: titolo grande in grassetto e tabella senza colonna Residenza
    - inner_sort: "nome" (Cognome e Nome A→Z + Inizio crescente per la stessa persona)
                  oppure "inizio" (orario, con Nome come tie-breaker)
    """
    doc = SimpleDocTemplate(str(path_out), pagesize=A4,
                            rightMargin=10*mm, leftMargin=10*mm,
                            topMargin=12*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleTight", parent=styles["Title"], spaceAfter=0, spaceBefore=0)
    group_style = ParagraphStyle("GroupTitle", parent=styles["Heading2"],
                                 fontSize=14, leading=16, spaceBefore=6, spaceAfter=2)

    elems = []

    # intestazione pagina (senza riga vuota sotto)
    header_text = f"{title} — {meta.get('data','')} — {meta.get('giorno','')}"
    elems.append(Paragraph(header_text, title_style))

    if logo_path and logo_path.exists():
        elems.append(Spacer(1, 2*mm))
        img = Image(str(logo_path)); img._restrictSize(40*mm, 15*mm)
        elems.append(img)

    # grouping per Residenza (A→Z)
    if "Residenza" not in df.columns:
        blocks = [("TUTTI", df)]
    else:
        df = df.sort_values(by=["Residenza"]).reset_index(drop=True)
        blocks = [(str(res), g.copy()) for res, g in df.groupby("Residenza", sort=True)]

    first_block = True
    for res_name, gdf in blocks:
        # sort interno per scelta
        if inner_sort == "inizio" and "Inizio" in gdf.columns:
            by = ["Inizio"] + (["Cognome e Nome"] if "Cognome e Nome" in gdf.columns else [])
            gdf = gdf.sort_values(by=by).reset_index(drop=True)
        else:
            # "nome": ordina per Nome e, per la stessa persona, per Inizio crescente
            by = []
            if "Cognome e Nome" in gdf.columns:
                by.append("Cognome e Nome")
            if "Inizio" in gdf.columns:
                by.append("Inizio")
            if by:
                gdf = gdf.sort_values(by=by).reset_index(drop=True)
            # compattamento ripetizioni su righe consecutive della stessa persona
            gdf = _collapse_repeats(gdf, key_cols=("Cognome e Nome","Matricola"),
                                    collapse_cols=("Cognome e Nome","Matricola"))

        if not first_block:
            elems.append(Spacer(1, 3*mm))
        first_block = False

        # titolo gruppo
        elems.append(Paragraph(res_name, group_style))

        # tabella (senza Residenza)
        data = _table_data_for(gdf)
        tbl = Table(data, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("ALIGN",(3,1),(4,-1),"CENTER"),  # Inizio/Fine
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))
        elems.append(tbl)

    doc.build(elems)
