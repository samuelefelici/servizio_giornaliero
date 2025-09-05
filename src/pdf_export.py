from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
import pandas as pd
from pathlib import Path

# colonne da stampare (senza Residenza)
_PRINT_COLS = ["Cognome e Nome", "Matricola", "Turno", "Inizio", "Fine", "Indennità e note"]

def _table_data_for(df: pd.DataFrame):
    keep = [c for c in _PRINT_COLS if c in df.columns]
    dfp = df[keep].copy()
    # rinomina "Indennità e note" -> "Note"
    dfp.columns = [c.replace("Indennità e note", "Note") for c in dfp.columns]
    header = list(dfp.columns)
    rows = dfp.fillna("").values.tolist()
    return [header] + rows

def build_pdf(path_out: Path, df: pd.DataFrame, meta: dict,
              logo_path: Path|None=None, title: str="Servizio Giornaliero",
              inner_sort: str = "nome"):
    """
    Crea un PDF raggruppato per Residenza.
    - Titolo pagina senza spazio dopo (niente riga vuota tra intestazione e dati)
    - Per ogni Residenza: titolo grande in grassetto e tabella senza colonna Residenza
    - inner_sort: "nome" (Cognome e Nome A→Z) oppure "inizio" (orario)
    """
    doc = SimpleDocTemplate(str(path_out), pagesize=A4,
                            rightMargin=10*mm, leftMargin=10*mm,
                            topMargin=12*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleTight", parent=styles["Title"], spaceAfter=0, spaceBefore=0
    )
    group_style = ParagraphStyle(
        "GroupTitle", parent=styles["Heading2"],
        fontSize=14, leading=16, spaceBefore=6, spaceAfter=2
    )

    elems = []

    # intestazione pagina
    header_text = f"{title} — {meta.get('data','')} — {meta.get('giorno','')}"
    elems.append(Paragraph(header_text, title_style))
    # niente Spacer qui: zero riga vuota tra intestazione e primo blocco

    if logo_path and logo_path.exists():
        # posizioniamo il logo subito sotto il titolo ma con poco spazio
        elems.append(Spacer(1, 2*mm))
        img = Image(str(logo_path)); img._restrictSize(40*mm, 15*mm)
        elems.append(img)

    # grouping per Residenza (ordinamento automatico A→Z)
    if "Residenza" not in df.columns:
        # se manca per qualche motivo, stampiamo tutto in un solo blocco
        blocks = [("TUTTI", df)]
    else:
        # ordina per Residenza
        df = df.sort_values(by=["Residenza"]).reset_index(drop=True)
        blocks = []
        for res, g in df.groupby("Residenza", sort=True):
            blocks.append((str(res), g.copy()))

    # render di ciascun deposito
    first_block = True
    for res_name, gdf in blocks:
        # sort interno
        if inner_sort == "inizio" and "Inizio" in gdf.columns:
            gdf = gdf.sort_values(by=["Inizio", "Cognome e Nome"] if "Cognome e Nome" in gdf.columns else ["Inizio"]).reset_index(drop=True)
        else:
            # default: per nome
            if "Cognome e Nome" in gdf.columns:
                gdf = gdf.sort_values(by=["Cognome e Nome"]).reset_index(drop=True)

        # separazione fra blocchi (leggera, non una riga vuota gigantesca)
        if not first_block:
            elems.append(Spacer(1, 3*mm))
        first_block = False

        # titolo gruppo
        elems.append(Paragraph(res_name, group_style))

        # tabella dati del gruppo (senza Residenza)
        data = _table_data_for(gdf)
        tbl = Table(data, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("ALIGN",(3,1),(4,-1),"CENTER"),  # Inizio/Fine sono col 3 e 4 in questo layout
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))
        elems.append(tbl)

    doc.build(elems)
