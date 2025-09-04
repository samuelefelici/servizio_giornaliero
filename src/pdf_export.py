from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd
from pathlib import Path

def df_to_table_data(df: pd.DataFrame):
    # Seleziona colonne principali per la stampa
    keep = [c for c in ["Residenza","Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Stato"] if c in df.columns]
    dfp = df[keep].copy()
    dfp.columns = [c.replace("Indennità e note","Note") for c in dfp.columns]
    header = list(dfp.columns)
    rows = dfp.fillna("").values.tolist()
    return [header] + rows

def build_pdf(path_out: Path, df: pd.DataFrame, riepilogo: pd.DataFrame, meta: dict, logo_path: Path|None=None, title: str="Servizio Giornaliero"):
    doc = SimpleDocTemplate(str(path_out), pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=12*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    elems = []

    header_text = f"{title} — {meta.get('data','')} — {meta.get('giorno','')}"
    elems.append(Paragraph(header_text, styles["Title"]))
    elems.append(Spacer(1, 4*mm))

    if logo_path and logo_path.exists():
        img = Image(str(logo_path))
        img._restrictSize(40*mm, 15*mm)
        elems.append(img)
        elems.append(Spacer(1, 3*mm))

    # Riepiloghi
    elems.append(Paragraph("Riepiloghi (Assenze per Sigla)", styles["Heading3"]))
    if len(riepilogo):
        rt = [["Sigla","Conteggio"]] + riepilogo.values.tolist()
        rtbl = Table(rt, hAlign="LEFT")
        rtbl.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ]))
        elems.append(rtbl)
        elems.append(Spacer(1, 6*mm))
    else:
        elems.append(Paragraph("Nessuna assenza rilevata.", styles["Normal"]))
        elems.append(Spacer(1, 6*mm))

    # Tabella principale
    data = df_to_table_data(df)
    col_widths = None  # lascia auto
    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("ALIGN",(4,1),(5,-1),"CENTER"),   # Inizio/Fine centrate se presenti
        ("VALIGN",(0,0),(-1,-1),"TOP"),
    ]))
    elems.append(tbl)

    doc.build(elems)
