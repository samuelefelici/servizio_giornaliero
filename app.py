import streamlit as st
import pandas as pd
import io, tempfile
from pathlib import Path
from src.process import read_input_excel, transform_dataframe
from src.pdf_export import build_pdf
from src.constants import TITLE

st.set_page_config(page_title="Servizio Giornaliero", layout="wide")

st.title("📋 Servizio Giornaliero – ExtraUrbano (Python)")
st.caption("Drag & drop del file Excel (.xls/.xlsx), pulizia automatica ed export in PDF.")

cfg_dir = Path("config")
assets_dir = Path("assets")
logo_path = assets_dir / "logo.jpg"

uploaded = st.file_uploader("Trascina qui il file oppure selezionalo", type=["xls","xlsx"])

col1, col2 = st.columns([1,1])
with col1:
    sort_by_res = st.checkbox("Ordina per Residenza", value=True)
    sort_by_cat = st.checkbox("Ordina per Categoria", value=True)
    sort_by_turno = st.checkbox("Ordina per Turno", value=True)
    sort_by_inizio = st.checkbox("Ordina per Inizio", value=True)

if st.button("▶️ Elabora"):
    if not uploaded:
        st.warning("Carica prima un file.")
        st.stop()
    try:
        df, meta = read_input_excel(uploaded)
        # aggiorna criteri di ordinamento runtime
        import src.constants as C
        sel = []
        if sort_by_res and "Residenza" in df.columns: sel.append("Residenza")
        if sort_by_cat and "Categoria" in df.columns: sel.append("Categoria")
        if sort_by_turno and "Turno" in df.columns: sel.append("Turno")
        if sort_by_inizio and "Inizio" in df.columns: sel.append("Inizio")
        if sel: C.DEFAULT_SORT = sel

        df_out = transform_dataframe(df, cfg_dir)

        st.success(f"File elaborato. Data: {meta.get('data','?')} – {meta.get('giorno','?')} (fonte: {meta.get('origine','?')})")
        st.subheader("Anteprima dati")
        st.dataframe(df_out, use_container_width=True, hide_index=True)

        # Download Excel (solo dati, niente riepilogo)
        xls_buf = io.BytesIO()
        with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
            df_out.to_excel(writer, sheet_name="ServizioGiornaliero", index=False)
        st.download_button("⬇️ Scarica Excel", data=xls_buf.getvalue(), file_name="ServizioGiornaliero.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF
        with tempfile.TemporaryDirectory() as td:
            pdf_path = Path(td) / "ServizioGiornaliero.pdf"
            build_pdf(pdf_path, df_out, meta, logo_path if logo_path.exists() else None, title=TITLE)
            st.download_button("⬇️ Scarica PDF", data=pdf_path.read_bytes(), file_name="ServizioGiornaliero.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"Errore durante l'elaborazione: {e}")
