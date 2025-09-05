import os, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd
import io, tempfile

try:
    from src.process import read_input_excel, transform_dataframe
    from src.pdf_export import build_pdf
    from src.constants import TITLE
except Exception as e:
    import traceback
    st.error(f"Errore durante l'import dei moduli: {e}")
    st.code("".join(traceback.format_exception(*sys.exc_info())))
    st.stop()

st.set_page_config(page_title="Servizio Giornaliero", layout="wide")
st.title("📋 Servizio Giornaliero – ExtraUrbano (Python)")
st.caption("Drag & drop del file Excel (.xls/.xlsx), pulizia automatica ed export in PDF (raggruppato per deposito).")

cfg_dir = Path("config")
assets_dir = Path("assets")
logo_path = assets_dir / "logo.jpg"

uploaded = st.file_uploader("Trascina qui il file oppure selezionalo", type=["xls","xlsx"])

col1, col2 = st.columns([1,1])
with col1:
    inner_sort_choice = st.radio(
        "Ordina dentro ciascun deposito per:",
        ["Cognome e Nome (A→Z)", "Inizio (orario)"],
        horizontal=True,
        index=0
    )

if st.button("▶️ Elabora"):
    if not uploaded:
        st.warning("Carica prima un file.")
        st.stop()

    try:
        df, meta = read_input_excel(uploaded)

        # pipeline base (filtri, rinomine, Assente, ecc.)
        df_out = transform_dataframe(df, cfg_dir)

        # anteprima a blocchi (senza colonna Residenza)
        st.success(f"File elaborato. Data: {meta.get('data','?')} – {meta.get('giorno','?')} (fonte: {meta.get('origine','?')})")
        st.subheader("Anteprima per deposito")

# ordine gruppi (depositi): A→Z
if "Residenza" in df_out.columns:
    res_list = sorted(df_out["Residenza"].dropna().astype(str).unique())
    for res in res_list:
        g = df_out[df_out["Residenza"].astype(str) == res].copy()

        # sort interno per scelta
        if inner_sort_choice.startswith("Inizio") and "Inizio" in g.columns:
            by = ["Inizio"]
            if "Cognome e Nome" in g.columns:
                by.append("Cognome e Nome")
            g = g.sort_values(by=by).reset_index(drop=True)
        else:
            # "Cognome e Nome (A→Z)": ordina per Nome e, per la stessa persona, per Inizio
            by = []
            if "Cognome e Nome" in g.columns:
                by.append("Cognome e Nome")
            if "Inizio" in g.columns:
                by.append("Inizio")
            if by:
                g = g.sort_values(by=by).reset_index(drop=True)

            # compattamento: non ripetere Nome/Matricola per la stessa persona
            if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
                same_person = (
                    g[["Cognome e Nome", "Matricola"]]
                    .eq(g[["Cognome e Nome", "Matricola"]].shift(1))
                ).all(axis=1)
                g.loc[same_person, ["Cognome e Nome", "Matricola"]] = ""

        st.markdown(f"### **{res}**")
        g_nosede = g.drop(columns=["Residenza"], errors="ignore")
        st.dataframe(g_nosede, use_container_width=True, hide_index=True)
else:
    st.dataframe(df_out, use_container_width=True, hide_index=True)


        # Download Excel (senza Residenza)
        xls_buf = io.BytesIO()
        with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
            df_out.drop(columns=["Residenza"], errors="ignore").to_excel(
                writer, sheet_name="ServizioGiornaliero", index=False
            )
        st.download_button("⬇️ Scarica Excel", data=xls_buf.getvalue(), file_name="ServizioGiornaliero.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF raggruppato per Residenza
        with tempfile.TemporaryDirectory() as td:
            pdf_path = Path(td) / "ServizioGiornaliero.pdf"
            inner_sort = "inizio" if "Inizio" in inner_sort_choice else "nome"
            build_pdf(pdf_path, df_out, meta, logo_path if logo_path.exists() else None,
                      title=TITLE, inner_sort=inner_sort)
            st.download_button("⬇️ Scarica PDF", data=pdf_path.read_bytes(),
                               file_name="ServizioGiornaliero.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"Errore durante l'elaborazione: {e}")
