import os, sys, re
from pathlib import Path
from src.constants import TITLE, DISPLAY_ORDER  # <-- usiamo anche DISPLAY_ORDER

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd
import io, tempfile

try:
    from src.process import read_input_excel, transform_dataframe
    from src.pdf_export import build_pdf
except Exception as e:
    import traceback
    st.set_page_config(page_title="Servizio Giornaliero", layout="wide")
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
with col2:
    show_absent = st.checkbox(
        "Mostra anche gli 'Assente'",
        value=True,
        help="Se deselezionato nasconde le righe con Turno = Assente (vale per anteprima ed export)."
    )

debug_mode = st.checkbox("🧪 Modalità debug", value=False)

# Helper: riordina e rimuove 'Residenza' per visualizzazione/export
def _reorder_for_display(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.drop(columns=["Residenza"], errors="ignore").copy()
    cols = [c for c in DISPLAY_ORDER if c in df2.columns]
    return df2[cols] if cols else df2

if st.button("▶️ Elabora"):
    if not uploaded:
        st.warning("Carica prima un file.")
        st.stop()

    # Debug opzionale (prima della lettura 'vera' per non alterare lo stream)
    if debug_mode:
        try:
            from src.process import debug_probe
            info = debug_probe(uploaded)
            with st.expander("Dettagli debug (header/sniffer)"):
                st.json(info)
        except Exception as e:
            st.warning(f"Debug non riuscito: {e}")

    try:
        # Lettura + pipeline base (filtri matricole/turni, rinomine, 'Assente', ordinamenti)
        df, meta = read_input_excel(uploaded)
        df_out = transform_dataframe(df, cfg_dir)

        # Applica filtro "Assente" alla vista (e agli export)
        df_view = df_out.copy()
        if not show_absent and "Turno" in df_view.columns:
            df_view = (
                df_view[df_view["Turno"].astype(str).str.strip() != "Assente"]
                .reset_index(drop=True)
            )
            nascosti = len(df_out) - len(df_view)
            if nascosti > 0:
                st.info(f"Righe 'Assente' nascoste: {nascosti}")

        # Anteprima per deposito
        st.success(
            f"File elaborato. Data: {meta.get('data','?')} – {meta.get('giorno','?')} "
            f"(fonte: {meta.get('origine','?')})"
        )
        st.subheader("Anteprima per deposito")

        if "Residenza" in df_view.columns:
            res_list = sorted(df_view["Residenza"].dropna().astype(str).unique())
            for res in res_list:
                g = df_view[df_view["Residenza"].astype(str) == res].copy()
                if g.empty:
                    continue

                # Ordinamento interno
                if inner_sort_choice.startswith("Inizio") and "Inizio" in g.columns:
                    by = ["Inizio"]
                    if "Cognome e Nome" in g.columns:
                        by.append("Cognome e Nome")
                    g = g.sort_values(by=by).reset_index(drop=True)
                else:
                    by = []
                    if "Cognome e Nome" in g.columns:
                        by.append("Cognome e Nome")
                    if "Inizio" in g.columns:
                        by.append("Inizio")
                    if by:
                        g = g.sort_values(by=by).reset_index(drop=True)

                    # Compattazione: non ripetere Nome/Matricola per righe consecutive della stessa persona
                    if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
                        same_person = g[["Cognome e Nome", "Matricola"]].eq(
                            g[["Cognome e Nome", "Matricola"]].shift(1)
                        ).all(axis=1)
                        g.loc[same_person, ["Cognome e Nome", "Matricola"]] = ""

st.markdown(f"### **{res}**")
g_disp = _reorder_for_display(g)

# calcola maschera di trasferta sulla tabella originale 'g'
mask = _trasferta_mask(g)

# applica bold solo a Turno/Inizio/Fine (se presenti)
subset_cols = [c for c in ["Turno", "Inizio", "Fine"] if c in g_disp.columns]
if subset_cols:
    def _bold_if_trasferta(row):
        # row contiene solo le colonne del subset (grazie a subset=...)
        return ["font-weight: bold"] * len(row) if mask.loc[row.name] else [""] * len(row)
    styled = g_disp.style.apply(_bold_if_trasferta, axis=1, subset=subset_cols)
    st.dataframe(styled, use_container_width=True, hide_index=True)
else:
    st.dataframe(g_disp, use_container_width=True, hide_index=True)

        else:
            df_tmp = _reorder_for_display(df_view)
            st.dataframe(df_tmp, use_container_width=True, hide_index=True)

        # Download Excel (coerente col filtro 'Assente' e ordine colonne)
        xls_buf = io.BytesIO()
        with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
            to_xls = _reorder_for_display(df_view)
            to_xls.to_excel(writer, sheet_name="ServizioGiornaliero", index=False)
        st.download_button(
            "⬇️ Scarica Excel",
            data=xls_buf.getvalue(),
            file_name="ServizioGiornaliero.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # PDF raggruppato per Residenza (coerente col filtro 'Assente')
        with tempfile.TemporaryDirectory() as td:
            pdf_path = Path(td) / "ServizioGiornaliero.pdf"
            inner_sort = "inizio" if inner_sort_choice.startswith("Inizio") else "nome"
            build_pdf(
                pdf_path, df_view, meta,
                logo_path if logo_path.exists() else None,
                title=TITLE, inner_sort=inner_sort
            )
            st.download_button(
                "⬇️ Scarica PDF",
                data=pdf_path.read_bytes(),
                file_name="ServizioGiornaliero.pdf",
                mime="application/pdf"
            )

    def _res_to_prefix(res: str) -> str | None:
    if not isinstance(res, str):
        return None
    r = res.upper().replace("_", " ")
    # gestisci alias/abbreviazioni frequenti
    if "JESI URBANO" in r or r.strip() == "JU":
        return "JU"
    if "JESI" in r or r.strip() == "J":
        return "J"
    if "MARINA" in r or r.strip() == "M":
        return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r.strip() == "C":
        return "C"
    if "OSIMO" in r or r.strip() == "O":
        return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r.strip() == "F":
        return "F"
    if "POLVERIGI" in r or r.strip() == "P":
        return "P"
    if "OSTRA" in r or r.strip() == "D":
        return "D"
    if "BELVED" in r or r.strip() == "B" or "DEPBELVE" in r:
        return "B"
    if "ANCONA" in r or r.strip() == "A":
        return "A"
    return None

def _turno_prefix(turno: str) -> str | None:
    if not isinstance(turno, str):
        turno = str(turno)
    m = re.match(r"[A-Za-z]+", turno.strip())
    return m.group(0).upper() if m else None

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    """True se il prefisso del turno è diverso da quello atteso per la residenza."""
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)
    expected = df["Residenza"].apply(_res_to_prefix)
    actual = df["Turno"].astype(str).apply(_turno_prefix)
    return expected.notna() & actual.notna() & (expected != actual)


    except Exception as e:
        import traceback
        st.error(f"Errore durante l'elaborazione: {e}")
        st.code("".join(traceback.format_exception(*sys.exc_info())))
