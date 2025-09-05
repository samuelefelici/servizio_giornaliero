# app.py
import os, sys, re, io, tempfile
from pathlib import Path

# --- Path per import locali ---
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd

REST_CODES = {"R", "RR"}  # riposo: non va in bold

# --- Import moduli del progetto (con gestione errore) ---
try:
    from src.process import read_input_excel, transform_dataframe, debug_probe  # debug opzionale
    from src.pdf_export import build_pdf
    from src.constants import TITLE, DISPLAY_ORDER
except Exception as e:
    import traceback
    st.set_page_config(page_title="Servizio Giornaliero", layout="wide")
    st.error(f"Errore durante l'import dei moduli: {e}")
    st.code("".join(traceback.format_exception(*sys.exc_info())))
    st.stop()

# --- Config pagina ---
st.set_page_config(page_title="Servizio Giornaliero", layout="wide")
# forza il container principale a occupare tutta la viewport
st.markdown("""
<style>
  .block-container { max-width: 100% !important; padding-left: 24px; padding-right: 24px; }
</style>
""", unsafe_allow_html=True)

st.title("📋 Servizio Giornaliero – ExtraUrbano (Python)")
st.caption("Drag & drop del file Excel (.xls/.xlsx), pulizia automatica, anteprima e export PDF/Excel (raggruppato per deposito).")

cfg_dir = Path("config")
assets_dir = Path("assets")
logo_path = assets_dir / "logo.jpg"

# ====================== Helper ======================

def _reorder_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Ordina colonne secondo DISPLAY_ORDER e nasconde 'Residenza' nella vista/export."""
    df2 = df.drop(columns=["Residenza"], errors="ignore").copy()
    cols = [c for c in DISPLAY_ORDER if c in df2.columns]
    return df2[cols] if cols else df2

def _res_to_prefix(res: str) -> str | None:
    """Mappa la residenza a un prefisso atteso (JU o J sono famigliari)."""
    if not isinstance(res, str): return None
    r = res.upper().replace("_", " ")
    if "JESI URBANO" in r or r.strip() == "JU": return "JU"
    if "JESI" in r or r.strip() == "J":       return "J"
    if "MARINA" in r or r.strip() == "M":     return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r.strip() == "C": return "C"
    if "OSIMO" in r or r.strip() == "O":      return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r.strip() == "F":    return "F"
    if "POLVERIGI" in r or r.strip() == "P":  return "P"
    if "OSTRA" in r or r.strip() == "D":      return "D"
    if "BELVED" in r or r.strip() == "B" or "DEPBELVE" in r:     return "B"
    if "ANCONA" in r or r.strip() == "A":     return "A"
    return None

def _turno_bucket(turno: str) -> str | None:
    """
    Riduce il codice turno a un 'bucket' di confronto:
      - 'JU…' -> 'JU'
      - 'J…' (non JU) -> 'J'
      - altrimenti prima lettera (M, C, O, F, P, D, A, B…)
      - esclude 'Assente' e R/RR
    """
    if not isinstance(turno, str):
        turno = str(turno)
    s = turno.strip()
    if not s or s.upper() == "ASSENTE" or s.upper() in REST_CODES:
        return None
    m = re.match(r"[A-Za-z]+", s)
    if not m: return None
    up = m.group(0).upper()
    if up.startswith("JU"): return "JU"
    if up.startswith("J"):  return "J"
    return up[0]

def _accepted_prefixes_for_res(prefix: str | None) -> set[str]:
    """Prefissi considerati 'di casa' per la residenza (J e JU sono equivalenti)."""
    if prefix in ("J", "JU"):
        return {"J", "JU"}
    return {prefix} if prefix else set()

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    """True se il turno è 'fuori deposito', con J/JU equivalenti ed esclusi R/RR e Assente."""
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)

    def _is_trasferta(row) -> bool:
        turn = str(row["Turno"]).strip().upper()
        if turn in REST_CODES or turn == "ASSENTE":
            return False
        rp = _res_to_prefix(row["Residenza"])
        tb = _turno_bucket(row["Turno"])
        if rp is None or tb is None:
            return False
        return tb not in _accepted_prefixes_for_res(rp)

    return df.apply(_is_trasferta, axis=1)

# ---- Styled HTML per anteprima full-width con grassetto su trasferte ----
def _style_bold_subset(df_disp: pd.DataFrame, mask: pd.Series, subset_cols: list[str]) -> str:
    """
    HTML di una tabella:
      - indice nascosto
      - larghezza vera 100vw (oltre i padding di Streamlit)
      - grassetto solo su colonne subset per le righe in trasferta (mask=True)
    """
    mask = mask.reindex(df_disp.index).fillna(False)

    def _bold_if_trasferta(row):
        return ["font-weight: bold"] * len(row) if mask.loc[row.name] else [""] * len(row)

    sty = df_disp.style.apply(_bold_if_trasferta, axis=1, subset=subset_cols)
    # nascondi indice (compat con varie versioni pandas)
    try:    sty = sty.hide(axis="index")
    except: 
        try: sty = sty.hide_index()
        except: pass

    sty = sty.set_table_styles([
        {"selector": "table",    "props": [("width", "100%"), ("border-collapse", "collapse"), ("table-layout", "fixed")]},
        {"selector": "th, td",   "props": [("border", "1px solid #e6e6e6"), ("padding", "6px 8px")]},
        {"selector": "thead th", "props": [("background-color", "#f7f7f7"), ("font-weight", "600")]},
    ], overwrite=False)

    html = sty.to_html()
    # wrapper che forza la tabella ad estendersi a tutta la viewport
    wrapper = f"<div style='width:100vw;margin-left:calc(50% - 50vw);margin-right:calc(50% - 50vw);'>{html}</div>"
    return wrapper

def _show_styled_table(df_disp: pd.DataFrame, mask: pd.Series):
    subset_cols = [c for c in ["Turno", "Inizio", "Fine"] if c in df_disp.columns]
    if not subset_cols:
        # fallback: Arrow dataframe comunque a piena larghezza del container
        st.dataframe(df_disp, use_container_width=True, hide_index=True)
        return
    html = _style_bold_subset(df_disp, mask, subset_cols)
    st.markdown(html, unsafe_allow_html=True)

# ====================== UI ======================

uploaded = st.file_uploader("Trascina qui il file oppure selezionalo", type=["xls","xlsx"])

col1, col2 = st.columns([1,1], vertical_alignment="center")
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

debug_mode = st.checkbox("🧪 Modalità debug", value=False, help="Mostra info sniffer/header del file caricato.")

# ====================== Azione ======================

if st.button("▶️ Elabora", type="primary", use_container_width=True):
    if not uploaded:
        st.warning("Carica prima un file.")
        st.stop()

    # Debug opzionale (si esegue su stream; la nostra lettura fa seek(0))
    if debug_mode:
        try:
            info = debug_probe(uploaded)
            with st.expander("Dettagli debug (header/sniffer)"):
                st.json(info)
        except Exception as e:
            st.warning(f"Debug non riuscito: {e}")

    try:
        # Lettura + pipeline base
        df, meta = read_input_excel(uploaded)
        df_out   = transform_dataframe(df, cfg_dir)

        # Filtro Assente per vista/export
        df_view = df_out.copy()
        if not show_absent and "Turno" in df_view.columns:
            before = len(df_view)
            df_view = df_view[df_view["Turno"].astype(str).str.strip() != "Assente"].reset_index(drop=True)
            hidden = before - len(df_view)
            if hidden > 0:
                st.info(f"Righe 'Assente' nascoste: {hidden}")

        # Feedback meta
        st.success(
            f"File elaborato. Data: {meta.get('data','?')} – {meta.get('giorno','?')} "
            f"(fonte: {meta.get('origine','?')})"
        )
        st.subheader("Anteprima per deposito")

        # Anteprima per deposito con grassetto per trasferte
        if "Residenza" in df_view.columns:
            res_list = sorted(df_view["Residenza"].dropna().astype(str).unique())
            for res in res_list:
                g = df_view[df_view["Residenza"].astype(str) == res].copy()
                if g.empty:
                    continue

                # Ordinamento interno
                if inner_sort_choice.startswith("Inizio") and "Inizio" in g.columns:
                    by = ["Inizio"] + (["Cognome e Nome"] if "Cognome e Nome" in g.columns else [])
                    g = g.sort_values(by=by).reset_index(drop=True)
                else:
                    by = []
                    if "Cognome e Nome" in g.columns: by.append("Cognome e Nome")
                    if "Inizio" in g.columns:         by.append("Inizio")
                    if by: g = g.sort_values(by=by).reset_index(drop=True)
                    # Compattazione
                    if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
                        same_person = g[["Cognome e Nome", "Matricola"]].eq(
                            g[["Cognome e Nome", "Matricola"]].shift(1)
                        ).all(axis=1)
                        g.loc[same_person, ["Cognome e Nome", "Matricola"]] = ""

                st.markdown(f"### **{res}**")
                g_disp = _reorder_for_display(g)
                mask   = _trasferta_mask(g)
                _show_styled_table(g_disp, mask)
        else:
            g_disp = _reorder_for_display(df_view)
            mask   = _trasferta_mask(df_view)
            _show_styled_table(g_disp, mask)

        # --- Export Excel (ordine colonne, filtro Assente applicato) ---
        xls_buf = io.BytesIO()
        with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
            _reorder_for_display(df_view).to_excel(writer, sheet_name="ServizioGiornaliero", index=False)
        st.download_button(
            "⬇️ Scarica Excel",
            data=xls_buf.getvalue(),
            file_name="ServizioGiornaliero.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # --- Export PDF ---
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

    except Exception as e:
        import traceback
        st.error(f"Errore durante l'elaborazione: {e}")
        st.code("".join(traceback.format_exception(*sys.exc_info())))
