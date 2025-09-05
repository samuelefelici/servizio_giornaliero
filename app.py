# app.py
import os, sys, re, io, tempfile
from pathlib import Path

# --- Path per import locali ---
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd

REST_CODES = {"R", "RR"}   # riposo

# --- Import moduli del progetto (con gestione errore) ---
try:
    from src.process import read_input_excel, transform_dataframe, debug_probe
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

# CSS: contenitore a tutta larghezza e stile tabelle HTML (pandas Styler)
st.markdown("""
<style>
.block-container {max-width: 100% !important; padding-left: 16px; padding-right: 16px;}
.serv-table-wrap {width: 100%;}
.serv-table-wrap table {width: 100% !important; table-layout: fixed; border-collapse: collapse;}
.serv-table-wrap th {background:#f5f5f5; text-align:left; padding:6px;}
.serv-table-wrap td {padding:6px; vertical-align: top;}
/* centro le colonne 4 e 5 (Inizio/Fine) */
.serv-table-wrap th:nth-child(4), .serv-table-wrap td:nth-child(4),
.serv-table-wrap th:nth-child(5), .serv-table-wrap td:nth-child(5) {text-align:center;}
/* colonna Note: usa il resto della larghezza e va a capo */
.serv-table-wrap th:nth-child(6), .serv-table-wrap td:nth-child(6) {
  width: 40%; white-space: normal; word-break: break-word;
}
</style>
""", unsafe_allow_html=True)

st.title("📋 Servizio Giornaliero – ExtraUrbano (Python)")
st.caption("Drag & drop del file Excel (.xls/.xlsx), pulizia automatica, anteprima e export PDF/Excel (raggruppato per deposito).")

cfg_dir   = Path("config")
assets_dir= Path("assets")
logo_path = assets_dir / "logo.jpg"

# ====================== Helper ======================

def _reorder_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Ordina colonne secondo DISPLAY_ORDER e nasconde 'Residenza'."""
    df2  = df.drop(columns=["Residenza"], errors="ignore").copy()
    cols = [c for c in DISPLAY_ORDER if c in df2.columns]
    return df2[cols] if cols else df2

def _res_to_prefix(res: str) -> str | None:
    """Mappa residenza a prefisso atteso (J e JU considerate ‘di casa’ insieme)."""
    if not isinstance(res, str): return None
    r = res.upper().replace("_", " ").strip()
    if "JESI URBANO" in r or r == "JU": return "JU"
    if "JESI" in r or r == "J":       return "J"
    if "MARINA" in r or r == "M":     return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r == "C": return "C"
    if "OSIMO" in r or r == "O":      return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r == "F":    return "F"
    if "POLVERIGI" in r or r == "P":  return "P"
    if "OSTRA" in r or r == "D":      return "D"
    if "BELVED" in r or r == "B" or "DEPBELVE" in r:     return "B"
    if "ANCONA" in r or r == "A":     return "A"
    return None

def _turno_bucket(turno: str) -> str | None:
    """Bucket turno: JU…, J… (non JU) o prima lettera; esclude Assente/R/RR."""
    if not isinstance(turno, str): turno = str(turno)
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
    if prefix in ("J", "JU"): return {"J", "JU"}
    return {prefix} if prefix else set()

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    """True se il turno è ‘fuori deposito’ (J/JU equivalenti; R/RR/Assente esclusi)."""
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

def _styled_html_table(g_disp: pd.DataFrame, trasferta_mask: pd.Series) -> str:
    """
    Ritorna HTML della tabella (pandas Styler) a larghezza piena.
    Bold su Turno/Inizio/Fine quando trasferta_mask è True.
    """
    # funzione per righe boldate (subset solo su colonne interessate)
    subset_cols = [c for c in ["Turno", "Inizio", "Fine"] if c in g_disp.columns]

    def _bold_if_trasferta(row):
        return (["font-weight: bold"] * len(row)) if trasferta_mask.loc[row.name] else ([""] * len(row))

    sty = (g_disp.style
           .hide(axis="index")
           .apply(_bold_if_trasferta, axis=1, subset=subset_cols)
           .set_table_styles([
               {"selector": "table", "props": [("width","100%"), ("table-layout","fixed"), ("border-collapse","collapse")]},
               {"selector": "th",    "props": [("background","#f5f5f5"), ("text-align","left"), ("padding","6px")]},
               {"selector": "td",    "props": [("padding","6px"), ("vertical-align","top")]},
               # centro Inizio/Fine
               {"selector": "th:nth-child(4), td:nth-child(4), th:nth-child(5), td:nth-child(5)",
                "props": [("text-align","center")]},
               # colonna Note: wrap e un po' più larga
               {"selector": "th:nth-child(6), td:nth-child(6)",
                "props": [("width","40%"), ("white-space","normal"), ("word-break","break-word")]}
           ], overwrite=False)
          )
    return f'<div class="serv-table-wrap">{sty.to_html()}</div>'

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

debug_mode = st.checkbox("🧪 Modalità debug", value=False,
                         help="Mostra info sniffer/header del file caricato.")

# ====================== Azione ======================

if st.button("▶️ Elabora", type="primary", use_container_width=True):
    if not uploaded:
        st.warning("Carica prima un file.")
        st.stop()

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
            before  = len(df_view)
            df_view = df_view[df_view["Turno"].astype(str).str.strip() != "Assente"].reset_index(drop=True)
            hidden  = before - len(df_view)
            if hidden > 0:
                st.info(f"Righe 'Assente' nascoste: {hidden}")

        # Feedback meta
        st.success(f"File elaborato. Data: {meta.get('data','?')} – {meta.get('giorno','?')} "
                   f"(fonte: {meta.get('origine','?')})")
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
                    by = ["Inizio"]
                    if "Cognome e Nome" in g.columns: by.append("Cognome e Nome")
                    g = g.sort_values(by=by).reset_index(drop=True)
                else:
                    by = []
                    if "Cognome e Nome" in g.columns: by.append("Cognome e Nome")
                    if "Inizio" in g.columns:         by.append("Inizio")
                    if by:
                        g = g.sort_values(by=by).reset_index(drop=True)
                    # compattazione Nome/Matricola ripetuti consecutivi
                    if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
                        same_person = g[["Cognome e Nome", "Matricola"]].eq(
                            g[["Cognome e Nome", "Matricola"]].shift(1)
                        ).all(axis=1)
                        g.loc[same_person, ["Cognome e Nome", "Matricola"]] = ""

                st.markdown(f"### **{res}**")

                g_disp = _reorder_for_display(g)
                mask   = _trasferta_mask(g)     # calcolata sull’indice originale
                html   = _styled_html_table(g_disp, mask)
                st.markdown(html, unsafe_allow_html=True)

        else:
            st.markdown("### **TUTTI**")
            g_disp = _reorder_for_display(df_view)
            mask   = _trasferta_mask(df_view)
            html   = _styled_html_table(g_disp, mask)
            st.markdown(html, unsafe_allow_html=True)

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
            pdf_path   = Path(td) / "ServizioGiornaliero.pdf"
            inner_sort = "inizio" if inner_sort_choice.startswith("Inizio") else "nome"
            build_pdf(pdf_path, df_view, meta,
                      logo_path if logo_path.exists() else None,
                      title=TITLE, inner_sort=inner_sort)
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
