import os, sys, re, io, tempfile
from pathlib import Path
from html import escape as html_escape
from datetime import datetime
import streamlit as st
import pandas as pd

# --- Path per import locali ---
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

REST_CODES = {"R", "RR"}   # riposo

# Eccezioni GLOBALI
GLOBAL_EXC_PATTERNS = (r"^IAST$", r"^N$",)

EXC_ANCONA_PREFIXES = ("D1R1","D1R2","D1R5","D2R1","D2R2","D2R3","D2R6","NP","ASC","V5","LU","MA","ME","GI","VE","SA","DO")
EXC_ANCONA_PATTERNS  = tuple(rf"^{re.escape(p)}" for p in EXC_ANCONA_PREFIXES)

# --- Import moduli del progetto ---
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

# CSS
st.markdown("""
<style>
.block-container {max-width: 100% !important; padding-left: 16px; padding-right: 16px;}
</style>
""", unsafe_allow_html=True)

st.title("📋 Servizio Giornaliero")
st.caption("Drag & drop del file Excel (.xls/.xlsx), pulizia automatica, anteprima e export PDF/Excel (raggruppato per deposito).")

cfg_dir   = Path("config")
assets_dir= Path("assest")
logo_path = assets_dir / "logo.jpg"

# ====================== Session state ======================
st.session_state.setdefault("df_view", None)
st.session_state.setdefault("meta", None)
st.session_state.setdefault("last_inner_sort", "nome")
st.session_state.setdefault("extra_rows", pd.DataFrame(columns=[
    "Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Residenza","_added"
]))
st.session_state.setdefault("transfer_menu", None)  # può essere None, "manuale", "file"
st.session_state.setdefault("imported_rows", [])    # per trasferte da file
st.session_state.setdefault("manual_rows", [])      # per trasferte manuali

# ====================== Helper ======================

def _turno_to_residenza_name(turno: str) -> str | None:
    t = str(turno or "").upper().strip().replace(".", "").replace(" ", "")
    prefix_to_res = {
        "JU": "JESI URBANO", "J": "JESI EXTRAURBANO", "M": "MARINA",
        "C": "CASTELFIDARDO", "O": "OSIMO", "F": "FILOTTRANO", "P": "POLVERIGI",
        "D": "OSTRA", "B": "BELVEDERE", "A": "ANCONA",
    }
    if t.startswith("JU"): return prefix_to_res["JU"]
    if t.startswith("J"):  return prefix_to_res["J"]
    return prefix_to_res.get(t[:1])

def _parse_time_like(s: str | None) -> str:
    s = (s or "").strip()
    if not s: return ""
    m = re.match(r"^([01]?\d|2[0-3]):([0-5]\d)$", s)
    return s if m else ""

def _build_extra_df(rows: list[dict]) -> pd.DataFrame:
    if not rows: return pd.DataFrame()
    df = pd.DataFrame(rows).rename(columns={
        "Nominativo": "Cognome e Nome",
        "Turno fuori residenza": "Turno",
    })
    for c in ["Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Residenza"]:
        if c not in df.columns: df[c] = ""
    df["Residenza"] = df.apply(lambda r: r["Residenza"] or (_turno_to_residenza_name(r["Turno"]) or ""), axis=1)
    df["Inizio"] = df["Inizio"].map(_parse_time_like)
    df["Fine"]   = df["Fine"].map(_parse_time_like)
    df["_added"] = True
    keep = ["Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Residenza","_added"]
    return df[keep]

def _do_rerun():
    fn = getattr(st, "rerun", None)
    if callable(fn): fn()
    else: st.experimental_rerun()

# ====================== UI ======================

uploaded = st.file_uploader("Trascina qui il file oppure selezionalo", type=["xls","xlsx"])

col1, col2 = st.columns([1,1], vertical_alignment="center")
with col1:
    inner_sort_choice = st.radio(
        "Ordina dentro ciascun deposito per:",
        ["Cognome e Nome (A→Z)", "Inizio (orario)"],
        horizontal=True, index=0
    )
with col2:
    show_absent = st.checkbox(
        "Mostra anche gli 'Assente'", value=True,
        help="Se deselezionato nasconde le righe con Turno = Assente (vale per anteprima ed export)."
    )

debug_mode = st.checkbox("🧪 Modalità debug", value=False,
                         help="Mostra info sniffer/header del file caricato.")

# ====================== Azione: ELABORA ======================

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
        df, meta = read_input_excel(uploaded)
        df_out   = transform_dataframe(df, cfg_dir)

        df_view = df_out.copy()
        if not show_absent and "Turno" in df_view.columns:
            before  = len(df_view)
            df_view = df_view[df_view["Turno"].astype(str).str.strip() != "Assente"].reset_index(drop=True)
            hidden  = before - len(df_view)
            if hidden > 0: st.info(f"Righe 'Assente' nascoste: {hidden}")

        if not st.session_state["extra_rows"].empty:
            df_view = pd.concat([df_view, st.session_state["extra_rows"]], ignore_index=True)

        st.session_state["df_view"] = df_view
        st.session_state["meta"] = meta
        st.session_state["last_inner_sort"] = "inizio" if inner_sort_choice.startswith("Inizio") else "nome"

    except Exception as e:
        import traceback
        st.error(f"Errore durante l'elaborazione: {e}")
        st.code("".join(traceback.format_exception(*sys.exc_info())))

# ====================== Trasferte UI: Scelta metodo inserimento ======================

if st.session_state["df_view"] is not None and st.session_state["meta"] is not None:

    # Menu di inserimento trasferte
    if st.session_state["transfer_menu"] is None:
        if st.button("➕ Inserisci Trasferte", use_container_width=True):
            st.session_state["transfer_menu"] = "menu"
            _do_rerun()

    elif st.session_state["transfer_menu"] == "menu":
        st.markdown("#### Scegli come inserire trasferte")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Inserisci da File"):
                st.session_state["transfer_menu"] = "file"
                _do_rerun()
        with col2:
            if st.button("Inserisci Manualmente"):
                st.session_state["transfer_menu"] = "manuale"
                _do_rerun()

    # Inserimento trasferte da file
    elif st.session_state["transfer_menu"] == "file":
        st.markdown("### Importa trasferte da file Excel")
        imported_file = st.file_uploader("Scegli il file XLS/XLSX", type=["xls", "xlsx"], key="upload_transfer_file")
        if imported_file:
            try:
                df_import = pd.read_excel(imported_file, header=0)
                preview_rows = []
                for _, row in df_import.iterrows():
                    preview_rows.append({
                        "Matricola": row.iloc[1],      # colonna 2
                        "Cognome e Nome": row.iloc[2], # colonna 3
                        "Turno": row.iloc[4],          # colonna 5
                        "Inizio": "",
                        "Fine": "",
                        "Indennità e note": "",
                        "Residenza": "",
                    })
                st.session_state["imported_rows"] = preview_rows
            except Exception as e:
                st.error(f"Errore import: {e}")
        # Tabella editabile di anteprima
        if st.session_state["imported_rows"]:
            st.markdown("#### Anteprima trasferte importate (modificabili)")
            df_preview = pd.DataFrame(st.session_state["imported_rows"])
            edited = st.data_editor(df_preview, num_rows="dynamic", use_container_width=True, key="imported_table")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Conferma trasferte importate"):
                    extra_df = _build_extra_df(edited.to_dict("records"))
                    st.session_state["extra_rows"] = pd.concat([st.session_state["extra_rows"], extra_df], ignore_index=True)
                    st.session_state["df_view"] = pd.concat([st.session_state["df_view"], extra_df], ignore_index=True)
                    st.session_state["transfer_menu"] = None
                    st.session_state["imported_rows"] = []
                    st.success(f"Inserite {len(extra_df)} trasferte dal file.")
                    _do_rerun()
            with col2:
                if st.button("Annulla"):
                    st.session_state["transfer_menu"] = None
                    st.session_state["imported_rows"] = []
                    _do_rerun()

    # Inserimento trasferte manuale
    elif st.session_state["transfer_menu"] == "manuale":
        st.markdown("### Inserisci trasferte manualmente")
        start_rows = 5
        manual_rows = st.session_state.get("manual_rows", [{"Matricola":"","Cognome e Nome":"","Turno":"","Inizio":"","Fine":""} for _ in range(start_rows)])
        df_manual = pd.DataFrame(manual_rows)
        edited_manual = st.data_editor(df_manual, num_rows="dynamic", use_container_width=True, key="manual_table")
        st.session_state["manual_rows"] = edited_manual.to_dict("records")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Conferma trasferte manuali"):
                rows = edited_manual.replace({pd.NA:"", None:""}).to_dict("records")
                rows = [r for r in rows if any(str(v).strip() for v in r.values())]
                extra_df = _build_extra_df(rows)
                st.session_state["extra_rows"] = pd.concat([st.session_state["extra_rows"], extra_df], ignore_index=True)
                st.session_state["df_view"] = pd.concat([st.session_state["df_view"], extra_df], ignore_index=True)
                st.session_state["manual_rows"] = []
                st.session_state["transfer_menu"] = None
                st.success(f"Inserite {len(extra_df)} trasferte manuali.")
                _do_rerun()
        with col2:
            if st.button("Annulla"):
                st.session_state["manual_rows"] = []
                st.session_state["transfer_menu"] = None
                _do_rerun()

    # Pulsante per svuotare trasferte extra
    if not st.session_state["extra_rows"].empty:
        if st.button("🗑️ Svuota trasferte aggiunte", use_container_width=True):
            st.session_state["extra_rows"] = st.session_state["extra_rows"].iloc[0:0]
            base = st.session_state["df_view"]
            if "_added" in base.columns:
                base = base[~base["_added"].fillna(False)].copy()
            st.session_state["df_view"] = base
            st.info("Trasferte cancellate.")
            _do_rerun()

    # ---- Anteprima
    st.markdown("### Anteprima per deposito")
    if st.session_state["df_view"] is not None:
        st.dataframe(st.session_state["df_view"])

    # --- Export PDF ---
    with tempfile.TemporaryDirectory() as td:
        temp_dir = Path(td)
        inner_sort = "inizio" if inner_sort_choice.startswith("Inizio") else "nome"
        pdf_path = build_pdf(
            temp_dir,
            st.session_state["df_view"],
            st.session_state["meta"],
            logo_path if logo_path.exists() else None,
            title=TITLE, inner_sort=inner_sort,
            exported_at=datetime.now()
        )
        st.download_button(
            "⬇️ Scarica PDF",
            data=pdf_path.read_bytes(),
            file_name=pdf_path.name,
            mime="application/pdf"
        )
