import os, sys, re, io, tempfile
from pathlib import Path
from datetime import datetime, timedelta

# --- Path per import locali ---
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd

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

# CSS (solo layout, nessuna colorazione speciale sulle celle)
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
st.session_state.setdefault("show_transfer_ui", False)
st.session_state.setdefault("extra_rows", pd.DataFrame(columns=[
    "Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Residenza","_added"
]))

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

# ====================== Trasferte + Anteprima Editabile + Export ======================

if st.session_state["df_view"] is not None and st.session_state["meta"] is not None:
    df_view = st.session_state["df_view"]

    c1, c2 = st.columns([1,3])
    with c1:
        if st.button("➕ Inserisci trasferte", use_container_width=True):
            st.session_state["show_transfer_ui"] = not st.session_state["show_transfer_ui"]

    # ---- ANTEPRIMA EDITABILE (NO COLORAZIONE) ----
    st.markdown("### Modifica trasferte e note direttamente nella tabella sottostante")

    mask_manual = df_view.get("_added", False)
    df_manual = df_view[mask_manual].copy()
    df_original = df_view[~mask_manual].copy()

    columns_all = ["Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Residenza"]

    # Tabella editabile per trasferte manuali (tutte le colonne)
    edited_manual = None
    if not df_manual.empty:
        st.markdown("#### Trasferte manuali (tutte le colonne editabili)")
        editable_manual = df_manual[columns_all].copy()
        edited_manual = st.data_editor(
            editable_manual,
            num_rows="fixed",
            use_container_width=True,
            key="edit_manual_rows"
        )

    # Tabella editabile per servizio da file (solo note)
    edited_original = None
    if not df_original.empty and "Indennità e note" in df_original.columns:
        st.markdown("#### Servizio da file (solo colonna 'Indennità e note' editabile)")
        editable_notes = df_original[["Indennità e note"]].copy()
        edited_original = st.data_editor(
            editable_notes,
            num_rows="fixed",
            use_container_width=True,
            key="edit_original_notes"
        )

    col_mod1, col_mod2 = st.columns([1,1])
    with col_mod1:
        if st.button("💾 Salva modifiche tabella", use_container_width=True):
            # Aggiorna trasferte manuali
            extra_df = pd.DataFrame()
            if edited_manual is not None:
                extra_df = edited_manual.copy()
                for c in columns_all:
                    if c not in extra_df.columns:
                        extra_df[c] = ""
                extra_df["Residenza"] = extra_df.apply(
                    lambda r: r["Residenza"] or (_turno_to_residenza_name(r["Turno"]) or ""),
                    axis=1
                )
                extra_df["Inizio"] = extra_df["Inizio"].map(_parse_time_like)
                extra_df["Fine"] = extra_df["Fine"].map(_parse_time_like)
                extra_df["_added"] = True
                st.session_state["extra_rows"] = extra_df

            # Aggiorna solo la colonna note per le righe originali
            df_original_new = df_original.copy()
            if edited_original is not None:
                df_original_new["Indennità e note"] = edited_original["Indennità e note"]

            # Aggiorna il DataFrame principale
            st.session_state["df_view"] = pd.concat([df_original_new, st.session_state["extra_rows"]], ignore_index=True)
            st.success("Modifiche salvate!")
            _do_rerun()

    with col_mod2:
        if not st.session_state["extra_rows"].empty:
            if st.button("🗑️ Svuota trasferte aggiunte", use_container_width=True):
                st.session_state["extra_rows"] = st.session_state["extra_rows"].iloc[0:0]
                base = st.session_state["df_view"]
                if "_added" in base.columns:
                    base = base[~base["_added"].fillna(False)].copy()
                st.session_state["df_view"] = base
                st.info("Trasferte cancellate.")
                _do_rerun()

    # --- UI per aggiunta nuove trasferte ---
    if st.session_state["show_transfer_ui"]:
        st.markdown("### Inserisci trasferte")
        st.caption("Compila le righe (max 30). Obbligatori: Matricola, Cognome e Nome, Turno.")

        start_rows = 5
        default_rows = pd.DataFrame(
            [{"Matricola":"","Cognome e Nome":"","Turno":"","Inizio":"","Fine":""} for _ in range(start_rows)]
        )
        grid = st.data_editor(default_rows, num_rows="dynamic", use_container_width=True, key="manual_grid")

        col_btn1, col_btn2 = st.columns([1,1])
        with col_btn1:
            if st.button("✅ Conferma trasferte", use_container_width=True):
                rows = grid.replace({pd.NA:"", None:""}).to_dict("records")
                rows = [r for r in rows if any(str(v).strip() for v in r.values())]
                if len(rows) > 30:
                    rows = rows[:30]
                    st.info("Sono state prese solo le prime 30 righe.")
                bad = [r for r in rows if not (str(r.get("Matricola")).strip()
                                               and str(r.get("Cognome e Nome")).strip()
                                               and str(r.get("Turno")).strip())]
                if bad:
                    st.error("Compila Matricola, Cognome e Nome e Turno per ogni riga non vuota.")
                else:
                    extra_df = _build_extra_df(rows)
                    st.session_state["extra_rows"] = pd.concat([st.session_state["extra_rows"], extra_df], ignore_index=True)
                    st.session_state["df_view"] = pd.concat([st.session_state["df_view"], extra_df], ignore_index=True)
                    st.session_state["show_transfer_ui"] = False
                    st.success(f"Inserite {len(extra_df)} trasferte.")
                    _do_rerun()

        with col_btn2:
            if not st.session_state["extra_rows"].empty:
                if st.button("🗑️ Svuota trasferte aggiunte", use_container_width=True):
                    st.session_state["extra_rows"] = st.session_state["extra_rows"].iloc[0:0]
                    base = st.session_state["df_view"]
                    if "_added" in base.columns:
                        base = base[~base["_added"].fillna(False)].copy()
                    st.session_state["df_view"] = base
                    st.info("Trasferte cancellate.")
                    _do_rerun()

    # --- Export Excel ---
    xls_buf = io.BytesIO()
    with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
        cols = [c for c in DISPLAY_ORDER if c in st.session_state["df_view"].columns]
        st.session_state["df_view"][cols].to_excel(
            writer, sheet_name="ServizioGiornaliero", index=False
        )
    st.download_button(
        "⬇️ Scarica Excel",
        data=xls_buf.getvalue(),
        file_name="ServizioGiornaliero.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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
            exported_at=datetime.now() + timedelta(hours=2)
        )
        st.download_button(
            "⬇️ Scarica PDF",
            data=pdf_path.read_bytes(),
            file_name=pdf_path.name,
            mime="application/pdf"
        )
