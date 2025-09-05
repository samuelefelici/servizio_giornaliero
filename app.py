# app.py
import os, sys, re, io, tempfile
from pathlib import Path

# --- Path per import locali ---
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd

# --- Import moduli del progetto (con gestione errore) ---
try:
    from src.process import read_input_excel, transform_dataframe, debug_probe  # debug_probe opzionale
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
    if not isinstance(res, str):
        return None
    r = res.upper().replace("_", " ")
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

def _turno_bucket(turno: str) -> str | None:
    """
    Riduce il codice turno a un 'bucket' di confronto:
      - 'JU…' -> 'JU'
      - 'J…' (non JU) -> 'J'
      - altrimenti prima lettera (M, C, O, F, P, D, A, B…)
    """
    if not isinstance(turno, str):
        turno = str(turno)
    s = turno.strip()
    if not s or s.upper() == "ASSENTE":
        return None
    m = re.match(r"[A-Za-z]+", s)
    if not m:
        return None
    up = m.group(0).upper()
    if up.startswith("JU"):
        return "JU"
    if up.startswith("J"):
        return "J"
    return up[0]  # prima lettera per gli altri depositi

def _accepted_prefixes_for_res(prefix: str | None) -> set[str]:
    """Prefissi considerati 'di casa' per la residenza (J e JU sono equivalenti)."""
    if prefix in ("J", "JU"):
        return {"J", "JU"}
    return {prefix} if prefix else set()

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    """
    True se la riga è 'trasferta' (Turno non appartiene ai prefissi accettati per la residenza).
    Esclude Assente e righe senza info utile.
    """
    if not {"Residenza", "Turno"}.issubset(df.columns):
        return pd.Series(False, index=df.index)

    def _is_trasferta(row) -> bool:
        rp = _res_to_prefix(row["Residenza"])
        tb = _turno_bucket(row["Turno"])
        if rp is None or tb is None:
            return False
        accepted = _accepted_prefixes_for_res(rp)
        return tb not in accepted

    return df.apply(_is_trasferta, axis=1)

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

    # Debug opzionale (si esegue su stream; la nostra lettura fa seek(0) quindi va bene)
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
                    # Compattazione: non ripetere Nome/Matricola su righe consecutive uguali
                    if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
                        same_person = g[["Cognome e Nome", "Matricola"]].eq(
                            g[["Cognome e Nome", "Matricola"]].shift(1)
                        ).all(axis=1)
                        g.loc[same_person, ["Cognome e Nome", "Matricola"]] = ""

                st.markdown(f"### **{res}**")

                # Reorder/drop per visualizzazione
                g_disp = _reorder_for_display(g)

                # Bold per trasferte: calcolo sulla tabella originale 'g' (indici allineati)
                mask = _trasferta_mask(g)
                subset_cols = [c for c in ["Turno", "Inizio", "Fine"] if c in g_disp.columns]

                if subset_cols:
                    def _bold_if_trasferta(row):
                        return (["font-weight: bold"] * len(row)) if mask.loc[row.name] else ([""] * len(row))

                    styled = g_disp.style.apply(_bold_if_trasferta, axis=1, subset=subset_cols)
                    st.dataframe(styled, use_container_width=True, hide_index=True)
                else:
                    st.dataframe(g_disp, use_container_width=True, hide_index=True)
        else:
            # Nessuna Residenza: mostra tabella piatta
            st.dataframe(_reorder_for_display(df_view), use_container_width=True, hide_index=True)

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

        # --- Export PDF (stessa logica, il grassetto per trasferte è gestito in pdf_export.py) ---
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
