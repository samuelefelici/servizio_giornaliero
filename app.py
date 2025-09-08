# app.py
import os, sys, re, io, tempfile
from pathlib import Path
from html import escape as html_escape
from datetime import datetime

# --- Path per import locali ---
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd

REST_CODES = {"R", "RR"}   # riposo

# Eccezioni GLOBALI: valgono per tutti i depositi (match ESATTO)
GLOBAL_EXC_PATTERNS = (
    r"^IAST$",   # esattamente IAST
    r"^N$",      # esattamente N
)

# >>>>>>>>>>>>>>>>> ECCEZIONI E PREFISSI <<<<<<<<<<<<<<<<<
EXC_ANCONA_PREFIXES = (
    "D1R1","D1R2","D1R5","D2R1","D2R2","D2R3","D2R6",
    "NP","ASC","V5",
    "LU","MA","ME","GI","VE","SA","DO",
)
# prefissi -> regex "inizia con"
EXC_ANCONA_PATTERNS  = tuple(rf"^{re.escape(p)}" for p in EXC_ANCONA_PREFIXES)

# NON usare 'N' come prefisso per gli altri depositi (sennò cattura anche NP).
EXC_OTHER_PATTERNS = tuple()  # lasciato vuoto: IAST/N sono già coperti da GLOBAL_EXC_PATTERNS
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

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
assets_dir= Path("assest")   # nome cartella come richiesto
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

def _norm_turno(s) -> str:
    """Normalizza il codice turno per match robusti (maiuscolo, senza punti/spazi)."""
    return str(s or "").upper().strip().replace(".", "").replace(" ", "")

def _match_any(token: str, patterns: tuple[str, ...]) -> bool:
    """True se 'token' normalizzato soddisfa uno dei pattern regex."""
    t = _norm_turno(token)
    return any(re.match(p, t) for p in patterns)

def _turno_bucket(turno: str) -> str | None:
    """
    Bucket turno: 'JU' se parte con JU, altrimenti 'J' se parte con J,
    altrimenti prima lettera; esclude Assente/R/RR.
    """
    t = _norm_turno(turno)
    if not t or t in {"ASSENTE", *REST_CODES}:
        return None
    if t.startswith("JU"): return "JU"
    if t.startswith("J"):  return "J"
    return t[0]

def _accepted_prefixes_for_res(prefix: str | None) -> set[str]:
    """Prefissi considerati 'di casa' per la residenza (J e JU sono equivalenti)."""
    if prefix in {"J", "JU"}:
        return {"J", "JU"}
    return {prefix} if prefix else set()

def _should_highlight_turno(residenza, turno) -> bool:
    t = _norm_turno(turno)
    if not t or t in {"ASSENTE", *REST_CODES}:
        return False

    # Eccezioni GLOBALI (match esatto)
    if _match_any(t, GLOBAL_EXC_PATTERNS):
        return False

    rp = _res_to_prefix(residenza)

    # Eccezioni specifiche per deposito
    if rp == "A":  # ANCONA
        if _match_any(t, EXC_ANCONA_PATTERNS):
            return False
    else:
        # niente: IAST/N già catturati come GLOBAL (evitiamo di escludere NP ecc.)
        pass

    # Regola standard J/JU equivalenti, altrimenti prima lettera
    if t.startswith("JU"):
        b = "JU"
    elif t.startswith("J"):
        b = "J"
    else:
        b = t[0] if t else None

    if not rp or not b:
        return False
    return b not in _accepted_prefixes_for_res(rp)

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    """
    True se il turno è ‘fuori deposito’ (J/JU equivalenti; R/RR/Assente esclusi)
    e NON cade nelle liste di eccezione definite.
    """
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)
    return df.apply(lambda r: _should_highlight_turno(r.get("Residenza"), r.get("Turno")), axis=1)

# ---------- NUOVO: supporto a trasferte inserite ----------
_PREFIX_TO_RES = {
    "JU": "JESI URBANO",
    "J" : "JESI EXTRAURBANO",
    "M" : "MARINA",
    "C" : "CASTELFIDARDO",
    "O" : "OSIMO",
    "F" : "FILOTTRANO",
    "P" : "POLVERIGI",
    "D" : "OSTRA",
    "B" : "BELVEDERE",
    "A" : "ANCONA",
}

def _turno_to_residenza_name(turno: str) -> str | None:
    """Dato un codice turno, decide in quale 'Residenza' inserirlo (bucket)."""
    t = _norm_turno(turno)
    if not t:
        return None
    if t.startswith("JU"):
        return _PREFIX_TO_RES["JU"]
    if t.startswith("J"):
        return _PREFIX_TO_RES["J"]
    return _PREFIX_TO_RES.get(t[0])

def _parse_time_like(s: str | None) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    m = re.match(r"^([01]?\d|2[0-3]):([0-5]\d)$", s)
    return s if m else ""

def _build_extra_df(rows: list[dict]) -> pd.DataFrame:
    """Costruisce il DataFrame dalle righe manuali/importate e aggiunge flag _added."""
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).rename(columns={
        "Nominativo": "Cognome e Nome",
        "Turno fuori residenza": "Turno",
    })
    for c in ["Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Residenza"]:
        if c not in df.columns:
            df[c] = ""
    df["Residenza"] = df.apply(lambda r: r["Residenza"] or (_turno_to_residenza_name(r["Turno"]) or ""), axis=1)
    df["Inizio"] = df["Inizio"].map(_parse_time_like)
    df["Fine"]   = df["Fine"].map(_parse_time_like)
    df["_added"] = True
    keep = ["Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Residenza","_added"]
    return df[keep]

def _styled_html_table(g_disp: pd.DataFrame, trasferta_mask: pd.Series, added_mask: pd.Series | None = None) -> str:
    """
    Ritorna HTML della tabella (pandas Styler) a larghezza piena.
    - Bold su Turno/Inizio/Fine quando trasferta_mask è True.
    - Righe aggiunte colorate in BLU.
    - Nelle Note l'asterisco * forza un a capo.
    """
    df_html = g_disp.copy()
    note_col = "Indennità e note"
    if note_col in df_html.columns:
        df_html[note_col] = (
            df_html[note_col]
            .astype(str)
            .apply(lambda x: html_escape(x).replace("*", "<br/>"))
        )

    subset_cols = [c for c in ["Turno", "Inizio", "Fine"] if c in df_html.columns]

    def _bold_if_trasferta(row):
        return (["font-weight: bold"] * len(row)) if trasferta_mask.loc[row.name] else ([""] * len(row))

    if added_mask is None:
        added_mask = pd.Series(False, index=df_html.index)

    def _blue_if_added(row):
        return (["color: #0b5ed7"] * len(row)) if added_mask.loc[row.name] else ([""] * len(row))

    sty = (df_html.style
           .hide(axis="index")
           .apply(_blue_if_added, axis=1)  # colore blu righe aggiunte
           .apply(_bold_if_trasferta, axis=1, subset=subset_cols)
           .set_table_styles([
               {"selector": "table", "props": [("width","100%"), ("table-layout","fixed"), ("border-collapse","collapse")]},
               {"selector": "th",    "props": [("background","#f5f5f5"), ("text-align","left"), ("padding","6px")]},
               {"selector": "td",    "props": [("padding","6px"), ("vertical-align","top")]},
               {"selector": "th:nth-child(4), td:nth-child(4), th:nth-child(5), td:nth-child(5)",
                "props": [("text-align","center")]},
               {"selector": "th:nth-child(6), td:nth-child(6)",
                "props": [("width","40%"), ("white-space","normal"), ("word-break","break-word")]}
           ], overwrite=False)
          )
    return f'<div class="serv-table-wrap">{sty.to_html(escape=False)}</div>'


# ====================== UI ======================

# stato per trasferte aggiunte
if "extra_rows" not in st.session_state:
    st.session_state["extra_rows"] = pd.DataFrame(columns=[
        "Matricola","Cognome e Nome","Turno","Inizio","Fine","Indennità e note","Residenza","_added"
    ])

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

        # --- Inserisci trasferte (manuale/import) ---
        st.divider()
        with st.expander("➕ Inserisci trasferte", expanded=False):
            mode = st.radio("Modalità", ["Manualmente", "Importa"], horizontal=True)
            if mode == "Manualmente":
                with st.form("trasferte_manual"):
                    nrows = st.number_input("Quante righe vuoi inserire?", min_value=1, max_value=30, value=1, step=1)
                    manual_rows = []
                    for i in range(int(nrows)):
                        st.markdown(f"**Riga {i+1}**")
                        c1, c2, c3, c4, c5 = st.columns([1,2,1,1,1])
                        matricola = c1.text_input("Matricola*", key=f"mat_{i}")
                        nome      = c2.text_input("Cognome e Nome*", key=f"nom_{i}")
                        turno     = c3.text_input("Turno*", key=f"tur_{i}")
                        inizio    = c4.text_input("Inizio (HH:MM)", key=f"ini_{i}", placeholder="es. 06:30")
                        fine      = c5.text_input("Fine (HH:MM)",   key=f"fin_{i}", placeholder="es. 13:10")
                        if matricola or nome or turno or inizio or fine:
                            manual_rows.append({
                                "Matricola": str(matricola).strip(),
                                "Cognome e Nome": nome.strip(),
                                "Turno": turno.strip(),
                                "Inizio": inizio.strip(),
                                "Fine":   fine.strip(),
                                "Indennità e note": "",
                                "Residenza": _turno_to_residenza_name(turno.strip()) or "",
                            })
                    ok = st.form_submit_button("Conferma inserimento")
                    if ok:
                        errs = [r for r in manual_rows if not (r["Matricola"] and r["Cognome e Nome"] and r["Turno"])]
                        if errs:
                            st.error("Compila Matricola, Cognome e Nome e Turno per tutte le righe non vuote.")
                        else:
                            extra_df = _build_extra_df(manual_rows)
                            st.session_state["extra_rows"] = pd.concat(
                                [st.session_state["extra_rows"], extra_df], ignore_index=True
                            )
                            st.success(f"Inserite {len(extra_df)} trasferte.")
            else:
                up_transfer = st.file_uploader("Carica file trasferte (.xls/.xlsx)", type=["xls","xlsx"], key="upl_trs")
                if up_transfer is not None:
                    try:
                        raw = pd.read_excel(up_transfer)
                        renamed = raw.rename(columns={
                            "Matricola": "Matricola",
                            "Nominativo": "Cognome e Nome",
                            "Turno fuori residenza": "Turno",
                        })
                        rows = renamed.to_dict("records")
                        extra_df = _build_extra_df(rows)
                        st.session_state["extra_rows"] = pd.concat(
                            [st.session_state["extra_rows"], extra_df], ignore_index=True
                        )
                        st.success(f"Importate {len(extra_df)} trasferte.")
                    except Exception as e:
                        st.error(f"Errore in import: {e}")

            if not st.session_state["extra_rows"].empty:
                if st.button("🗑️ Svuota trasferte aggiunte"):
                    st.session_state["extra_rows"] = st.session_state["extra_rows"].iloc[0:0]
                    st.info("Trasferte cancellate.")
        st.divider()

        # Aggiungi eventuali trasferte inserite/importate
        if not st.session_state["extra_rows"].empty:
            df_view = pd.concat([df_view, st.session_state["extra_rows"]], ignore_index=True)

        # Feedback meta
        st.success(f"File elaborato. Data: {meta.get('data','?')} – {meta.get('giorno','?')} "
                   f"(fonte: {meta.get('origine','?')})")
        # Titolo rosso dinamico (anteprima)
        st.markdown(
            f"""
            <h2 style="color:#d00; font-weight:800; margin: 0.5rem 0 0.5rem 0; text-align:center;">
              Servizio Giornaliero: {meta.get('giorno','')} {meta.get('data','')}
            </h2>
            """,
            unsafe_allow_html=True,
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

                st.markdown(
                    f"<h3 style='text-align:center; margin: 0.5rem 0 0.25rem 0;'>{res}</h3>",
                    unsafe_allow_html=True,
                )

                g_disp = _reorder_for_display(g)
                mask   = _trasferta_mask(g)     # calcolata sull’indice originale
                added_mask = g["_added"].fillna(False) if "_added" in g.columns else pd.Series(False, index=g.index)
                html   = _styled_html_table(g_disp, mask, added_mask)
                st.markdown(html, unsafe_allow_html=True)

        else:
            st.markdown("### **TUTTI**")
            g_disp = _reorder_for_display(df_view)
            mask   = _trasferta_mask(df_view)
            added_mask = df_view["_added"].fillna(False) if "_added" in df_view.columns else pd.Series(False, index=df_view.index)
            html   = _styled_html_table(g_disp, mask, added_mask)
            st.markdown(html, unsafe_allow_html=True)

        # --- Export Excel ---
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
                      title=TITLE, inner_sort=inner_sort,
                      exported_at=datetime.now()
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
