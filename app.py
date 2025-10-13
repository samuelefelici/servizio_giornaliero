import os, sys, re, io, tempfile
from pathlib import Path
from html import escape as html_escape
from datetime import datetime, timedelta

# --- Path per import locali ---
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd

REST_CODES = {"R", "RR"}   # riposo

# Eccezioni GLOBALI
GLOBAL_EXC_PATTERNS = (r"^IAST$", r"^N$",)

# Prefissi eccezioni ANCONA
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
.serv-table-wrap {width: 100%;}
.serv-table-wrap table {width: 100% !important; table-layout: fixed; border-collapse: collapse;}
.serv-table-wrap th {background:#f5f5f5; text-align:left; padding:6px;}
.serv-table-wrap td {padding:6px; vertical-align: top;}
.serv-table-wrap th:nth-child(4), .serv-table-wrap td:nth-child(4),
.serv-table-wrap th:nth-child(5), .serv-table-wrap td:nth-child(5) {text-align:center;}
.serv-table-wrap th:nth-child(6), .serv-table-wrap td:nth-child(6) {
  width: 40%; white-space: normal; word-break: break-word;
}
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

def _reorder_for_display(df: pd.DataFrame) -> pd.DataFrame:
    df2  = df.drop(columns=["Residenza"], errors="ignore").copy()
    cols = [c for c in DISPLAY_ORDER if c in df2.columns]
    return df2[cols] if cols else df2

def _res_to_prefix(res: str) -> str | None:
    if not isinstance(res, str): return None
    r = res.upper().replace("_", " ").strip()
    if "JESI URBANO" in r or r == "JU": return "JU"
    if "JESI" in r or r == "J":        return "J"
    if "MARINA" in r or r == "M":      return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r == "C": return "C"
    if "OSIMO" in r or r == "O":       return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r == "F":    return "F"
    if "POLVERIGI" in r or r == "P":   return "P"
    if "OSTRA" in r or r == "D":       return "D"
    if "BELVED" in r or r == "B" or "DEPBELVE" in r:     return "B"
    if "ANCONA" in r or r == "A":      return "A"
    return None

def _norm_turno(s) -> str:
    return str(s or "").upper().strip().replace(".", "").replace(" ", "")

def _match_any(token: str, patterns: tuple[str, ...]) -> bool:
    t = _norm_turno(token)
    return any(re.match(p, t) for p in patterns)

def _turno_bucket(turno: str) -> str | None:
    t = _norm_turno(turno)
    if not t or t in {"ASSENTE", *REST_CODES}: return None
    if t.startswith("JU"): return "JU"
    if t.startswith("J"):  return "J"
    return t[0]

def _accepted_prefixes_for_res(prefix: str | None) -> set[str]:
    if prefix in {"J", "JU"}: return {"J", "JU"}
    return {prefix} if prefix else set()

def _should_highlight_turno(residenza, turno) -> bool:
    t = _norm_turno(turno)
    if not t or t in {"ASSENTE", *REST_CODES}: return False
    if _match_any(t, GLOBAL_EXC_PATTERNS): return False
    rp = _res_to_prefix(residenza)
    if rp == "A":
        if _match_any(t, EXC_ANCONA_PATTERNS): return False
    if t.startswith("JU"): b = "JU"
    elif t.startswith("J"): b = "J"
    else: b = t[0] if t else None
    if not rp or not b: return False
    return b not in _accepted_prefixes_for_res(rp)

def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)
    return df.apply(lambda r: _should_highlight_turno(r.get("Residenza"), r.get("Turno")), axis=1)

# ---------- supporto a trasferte inserite ----------
_PREFIX_TO_RES = {
    "JU": "JESI URBANO", "J" : "JESI EXTRAURBANO", "M" : "MARINA",
    "C" : "CASTELFIDARDO","O" : "OSIMO","F" : "FILOTTRANO","P" : "POLVERIGI",
    "D" : "OSTRA","B" : "BELVEDERE","A" : "ANCONA",
}

def _turno_to_residenza_name(turno: str) -> str | None:
    t = _norm_turno(turno)
    if not t: return None
    if t.startswith("JU"): return _PREFIX_TO_RES["JU"]
    if t.startswith("J"):  return _PREFIX_TO_RES["J"]
    return _PREFIX_TO_RES.get(t[0])

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

def _styled_html_table(g_disp: pd.DataFrame, trasferta_mask: pd.Series, added_mask: pd.Series | None = None) -> str:
    """
    - Bold su Turno/Inizio/Fine per trasferte
    - Righe aggiunte in blu
    - * nelle note = a capo
    - Freccia ↳ in colonna 'Cognome e Nome' allineata a destra
    """
    df_html = g_disp.copy()
    note_col = "Indennità e note"
    if note_col in df_html.columns:
        df_html[note_col] = (
            df_html[note_col].astype(str).apply(lambda x: html_escape(x).replace("*", "<br/>"))
        )

    subset_cols = [c for c in ["Turno", "Inizio", "Fine"] if c in df_html.columns]
    if added_mask is None:
        added_mask = pd.Series(False, index=df_html.index)

    def _bold_if_trasferta(row):
        return (["font-weight: bold"] * len(row)) if trasferta_mask.loc[row.name] else ([""] * len(row))

    def _blue_if_added(row):
        return (["color:#0b5ed7"] * len(row)) if added_mask.loc[row.name] else ([""] * len(row))

    # stile cella: se la colonna è "Cognome e Nome" e il valore è ↳, allinea a destra
    def _right_if_arrow(val):
        return "text-align: right;" if str(val).strip() == "↳" else ""

    sty = (df_html.style
           .hide(axis="index")
           .apply(_blue_if_added, axis=1)
           .apply(_bold_if_trasferta, axis=1, subset=subset_cols)
           .applymap(_right_if_arrow, subset=["Cognome e Nome"] if "Cognome e Nome" in df_html.columns else None)
           .set_table_styles([
               {"selector": "table", "props": [("width","100%"), ("table-layout","fixed"), ("border-collapse","collapse")]},
               {"selector": "th",    "props": [("background","#f5f5f5"), ("text-align","left"), ("padding","6px")]},
               {"selector": "td",    "props": [("padding","6px"), ("vertical-align","top")]},
               {"selector": "th:nth-child(4), td:nth-child(4), th:nth-child(5), td:nth-child(5)", "props": [("text-align","center")]},
               {"selector": "th:nth-child(6), td:nth-child(6)", "props": [("width","40%"), ("white-space","normal"), ("word-break","break-word")]}
           ], overwrite=False)
          )
    return f'<div class="serv-table-wrap">{sty.to_html(escape=False)}</div>'


def render_preview(df_view: pd.DataFrame, meta: dict, inner_sort_choice: str):
    st.success(f"File elaborato. Data: {meta.get('data','?')} – {meta.get('giorno','?')} (fonte: {meta.get('origine','?')})")
    st.markdown(
        f"""
        <h2 style="color:#d00; font-weight:800; margin: 0.5rem 0 0.5rem 0; text-align:center;">
          Servizio Giornaliero: {meta.get('giorno','')} {meta.get('data','')}
        </h2>
        """,
        unsafe_allow_html=True,
    )
    st.subheader("Anteprima per deposito")

    if "Residenza" in df_view.columns:
        res_list = sorted(df_view["Residenza"].dropna().astype(str).unique())
        for res in res_list:
            g = df_view[df_view["Residenza"].astype(str) == res].copy()
            if g.empty: continue
            if inner_sort_choice.startswith("Inizio") and "Inizio" in g.columns:
                by = ["Inizio"] + (["Cognome e Nome"] if "Cognome e Nome" in g.columns else [])
                g = g.sort_values(by=by).reset_index(drop=True)
            else:
                by = []
                if "Cognome e Nome" in g.columns: by.append("Cognome e Nome")
                if "Inizio" in g.columns:         by.append("Inizio")
                if by: g = g.sort_values(by=by).reset_index(drop=True)
                # --- qui: compattazione + simbolo ↳ sotto il nome per turni successivi
                if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
                    same_person = g[["Cognome e Nome", "Matricola"]].eq(
                        g[["Cognome e Nome", "Matricola"]].shift(1)
                    ).all(axis=1)
                    g.loc[same_person, "Cognome e Nome"] = "↳"
                    g.loc[same_person, "Matricola"] = ""

            st.markdown(f"<h3 style='text-align:center; margin: 0.5rem 0 0.25rem 0;'>{res}</h3>", unsafe_allow_html=True)
            g_disp = _reorder_for_display(g)
            mask   = _trasferta_mask(g)
            added_mask = g["_added"].fillna(False) if "_added" in g.columns else pd.Series(False, index=g.index)
            html   = _styled_html_table(g_disp, mask, added_mask)
            st.markdown(html, unsafe_allow_html=True)
    else:
        st.markdown("### **TUTTI**")
        g = df_view.copy()
        if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
            same_person = g[["Cognome e Nome", "Matricola"]].eq(g[["Cognome e Nome", "Matricola"]].shift(1)).all(axis=1)
            g.loc[same_person, "Cognome e Nome"] = "↳"
            g.loc[same_person, "Matricola"] = ""
        g_disp = _reorder_for_display(g)
        mask   = _trasferta_mask(g)
        added_mask = g["_added"].fillna(False) if "_added" in g.columns else pd.Series(False, index=g.index)
        html   = _styled_html_table(g_disp, mask, added_mask)
        st.markdown(html, unsafe_allow_html=True)

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

        # Aggiungi eventuali trasferte inserite
        if not st.session_state["extra_rows"].empty:
            df_view = pd.concat([df_view, st.session_state["extra_rows"]], ignore_index=True)
        # >>>> NON APPLICARE OFFSET SUGLI ORARI <<<<
        # df_view = _apply_time_offset(df_view, hours=2)

        # salva stato
        st.session_state["df_view"] = df_view
        st.session_state["meta"] = meta
        st.session_state["last_inner_sort"] = "inizio" if inner_sort_choice.startswith("Inizio") else "nome"

    except Exception as e:
        import traceback
        st.error(f"Errore durante l'elaborazione: {e}")
        st.code("".join(traceback.format_exception(*sys.exc_info())))

# ====================== Trasferte + Preview + Export ======================

if st.session_state["df_view"] is not None and st.session_state["meta"] is not None:
    c1, c2 = st.columns([1,3])
    with c1:
        if st.button("➕ Inserisci trasferte", use_container_width=True):
            st.session_state["show_transfer_ui"] = not st.session_state["show_transfer_ui"]

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
                    # non applichiamo offset alle manuali
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

    # ---- Anteprima
    render_preview(st.session_state["df_view"], st.session_state["meta"], inner_sort_choice)

    # --- Export Excel ---
    xls_buf = io.BytesIO()
    with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
        _reorder_for_display(st.session_state["df_view"]).to_excel(
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
        pdf_path   = Path(td) / "ServizioGiornaliero.pdf"
        inner_sort = "inizio" if inner_sort_choice.startswith("Inizio") else "nome"
        build_pdf(
            pdf_path,
            st.session_state["df_view"],
            st.session_state["meta"],
            logo_path if logo_path.exists() else None,
            title=TITLE, inner_sort=inner_sort,
            exported_at=datetime.now() + timedelta(hours=2)  # <<<<< OFFSET SOLO SULLA DATA DI EXPORT
        )
        st.download_button(
            "⬇️ Scarica PDF",
            data=pdf_path.read_bytes(),
            file_name="ServizioGiornaliero.pdf",
            mime="application/pdf"
        )
