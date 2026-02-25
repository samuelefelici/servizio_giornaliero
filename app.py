import os
import sys
import re
import io
import tempfile
from pathlib import Path
from html import escape as html_escape
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ============================================================
# Config & import locali
# ============================================================

# --- Path per import locali ---
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

REST_CODES = {"R", "RR"}  # riposo

# Eccezioni GLOBALI
GLOBAL_EXC_PATTERNS = (r"^IAST$", r"^N$",)

# Prefissi eccezioni ANCONA
EXC_ANCONA_PREFIXES = (
    "D1R1", "D1R2", "D1R5", "D2R1", "D2R2", "D2R3", "D2R6",
    "NP", "ASC", "V5", "LU", "MA", "ME", "GI", "VE", "SA", "DO",
    "STRA",
)
EXC_ANCONA_PATTERNS = tuple(rf"^{re.escape(p)}" for p in EXC_ANCONA_PREFIXES)

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

# ============================================================
# Stili pagina
# ============================================================

st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)

st.title("Servizio Giornaliero")
st.caption("Carica il file Excel, visualizza l’anteprima, modifica le note e gestisci le trasferte.")
st.divider()

cfg_dir = Path("config")
assets_dir = Path("assest")
logo_path = assets_dir / "logo.jpg"

# ============================================================
# Session state
# ============================================================

st.session_state.setdefault("df_view", None)
st.session_state.setdefault("meta", None)
st.session_state.setdefault("last_inner_sort", "nome")
st.session_state.setdefault("show_transfer_ui", False)
st.session_state.setdefault(
    "extra_rows",
    pd.DataFrame(
        columns=[
            "Matricola",
            "Cognome e Nome",
            "Turno",
            "Inizio",
            "Fine",
            "Indennità e note",
            "Residenza",
            "_added",
        ]
    ),
)
st.session_state.setdefault("transfer_mode", None)

# ============================================================
# Helper: colonne, evidenziazione e trasferte
# ============================================================


def _reorder_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Riordina le colonne per la visualizzazione secondo DISPLAY_ORDER."""
    df2 = df.drop(columns=["Residenza"], errors="ignore").copy()
    cols = [c for c in DISPLAY_ORDER if c in df2.columns]
    return df2[cols] if cols else df2


def _res_to_prefix(res: str) -> str | None:
    """Ricava il prefisso dalla stringa di residenza (JU/J/M/C/...)."""
    if not isinstance(res, str):
        return None
    r = res.upper().replace("_", " ").strip()
    if "JESI URBANO" in r or r == "JU":
        return "JU"
    if "JESI" in r or r == "J":
        return "J"
    if "MARINA" in r or r == "M":
        return "M"
    if "CASTELFIDARDO" in r or "C.FID" in r or r == "C":
        return "C"
    if "OSIMO" in r or r == "O":
        return "O"
    if "FILOTTRANO" in r or "FILOT" in r or r == "F":
        return "F"
    if "POLVERIGI" in r or r == "P":
        return "P"
    if "OSTRA" in r or r == "D":
        return "D"
    if "BELVED" in r or r == "B" or "DEPBELVE" in r:
        return "B"
    if "ANCONA" in r or r == "A":
        return "A"
    return None


def _norm_turno(value) -> str:
    """Normalizza il turno per confronti: maiuscolo, no spazi/punti."""
    return str(value or "").upper().strip().replace(".", "").replace(" ", "")


def _match_any(token: str, patterns: tuple[str, ...]) -> bool:
    """True se token matcha almeno uno dei regex patterns."""
    t = _norm_turno(token)
    return any(re.match(p, t) for p in patterns)


def _accepted_prefixes_for_res(prefix: str | None) -> set[str]:
    """Prefissi accettati per una residenza (J e JU sono compatibili tra loro)."""
    if prefix in {"J", "JU"}:
        return {"J", "JU"}
    return {prefix} if prefix else set()


def _should_highlight_turno(residenza, turno) -> bool:
    """
    True se il turno appare "fuori residenza" (trasferta automatica),
    escluse alcune eccezioni.
    """
    t = _norm_turno(turno)
    if not t or t in {"ASSENTE", *REST_CODES}:
        return False
    if _match_any(t, GLOBAL_EXC_PATTERNS):
        return False

    rp = _res_to_prefix(residenza)

    # Eccezioni Ancona: alcuni prefissi non vanno considerati trasferta
    if rp == "A":
        if _match_any(t, EXC_ANCONA_PATTERNS):
            return False

    # Bucket del turno (JU/J/iniziale)
    if t.startswith("JU"):
        bucket = "JU"
    elif t.startswith("J"):
        bucket = "J"
    else:
        bucket = t[0] if t else None

    if not rp or not bucket:
        return False

    return bucket not in _accepted_prefixes_for_res(rp)


def _trasferta_mask(df: pd.DataFrame) -> pd.Series:
    """Maschera booleana: True sulle righe da evidenziare come trasferta."""
    if "Residenza" not in df.columns or "Turno" not in df.columns:
        return pd.Series(False, index=df.index)
    return df.apply(
        lambda r: _should_highlight_turno(r.get("Residenza"), r.get("Turno")),
        axis=1,
    )


def _is_turno_numero(turno: str) -> bool:
    """True se il turno inizia con 3 cifre (es. 510, 520...)."""
    t = str(turno).strip()
    return bool(re.match(r"^\d{3}", t))


_PREFIX_TO_RES = {
    "JU": "JESI URBANO",
    "J": "JESI EXTRAURBANO",
    "M": "MARINA",
    "C": "CASTELFIDARDO",
    "O": "OSIMO",
    "F": "FILOTTRANO",
    "P": "POLVERIGI",
    "D": "OSTRA",
    "B": "BELVEDERE",
    "A": "ANCONA",
}


def _turno_to_residenza_name(turno: str) -> str | None:
    """Prova a ricavare il nome residenza dal turno (J/JU o iniziale)."""
    t = _norm_turno(turno)
    if not t:
        return None
    if t.startswith("JU"):
        return _PREFIX_TO_RES["JU"]
    if t.startswith("J"):
        return _PREFIX_TO_RES["J"]
    return _PREFIX_TO_RES.get(t[0])


def _parse_time_like(s: str | None) -> str:
    """Accetta orari in formato HH:MM, altrimenti ritorna stringa vuota."""
    s = (s or "").strip()
    if not s:
        return ""
    m = re.match(r"^([01]?\d|2[0-3]):([0-5]\d)$", s)
    return s if m else ""


def _build_extra_df(rows: list[dict]) -> pd.DataFrame:
    """Costruisce un DataFrame standard per le trasferte aggiunte manualmente/da file."""
    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows).rename(
        columns={
            "Nominativo": "Cognome e Nome",
            "Turno fuori residenza": "Turno",
        }
    )

    for c in [
        "Matricola",
        "Cognome e Nome",
        "Turno",
        "Inizio",
        "Fine",
        "Indennità e note",
        "Residenza",
    ]:
        if c not in df.columns:
            df[c] = ""

    df["Residenza"] = df.apply(
        lambda r: r["Residenza"] or (_turno_to_residenza_name(r["Turno"]) or ""),
        axis=1,
    )
    df["Inizio"] = df["Inizio"].map(_parse_time_like)
    df["Fine"] = df["Fine"].map(_parse_time_like)
    df["_added"] = True

    keep = [
        "Matricola",
        "Cognome e Nome",
        "Turno",
        "Inizio",
        "Fine",
        "Indennità e note",
        "Residenza",
        "_added",
    ]
    return df[keep]


def _styled_html_table(
    g_disp: pd.DataFrame,
    trasferta_mask: pd.Series,
    added_mask: pd.Series | None = None,
) -> str:
    """Ritorna HTML della tabella con gli stili (blu/grassetto) come in app."""
    df_html = g_disp.copy()
    note_col = "Indennità e note"

    if note_col in df_html.columns:
        df_html[note_col] = df_html[note_col].astype(str).apply(
            lambda x: html_escape(x).replace("*", "<br/>")
        )

    style_cols = [
        c for c in ["Matricola", "Cognome e Nome", "Turno", "Inizio", "Fine"]
        if c in df_html.columns
    ]

    if added_mask is None:
        added_mask = pd.Series(False, index=df_html.index)

    def _custom_style(row: pd.Series) -> list[str]:
        # Riga aggiunta manualmente → tutta blu+grassetto
        if added_mask.loc[row.name]:
            return ["color:#0b5ed7;font-weight:bold"] * len(row)

        # Turno numerico (510/520...) → tutta in grassetto
        if _is_turno_numero(row.get("Turno", "")):
            return ["font-weight:bold"] * len(row)

        # Trasferta automatica → blu+grassetto su alcune colonne
        if trasferta_mask.loc[row.name]:
            return [
                "color:#0b5ed7;font-weight:bold" if col in style_cols else ""
                for col in row.index
            ]

        return [""] * len(row)

    def _right_if_arrow(val) -> str:
        return "text-align: right;" if str(val).strip() == "↳" else ""

    sty = (
        df_html.style
        .hide(axis="index")
        .apply(_custom_style, axis=1)
        .applymap(
            _right_if_arrow,
            subset=["Cognome e Nome"] if "Cognome e Nome" in df_html.columns else None,
        )
        .set_table_styles(
            [
                {
                    "selector": "table",
                    "props": [
                        ("width", "100%"),
                        ("table-layout", "fixed"),
                        ("border-collapse", "collapse"),
                    ],
                },
                {"selector": "th", "props": [("background", "#f5f5f5"), ("text-align", "left"), ("padding", "6px")]},
                {"selector": "td", "props": [("padding", "6px"), ("vertical-align", "top")]},
                {
                    "selector": "th:nth-child(4), td:nth-child(4), th:nth-child(5), td:nth-child(5)",
                    "props": [("text-align", "center")],
                },
                {
                    "selector": "th:nth-child(6), td:nth-child(6)",
                    "props": [("width", "40%"), ("white-space", "normal"), ("word-break", "break-word")],
                },
            ],
            overwrite=False,
        )
    )
    return f'<div class="serv-table-wrap">{sty.to_html(escape=False)}</div>'


def render_preview(df_view: pd.DataFrame, meta: dict, inner_sort_choice: str) -> None:
    """Render anteprima per deposito/residenza."""
    st.success(
        f"File elaborato. Data: {meta.get('data','?')} – {meta.get('giorno','?')} "
        f"(fonte: {meta.get('origine','?')})"
    )

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
            if g.empty:
                continue

            if inner_sort_choice.startswith("Inizio") and "Inizio" in g.columns:
                by = ["Inizio"] + (["Cognome e Nome"] if "Cognome e Nome" in g.columns else [])
                g = g.sort_values(by=by).reset_index(drop=True)
            else:
                by = []
                if "Cognome e Nome" in g.columns:
                    by.append("Cognome e Nome")
                if "Inizio" in g.columns:
                    by.append("Inizio")

                if by:
                    g = g.sort_values(by=by).reset_index(drop=True)

                # Collassa ripetizioni: se stesso nominativo+matricola della riga sopra
                if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
                    same_person = g[["Cognome e Nome", "Matricola"]].eq(
                        g[["Cognome e Nome", "Matricola"]].shift(1)
                    ).all(axis=1)
                    g.loc[same_person, "Cognome e Nome"] = "↳"
                    g.loc[same_person, "Matricola"] = ""

            st.markdown(
                f"<h3 style='text-align:center; margin: 0.5rem 0 0.25rem 0;'>{res}</h3>",
                unsafe_allow_html=True,
            )

            g_disp = _reorder_for_display(g)
            mask = _trasferta_mask(g)
            added_mask = (
                g["_added"].fillna(False)
                if "_added" in g.columns
                else pd.Series(False, index=g.index)
            )

            html = _styled_html_table(g_disp, mask, added_mask)
            st.markdown(html, unsafe_allow_html=True)

    else:
        st.markdown("### **TUTTI**")
        g = df_view.copy()

        if {"Cognome e Nome", "Matricola"}.issubset(g.columns):
            same_person = g[["Cognome e Nome", "Matricola"]].eq(
                g[["Cognome e Nome", "Matricola"]].shift(1)
            ).all(axis=1)
            g.loc[same_person, "Cognome e Nome"] = "↳"
            g.loc[same_person, "Matricola"] = ""

        g_disp = _reorder_for_display(g)
        mask = _trasferta_mask(g)
        added_mask = (
            g["_added"].fillna(False)
            if "_added" in g.columns
            else pd.Series(False, index=g.index)
        )

        html = _styled_html_table(g_disp, mask, added_mask)
        st.markdown(html, unsafe_allow_html=True)


def _do_rerun() -> None:
    """Forza un rerun Streamlit (compatibile con versioni diverse)."""
    fn = getattr(st, "rerun", None)
    if callable(fn):
        fn()
    else:
        st.experimental_rerun()


# ============================================================
# UI: Caricamento file
# ============================================================

st.header("Carica il file")
uploaded = st.file_uploader("Trascina qui il file oppure selezionalo", type=["xls", "xlsx"])
if uploaded:
    st.success("File caricato correttamente.")

st.divider()

# ============================================================
# UI: Opzioni
# ============================================================

with st.expander("Opzioni di visualizzazione", expanded=True):
    col1, col2 = st.columns([2, 1])

    with col1:
        inner_sort_choice = st.radio(
            "Ordina ciascun deposito per:",
            ["Cognome e Nome (A→Z)", "Inizio (orario)"],
            horizontal=True,
            index=0,
        )

    with col2:
        show_absent = st.checkbox(
            "Mostra anche gli 'Assente'",
            value=True,
            help="Se deselezionato nasconde le righe con Turno = Assente (vale per anteprima ed export).",
        )

with st.expander("🧪 Debug"):
    debug_mode = st.checkbox("Attiva modalità debug", value=False)

st.divider()

# ============================================================
# Azione: ELABORA
# ============================================================

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
        df_out = transform_dataframe(df, cfg_dir)

        df_view = df_out.copy()

        # Nascondi Assente se richiesto
        if not show_absent and "Turno" in df_view.columns:
            before = len(df_view)
            df_view = df_view[df_view["Turno"].astype(str).str.strip() != "Assente"].reset_index(drop=True)
            hidden = before - len(df_view)
            if hidden > 0:
                st.info(f"Righe 'Assente' nascoste: {hidden}")

        # Applica eventuali trasferte già aggiunte
        if not st.session_state["extra_rows"].empty:
            df_view = pd.concat([df_view, st.session_state["extra_rows"]], ignore_index=True)

        st.session_state["df_view"] = df_view
        st.session_state["meta"] = meta
        st.session_state["last_inner_sort"] = "inizio" if inner_sort_choice.startswith("Inizio") else "nome"

    except Exception as e:
        import traceback

        st.error(f"Errore durante l'elaborazione: {e}")
        st.code("".join(traceback.format_exception(*sys.exc_info())))

# ============================================================
# Trasferte + Preview + Export
# ============================================================

if st.session_state["df_view"] is not None and st.session_state["meta"] is not None:
    st.header("Azioni rapide")
    st.divider()

    btn_style = dict(use_container_width=True)

    # ---------- Barra azioni principali ----------
    colA, colB, colC = st.columns([2, 2, 1])

    with colA:
        with st.popover("➕ Trasferte", use_container_width=True):
            st.caption("Aggiungi nuove trasferte manualmente o da file.")

            c1, c2 = st.columns(2)
            with c1:
                if st.button("📝 Manuale", **btn_style):
                    st.session_state["show_transfer_ui"] = True
                    st.session_state["transfer_mode"] = "manuale"
                    _do_rerun()
            with c2:
                if st.button("📄 Da File", **btn_style):
                    st.session_state["show_transfer_ui"] = True
                    st.session_state["transfer_mode"] = "file"
                    _do_rerun()

    with colB:
        if "Indennità e note" in st.session_state["df_view"].columns:
            with st.popover("✏️ Modifica Note", use_container_width=True):
                df_notes = st.session_state["df_view"].copy()

                search_value = st.text_input(
                    "🔎 Cerca Cognome e Nome",
                    placeholder="Digita il cognome...",
                    key="search_notes",
                )

                if search_value:
                    mask = df_notes["Cognome e Nome"].astype(str).str.contains(search_value, case=False, na=False)
                    df_notes = df_notes[mask].copy()

                cols_for_edit = [c for c in ["Matricola", "Cognome e Nome", "Indennità e note"] if c in df_notes.columns]
                editable = df_notes[cols_for_edit].copy()

                edited = st.data_editor(
                    editable,
                    use_container_width=True,
                    num_rows="fixed",
                    column_config={
                        "Indennità e note": st.column_config.TextColumn("Note", required=False),
                    },
                )

                if st.button("💾 Salva Note", **btn_style):
                    key_cols = [c for c in ["Matricola", "Cognome e Nome"] if c in st.session_state["df_view"].columns]

                    for _, row in edited.iterrows():
                        # Match riga base su Matricola + Cognome e Nome
                        base_mask = pd.Series(True, index=st.session_state["df_view"].index)
                        for col in key_cols:
                            base_mask &= (st.session_state["df_view"][col] == row[col])

                        note_value = row.get("Indennità e note", "")
                        if note_value is None or (isinstance(note_value, float) and pd.isna(note_value)):
                            note_value = ""

                        st.session_state["df_view"].loc[base_mask, "Indennità e note"] = note_value

                    st.success("Note aggiornate!")
                    _do_rerun()

    with colC:
        if not st.session_state["extra_rows"].empty:
            if st.button("🗑️ Svuota trasferte", **btn_style):
                st.session_state["extra_rows"] = st.session_state["extra_rows"].iloc[0:0]

                base = st.session_state["df_view"]
                if "_added" in base.columns:
                    base = base[~base["_added"].fillna(False)].copy()

                st.session_state["df_view"] = base
                st.info("Trasferte cancellate.")
                _do_rerun()

    # ---------- UI inserimento trasferte ----------
    if st.session_state.get("show_transfer_ui", False):
        st.divider()
        st.subheader("Inserisci trasferte")

        mode = st.session_state.get("transfer_mode")

        # Se non è impostata (caso limite), chiedi di scegliere
        if mode is None:
            cm, cf = st.columns(2)
            with cm:
                if st.button("Manuale", **btn_style):
                    st.session_state["transfer_mode"] = "manuale"
                    _do_rerun()
            with cf:
                if st.button("Da File", **btn_style):
                    st.session_state["transfer_mode"] = "file"
                    _do_rerun()

        elif mode == "manuale":
            st.caption("Compila le righe (max 30). Obbligatori: Matricola, Cognome e Nome, Turno.")

            start_rows = 5
            default_rows = pd.DataFrame(
                [{"Matricola": "", "Cognome e Nome": "", "Turno": "", "Inizio": "", "Fine": ""} for _ in range(start_rows)]
            )
            grid = st.data_editor(default_rows, num_rows="dynamic", use_container_width=True, key="manual_grid")

            b1, b2 = st.columns([2, 1])
            with b1:
                if st.button("✅ Conferma trasferte", **btn_style):
                    rows = grid.replace({pd.NA: "", None: ""}).to_dict("records")
                    rows = [r for r in rows if any(str(v).strip() for v in r.values())]

                    if len(rows) > 30:
                        rows = rows[:30]
                        st.info("Sono state prese solo le prime 30 righe.")

                    bad = [
                        r
                        for r in rows
                        if not (
                            str(r.get("Matricola", "")).strip()
                            and str(r.get("Cognome e Nome", "")).strip()
                            and str(r.get("Turno", "")).strip()
                        )
                    ]

                    if bad:
                        st.error("Compila Matricola, Cognome e Nome e Turno per ogni riga non vuota.")
                    else:
                        extra_df = _build_extra_df(rows)

                        st.session_state["extra_rows"] = pd.concat([st.session_state["extra_rows"], extra_df], ignore_index=True)
                        st.session_state["df_view"] = pd.concat([st.session_state["df_view"], extra_df], ignore_index=True)

                        st.session_state["show_transfer_ui"] = False
                        st.session_state["transfer_mode"] = None

                        st.success(f"Inserite {len(extra_df)} trasferte.")
                        _do_rerun()

            with b2:
                if st.button("Annulla", **btn_style):
                    st.session_state["show_transfer_ui"] = False
                    st.session_state["transfer_mode"] = None
                    _do_rerun()

        elif mode == "file":
            st.caption(
                "Importa trasferte da file XLS/XLSX. "
                "In questo flusso il tuo codice attuale legge in realtà un file tabellare (TSV/CSV)."
            )

            file_import = st.file_uploader("Scegli file trasferte", type=["xls", "xlsx"], key="upl_transfer_xls")

            preview_rows: list[dict] = []
            if file_import is not None:
                try:
                    # NB: lasciamo la tua logica originale (read_csv sep=\t)
                    df_import = pd.read_csv(file_import, sep="\t", header=0, encoding="cp1252")

                    for _, row in df_import.iterrows():
                        preview_rows.append(
                            {
                                "Matricola": row.get("Matricola", ""),
                                "Nominativo": row.get("Nominativo", ""),
                                "Turno fuori residenza": row.get("Turno fuori residenza", ""),
                                "Inizio": "",
                                "Fine": "",
                                "Indennità e note": "",
                                "Residenza": "",
                            }
                        )
                    st.success(f"Trasferte importate dal file: {len(preview_rows)}")

                except Exception as e:
                    st.error(f"Errore import: {e}")

            df_preview = pd.DataFrame(
                preview_rows
                or [{"Matricola": "", "Nominativo": "", "Turno fuori residenza": "", "Inizio": "", "Fine": ""}]
            )

            grid = st.data_editor(df_preview, num_rows="dynamic", use_container_width=True, key="file_grid")

            b1, b2 = st.columns([2, 1])

            with b1:
                if st.button("✅ Conferma trasferte importate", **btn_style):
                    rows = grid.replace({pd.NA: "", None: ""}).to_dict("records")
                    rows = [r for r in rows if any(str(v).strip() for v in r.values())]

                    bad = [
                        r
                        for r in rows
                        if not (
                            str(r.get("Matricola", "")).strip()
                            and str(r.get("Nominativo", "")).strip()
                            and str(r.get("Turno fuori residenza", "")).strip()
                        )
                    ]

                    if bad:
                        st.error("Compila Matricola, Cognome e Nome e Turno per ogni riga non vuota.")
                    else:
                        extra_df = _build_extra_df(rows)

                        st.session_state["extra_rows"] = pd.concat([st.session_state["extra_rows"], extra_df], ignore_index=True)
                        st.session_state["df_view"] = pd.concat([st.session_state["df_view"], extra_df], ignore_index=True)

                        st.session_state["show_transfer_ui"] = False
                        st.session_state["transfer_mode"] = None

                        st.success(f"Inserite {len(extra_df)} trasferte dal file.")
                        _do_rerun()

            with b2:
                if st.button("Annulla", **btn_style):
                    st.session_state["show_transfer_ui"] = False
                    st.session_state["transfer_mode"] = None
                    _do_rerun()

    st.divider()

    # ---------- Anteprima ----------
    render_preview(st.session_state["df_view"], st.session_state["meta"], inner_sort_choice)

    st.divider()

    # ============================================================
    # Export PDF (download + "apri PDF per stampare")
    # ============================================================

    st.subheader("Esporta in PDF")

    with tempfile.TemporaryDirectory() as td:
        temp_dir = Path(td)
        inner_sort = "inizio" if inner_sort_choice.startswith("Inizio") else "nome"

        pdf_path = build_pdf(
            temp_dir,
            st.session_state["df_view"],
            st.session_state["meta"],
            logo_path if logo_path.exists() else None,
            title=TITLE,
            inner_sort=inner_sort,
            exported_at=datetime.now(),
        )

        pdf_bytes = pdf_path.read_bytes()
        pdf_filename = pdf_path.name

        col_dl, col_open = st.columns([1, 1])

        with col_dl:
            st.download_button(
                "⬇️ Scarica PDF",
                data=pdf_bytes,
                file_name=pdf_filename,
                mime="application/pdf",
                use_container_width=True,
            )

        with col_open:
            # Chrome blocca spesso la stampa automatica da iframe/data-url.
            # Soluzione affidabile: apri il PDF nel viewer del browser e poi l'utente stampa (Ctrl+P).
            import base64

            b64 = base64.b64encode(pdf_bytes).decode("utf-8")
            pdf_data_url = f"data:application/pdf;base64,{b64}"

            st.markdown(
                f"""
                <a href="{pdf_data_url}" target="_blank" style="text-decoration:none;">
                  <button style="
                      width:100%;
                      padding:0.6rem 1rem;
                      border-radius:0.5rem;
                      border:1px solid rgba(49, 51, 63, 0.2);
                      background-color:white;
                      cursor:pointer;
                    ">
                    🖨️ Apri PDF (poi stampa con Ctrl+P)
                  </button>
                </a>
                """,
                unsafe_allow_html=True,
            )

else:
    st.info("Carica un file e premi ▶️ Elabora per abilitare anteprima ed export.")
