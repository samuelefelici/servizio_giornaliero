from typing import Optional
import pandas as pd

# --- normalizzatori di testo ---

def _norm_text(x) -> str:
    """
    Converte qualunque valore in stringa SOLO per confronto, sostituisce NBSP con spazio
    e fa strip. Usala per confronti (header, probe), non per scrivere nel df.
    """
    if x is None:
        return ""
    return str(x).replace("\u00A0", " ").strip()

def _strip_cell(x):
    """
    Rimuove NBSP e spazi ai margini SOLO se x è stringa.
    Non tocca numeri, date, NaN.
    """
    if isinstance(x, str):
        return x.replace("\u00A0", " ").strip()
    return x

# --- API pubbliche ---

def find_header_row(df_raw: pd.DataFrame, header_probe: str) -> Optional[int]:
    """
    Cerca la riga di intestazione confrontando in modo normalizzato (trim + NBSP fix).
    """
    probe = _norm_text(header_probe)
    # cerchiamo un po' più a fondo (alcuni export hanno parecchie righe top)
    max_scan = min(60, len(df_raw))
    for i in range(max_scan):
        row_vals = df_raw.iloc[i].tolist()
        if any(_norm_text(v) == probe for v in row_vals):
            return i
    return None

def coerce_time(series: pd.Series) -> pd.Series:
    """
    Converte “HH:MM” in una stringa normalizzata “HH:MM”.
    Sostituisce '00:00' con vuoto.
    Non lancia eccezioni su valori non parsabili (li lascia com’erano).
    """
    # lavora su una copia, senza toccare il dtype originale della serie chiamante
    s = series.copy()

    # solo celle stringa: trim + NBSP fix
    s = s.apply(lambda x: x.replace("\u00A0", " ").strip() if isinstance(x, str) else x)

    # mappa '00:00' -> '' (solo se è stringa '00:00')
    s = s.apply(lambda x: "" if isinstance(x, str) and x == "00:00" else x)

    # prova il parsing in HH:MM per tutte le celle (stringhe o meno)
    # errors='coerce' mette NaT dove non parsabile
    parsed = pd.to_datetime(s, format="%H:%M", errors="coerce")

    # dove parsed è valido, formattiamo; altrimenti teniamo il valore originale
    out = s.copy()
    mask = parsed.notna()
    out.loc[mask] = parsed[mask].dt.strftime("%H:%M")
    return out

def clean_spaces(df: pd.DataFrame) -> pd.DataFrame:
    """
    Trim su tutte le celle stringa (NBSP compresi). Non converte numeri/date in stringhe.
    """
    return df.applymap(_strip_cell)

def clean_column_names(cols) -> list[str]:
    """
    Pulisce i nomi colonna: NBSP→spazio e strip.
    """
    return [_norm_text(c) for c in cols]
