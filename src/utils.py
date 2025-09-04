from typing import Tuple, Optional
import pandas as pd

def find_header_row(df_raw: pd.DataFrame, header_probe: str) -> Optional[int]:
    for i in range(min(30, len(df_raw))):  # cerca nei primi 30
        row_vals = df_raw.iloc[i].astype(str).tolist()
        if any(str(v).strip() == header_probe for v in row_vals):
            return i
    return None

def coerce_time(series: pd.Series) -> pd.Series:
    # converte “HH:MM” in timedelta o stringa vuota
    s = series.astype(str).str.strip()
    s = s.replace({"00:00": ""})
    # prova a parse solo non vuoti
    out = pd.to_datetime(s, format="%H:%M", errors="coerce").dt.strftime("%H:%M")
    out = out.fillna(s.where(s==""))  # preserva stringhe già vuote/altre sigle
    return out

def clean_spaces(df: pd.DataFrame) -> pd.DataFrame:
    return df.applymap(lambda x: str(x).strip() if pd.notna(x) else x)
