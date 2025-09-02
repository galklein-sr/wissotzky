import pandas as pd
from typing import Tuple

# דפוס ללא קבוצות לוכדות, כדי למנוע UserWarning
PATTERN = r'(?:סה\"כ|סהכ|Total|סיכום)'

def remove_summary_rows(df: pd.DataFrame, col: str = "קוד סוכן") -> Tuple[pd.DataFrame, int]:
    if col not in df.columns:
        return df.copy(), 0
    s = df[col].astype(str).fillna("").str.strip()
    mask = s.str.contains(PATTERN, case=False, regex=True)
    removed = int(mask.sum())
    out = df.loc[~mask].reset_index(drop=True)
    return out, removed
