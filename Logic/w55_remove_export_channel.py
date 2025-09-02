import pandas as pd
from typing import Tuple

def remove_export_channel(df: pd.DataFrame, col: str = "ערוץ") -> Tuple[pd.DataFrame, int]:
    if col not in df.columns:
        return df.copy(), 0
    s = df[col].astype(str).fillna("").str.strip()
    mask = s.eq("ייצוא")
    removed = int(mask.sum())
    out = df.loc[~mask].reset_index(drop=True)
    return out, removed