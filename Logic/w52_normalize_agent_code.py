import pandas as pd
from typing import Tuple

def normalize_agent_code(df: pd.DataFrame, col: str = "קוד סוכן") -> Tuple[pd.DataFrame, int]:
    """
    משאיר רק ספרות ב'קוד סוכן'; מחרוזות שמתרוקנות -> NaN.
    מחזיר גם כמות שינויים.
    """
    if col not in df.columns:
        return df.copy(), 0
    s = df[col].astype(str)
    new = s.str.replace(r"\D+", "", regex=True).str.strip()
    changes = int((s != new).sum())
    out = df.copy()
    out[col] = new.replace({"": None})
    return out, changes
