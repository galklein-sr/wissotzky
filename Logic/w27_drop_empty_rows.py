import pandas as pd
from typing import Iterable, Tuple

DEFAULT_REQUIRED_ANY = ["קוד סוכן", "קוד לקוח קצה", "לקוח קצה"]

def _is_empty_series(s: pd.Series) -> pd.Series:
    return s.isna() | (s.astype(str).str.strip() == "")

def drop_empty_rows(df: pd.DataFrame, required_any: Iterable[str] = DEFAULT_REQUIRED_ANY) -> Tuple[pd.DataFrame, int, int]:
    """
    מוחק:
      1) שורות ריקות לגמרי
      2) שורות שאין בהן אף ערך בעמודות ה-required_any (אלו שמופיעות בפועל ב-DF)
    """
    out = df.copy()
    empty_mask = out.apply(_is_empty_series).all(axis=1)
    removed_fully = int(empty_mask.sum())
    out = out.loc[~empty_mask].reset_index(drop=True)

    req_cols = [c for c in required_any if c in out.columns]
    removed_required = 0
    if req_cols:
        req_empty = out[req_cols].apply(_is_empty_series).all(axis=1)
        removed_required = int(req_empty.sum())
        out = out.loc[~req_empty].reset_index(drop=True)

    return out, removed_fully, removed_required
