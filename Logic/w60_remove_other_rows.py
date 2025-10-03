import pandas as pd
from typing import Tuple, List, Iterable

TARGET_VALUES = {"אחר", "אחר אחר"}
# DEFAULT_COLS = ["מנהל סחר", "מנהל אזור", "מנהל איזור"]  # נזהה גם 'אזור' וגם 'איזור'
DEFAULT_COLS = ["מנהל סחר", "סוכן"]  # נזהה גם 'אזור' וגם 'איזור'
def remove_other_rows(df: pd.DataFrame, cols: Iterable[str] = DEFAULT_COLS) -> Tuple[pd.DataFrame, int]:
    present = [c for c in cols if c in df.columns]
    if not present:
        return df.copy(), 0
    # בנה מסכה: אמת אם באחת העמודות הערך בדיוק 'אחר' או 'אחר אחר' (אחרי strip)
    mask_any = None
    for c in present:
        s = df[c].astype(str).fillna("").str.strip()
        m = s.isin(TARGET_VALUES)
        mask_any = m if mask_any is None else (mask_any | m)
    removed = int(mask_any.sum())
    out = df.loc[~mask_any].reset_index(drop=True)
    return out, removed
