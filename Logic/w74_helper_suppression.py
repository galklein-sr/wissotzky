from typing import List
import pandas as pd

HELPER_COL_NAME = "טור עזר"

def suppress_rows_by_helper(
    df: pd.DataFrame,
    month_cols: List[str],
    helper_col: str = HELPER_COL_NAME,
    threshold: float = -1000
) -> pd.DataFrame:
    """
    עבור כל שורה שבה ערך 'טור עזר' < threshold (ברירת מחדל -1000):
      - מנקים את עמודות החודשים (month_cols) לריק (None)
      - מציבים '-' בטור העזר (כטקסט)
    עמודות זהות (A..G) וסכומים H/I לא משתנים.
    """
    out = df.copy()
    if helper_col not in out.columns:
        return out

    helper_num = pd.to_numeric(out[helper_col], errors="coerce")
    mask = helper_num < threshold

    if not mask.any():
        return out

    # מנקים את עמודות החודשים
    for c in month_cols:
        if c in out.columns:
            out.loc[mask, c] = None

    # לפני הצבה של טקסט '-', נמיר את העמודה ל-object כדי למנוע FutureWarning
    out[helper_col] = out[helper_col].astype(object)
    out.loc[mask, helper_col] = "-"

    return out
