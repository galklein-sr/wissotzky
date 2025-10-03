# import pandas as pd
# from typing import Tuple

# def normalize_agent_code(df: pd.DataFrame, col: str = "קוד סוכן") -> Tuple[pd.DataFrame, int]:
#     """
#     משאיר רק ספרות ב'קוד סוכן' וגם "אחר"; מחרוזות שמתרוקנות -> NaN.
#     מחזיר גם כמות שינויים.
#     """
#     if col not in df.columns:
#         return df.copy(), 0
#     s = df[col].astype(str)
#     new = s.str.replace(r"\D+", "", regex=True).str.strip()
#     changes = int((s != new).sum())
#     out = df.copy()
#     out[col] = new.replace({"": None})
#     return out, changes



from typing import Tuple
import pandas as pd

def normalize_agent_code(
    df: pd.DataFrame,
    col: str = "קוד סוכן",
    keep_values = ("אחר",),   
) -> Tuple[pd.DataFrame, int]:
    """
    מנרמל את עמודת 'קוד סוכן':
    - שומר ערכים ברשימת keep_values (כאן: "אחר") כמו שהם.
    - בכל הערכים האחרים משאיר רק ספרות (0-9).
    - ריק לאחר הנירמול ⇒ None.
    """
    if col not in df.columns:
        return df.copy(), 0
    # הגנה: אם בטעות יעבור מחרוזת ולא רשימה/טופל, נהפוך לרשימה
    if isinstance(keep_values, str):
        keep_values = [keep_values]
    out = df.copy()
    s_raw = out[col]
    s = s_raw.astype("string").str.strip()
    keep_mask  = s.isin(keep_values)
    empty_mask = s.isna() | (s == "") | s.str.fullmatch(r"[-–—]+", na=False)
    digits_only = s.str.replace(r"\D+", "", regex=True).str.strip()
    new = digits_only.where(~keep_mask, s)
    new = new.mask(empty_mask | (new == ""), other=None)
    before = s_raw.astype("string")
    after  = new.astype("string")
    changes = int((before.fillna("__NA__") != after.fillna("__NA__")).sum())
    out[col] = new
    return out, changes

