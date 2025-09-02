import pandas as pd
from typing import Tuple

def filter_agent_code_required(df: pd.DataFrame, col: str = "קוד סוכן") -> Tuple[pd.DataFrame, int]:
    """
    משאיר רק שורות שבהן 'קוד סוכן' הוא ספרתי לגמרי (לפחות ספרה אחת).
    שורות ריקות / תווי לא-ספרות -> נמחקות.
    מחזיר (df מסונן, מספר שורות שהוסרו).
    הערה: יש להריץ את זה אחרי normalize_agent_code כדי שהעמודה כבר תהיה ספרות בלבד או ריקה.
    """
    if col not in df.columns:
        return df.copy(), 0
    s = df[col].astype(str).str.strip()
    mask_valid = s.str.fullmatch(r"\d+")
    removed = int((~mask_valid).sum())
    out = df.loc[mask_valid].reset_index(drop=True)
    return out, removed
