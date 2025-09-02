import re
import pandas as pd

def _to_number(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s == "" or s == "-":
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    s = re.sub(r"[^\d.\-]", "", s)
    if s.count(".") > 1:
        last = s.rfind(".")
        s = s[:last].replace(".", "") + s[last:]
    try:
        num = float(s)
        return -num if neg else num
    except Exception:
        return None

def normalize_numeric_columns(df: pd.DataFrame, headers):
    out = df.copy()
    for h in headers:
        if h in out.columns:
            out[h] = out[h].apply(_to_number)
    return out
