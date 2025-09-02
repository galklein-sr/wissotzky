import pandas as pd
from typing import List
from .utils import dedupe_headers

def select_and_order_columns(df: pd.DataFrame, desired_headers: List[str]) -> pd.DataFrame:
    desired_unique = dedupe_headers(desired_headers)
    out_cols = []
    for hdr in desired_unique:
        if hdr in df.columns:
            out_cols.append(hdr)
        else:
            candidates = [c for c in df.columns if c.strip() == hdr.strip()]
            if candidates:
                out_cols.append(candidates[0])
            else:
                df[hdr] = None
                out_cols.append(hdr)
    return df[out_cols].copy()
