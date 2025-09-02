import pandas as pd
from typing import Iterable, List, Tuple

def drop_columns(df: pd.DataFrame, columns: Iterable[str] = ("שיטת תשלום לקוח משלם", "שיטת תשלום")) -> Tuple[pd.DataFrame, List[str]]:
    present = [c for c in columns if c in df.columns]
    if not present:
        return df.copy(), []
    out = df.drop(columns=present, errors="ignore")
    return out, present
