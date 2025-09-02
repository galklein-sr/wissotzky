import pandas as pd

NOV = "סכום יתרת חוב לחודש Nov"
DEC = "סכום יתרת חוב לחודש Dec"
OCT = "סכום יתרת חוב לחודש Oct "
BEYOND_OCT = "סכום יתרת חוב מעבר לחודש Oct"

def add_optional_nov_dec(df: pd.DataFrame, insert_after: str = OCT, beyond_oct: str = BEYOND_OCT) -> pd.DataFrame:
    cols = list(df.columns)
    out = df.copy()

    def ensure_col(name: str):
        nonlocal out
        if name not in out.columns:
            out[name] = None
        cols2 = list(out.columns)
        if name in cols2 and insert_after in cols2:
            cols2.remove(name)
            pos = cols2.index(insert_after) + 1
            cols2.insert(pos, name)
            out = out[cols2]
        return out

    out = ensure_col(NOV)
    out = ensure_col(DEC)
    return out
