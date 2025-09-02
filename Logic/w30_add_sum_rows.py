import pandas as pd

def append_sum_rows(df: pd.DataFrame, sum_column_header: str) -> pd.DataFrame:
    df2 = df.copy()
    df2[sum_column_header] = pd.to_numeric(df2[sum_column_header], errors="coerce")
    total_sum = df2[sum_column_header].sum(skipna=True)
    blank = {col: None for col in df2.columns}
    return pd.concat(
        [df2, pd.DataFrame([blank]), pd.DataFrame([{**blank, sum_column_header: "סכום"}]), pd.DataFrame([{**blank, sum_column_header: float(total_sum)}])],
        ignore_index=True,
    )
