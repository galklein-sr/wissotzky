import pandas as pd
from typing import Optional, List

def detect_header_and_frame(
    path: str,
    sheet_name: Optional[str] = None,
    desired_headers: Optional[List[str]] = None,
    search_rows: int = 120
) -> pd.DataFrame:
    """
    קורא את הגיליון (או הראשון אם sheet_name לא קיים/לא סופק), מאתר את שורת הכותרות
    (השורה הראשונה עם >=5 תאים לא ריקים), ומחזיר DataFrame מהשורה שאחרי הכותרות.
    """
    target = (sheet_name if sheet_name else 0)
    try:
        df_all = pd.read_excel(path, sheet_name=target, header=None, dtype=object)
    except Exception:
        # אם השם לא נמצא – ניפול לגיליון הראשון
        df_all = pd.read_excel(path, sheet_name=0, header=None, dtype=object)

    header_row_idx = None
    for i in range(min(search_rows, len(df_all))):
        if df_all.iloc[i].notna().sum() >= 5:
            header_row_idx = i
            break
    if header_row_idx is None:
        raise RuntimeError("Header row not found within first rows")

    columns = df_all.iloc[header_row_idx].astype(str).str.strip().tolist()
    df_all.columns = columns
    df = df_all.iloc[header_row_idx + 1:].reset_index(drop=True)
    return df
