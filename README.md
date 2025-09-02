# wissotzky_stage1
שלב 1 – בנוי בסגנון של שלב 2 (Logic / tests / pipeline).

## התקנה
1. סביבה וירטואלית (מומלץ):
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```
2. התקנת תלויות:
   ```bash
   pip install -r requirements.txt
   ```

## הרצה
```bash
python pipeline/run_stage1.py --input "C:\path\QS.xlsx" --output-dir "C:\out" --sheet-name "sheet1"

in the terminal- vscode:
python -m pipeline.run_stage1 `                                                                           
>>   --input "C:\path\QS.xlsx" `
>>   --output-dir "C:\out" --sheet-name "sheet1" `
>>   --drop-empty `
>>   --split-by-manager `
>>   --market-private `
>>   --market-tedmiti `
>>   --pivot-private ` 
>>   --pivot-tedmiti `
>>   --by-agent       
```

- הפלט יישמר בשם: `<שם_מקור>_<YYYY-MM-DD>_<HH-MM>.xlsx` בתיקיית output שהוגדרה.
- הגיליון שנוצר נקרא **"מעובד"**.
- בעמודה H מתווספות שורות: ריקה, "סכום", והסכום המספרי.

## מבנה
```
wissotzky_stage1/
  Logic/
    headers_stage1.py
    utils.py
    w10_load_and_unmerge.py
    w15_detect_header.py
    w20_select_columns.py
    w30_add_sum_rows.py
    w40_finalize_save.py
  pipeline/
    run_stage1.py
  tests/
    test_smoke_pipeline.py
  output/
  requirements.txt
  README.md
```

## הערות
- אם יש כפילות שמות עמודות (למשל Sep פעמיים) – הלוגיקה יודעת לטפל.
- ניתן להרחיב בהמשך להוספת עמודות לחודשים Nov/Dec או כל לוגיקה נוספת.
