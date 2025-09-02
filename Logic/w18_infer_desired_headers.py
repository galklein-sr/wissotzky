import re
from typing import List, Tuple

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
MON_IDX = {m:i for i,m in enumerate(MONTHS)}

FIXED_PREFIX = [
    "מנהל סחר",
    "מנהל אזור",
    "סוכן",
    "ערוץ",
    "קוד לקוח קצה",
    "לקוח קצה",
    "קוד סוכן",
    'סה"כ סכום יתרת חוב',
    'סה"כ סכום יתרת חוב עד היום',
]

RE_PRE    = re.compile(r"סכום יתרת חוב\s*טרם\s*חודש\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*$")
RE_MONTH  = re.compile(r"סכום יתרת חוב\s*לחודש\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*$")
RE_TODAY  = re.compile(r"סכום יתרת חוב\s*לחודש\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*עד היום\s*$")
RE_FUTURE = re.compile(r"סכום יתרת חוב\s*לחודש\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*ממחר עד סוף החודש\s*$")
RE_BEYOND = re.compile(r"סכום יתרת חוב\s*מעבר לחודש\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*$")

def infer_month_columns(all_headers: List[str]) -> List[str]:
    """
    מקבלת את כל שמות העמודות מהקובץ, ומחזירה סדר עמודות 'חודשיות' הגיוני לפי מה שקיים בפועל.
    """
    out = [h for h in FIXED_PREFIX if h in all_headers]

    pre = []      # (idx, header)
    by_mon = {}   # mon -> list of variants in order
    beyond = []   # (idx, header)

    for h in all_headers:
        m = RE_PRE.match(h)
        if m:
            mon = m.group(1)
            pre.append((MON_IDX.get(mon, 99), h))
            continue
        m = RE_TODAY.match(h)
        if m:
            mon = m.group(1)
            by_mon.setdefault(mon, [])
            by_mon[mon].append(("TODAY", h))
            continue
        m = RE_FUTURE.match(h)
        if m:
            mon = m.group(1)
            by_mon.setdefault(mon, [])
            by_mon[mon].append(("FUTURE", h))
            continue
        m = RE_MONTH.match(h)
        if m:
            mon = m.group(1)
            by_mon.setdefault(mon, [])
            by_mon[mon].insert(0, ("MONTH", h))
            continue
        m = RE_BEYOND.match(h)
        if m:
            mon = m.group(1)
            beyond.append((MON_IDX.get(mon, 99), h))
            continue

    pre.sort(key=lambda x: x[0])
    out.extend([h for _,h in pre])

    order = {"MONTH":0, "TODAY":1, "FUTURE":2}
    for mon in MONTHS:
        if mon in by_mon:
            variants = by_mon[mon]
            variants.sort(key=lambda t: order.get(t[0], 9))
            out.extend([h for _,h in variants])

    beyond.sort(key=lambda x: x[0])
    out.extend([h for _,h in beyond])

    return out