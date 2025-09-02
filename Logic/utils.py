import os
from datetime import datetime
from typing import List, Dict

def ts_now():
    return datetime.now().strftime("%Y-%m-%d_%H-%M")

def safe_mkdir(path: str):
    os.makedirs(path, exist_ok=True)
    return path

def dedupe_headers(headers: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    out: List[str] = []
    for h in headers:
        if h not in seen:
            seen[h] = 1
            out.append(h)
        else:
            seen[h] += 1
            out.append(f"{h} ({seen[h]})")
    return out

def base_from_path(path: str) -> str:
    return os.path.splitext(os.path.basename(path))[0]
