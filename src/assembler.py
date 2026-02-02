import pandas as pd
from typing import List, Dict, Any

def build_dataframe(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    # 컬럼 순서 고정
    cols = [
        "tag", "type", "color_group", "rgb", "hex",
        "rep_system", "page",
        "x1","y1","x2","y2"
    ]
    for c in cols:
        if c not in df.columns: df[c] = None
    return df[cols]
