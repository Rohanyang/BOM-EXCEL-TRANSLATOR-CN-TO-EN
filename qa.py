from __future__ import annotations
from typing import List, Dict, Any
import csv

def write_qa_csv(path: str, rows: List[Dict[str, Any]]) -> None:
    if not rows:
        # 也创建空文件，方便流程自动化
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            f.write("sheet,row,col,cell,original\n")
        return

    fieldnames = ["sheet", "row", "col", "cell", "original"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)
