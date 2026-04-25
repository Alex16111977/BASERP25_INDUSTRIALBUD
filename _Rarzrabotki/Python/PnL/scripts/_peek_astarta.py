"""Peek Астарта. Тищенки per-period parsed data."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config

d = json.loads((config.JSON_DIR / "01_raw_sheets.json").read_text(encoding="utf-8"))
for p in d:
    for s in p["sheets"]:
        if "Астарта" in s["sheet_name"]:
            print(f"\n=== {p['label']} / {s['sheet_name']} ({len(s['rows'])} rows) ===")
            for r in s["rows"][:10]:
                print(f"  r{r['row']:3} {r['article'][:40]:40} f1={r['sum_f1']:>14.2f} f2={r['sum_f2']:>10.2f} tot={r['total']:>14.2f}")
