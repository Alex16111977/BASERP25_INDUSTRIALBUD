"""Step 01: Excel → 01_raw_sheets.json."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import config
from utils.excel_parser import parse_workbook


def main():
    config.JSON_DIR.mkdir(parents=True, exist_ok=True)
    out = []
    for f in config.EXCEL_FILES:
        print(f"Parsing: {f['path']}  (month={f['month_header']})")
        sheets = list(parse_workbook(f["path"], f["month_header"]))
        print(f"  {len(sheets)} project sheets")
        out.append({
            "label": f["label"],
            "period": f["period"],
            "source_file": f["path"],
            "sheets": sheets,
        })

    dst = config.JSON_DIR / "01_raw_sheets.json"
    dst.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    total_rows = sum(len(s["rows"]) for f in out for s in f["sheets"])
    print(f"Wrote {dst}  (periods: {len(out)}, rows total: {total_rows})")


if __name__ == "__main__":
    main()
