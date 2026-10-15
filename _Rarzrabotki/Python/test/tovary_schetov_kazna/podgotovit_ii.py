# -*- coding: utf-8 -*-
"""Подготовка счетов для распознавания ИИ: pdf → png (до 3 стр.), xls/xlsx/doc/docx/rtf → текст сетки, jpg/png как есть.
Результат — RAB/todo_ii.json с полем views (пути к картинкам и txt)."""
import json, os, sys
sys.stdout.reconfigure(encoding='utf-8')
import raspoznat as R
RAB = R.RAB
todo = json.load(open(os.path.join(RAB, "todo_ii.json"), encoding="utf-8"))
import pymupdf
for t in todo:
    views = []
    for p in t["files"]:
        ext = p.rsplit(".", 1)[-1].lower()
        base = os.path.splitext(p)[0]
        try:
            if ext == "pdf":
                doc = pymupdf.open(p)
                for i, page in enumerate(doc):
                    if i >= 3: break
                    out = f"{base}.p{i + 1}.png"
                    if not os.path.exists(out):
                        page.get_pixmap(dpi=110).save(out)
                    views.append(out)
            elif ext in ("jpg", "jpeg", "png"):
                views.append(p)
            else:
                grid = (R.grid_xls(p) if ext == "xls" else R.grid_xlsx(p) if ext == "xlsx"
                        else R.grid_docx(p) if ext == "docx" else R.grid_docx(R.to_docx(p)))
                out = base + ".grid.txt"
                with open(out, "w", encoding="utf-8") as fh:
                    for row in grid:
                        vals = [str(v).strip() for v in row if v is not None and str(v).strip()]
                        if vals:
                            fh.write(" | ".join(vals) + "\n")
                views.append(out)
        except Exception as e:
            views.append(f"ОШИБКА {p}: {e}")
    t["views"] = views
json.dump(todo, open(os.path.join(RAB, "todo_ii.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("подготовлено", len(todo), "счетов, видов", sum(len(t["views"]) for t in todo))
print([v for t in todo for v in t["views"] if v.startswith("ОШИБКА")][:5])
