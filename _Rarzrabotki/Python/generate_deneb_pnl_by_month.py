# -*- coding: utf-8 -*-
"""
Generate Deneb (Vega-Lite) visual for PnL by month on page
"Дубликат Cashflow об'єкт (казна) динаміка" (id=dbe7e85d8424dfc825fa).

Source: Table_Measures[Сума грн місяць PL] = SUM(Fact_PnL[Sum_Fact])
Axis:   Calendar.year_month (sorted by Year_Month_Sort)
Style:  IBCS-like — green for positive bars, red for negative
"""
import json
import os
import sys
import io

# UTF-8 stdout
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ROOT = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\PowerBi\Industrial"
REPORT = os.path.join(ROOT, "PL.Report")
PAGE_DIR = os.path.join(REPORT, "definition", "pages", "dbe7e85d8424dfc825fa")
VISUAL_DIR = os.path.join(PAGE_DIR, "visuals", "pnl_by_month_deneb")
REPORT_JSON = os.path.join(REPORT, "definition", "report.json")

os.makedirs(VISUAL_DIR, exist_ok=True)

# --- Vega-Lite spec (standalone form) ----------------------------------------
spec = {
    "$schema": "https://vega.github.io/schema/vega-lite/v6.json",
    "data": {"name": "dataset"},
    "title": {
        "text": "PnL по місяцях (Sum_Fact)",
        "anchor": "start",
        "fontSize": 14,
        "fontWeight": 600,
        "color": "#252423",
        "offset": 10,
    },
    "width": {"signal": "pbiContainerWidth - 25"},
    "height": {"signal": "pbiContainerHeight - 40"},
    "mark": {"type": "bar", "tooltip": True, "cornerRadiusEnd": 2},
    "encoding": {
        "x": {
            "field": "Місяць",
            "type": "ordinal",
            "axis": {
                "title": None,
                "labelAngle": -45,
                "labelFontSize": 11,
                "grid": False,
                "domainColor": "#a0a0a0",
                "tickColor": "#a0a0a0",
            },
        },
        "y": {
            "field": "Сума грн місяць PL",
            "type": "quantitative",
            "axis": {
                "title": None,
                "format": ",.0f",
                "grid": True,
                "gridDash": [2, 2],
                "gridColor": "#e0e0e0",
                "labelFontSize": 11,
                "domain": False,
                "ticks": False,
            },
        },
        "color": {
            "condition": {
                "test": "datum['Сума грн місяць PL'] >= 0",
                "value": "#107C41",
            },
            "value": "#C00000",
        },
        "tooltip": [
            {"field": "Місяць", "type": "ordinal"},
            {
                "field": "Сума грн місяць PL",
                "type": "quantitative",
                "format": ",.2f",
                "title": "PnL грн",
            },
        ],
    },
}

config = {
    "autosize": {"type": "fit", "contains": "padding"},
    "view": {"stroke": "transparent"},
    "font": "Segoe UI",
}


def pbir_str(obj: dict) -> str:
    """Stringify JSON, escape inner single quotes, wrap with single quotes."""
    s = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    s = s.replace("'", "''")
    return "'" + s + "'"


spec_lit = pbir_str(spec)
cfg_lit = pbir_str(config)

# --- visual.json -------------------------------------------------------------
visual = {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.8.0/schema.json",
    "name": "pnl_by_month_deneb",
    "position": {
        "x": 40,
        "y": 3120,
        "z": 30000,
        "height": 360,
        "width": 1890,
        "tabOrder": 30000,
    },
    "visual": {
        "visualType": "deneb7E15AEF80B9E4D4F8E12924291ECE89A",
        "query": {
            "queryState": {
                "dataset": {
                    "projections": [
                        {
                            "field": {
                                "Column": {
                                    "Expression": {"SourceRef": {"Entity": "Calendar"}},
                                    "Property": "year_month",
                                }
                            },
                            "queryRef": "Calendar.year_month",
                            "nativeQueryRef": "Місяць",
                            "active": True,
                        },
                        {
                            "field": {
                                "Measure": {
                                    "Expression": {
                                        "SourceRef": {"Entity": "Table_Measures"}
                                    },
                                    "Property": "Сума грн місяць PL",
                                }
                            },
                            "queryRef": "Table_Measures.Сума грн місяць PL",
                            "nativeQueryRef": "Сума грн місяць PL",
                            "active": True,
                        },
                    ]
                }
            },
            "sortDefinition": {
                "sort": [
                    {
                        "field": {
                            "Column": {
                                "Expression": {"SourceRef": {"Entity": "Calendar"}},
                                "Property": "Year_Month_Sort",
                            }
                        },
                        "direction": "Ascending",
                    }
                ],
                "isDefaultSort": True,
            },
        },
        "objects": {
            "vega": [
                {
                    "properties": {
                        "provider": {"expr": {"Literal": {"Value": "'vegaLite'"}}},
                        "jsonSpec": {"expr": {"Literal": {"Value": spec_lit}}},
                        "jsonConfig": {"expr": {"Literal": {"Value": cfg_lit}}},
                        "enableTooltips": {"expr": {"Literal": {"Value": "true"}}},
                        "enableContextMenu": {"expr": {"Literal": {"Value": "true"}}},
                        "enableSelection": {"expr": {"Literal": {"Value": "true"}}},
                        "enableHighlight": {"expr": {"Literal": {"Value": "true"}}},
                        "selectionMaxDataPoints": {
                            "expr": {"Literal": {"Value": "50D"}}
                        },
                        "logLevel": {"expr": {"Literal": {"Value": "3D"}}},
                        "isNewDialogOpen": {"expr": {"Literal": {"Value": "false"}}},
                    }
                }
            ],
            "stateManagement": [
                {
                    "properties": {
                        "viewportHeight": {"expr": {"Literal": {"Value": "300D"}}},
                        "viewportWidth": {"expr": {"Literal": {"Value": "1880D"}}},
                    }
                }
            ],
        },
        "visualContainerObjects": {
            "title": [
                {
                    "properties": {
                        "text": {
                            "expr": {"Literal": {"Value": "'PnL по місяцях (Deneb)'"}}
                        },
                        "show": {"expr": {"Literal": {"Value": "true"}}},
                    }
                }
            ]
        },
    },
}

visual_path = os.path.join(VISUAL_DIR, "visual.json")
with open(visual_path, "w", encoding="utf-8") as f:
    json.dump(visual, f, ensure_ascii=False, indent=2)
print(f"Wrote: {visual_path} size={os.path.getsize(visual_path)}")

# --- report.json: register Deneb in publicCustomVisuals ----------------------
with open(REPORT_JSON, "r", encoding="utf-8") as f:
    report = json.load(f)

pcv = report.get("publicCustomVisuals", [])
DENEB = "deneb7E15AEF80B9E4D4F8E12924291ECE89A"
if DENEB not in pcv:
    pcv.append(DENEB)
    report["publicCustomVisuals"] = pcv
    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"Added Deneb to publicCustomVisuals (total={len(pcv)})")
else:
    print("Deneb already registered")

# --- Validate JSON -----------------------------------------------------------
for p in (visual_path, REPORT_JSON):
    with open(p, "r", encoding="utf-8") as f:
        json.load(f)
    print(f"JSON valid: {p}")
