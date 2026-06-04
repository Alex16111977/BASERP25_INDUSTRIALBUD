"""
Cashflow analytics dashboard for project "Глобино-2" (source=Казна).
Generates 5 plotly charts (HTML + PNG) — inspired by Skanska/Strabag/Bechtel public reports.

Usage:
    .venv/Scripts/python.exe cashflow_analytics_globino2.py [--month YYYY-MM]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Final

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import pyodbc
from plotly.subplots import make_subplots

sys.stdout.reconfigure(encoding="utf-8")

DSN: Final[str] = (
    "DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;"
    "DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;TrustServerCertificate=yes"
)
OBJECT_NAME: Final[str] = "Глобино-2"
SOURCE: Final[str] = "Казна"
OUT_DIR: Final[Path] = Path(__file__).parent / "output"
IMG_DIR: Final[Path] = OUT_DIR / "img"
OUT_DIR.mkdir(parents=True, exist_ok=True)
IMG_DIR.mkdir(parents=True, exist_ok=True)

COLORS = {
    "Operating": "#2ca02c",
    "Investing": "#1f77b4",
    "Financing": "#ff7f0e",
    "Internal":  "#7f7f7f",
    "inflow":    "#2ca02c",
    "outflow":   "#d62728",
    "net":       "#222222",
    "plan":      "#9467bd",
    "bg":        "#fafafa",
}

CFS_MAP_KEYWORDS: dict[str, list[str]] = {
    "Financing": [
        "финансов", "финагенту", "кредит", "займ", "ссуд",
        "внутригруп", "дивиденд", "учредител",
    ],
    "Investing": [
        "основн", "капитальн", "капвлож", "оборудован",
        "приобрет", "инвестиц",
    ],
    "Internal": [
        "транзит", "перемещен", "конвертац",
    ],
}


def classify_cfs(name: str | None) -> str:
    if not name:
        return "Operating"
    lo = name.lower()
    for section, kws in CFS_MAP_KEYWORDS.items():
        if any(kw in lo for kw in kws):
            return section
    return "Operating"


def load_data() -> dict[str, pd.DataFrame]:
    cn = pyodbc.connect(DSN, timeout=15)
    sql_fact = """
    SELECT f.Period_Month,
           f.Movement_Type,
           f.Cash_Type,
           dds.DDS_Article_Name,
           f.Sum_Kazna AS Amount,
           f.Source
    FROM Fact_Cashflow f
    JOIN Dim_Departments dep ON f.Department_ID = dep.Department_ID
    LEFT JOIN Dim_DDS_Articles dds ON f.DDS_Article_ID = dds.DDS_Article_ID
    WHERE dep.Level2 = N'Глобино-2' AND f.Source = N'Казна'
    """
    df = pd.read_sql(sql_fact, cn)
    df["Period_Month"] = pd.to_datetime(df["Period_Month"])
    df["DDS_Article_Name"] = df["DDS_Article_Name"].fillna("Без статьи")
    df["CFS_Section"] = df["DDS_Article_Name"].apply(classify_cfs)
    df["Sign"] = df["Movement_Type"].map({"Поступление": 1, "Списание": -1}).fillna(1)
    df["Amount"] = df["Amount"].astype(float)

    sql_plan = """
    SELECT f.Period_Month, SUM(f.Sum_Plan_Object) AS PlanAmount
    FROM Fact_Cashflow f
    JOIN Dim_Departments dep ON f.Department_ID = dep.Department_ID
    WHERE dep.Level2 = N'Глобино-2' AND f.Source = N'ПланОбъекта'
    GROUP BY f.Period_Month
    """
    df_plan = pd.read_sql(sql_plan, cn)
    df_plan["Period_Month"] = pd.to_datetime(df_plan["Period_Month"])
    df_plan = df_plan.rename(columns={"PlanAmount": "Plan"})
    cn.close()
    return {"fact": df, "plan": df_plan}


def fmt_short(value: float) -> str:
    a = abs(value)
    if a >= 1e9:
        return f"{value/1e9:.2f} млрд"
    if a >= 1e6:
        return f"{value/1e6:.1f} млн"
    if a >= 1e3:
        return f"{value/1e3:.0f} тис"
    return f"{value:.0f}"


def chart1_waterfall_methods(df: pd.DataFrame) -> go.Figure:
    monthly = (df.groupby(["Period_Month", "CFS_Section"])["Amount"].sum().unstack(fill_value=0.0))
    for sec in ["Operating", "Investing", "Financing", "Internal"]:
        if sec not in monthly.columns:
            monthly[sec] = 0.0
    monthly = monthly[["Operating", "Investing", "Financing", "Internal"]]
    monthly["Net"] = monthly.sum(axis=1)

    x = [d.strftime("%Y-%m") for d in monthly.index]
    fig = go.Figure()
    for sec in ["Operating", "Investing", "Financing", "Internal"]:
        fig.add_trace(go.Bar(
            x=x, y=monthly[sec], name=sec, marker_color=COLORS[sec],
            hovertemplate=f"{sec}<br>%{{x}}<br>%{{y:,.0f}} ₴<extra></extra>",
        ))
    fig.add_trace(go.Scatter(
        x=x, y=monthly["Net"], name="Net Change", mode="lines+markers",
        line=dict(color=COLORS["net"], width=3),
        marker=dict(size=8, color=COLORS["net"]),
        hovertemplate="Net<br>%{x}<br>%{y:,.0f} ₴<extra></extra>",
    ))
    fig.update_layout(
        title=f"<b>1. Direct Method — місячні грошові потоки по типах діяльності</b><br>"
              f"<sub>{OBJECT_NAME} | джерело: {SOURCE} | {monthly.index.min().date()} … {monthly.index.max().date()}</sub>",
        barmode="relative",
        xaxis_title="Місяць",
        yaxis_title="₴",
        height=600,
        plot_bgcolor=COLORS["bg"],
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
    )
    fig.update_yaxes(tickformat=",.0f", zeroline=True, zerolinewidth=2, zerolinecolor="#444")
    return fig


def chart2_cash_bridge(df: pd.DataFrame) -> go.Figure:
    df_g = df.groupby(["DDS_Article_Name", "Movement_Type"])["Amount"].sum().reset_index()
    inflow_top = (df_g[df_g.Movement_Type == "Поступление"].nlargest(5, "Amount")[["DDS_Article_Name", "Amount"]])
    outflow_top = (df_g[df_g.Movement_Type == "Списание"].nsmallest(5, "Amount")[["DDS_Article_Name", "Amount"]])

    inflow_total = df[df.Movement_Type == "Поступление"]["Amount"].sum()
    outflow_total = df[df.Movement_Type == "Списание"]["Amount"].sum()
    inflow_other = inflow_total - inflow_top["Amount"].sum()
    outflow_other = outflow_total - outflow_top["Amount"].sum()

    labels = ["Старт (опер. початок)"]
    values = [0.0]
    measures = ["absolute"]

    for _, r in inflow_top.iterrows():
        labels.append(r["DDS_Article_Name"])
        values.append(float(r["Amount"]))
        measures.append("relative")
    if abs(inflow_other) > 1:
        labels.append("Інші надходження")
        values.append(float(inflow_other))
        measures.append("relative")

    for _, r in outflow_top.iterrows():
        labels.append(r["DDS_Article_Name"])
        values.append(float(r["Amount"]))
        measures.append("relative")
    if abs(outflow_other) > 1:
        labels.append("Інші витрати")
        values.append(float(outflow_other))
        measures.append("relative")

    labels.append("Накопичений Net")
    values.append(0.0)
    measures.append("total")

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=measures,
        x=labels,
        y=values,
        text=[fmt_short(v) if v != 0 else "" for v in values],
        textposition="outside",
        increasing=dict(marker=dict(color=COLORS["inflow"])),
        decreasing=dict(marker=dict(color=COLORS["outflow"])),
        totals=dict(marker=dict(color="#444")),
        connector=dict(line=dict(color="#bbb", dash="dot")),
    ))
    fig.update_layout(
        title=f"<b>2. Cash Bridge — від чого і куди йдуть гроші (увесь період)</b><br>"
              f"<sub>{OBJECT_NAME} | джерело: {SOURCE} | агрегація по top-5 інфлоу та top-5 аутфлоу</sub>",
        height=620,
        plot_bgcolor=COLORS["bg"],
        xaxis_tickangle=-30,
        yaxis_title="₴ (накопичено)",
        showlegend=False,
        margin=dict(b=180),
    )
    fig.update_yaxes(tickformat=",.0f", zeroline=True, zerolinecolor="#444")
    return fig


def chart3_sankey(df: pd.DataFrame, month: pd.Timestamp | None = None) -> go.Figure:
    if month is None:
        month = df["Period_Month"].max()
    period_label = pd.Timestamp(month).strftime("%Y-%m")
    df_m = df[df["Period_Month"] == month].copy()
    df_m["AbsAmount"] = df_m["Amount"].abs()

    cash_types = sorted(df_m["Cash_Type"].dropna().unique())
    cfs_sections = ["Operating", "Investing", "Financing", "Internal"]

    top_articles = (df_m.groupby("DDS_Article_Name")["AbsAmount"].sum().nlargest(12).index.tolist())
    df_m["Article_Bucket"] = df_m["DDS_Article_Name"].where(df_m["DDS_Article_Name"].isin(top_articles), "Інші статті")
    articles_in_use = list(dict.fromkeys(list(top_articles) + (["Інші статті"] if (df_m["Article_Bucket"] == "Інші статті").any() else [])))

    nodes = cash_types + cfs_sections + articles_in_use
    idx = {n: i for i, n in enumerate(nodes)}

    sources, targets, values, colors = [], [], [], []

    g1 = df_m.groupby(["Cash_Type", "CFS_Section"])["AbsAmount"].sum().reset_index()
    for _, r in g1.iterrows():
        if r["AbsAmount"] < 1: continue
        sources.append(idx[r["Cash_Type"]])
        targets.append(idx[r["CFS_Section"]])
        values.append(float(r["AbsAmount"]))
        colors.append("rgba(31,119,180,0.35)")

    g2 = df_m.groupby(["CFS_Section", "Article_Bucket", "Movement_Type"])["AbsAmount"].sum().reset_index()
    for _, r in g2.iterrows():
        if r["AbsAmount"] < 1: continue
        sources.append(idx[r["CFS_Section"]])
        targets.append(idx[r["Article_Bucket"]])
        values.append(float(r["AbsAmount"]))
        colors.append("rgba(44,160,44,0.45)" if r["Movement_Type"] == "Поступление" else "rgba(214,39,40,0.45)")

    node_colors = (["#9ecae1"]*len(cash_types) + [COLORS[s] for s in cfs_sections] + ["#e0e0e0"]*len(articles_in_use))

    fig = go.Figure(go.Sankey(
        node=dict(label=nodes, color=node_colors, pad=18, thickness=18, line=dict(color="#888", width=0.5)),
        link=dict(source=sources, target=targets, value=values, color=colors,
                  hovertemplate="%{source.label} → %{target.label}<br>%{value:,.0f} ₴<extra></extra>"),
    ))
    fig.update_layout(
        title=f"<b>3. Sankey — трубопровід грошей за {period_label}</b><br>"
              f"<sub>Cash Type → CFS Section → DDS-стаття | зелені/червоні стрілки = inflow/outflow</sub>",
        height=620,
        font=dict(size=11),
        plot_bgcolor=COLORS["bg"],
    )
    return fig


def chart4_treemap_heatmap(df: pd.DataFrame) -> go.Figure:
    df_in = df[df.Movement_Type == "Поступление"].groupby("DDS_Article_Name", as_index=False)["Amount"].sum()
    df_out = df[df.Movement_Type == "Списание"].groupby("DDS_Article_Name", as_index=False)["Amount"].sum()
    df_out["Amount"] = df_out["Amount"].abs()
    df_in = df_in.nlargest(10, "Amount")
    df_out = df_out.nlargest(10, "Amount")

    heat = (df.groupby(["Period_Month", "CFS_Section"])["Amount"].sum().unstack(fill_value=0.0))
    if "Operating" not in heat.columns: heat["Operating"] = 0.0
    if "Financing" not in heat.columns: heat["Financing"] = 0.0
    if "Investing" not in heat.columns: heat["Investing"] = 0.0
    if "Internal"  not in heat.columns: heat["Internal"]  = 0.0
    heat["Net"] = heat.sum(axis=1)
    heat = heat[["Operating", "Investing", "Financing", "Internal", "Net"]]
    heat.index = [d.strftime("%Y-%m") for d in heat.index]

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            "Top-10 надходжень (₴, увесь період)",
            "Top-10 виплат (₴, увесь період)",
            "Heatmap: Net Cashflow по місяцях × секція",
            None,
        ),
        specs=[[{"type":"treemap"},{"type":"treemap"}], [{"type":"heatmap","colspan":2}, None]],
        vertical_spacing=0.12, row_heights=[0.45, 0.55],
    )
    fig.add_trace(go.Treemap(
        labels=df_in["DDS_Article_Name"], parents=[""]*len(df_in), values=df_in["Amount"],
        marker=dict(colorscale="Greens", colors=df_in["Amount"]),
        texttemplate="<b>%{label}</b><br>%{value:,.0f} ₴<br>(%{percentParent})",
        hovertemplate="%{label}<br>%{value:,.0f} ₴<extra></extra>",
    ), row=1, col=1)
    fig.add_trace(go.Treemap(
        labels=df_out["DDS_Article_Name"], parents=[""]*len(df_out), values=df_out["Amount"],
        marker=dict(colorscale="Reds", colors=df_out["Amount"]),
        texttemplate="<b>%{label}</b><br>%{value:,.0f} ₴<br>(%{percentParent})",
        hovertemplate="%{label}<br>%{value:,.0f} ₴<extra></extra>",
    ), row=1, col=2)

    fig.add_trace(go.Heatmap(
        z=heat.values, x=heat.columns.tolist(), y=heat.index.tolist(),
        colorscale="RdYlGn", zmid=0,
        text=[[fmt_short(v) for v in row] for row in heat.values],
        texttemplate="%{text}", textfont=dict(size=10),
        hovertemplate="%{y} | %{x}<br>%{z:,.0f} ₴<extra></extra>",
        colorbar=dict(title="₴", thickness=12, len=0.5, y=0.22),
    ), row=2, col=1)

    fig.update_layout(
        title=f"<b>4. Структура й сезонність — Top-Inflow / Top-Outflow / Heatmap</b><br>"
              f"<sub>{OBJECT_NAME} | джерело: {SOURCE}</sub>",
        height=900,
        plot_bgcolor=COLORS["bg"],
        showlegend=False,
    )
    return fig


def chart5_cumulative_scurve(df: pd.DataFrame, df_plan: pd.DataFrame) -> go.Figure:
    monthly = df.groupby("Period_Month")["Amount"].sum().sort_index()
    cum = monthly.cumsum()
    x = [d.strftime("%Y-%m") for d in cum.index]

    pos_idx = (cum > 0).idxmax() if (cum > 0).any() else None
    min_idx = cum.idxmin()

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=cum.values, name="Cumulative Net Cashflow (факт, Казна)",
        mode="lines+markers", line=dict(color="#1f77b4", width=3),
        fill="tozeroy", fillcolor="rgba(31,119,180,0.18)",
        hovertemplate="%{x}<br>%{y:,.0f} ₴<extra></extra>",
    ))

    if not df_plan.empty:
        df_plan_s = df_plan.sort_values("Period_Month")
        df_plan_s["Cum"] = df_plan_s["Plan"].cumsum()
        fig.add_trace(go.Scatter(
            x=[d.strftime("%Y-%m") for d in df_plan_s["Period_Month"]],
            y=df_plan_s["Cum"].values,
            name="Plan (ПланОбъекта, накоп.)",
            mode="lines", line=dict(color=COLORS["plan"], width=2, dash="dash"),
            hovertemplate="%{x}<br>Plan %{y:,.0f} ₴<extra></extra>",
        ))

    if pos_idx is not None:
        be_x = pos_idx.strftime("%Y-%m")
        fig.add_annotation(x=be_x, y=cum.loc[pos_idx], text=f"★ Break-even<br>{be_x}<br>{fmt_short(cum.loc[pos_idx])} ₴",
                           showarrow=True, arrowhead=2, ax=20, ay=-50, font=dict(size=11, color="#0a0"))
    fig.add_annotation(x=min_idx.strftime("%Y-%m"), y=cum.loc[min_idx],
                       text=f"▼ Дно<br>{min_idx.strftime('%Y-%m')}<br>{fmt_short(cum.loc[min_idx])} ₴",
                       showarrow=True, arrowhead=2, ax=-20, ay=40, font=dict(size=11, color="#c00"))

    fig.update_layout(
        title=f"<b>5. Cumulative S-Curve — накопичений Net Cashflow і план</b><br>"
              f"<sub>{OBJECT_NAME} | факт: Казна | план: ПланОбъекта</sub>",
        height=560,
        plot_bgcolor=COLORS["bg"],
        xaxis_title="Місяць", yaxis_title="Накопичений Net ₴",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
    )
    fig.update_yaxes(tickformat=",.0f", zeroline=True, zerolinewidth=2, zerolinecolor="#444")
    return fig


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", type=str, default=None, help="Month for Sankey, format YYYY-MM (default: last)")
    args = ap.parse_args()

    print(f"[1/4] Loading data for {OBJECT_NAME} (source={SOURCE})…")
    data = load_data()
    df = data["fact"]
    print(f"  fact rows: {len(df)}, periods: {df['Period_Month'].min().date()}..{df['Period_Month'].max().date()}")
    print(f"  inflow total : {df.loc[df.Movement_Type=='Поступление','Amount'].sum():>20,.2f} ₴")
    print(f"  outflow total: {df.loc[df.Movement_Type=='Списание','Amount'].sum():>20,.2f} ₴")
    print(f"  NET total    : {df['Amount'].sum():>20,.2f} ₴")
    print(f"  plan rows: {len(data['plan'])}")

    sankey_month = pd.Timestamp(args.month + "-01") if args.month else None

    print("[2/4] Building 5 plotly figures…")
    figs = [
        ("chart1_waterfall_methods",  chart1_waterfall_methods(df)),
        ("chart2_cash_bridge",        chart2_cash_bridge(df)),
        ("chart3_sankey",             chart3_sankey(df, sankey_month)),
        ("chart4_treemap_heatmap",    chart4_treemap_heatmap(df)),
        ("chart5_scurve",             chart5_cumulative_scurve(df, data["plan"])),
    ]

    print("[3/4] Exporting PNGs (kaleido)…")
    for name, fig in figs:
        png = IMG_DIR / f"{name}.png"
        fig.write_image(str(png), width=1500, height=fig.layout.height or 600, scale=2)
        print(f"  ✓ {png.name} ({png.stat().st_size//1024} KB)")

    print("[4/4] Writing combined HTML…")
    html_path = OUT_DIR / "cashflow_globino2.html"
    parts = [
        "<!doctype html><html lang='uk'><head><meta charset='utf-8'>",
        f"<title>Cashflow {OBJECT_NAME} ({SOURCE})</title>",
        "<style>body{font-family:Segoe UI,Arial,sans-serif;background:#f4f4f4;margin:0;padding:24px}",
        "h1{color:#222;margin:8px 0 24px}.chart{background:#fff;border-radius:10px;",
        "box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:24px;padding:8px}</style></head><body>",
        f"<h1>📊 Cashflow analytics — {OBJECT_NAME} (джерело: {SOURCE})</h1>",
    ]
    for name, fig in figs:
        parts.append("<div class='chart'>")
        parts.append(fig.to_html(full_html=False, include_plotlyjs="cdn" if name == "chart1_waterfall_methods" else False))
        parts.append("</div>")
    parts.append("</body></html>")
    html_path.write_text("\n".join(parts), encoding="utf-8")
    print(f"  ✓ {html_path} ({html_path.stat().st_size//1024} KB)")
    print("Done.")


if __name__ == "__main__":
    main()
