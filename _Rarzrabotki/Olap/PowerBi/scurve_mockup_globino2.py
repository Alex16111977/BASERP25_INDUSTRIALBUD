"""
Мокап PBI Combo Chart (Line and Stacked Column) для S-curve прибутку проекту Глобино-2.
Імітує те, що користувач побачить у PL.pbix після додавання DAX-мір.

X-axis:           Calendar[year_month_short_ua]
Column (Y-prim):  [Сума казна ДДС] — місячний Net Cashflow (помісячна "пульсація")
Line  (Y-sec):    [Накопичена сума казна ДДС] — кумулятивна S-крива
Line  (Y-sec):    [Накопичена сума план ДДС (об'єкт)] — план для порівняння
Ref line: y=0 (Break-even)
"""
from __future__ import annotations
import sys
from pathlib import Path
import pandas as pd
import pyodbc
import plotly.graph_objects as go
from plotly.subplots import make_subplots
sys.stdout.reconfigure(encoding="utf-8")

DSN = ("DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;"
       "DATABASE=OlapBASERP;UID=sa;PWD=Brw739182465!;TrustServerCertificate=yes")

OUT = Path(__file__).parent / "output" / "img" / "scurve_pbi_mockup.png"
OUT.parent.mkdir(parents=True, exist_ok=True)

cn = pyodbc.connect(DSN, timeout=10)

sql_fact = """
SELECT f.Period_Month, SUM(f.Sum_Kazna) AS NetKazna
FROM Fact_Cashflow f
JOIN Dim_Departments dep ON f.Department_ID = dep.Department_ID
WHERE dep.Level2 = N'Глобино-2' AND f.Source = N'Казна'
GROUP BY f.Period_Month
"""
df = pd.read_sql(sql_fact, cn)
df["Period_Month"] = pd.to_datetime(df["Period_Month"])
df = df.sort_values("Period_Month").reset_index(drop=True)
df["Cum"] = df["NetKazna"].cumsum()

sql_plan = """
SELECT f.Period_Month, SUM(f.Sum_Plan_Object) AS PlanObj
FROM Fact_Cashflow f
JOIN Dim_Departments dep ON f.Department_ID = dep.Department_ID
WHERE dep.Level2 = N'Глобино-2' AND f.Source = N'ПланОбъекта'
GROUP BY f.Period_Month
"""
df_plan = pd.read_sql(sql_plan, cn)
df_plan["Period_Month"] = pd.to_datetime(df_plan["Period_Month"])
df_plan = df_plan.sort_values("Period_Month").reset_index(drop=True)
df_plan["CumPlan"] = df_plan["PlanObj"].cumsum()
cn.close()

df["ym"] = df["Period_Month"].dt.strftime("%Y-%m")
df_plan["ym"] = df_plan["Period_Month"].dt.strftime("%Y-%m")

fig = make_subplots(specs=[[{"secondary_y": True}]])

fig.add_trace(go.Bar(
    x=df["ym"], y=df["NetKazna"],
    name="[Сума казна ДДС] — місячна",
    marker=dict(color=["#2ca02c" if v >= 0 else "#d62728" for v in df["NetKazna"]],
                line=dict(width=0)),
    opacity=0.55,
    hovertemplate="%{x}<br>Місячна Net = %{y:,.0f} ₴<extra></extra>",
), secondary_y=False)

fig.add_trace(go.Scatter(
    x=df["ym"], y=df["Cum"],
    name="[Накопичена сума казна ДДС] — S-curve",
    mode="lines+markers",
    line=dict(color="#1f77b4", width=3.5),
    marker=dict(size=8, color="#1f77b4", line=dict(color="#fff", width=1.5)),
    hovertemplate="%{x}<br>Накопичено = %{y:,.0f} ₴<extra></extra>",
), secondary_y=True)

if not df_plan.empty:
    fig.add_trace(go.Scatter(
        x=df_plan["ym"], y=df_plan["CumPlan"],
        name="[Накопичена сума план ДДС (об'єкт)]",
        mode="lines+markers",
        line=dict(color="#9467bd", width=2, dash="dash"),
        marker=dict(size=6, color="#9467bd"),
        hovertemplate="%{x}<br>Plan накоп. = %{y:,.0f} ₴<extra></extra>",
    ), secondary_y=True)

fig.add_hline(y=0, line=dict(color="#666", width=1, dash="dot"), secondary_y=True,
              annotation_text="Break-even (y=0)", annotation_position="top right",
              annotation_font=dict(size=10, color="#666"))

if (df["Cum"] > 0).any():
    be = df[df["Cum"] > 0].iloc[0]
    fig.add_annotation(x=be["ym"], y=be["Cum"], xref="x", yref="y2",
                       text=f"★ Break-even<br>{be['ym']}<br>{be['Cum']/1e6:.1f} млн",
                       showarrow=True, arrowhead=2, ax=30, ay=-50,
                       font=dict(size=10, color="#0a0"),
                       bgcolor="rgba(255,255,255,0.85)", bordercolor="#0a0", borderwidth=1)

last = df.iloc[-1]
fig.add_annotation(x=last["ym"], y=last["Cum"], xref="x", yref="y2",
                   text=f"<b>{last['Cum']/1e6:.1f} млн ₴</b>",
                   showarrow=False, xshift=20, yshift=5,
                   font=dict(size=14, color="#1f77b4", family="Segoe UI"))

fig.update_layout(
    title=dict(
        text="<b>S-curve прибутку — Глобино-2 (Казна)</b><br>"
             "<sub>Мокап Combo Chart для PL.pbix · Стовпчик = місяць · Лінія = накопичено · Пунктир = план об'єкта</sub>",
        x=0, font=dict(size=16)),
    height=620,
    plot_bgcolor="#fafafa",
    paper_bgcolor="#ffffff",
    barmode="relative",
    bargap=0.25,
    hovermode="x unified",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                bgcolor="rgba(255,255,255,0.85)", bordercolor="#bbb", borderwidth=1),
    margin=dict(l=80, r=80, t=110, b=80),
    xaxis=dict(title="Calendar[year_month_short_ua]", tickangle=-45, gridcolor="#e7e7e7"),
)
fig.update_yaxes(title_text="Місячний Net (₴)", secondary_y=False,
                 tickformat=",.0f", gridcolor="#eee", zeroline=True, zerolinecolor="#888")
fig.update_yaxes(title_text="Накопичений Net (₴) — S-curve", secondary_y=True,
                 tickformat=",.0f", showgrid=False, zeroline=False)

fig.write_image(str(OUT), width=1600, height=620, scale=2)
print(f"OK: {OUT} ({OUT.stat().st_size//1024} KB)")
print(f"Net total: {df['NetKazna'].sum():,.0f} ₴ | Cum last: {df['Cum'].iloc[-1]:,.0f} ₴ | Plan cum: {df_plan['CumPlan'].iloc[-1] if not df_plan.empty else 0:,.0f} ₴")
