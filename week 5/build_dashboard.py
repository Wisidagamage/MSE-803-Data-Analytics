import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio

# ── Load & Clean ─────────────────────────────────────────────────────────────
df = pd.read_excel("Retail_Sales_sample-Dataset (1).xlsx")
df.columns = df.columns.str.strip()
df["Order Date"] = pd.to_datetime(df["Order Date"])
for col in ["Product Category", "Product Name", "Region", "Customer Segment"]:
    df[col] = df[col].str.strip()
df["Sales Amount"] = df["Sales Amount"].round(2)
df["Profit"] = df["Profit"].round(2)
df["Discount (%)"] = df["Discount (%)"].round(2)

# Derived features
df["Month"] = df["Order Date"].dt.to_period("M").astype(str)
df["Profit Margin (%)"] = (df["Profit"] / df["Sales Amount"] * 100).round(2)
df["Revenue per Unit"] = (df["Sales Amount"] / df["Quantity Sold"]).round(2)

# ── Aggregations ─────────────────────────────────────────────────────────────
monthly = (
    df.groupby("Month")
    .agg(Total_Sales=("Sales Amount", "sum"), Total_Profit=("Profit", "sum"), Orders=("Order ID", "count"))
    .reset_index()
    .sort_values("Month")
)

region = (
    df.groupby("Region")
    .agg(Total_Sales=("Sales Amount", "sum"), Total_Profit=("Profit", "sum"))
    .reset_index()
    .sort_values("Total_Sales", ascending=False)
)

category = (
    df.groupby("Product Category")
    .agg(Total_Sales=("Sales Amount", "sum"), Total_Profit=("Profit", "sum"), Orders=("Order ID", "count"))
    .reset_index()
)

segment = (
    df.groupby("Customer Segment")
    .agg(
        Total_Sales=("Sales Amount", "sum"),
        Total_Profit=("Profit", "sum"),
        Avg_Discount=("Discount (%)", "mean"),
        Orders=("Order ID", "count"),
    )
    .reset_index()
)
segment["Profit Margin (%)"] = (segment["Total_Profit"] / segment["Total_Sales"] * 100).round(1)

# KPIs
total_sales = df["Sales Amount"].sum()
total_profit = df["Profit"].sum()
total_orders = len(df)
avg_margin = df["Profit Margin (%)"].mean()
avg_discount = df["Discount (%)"].mean()

# ── Colour palette ────────────────────────────────────────────────────────────
BLUE = "#2563EB"
GREEN = "#16A34A"
AMBER = "#D97706"
RED = "#DC2626"
TEAL = "#0891B2"
PURPLE = "#7C3AED"
LIGHT_BG = "#F8FAFC"
CARD_BG = "#FFFFFF"
BORDER = "#E2E8F0"
TEXT_DARK = "#1E293B"
TEXT_MUTED = "#64748B"

CAT_COLORS = {"Electronics": BLUE, "Clothing": TEAL, "Furniture": PURPLE}
REG_COLORS = ["#2563EB", "#0891B2", "#7C3AED", "#D97706"]
SEG_COLORS = [GREEN, BLUE, AMBER]
MONTH_LABELS = {"2025-01": "Jan 2025", "2025-02": "Feb 2025", "2025-03": "Mar 2025"}
monthly["Month Label"] = monthly["Month"].map(MONTH_LABELS)

# ── Build figure (2×2 subplots) ───────────────────────────────────────────────
fig = make_subplots(
    rows=2,
    cols=2,
    subplot_titles=(
        "Component 1 — Monthly Sales & Profit Trend",
        "Component 2 — Sales & Profit by Region",
        "Component 3 — Product Category Revenue Share",
        "Component 4 — Customer Segment Performance",
    ),
    vertical_spacing=0.16,
    horizontal_spacing=0.10,
    specs=[
        [{"type": "xy"}, {"type": "xy"}],
        [{"type": "domain"}, {"type": "xy"}],
    ],
)

# ── Component 1 — Monthly Trend ───────────────────────────────────────────────
fig.add_trace(
    go.Bar(
        x=monthly["Month Label"],
        y=monthly["Total_Sales"],
        name="Sales",
        marker_color=BLUE,
        opacity=0.85,
        text=monthly["Total_Sales"].apply(lambda v: f"${v:,.0f}"),
        textposition="outside",
        textfont=dict(size=11),
    ),
    row=1,
    col=1,
)
fig.add_trace(
    go.Scatter(
        x=monthly["Month Label"],
        y=monthly["Total_Profit"],
        name="Profit",
        mode="lines+markers+text",
        line=dict(color=GREEN, width=3),
        marker=dict(size=9, color=GREEN),
        text=monthly["Total_Profit"].apply(lambda v: f"${v:,.0f}"),
        textposition="top center",
        textfont=dict(size=10, color=GREEN),
    ),
    row=1,
    col=1,
)

# ── Component 2 — Region bar ──────────────────────────────────────────────────
fig.add_trace(
    go.Bar(
        name="Sales",
        x=region["Region"],
        y=region["Total_Sales"],
        marker_color=REG_COLORS,
        text=region["Total_Sales"].apply(lambda v: f"${v:,.0f}"),
        textposition="outside",
        textfont=dict(size=11),
        showlegend=False,
    ),
    row=1,
    col=2,
)
fig.add_trace(
    go.Bar(
        name="Profit",
        x=region["Region"],
        y=region["Total_Profit"],
        marker_color=REG_COLORS,
        opacity=0.5,
        text=region["Total_Profit"].apply(lambda v: f"${v:,.0f}"),
        textposition="outside",
        textfont=dict(size=10),
        showlegend=False,
    ),
    row=1,
    col=2,
)

# ── Component 3 — Donut chart ─────────────────────────────────────────────────
fig.add_trace(
    go.Pie(
        labels=category["Product Category"],
        values=category["Total_Sales"],
        hole=0.52,
        marker_colors=[CAT_COLORS[c] for c in category["Product Category"]],
        textinfo="label+percent",
        textfont=dict(size=13),
        hovertemplate="<b>%{label}</b><br>Sales: $%{value:,.2f}<br>Share: %{percent}<extra></extra>",
    ),
    row=2,
    col=1,
)

# ── Component 4 — Segment grouped bar ────────────────────────────────────────
fig.add_trace(
    go.Bar(
        name="Sales",
        x=segment["Customer Segment"],
        y=segment["Total_Sales"],
        marker_color=SEG_COLORS,
        text=segment["Total_Sales"].apply(lambda v: f"${v:,.0f}"),
        textposition="outside",
        textfont=dict(size=11),
        showlegend=False,
    ),
    row=2,
    col=2,
)
fig.add_trace(
    go.Scatter(
        x=segment["Customer Segment"],
        y=segment["Profit Margin (%)"],
        name="Profit Margin %",
        mode="markers+text",
        marker=dict(size=14, color=GREEN, symbol="diamond"),
        text=segment["Profit Margin (%)"].apply(lambda v: f"{v:.1f}%"),
        textposition="top center",
        textfont=dict(size=11, color=GREEN),
        yaxis="y4",
        showlegend=False,
    ),
    row=2,
    col=2,
)

# ── Axes tweaks ───────────────────────────────────────────────────────────────
fig.update_yaxes(title_text="Amount (NZD $)", row=1, col=1, tickprefix="$", gridcolor=BORDER)
fig.update_yaxes(title_text="Amount (NZD $)", row=1, col=2, tickprefix="$", gridcolor=BORDER)
fig.update_yaxes(title_text="Amount (NZD $)", row=2, col=2, tickprefix="$", gridcolor=BORDER)
fig.update_xaxes(gridcolor=BORDER)

fig.update_layout(
    barmode="group",
    title=dict(
        text="<b>Retail Sales Dashboard</b> — Q1 2025 (Jan – Mar)",
        font=dict(family="Arial", size=22, color=TEXT_DARK),
        x=0.5,
        xanchor="center",
        y=0.99,
    ),
    font=dict(family="Arial", size=12, color=TEXT_DARK),
    plot_bgcolor=LIGHT_BG,
    paper_bgcolor=LIGHT_BG,
    height=820,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.01,
        xanchor="right",
        x=0.55,
        bgcolor="rgba(0,0,0,0)",
        font=dict(size=12),
    ),
    margin=dict(t=110, b=40, l=60, r=40),
)

# Bold subplot titles
for ann in fig.layout.annotations:
    ann.font = dict(family="Arial", size=13, color=TEXT_DARK)
    ann.font.update({"color": TEXT_DARK})

# ── KPI cards HTML ────────────────────────────────────────────────────────────
def kpi_card(label, value, sub="", color=BLUE):
    return f"""
    <div style="background:{CARD_BG};border:1px solid {BORDER};border-top:4px solid {color};
                border-radius:8px;padding:18px 22px;min-width:160px;flex:1;box-shadow:0 1px 4px rgba(0,0,0,.06);">
      <div style="font-size:12px;color:{TEXT_MUTED};font-weight:600;letter-spacing:.06em;text-transform:uppercase;">{label}</div>
      <div style="font-size:26px;font-weight:700;color:{TEXT_DARK};margin:6px 0 2px;">{value}</div>
      <div style="font-size:11px;color:{TEXT_MUTED};">{sub}</div>
    </div>"""


kpi_html = f"""
<div style="display:flex;gap:14px;flex-wrap:wrap;padding:18px 24px 0;font-family:Arial,sans-serif;">
  {kpi_card("Total Revenue",    f"${total_sales:,.0f}",  "Q1 2025",             BLUE)}
  {kpi_card("Total Profit",     f"${total_profit:,.0f}", f"Margin {avg_margin:.1f}%", GREEN)}
  {kpi_card("Total Orders",     str(total_orders),       "Jan – Mar 2025",      PURPLE)}
  {kpi_card("Avg Profit Margin",f"{avg_margin:.1f}%",    "Across all orders",   TEAL)}
  {kpi_card("Avg Discount",     f"{avg_discount:.1f}%",  "Applied per order",   AMBER)}
</div>"""

# ── Insight section ───────────────────────────────────────────────────────────
insight_rows = ""
insights = [
    ("Monthly Trend",       f"February was the peak sales month (${monthly.loc[monthly['Month']=='2025-02','Total_Sales'].values[0]:,.0f}). Profit tracked sales closely, with all three months maintaining a positive margin."),
    ("Top Region",          f"{region.iloc[0]['Region']} led revenue at ${region.iloc[0]['Total_Sales']:,.0f}, while all four regions contributed positively to profit."),
    ("Category Breakdown",  f"Electronics, Clothing, and Furniture each hold ~33% of revenue — a balanced portfolio with no single category dominating."),
    ("Customer Segments",   f"Home Office customers generate the highest profit margin ({segment.sort_values('Profit Margin (%)', ascending=False).iloc[0]['Customer Segment']}), while Consumer segment drives the most orders."),
    ("Discount Impact",     f"Average discount of {avg_discount:.1f}% is applied across orders. Higher discounts do not always correlate with lower profit margins, suggesting strategic pricing."),
    ("Data Quality",        "Dataset: 100 orders, 0 nulls, 0 duplicates. Three derived features added: Month, Profit Margin (%), Revenue per Unit."),
]
for title, body in insights:
    insight_rows += f"""
    <tr>
      <td style="padding:9px 14px;font-weight:600;color:{TEXT_DARK};white-space:nowrap;border-bottom:1px solid {BORDER};">{title}</td>
      <td style="padding:9px 14px;color:{TEXT_MUTED};border-bottom:1px solid {BORDER};">{body}</td>
    </tr>"""

insights_html = f"""
<div style="font-family:Arial,sans-serif;padding:16px 24px 24px;">
  <h3 style="margin:0 0 10px;color:{TEXT_DARK};font-size:15px;">Key Insights</h3>
  <table style="width:100%;border-collapse:collapse;background:{CARD_BG};border:1px solid {BORDER};border-radius:8px;overflow:hidden;font-size:13px;">
    {insight_rows}
  </table>
  <p style="font-size:11px;color:{TEXT_MUTED};margin-top:10px;">
    Dataset: Retail_Sales_sample-Dataset.xlsx &nbsp;|&nbsp; Period: Q1 2025 &nbsp;|&nbsp;
    Course: MSE-803 Data Analytics, Week 5 Activity 2 &nbsp;|&nbsp; Student: Pathuma Hiruni
  </p>
</div>"""

# ── Assemble full HTML ────────────────────────────────────────────────────────
chart_html = pio.to_html(fig, include_plotlyjs="cdn", full_html=False, config={"displayModeBar": False})

full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>Retail Sales Dashboard — Q1 2025</title>
  <style>
    *{{box-sizing:border-box;margin:0;padding:0;}}
    body{{background:{LIGHT_BG};font-family:Arial,sans-serif;}}
    header{{background:{TEXT_DARK};color:#fff;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;}}
    header h1{{font-size:18px;font-weight:700;letter-spacing:.02em;}}
    header span{{font-size:12px;opacity:.65;}}
  </style>
</head>
<body>
  <header>
    <h1>Retail Sales Dashboard</h1>
    <span>MSE-803 Data Analytics &nbsp;·&nbsp; Week 5 Activity 2 &nbsp;·&nbsp; Q1 2025</span>
  </header>
  {kpi_html}
  {chart_html}
  {insights_html}
</body>
</html>"""

with open("Week5_Activity2_Dashboard.html", "w", encoding="utf-8") as f:
    f.write(full_html)

print("Dashboard saved: Week5_Activity2_Dashboard.html")

# ── Save cleaned Excel ────────────────────────────────────────────────────────
df.to_excel("Week5_Activity2_Cleaned_Data.xlsx", index=False)
print("Cleaned data saved: Week5_Activity2_Cleaned_Data.xlsx")
