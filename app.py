import pandas as pd
import numpy as np
from dash import Dash, html, dcc
import plotly.graph_objects as go
from plotly.subplots import make_subplots


# ============================================================
# CONFIG
# ============================================================
FILE_PATH = "SN_2026.xlsx"

BG = "#0D1117"
CARD_BG = "#161B22"
PANEL_BG = "#161B22"
BRADESCO_RED = "#CC092F"
GOLD = "#F0B429"
TEAL = "#2DD4BF"
LIGHT_BLUE = "#58A6FF"
SOFT_PURPLE = "#A78BFA"
SOFT_PINK = "#F472B6"
WHITE = "#E6EDF3"
GRAY = "#8B949E"
GRID = "rgba(139, 148, 158, 0.12)"

ISSUER_COLORS = {
    "Citibank": "#58A6FF",
    "BBVA": "#2DD4BF",
    "JPM": "#F0B429",
    "Bradesco": "#CC092F",
    "Santander": "#A78BFA",
}

STRUCTURE_COLORS = {
    "CLN": "#58A6FF",
    "Phoenix Autocall": "#2DD4BF",
    "Participation": "#F0B429",
    "Catapult": "#A78BFA",
    "DRA": "#F472B6",
    "Reverse Convertible": "#FF7B72",
    "Range Accrual": "#8B949E",
    "Other": "#6E7681",
}

STEERCO_COLORS = {
    "Banker": "#58A6FF",
    "Investor": "#2DD4BF",
    "Investments": "#F0B429",
    "Calendar": "#F472B6",
}

from datetime import datetime
today = datetime.now().strftime("%B %d, %Y")  # Ex: "May 27, 2026"


# ============================================================
# LOAD + PREP DATA
# ============================================================
def categorize_structure(value):
    s = str(value).lower()
    if "autocall" in s or "phoenix" in s:
        return "Phoenix Autocall"
    if "catapult" in s:
        return "Catapult"
    if "cln" in s:
        return "CLN"
    if "participation" in s:
        return "Participation"
    if "range accrual" in s:
        return "Range Accrual"
    if "reverse convertible" in s:
        return "Reverse Convertible"
    if "dra" in s:
        return "DRA"
    return "Other"


def money_m(x):
    return f"${x / 1_000_000:.1f}M"


def money_k(x):
    return f"${x / 1_000:.0f}K"


def money_full(x):
    return f"${x:,.0f}"


def pct(x):
    return f"{x:.2%}"


def load_data(path):
    df = pd.read_excel(path, engine="openpyxl")
    df["Trade Date"] = pd.to_datetime(df["Trade Date"], errors="coerce")
    
    # Clean whitespace from string columns
    df["Sub Asset Class"] = df["Sub Asset Class"].str.strip()

    if "Settle Date" in df.columns:
        df["Settle Date"] = pd.to_datetime(df["Settle Date"], errors="coerce")
    df["Volume"] = pd.to_numeric(df["Volume"], errors="coerce").fillna(0)
    df["Fee"] = pd.to_numeric(df["Fee"], errors="coerce").fillna(0)
    if "Fee $" not in df.columns:
        df["Fee $"] = df["Volume"] * df["Fee"]
    else:
        df["Fee $"] = pd.to_numeric(df["Fee $"], errors="coerce").fillna(df["Volume"] * df["Fee"])
    df["Structure Type"] = df["Strutcutre"].apply(categorize_structure)
    df["Month"] = df["Trade Date"].dt.to_period("M").dt.to_timestamp()
    df["Month Label"] = df["Month"].dt.strftime("%b %Y")
    return df


df = load_data(FILE_PATH)


# ============================================================
# AGGREGATIONS
# ============================================================
monthly = (
    df.groupby(["Month", "Month Label"], as_index=False)
    .agg(Volume=("Volume", "sum"), Trades=("Note", "count"), Fees=("Fee $", "sum"))
    .sort_values("Month")
)
monthly["Cumulative"] = monthly["Volume"].cumsum()


issuer = (
    df.groupby("Issuer", as_index=False)
    .agg(
        Volume=("Volume", "sum"),
        Fees=("Fee $", "sum"),
        Trades=("Note", "count"),
    )
)

issuer["Avg_Fee"] = issuer["Fees"] / issuer["Volume"]
issuer = issuer.sort_values("Volume", ascending=False)


structure = (
    df.groupby("Structure Type", as_index=False)
    .agg(Volume=("Volume", "sum"), Trades=("Note", "count"))
    .sort_values("Volume", ascending=False)
)

steerco_monthly = df.groupby(["Month Label", "Steerco"], as_index=False)["Volume"].sum()

top10 = (
    df.nlargest(10, "Volume")[["Strutcutre", "Issuer", "Volume", "Fee $", "Trade Date"]]
    .sort_values("Volume", ascending=True)
    .copy()
)
top10["Label"] = top10["Strutcutre"].apply(lambda x: x if len(str(x)) <= 70 else str(x)[:67] + "...")

issuer_structure = df.groupby(["Issuer", "Structure Type"], as_index=False)["Volume"].sum()
top_structures = structure["Structure Type"].head(5).tolist()
issuer_structure_top = issuer_structure[issuer_structure["Structure Type"].isin(top_structures)].copy()

# Sub Asset Class breakdown
sub_asset = (
    df.groupby("Sub Asset Class", as_index=False)
    .agg(
        Volume=("Volume", "sum"),
        Trades=("Note", "count"),
        Fees=("Fee $", "sum"),
    )
    .sort_values("Volume", ascending=False)
)

sub_asset["Avg_Fee"] = sub_asset["Fees"] / sub_asset["Volume"]
sub_asset["Pct"] = sub_asset["Volume"] / sub_asset["Volume"].sum() * 100

# KPIs
total_volume = df["Volume"].sum()
total_trades = df["Note"].count()
total_fees = df["Fee $"].sum()
avg_trade_size = df["Volume"].mean()
avg_fee_rate = df["Fee $"].sum() / df["Volume"].sum()

date_min = df["Trade Date"].min()
date_max = df["Trade Date"].max()
date_range_label = f"{date_min.strftime('%B')} \u2014 {date_max.strftime('%B %Y')} | Year-to-Date Performance"





# ============================================================
# COMMON FIGURE STYLING
# ============================================================
def base_layout(fig, title=None):
    fig.update_layout(
        title=dict(text=title or "", x=0.01, xanchor="left", font=dict(size=20, color=WHITE, family="Arial Black")),
        paper_bgcolor=PANEL_BG,
        plot_bgcolor=PANEL_BG,
        font=dict(color=WHITE, family="Arial"),
        margin=dict(l=40, r=30, t=70, b=40),
        hoverlabel=dict(bgcolor=CARD_BG, font_color=WHITE),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=WHITE), orientation="v"),
    )
    return fig

# ============================================================
# UNDERLYING CONCENTRATION (reads from Excel columns)
# ============================================================
und_cols = ["Underlying 1", "Underlying 2", "Underlying 3"]

rows = []
for _, trade in df.iterrows():
    for col in und_cols:
        if col in df.columns:
            val = trade.get(col)
            if pd.notna(val) and str(val).strip():
                rows.append({
                    "Underlying": str(val).strip(),
                    "Volume": trade["Volume"],
                })

und_df = pd.DataFrame(rows)

underlying = (
    und_df.groupby("Underlying", as_index=False)
    .agg(
        Appearances=("Volume", "count"),
        Total_Volume=("Volume", "sum"),
    )
    .sort_values("Total_Volume", ascending=False)
)
underlying["Pct"] = underlying["Total_Volume"] / underlying["Total_Volume"].sum() * 100
underlying["Cumulative_Pct"] = underlying["Pct"].cumsum()

top5_concentration = underlying.head(5)["Pct"].sum()
n_underlyings = len(underlying)

# --- Top 10 Bar Chart ---
top_und = underlying.head(10).sort_values("Total_Volume", ascending=True)

max_vol = top_und["Total_Volume"].max()
bar_colors = [
    f"rgba(88, 166, 255, {0.4 + 0.6 * (v / max_vol)})"
    for v in top_und["Total_Volume"]
]

fig_underlying = go.Figure(
    go.Bar(
        x=top_und["Total_Volume"] / 1_000_000,
        y=top_und["Underlying"],
        orientation="h",
        marker=dict(color=bar_colors, line=dict(color=PANEL_BG, width=1.5)),
        text=[
            f"  {money_m(v)}  \u00b7  {int(a)} trades  \u00b7  {p:.1f}%"
            for v, a, p in zip(top_und["Total_Volume"], top_und["Appearances"], top_und["Pct"])
        ],
        textposition="outside",
        textfont=dict(size=12, color=WHITE),
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Volume: $%{x:.2f}M<br>"
            "Trades: %{customdata[0]}<br>"
            "Share: %{customdata[1]:.1f}%<br>"
            "Cumulative: %{customdata[2]:.1f}%"
            "<extra></extra>"
        ),
        customdata=np.stack(
            [top_und["Appearances"], top_und["Pct"], top_und["Cumulative_Pct"]],
            axis=-1,
        ),
    )
)

fig_underlying.update_yaxes(color=WHITE, tickfont=dict(size=13, color=WHITE))
fig_underlying.update_xaxes(title="Volume ($M)", showgrid=True, gridcolor=GRID, zeroline=False, color=GRAY)
fig_underlying.update_layout(height=450, margin=dict(l=120, r=200, t=70, b=40))
base_layout(fig_underlying, "\u25c6 Top 10 Underlyings by Exposure")


# --- Donut: Top 10 vs Others ---
top10_vol = underlying.head(10)["Total_Volume"].sum()
others_vol = underlying.iloc[10:]["Total_Volume"].sum()

donut_labels = underlying.head(10)["Underlying"].tolist() + ["Others"]
donut_values = underlying.head(10)["Total_Volume"].tolist() + [others_vol]
donut_colors = [
    LIGHT_BLUE, TEAL, GOLD, SOFT_PURPLE, BRADESCO_RED,
    "#79C0FF", "#56D364", "#FFA657", "#FF7B72", "#D2A8FF", GRAY,
]

fig_und_donut = go.Figure(
    data=[
        go.Pie(
            labels=donut_labels,
            values=donut_values,
            hole=0.58,
            marker=dict(
                colors=donut_colors[:len(donut_labels)],
                line=dict(color=PANEL_BG, width=3),
            ),
            textinfo="percent",
            textfont=dict(size=11, color=WHITE),
            hovertemplate="<b>%{label}</b><br>Volume: $%{value:,.0f}<br>Share: %{percent}<extra></extra>",
        )
    ]
)

fig_und_donut.add_annotation(
    text=f"<b>{n_underlyings}</b><br><span style='font-size:12px;color:{GRAY}'>Underlyings</span>",
    x=0.5, y=0.5, showarrow=False, font=dict(color=WHITE, size=24),
)

fig_und_donut.update_layout(
    height=450,
    legend=dict(
        font=dict(size=11, color=WHITE),
        bgcolor="rgba(0,0,0,0)",
        orientation="v",
        yanchor="middle", y=0.5,
        xanchor="left", x=1.05,
    ),
)
base_layout(fig_und_donut, "\u25c6 Underlying Concentration Mix")
# ============================================================
# FIGURES
# ============================================================

# 1) Monthly Volume + Cumulative Trend
fig_monthly = make_subplots(specs=[[{"secondary_y": True}]])

fig_monthly.add_trace(
    go.Bar(
        x=monthly["Month Label"],
        y=monthly["Volume"] / 1_000_000,
        name="Monthly Volume",
        marker=dict(color=LIGHT_BLUE, line=dict(color=PANEL_BG, width=1.5)),
        text=[money_m(v) for v in monthly["Volume"]],
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Volume: %{text}<br>Trades: %{customdata}<extra></extra>",
        customdata=monthly["Trades"],
    ),
    secondary_y=False,
)

fig_monthly.add_trace(
    go.Scatter(
        x=monthly["Month Label"],
        y=monthly["Cumulative"] / 1_000_000,
        name="Cumulative",
        mode="lines+markers+text",
        line=dict(color=BRADESCO_RED, width=4),
        marker=dict(size=10, color=BRADESCO_RED, line=dict(color=WHITE, width=1)),
        text=[money_m(v) for v in monthly["Cumulative"]],
        textposition="top center",
        hovertemplate="<b>%{x}</b><br>Cumulative: %{text}<extra></extra>",
    ),
    secondary_y=True,
)

fig_monthly.update_yaxes(title_text="Monthly Volume ($M)", secondary_y=False, showgrid=True, gridcolor=GRID, zeroline=False, color=GRAY)
fig_monthly.update_yaxes(title_text="Cumulative Volume ($M)", secondary_y=True, showgrid=False, zeroline=False, color=BRADESCO_RED)
fig_monthly.update_xaxes(color=GRAY)
base_layout(fig_monthly, "\u25c6 Monthly Volume & Cumulative Trend")


# 2) Donut by Issuer
fig_issuer = go.Figure(
    data=[
        go.Pie(
            labels=issuer["Issuer"],
            values=issuer["Volume"],
            hole=0.58,
            marker=dict(colors=[ISSUER_COLORS.get(i, GRAY) for i in issuer["Issuer"]], line=dict(color=PANEL_BG, width=3)),
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>Volume: $%{value:,.0f}<br>Share: %{percent}<extra></extra>",
        )
    ]
)
fig_issuer.add_annotation(
    text=f"<b>{money_m(total_volume)}</b><br><span style='font-size:12px;color:{GRAY}'>Total Volume</span>",
    x=0.5, y=0.5, showarrow=False, font=dict(color=WHITE, size=22),
)
base_layout(fig_issuer, "\u25c6 Volume by Issuer")


# 3) Donut by Structure Type
fig_structure = go.Figure(
    data=[
        go.Pie(
            labels=structure["Structure Type"],
            values=structure["Volume"],
            hole=0.58,
            marker=dict(colors=[STRUCTURE_COLORS.get(s, GRAY) for s in structure["Structure Type"]], line=dict(color=PANEL_BG, width=3)),
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>Volume: $%{value:,.0f}<br>Share: %{percent}<extra></extra>",
        )
    ]
)
fig_structure.add_annotation(
    text=f"<b>{total_trades}</b><br><span style='font-size:12px;color:{GRAY}'>Total Trades</span>",
    x=0.5, y=0.5, showarrow=False, font=dict(color=WHITE, size=24),
)
base_layout(fig_structure, "\u25c6 Volume by Structure Type")


# 4) Stacked Bar by Steerco  --- FIXED: reindex only the Volume series, not the full DataFrame
fig_steerco = go.Figure()
month_order = monthly["Month Label"].tolist()

for cat in ["Banker", "Investor", "Investments", "Calendar"]:
    vol_series = (
        steerco_monthly[steerco_monthly["Steerco"] == cat]
        .set_index("Month Label")["Volume"]
        .reindex(month_order, fill_value=0)
    )
    fig_steerco.add_trace(
        go.Bar(
            x=vol_series.index.tolist(),
            y=vol_series.values / 1_000_000,
            name=cat,
            marker=dict(color=STEERCO_COLORS.get(cat, GRAY)),
            hovertemplate="<b>%{x}</b><br>" + cat + ": %{y:.2f}M<extra></extra>",
        )
    )

fig_steerco.update_layout(barmode="stack")
fig_steerco.update_yaxes(title="Volume ($M)", showgrid=True, gridcolor=GRID, zeroline=False, color=GRAY)
fig_steerco.update_xaxes(color=GRAY)
base_layout(fig_steerco, "\u25c6 Monthly Volume by Sponsor")


# 5) Top 10 largest trades by volume
fig_top10 = go.Figure(
    go.Bar(
        x=top10["Volume"] / 1_000_000,
        y=top10["Label"],
        orientation="h",
        marker=dict(color=[ISSUER_COLORS.get(i, GRAY) for i in top10["Issuer"]], line=dict(color=PANEL_BG, width=1.5)),
        text=[
            f"{money_m(v)} \u00b7 {issuer_name} \u00b7 Fee: {money_k(fee)}"
            for v, issuer_name, fee in zip(top10["Volume"], top10["Issuer"], top10["Fee $"])
        ],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Volume: %{x:.2f}M<extra></extra>",
        customdata=np.stack([top10["Issuer"], top10["Fee $"], top10["Trade Date"].dt.strftime("%Y-%m-%d")], axis=-1),
    )
)

fig_top10.update_yaxes(color=WHITE)
fig_top10.update_xaxes(title="Volume ($M)", showgrid=True, gridcolor=GRID, zeroline=False, color=GRAY)
base_layout(fig_top10, "\u25c6 Top 10 Largest Trades by Volume")


# 6) Average Fee Rate by Issuer
issuer_fee_sorted = issuer.sort_values("Avg_Fee", ascending=True).copy()

fig_fee = go.Figure(
    go.Bar(
        x=issuer_fee_sorted["Avg_Fee"] * 100,
        y=issuer_fee_sorted["Issuer"],
        orientation="h",
        marker=dict(color=[ISSUER_COLORS.get(i, GRAY) for i in issuer_fee_sorted["Issuer"]], line=dict(color=PANEL_BG, width=1.5)),
        text=[f"{x:.2f}%" for x in issuer_fee_sorted["Avg_Fee"] * 100],
        textposition="outside",
        customdata=np.stack([issuer_fee_sorted["Fees"]], axis=-1),
        hovertemplate="<b>%{y}</b><br>Avg Fee Rate: %{x:.2f}%<extra></extra>",
    )
)

fig_fee.update_xaxes(title="Average Fee Rate (%)", showgrid=True, gridcolor=GRID, zeroline=False, color=GRAY)
fig_fee.update_yaxes(color=WHITE)
base_layout(fig_fee, "\u25c6 Average Fee Rate by Issuer")


# 7) Grouped Bar: Volume by Issuer & Structure Type  --- FIXED: reindex only Volume series
fig_issuer_structure = go.Figure()

# 8) Sub Asset Class Breakdown — Horizontal Bar
sub_asset_sorted = sub_asset.sort_values("Volume", ascending=True).copy()

SUB_ASSET_COLORS = {
    "US Equities": "#58A6FF",
    "Emerging Markets Bonds": "#2DD4BF",
    "Emerging Markets Equities": "#2DD4BF",
    "Global Fixed Income": "#F0B429",
    "Commodities": "#A78BFA",
    "Investment Grade": "#FF7B72",
    "High Yield": "#79C0FF",
}

fig_sub_asset = go.Figure(
    go.Bar(
        x=sub_asset_sorted["Volume"] / 1_000_000,
        y=sub_asset_sorted["Sub Asset Class"],
        orientation="h",
        marker=dict(
            color=[SUB_ASSET_COLORS.get(s, GRAY) for s in sub_asset_sorted["Sub Asset Class"]],
            line=dict(color=PANEL_BG, width=1.5),
        ),
        text=[
            f"{money_m(v)}  ·  {t} trades  ·  {p:.1f}%"
            for v, t, p in zip(
                sub_asset_sorted["Volume"],
                sub_asset_sorted["Trades"],
                sub_asset_sorted["Pct"],
            )
        ],
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Volume: %{x:.2f}M<br>"
            "Trades: %{customdata[0]}<br>"
            "Avg Fee: %{customdata[1]:.2f}%<br>"
            "<extra></extra>"
        ),
        customdata=np.stack(
            [sub_asset_sorted["Trades"], sub_asset_sorted["Avg_Fee"] * 100],
            axis=-1,
        ),
    )
)

fig_sub_asset.update_yaxes(color=WHITE)
fig_sub_asset.update_xaxes(
    title="Volume ($M)",
    showgrid=True,
    gridcolor=GRID,
    zeroline=False,
    color=GRAY,
)
base_layout(fig_sub_asset, "\u25c6 Volume by Sub Asset Class")
for struct_name in top_structures:
    vol_series = (
        issuer_structure_top[issuer_structure_top["Structure Type"] == struct_name]
        .set_index("Issuer")["Volume"]
        .reindex(issuer["Issuer"], fill_value=0)
    )
    fig_issuer_structure.add_trace(
        go.Bar(
            x=vol_series.index.tolist(),
            y=vol_series.values / 1_000_000,
            name=struct_name,
            marker=dict(color=STRUCTURE_COLORS.get(struct_name, GRAY)),
            hovertemplate="<b>%{x}</b><br>" + struct_name + ": %{y:.2f}M<extra></extra>",
        )
    )

fig_issuer_structure.update_layout(barmode="group")
fig_issuer_structure.update_xaxes(color=WHITE)
fig_issuer_structure.update_yaxes(title="Volume ($M)", showgrid=True, gridcolor=GRID, zeroline=False, color=GRAY)
base_layout(fig_issuer_structure, "\u25c6 Volume by Issuer & Structure Type")


fig_sub_asset.update_yaxes(color=WHITE)
fig_sub_asset.update_xaxes(
    title="Volume ($M)",
    showgrid=True,
    gridcolor=GRID,
    zeroline=False,
    color=GRAY,
)
base_layout(fig_sub_asset, "\u25c6 Volume by Sub Asset Class")

# Donut: Sub Asset Class
fig_sub_asset_donut = go.Figure(
    data=[
        go.Pie(
            labels=sub_asset["Sub Asset Class"],
            values=sub_asset["Volume"],
            hole=0.58,
            marker=dict(
                colors=[SUB_ASSET_COLORS.get(s, GRAY) for s in sub_asset["Sub Asset Class"]],
                line=dict(color=PANEL_BG, width=3),
            ),
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>Volume: $%{value:,.0f}<br>Share: %{percent}<extra></extra>",
        )
    ]
)
fig_sub_asset_donut.add_annotation(
    text=f"<b>{len(sub_asset)}</b><br><span style='font-size:12px;color:{GRAY}'>Sub Classes</span>",
    x=0.5, y=0.5, showarrow=False, font=dict(color=WHITE, size=22),
)
base_layout(fig_sub_asset_donut, "\u25c6 Sub Asset Class Mix")


# ============================================================
# SANKEY: Issuer → Structure Type
# ============================================================
sankey_links = df.groupby(["Issuer", "Structure Type"], as_index=False)["Volume"].sum()
sankey_links = sankey_links[sankey_links["Volume"] > 0]

issuer_order = df.groupby("Issuer")["Volume"].sum().sort_values(ascending=False).index.tolist()
struct_order = df.groupby("Structure Type")["Volume"].sum().sort_values(ascending=False).index.tolist()

sankey_labels = issuer_order + struct_order
sankey_idx = {label: i for i, label in enumerate(sankey_labels)}

sankey_node_colors = []
for label in sankey_labels:
    if label in ISSUER_COLORS:
        sankey_node_colors.append(ISSUER_COLORS[label])
    elif label in STRUCTURE_COLORS:
        sankey_node_colors.append(STRUCTURE_COLORS[label])
    else:
        sankey_node_colors.append(GRAY)

sankey_node_vols = {}
for _, row in df.iterrows():
    for col in ["Issuer", "Structure Type"]:
        sankey_node_vols[row[col]] = sankey_node_vols.get(row[col], 0) + row["Volume"]

sankey_display = [
    f"{label}  \u00b7  ${sankey_node_vols.get(label, 0)/1e6:.1f}M"
    for label in sankey_labels
]

def hex_to_rgba(hex_color, alpha=0.45):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r}, {g}, {b}, {alpha})"

sk_src, sk_tgt, sk_val, sk_clr = [], [], [], []
for _, row in sankey_links.iterrows():
    sk_src.append(sankey_idx[row["Issuer"]])
    sk_tgt.append(sankey_idx[row["Structure Type"]])
    sk_val.append(row["Volume"])
    sk_clr.append(hex_to_rgba(ISSUER_COLORS.get(row["Issuer"], GRAY)))

fig_sankey = go.Figure(data=[go.Sankey(
    arrangement="snap",
    node=dict(
        pad=30, thickness=35,
        line=dict(color=PANEL_BG, width=2),
        label=sankey_display,
        color=sankey_node_colors,
        hovertemplate="<b>%{label}</b><extra></extra>",
    ),
    link=dict(
        source=sk_src, target=sk_tgt, value=sk_val, color=sk_clr,
        hovertemplate="<b>%{source.label}</b> \u2192 <b>%{target.label}</b><br>Volume: $%{value:,.0f}<extra></extra>",
    ),
)])

fig_sankey.update_layout(height=500, margin=dict(l=30, r=30, t=100, b=40))
fig_sankey.add_annotation(x=0.0, y=1.10, text="<b>ISSUER</b>", showarrow=False, font=dict(size=16, color=GRAY), xref="paper", yref="paper")
fig_sankey.add_annotation(x=1.0, y=1.10, text="<b>STRUCTURE TYPE</b>", showarrow=False, font=dict(size=16, color=GRAY), xref="paper", yref="paper")
base_layout(fig_sankey, "\u25c6 Trade Flow: Issuer \u2192 Structure Type")

# ============================================================
# DASH APP
# ============================================================
app = Dash(__name__, title="Structured Notes Analytics - Bradesco")
server = app.server

card_style = {
    "backgroundColor": CARD_BG,
    "borderRadius": "18px",
    "padding": "18px 20px",
    "border": "1px solid rgba(255,255,255,0.06)",
    "boxShadow": "0 10px 30px rgba(0,0,0,0.20)",
    "flex": "1",
    "minWidth": "180px",
}

panel_style = {
    "backgroundColor": PANEL_BG,
    "borderRadius": "20px",
    "padding": "14px",
    "border": "1px solid rgba(255,255,255,0.06)",
    "boxShadow": "0 10px 30px rgba(0,0,0,0.20)",
}


def kpi_card(title, value, accent):
    return html.Div(
        [
            html.Div(
                style={
                    "height": "6px",
                    "width": "100%",
                    "backgroundColor": accent,
                    "borderRadius": "999px",
                    "marginBottom": "14px",
                },
                className="accent-bar-glow",
            ),
            html.Div(
                value,
                style={
                    "fontSize": "30px",
                    "fontWeight": "800",
                    "color": WHITE,
                    "marginBottom": "6px",
                },
                className="kpi-value",
            ),
            html.Div(
                title,
                style={
                    "fontSize": "14px",
                    "color": GRAY,
                    "fontWeight": "600",
                    "letterSpacing": "0.2px",
                },
            ),
        ],
        className="glass-kpi kpi-animate",
        style={
            "padding": "18px 20px",
            "flex": "1",
            "minWidth": "180px",
        },
    )


app.layout = html.Div(
    style={"backgroundColor": BG, "minHeight": "100vh", "padding": "24px", "fontFamily": "Arial, sans-serif"},
    children=[
# Header
        html.Div(
            [
                html.Div(
                    "STRUCTURED NOTES ANALYTICS",
                    style={"textAlign": "center", "fontSize": "42px", "fontWeight": "900", "color": WHITE, "letterSpacing": "1px", "textTransform": "uppercase"},
                ),
                html.Div(
                    "Bradesco Investments",
                    style={"textAlign": "center", "fontSize": "24px", "fontStyle": "italic", "color": BRADESCO_RED, "marginTop": "6px"},
                ),
                html.Div(
                    date_range_label,
                    style={"textAlign": "center", "fontSize": "14px", "color": GRAY, "marginTop": "8px", "marginBottom": "14px"},
                ),
                html.Div(
                    style={"height": "4px", "width": "70%", "margin": "0 auto 28px auto", "backgroundColor": BRADESCO_RED, "borderRadius": "999px"},
                    className="header-line",
                ),
            ],
            className="header-animate",
        ),
        # KPI Row
        html.Div(
            [
                kpi_card("Total Volume", money_m(total_volume), LIGHT_BLUE),
                kpi_card("Total Trades", f"{total_trades}", TEAL),
                kpi_card("Fee Revenue", f"${total_fees / 1_000_000:.2f}M", GOLD),
                kpi_card("Avg Trade Size", money_k(avg_trade_size), SOFT_PURPLE),
                kpi_card("Avg Fee Rate", pct(avg_fee_rate), BRADESCO_RED),
                kpi_card("Top 5 Concentration", f"{top5_concentration:.1f}%", SOFT_PINK),
            ],
            style={"display": "flex", "gap": "18px", "flexWrap": "wrap", "marginBottom": "24px"},
        ),
        # Row 1
        html.Div(
            [
                html.Div(dcc.Graph(figure=fig_monthly, config={"displayModeBar": False}), style={**panel_style, "flex": "1"}),
                html.Div(dcc.Graph(figure=fig_issuer, config={"displayModeBar": False}), style={**panel_style, "flex": "1"}),
            ],
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "20px", "marginBottom": "20px"},
        ),
        # Row 2
        html.Div(
            [
                html.Div(dcc.Graph(figure=fig_structure, config={"displayModeBar": False}), style={**panel_style, "flex": "1"}),
                html.Div(dcc.Graph(figure=fig_steerco, config={"displayModeBar": False}), style={**panel_style, "flex": "1"}),
            ],
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "20px", "marginBottom": "20px"},
        ),
        


        # Row 3 - full width
        html.Div(dcc.Graph(figure=fig_top10, config={"displayModeBar": False}), style={**panel_style, "marginBottom": "20px"}),
        # Row 4
        html.Div(
            [
                html.Div(dcc.Graph(figure=fig_fee, config={"displayModeBar": False}), style={**panel_style, "flex": "1"}),
                html.Div(dcc.Graph(figure=fig_issuer_structure, config={"displayModeBar": False}), style={**panel_style, "flex": "1"}),
            ],
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "20px", "marginBottom": "20px"},
        ),
        # Row 7 - Sankey (full width)
        html.Div(
            dcc.Graph(figure=fig_sankey, config={"displayModeBar": False}),
            style={**panel_style, "marginBottom": "20px"},
        ),
        
# Row 5 - Sub Asset Class
        html.Div(
            [
                html.Div(
                    dcc.Graph(figure=fig_sub_asset, config={"displayModeBar": False}),
                    style={**panel_style, "flex": "1"},
                ),
                html.Div(
                    dcc.Graph(figure=fig_sub_asset_donut, config={"displayModeBar": False}),
                    style={**panel_style, "flex": "1"},
                ),
            ],
            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "20px", "marginBottom": "20px"},
        ),
   # Row 6 - Underlying Concentration
        html.Div(
            [
                html.Div(
                    dcc.Graph(figure=fig_underlying, config={"displayModeBar": False}),
                    style={**panel_style, "flex": "3"},
                ),
                html.Div(
                    dcc.Graph(figure=fig_und_donut, config={"displayModeBar": False}),
                    style={**panel_style, "flex": "2"},
                ),
            ],
            style={
                "display": "grid",
                "gridTemplateColumns": "3fr 2fr",
                "gap": "20px",
                "marginBottom": "20px",
            },
        ),
            # Footer
        

html.Div(
    [
        html.Div(
            style={
                "height": "4px",
                "width": "70%",
                "margin": "18px auto 14px auto",
                "backgroundColor": BRADESCO_RED,
                "borderRadius": "999px",
            }
        ),
        html.Div(
            f"Portfolio Solutions | Data as of {today} | Confidential",
            style={
                "textAlign": "center",
                "fontSize": "13px",
                "color": GRAY,
                "fontStyle": "italic",
            },
        ),
    ]
),
    ],
)


if __name__ == "__main__":
    app.run(debug=True)
