"""Portfolio dashboard over Sales_Master.xlsx (Part 5).

    uv run streamlit run dashboard.py

Everything is read from the master that process_reports.py maintains, and the
alerts come from the same process_reports.analyse used for the weekly brief.
Flagged values are blank in the master, so they never reach a chart.
"""

from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from process_reports import SWING_ALERT_POINTS, TREND_WEEKS, analyse, load_directory, load_master

HERE = Path(__file__).resolve().parent
MASTER = HERE / "Sales_Master.xlsx"
BRIEF = HERE / "weekly_ops_brief.txt"

# Categorical palette, fixed order. A store keeps its colour whatever is filtered.
SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
UP, DOWN, MUTED = "#2a78d6", "#e34948", "#898781"

st.set_page_config(page_title="Bluebonnet Sales", layout="wide")


@st.cache_data
def load(modified: float) -> tuple[pd.DataFrame, dict[str, str]]:
    """Cached until the master file changes (``modified`` is its mtime)."""
    return load_master(MASTER), load_directory(HERE / "Store_Directory.csv")


def money(value: float) -> str:
    # Escaped: Streamlit markdown treats text between two "$" as a formula.
    return f"\\${value:,.0f}"


if not MASTER.exists():
    st.error("Sales_Master.xlsx not found. Run `uv run process_reports.py Input_Files` first.")
    st.stop()

master, names = load(MASTER.stat().st_mtime)
analysis = analyse(master)
all_stores = sorted(master["Store Code"].unique())
store_colors = alt.Scale(domain=all_stores, range=SERIES_COLORS[: len(all_stores)])


def label(code: str) -> str:
    return f"{code} {names.get(code, '')}".strip()


# --- Sidebar -----------------------------------------------------------------

with st.sidebar:
    st.header("Filters")
    stores = st.multiselect("Stores", all_stores, default=all_stores, format_func=label)
    weeks = sorted(master["Week Ending"].dt.date.unique())
    start, end = st.select_slider("Chart date range", options=weeks, value=(weeks[0], weeks[-1]))
    st.divider()
    st.download_button("Download Sales_Master.xlsx", MASTER.read_bytes(), MASTER.name,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    if BRIEF.exists():
        st.download_button("Download weekly ops brief", BRIEF.read_bytes(), BRIEF.name, "text/plain", width="stretch")
    st.caption(f"{len(master)} store-weeks in {MASTER.name}, latest week ending {analysis.week.date()}.")

if not stores:
    st.info("Pick at least one store.")
    st.stop()

view = master[master["Store Code"].isin(stores) & master["Week Ending"].dt.date.between(start, end)]
snapshot = analysis.snapshot[analysis.snapshot["Store Code"].isin(stores)].assign(Store=lambda d: d["Store Code"].map(label))
comparable = snapshot.dropna(subset=["Swing (pts)"])
urgent = comparable[comparable["Swing (pts)"].abs() > SWING_ALERT_POINTS]
trends = snapshot.dropna(subset=["Trend"])

# --- Headline ----------------------------------------------------------------

st.title("Bluebonnet Blend Co. - Weekly Sales")
scope = "all stores" if len(stores) == len(all_stores) else f"{len(stores)} of {len(all_stores)} stores"
st.caption(f"Week ending {analysis.week.date()} | {scope} | % changes are vs the same period last year")

latest = master[(master["Week Ending"] == analysis.week) & master["Store Code"].isin(stores)]


def portfolio(period: str) -> tuple[float, float]:
    """Total net sales and its % change vs LY. LY dollars are not stored, but LY = TY / (1 + %chg)."""
    rows = latest.dropna(subset=[f"{period} Net Sales", f"{period} Net Sales % Chg"])
    now = rows[f"{period} Net Sales"].sum()
    last_year = (rows[f"{period} Net Sales"] / (1 + rows[f"{period} Net Sales % Chg"] / 100)).sum()
    return now, (now / last_year - 1) * 100 if last_year else float("nan")


tiles = st.columns(5)
for tile, period, title in ((tiles[0], "28D", "28-day net sales"), (tiles[1], "7D", "7-day net sales")):
    now, change = portfolio(period)
    tile.metric(title, f"${now:,.0f}", f"{change:+.1f}% vs LY")
tiles[2].metric("Improving week over week", f"{(comparable['Swing (pts)'] > 0).sum()} of {len(snapshot)}",
                help="Stores whose 7-day % change vs LY rose compared with the previous week")
tiles[3].metric(f"Swings over {SWING_ALERT_POINTS:g} pts", len(urgent),
                help="Week-over-week move in 7-day % change vs LY bigger than 5 points either way")
tiles[4].metric(f"At a {TREND_WEEKS}-week low", int((trends["Trend"] == "LOW").sum()),
                help="28-day net sales at the lowest point of the trailing 26 weeks")

alerts = []
for _, row in urgent.sort_values("Swing (pts)").iterrows():
    arrow = "▼" if row["Swing (pts)"] < 0 else "▲"
    alerts.append(f"{arrow} **{row['Store']}** swung **{row['Swing (pts)']:+.1f} pts** week over week "
                  f"(7-day vs LY {row['Previous 7D Net Sales % Chg']:+.1f}% → {row['7D Net Sales % Chg']:+.1f}%)")
for _, row in trends.sort_values("Trend").iterrows():
    kind = "high" if row["Trend"] == "HIGH" else "low"
    arrow = "▲" if kind == "high" else "▼"
    alerts.append(f"{arrow} **{row['Store']}** is at a {row['Trend Weeks']}-week **{kind}** in 28-day net sales: "
                  f"{money(row['28D Net Sales'])} (previous {kind} {money(row[f'Previous {kind.title()}'])})")
with st.container(border=True):
    st.markdown("**Needs attention this week**")
    st.markdown("\n".join(f"- {alert}" for alert in alerts) if alerts else "Nothing unusual this week.")

# --- 28D Net Sales over time (the required chart) ----------------------------

st.subheader("28-day net sales by store")
# Re-index to every week so a missing report shows as a gap, not a line drawn across it.
full_weeks = pd.date_range(view["Week Ending"].min(), view["Week Ending"].max(), freq="7D")
wide = (
    view.set_index(["Store Code", "Week Ending"])["28D Net Sales"]
    .unstack("Store Code").reindex(full_weeks).rename_axis("Week Ending")
)
long = wide.reset_index().melt("Week Ending", var_name="Store", value_name="28D Net Sales")
long["Store name"] = long["Store"].map(label)
# Each store's latest point, labelled with its value and any 26-week high/low, for the per-store panels.
status = snapshot.set_index("Store Code")["Trend"].to_dict()
latest_week = long.dropna(subset=["28D Net Sales"]).groupby("Store")["Week Ending"].max()
long["Latest"] = long["Week Ending"].eq(long["Store"].map(latest_week)) & long["28D Net Sales"].notna()
# A falling line approaches its last point from above, so its label goes below the point instead.
falling = long.dropna(subset=["28D Net Sales"]).groupby("Store")["28D Net Sales"].apply(lambda s: len(s) > 1 and s.iloc[-1] < s.iloc[-2])
long["Falling"] = long["Store"].map(falling).fillna(False).astype(bool)
long["Label"] = [
    f"${value / 1000:,.1f}K" + {"HIGH": " ▲ 26-wk high", "LOW": " ▼ 26-wk low"}.get(status.get(store), "") if is_latest else ""
    for store, value, is_latest in zip(long["Store"], long["28D Net Sales"], long["Latest"])
]
money_axis = alt.Axis(format="$,.0f", title=None)

overlay_tab, multiples_tab = st.tabs(["All stores on one chart", "One chart per store"])
with overlay_tab:
    st.caption("Click a store in the legend to highlight it (shift-click for several). Hover for that week's figures.")
    pick = alt.selection_point(fields=["Store"], bind="legend")
    hover = alt.selection_point(nearest=True, on="pointerover", fields=["Week Ending"], empty=False)
    base = alt.Chart(long).encode(x=alt.X("Week Ending:T", title=None))
    lines = base.mark_line(strokeWidth=2).encode(
        y=alt.Y("28D Net Sales:Q", scale=alt.Scale(zero=False), axis=money_axis),
        color=alt.Color("Store:N", scale=store_colors, legend=alt.Legend(title=None, orient="top")),
        opacity=alt.condition(pick, alt.value(1), alt.value(0.12)),
    ).add_params(pick)
    dots = base.mark_point(filled=True, size=60).encode(
        y="28D Net Sales:Q", color=alt.Color("Store:N", scale=store_colors, legend=None),
        opacity=alt.condition(hover, alt.value(1), alt.value(0)),
    )
    rule = (
        alt.Chart(wide.reset_index()).mark_rule(color=MUTED)
        .encode(
            x="Week Ending:T",
            opacity=alt.condition(hover, alt.value(0.5), alt.value(0)),
            tooltip=[alt.Tooltip("Week Ending:T", title="Week ending")]
            + [alt.Tooltip(f"{code}:Q", format="$,.0f", title=label(code)) for code in wide.columns],
        )
        .add_params(hover)
    )
    st.altair_chart((lines + dots + rule).properties(height=400), width="stretch")

with multiples_tab:
    st.caption("Same scale on every chart. The dot is the latest week, labelled with its 28-day net sales.")
    # Every layer uses the faceted data, so each panel only draws its own store.
    panel_base = alt.Chart().encode(
        x=alt.X("Week Ending:T", title=None, axis=alt.Axis(format="%b", tickCount=6)),
        y=alt.Y("28D Net Sales:Q", scale=alt.Scale(zero=False), axis=alt.Axis(format="$.3~s", title=None, tickCount=4)),
        color=alt.Color("Store:N", scale=store_colors, legend=None),
    )
    panel = alt.layer(
        panel_base.mark_line(strokeWidth=2).encode(
            tooltip=[alt.Tooltip("Week Ending:T", title="Week ending"), alt.Tooltip("28D Net Sales:Q", format="$,.0f")]),
        panel_base.mark_point(filled=True, size=50, opacity=1).transform_filter("datum.Latest"),
        panel_base.mark_text(align="right", dx=-7, dy=-11, fontSize=11)
        .encode(text="Label:N", color=alt.value(MUTED)).transform_filter("datum.Latest && !datum.Falling"),
        panel_base.mark_text(align="right", dx=-7, dy=13, fontSize=11)
        .encode(text="Label:N", color=alt.value(MUTED)).transform_filter("datum.Latest && datum.Falling"),
        data=long,
    ).properties(width=195, height=140)
    faceted = panel.facet(
        facet=alt.Facet("Store name:N", title=None, header=alt.Header(labelFontWeight="bold", labelFontSize=12)),
        columns=4,
    )
    st.altair_chart(faceted, width="content")

if gaps := {code: missing for code, missing in analysis.missing_weeks.items() if code in stores}:
    st.caption("Gaps in a line are weeks with no readable report: " + "; ".join(f"{c} {', '.join(map(str, w))}" for c, w in gaps.items()))

# --- This week vs last year, and this week vs last week -----------------------

def store_bars(data: pd.DataFrame, field: str, title: str, unit: str, threshold: float | None = None) -> alt.Chart:
    """Horizontal bars, blue for up and red for down, labelled with their value."""
    data = data.dropna(subset=[field]).assign(Value=lambda d: d[field].map(lambda v: f"{v:+.1f}{unit}"))
    order = data.sort_values(field, ascending=False)["Store"].tolist()  # an explicit list survives layering
    # Pad the axis so a value label at the end of the longest bar never runs into the store names.
    low, high = min(data[field].min(), -(threshold or 0), 0), max(data[field].max(), threshold or 0, 0)
    pad = (high - low) * 0.22
    domain = [low - pad if low < 0 else 0, high + pad if high > 0 else 0]
    base = alt.Chart(data).encode(
        y=alt.Y("Store:N", sort=order, title=None, axis=alt.Axis(labelLimit=220)),
        x=alt.X(f"{field}:Q", title=title, scale=alt.Scale(domain=domain, nice=False)),
        tooltip=["Store:N", alt.Tooltip("Value:N", title=title)],
    )
    color = alt.condition(f"datum['{field}'] >= 0", alt.value(UP), alt.value(DOWN))
    if threshold is None:
        bars = base.mark_bar(cornerRadiusEnd=4, height=16).encode(color=color)
    else:
        beyond = f"abs(datum['{field}']) > {threshold}"
        bars = base.mark_bar(cornerRadiusEnd=4, height=16).encode(color=color, opacity=alt.condition(beyond, alt.value(1), alt.value(0.35)))
    # Explicit grey: Vega's default black text disappears in dark mode.
    labels = (
        base.mark_text(align="left", dx=4, color=MUTED).encode(text="Value:N").transform_filter(f"datum['{field}'] >= 0")
        + base.mark_text(align="right", dx=-4, color=MUTED).encode(text="Value:N").transform_filter(f"datum['{field}'] < 0")
    )
    layers = [bars, labels, alt.Chart(pd.DataFrame({"x": [0]})).mark_rule(color=MUTED).encode(x="x:Q")]
    if threshold is not None:
        layers.append(alt.Chart(pd.DataFrame({"x": [-threshold, threshold]})).mark_rule(color=MUTED, strokeDash=[4, 3]).encode(x="x:Q"))
    return alt.layer(*layers).properties(height=36 * len(data) + 30)


left, right = st.columns(2, gap="large")
with left:
    st.subheader("28-day net sales vs last year")
    st.caption("Ranked best to worst. Hover a bar for the store's figure.")
    st.altair_chart(store_bars(snapshot, "28D Net Sales % Chg", "% change vs LY", "%"), width="stretch")
with right:
    st.subheader("Week-over-week swing")
    st.caption(f"Change in 7-day % vs LY since last week. Solid bars crossed the ±{SWING_ALERT_POINTS:g} pt alert line (dashed).")
    st.altair_chart(store_bars(snapshot, "Swing (pts)", "points", " pts", SWING_ALERT_POINTS), width="stretch")

# --- Scorecard ----------------------------------------------------------------

st.subheader("Store scorecard")


def flags(row: pd.Series) -> str:
    parts = []
    if row["Trend"] == "HIGH":
        parts.append(f"▲ {row['Trend Weeks']}-wk high")
    elif row["Trend"] == "LOW":
        parts.append(f"▼ {row['Trend Weeks']}-wk low")
    if pd.notna(row["Swing (pts)"]) and abs(row["Swing (pts)"]) > SWING_ALERT_POINTS:
        parts.append(f"⚠ swing {row['Swing (pts)']:+.1f}")
    return "   ".join(parts)


card = snapshot.sort_values("28D Net Sales % Chg", ascending=False).assign(Flags=lambda d: d.apply(flags, axis=1))
st.dataframe(
    card[["Store", "28D Net Sales", "28D Net Sales % Chg", "28D Transactions % Chg", "28D Avg Ticket % Chg", "28D Driver",
          "7D Net Sales % Chg", "Swing (pts)", "Flags"]],
    hide_index=True,
    width="stretch",
    column_config={
        "Store": st.column_config.TextColumn("Store", width="medium"),
        "28D Net Sales": st.column_config.NumberColumn("28-day sales", format="$%,.0f"),
        "28D Net Sales % Chg": st.column_config.NumberColumn("28-day vs LY", format="%+.1f%%"),
        "28D Transactions % Chg": st.column_config.NumberColumn("Traffic vs LY", format="%+.1f%%", help="28-day transactions"),
        "28D Avg Ticket % Chg": st.column_config.NumberColumn("Avg ticket vs LY", format="%+.1f%%", help="28-day average ticket"),
        "28D Driver": st.column_config.TextColumn("Driven by", help="Which of traffic or ticket moved 28-day sales more"),
        "7D Net Sales % Chg": st.column_config.NumberColumn("7-day vs LY", format="%+.1f%%"),
        "Swing (pts)": st.column_config.NumberColumn("WoW swing", format="%+.1f", help="Change in 7-day vs LY since last week, points"),
        "Flags": st.column_config.TextColumn("Flags", width="medium"),
    },
)

# --- One store in detail -----------------------------------------------------

st.subheader("Why is a store up or down?")
focus = st.selectbox("Store", stores, format_func=label, index=stores.index("PLN01") if "PLN01" in stores else 0)
parts = {"Net sales": "28D Net Sales % Chg", "Traffic (transactions)": "28D Transactions % Chg", "Avg ticket": "28D Avg Ticket % Chg"}
detail = (
    view[view["Store Code"] == focus]
    .rename(columns={column: part for part, column in parts.items()})
    .melt("Week Ending", value_vars=list(parts), var_name="Measure", value_name="% vs LY")
)
detail_chart = alt.Chart(detail).mark_line(strokeWidth=2, point=alt.OverlayMarkDef(size=30)).encode(
    x=alt.X("Week Ending:T", title=None),
    y=alt.Y("% vs LY:Q", title="28-day % change vs LY"),
    color=alt.Color("Measure:N", scale=alt.Scale(domain=list(parts), range=SERIES_COLORS[:3]), legend=alt.Legend(title=None, orient="top")),
    tooltip=[alt.Tooltip("Week Ending:T", title="Week ending"), "Measure:N", alt.Tooltip("% vs LY:Q", format="+.1f")],
)
zero = alt.Chart(pd.DataFrame({"y": [0]})).mark_rule(color=MUTED).encode(y="y:Q")
st.altair_chart((zero + detail_chart).properties(height=300), width="stretch")
st.caption("Net sales = transactions × average ticket. Whichever line tracks net sales more closely is driving it.")

# --- Data quality ------------------------------------------------------------

with st.expander("Data quality"):
    st.markdown(
        "Impossible values found in POS exports are blank in the master and excluded from every chart and "
        "figure above. The original value is kept in the master's Data Quality Flag column."
    )
    flagged = analysis.flagged[["Store Code", "Week Ending", "Data Quality Flag"]]
    if len(flagged):
        st.dataframe(flagged, hide_index=True, width="stretch")
    for code, missing in analysis.missing_weeks.items():
        st.markdown(f"- **{label(code)}**: no readable report for {', '.join(map(str, missing))}")
    for _, row in analysis.snapshot.iterrows():
        if row["Trend Weeks"] < TREND_WEEKS:
            st.markdown(f"- **{label(row['Store Code'])}**: {row['Trend Weeks']} weeks of history; highs and lows are judged on those weeks")
