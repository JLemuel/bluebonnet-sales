"""Ask questions about Sales_Master.xlsx in plain English (Part 6).

    uv run ask_sales.py                                    # conversation (follow-ups keep context)
    uv run ask_sales.py "How did the Austin store do last week?"

How it works:

* The model never sees the spreadsheet. It gets seven typed tools (``TOOLS``)
  backed by ``SalesData``, and decides which to call. The tools do all the
  arithmetic -- changes, differences between stores, week-in-a-row streaks,
  rankings, drivers, highs/lows -- so the model only has to read numbers,
  never compute them. The highs/lows and swings come from
  process_reports.analyse, so the agent, the brief and the dashboard agree.
* Values the pipeline flagged as impossible are blank in the master, and the
  tools strip the raw value from the flag text: the model cannot repeat a
  number nobody should trust.
* After each answer, ``ungrounded_numbers`` checks every figure in the text
  against the numbers the tools returned. An answer that fails gets one
  automatic revision; if it still fails, the figures are shown with a warning.
* Cost and speed: the tool definitions, system prompt and conversation are
  prompt-cached between rounds of the tool loop, tool results are compact JSON,
  and the model is told to request independent lookups in parallel.

Provider: OpenRouter's OpenAI-compatible API, model anthropic/claude-opus-5.5 by
default (override with --model or ASK_SALES_MODEL). The key is read from .env
(OPENROUTER_API_KEY, or OPENROUTE_KEY), which takes precedence over the shell.
"""

import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

import openai
import pandas as pd
from dotenv import dotenv_values

from process_reports import (
    METRIC_COLUMNS,
    METRICS,
    PERIODS,
    SWING_ALERT_POINTS,
    TREND_WEEKS,
    analyse,
    column_name,
    load_master,
    sales_driver,
)

HERE = Path(__file__).resolve().parent
OPENROUTER_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "anthropic/claude-opus-5.5"
DEFAULT_EFFORT = "medium"  # reasoning effort: low | medium | high
MAX_TOOL_ROUNDS = 8  # model round-trips per question before giving up


class ToolError(Exception):
    """A tool call the data cannot satisfy. The message goes back to the model."""


# --- The data, exposed as tools ----------------------------------------------


def as_number(value: object, digits: int = 2) -> float | int | None:
    if pd.isna(value):
        return None
    number = round(float(value), digits)
    return int(number) if number.is_integer() else number


def latest_run(values: pd.Series) -> dict:
    """How many weeks in a row a metric has moved the same way, ending at its latest week.

    Uses the store's whole history, and stops at a missing or excluded week so a
    gap is never counted as a move.
    """
    points = list(values.dropna().sort_index().items())
    direction, weeks, start = None, 0, None
    for (previous_week, previous), (week, value) in reversed(list(zip(points, points[1:]))):
        step = "up" if value > previous else "down" if value < previous else None
        if (week - previous_week).days != 7 or step is None or direction not in (None, step):
            break
        direction, weeks, start = step, weeks + 1, previous_week
    if not weeks:
        return {"direction": "unchanged", "weeks_in_a_row": 0}
    return {"direction": direction, "weeks_in_a_row": weeks, "after_week_ending": str(start.date())}


FLAG_PART = re.compile(r"^(?P<field>.+?): (?P<reason>.+) \(.*\)$")


def excluded_fields(flag: str) -> list[dict]:
    """Field and reason for each flagged value -- deliberately without the raw value."""
    parts = [FLAG_PART.match(part) for part in flag.split("; ") if part]
    return [{"field": m["field"], "reason": m["reason"]} for m in parts if m]


class SalesData:
    """Read-only questions over the master. Each public method is one tool."""

    def __init__(self, master_path: Path, directory_path: Path):
        self.master = load_master(master_path).sort_values(["Store Code", "Week Ending"], ignore_index=True)
        if self.master.empty:
            raise SystemExit(f"{master_path.name} has no data yet -- run process_reports.py first.")
        directory = pd.read_csv(directory_path, dtype=str).fillna("")
        self.stores = {row.store_code.strip().upper(): row for row in directory.itertuples(index=False)}
        self.analysis = analyse(self.master)
        self.latest = self.analysis.week
        self.first = self.master["Week Ending"].min()

    # -- helpers --

    def _store(self, store_code: str) -> str:
        code = str(store_code or "").strip().upper()
        if code not in self.stores or code not in set(self.master["Store Code"]):
            known = ", ".join(f"{c} ({s.store_name})" for c, s in self.stores.items())
            raise ToolError(f"Unknown store code {store_code!r}. Use find_stores to look a store up. Known stores: {known}.")
        return code

    def _store_info(self, code: str) -> dict:
        store = self.stores[code]
        return {"store_code": code, "store_name": store.store_name, "city": store.city}

    def _metric(self, metric: str) -> str:
        wanted = str(metric or "").strip().casefold()
        for column in METRIC_COLUMNS:
            if column.casefold() == wanted:
                return column
        raise ToolError(f"Unknown metric {metric!r}. Use one of: {', '.join(METRIC_COLUMNS)}.")

    def _week(self, week_ending: str | None) -> pd.Timestamp:
        """A date -> the Sunday week ending on or after it. None means the latest week."""
        if not week_ending or str(week_ending).strip().casefold() in {"latest", "last week", "this week"}:
            return self.latest
        try:
            day = pd.Timestamp(str(week_ending)).normalize()
        except (ValueError, TypeError):
            raise ToolError(f"Could not read {week_ending!r} as a date. Use YYYY-MM-DD, e.g. {self.latest.date()}.") from None
        return day + pd.Timedelta(days=6 - day.weekday())

    def _history(self, code: str) -> pd.DataFrame:
        return self.master[self.master["Store Code"] == code]

    def _period(self, row: pd.Series, period: str) -> dict:
        out = {}
        for metric in METRICS:
            key = metric.lower().replace(" ", "_")
            out[key] = as_number(row[column_name(period, metric)])
            out[f"{key}_pct_chg_vs_ly"] = as_number(row[column_name(period, metric, pct=True)], 1)
        out["net_sales_change_driven_mainly_by"] = sales_driver(
            row[column_name(period, "Transactions", pct=True)], row[column_name(period, "Avg Ticket", pct=True)]
        )
        return out

    # -- tools --

    def find_stores(self, query: str | None = None) -> dict:
        tokens = [t for t in re.split(r"\W+", (query or "").casefold()) if t and t not in {"store", "the", "location", "shop"}]
        matches = []
        for code, store in self.stores.items():
            haystack = f"{code} {store.store_name} {store.city} {store.state}".casefold()
            if all(token in haystack for token in tokens):
                history = self._history(code)
                matches.append({
                    **self._store_info(code),
                    "state": store.state,
                    "first_week": str(history["Week Ending"].min().date()) if len(history) else None,
                    "latest_week": str(history["Week Ending"].max().date()) if len(history) else None,
                    "weeks_of_data": len(history),
                })
        result = {"query": query, "matches": matches}
        if not matches:
            result["note"] = (
                f"No store in the store directory matches {query!r}. Only these stores are tracked: "
                + ", ".join(f"{c} ({s.store_name}, {s.city})" for c, s in self.stores.items())
            )
        elif len(matches) > 1 and tokens:
            result["note"] = f"{len(matches)} stores match -- answer for each or ask which one is meant."
        return result

    def get_store_metrics(self, store_code: str, week_ending: str | None = None) -> dict:
        code, week = self._store(store_code), self._week(week_ending)
        history = self._history(code)
        rows = history[history["Week Ending"] == week]
        if rows.empty:
            first, last = history["Week Ending"].min().date(), history["Week Ending"].max().date()
            if first <= week.date() <= last:
                raise ToolError(f"No report for {code} for the week ending {week.date()}: that week is missing (the file was unreadable or never arrived).")
            raise ToolError(f"No data for {code} for the week ending {week.date()}. {code} has data from {first} to {last}.")
        row = rows.iloc[0]
        previous = history[history["Week Ending"] == week - pd.Timedelta(weeks=1)]
        previous_pct = previous["7D Net Sales % Chg"].iloc[0] if len(previous) else float("nan")
        result = {
            **self._store_info(code),
            "week_ending": str(week.date()),
            "is_latest_week": week == self.latest,
            "last_7_days": self._period(row, "7D"),
            "last_28_days": self._period(row, "28D"),
            "week_over_week": {
                "previous_week_ending": str((week - pd.Timedelta(weeks=1)).date()),
                "previous_week_7d_net_sales_pct_chg_vs_ly": as_number(previous_pct, 1),
                "swing_in_7d_net_sales_pct_chg_pts": as_number(row["7D Net Sales % Chg"] - previous_pct, 1),
            },
        }
        if row["Data Quality Flag"]:
            result["excluded_values"] = excluded_fields(row["Data Quality Flag"])
            result["excluded_values_note"] = "These values were impossible in the POS export and are not available (shown as null)."
        return result

    def get_history(self, store_code: str, metric: str, weeks: int = 8) -> dict:
        code, column = self._store(store_code), self._metric(metric)
        weeks = max(1, min(int(weeks), 52))
        by_week = self._history(code).set_index("Week Ending")[[column]]
        start, end = by_week.index.min(), by_week.index.max()
        series = []
        for week in pd.date_range(end - pd.Timedelta(weeks=weeks - 1), end, freq="7D"):
            if week < start:
                continue
            point = {"week_ending": str(week.date()), "value": None}
            if week not in by_week.index:
                point["note"] = "missing week (no readable report)"
            elif pd.isna(value := by_week.at[week, column]):
                point["note"] = "value excluded as impossible"
            else:
                point["value"] = as_number(value)
            series.append(point)

        result = {**self._store_info(code), "metric": column, "weeks_requested": weeks, "series": series}
        if len(series) < weeks:
            result["note"] = f"{code} only has data from {start.date()}, so fewer weeks are available."
        valid = [p for p in series if p["value"] is not None]
        if len(valid) >= 2:
            first, last = valid[0], valid[-1]
            summary = {
                "from_week": first["week_ending"], "from_value": first["value"],
                "to_week": last["week_ending"], "to_value": last["value"],
                "highest": max(valid, key=lambda p: p["value"]), "lowest": min(valid, key=lambda p: p["value"]),
            }
            if column.endswith("% Chg"):
                summary["change_pts"] = as_number(last["value"] - first["value"], 1)
            else:
                summary["change"] = as_number(last["value"] - first["value"])
                summary["change_pct"] = as_number((last["value"] / first["value"] - 1) * 100, 1)
            result["summary"] = summary
        result["latest_run"] = latest_run(by_week[column])
        return result

    def compare_stores(self, store_codes: list[str], week_ending: str | None = None) -> dict:
        codes = list(dict.fromkeys(self._store(code) for code in store_codes or []))
        if len(codes) < 2:
            raise ToolError("Give at least two different store codes to compare.")
        week = self._week(week_ending)
        rows, missing = {}, []
        for code in codes:
            found = self.master[(self.master["Store Code"] == code) & (self.master["Week Ending"] == week)]
            if found.empty:
                missing.append(code)
            else:
                rows[code] = found.iloc[0]
        if len(rows) < 2:
            raise ToolError(f"Need at least two stores with a report for the week ending {week.date()} (missing: {', '.join(missing) or 'none'}).")
        base, *others = rows
        differences = {}
        for other in others:
            differences[f"{other} minus {base}"] = {
                column + (" (pts)" if column.endswith("% Chg") else ""): as_number(rows[other][column] - rows[base][column], 1 if column.endswith("% Chg") else 2)
                for column in METRIC_COLUMNS
            }
        result = {
            "week_ending": str(week.date()),
            "stores": [{**self._store_info(c), "last_7_days": self._period(r, "7D"), "last_28_days": self._period(r, "28D")} for c, r in rows.items()],
            "differences": differences,
            "note": "Differences in % Chg metrics are percentage points. Null means a value was excluded as impossible.",
        }
        if missing:
            result["not_reported_this_week"] = missing
        return result

    def rank_stores(self, metric: str, n: int = 3, order: str = "top", week_ending: str | None = None) -> dict:
        column, week = self._metric(metric), self._week(week_ending)
        n = max(1, min(int(n), len(self.stores)))
        rows = self.master[self.master["Week Ending"] == week]
        if rows.empty:
            raise ToolError(f"No data for the week ending {week.date()}. Data covers {self.first.date()} to {self.latest.date()}.")
        bottom = str(order).strip().casefold() == "bottom"
        valid = rows.dropna(subset=[column]).sort_values(column, ascending=bottom)
        excluded = [{"store_code": c, "reason": "no report for this week"} for c in sorted(set(self.master["Store Code"]) - set(rows["Store Code"]))]
        excluded += [{"store_code": c, "reason": "value excluded as impossible"} for c in rows.loc[rows[column].isna(), "Store Code"]]
        return {
            "metric": column,
            "week_ending": str(week.date()),
            "order": "lowest first" if bottom else "highest first",
            "stores_ranked": len(valid),
            "ranking": [
                {"rank": rank, **self._store_info(row["Store Code"]), "value": as_number(row[column])}
                for rank, (_, row) in enumerate(valid.head(n).iterrows(), start=1)
            ],
            "excluded": excluded,
        }

    def get_weekly_highlights(self) -> dict:
        a = self.analysis
        snapshot = a.snapshot.set_index("Store Code")

        def store(code: str) -> dict:
            return {"store_code": code, "store_name": self.stores[code].store_name}

        trend_alerts = []
        for _, row in a.trend_alerts.iterrows():
            kind = row["Trend"].lower()
            limited = row["Trend Weeks"] < TREND_WEEKS
            trend_alerts.append({
                **store(row["Store Code"]),
                "alert": f"{row['Trend Weeks']}-week {kind}" + (" (limited history)" if limited else ""),
                "28d_net_sales": as_number(row["28D Net Sales"]),
                f"previous_{kind}": as_number(row[f"Previous {kind.title()}"]),
            })
        return {
            "week_ending": str(self.latest.date()),
            "momentum": {
                "definition": "a store improved if its 7D Net Sales % Chg vs LY rose versus the previous week",
                "improved_count": len(a.improved),
                "store_count": len(a.snapshot),
                "stores": [
                    {**store(row["Store Code"]), "7d_net_sales_pct_chg": as_number(row["7D Net Sales % Chg"], 1),
                     "previous_week": as_number(row["Previous 7D Net Sales % Chg"], 1),
                     "swing_pts": as_number(row["Swing (pts)"], 1), "improved": bool(row["Swing (pts)"] > 0)}
                    for _, row in a.comparable.sort_values("Swing (pts)", ascending=False).iterrows()
                ],
            },
            "trend_alerts_28d_net_sales": {
                "definition": f"28D Net Sales at its highest or lowest point of the trailing {TREND_WEEKS} weeks",
                "alerts": trend_alerts,
                "stores_with_limited_history": [
                    {"store_code": code, "weeks_of_history": int(weeks)}
                    for code, weeks in snapshot["Trend Weeks"].items() if weeks < TREND_WEEKS
                ],
            },
            "urgent_weekly_swings": {
                "threshold_pts": SWING_ALERT_POINTS,
                "alerts": [
                    {**store(row["Store Code"]), "swing_pts": as_number(row["Swing (pts)"], 1),
                     "this_week_7d_net_sales_pct_chg": as_number(row["7D Net Sales % Chg"], 1),
                     "last_week_7d_net_sales_pct_chg": as_number(row["Previous 7D Net Sales % Chg"], 1)}
                    for _, row in a.swing_alerts.iterrows()
                ],
            },
            "ranking_28d_net_sales_pct_chg": [
                {"rank": rank, **store(row["Store Code"]), "value": as_number(row["28D Net Sales % Chg"], 1),
                 "driven_mainly_by": row["28D Driver"]}
                for rank, (_, row) in enumerate(a.ranked().iterrows(), start=1)
            ],
        }

    def get_data_quality(self, store_code: str | None = None) -> dict:
        code = self._store(store_code) if store_code else None
        flagged = self.analysis.flagged
        coverage = self.master.groupby("Store Code")["Week Ending"].agg(["min", "max", "count"])
        if code:
            flagged, coverage = flagged[flagged["Store Code"] == code], coverage.loc[[code]]
        return {
            "excluded_values": [
                {"store_code": row["Store Code"], "week_ending": str(row["Week Ending"].date()),
                 "fields": excluded_fields(row["Data Quality Flag"])}
                for _, row in flagged.iterrows()
            ],
            "missing_weeks": [
                {"store_code": c, "week_ending": str(week)}
                for c, weeks in self.analysis.missing_weeks.items() if code in (None, c) for week in weeks
            ],
            "coverage": [
                {"store_code": c, "first_week": str(r["min"].date()), "latest_week": str(r["max"].date()), "weeks_with_data": int(r["count"])}
                for c, r in coverage.iterrows()
            ],
            "notes": [
                "Excluded values were impossible in the POS export (e.g. negative transactions) and are not available.",
                "Only stores listed in the store directory are loaded.",
            ],
        }


METRIC_SCHEMA = {"type": "string", "enum": METRIC_COLUMNS, "description": "Exact metric column name."}
STORE_SCHEMA = {"type": "string", "description": "Store code such as PLN01. Use find_stores to resolve names or cities."}
WEEK_SCHEMA = {"type": "string", "description": "Week ending date, YYYY-MM-DD (a Sunday). Omit for the latest week."}


def tool(name: str, description: str, properties: dict, required: list[str]) -> dict:
    return {"type": "function", "function": {
        "name": name, "description": description,
        "parameters": {"type": "object", "properties": properties, "required": required, "additionalProperties": False},
    }}


TOOLS = [
    tool("find_stores",
         "Look up stores by city, store name or code. Use it whenever the user names a store or city. "
         "Returns every match with its code and the weeks of data available; an empty query lists all stores.",
         {"query": {"type": "string", "description": "City, store name or code, e.g. 'Austin' or 'Plano'."}}, []),
    tool("get_store_metrics",
         "All metrics for one store and week: Net Sales, Transactions and Avg Ticket over the last 7 and 28 days, "
         "each with % change vs last year, what mainly drove the Net Sales change (traffic or average ticket), "
         "and the week-over-week swing in 7D Net Sales % Chg. Also reports values excluded as impossible.",
         {"store_code": STORE_SCHEMA, "week_ending": WEEK_SCHEMA}, ["store_code"]),
    tool("compare_stores",
         "Side-by-side metrics for two or more stores in one week (default the latest), with the differences between "
         "them already calculated (each store minus the first one listed). Use it for any comparison or 'by how much'.",
         {"store_codes": {"type": "array", "items": {"type": "string"}, "minItems": 2, "maxItems": 8,
                          "description": "Store codes to compare; differences are relative to the first."},
          "week_ending": WEEK_SCHEMA},
         ["store_codes"]),
    tool("get_history",
         "Weekly values of one metric for one store over the last N weeks (ending at the latest week), with a summary: "
         "start and end values, the change between them, the highest and lowest weeks, and latest_run -- how many "
         "weeks in a row the metric has risen or fallen. Use it for trends.",
         {"store_code": STORE_SCHEMA, "metric": METRIC_SCHEMA,
          "weeks": {"type": "integer", "minimum": 1, "maximum": 52, "description": "How many weeks; 4 is about a month."}},
         ["store_code", "metric"]),
    tool("rank_stores",
         "Rank all stores on one metric for one week (default the latest). Use for best/worst/top/bottom questions.",
         {"metric": METRIC_SCHEMA, "n": {"type": "integer", "minimum": 1, "maximum": 8, "description": "How many stores to return."},
          "order": {"type": "string", "enum": ["top", "bottom"], "description": "top = highest values first, bottom = lowest first."},
          "week_ending": WEEK_SCHEMA},
         ["metric"]),
    tool("get_weekly_highlights",
         "This week's portfolio analysis: momentum (stores whose 7D Net Sales % Chg improved on last week), stores at a "
         "26-week high or low in 28D Net Sales, urgent week-over-week swings over 5 points, and the full ranking by "
         "28D Net Sales % Chg with its main driver.",
         {}, []),
    tool("get_data_quality",
         "Data quality: values excluded as impossible, missing weeks, and the weeks of data each store has.",
         {"store_code": {"type": "string", "description": "Optional store code to filter to one store."}}, []),
]
TOOL_NAMES = {t["function"]["name"] for t in TOOLS}


# --- Grounding check ---------------------------------------------------------

# Things in an answer that look numeric but are not data: store codes, dates,
# years, ordinals and window names like "28-day".
NOT_DATA = [
    re.compile(r"\b[A-Za-z]{3}\d{2}\b"),
    re.compile(r"\b\d{4}-\d{2}-\d{2}\b"),
    re.compile(r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}(?:st|nd|rd|th)?\b(?:,?\s+\d{4})?", re.I),
    re.compile(r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b"),
    re.compile(r"\b(?:19|20)\d{2}\b"),
    re.compile(r"\b\d+(?:st|nd|rd|th)\b"),
    re.compile(r"\b\d+\s?-?\s?(?:day|week|month)s?\b", re.I),
    re.compile(r"\b\d+D\b"),
]
ANSWER_NUMBER = re.compile(
    r"(?<![\w.])(?P<dollar>\$)?(?P<int>\d{1,3}(?:,\d{3})+|\d+)(?P<dec>\.\d+)?"
    r"(?:\s?(?P<suffix>[kKmM])(?![A-Za-z]))?(?P<pct>\s?%)?"
)
JSON_NUMBER = re.compile(r"-?\d+(?:\.\d+)?")


def numbers_in(text: str) -> set[float]:
    return {abs(float(match)) for match in JSON_NUMBER.findall(text)}


def ungrounded_numbers(answer: str, pool: set[float]) -> list[str]:
    """Figures in ``answer`` that match no number a tool returned (or the user typed).

    A figure matches if it equals a tool number at the precision it is written
    in: "$84,680" and "$84.7K" both match 84680.45. Sign is ignored ("down
    16.4%" matches -16.4). Small whole numbers (counts, "top 3") are not checked.
    """
    text = answer
    for pattern in NOT_DATA:
        text = pattern.sub(" ", text)
    problems = []
    for match in ANSWER_NUMBER.finditer(text):
        whole, decimals, suffix = match["int"], match["dec"] or "", (match["suffix"] or "").lower()
        is_plain_count = not (match["dollar"] or decimals or "," in whole or suffix or match["pct"])
        value = float(whole.replace(",", "") + decimals)
        if is_plain_count and value <= 31:
            continue
        scale = {"k": 1e3, "m": 1e6}.get(suffix, 1)
        tolerance = 0.5 * 10 ** -(len(decimals) - 1 if decimals else 0) * scale + 1e-9
        if not any(abs(value * scale - known) <= tolerance for known in pool):
            problems.append(match.group(0).strip())
    return problems


# --- The agent ---------------------------------------------------------------

SYSTEM_PROMPT = """\
You are the sales analyst for Bluebonnet Blend Co., a smoothie franchise in Texas. You answer questions from \
franchise owners and store managers about weekly store sales, using only the tools provided.

The data
- One report per store per week (weeks end on Sunday), from {first} to {latest}. "Last week", "this week" and \
"latest" mean the week ending {latest} unless the user names another week.
- For each store and week: Net Sales, Transactions (traffic) and Avg Ticket over the trailing 7 days and 28 days, \
each with its % change vs the same period last year.
- Net Sales = Transactions x Avg Ticket, so a change in Net Sales is driven by traffic, by ticket, or both.
- Not in the data: labor or staffing, costs, profit, last-year dollar amounts, daily or hourly sales, products, \
weather, local events, and the reasons behind any change.
- Some impossible values were excluded and some weeks are missing; the tools say so when it matters.

Rules
1. Every number you state must come from a tool result in this conversation. Do not calculate new figures \
(differences, sums, averages, counts of weeks). compare_stores returns differences between stores and get_history \
returns changes over time and how many weeks in a row a metric has moved; otherwise describe the comparison in words.
2. If the tools cannot answer, say so plainly and say what the data does cover. Never estimate or fill gaps from \
general knowledge. Questions that are not about these stores' sales (laws, wages, recipes, general facts) are out of \
scope: say you can only answer questions about the sales data, and do not answer them even if you know the answer.
3. Resolve store names and cities with find_stores. If more than one store matches (there are two in Austin), \
answer for each one or ask which is meant -- never pick one silently.
4. For vague questions ("is performance good?"), state the yardstick you used (by default % change vs last year, \
plus week-over-week momentum), answer briefly, and offer a more specific follow-up.
5. For "why" questions, explain with the data's own drivers (traffic vs average ticket, the trend over recent \
weeks). Say that outside causes are not in the data if the user asks for them.
6. Mention a data quality issue (excluded value, missing week, short history) when it affects the answer.
7. Work efficiently: request lookups that do not depend on each other together in one turn (for example \
get_history for two metrics), and do not call a tool whose answer you already have.
8. Write for a terminal: plain text, no markdown headings, tables or bold. Lead with the answer in one or two \
sentences, then at most five short "-" bullets. Refer to stores as code plus name, e.g. PLN01 (Plano Legacy). \
Money like $84,680; percentages with one decimal and a sign, e.g. -16.4%."""

GROUNDING_RETRY = (
    "[Automated check, not from the user] These figures in your answer do not match any number returned by a tool: "
    "{figures}. Rewrite the answer so that every number comes from a tool result -- call a tool if you need the figure, "
    "or describe the comparison in words. Reply with the corrected answer only."
)


@dataclass
class ToolCall:
    name: str
    arguments: dict
    result: str


@dataclass
class Usage:
    """Token and dollar cost of one answer, as reported by OpenRouter."""

    model_calls: int = 0
    input_tokens: int = 0
    cached_tokens: int = 0  # part of input_tokens served from the prompt cache
    output_tokens: int = 0
    cost: float = 0.0

    def add(self, usage) -> None:
        if usage is None:
            return
        self.model_calls += 1
        self.input_tokens += usage.prompt_tokens or 0
        self.output_tokens += usage.completion_tokens or 0
        self.cached_tokens += getattr(getattr(usage, "prompt_tokens_details", None), "cached_tokens", 0) or 0
        self.cost += float(getattr(usage, "cost", 0) or 0)


@dataclass
class Answer:
    text: str
    tool_calls: list[ToolCall] = field(default_factory=list)
    ungrounded: list[str] = field(default_factory=list)
    revised: bool = False  # True if the grounding check sent the first draft back
    usage: Usage = field(default_factory=Usage)


def load_api_key() -> str | None:
    """The project's .env wins over the shell, so a stale exported key cannot shadow it."""
    file_values = dotenv_values(HERE / ".env")
    for name in ("OPENROUTER_API_KEY", "OPENROUTE_KEY"):
        if file_values.get(name):
            return file_values[name]
    return os.environ.get("OPENROUTER_API_KEY")


class SalesAgent:
    """A conversation with the model. Follow-up questions keep the earlier context."""

    def __init__(self, data: SalesData, client: openai.OpenAI, model: str, on_tool_call=None, effort: str = DEFAULT_EFFORT):
        self.data, self.client, self.model, self.effort = data, client, model, effort
        self.on_tool_call = on_tool_call or (lambda call: None)
        self.system = SYSTEM_PROMPT.format(first=data.first.date(), latest=data.latest.date())
        self.reset()

    def reset(self) -> None:
        self.messages: list[dict] = []
        self.known_numbers: set[float] = set()  # every number tools have returned in this conversation

    def ask(self, question: str) -> Answer:
        question = question.strip()
        if not question:
            return Answer("Ask me a question about store sales, for example: How did Plano do last week?")
        checkpoint, usage = len(self.messages), Usage()
        try:
            self.messages.append({"role": "user", "content": question})
            self.known_numbers |= numbers_in(question)
            answer = Answer(self._complete(usage), usage=usage)
            answer.ungrounded = ungrounded_numbers(answer.text, self.known_numbers)
            if answer.ungrounded:
                answer.revised = True
                self.messages.append({"role": "user", "content": GROUNDING_RETRY.format(figures=", ".join(answer.ungrounded))})
                answer.text = self._complete(usage)
                answer.ungrounded = ungrounded_numbers(answer.text, self.known_numbers)
            answer.tool_calls = self._calls_since(checkpoint)
            return answer
        except Exception:
            del self.messages[checkpoint:]  # keep the history valid for the next question
            raise

    def _request_messages(self) -> list[dict]:
        """System prompt + history, with two prompt-cache breakpoints.

        One after the system prompt (tools and system are identical on every
        call) and one on the newest message, so each round of the tool loop
        re-reads the conversation so far from cache instead of paying for it
        again. The breakpoints are added to copies; the history is untouched.
        """
        def cached(text: str) -> list[dict]:
            return [{"type": "text", "text": text, "cache_control": {"type": "ephemeral"}}]

        messages = [{"role": "system", "content": cached(self.system)}, *self.messages]
        if isinstance(messages[-1].get("content"), str):
            messages[-1] = {**messages[-1], "content": cached(messages[-1]["content"])}
        return messages

    def _complete(self, usage: Usage) -> str:
        """Run the tool loop until the model answers in text."""
        for _ in range(MAX_TOOL_ROUNDS):
            response = self.client.chat.completions.create(
                model=self.model,
                messages=self._request_messages(),
                tools=TOOLS,
                max_tokens=8000,
                extra_body={"reasoning": {"effort": self.effort}, "usage": {"include": True}},
            )
            usage.add(response.usage)
            choice = response.choices[0]
            message = choice.message
            self.messages.append(message.model_dump(exclude_none=True))
            if not message.tool_calls:
                text = (message.content or "").strip()
                if choice.finish_reason == "length":
                    text += "\n(The answer was cut off.)"
                return text or "I could not produce an answer to that."
            for call in message.tool_calls:
                result = self._run_tool(call.function.name, call.function.arguments)
                self.known_numbers |= numbers_in(result)
                self.messages.append({"role": "tool", "tool_call_id": call.id, "content": result})
                self.on_tool_call(ToolCall(call.function.name, self._parse(call.function.arguments), result))
        return "I could not finish working that out. Try asking a narrower question."

    @staticmethod
    def _parse(arguments: str) -> dict:
        try:
            parsed = json.loads(arguments or "{}")
        except json.JSONDecodeError:
            return {}
        return parsed if isinstance(parsed, dict) else {}

    def _run_tool(self, name: str, arguments: str) -> str:
        """Execute one tool call. Any failure goes back to the model as an error, never a crash."""
        try:
            if name not in TOOL_NAMES:
                raise ToolError(f"There is no tool called {name!r}.")
            result = getattr(self.data, name)(**self._parse(arguments))
        except ToolError as exc:
            result = {"error": str(exc)}
        except (TypeError, ValueError) as exc:
            result = {"error": f"Invalid arguments for {name}: {exc}"}
        return json.dumps(result, default=str, separators=(",", ":"))  # compact: fewer tokens

    def _calls_since(self, checkpoint: int) -> list[ToolCall]:
        calls, results = [], {}
        for message in self.messages[checkpoint:]:
            if message.get("role") == "tool":
                results[message["tool_call_id"]] = message["content"]
        for message in self.messages[checkpoint:]:
            for call in message.get("tool_calls") or []:
                calls.append(ToolCall(call["function"]["name"], self._parse(call["function"]["arguments"]), results.get(call["id"], "")))
        return calls


def build_agent(model: str | None = None, master: Path = HERE / "Sales_Master.xlsx", on_tool_call=None,
                effort: str | None = None) -> SalesAgent:
    key = load_api_key()
    if not key:
        raise SystemExit("No OpenRouter key found. Put OPENROUTER_API_KEY=... in .env next to ask_sales.py.")
    client = openai.OpenAI(base_url=OPENROUTER_URL, api_key=key, default_headers={"X-Title": "Bluebonnet ask_sales"})
    data = SalesData(master, HERE / "Store_Directory.csv")
    return SalesAgent(data, client, model or os.environ.get("ASK_SALES_MODEL", DEFAULT_MODEL), on_tool_call,
                      effort or os.environ.get("ASK_SALES_EFFORT", DEFAULT_EFFORT))


# --- Terminal ----------------------------------------------------------------


def show_tool_call(call: ToolCall) -> None:
    arguments = ", ".join(f"{k}={v!r}" for k, v in call.arguments.items())
    status = "  (error)" if call.result.startswith('{"error"') else ""
    print(f"   > {call.name}({arguments}){status}", flush=True)


def show_answer(answer: Answer, seconds: float) -> None:
    print()
    for line in answer.text.splitlines():
        print(f"  {line}")
    if answer.ungrounded:
        print(f"\n  [!] Could not verify these figures against the data: {', '.join(answer.ungrounded)}")
    cost = f", ${answer.usage.cost:.3f}" if answer.usage.cost else ""
    print(f"\n  ({len(answer.tool_calls)} tool call{'s' * (len(answer.tool_calls) != 1)}, {seconds:.1f}s{cost})\n")


def ask_and_show(agent: SalesAgent, question: str) -> None:
    started = time.monotonic()
    try:
        answer = agent.ask(question)
    except openai.AuthenticationError:
        print("\n  OpenRouter rejected the API key. Check OPENROUTER_API_KEY in .env.\n")
        return
    except openai.RateLimitError:
        print("\n  Rate limited by OpenRouter. Wait a moment and ask again.\n")
        return
    except openai.APIConnectionError:
        print("\n  Could not reach OpenRouter. Check the internet connection.\n")
        return
    except openai.APIStatusError as exc:
        print(f"\n  The model provider returned an error ({exc.status_code}). Try again.\n")
        return
    show_answer(answer, time.monotonic() - started)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Ask questions about Sales_Master.xlsx in plain English.")
    parser.add_argument("question", nargs="*", help="a single question; omit for a conversation")
    parser.add_argument("--model", help=f"OpenRouter model id (default {DEFAULT_MODEL})")
    parser.add_argument("--quiet", action="store_true", help="hide the tool calls")
    args = parser.parse_args(argv)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    agent = build_agent(args.model, on_tool_call=None if args.quiet else show_tool_call)
    if args.question:
        ask_and_show(agent, " ".join(args.question))
        return 0

    print(f"Bluebonnet sales assistant ({agent.model}). Data to week ending {agent.data.latest.date()}.")
    print("Ask about store sales. Follow-ups keep context. 'new' starts over, 'exit' quits.\n")
    while True:
        try:
            question = input("You> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return 0
        command = re.sub(r"[^a-z]", "", question.casefold())  # ignores a BOM or stray characters from piped input
        if command in {"exit", "quit", "q"}:
            return 0
        if command in {"new", "reset"}:
            agent.reset()
            print("  (started a new conversation)\n")
            continue
        ask_and_show(agent, question)


if __name__ == "__main__":
    sys.exit(main())
