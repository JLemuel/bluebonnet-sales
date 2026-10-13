# Development decisions - Bluebonnet sales pipeline and agent

## Summary

| Area | Decision | Why |
|---|---|---|
| Data discovery | Surveyed every file before writing parsing code | Found the traps first: junk files, a corrupt file, an unknown store, a nested zip, impossible values |
| Robustness | A bad file is logged and skipped, never crashes the run | One broken file shouldn't stop a weekly job |
| Validation | Flag only values that *cannot* be true | Unusual but real numbers, like +38% growth, are news, not errors |
| Flagged values | Blanked in the master, original kept in the flag column | Nothing downstream can use a bad number by accident |
| Master file | Store + week is the key; add new weeks, replace resubmitted ones | No duplicates, history never lost, corrections take effect |
| Safety | Atomic writes; refuses to overwrite a master it can't read | A crash or locked file can't wipe history |
| Short history | Judge a store on the weeks it has, labelled "limited history" | Honest about El Paso's 9 weeks |
| Shared analysis | Brief, dashboard and agent use one analysis function | All three always show the same numbers |
| Dashboard | Two views of 28-day sales, swing chart, driver scorecard | Meets the requirement and answers "what needs attention?" |
| Agent design | Typed tools, no spreadsheet in the prompt | Exact data; the model chooses what to look up |
| No model arithmetic | Tools compute every difference, ranking and streak | Models read numbers reliably but calculate unreliably |
| Grounding check | Code verifies every number in an answer | Catches invented figures before the user sees them |
| Tools, not RAG | Exact queries over a small table | Search can return the wrong week; rankings need calculation |
| Evals | 15 live cases aimed at the risky behaviors | Proves the agent is safe, not just that it works |
| Optimization | Measured first, then cached and cut round trips | 61% cheaper per question, every eval still passing |
| Publishing | Public repo without the Hermes files or the key | Shareable code with nothing private exposed |

---

## Detailed report

### 1. Consolidation (Part 1)

**Survey the data first.** Before writing parsing code, I read every file, including those inside the zips, and checked values for consistency. That revealed:
- an Excel lock file (`~$AUS01...`), `Thumbs.db` and a README;
- a corrupt workbook (`PLN01-20260607.xlsx`);
- Round Rock Crossing, which isn't in the store directory;
- Houston's backfill in a zip inside another zip;
- one impossible record (Houston 2026-08-09).

**Zips are read in memory, at any depth.** Nested archives are opened without extracting to disk, so a zip inside a zip needs no temp files or cleanup.

**The filename is the source of truth** for store and week, as the brief specifies. Name matching ignores case and extra spaces, and a store code missing from the directory is treated like a missing name: skipped with a warning.

**Every skip has a category:** not a report, unmapped store, unreadable, or superseded. The brief's data quality section is built from these.

**Layout check before reading cells.** The script confirms B11–B13 say Net Sales, Transactions and Avg Ticket before reading the fixed cells. If the layout ever shifts, the file is rejected rather than read wrongly.

**Validation rules.** A value is flagged only when it cannot be true:
- a negative count;
- a fractional transaction count;
- a % change below -100%;
- Net Sales that doesn't equal Transactions × Avg Ticket;
- a 28-day total smaller than the 7-day total inside it.

Houston's -3,027 transactions and -$8.91 ticket multiply to a believable Net Sales, which is why each field is checked separately.

**Not flagged on purpose:** San Antonio's +38% growth and Plano's slide are consistent week to week, so they are real trends.

**Flagged values are blanked.** The original value goes into the Data Quality Flag text. The rest of the row stays usable; Houston's Net Sales that week is still valid.

**Duplicates within one run:** a file in a Corrections folder wins, otherwise the newest file wins.

### 2. Master file (Part 2)

- **Upsert on Store Code + Week Ending.** New weeks are appended, a resubmitted report replaces its row, and untouched rows are carried over unchanged.
- **Atomic write:** the file is written to a temp file and swapped in.
- **Refuse rather than overwrite:** if the master can't be read because it's corrupt or open in Excel, the script stops with a clear message.
- **Re-running a drop changes nothing.**
- **Formatting:** dates, money and percent formats, frozen headers and filters, so it reads well in Excel.
- **The correction mattered:** Dallas 08-30 went from $35,391 to $32,709. Against the original, Dallas showed a false -12.9 point alert; with the correction it's -4.6.
- **Known trade-off:** the latest run wins. Re-running an old drop after a correction would restore the old numbers, so drops must run in arrival order.

### 3. Analysis (Part 3)

- **Based on each store's most recent entry,** as the brief specifies.
- **Momentum:** swing = this week's 7-day % vs LY minus last week's. A store with no valid previous week is left out of the count, not guessed.
- **26-week highs and lows:** the window is the latest week plus the 25 before it.
  - Stores with less history (El Paso, 9 weeks) are judged on what they have and labelled "N-week, limited history".
  - Below 4 weeks, no high or low is called.
  - A missing week (Plano 06-07) just drops out of the window.
- **Urgent swings:** strictly greater than 5 points. Swings are rounded so computer rounding noise can't create a false alert.
- **Rankings:** by 28-day % vs LY, with the main driver (traffic or average ticket) attached.
- **One analysis function** feeds the brief, dashboard and agent.

**Results (week ending 2026-09-06):**

| Analysis | Result |
|---|---|
| Momentum | 2 of 8 stores improved (AUS02, SAT01) |
| 26-week high | SAT01 San Antonio |
| 26-week lows | PLN01 Plano, HOU01 Houston |
| Urgent swings | FTW01 -5.9 pts, AUS02 +8.4 pts |
| Top 3 (28-day vs LY) | SAT01 +36.8%, ELP01 +8.9%, AUS02 +7.1% |
| Bottom 3 | PLN01 -16.4% (traffic), FTW01 -8.6% (avg ticket), HOU01 -6.1% (traffic) |

### 4. Weekly brief (Part 4)

It follows the brief's example layout, adds store names and the driver behind each top and bottom store, and its data quality section lists:
- skipped files, grouped by reason;
- unmapped stores;
- revised reports;
- excluded values;
- missing weeks.

### 5. Dashboard (Part 5)

- **Streamlit:** pure Python, quick to build, and a free public host.
- **28-day chart in two views:** all stores on one chart (click the legend to highlight a store) and one chart per store with its latest value and any 26-week high or low. Missing weeks show as gaps, not joined lines.
- **Week-over-week swing chart** with the ±5 alert line, so urgency is visible, not just listed.
- **Scorecard** showing traffic vs ticket and the main driver, plus flags.
- **Drill-down chart** for "why is a store up or down?"
- **Fixed store colours, validated palette, light and dark mode,** and filters that apply everywhere.
- **Checked with real screenshots,** which caught bugs the automated test missed: `$` rendering as maths, a cut-off table and wrong sort order.

### 6. AI agent (Part 6)

- **Provider:** OpenRouter (the key available), running Claude Opus 5.5; the model is one setting.
- **Seven typed tools:** find_stores, get_store_metrics, compare_stores, get_history, rank_stores, get_weekly_highlights and get_data_quality. The spreadsheet is never in the prompt.
- **Tools do all the arithmetic:** differences, rankings, streaks, drivers and highs/lows.
- **Grounding check:** code verifies each number against tool results at the precision written. A failure gets one revision, then a visible warning.
- **Impossible values never reach the model;** tools return null with a reason.
- **Ambiguity and scope:**
  - all store matches are returned (two in Austin);
  - vague questions get their yardstick stated;
  - off-topic questions are declined;
  - missing data is said plainly.
- **Tools, not RAG:** small numeric table, exact answers needed.
- **Never crashes:** bad tool calls go back to the model as errors, there is a step limit, and an API failure rolls the conversation back.
- **Multi-turn:** follow-ups keep context; `new` resets it.

### 7. Evals and testing

- **15 live evals, one command,** aimed at risky behaviors:
  - repeating the flagged value;
  - using the pre-correction Dallas figure;
  - inventing labor %;
  - a missing week;
  - an unknown store;
  - off-topic questions;
  - ambiguity;
  - follow-up questions.
- **Offline tests** with a fake model check the agent's machinery without API cost.
- **Pipeline tests:**
  - Part 1: 2,376 values match an independent read of the files;
  - Part 2: 19 checks;
  - Part 3: 26 checks.
- **One honest miss:** an eval expected `$32,708` while the agent correctly rounded to `$32,709`. I fixed the test, not the agent.

### 8. Optimization

- **Measured a baseline first,** then added prompt caching, compact tool results and parallel tool calls.
- **Cost per question:** $0.059 to $0.023 (−61%).
- **Time per question:** 13.1 s to 10.9 s.
- **Added `compare_stores` and week streaks** so "by how much" and "how many weeks in a row" are answered from code.
- **Kept medium effort:** low also passed every case, but one run isn't enough evidence that quality holds.

### 9. Publishing

- **Public GitHub repo** with code and outputs only: no Hermes PDF, sample data or API key, confirmed with a scan.
- **Dashboard prepared for Streamlit Community Cloud** and tested in a clean Python 3.12 environment.
- **The agent isn't public:** the OpenRouter key would pay for every visitor's questions.

### 10. Known limitations and next steps

- **The grounding check verifies numbers, not attribution** (the 7-day vs 28-day ticket mix-up). Next: match each figure to the field it's described as.
- **Latest run wins.** Next: an ingest log recording where each row came from.
- **Test scripts aren't in the repo yet.** Next: move them in and add CI.
- **The Azure bonus wasn't done.**
