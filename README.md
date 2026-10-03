# Bluebonnet Blend Co. - weekly sales pipeline and sales agent

```
uv sync                                    # or: pip install -r requirements.txt
uv run process_reports.py Input_Files      # Parts 1-4, first drop
uv run process_reports.py Next_Week_Drop   # second drop -> Sales_Master.xlsx, weekly_ops_brief.txt
uv run streamlit run dashboard.py          # Part 5
uv run ask_sales.py                        # Part 6 (key in .env, see .env.example)
```

The challenge's sample data is not in this repo. To re-run the pipeline, copy `Input_Files/` and `Next_Week_Drop/` into the project folder. The dashboard runs from the included `Sales_Master.xlsx`.

**Pipeline judgment calls.** Impossible values (negative counts, % changes below -100%, Net Sales not equal to Transactions x Avg Ticket, 28D below 7D) are blanked in the master. Their original values stay in Data Quality Flag. The one real case is HOU01 2026-08-09. A store with less than 26 weeks of history (ELP01, 9 weeks) is judged over the weeks it has, and the brief labels it as limited history rather than a 26-week high or low. The DAL01 08-30 correction replaces the original; without it, Dallas would have triggered a false -12.9 point alert.

**Dashboard.** The page has six parts:
- Headline tiles and a "needs attention" list.
- 28-day net sales per store, either on one chart (click the legend to highlight a store) or one chart per store, marked with 26-week highs and lows.
- Two bar charts: % vs last year, and the week-over-week swing against the ±5 pt alert line.
- A scorecard showing whether traffic or ticket drove each store's change.
- A drill-down chart for one store.
- The data quality notes.

Store filters apply to every view, and it works in light and dark mode.

## Part 6 - AI sales agent

**Provider and model:** OpenRouter (OpenAI-compatible API), `anthropic/claude-opus-5.5`, reasoning effort medium. The prompt is cached across the tool loop.

**D1 - How data reaches the model.** Through seven typed tools: `find_stores`, `get_store_metrics`, `compare_stores`, `get_history`, `rank_stores`, `get_weekly_highlights` and `get_data_quality`. The model chooses which to call, and the spreadsheet is never put in the prompt. The tools do all the arithmetic (changes, differences between stores, how many weeks in a row a metric has moved, rankings, traffic-vs-ticket driver, 26-week highs/lows), so the model reads numbers and never computes them. The tools reuse `process_reports.analyse`, so the agent, brief and dashboard always agree.

**D2 - Preventing made-up numbers.** There are three layers:
1. The prompt rule: every number must come from a tool result, and the agent may not do its own arithmetic.
2. A check in code: after each answer, every figure is matched against the numbers the tools returned, at the precision written ($84.7K matches 84680.45). A failing answer is sent back once for revision, and if it still fails it is shown with a warning.
3. Impossible values never reach the model: tools return them as null with a reason, not the raw value.

**D3 - Ambiguity and scope.** `find_stores` returns every match, and the agent must cover each store or ask which one is meant; "the Austin store" has two. For vague questions the agent states its yardstick (% vs last year, plus week-over-week momentum) and offers a follow-up. Out-of-scope questions are declined without using general knowledge. Unknown stores, missing weeks and missing metrics are reported plainly.

**Optimization.** Measured on the eval suite:
- Cost per question fell from $0.059 to $0.023 (−61%).
- Time per question fell from 13.1 s to 10.9 s.
- Two weaknesses were fixed: the agent can now say "by how much" and "how many weeks in a row".

Low effort (`ASK_SALES_EFFORT=low`) also passes every case, at $0.017 per question.

**Known limitation.** The grounding check proves a number came from a tool, not that it is attached to the right metric, and it does not check wording. In one early run the agent quoted Plano's 7-day ticket change as the 28-day one and corrected itself only in a follow-up.

### Transcripts (`uv run ask_sales.py "<question>"`)

```
Q1> Which store fell furthest behind last year over the last 28 days, and was it traffic or average ticket that drove it?
   > rank_stores(metric='28D Net Sales % Chg', order='bottom', n=3)
   > get_store_metrics(store_code='PLN01')
  PLN01 (Plano Legacy) fell furthest behind last year over the last 28 days, with Net Sales of $84,680, down -16.4%. The cause was traffic, not average ticket. This is for the week ending 2026-09-06.
  - Traffic: 28-day Transactions were 9,562, down -17.5% vs last year.
  - Average ticket: $8.86, up +1.5%. That small gain made up for only a little of the lost traffic.
  - It's getting worse: 7-day Net Sales were down -19.7%, compared with -17.3% the week before (a -2.4 point swing). 7-day Transactions were down -20.8%.
  - Next-weakest stores: FTW01 (Fort Worth Sundance) at -8.6% and HOU01 (Houston Heights) at -6.1%. No values were excluded from this ranking.

Q2> How did the Austin store do last week?
   > find_stores(query='Austin')
   > compare_stores(store_codes=['AUS01', 'AUS02'])
   > get_data_quality(store_code='AUS02')
  There are two Austin stores, and both grew compared with the same week last year (week ending 2026-09-06). AUS02 (Austin Mueller) grew faster, with net sales up +13.7%. AUS01 (Austin Downtown) sold more in dollars, but its growth was smaller at +3.9%.
  - AUS01 (Austin Downtown): net sales were $29,452.58 for the week. Transactions were up +3.0% to 3,250, and the average ticket was up +0.9% to $9.06. Most of the growth came from more customers.
  - AUS02 (Austin Mueller): net sales were $25,000 for the week. Transactions were up +11.9% to 2,908, and the average ticket was up +1.5% to $8.60. Most of the growth here also came from more customers.
  - Last 4 weeks: AUS01 was up +4.0% to $117,827.18, and AUS02 was up +7.1% to $95,650.22. That means AUS02's +13.7% last week is well above its 4-week rate.
  - Gap between the stores: AUS02 took in $4,452.58 less than AUS01 last week and had 342 fewer transactions. Its sales growth was 9.8 points higher.
  - Data quality: AUS02's $25,000 is a suspiciously round number, so I checked it. The data shows no excluded values or missing weeks for that store.

Q3> What's the minimum wage in Texas?
  I can only answer questions about Bluebonnet Blend Co. store sales, so I can't help with minimum wage or other legal or labor questions.
  The data covers each store's weekly Net Sales, Transactions (traffic) and Avg Ticket over the last 7 and 28 days, each compared with the same period last year. It runs from the week ending 2026-03-08 to the week ending 2026-09-06. Labor, staffing and wage data aren't included.
```
(Each answer's closing offer of a follow-up is trimmed.)

### Evals - `uv run evals/run_evals.py`

15 cases in `evals/cases.json`. Each lists required and forbidden patterns and the tools it must call, and must also pass the grounding check. Full answers are saved to `evals/results.json`.

```
PASS  q1_furthest_behind_and_driver    2 tool calls   10.5s  $0.031   multi-step reasoning: rank the stores, then explain the driver from the store's own numbers
PASS  q2_ambiguous_austin              3 tool calls   17.2s  $0.042   ambiguity: two stores are in Austin -- cover both or ask, never pick one silently
PASS  q3_out_of_scope_minimum_wage     0 tool calls    5.1s  $0.020   out of scope: decline, and do not answer from general knowledge
PASS  top3_28d_change                  1 tool calls    7.8s  $0.025   ranking with exact figures
PASS  any_26_week_low                  1 tool calls    8.7s  $0.020   uses the pipeline's 26-week analysis, not its own guess
PASS  ftw01_last_month_trend           4 tool calls   11.3s  $0.029   trend over a window, driver identified (avg ticket, not traffic)
PASS  not_in_data_labor                0 tool calls    5.2s  $0.007   data doesn't show it: labor % is not in the master -- say so, invent nothing
PASS  not_in_data_missing_week         3 tool calls   12.6s  $0.016   data doesn't show it: PLN01's 2026-06-07 report was unreadable -- no figure may be made up
PASS  flagged_value_not_repeated       2 tool calls    9.1s  $0.015   an impossible value (-3027 transactions) must not reach the answer
PASS  resubmitted_report_used          2 tool calls    9.2s  $0.012   the corrected DAL01 report replaced the original
PASS  unknown_store_round_rock         1 tool calls    6.4s  $0.009   a store that is not in the directory: say it is not tracked, show no figures for it
PASS  vague_performance                1 tool calls   11.4s  $0.025   ambiguous question: states the yardstick it used
PASS  compare_by_how_much              2 tool calls    9.3s  $0.017   a difference between stores comes from a tool, not left out or made up ($9.01 vs $8.52 = $0.49)
PASS  streak_weeks_in_a_row            2 tool calls   10.2s  $0.017   counts of weeks come from the data (FTW01 7D Net Sales % Chg fell 10 weeks running), not guessed
PASS  follow_up_compare_dallas         4 tool calls   29.0s  $0.057   multi-turn: 'it' refers to the store from the previous question

15/15 passed | per case: 10.9s, 2.8 model calls, 10,523 tokens, $0.023 | total $0.34
```
