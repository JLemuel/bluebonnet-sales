"""Run every case in evals/cases.json against the live agent and print PASS/FAIL.

    uv run evals/run_evals.py                 # all cases
    uv run evals/run_evals.py --case q2_ambiguous_austin

A case passes when the final answer
  * matches every ``must_contain`` regex (case-insensitive),
  * matches no ``must_not_contain`` regex,
  * called a tool matching each ``must_call`` regex, and
  * passed the grounding check (every figure traced to a tool result).

Answers and tool calls are saved to evals/results.json. Exit code 1 if any case fails.
"""

import argparse
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

from ask_sales import build_agent  # noqa: E402


def run_case(case: dict, model: str | None, effort: str | None = None) -> dict:
    agent = build_agent(model, effort=effort)  # a fresh conversation per case
    started = time.monotonic()
    calls, answer, model_calls, cost, tokens = [], None, 0, 0.0, 0
    try:
        for question in case.get("turns") or [case["question"]]:
            answer = agent.ask(question)
            calls += [call.name for call in answer.tool_calls]
            model_calls += answer.usage.model_calls
            cost += answer.usage.cost
            tokens += answer.usage.input_tokens + answer.usage.output_tokens
    except Exception as exc:  # an API failure fails the case but not the run
        return {"id": case["id"], "passed": False, "failures": [f"error: {type(exc).__name__}: {exc}"], "answer": "",
                "tools": calls, "seconds": 0, "model_calls": model_calls, "cost": cost, "tokens": tokens}

    text, failures = answer.text, []
    failures += [f"missing /{p}/" for p in case.get("must_contain", []) if not re.search(p, text, re.I)]
    failures += [f"forbidden /{p}/" for p in case.get("must_not_contain", []) if re.search(p, text, re.I)]
    failures += [f"no tool call matching /{p}/" for p in case.get("must_call", []) if not any(re.fullmatch(p, c) for c in calls)]
    if answer.ungrounded:
        failures.append(f"ungrounded figures: {', '.join(answer.ungrounded)}")
    return {
        "id": case["id"], "passed": not failures, "failures": failures, "answer": text, "tools": calls,
        "revised_by_grounding_check": answer.revised, "seconds": round(time.monotonic() - started, 1),
        "model_calls": model_calls, "cost": round(cost, 4), "tokens": tokens,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--case", action="append", help="run only this case id (repeatable)")
    parser.add_argument("--model", help="OpenRouter model id (default: the agent's default)")
    parser.add_argument("--effort", choices=["low", "medium", "high"], help="reasoning effort (default: the agent's default)")
    parser.add_argument("--workers", type=int, default=4, help="cases run in parallel")
    args = parser.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    cases = json.loads((HERE / "cases.json").read_text(encoding="utf-8"))
    if args.case:
        cases = [case for case in cases if case["id"] in args.case]
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        results = list(pool.map(lambda case: run_case(case, args.model, args.effort), cases))

    for case, result in zip(cases, results):
        verdict = "PASS" if result["passed"] else "FAIL"
        print(f"{verdict}  {result['id']:<32} {len(result['tools'])} tool calls  {result['seconds']:>5.1f}s  ${result['cost']:.3f}   {case['tests']}")
        for failure in result["failures"]:
            print(f"      - {failure}")
    passed = sum(result["passed"] for result in results)
    n = len(results) or 1
    print(f"\n{passed}/{len(results)} passed | per case: {sum(r['seconds'] for r in results) / n:.1f}s, "
          f"{sum(r['model_calls'] for r in results) / n:.1f} model calls, {sum(r['tokens'] for r in results) / n:,.0f} tokens, "
          f"${sum(r['cost'] for r in results) / n:.3f} | total ${sum(r['cost'] for r in results):.2f}")
    (HERE / "results.json").write_text(json.dumps(results, indent=2), encoding="utf-8")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
