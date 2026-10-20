"""Ask Data parity tooling (read-only dev tool: never writes to Supabase or any data store).

  py -3 scripts/ask_data_parity.py ast-check [--rev main]
      AST-compares every definition moved into services/ask_data/* with the original in
      `git show <rev>:streamlit_app.py` (formatting/comments ignored).

  py -3 scripts/ask_data_parity.py golden [--rev main]
      Regenerates tests/golden/ask_data_v2_characterization.json by running the synthetic scenarios
      through the ORIGINAL code (extracted by AST from `git show <rev>:streamlit_app.py`).

  py -3 scripts/ask_data_parity.py compare [--rev main]
      Runs the synthetic scenarios through (1) the original code, (2) the refactored working-tree
      streamlit_app.py and (3) the extracted services/ask_data modules, and reports any difference.

  py -3 scripts/ask_data_parity.py live [--rev main]
      Same three-way comparison on LIVE CRM data (read-only Supabase selects), including a check
      of the current outstanding total.

Used to prove that the Phase 1 extraction is a pure move with no behavior change.
"""
from __future__ import annotations

import argparse
import ast
import datetime
import difflib
import json
import subprocess
import sys
from decimal import Decimal
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

GOLDEN = ROOT / "tests" / "golden" / "ask_data_v2_characterization.json"

# module file -> names that must be identical to the original streamlit_app.py definitions
MOVED = {
    "services/ask_data/helpers.py": [
        "_safe_int", "_safe_float", "_safe_str", "_normalize_project_name_key", "_suggest_best_project_match",
        "_invoice_category_label", "_is_paid_trial_category", "_is_new_installation_category",
        "_is_maintenance_category", "ORDER_STATUS_OPTIONS", "PROJECT_STATUS_OPTIONS", "DETECTION_CAMERA_FIELDS",
        "_normalize_query_text", "_normalize_project_status", "_normalize_country", "_project_license_date",
        "_license_status", "_normalize_detection_key", "_parse_detection_camera_counts",
        "_project_detection_counts", "_match_countries_from_question", "_build_invoice_answer_df",
        "_build_project_answer_df",
    ],
    "services/ask_data/tool_specs.py": [
        "_ASK_DATA_TOOL_SPECS", "_ASK_DATA_TOOL_DESCRIPTIONS", "_ask_data_tool_catalog_text",
    ],
    "services/ask_data/validation.py": [
        "_resolve_project_arg", "_resolve_single_country", "_resolve_ask_data_arg", "_validate_ask_data_tool_call",
    ],
    "services/ask_data/executor.py": [
        "_camera_metric_for_project", "_filter_invoices_for_debt_tools", "_filter_license_entries",
        "_annual_payment_month", "_latest_annual_invoice_year", "_annual_license_payment_row",
        "_execute_ask_data_tool",
    ],
    "services/ask_data/planner.py": ["_llm_parse_data_question"],
    "services/ask_data/router.py": ["_answer_data_question_smart"],
}

# Definitions that are intentionally NOT byte-identical in Phase 1 (each documented in the new module).
EXPECTED_DEVIATIONS = {
    # Phase 1 injection points
    "_execute_ask_data_tool": "Phase 1: keyword-only `next_invoice_number`; Phase 2: month fix + canonical matching; Phase 3: thin tuple wrapper over execute_tool/ToolResult (Decimal money, rows, calculation)",
    "_answer_data_question_smart": "Phase 1: injected legacy fallbacks; Phase 2: validate_tool_call + one-question clarification",
    # Phase 2 (approved scope only)
    "_build_invoice_answer_df": "Phase 2: Country via canonical ProjectIndex; Phase 3: Amount (€) as exact Decimal",
    "_ASK_DATA_TOOL_SPECS": "Phase 3: adds the whitelisted debt_by_country tool (appended last)",
    "_ASK_DATA_TOOL_DESCRIPTIONS": "Phase 3: adds the debt_by_country description (appended last)",
    "_annual_license_payment_row": "Phase 3: all-time/period amounts computed with Decimal instead of float",
    "_resolve_project_arg": "Phase 2: canonical resolver, thresholds 0.90 / 0.60",
    "_resolve_single_country": "Phase 2: canonical resolver (accents, Hebrew, variants), thresholds 0.90 / 0.60",
    "_resolve_ask_data_arg": "Phase 2: month resolved via resolve_month (Hebrew, 'sept'); project/country via the resolver",
    "_validate_ask_data_tool_call": "Phase 2: thin wrapper over validate_tool_call (adds clarification)",
    "_filter_invoices_for_debt_tools": "Phase 2: country/project matched via ProjectIndex instead of exact project-name equality",
}


def git_show(rev: str, path: str) -> str:
    return subprocess.run(["git", "show", f"{rev}:{path}"], cwd=ROOT, capture_output=True, check=True, encoding="utf-8").stdout


def top_level_defs(source: str) -> dict[str, ast.AST]:
    out: dict[str, ast.AST] = {}
    for node in ast.parse(source).body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            out[node.name] = node
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    out[target.id] = node
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            out[node.target.id] = node
    return out


def ast_check(rev: str) -> int:
    original = top_level_defs(git_show(rev, "streamlit_app.py"))
    identical = unexpected = missing = expected = 0
    for rel, names in MOVED.items():
        new_defs = top_level_defs((ROOT / rel).read_text(encoding="utf-8"))
        for name in names:
            old_node, new_node = original.get(name), new_defs.get(name)
            if old_node is None or new_node is None:
                print(f"MISSING  {name}: original={'yes' if old_node else 'NO'} new={'yes' if new_node else 'NO'} ({rel})")
                missing += 1
                continue
            if ast.dump(old_node) == ast.dump(new_node):
                identical += 1
                continue
            is_expected = name in EXPECTED_DEVIATIONS
            expected += is_expected
            unexpected += not is_expected
            print(f"{'EXPECTED' if is_expected else 'UNEXPECTED'} DIFFERENCE {name} ({rel})"
                  + (f" - {EXPECTED_DEVIATIONS[name]}" if is_expected else ""))
            for line in list(difflib.unified_diff(
                ast.unparse(old_node).splitlines(), ast.unparse(new_node).splitlines(), "original", "new", lineterm="", n=1
            ))[:40]:
                print("    " + line)
    total = sum(len(v) for v in MOVED.values())
    print(f"\nchecked {total} definitions: {identical} AST-identical, {expected} expected deviations, "
          f"{unexpected} UNEXPECTED, {missing} missing")
    return 1 if (unexpected or missing) else 0


def _impls(rev: str):
    from tests._ask_data_harness import NewImpl, OracleImpl
    original = OracleImpl(git_show(rev, "streamlit_app.py"), f"original@{rev}")
    refactored_app = OracleImpl((ROOT / "streamlit_app.py").read_text(encoding="utf-8"), "refactored-app", delegate=NewImpl())
    return original, refactored_app, NewImpl()


def golden(rev: str) -> int:
    from tests._ask_data_harness import OracleImpl, run_scenarios
    result = run_scenarios(OracleImpl(git_show(rev, "streamlit_app.py"), f"original@{rev}"))
    GOLDEN.parent.mkdir(parents=True, exist_ok=True)
    GOLDEN.write_text(json.dumps(result, indent=1, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")
    counts = {k: (len(v) if isinstance(v, dict) else 1) for k, v in result.items()}
    print(f"wrote {GOLDEN.relative_to(ROOT)} from original@{rev}: {counts}")
    return 0


def compare(rev: str) -> int:
    from tests._ask_data_harness import diff_results, run_scenarios
    original, app, new = _impls(rev)
    results = {impl.name: run_scenarios(impl) for impl in (original, app, new)}
    scenarios = sum(len(v) if isinstance(v, dict) else 1 for v in results[original.name].values())
    bad = 0
    for other in (app, new):
        diffs = diff_results(results[original.name], results[other.name])
        print(f"original@{rev}  vs  {other.name}: {len(diffs)} difference(s) over {scenarios} scenario groups")
        for line in diffs[:30]:
            print("   ", line)
        bad += len(diffs)
    return 1 if bad else 0


def _live_battery(projects, invoices):
    """Tool calls derived from live data (all argument values are real countries/projects/years/invoice numbers)."""
    from tests._ask_data_harness import D
    countries = sorted({p.country for p in projects if p.country})
    years = sorted({i.year for i in invoices if i.year})[-3:]
    unpaid_by_project: dict[str, float] = {}
    for inv in invoices:
        if inv.is_unpaid():
            unpaid_by_project[inv.project_name] = unpaid_by_project.get(inv.project_name, 0.0) + float(inv.payment_amount)
    top_projects = [n for n, _ in sorted(unpaid_by_project.items(), key=lambda kv: -kv[1])[:4]]
    numbers = sorted({int(i.invoice_number) for i in invoices if i.invoice_number})
    pick_numbers = [numbers[-1], numbers[len(numbers) // 2], 8684, 8683, 8713, 99999]
    cases: list[tuple[str, dict]] = []
    for n in pick_numbers:
        cases.append(("get_invoice", {"invoice_number": n}))
    for tool in ("get_invoices", "get_unpaid_invoices", "get_debt", "get_paid_amount", "get_outstanding_amount"):
        cases.append((tool, {}))
        for c in countries:
            cases.append((tool, {"country": c}))
        for y in years:
            cases.append((tool, {"year": y}))
        for p in top_projects:
            cases.append((tool, {"project": p}))
        if tool in ("get_unpaid_invoices", "get_debt"):
            for cat in ("y1", "y2plus", "trial"):
                cases.append((tool, {"category": cat}))
    cases += [("get_invoices", {"paid": v}) for v in ("yes", "no", "cancelled")]
    cases += [("get_paid_amount", {"paid_year": y}) for y in years]
    cases += [("get_invoices", {"paid_year": y}) for y in years]
    cases += [("get_top_debt_projects", a) for a in ({}, {"limit": 5}, {"year": years[-1]})]
    cases += [("get_sent_invoices", a) for a in ({}, {"invoice_number": 8688})]
    cases += [("get_project", {"project": p}) for p in top_projects]
    for tool in ("get_projects", "count_projects"):
        cases += [(tool, {}), (tool, {"status": "All"}), (tool, {"status": "Cancelled"}), (tool, {"min_cameras": 8})]
        cases += [(tool, {"country": c}) for c in countries]
    cases += [("get_projects", {"month": "March"})]  # latent NameError, expected identical on both sides
    for agg in ("sum", "count_projects", "average_per_project", "group_by_country", "list_projects"):
        cases.append(("get_camera_statistics", {"aggregation": agg}))
        for ct in ("TopDown", "Backtray", "Pushout"):
            cases.append(("get_camera_statistics", {"aggregation": agg, "camera_type": ct}))
    for c in countries:
        cases.append(("get_camera_statistics", {"country": c}))
    for tool in ("get_licenses", "get_license_payment_status"):
        cases += [(tool, {}), (tool, {"license_status": "Expired"}), (tool, {"license_status": "Update Next Month"}),
                  (tool, {"project_status": "Active"}), (tool, {"eop_year": datetime.date.today().year}),
                  (tool, {"eop_month": "October"})]
        cases += [(tool, {"country": c}) for c in countries]
    cases += [("get_license_eop_reconciliation", a) for a in ({}, {"paid_year": 1}, {"paid_year": 2}, {"needs_update": "yes"})]
    cases += [("get_last_bank_payment", {}), ("get_last_bank_payment", {"project": top_projects[0]})]
    cases += [("get_next_invoice_number", {})]
    return cases, countries, top_projects


def live(rev: str) -> int:
    from services.ask_data.helpers import _normalize_project_status
    from services.supabase_service import load_invoices, load_projects
    from tests._ask_data_harness import (
        BASE_CFG, _gemini_ok, _make_post, diff_results, fake_streamlit, norm_value, snapshot_call,
    )

    projects, invoices = load_projects(), load_invoices()
    for p in projects:
        p.status = _normalize_project_status(getattr(p, "status", ""))
    print(f"live data: {len(projects)} projects, {len(invoices)} invoice rows (read-only)")
    original, app, new = _impls(rev)
    impls = (original, app, new)
    battery, countries, top_projects = _live_battery(projects, invoices)

    # ---- 1. every tool, three-way ------------------------------------------------------------
    results = {impl.name: {} for impl in impls}
    for impl in impls:
        with impl.env(freeze=False):
            for index, (tool, args) in enumerate(battery):
                key = f"{index:03d} {tool} {json.dumps(args, default=str, sort_keys=True)}"
                results[impl.name][key] = snapshot_call(lambda: impl.execute(tool, dict(args), projects, invoices, next_invoice=lambda: -1))
    bad = 0
    for other in (app, new):
        diffs = diff_results(results[original.name], results[other.name])
        print(f"[tools]      {len(battery)} live tool calls: original@{rev} vs {other.name}: {len(diffs)} difference(s)")
        for line in diffs[:20]:
            print("    ", line)
        bad += len(diffs)

    # ---- 2. validator, three-way ---------------------------------------------------------------
    names = sorted({p.project_name for p in projects})
    parsed_cases = []
    for name in names:
        parsed_cases.append({"tool": "get_project", "arguments": {"project": name}})
        parsed_cases.append({"tool": "get_debt", "arguments": {"project": name.lower().replace("-", " ").replace("  ", " ")}})
    for c in countries + [_c.lower() for _c in countries] + ["belgium", "germany", "holland", "netherlands", "blegium", "Atlantis"]:
        parsed_cases.append({"tool": "get_debt", "arguments": {"country": c}})
    for extra in ("intermarché Heusy", "CM Essen (coplementary)", "AD Delhaize", "PPS", "iRetailCheck", "Tom"):
        parsed_cases.append({"tool": "get_debt", "arguments": {"project": extra}})
    val = {impl.name: {} for impl in impls}
    for impl in impls:
        with impl.env(freeze=False):
            for index, parsed in enumerate(parsed_cases):
                tool_name, cleaned, error = impl.validate(parsed, projects)
                val[impl.name][f"{index:04d}"] = {"tool": tool_name, "cleaned": norm_value(cleaned), "error": error}
    for other in (app, new):
        diffs = diff_results(val[original.name], val[other.name])
        print(f"[validator]  {len(parsed_cases)} live validation calls: original@{rev} vs {other.name}: {len(diffs)} difference(s)")
        for line in diffs[:20]:
            print("    ", line)
        bad += len(diffs)

    # ---- 3. full smart path (mock Gemini -> validate -> execute), three-way -----------------------
    smart_cases = [(t, a) for (t, a) in battery if t != "get_next_invoice_number"][::3]
    smart_results = {impl.name: {} for impl in impls}
    for impl in impls:
        for index, (tool, args) in enumerate(smart_cases):
            payload = json.dumps({"tool": tool, "arguments": {k: (v.isoformat() if hasattr(v, "isoformat") else v) for k, v in args.items()}})
            legacy_calls: list = []

            def normalize(q, _calls=legacy_calls):
                _calls.append(("normalize", q))
                return None

            def answer(q, p, i, d, _calls=legacy_calls):
                _calls.append(("answer", q))
                return f"LEGACY[{q}]", None

            with fake_streamlit() as state, impl.env(cfg=BASE_CFG, freeze=False), \
                    mock.patch("requests.post", _make_post([_gemini_ok(payload)], [])):
                res = snapshot_call(lambda: impl.smart(f"live question {index}", projects, invoices, normalize, answer, next_invoice=lambda: -1))
            smart_results[impl.name][f"{index:03d} {tool}"] = {"result": res, "state": norm_value(state), "legacy_calls": legacy_calls}
    for other in (app, new):
        diffs = diff_results(smart_results[original.name], smart_results[other.name])
        print(f"[smart path] {len(smart_cases)} live planner->validator->executor runs: original@{rev} vs {other.name}: {len(diffs)} difference(s)")
        for line in diffs[:20]:
            print("    ", line)
        bad += len(diffs)

    # ---- 4. legacy keyword fallback (no Gemini key): original vs refactored app -----------------
    legacy_questions = [
        "How many cameras do we have in total?", "How many TopDown cameras are in Belgium?", "Which country has the most cameras?",
        "How much Y1 debt do we have in 2026?", "What is the Y2+ debt for 2026?", "Show unpaid invoices for AD Anderlecht",
        "Show invoice 8676", "Was invoice 8676 sent?", "Show invoice 8684", "Show invoice 8713", "How many active projects are in Belgium?",
        "How many active projects are in Germany?", "Show sent PDF invoices for 2026", "What is the next invoice number?",
        "Show top debt projects", "Show project details for AD Alpha", "What invoices exist for Edeka Herweg?",
        "Which projects are billed in April?", "debt in belgium", "unpaid germany", "hello",
    ]
    legacy = {original.name: {}, app.name: {}}
    for impl in (original, app):
        for question in legacy_questions:
            with fake_streamlit(), impl.env(cfg={"api_key": "", "model": "m"}, freeze=False):
                impl.ns["_data_path"] = "Supabase"
                legacy[impl.name][question] = snapshot_call(lambda: impl.ns["_answer_data_question_smart"](question, projects, invoices, []))
    diffs = diff_results(legacy[original.name], legacy[app.name])
    print(f"[legacy]     {len(legacy_questions)} keyword-matcher questions (no Gemini): original@{rev} vs {app.name}: {len(diffs)} difference(s)")
    for line in diffs[:20]:
        print("    ", line)
    bad += len(diffs)

    # ---- 5. the headline number -----------------------------------------------------------------
    unpaid_total = sum((Decimal(str(i.payment_amount)) for i in invoices if i.is_unpaid()), Decimal("0"))
    print("\nOutstanding check (current ledger)")
    print(f"  independent Decimal sum of unpaid rows : EUR {unpaid_total:,.2f}  ({sum(1 for i in invoices if i.is_unpaid())} rows)")
    for impl in impls:
        with impl.env(freeze=False):
            text, _ = impl.execute("get_outstanding_amount", {}, projects, invoices, next_invoice=lambda: -1)
            debt_text, _ = impl.execute("get_debt", {}, projects, invoices, next_invoice=lambda: -1)
        print(f"  {impl.name:22} get_outstanding_amount: {text}")
        print(f"  {'':22} get_debt              : {debt_text}")
    print(f"\nRESULT: {'ALL IDENTICAL' if not bad else str(bad) + ' DIFFERENCE(S)'}")
    return 1 if bad else 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)
    for name in ("ast-check", "golden", "compare", "live"):
        sub.add_parser(name).add_argument("--rev", default="main")
    args = parser.parse_args(argv)
    return {"ast-check": ast_check, "golden": golden, "compare": compare, "live": live}[args.cmd](args.rev)


if __name__ == "__main__":
    sys.exit(main())
