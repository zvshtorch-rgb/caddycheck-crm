"""Shared harness for the Ask Data characterization tests and the parity script.

It defines (1) a deterministic synthetic CRM fixture, (2) scenario matrices covering every Ask Data v2
tool, the validator, the Gemini planner and the router, and (3) two interchangeable runners:

  OracleImpl  - executes the *original* definitions, extracted by AST from a streamlit_app.py source
                (e.g. `git show main:streamlit_app.py`), exactly as they were written.
  NewImpl     - calls the extracted services/ask_data modules.

Running the same scenarios through both and comparing is the parity proof for Phase 1. Nothing here
touches the network, Supabase or any production data (Gemini and the bank-payment log are mocked).
"""
from __future__ import annotations

import ast
import builtins
import contextlib
import datetime
import json
import math
import sys
import types
from decimal import Decimal
from unittest import mock

import pandas as pd

from models.invoice import Invoice
from models.project import Project

FROZEN_TODAY = datetime.date(2026, 10, 20)
FIXED_LOADED_AT = datetime.datetime(2026, 10, 20, 9, 5, tzinfo=datetime.timezone.utc)   # shown as "Data as of 2026-10-20 09:05 UTC"


# --------------------------------------------------------------------------------------------------
# Frozen clock
# --------------------------------------------------------------------------------------------------
class _DateMeta(type):
    def __instancecheck__(cls, obj):
        return isinstance(obj, datetime.date)

    def __subclasscheck__(cls, sub):
        return issubclass(sub, datetime.date)


class FrozenDate(datetime.date, metaclass=_DateMeta):
    @classmethod
    def today(cls):
        return FROZEN_TODAY


FROZEN_DATETIME_MODULE = types.SimpleNamespace(
    date=FrozenDate, datetime=datetime.datetime, timedelta=datetime.timedelta, timezone=datetime.timezone,
)


# --------------------------------------------------------------------------------------------------
# Fake streamlit (the extracted code only uses `st.session_state`, imported lazily inside functions)
# --------------------------------------------------------------------------------------------------
@contextlib.contextmanager
def fake_streamlit():
    state: dict = {}
    module = types.ModuleType("streamlit")
    module.session_state = state
    previous = sys.modules.get("streamlit")
    sys.modules["streamlit"] = module
    try:
        yield state
    finally:
        if previous is None:
            sys.modules.pop("streamlit", None)
        else:
            sys.modules["streamlit"] = previous


# --------------------------------------------------------------------------------------------------
# Synthetic fixture
# --------------------------------------------------------------------------------------------------
def _dt(y, m, d):
    return datetime.datetime(y, m, d)


def build_fixture():
    projects = [
        Project("AD Alpha", country="Bel", num_cams=10, payment_month="March", installation_year=2024,
                topdown_cameras=6, backtray_cameras=4, status="Active", license_eop=_dt(2099, 3, 1)),
        Project("AD Beta", country="Bel", num_cams=8, payment_month="April", installation_year=2022,
                topdown_cameras=8, status="Active", license_eop=_dt(2020, 5, 1)),
        Project("CM Gamma", country="Bel", num_cams=5, payment_month="May", installation_year=2023,
                backtray_cameras=5, status="Offline"),
        Project("Proxy Delta", country="Bel", num_cams=3, payment_month="June", installation_year=2021,
                pushout_cameras=3, status="Cancelled", license_eop=_dt(2030, 1, 1)),
        Project("Edeka Epsilon - Town", country="Ger", num_cams=12, payment_month="June", installation_year=2025,
                pushout_cameras=12, status="Active", license_eop=_dt(2099, 6, 1)),
        Project("Plus Zeta", country="Hol", num_cams=4, payment_month="November", installation_year=2020,
                detection_type="2 topdown, 2 backtray", status="Installed", license_eop=_dt(2026, 11, 15)),
        Project("Spar Eta", country="Lux", num_cams=2, status="Trial"),
        Project("Intermarche Heusy", country="Bel", num_cams=6, payment_month="", installation_year=2026,
                topdown_cameras=6, status="Active", license_eop=_dt(2026, 10, 1)),
        Project("AD Theta", country="Bel", num_cams=7, payment_month="Foo", installation_year=2030,
                topdown_cameras=7, status="active", license_eop=_dt(2027, 1, 31)),
    ]
    inv = Invoice
    invoices = [
        inv(8001, "AD Alpha", "Y1", 7780.0, 10, None, "No", 2026),
        inv(8001, "Edeka Epsilon - Town", "Y1", 3890.0, 5, None, "No", 2026),
        inv(8002, "AD Alpha", "Y2", 2280.0, 10, _dt(2026, 3, 5), "Yes", 2026),
        inv(8003, "AD Beta", "Y5", 1824.0, 8, _dt(2025, 6, 2), "Yes", 2025),
        inv(8003, "AD Beta", "Y6", 1824.0, 8, _dt(2026, 6, 2), "Yes", 2026),
        inv(8004, "CM Gamma", "Y2", 1140.0, 5, None, "No", 2025),
        inv(8005, "Proxy Delta", "Y3", 684.0, 3, None, "cancelled", 2025),
        inv(8006, "Plus Zeta", "Y6", 912.0, 4, None, "No", 2026),
        inv(8006, "Plus Zeta", "Y6", 1000.5, 4, None, "No", 2026),
        inv(8007, "Spar Eta", "Paid Trial-0.5Y", 389.0, 2, _dt(2026, 2, 1), "Yes", 2026),
        inv(8008, "Intermarche Heusy", "Y1", 4668.0, 6, None, "No", 2026),
        inv(8009, "intermarche heusy", "Y1", 3112.0, 4, None, "No", 2026),
        inv(8010, "CM Essen (coplementary)", "Y1", 3112.0, 4, None, "No", 2026),
        inv(8011, "AD Theta", "Y1", 5446.0, 7, _dt(2025, 12, 31), "Yes", 2025),
        inv(8012, "AD Theta", "Y1", 778.0, 1, None, "No", 2024),
        inv(8013, "AD Alpha", "Credit", -2723.0, None, None, "No", 2026),
        inv(8014, "AD Alpha", "Y3", 456.0, 2, _dt(2025, 1, 15), "Yes", 2025),
        inv(None, "CM Gamma", "Y3", 228.0, 1, None, "No", None),
        inv(8015, "Edeka Epsilon - Town", "Y2", 2736.0, 12, _dt(2026, 7, 9), "Yes", 2026),
        inv(8016, "AD Beta", "Y2", 5016.0, 22, None, "No", 2026),
    ]
    return projects, invoices


SENT_LOG = [
    {"sent_at": "2026-05-14T08:58:25", "invoice_number": 8001, "month": "May", "year": 2026, "pdf_filename": "a.pdf",
     "recipients": ["a@example.com"], "subject": "Monthly Invoice AD Alpha"},
    {"sent_at": "2026-06-14T08:58:25", "invoice_number": 8002, "month": "June", "year": 2026, "pdf_filename": "b.pdf",
     "recipients": ["a@example.com", "b@example.com"], "subject": "Monthly Invoice AD Beta"},
    {"sent_at": "2025-01-14T08:58:25", "invoice_number": 8003, "month": "January", "year": 2025, "pdf_filename": "c.pdf",
     "recipients": [], "subject": "Historical invoice"},
]

BANK_PAYMENTS = [
    {"id": 3, "payment_date": "2026-09-28", "invoice_number": "8001", "applied_amount": 7780.0, "received_amount": 7700.0, "source_name": "p3.pdf"},
    {"id": 2, "payment_date": "2026-09-14", "invoice_number": None, "applied_amount": None, "received_amount": 912.0, "source_name": "p2.pdf"},
    {"id": 1, "payment_date": None, "invoice_number": None, "applied_amount": 0, "received_amount": 0, "source_name": None},
]
BANK_ALLOCATIONS = {
    3: [
        {"invoice_number": 8001, "project_name": "AD Alpha", "maintenance_year": "Y1", "amount_applied": 7780.0, "year": 2026},
        {"invoice_number": 8001, "project_name": "Edeka Epsilon - Town", "maintenance_year": "Y1", "amount_applied": 3890.0, "year": 2026},
    ],
    2: [],
    1: [],
}


# --------------------------------------------------------------------------------------------------
# Result normalisation (JSON-safe, order preserving)
# --------------------------------------------------------------------------------------------------
def norm_value(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, Decimal):
        return str(v)
    if hasattr(v, "summary") and hasattr(v, "display_text"):   # a ToolResult: the random result_id is not part of behaviour
        summary = dict(v.summary())
        summary.pop("result_id", None)
        return norm_value(summary)
    if v.__class__.__name__ == "ResultStore":
        return "<ResultStore>"
    if isinstance(v, float):
        return None if math.isnan(v) else v
    if isinstance(v, (int, str)):
        return v
    if isinstance(v, (datetime.datetime, pd.Timestamp)):
        return v.isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    if hasattr(v, "item"):
        return norm_value(v.item())
    if isinstance(v, (list, tuple)):
        return [norm_value(x) for x in v]
    if isinstance(v, dict):
        return {str(k): norm_value(x) for k, x in v.items()}
    return repr(v)


def snapshot_result(result):
    text, df = result
    snap = {"text": text, "df": None}
    if df is not None:
        snap["df"] = {
            "columns": [str(c) for c in df.columns],
            "records": [{str(k): norm_value(x) for k, x in row.items()} for row in df.to_dict(orient="records")],
        }
    return snap


def snapshot_call(fn):
    try:
        return snapshot_result(fn())
    except Exception as exc:  # exceptions are part of the behavior being pinned
        return {"exception": type(exc).__name__, "message": str(exc)}


# --------------------------------------------------------------------------------------------------
# Oracle: original definitions extracted by AST from a streamlit_app.py source
# --------------------------------------------------------------------------------------------------
ORACLE_ROOTS = [
    "_answer_data_question_smart", "_execute_ask_data_tool", "_validate_ask_data_tool_call",
    "_ask_data_tool_catalog_text", "_llm_parse_data_question",
]


def load_app_namespace(source: str, roots=ORACLE_ROOTS) -> dict:
    tree = ast.parse(source)
    top: dict[str, ast.AST] = {}
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            top.setdefault(node.name, node)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                for n in ast.walk(target):
                    if isinstance(n, ast.Name):
                        top.setdefault(n.id, node)
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            top.setdefault(node.target.id, node)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            for alias in node.names:
                top.setdefault((alias.asname or alias.name).split(".")[0], node)

    builtin_names = set(dir(builtins))

    def free_names(node: ast.AST) -> set[str]:
        loads = {n.id for n in ast.walk(node) if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load)}
        local: set[str] = set()
        for n in ast.walk(node):
            if isinstance(n, ast.arg):
                local.add(n.arg)
            elif isinstance(n, ast.Name) and isinstance(n.ctx, (ast.Store, ast.Del)):
                local.add(n.id)
            elif isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)) and n is not node:
                local.add(n.name)
            elif isinstance(n, ast.ExceptHandler) and n.name:
                local.add(n.name)
            elif isinstance(n, (ast.Import, ast.ImportFrom)):
                for alias in n.names:
                    local.add((alias.asname or alias.name).split(".")[0])
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            return loads - local - builtin_names
        return loads - builtin_names

    needed: dict[int, ast.AST] = {}
    seen: set[str] = set()
    stack = list(roots)
    while stack:
        name = stack.pop()
        if name in seen or name not in top:
            continue
        seen.add(name)
        node = top[name]
        needed[id(node)] = node
        for dep in free_names(node):
            if dep in top and dep not in seen:
                stack.append(dep)

    namespace: dict = {"__name__": "ask_data_oracle_app"}
    for node in sorted(needed.values(), key=lambda n: n.lineno):
        exec(compile(ast.Module(body=[node], type_ignores=[]), "<ask-data-app>", "exec"), namespace)
    return namespace


# --------------------------------------------------------------------------------------------------
# Implementations under test
# --------------------------------------------------------------------------------------------------
@contextlib.contextmanager
def _bank_env(bank):
    """The bank-payment tool imports these lazily inside the function, so patching the source modules works for both."""
    if bank is None:
        yield
        return

    def load_payments():
        if bank.get("load_error"):
            raise RuntimeError("bank log unavailable")
        return list(bank.get("payments", []))

    def load_allocs(payment_id):
        return list(bank.get("allocations", {}).get(payment_id, []))

    with mock.patch("config.settings.load_bank_payments_log", load_payments), \
            mock.patch("services.supabase_service.load_bank_payment_allocations", load_allocs):
        yield


class OracleImpl:
    """Original code, extracted from a streamlit_app.py source (the pre-refactor file, or the refactored working copy)."""

    def __init__(self, source: str, label: str = "oracle", delegate=None):
        self.name = label
        self.ns = load_app_namespace(source)
        # When the app source delegates to the extracted modules, patches must reach those modules too.
        self.delegate = delegate

    @contextlib.contextmanager
    def env(self, cfg=None, sent_log=None, bank=None, freeze=True):
        with contextlib.ExitStack() as stack:
            if self.delegate is not None:
                stack.enter_context(self.delegate.env(cfg=cfg, sent_log=sent_log, bank=bank, freeze=freeze))
            stack.enter_context(self._env_namespace(cfg, sent_log, bank, freeze))
            yield

    @contextlib.contextmanager
    def _env_namespace(self, cfg, sent_log, bank, freeze):
        ns = self.ns
        saved = {k: ns.get(k, _MISSING) for k in ("datetime", "get_gemini_config", "load_sent_invoices_log", "_data_path", "_get_next_invoice_number")}
        if freeze:
            ns["datetime"] = FROZEN_DATETIME_MODULE
        if cfg is not None:
            ns["get_gemini_config"] = lambda: dict(cfg)
        if sent_log is not None:
            ns["load_sent_invoices_log"] = lambda: list(sent_log)
        try:
            with _bank_env(bank):
                yield
        finally:
            for key, value in saved.items():
                if value is _MISSING:
                    ns.pop(key, None)
                else:
                    ns[key] = value

    def catalog(self):
        return self.ns["_ask_data_tool_catalog_text"]()

    def validate(self, parsed, projects):
        return self.ns["_validate_ask_data_tool_call"](parsed, projects)

    def execute(self, tool, args, projects, invoices, next_invoice=None):
        self.ns["_data_path"] = "SOURCE"
        self.ns["_get_next_invoice_number"] = lambda _inv, _src: next_invoice()
        return self.ns["_execute_ask_data_tool"](tool, args, projects, invoices, [])

    def parse(self, question):
        return self.ns["_llm_parse_data_question"](question)

    def smart(self, question, projects, invoices, normalize, answer, next_invoice=None):
        self.ns["_data_path"] = "SOURCE"
        self.ns["_get_next_invoice_number"] = lambda _inv, _src: next_invoice()
        self.ns["_llm_normalize_question"] = normalize
        self.ns["_answer_data_question"] = answer
        return self.ns["_answer_data_question_smart"](question, projects, invoices, [])


_MISSING = object()


class NewImpl:
    """The extracted services/ask_data modules."""

    name = "new"

    def __init__(self):
        from services.ask_data import executor, helpers, planner, router, tool_specs, validation
        self.m = types.SimpleNamespace(executor=executor, helpers=helpers, planner=planner, router=router,
                                       tool_specs=tool_specs, validation=validation)

    @contextlib.contextmanager
    def env(self, cfg=None, sent_log=None, bank=None, freeze=True):
        m = self.m
        patches = [mock.patch.dict("services.ask_data.data_clock._state", {"loaded_at": FIXED_LOADED_AT, "source": "test"})]
        if freeze:
            patches += [
                mock.patch.object(m.executor, "datetime", FROZEN_DATETIME_MODULE),
                mock.patch.object(m.helpers, "datetime", FROZEN_DATETIME_MODULE),
                mock.patch.object(m.validation, "datetime", FROZEN_DATETIME_MODULE),
            ]
        if cfg is not None:
            patches.append(mock.patch.object(m.planner, "get_gemini_config", lambda: dict(cfg)))
        if sent_log is not None:
            patches.append(mock.patch.object(m.executor, "load_sent_invoices_log", lambda: list(sent_log)))
        with contextlib.ExitStack() as stack:
            for p in patches:
                stack.enter_context(p)
            stack.enter_context(_bank_env(bank))
            yield

    def catalog(self):
        return self.m.tool_specs._ask_data_tool_catalog_text()

    def validate(self, parsed, projects):
        return self.m.validation._validate_ask_data_tool_call(parsed, projects)

    def execute(self, tool, args, projects, invoices, next_invoice=None):
        return self.m.executor._execute_ask_data_tool(tool, args, projects, invoices, [], next_invoice_number=next_invoice)

    def parse(self, question):
        return self.m.planner._llm_parse_data_question(question)

    def smart(self, question, projects, invoices, normalize, answer, next_invoice=None):
        return self.m.router._answer_data_question_smart(
            question, projects, invoices, [],
            legacy_normalize_question=normalize, legacy_answer=answer, next_invoice_number=next_invoice,
        )


# --------------------------------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------------------------------
D = datetime.date

TOOL_CASES: list[tuple[str, str, dict]] = [
    ("get_invoice/multi_row", "get_invoice", {"invoice_number": 8001}),
    ("get_invoice/year_filter", "get_invoice", {"invoice_number": 8001, "year": 2026}),
    ("get_invoice/year_filter_empty", "get_invoice", {"invoice_number": 8001, "year": 1999}),
    ("get_invoice/single_paid", "get_invoice", {"invoice_number": 8002}),
    ("get_invoice/single_unpaid", "get_invoice", {"invoice_number": 8004}),
    ("get_invoice/single_cancelled", "get_invoice", {"invoice_number": 8005}),
    ("get_invoice/missing", "get_invoice", {"invoice_number": 99999}),
    ("get_invoices/all", "get_invoices", {}),
    ("get_invoices/paid_yes", "get_invoices", {"paid": "yes"}),
    ("get_invoices/paid_no", "get_invoices", {"paid": "no"}),
    ("get_invoices/paid_cancelled", "get_invoices", {"paid": "cancelled"}),
    ("get_invoices/project", "get_invoices", {"project": "AD Alpha"}),
    ("get_invoices/year", "get_invoices", {"year": 2026}),
    ("get_invoices/paid_year", "get_invoices", {"paid_year": 2026}),
    ("get_invoices/country", "get_invoices", {"country": "Bel"}),
    ("get_invoices/country_unpaid", "get_invoices", {"country": "Ger", "paid": "no"}),
    ("get_invoices/empty", "get_invoices", {"project": "Spar Eta", "paid": "no"}),
    ("get_unpaid_invoices/all", "get_unpaid_invoices", {}),
    ("get_unpaid_invoices/y1", "get_unpaid_invoices", {"category": "y1"}),
    ("get_unpaid_invoices/y2plus", "get_unpaid_invoices", {"category": "y2plus"}),
    ("get_unpaid_invoices/trial", "get_unpaid_invoices", {"category": "trial"}),
    ("get_unpaid_invoices/country_year", "get_unpaid_invoices", {"country": "Bel", "year": 2026}),
    ("get_unpaid_invoices/project", "get_unpaid_invoices", {"project": "Plus Zeta"}),
    ("get_project/found", "get_project", {"project": "AD Alpha"}),
    ("get_project/missing", "get_project", {"project": "Nope"}),
    ("get_projects/default_active", "get_projects", {}),
    ("get_projects/country", "get_projects", {"country": "Bel"}),
    ("get_projects/status_all", "get_projects", {"status": "All"}),
    ("get_projects/status_offline", "get_projects", {"status": "Offline"}),
    ("get_projects/min_cameras", "get_projects", {"min_cameras": 5}),
    ("get_projects/empty", "get_projects", {"country": "Lux"}),
    ("get_projects/month_filter_latent_bug", "get_projects", {"month": "March"}),
    ("count_projects/default", "count_projects", {}),
    ("count_projects/country", "count_projects", {"country": "Bel"}),
    ("count_projects/status_cancelled", "count_projects", {"status": "Cancelled"}),
    ("count_projects/all_ger", "count_projects", {"status": "All", "country": "Ger"}),
    ("get_debt/all", "get_debt", {}),
    ("get_debt/country", "get_debt", {"country": "Bel"}),
    ("get_debt/year", "get_debt", {"year": 2026}),
    ("get_debt/y1", "get_debt", {"category": "y1"}),
    ("get_debt/project", "get_debt", {"project": "AD Alpha"}),
    ("get_debt/ger_y2plus", "get_debt", {"country": "Ger", "category": "y2plus"}),
    ("get_debt/empty", "get_debt", {"project": "Spar Eta"}),
    ("get_paid_amount/all", "get_paid_amount", {}),
    ("get_paid_amount/country", "get_paid_amount", {"country": "Bel"}),
    ("get_paid_amount/paid_year", "get_paid_amount", {"paid_year": 2025}),
    ("get_paid_amount/project_fully_paid", "get_paid_amount", {"project": "AD Theta", "year": 2025}),
    ("get_paid_amount/project_partial", "get_paid_amount", {"project": "AD Beta"}),
    ("get_paid_amount/nothing", "get_paid_amount", {"project": "Plus Zeta", "paid_year": 2030}),
    ("get_paid_amount/year", "get_paid_amount", {"year": 2026}),
    ("get_outstanding_amount/all", "get_outstanding_amount", {}),
    ("get_outstanding_amount/country", "get_outstanding_amount", {"country": "Ger"}),
    ("get_outstanding_amount/project", "get_outstanding_amount", {"project": "Intermarche Heusy"}),
    ("get_outstanding_amount/none", "get_outstanding_amount", {"project": "Spar Eta"}),
    ("get_outstanding_amount/year", "get_outstanding_amount", {"year": 2025}),
    ("get_top_debt_projects/default", "get_top_debt_projects", {}),
    ("get_top_debt_projects/year", "get_top_debt_projects", {"year": 2026}),
    ("get_top_debt_projects/limit2", "get_top_debt_projects", {"limit": 2}),
    ("get_top_debt_projects/limit0", "get_top_debt_projects", {"limit": 0}),
    ("get_top_debt_projects/empty", "get_top_debt_projects", {"year": 1999}),
    ("get_sent_invoices/all", "get_sent_invoices", {}),
    ("get_sent_invoices/number", "get_sent_invoices", {"invoice_number": 8001}),
    ("get_sent_invoices/year", "get_sent_invoices", {"year": 2026}),
    ("get_sent_invoices/project", "get_sent_invoices", {"project": "AD Alpha"}),
    ("get_sent_invoices/empty", "get_sent_invoices", {"invoice_number": 424242}),
    ("get_next_invoice_number/basic", "get_next_invoice_number", {}),
    ("get_camera_statistics/sum", "get_camera_statistics", {}),
    ("get_camera_statistics/country", "get_camera_statistics", {"country": "Bel"}),
    ("get_camera_statistics/topdown", "get_camera_statistics", {"camera_type": "TopDown"}),
    ("get_camera_statistics/backtray_bel", "get_camera_statistics", {"camera_type": "Backtray", "country": "Bel"}),
    ("get_camera_statistics/pushout", "get_camera_statistics", {"camera_type": "Pushout"}),
    ("get_camera_statistics/count_projects", "get_camera_statistics", {"aggregation": "count_projects"}),
    ("get_camera_statistics/average", "get_camera_statistics", {"aggregation": "average_per_project"}),
    ("get_camera_statistics/group_by_country", "get_camera_statistics", {"aggregation": "group_by_country"}),
    ("get_camera_statistics/group_by_country_topdown", "get_camera_statistics", {"aggregation": "group_by_country", "camera_type": "TopDown"}),
    ("get_camera_statistics/list_all_status", "get_camera_statistics", {"aggregation": "list_projects", "status": "All"}),
    ("get_camera_statistics/average_cancelled", "get_camera_statistics", {"status": "Cancelled", "aggregation": "average_per_project"}),
    ("get_camera_statistics/list_empty", "get_camera_statistics", {"country": "Lux", "aggregation": "list_projects"}),
    ("get_camera_statistics/group_empty", "get_camera_statistics", {"country": "Lux", "aggregation": "group_by_country"}),
    ("get_licenses/all", "get_licenses", {}),
    ("get_licenses/expired", "get_licenses", {"license_status": "Expired"}),
    ("get_licenses/active_bel", "get_licenses", {"license_status": "Active", "country": "Bel"}),
    ("get_licenses/project_status_active", "get_licenses", {"project_status": "Active"}),
    ("get_licenses/project_status_all", "get_licenses", {"project_status": "All"}),
    ("get_licenses/eop_year", "get_licenses", {"eop_year": 2099}),
    ("get_licenses/eop_month", "get_licenses", {"eop_month": "May"}),
    ("get_licenses/eop_date", "get_licenses", {"eop_date": D(2020, 5, 1)}),
    ("get_licenses/project", "get_licenses", {"project": "AD Beta"}),
    ("get_licenses/update_next_month", "get_licenses", {"license_status": "Update Next Month"}),
    ("get_licenses/empty", "get_licenses", {"eop_year": 1999}),
    ("get_license_payment_status/all", "get_license_payment_status", {}),
    ("get_license_payment_status/country", "get_license_payment_status", {"country": "Bel"}),
    ("get_license_payment_status/project", "get_license_payment_status", {"project": "AD Alpha"}),
    ("get_license_payment_status/expired", "get_license_payment_status", {"license_status": "Expired"}),
    ("get_license_payment_status/empty", "get_license_payment_status", {"eop_year": 1999}),
    ("get_license_eop_reconciliation/all", "get_license_eop_reconciliation", {}),
    ("get_license_eop_reconciliation/paid_year1", "get_license_eop_reconciliation", {"paid_year": 1}),
    ("get_license_eop_reconciliation/needs_update", "get_license_eop_reconciliation", {"needs_update": "yes"}),
    ("get_license_eop_reconciliation/project", "get_license_eop_reconciliation", {"project": "AD Alpha"}),
    ("get_license_eop_reconciliation/country", "get_license_eop_reconciliation", {"country": "Ger"}),
    ("get_license_eop_reconciliation/empty", "get_license_eop_reconciliation", {"project": "Nope"}),
    ("unknown_tool", "definitely_not_a_tool", {}),
    # Phase 3: the whitelisted debt_by_country tool
    ("debt_by_country/overview", "debt_by_country", {}),
    ("debt_by_country/year", "debt_by_country", {"year": 2026}),
    ("debt_by_country/category_y1", "debt_by_country", {"category": "y1"}),
    ("debt_by_country/single_country", "debt_by_country", {"country": "Bel"}),
    ("debt_by_country/compare", "debt_by_country", {"country": "Bel", "compare_country": "Ger"}),
    ("debt_by_country/compare_zero_side", "debt_by_country", {"country": "Ger", "compare_country": "Lux"}),
    ("debt_by_country/same_country_twice", "debt_by_country", {"country": "Bel", "compare_country": "Bel"}),
    ("debt_by_country/empty", "debt_by_country", {"year": 1999}),
]

BANK_CASES: list[tuple[str, dict, dict]] = [
    ("bank/no_payments", {"payments": []}, {}),
    ("bank/latest_with_allocations", {"payments": BANK_PAYMENTS, "allocations": BANK_ALLOCATIONS}, {}),
    ("bank/project_found", {"payments": BANK_PAYMENTS, "allocations": BANK_ALLOCATIONS}, {"project": "AD Alpha"}),
    ("bank/project_not_found", {"payments": BANK_PAYMENTS, "allocations": BANK_ALLOCATIONS}, {"project": "Spar Eta"}),
    ("bank/latest_without_allocations_applied", {"payments": BANK_PAYMENTS[:1], "allocations": {}}, {}),
    ("bank/latest_without_allocations_received", {"payments": BANK_PAYMENTS[1:2], "allocations": {}}, {}),
    ("bank/latest_nothing_recorded", {"payments": BANK_PAYMENTS[2:], "allocations": {}}, {}),
    ("bank/load_error", {"payments": [], "load_error": True}, {}),
]

VALIDATION_CASES: list[tuple[str, object]] = [
    ("not_a_dict", ["get_debt"]),
    ("none", None),
    ("missing_tool", {"arguments": {}}),
    ("unknown_tool", {"tool": "drop_table", "arguments": {}}),
    ("arguments_not_dict", {"tool": "get_debt", "arguments": "country=Bel"}),
    ("arguments_none", {"tool": "get_debt", "arguments": None}),
    ("no_arguments_key", {"tool": "get_next_invoice_number"}),
    ("tool_with_spaces", {"tool": "  get_debt  ", "arguments": {"country": "belgium"}}),
    ("country_code", {"tool": "get_debt", "arguments": {"country": "Bel"}}),
    ("country_full_name", {"tool": "get_debt", "arguments": {"country": "Belgium"}}),
    ("country_typo", {"tool": "get_debt", "arguments": {"country": "blegium"}}),
    ("country_unknown", {"tool": "get_debt", "arguments": {"country": "Atlantis"}}),
    ("project_exact", {"tool": "get_debt", "arguments": {"project": "AD Alpha"}}),
    ("project_fuzzy_case_punct", {"tool": "get_debt", "arguments": {"project": "ad-alpha"}}),
    ("project_substring", {"tool": "get_debt", "arguments": {"project": "Edeka"}}),
    ("project_unknown", {"tool": "get_debt", "arguments": {"project": "Zzzzzz Qqqq"}}),
    ("project_required_missing", {"tool": "get_project", "arguments": {}}),
    ("project_required_ok", {"tool": "get_project", "arguments": {"project": "Plus Zeta"}}),
    ("int_string", {"tool": "get_invoice", "arguments": {"invoice_number": "8001", "year": "2026"}}),
    ("int_float", {"tool": "get_top_debt_projects", "arguments": {"limit": 3.0}}),
    ("int_bad", {"tool": "get_invoice", "arguments": {"invoice_number": "abc"}}),
    ("required_int_missing", {"tool": "get_invoice", "arguments": {"year": 2026}}),
    ("enum_case_insensitive", {"tool": "get_invoices", "arguments": {"paid": "YES"}}),
    ("enum_invalid", {"tool": "get_invoices", "arguments": {"paid": "maybe"}}),
    ("category_enum", {"tool": "get_debt", "arguments": {"category": "Y2PLUS"}}),
    ("status_all", {"tool": "get_projects", "arguments": {"status": "all"}}),
    ("status_known_case", {"tool": "get_projects", "arguments": {"status": "offline"}}),
    ("status_unknown", {"tool": "get_projects", "arguments": {"status": "retired"}}),
    ("month_short", {"tool": "get_projects", "arguments": {"month": "Apr"}}),
    ("month_full_lower", {"tool": "get_licenses", "arguments": {"eop_month": "october"}}),
    ("month_bad", {"tool": "get_projects", "arguments": {"month": "Smarch"}}),
    ("date_ok", {"tool": "get_licenses", "arguments": {"eop_date": "2026-10-01"}}),
    ("date_with_time", {"tool": "get_licenses", "arguments": {"eop_date": "2026-10-01T12:00:00"}}),
    ("date_bad", {"tool": "get_licenses", "arguments": {"eop_date": "01/10/2026"}}),
    ("license_status_enum", {"tool": "get_licenses", "arguments": {"license_status": "update next month"}}),
    ("combined_license_filters", {"tool": "get_licenses", "arguments": {"license_status": "Expired", "project_status": "Active", "country": "Bel"}}),
    ("unknown_keys_ignored", {"tool": "get_debt", "arguments": {"country": "Bel", "sql": "DROP TABLE invoices", "limit": 5}}),
    ("reconciliation_args", {"tool": "get_license_eop_reconciliation", "arguments": {"paid_year": 2, "needs_update": "Yes"}}),
    ("camera_enums", {"tool": "get_camera_statistics", "arguments": {"camera_type": "topdown", "aggregation": "GROUP_BY_COUNTRY"}}),
    ("debt_by_country_args", {"tool": "debt_by_country", "arguments": {"country": "belgium", "compare_country": "Germany", "year": "2026", "category": "Y1"}}),
    ("debt_by_country_hebrew", {"tool": "debt_by_country", "arguments": {"country": "בלגיה", "compare_country": "גרמניה"}}),
    ("debt_by_country_bad_country", {"tool": "debt_by_country", "arguments": {"country": "Atlantis"}}),
]

API_KEY = "TESTKEY-123"
BASE_CFG = {"api_key": API_KEY, "model": "gemini-test-model"}


def _gemini_ok(text):
    return {"status": 200, "json": {"candidates": [{"content": {"parts": [{"text": text}]}, "finishReason": "STOP"}],
                                    "usageMetadata": {"totalTokenCount": 10}}}


PLANNER_CASES: list[tuple[str, dict, str, list]] = [
    ("no_api_key", {"api_key": "", "model": "m"}, "how much debt?", []),
    ("blank_question", BASE_CFG, "   ", []),
    ("valid_tool_call", BASE_CFG, "How much does Belgium owe?", [_gemini_ok('{"tool": "get_debt", "arguments": {"country": "Belgium"}}')]),
    ("custom_model_in_url", {"api_key": API_KEY, "model": "gemini-x"}, "q", [_gemini_ok('{"tool": "get_next_invoice_number"}')]),
    ("json_list_is_none", BASE_CFG, "q", [_gemini_ok("[1, 2]")]),
    ("invalid_json", BASE_CFG, "q", [_gemini_ok("not json at all")]),
    ("empty_text", BASE_CFG, "q", [_gemini_ok("")]),
    ("retry_after_500", BASE_CFG, "q", [{"status": 503, "json": {}}, _gemini_ok('{"tool": "get_next_invoice_number"}')]),
    ("two_500s", BASE_CFG, "q", [{"status": 500, "json": {}}, {"status": 502, "json": {}}]),
    ("connection_errors_redacted", BASE_CFG, "q", [
        {"raise": "ConnectionError", "message": f"HTTPSConnectionPool: failed ?key={API_KEY}&alt=json {API_KEY}"},
        {"raise": "ConnectionError", "message": f"HTTPSConnectionPool: failed ?key={API_KEY}&alt=json {API_KEY}"},
    ]),
    ("http_400", BASE_CFG, "q", [{"status": 400, "json": {}, "http_error": f"400 Client Error for url: https://x/y?key={API_KEY}"}]),
]

# router: (id, cfg, gemini responses, question, normalize_returns, kwargs)
ROUTER_CASES: list[tuple[str, dict, list, str, object]] = [
    ("tool_call_success", BASE_CFG, [_gemini_ok('{"tool": "get_debt", "arguments": {"country": "Bel"}}')], "Belgium debt", None),
    ("outstanding_total", BASE_CFG, [_gemini_ok('{"tool": "get_outstanding_amount", "arguments": {}}')], "total outstanding", None),
    ("next_invoice", BASE_CFG, [_gemini_ok('{"tool": "get_next_invoice_number", "arguments": {}}')], "next invoice?", None),
    ("unknown_tool_falls_back", BASE_CFG, [_gemini_ok('{"tool": "drop_table", "arguments": {}}')], "drop it", None),
    ("fallback_uses_normalized_question", BASE_CFG, [_gemini_ok('{"tool": "nope"}')], "weird phrasing", "Show top debt projects"),
    ("unresolvable_project_falls_back", BASE_CFG, [_gemini_ok('{"tool": "get_debt", "arguments": {"project": "Zzzzzz Qqqq"}}')], "q", None),
    ("tool_exception_falls_back", BASE_CFG, [_gemini_ok('{"tool": "get_projects", "arguments": {"month": "March"}}')], "projects billed in March", None),
    ("no_api_key_goes_straight_to_legacy", {"api_key": "", "model": "m"}, [], "anything", None),
    ("gemini_failure_goes_to_legacy", BASE_CFG, [{"status": 500, "json": {}}, {"status": 500, "json": {}}], "q", None),
    ("debt_by_country_comparison", BASE_CFG, [_gemini_ok('{"tool": "debt_by_country", "arguments": {"country": "Belgium", "compare_country": "Germany"}}')], "Compare Belgium and Germany debt", None),
    ("debt_by_country_overview", BASE_CFG, [_gemini_ok('{"tool": "debt_by_country", "arguments": {}}')], "How much debt do we have by country?", None),
]


class _FakeResponse:
    def __init__(self, spec):
        self.status_code = spec.get("status", 200)
        self._json = spec.get("json", {})
        self._http_error = spec.get("http_error")

    def json(self):
        return self._json

    def raise_for_status(self):
        import requests
        if self._http_error or self.status_code >= 400:
            raise requests.exceptions.HTTPError(self._http_error or f"{self.status_code} error")


def _make_post(responses, calls):
    queue = list(responses)

    def post(url, headers=None, json=None, timeout=None):
        import requests
        calls.append({"url": url, "headers": dict(headers or {}), "body": json, "timeout": timeout})
        if not queue:
            raise AssertionError("unexpected extra Gemini request")
        spec = queue.pop(0)
        if "raise" in spec:
            raise getattr(requests.exceptions, spec["raise"])(spec["message"])
        return _FakeResponse(spec)

    return post


def _prompt_fingerprint(body):
    import hashlib
    prompt = body["contents"][0]["parts"][0]["text"]
    return {
        "prompt_sha256": hashlib.sha256(prompt.encode("utf-8")).hexdigest(),
        "prompt_len": len(prompt),
        "prompt_tail": prompt[-120:],
        "generationConfig": body["generationConfig"],
    }


def run_scenarios(impl) -> dict:
    """Run every scenario against `impl`; returns a JSON-safe dict of observed behavior."""
    projects, invoices = build_fixture()
    out: dict = {"catalog": None, "tools": {}, "bank": {}, "validation": {}, "planner": {}, "router": {}}

    with impl.env(sent_log=SENT_LOG):
        out["catalog"] = impl.catalog()

        for case_id, tool, args in TOOL_CASES:
            out["tools"][case_id] = snapshot_call(lambda: impl.execute(tool, dict(args), projects, invoices, next_invoice=lambda: 8042))

        for case_id, parsed in VALIDATION_CASES:
            tool_name, cleaned, error = impl.validate(parsed, projects)
            out["validation"][case_id] = {"tool": tool_name, "cleaned": norm_value(cleaned), "error": error}

    for case_id, bank, args in BANK_CASES:
        with impl.env(bank=bank):
            out["bank"][case_id] = snapshot_call(lambda: impl.execute("get_last_bank_payment", dict(args), projects, invoices))

    for case_id, cfg, question, responses in PLANNER_CASES:
        calls: list = []
        with fake_streamlit() as state, impl.env(cfg=cfg), mock.patch("requests.post", _make_post(responses, calls)):
            try:
                result = norm_value(impl.parse(question))
                exc = None
            except Exception as e:  # the planner is expected to swallow everything
                result, exc = None, f"{type(e).__name__}: {e}"
        out["planner"][case_id] = {
            "result": result, "exception": exc, "state": norm_value(state), "posts": len(calls),
            "first_request": None if not calls else {
                "url": calls[0]["url"], "headers": calls[0]["headers"], "timeout": calls[0]["timeout"],
                **_prompt_fingerprint(calls[0]["body"]),
            },
        }

    for case_id, cfg, responses, question, normalized in ROUTER_CASES:
        calls = []
        legacy_calls: list = []

        def normalize(q):
            legacy_calls.append(("normalize", q))
            return normalized

        def answer(q, p, i, d):
            legacy_calls.append(("answer", q))
            return f"LEGACY[{q}]", None

        with fake_streamlit() as state, impl.env(cfg=cfg, sent_log=SENT_LOG), mock.patch("requests.post", _make_post(responses, calls)):
            result = snapshot_call(lambda: impl.smart(question, projects, invoices, normalize, answer, next_invoice=lambda: 8042))
        out["router"][case_id] = {"result": result, "state": norm_value(state), "legacy_calls": norm_value(legacy_calls), "posts": len(calls)}

    return out


def diff_results(a, b, path="") -> list[str]:
    """Human-readable differences between two run_scenarios() results (empty list = identical)."""
    diffs: list[str] = []
    if isinstance(a, dict) and isinstance(b, dict):
        for key in sorted(set(a) | set(b)):
            if key not in a:
                diffs.append(f"{path}/{key}: only in second")
            elif key not in b:
                diffs.append(f"{path}/{key}: only in first")
            else:
                diffs.extend(diff_results(a[key], b[key], f"{path}/{key}"))
    elif isinstance(a, list) and isinstance(b, list) and len(a) == len(b):
        for index, (x, y) in enumerate(zip(a, b)):
            diffs.extend(diff_results(x, y, f"{path}[{index}]"))
    elif a != b:
        diffs.append(f"{path}: {json.dumps(a, default=str)[:160]}  !=  {json.dumps(b, default=str)[:160]}")
    return diffs


_EUR = None


def _split_amounts(text: str):
    """(text with every euro amount replaced by a placeholder, [signed Decimal amounts])."""
    import re
    global _EUR
    if _EUR is None:
        _EUR = re.compile(r"(-?)€(-?)([\d,]+(?:\.\d+)?)")
    amounts = []

    def repl(match):
        value = Decimal(match.group(3).replace(",", ""))
        amounts.append(-value if (match.group(1) or match.group(2)) else value)
        return "€#"

    return _EUR.sub(repl, text), amounts


def equivalent_snapshot(old: dict, new: dict):
    """(True, reason) when `new` differs from `old` only by Phase 3 presentation: cents on amounts (each within
    EUR 1 of the old whole-euro figure), appended calculation/scope/freshness/caveat paragraphs, and exact Decimal
    table cells. Anything else (different wording, different rows, exceptions) is not equivalent."""
    if "exception" in old or "exception" in new:
        return False, "exception"
    headline = new["text"].split("\n\n")[0]
    old_shape, old_amounts = _split_amounts(old["text"])
    new_shape, new_amounts = _split_amounts(headline)
    if old_shape != new_shape or len(old_amounts) != len(new_amounts):
        return False, "wording changed"
    if any(abs(a - b) >= 1 for a, b in zip(old_amounts, new_amounts)):
        return False, "amount changed by EUR 1 or more"
    odf, ndf = old["df"], new["df"]
    if odf is None and ndf is None:
        return True, "headline cents/appended paragraphs only"
    if odf is None or ndf is None:
        return False, "table appeared/disappeared"
    if odf["columns"] != ndf["columns"] or len(odf["records"]) != len(ndf["records"]):
        return False, "table shape changed"
    for orow, nrow in zip(odf["records"], ndf["records"]):
        for column in odf["columns"]:
            a, b = orow[column], nrow[column]
            if isinstance(a, (int, float)) and not isinstance(a, bool) and isinstance(b, str):
                try:
                    if Decimal(repr(float(a))) != Decimal(b):
                        return False, f"cell {column} changed"
                except Exception:
                    return False, f"cell {column} changed"
            elif a != b:
                return False, f"cell {column} changed"
    return True, "cents/Decimal cells/appended paragraphs only"
