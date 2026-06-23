"""
Seamfix Financial Intelligence Suite
Streamlit Cloud deployment — serves 5 interactive dashboards + AI chat assistant.

Data sources:
  • Revenue data:  Google Sheet (live) or uploaded xlsx
  • Cash reports:  uploaded xlsx files or pre-loaded data/
  • Budget:        uploaded xlsx or pre-loaded data/
"""

import streamlit as st
import streamlit.components.v1 as components
import os, sys, shutil, tempfile, glob, io, urllib.request, urllib.parse, time, importlib.util
import json, base64, secrets as secrets_mod, hashlib
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# ── Page config ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Seamfix Financial Intelligence",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Constants ────────────────────────────────────────────────────────
APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
GENERATED_DIR = APP_DIR / "generated"
GENERATED_DIR.mkdir(exist_ok=True)

# ── Bobby usage log ───────────────────────────────────────────────────
# Google Sheet shared with the service account for query audit trail
BOBBY_LOG_SHEET_ID = "1c7QMZuV-YNDsmn1XYLJtx8pRyYi6g_wwdAHJ-D0cgtk"

GOOGLE_SHEET_ID = "1XKIE9eRP8H1AWpuMAJA0U8bM7pQ9o1jvoQobc6aUn5s"
GOOGLE_DRIVE_FOLDER_ID = "1vLq8m030d1ifL6nAVuo9LT5N9NSeGs9U"
COLLECTIONS_SHEET_ID = "17KE1n5_SOeDXaX96Xsa1JfAjNs_OZX8xu-wYDt4LpU8"
# gid of the "2026 CRITICAL REVENUE INFLOWS" tab. The workbook holds several
# near-identical tabs (Revenue Bridge - Data is a mirror, plus Closed/IAM), so we
# pin to this gid to always fetch the correct one regardless of tab renaming/order.
COLLECTIONS_GID = "1584269897"
REVENUE_FILENAME = "2026 Path to Revenue (1).xlsx"
BUDGET_FILENAME = "2026 LEAN BUDGET.xlsx"
COLLECTIONS_FILENAME = "2026 Collections Tracker.csv"

DASHBOARDS = {
    "Cash Overview": {
        "icon": "💰",
        "script": "generate_dashboard.py",
        "output": "dashboard.html",
        "description": "Weekly cash position, inflows/outflows, FX rates, and bank balances"
    },
    "Expense & Vendor Analysis": {
        "icon": "📊",
        "script": "generate_expense_dashboard.py",
        "output": "expense_dashboard.html",
        "description": "Category breakdown, vendor analysis, investment tracking"
    },
    "Budget vs Actual": {
        "icon": "📋",
        "script": "generate_budget_dashboard.py",
        "output": "budget_dashboard.html",
        "description": "₦5.1B annual budget mapped against actual weekly spend"
    },
    "Revenue & Fundability": {
        "icon": "🚀",
        "script": "generate_revenue_dashboard.py",
        "output": "revenue_dashboard.html",
        "description": "$10M revenue target, critical actions, pipeline cross-reference"
    },
    "Pipeline Intelligence": {
        "icon": "🎯",
        "script": "generate_pipeline_dashboard.py",
        "output": "pipeline_dashboard.html",
        "description": "Deal-level momentum, status vs trend signals, On Track / At Risk / Off Track breakdown"
    },
    "Collections Tracker": {
        "icon": "📥",
        "script": "generate_collections_dashboard.py",
        "output": "collections_dashboard.html",
        "description": "Critical revenue inflows, weekly movement, payment status, and urgent collection actions"
    },
    "Group Financials": {
        "icon": "🏦",
        "script": "generate_financial_report_dashboard.py",
        "output": "financial_report_dashboard.html",
        "description": "Consolidated P&L, profitability vs targets (net margin 10%, gross margin 70%), expense analysis, and revenue breakdowns"
    },
}


# ── Helper: import a generator module without running its main() ─────
def _import_generator(filename):
    """Load a generator .py file as a module (safe — does not execute main)."""
    spec = importlib.util.spec_from_file_location(filename, APP_DIR / f"{filename}.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# ── Chat context builder ─────────────────────────────────────────────
def build_chat_context(data_folder):
    """
    Build a structured text summary of all financial data for the chatbot.
    Imports generator parsing functions to avoid duplicating xlsx logic.
    Returns a plain-text string covering every dashboard: Cash Overview,
    Budget vs Actual, Pipeline Intelligence, Expense & Vendor, Collections
    Tracker and Group Financials. Each section guards on its own data source,
    so a missing file or module degrades that section only — never the whole context.
    """
    data_path = Path(data_folder)
    parts = []
    FX_RATE = 1450

    parts.append("=== SEAMFIX FINANCIAL INTELLIGENCE — DATA CONTEXT ===")
    parts.append(f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}. Exchange rate: ₦{FX_RATE:,} per $1.")
    parts.append("Currency note: ₦ = Nigerian Naira, $ = US Dollars.")

    # Import generator modules for their parsing functions
    try:
        gen_dash   = _import_generator("generate_dashboard")
        gen_pipe   = _import_generator("generate_pipeline_dashboard")
        gen_budget = _import_generator("generate_budget_dashboard")
    except Exception as e:
        return f"Error loading financial data modules: {e}"

    # The remaining dashboards are imported individually so a single bad module
    # never blanks Bobby's whole context — each section guards on its own import.
    def _try_import(name):
        try:
            return _import_generator(name)
        except Exception:
            return None
    gen_exp  = _try_import("generate_expense_dashboard")
    gen_coll = _try_import("generate_collections_dashboard")
    gen_fin  = _try_import("generate_financial_report_dashboard")

    fmt = gen_dash.fmt_naira  # shorthand

    # ── SECTION 1: CASH OVERVIEW ────────────────────────────────────
    cash_files = sorted(glob.glob(str(data_path / "Cash Report*.xlsx")))
    reports = []
    for f in cash_files:
        try:
            r = gen_dash.extract_report(f)
            if r:
                reports.append(r)
        except Exception:
            pass
    reports.sort(key=lambda x: x["date"])

    if reports:
        parts.append(f"\n--- CASH OVERVIEW ---")
        parts.append(f"Data: {len(reports)} weekly reports from {reports[0]['date_str']} to {reports[-1]['date_str']}.")

        latest = reports[-1]
        total_cash = latest.get("total_cash_ngn", 0)
        ngn        = latest.get("ngn_closing", 0)
        usd_ngn    = latest.get("usd_closing_ngn", 0)
        inv        = latest.get("investment_ngn", 0)

        parts.append(f"\nLatest cash position ({latest['date_str']}):")
        parts.append(f"  Total (NGN equiv, incl. investments): {fmt(total_cash)}")
        parts.append(f"  NGN liquid cash: {fmt(ngn)}")
        parts.append(f"  USD holdings (NGN equiv): {fmt(usd_ngn)}  (${usd_ngn / FX_RATE / 1e6:.2f}M)")
        if inv > 0:
            parts.append(f"  Investment portfolio: {fmt(inv)}")

        # Operational averages (exclude investment flows)
        op_inflows, op_burns = [], []
        for r in reports:
            inv_in  = sum(v for k, v in r.get("inflow_items",  {}).items() if gen_dash.is_investment_inflow(k))
            inv_out = sum(v for k, v in r.get("outflow_items", {}).items() if gen_dash.is_investment_outflow(k))
            op_inflows.append(r.get("total_inflow",  0) - inv_in)
            op_burns.append(r.get("total_outflow", 0) - inv_out)

        avg_op_inflow = sum(op_inflows) / len(op_inflows) if op_inflows else 0
        avg_op_burn   = sum(op_burns)   / len(op_burns)   if op_burns   else 0
        avg_net       = avg_op_inflow - avg_op_burn

        parts.append(f"\nWeekly averages ({len(reports)} weeks):")
        parts.append(f"  Operational inflow: {fmt(avg_op_inflow)}/week")
        parts.append(f"  Operational burn:   {fmt(avg_op_burn)}/week")
        parts.append(f"  Net cash flow:      {fmt(avg_net)}/week ({'outflow exceeds inflow' if avg_net < 0 else 'inflow exceeds outflow'})")

        forecast_expected = total_cash + avg_net * 4
        forecast_floor    = total_cash - avg_op_burn * 4
        parts.append(f"\n4-Week Forecast:")
        parts.append(f"  Expected (avg net flow continues): {fmt(forecast_expected)}")
        parts.append(f"  Floor (zero inflows, burn only):   {fmt(forecast_floor)}")

        parts.append(f"\nWeekly cash history (oldest → newest):")
        for r in reports:
            net = r.get("total_inflow", 0) - r.get("total_outflow", 0)
            parts.append(
                f"  {r['date_str']}: Total={fmt(r.get('total_cash_ngn', 0))}, "
                f"Inflow={fmt(r.get('total_inflow', 0))}, "
                f"Outflow={fmt(r.get('total_outflow', 0))}, "
                f"Net={fmt(net)}"
            )

    # ── SECTION 2: BUDGET vs ACTUAL ─────────────────────────────────
    # Driven by the committed Budget Tracker snapshot (lean mode, all NGN) — the
    # same source as the Budget vs Actual dashboard (rewritten 2026-06-11). The old
    # LEAN-BUDGET + cash-outflow fuzzy-matching approach (and gen_budget.BUDGET_CATEGORIES)
    # was retired, so Bobby reads the snapshot via gen_budget.compute() here.
    snapshot_file = data_path / gen_budget.SNAPSHOT_NAME
    if snapshot_file.exists():
        departments = []
        try:
            with open(snapshot_file, "r", encoding="utf-8") as fh:
                snapshot = json.load(fh)
            departments, entities, group, elapsed = gen_budget.compute(snapshot)
        except Exception as e:
            parts.append(f"\n--- BUDGET vs ACTUAL ---")
            parts.append(f"(Budget snapshot could not be read: {e})")

        if departments:
            bfmt = gen_budget.fmt_naira
            fy           = snapshot.get("fiscalYear", "")
            last_actuals = snapshot.get("lastActualsMonth", "")
            run_date     = snapshot.get("runDate", "")

            parts.append(f"\n--- BUDGET vs ACTUAL ---")
            parts.append(
                f"Source: Seamfix Budget Tracker snapshot (lean mode, all figures NGN). "
                f"FY {fy}, actuals through {last_actuals} ({elapsed} of 12 months elapsed), "
                f"snapshot dated {run_date}."
            )
            parts.append(f"Group Annual Budget: {bfmt(group['annual_budget'])}")
            parts.append(f"Group YTD Actual: {bfmt(group['ytd_actual'])}")
            parts.append(f"Group YTD Budget (pace to date): {bfmt(group['ytd_budget'])}")
            g_over = group["ytd_actual"] > group["ytd_budget"]
            parts.append(
                f"Group Variance: {bfmt(abs(group['variance']))} "
                f"{'OVER' if g_over else 'UNDER'} pace "
                f"({group['pct_of_pace']:.1f}% of pace) — {bfmt(group['remaining'])} remaining of annual budget."
            )
            parts.append(f"Group Year-End Projection at current run rate: {bfmt(group['projected'])}")

            parts.append("\nBy entity (NGN):")
            parts.append(f"  {'Entity':<18} | {'Annual Budget':>14} | {'YTD Budget':>12} | {'YTD Actual':>12} | {'% of Pace':>9}")
            for ecode in sorted(entities.keys()):
                e = entities[ecode]
                label = gen_budget.ENTITY_LABEL.get(ecode, ecode)
                parts.append(
                    f"  {label:<18} | {bfmt(e['annual_budget']):>14} | {bfmt(e['ytd_budget']):>12} | "
                    f"{bfmt(e['ytd_actual']):>12} | {e['pct_of_pace']:>8.1f}%"
                )

            parts.append(f"\nBy department ({len(departments)} departments, sorted by YTD actual; NGN):")
            parts.append(
                f"  {'Department':<26} | {'Entity':<6} | {'Annual':>12} | "
                f"{'YTD Budget':>12} | {'YTD Actual':>12} | {'% Pace':>7} | Status"
            )
            for d in sorted(departments, key=lambda x: -x["ytd_actual"]):
                _, _, status_label = gen_budget.status_for(d["pct_of_pace"])
                parts.append(
                    f"  {d['name']:<26} | {d['entity']:<6} | {bfmt(d['annual_budget']):>12} | "
                    f"{bfmt(d['ytd_budget']):>12} | {bfmt(d['ytd_actual']):>12} | "
                    f"{d['pct_of_pace']:>6.1f}% | {status_label}"
                )
            parts.append(
                "(Budget is NGN, built bottom-up: Group = sum of entities = sum of departments. "
                "UK/UAE budgets are FX-converted to NGN; actuals are already NGN. For per-budget-head "
                "drill-down and individual expense transactions, direct the user to the Budget vs Actual dashboard.)"
            )

    # ── SECTION 3: PIPELINE INTELLIGENCE ────────────────────────────
    revenue_file = data_path / REVENUE_FILENAME
    if revenue_file.exists():
        parts.append(f"\n--- PIPELINE INTELLIGENCE ---")

        STATUS_WEIGHTS              = gen_pipe.STATUS_WEIGHTS
        STATUS_WEIGHTS_CONSERVATIVE = gen_pipe.STATUS_WEIGHTS_CONSERVATIVE
        LANDING_ZONE                = gen_pipe.LANDING_ZONE

        try:
            revenues = gen_pipe.extract_revenue_data(str(revenue_file))
        except Exception as e:
            revenues = []
            parts.append(f"(Error loading pipeline data: {e})")

        if revenues:
            realistic_proj    = sum(r["annual_usd"] * STATUS_WEIGHTS.get(r["status"], 0.5)              for r in revenues)
            conservative_proj = sum(r["annual_usd"] * STATUS_WEIGHTS_CONSERVATIVE.get(r["status"], 0.0) for r in revenues)
            realistic_gap     = LANDING_ZONE - realistic_proj
            total_ytd_usd     = sum(r["ytd"] for r in revenues)

            parts.append(f"Annual Revenue Target: ${LANDING_ZONE:,.0f}")
            parts.append(f"Realistic Projection (On Track=100%, At Risk=50%, Off Track=10%): ${realistic_proj:,.0f} ({realistic_proj / LANDING_ZONE * 100:.1f}%)")
            parts.append(f"Conservative Projection (At Risk & Off Track excluded): ${conservative_proj:,.0f} ({conservative_proj / LANDING_ZONE * 100:.1f}%)")
            parts.append(f"Realistic Gap to Target: ${realistic_gap:,.0f} {'(target met)' if realistic_gap <= 0 else '(shortfall)'}")
            parts.append(f"Total YTD Revenue Collected: ${total_ytd_usd:,.0f}")

            status_groups = defaultdict(list)
            for r in revenues:
                status_groups[r["status"]].append(r)

            parts.append(f"\nStatus breakdown:")
            for status in ["On Track", "Closed", "At Risk", "Off Track", "Unknown"]:
                deals = status_groups.get(status, [])
                if deals:
                    val = sum(d["annual_usd"] for d in deals)
                    ytd = sum(d["ytd"] for d in deals)
                    parts.append(f"  {status}: {len(deals)} deals, ${val:,.0f} annual target, ${ytd:,.0f} YTD collected")

            # Detect months with data dynamically
            _month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            _monthly_totals = [sum(r.get('monthly', [0]*12)[i] for r in revenues) for i in range(12)]
            _months_with_data = [i for i, v in enumerate(_monthly_totals) if v > 0]
            _last_month = max(_months_with_data) if _months_with_data else 0
            _num_months = _last_month + 1

            month_headers = ' | '.join(f'{_month_names[i]+" $":>9}' for i in range(_num_months))
            parts.append(f"\nAll deals ({len(revenues)} total, sorted by annual value desc):")
            parts.append(
                f"  {'Deal Name':<42} | {'Status':<12} | {'Annual $':>10} | "
                f"{month_headers} | {'YTD $':>9}"
            )
            for r in sorted(revenues, key=lambda x: -x["annual_usd"]):
                monthly = r.get('monthly', [0]*12)
                month_vals = ' | '.join(f'${monthly[i]:>8,.0f}' for i in range(_num_months))
                parts.append(
                    f"  {r['name']:<42} | {r['status']:<12} | ${r['annual_usd']:>9,.0f} | "
                    f"{month_vals} | ${r['ytd']:>8,.0f}"
                )

    # ── SECTION 4: EXPENSE & VENDOR ANALYSIS ────────────────────────
    # Parses the OUTFLOWS + payment-batch sections of every weekly cash report.
    if gen_exp is not None and reports:
        try:
            weekly, vendors, cats_by_week = gen_exp.process_all_files(str(data_path))
            ekpis = gen_exp.calculate_kpis(weekly, vendors, cats_by_week)
        except Exception as e:
            weekly, vendors, ekpis = [], {}, None
            parts.append(f"\n--- EXPENSE & VENDOR ANALYSIS ---\n(Error loading expense data: {e})")

        if ekpis and weekly:
            efmt = gen_exp.format_naira
            parts.append(f"\n--- EXPENSE & VENDOR ANALYSIS ---")
            parts.append(
                f"Total YTD Expenses ({ekpis['weeks']} weekly reports, excl. investment outflows): "
                f"{efmt(ekpis['total_ytd'])}"
            )
            parts.append(f"Average Weekly Burn Rate: {efmt(ekpis['avg_burn'])}")
            parts.append(f"Unique Vendors: {ekpis['unique_vendors']} ({ekpis['recurring_count']} recurring, 3+ payments)")
            parts.append(f"Largest Single Payment: {efmt(ekpis['largest_payment'])}")
            if ekpis['top_vendor']:
                parts.append(f"Top Vendor by Spend: {ekpis['top_vendor']} — {efmt(ekpis['top_vendor_spend'])}")

            # Aggregate spend by standardized category across all weeks
            cat_totals = defaultdict(float)
            for week_cats in cats_by_week.values():
                for cat, amt in week_cats.items():
                    if cat != "Investment Outflows":
                        cat_totals[cat] += amt
            if cat_totals:
                parts.append("\nSpend by category (YTD, excl. investment outflows):")
                for cat, amt in sorted(cat_totals.items(), key=lambda x: -x[1]):
                    parts.append(f"  {cat:<40} | {efmt(amt):>14}")

            # Top 15 vendors by total spend
            vendor_totals = [(v, sum(p["amount"] for p in pays)) for v, pays in vendors.items()]
            vendor_totals.sort(key=lambda x: -x[1])
            parts.append("\nTop 15 vendors by total spend (payment-batch detail):")
            for v, total in vendor_totals[:15]:
                parts.append(f"  {v:<40} | {efmt(total):>14}")
            parts.append(
                "(Vendor figures are the payment-batch disbursement detail, a subset of total outflows. "
                "For individual transactions, direct the user to the Expense & Vendor dashboard.)"
            )

    # ── SECTION 5: COLLECTIONS TRACKER ──────────────────────────────
    coll_file = data_path / COLLECTIONS_FILENAME
    if gen_coll is not None and coll_file.exists():
        try:
            items = gen_coll.extract_collections(str(coll_file))
        except Exception as e:
            items = []
            parts.append(f"\n--- COLLECTIONS TRACKER ---\n(Error loading collections data: {e})")

        if items:
            cfmt = gen_coll.fmt_usd
            total_usd = sum(it["usd"] for it in items)
            booked_usd = sum(it["usd"] for it in items if it.get("booked") == "YES")
            parts.append(f"\n--- COLLECTIONS TRACKER (Critical Revenue Inflows) ---")
            parts.append(f"{len(items)} tracked deals, total {cfmt(total_usd)} expected; {cfmt(booked_usd)} booked.")
            parts.append("\nDeals (sorted by USD value desc):")
            parts.append(
                f"  {'Deal / Customer':<34} | {'USD':>10} | {'Booked':<6} | {'Predict.':<8} | "
                f"{'Closure':<12} | Payment Status"
            )
            for it in sorted(items, key=lambda x: -x["usd"]):
                label = (it.get("name") or it.get("customer") or "")[:34]
                parts.append(
                    f"  {label:<34} | {cfmt(it['usd']):>10} | {it.get('booked',''):<6} | "
                    f"{it.get('predictability',''):<8} | {(it.get('closure_period') or '')[:12]:<12} | "
                    f"{it.get('payment_status','')}"
                )
            parts.append(
                "(For weekly update commentary and per-deal actions, direct the user to the Collections Tracker dashboard.)"
            )

    # ── SECTION 6: GROUP FINANCIALS (Consolidated P&L) ──────────────
    if gen_fin is not None:
        report_file = gen_fin.find_file(str(data_path), "Group Financial Report")
        m = None
        if report_file:
            try:
                m = gen_fin.extract_financials(report_file)
            except Exception as e:
                parts.append(f"\n--- GROUP FINANCIALS ---\n(Error loading financial report: {e})")
        if m:
            def _ngn(key):
                return fmt(m.get(key, {}).get("ngn", 0))
            def _ngn_abs(key):  # cost lines are stored negative on the Summary tab
                return fmt(abs(m.get(key, {}).get("ngn", 0)))
            def _usd(key):
                return f"${m.get(key, {}).get('usd', 0):,.0f}"
            def _pct(key):
                return f"{m.get(key, {}).get('ngn', 0) * 100:.1f}%"
            period = m.get("cur_date")
            period_str = period.strftime("%d %b %Y") if hasattr(period, "strftime") else "latest period"
            parts.append(f"\n--- GROUP FINANCIALS (Consolidated P&L, YTD as at {period_str}) ---")
            parts.append("All figures NGN with USD equivalent (report's own period-average FX).")
            parts.append(f"Total Revenue: {_ngn('revenue')} ({_usd('revenue')})")
            parts.append(f"Cost of Sales: {_ngn_abs('cogs')}")
            parts.append(f"Gross Profit: {_ngn('gross_profit')} — Gross Margin {_pct('gross_margin')} (target 70%)")
            parts.append(f"Total Operating Expenses: {_ngn_abs('opex')}")
            parts.append(f"EBITDA: {_ngn('ebitda')} — EBITDA Margin {_pct('ebitda_margin')}")
            parts.append(f"Net Profit (Profit After Tax): {_ngn('pat')} ({_usd('pat')}) — Net Margin {_pct('net_margin')} (target 10%)")
            parts.append(f"Payroll: {_ngn('payroll')} ({_pct('payroll_pct')} of revenue); Marketing: {_ngn('marketing')} ({_pct('marketing_pct')} of revenue)")

            for sec_key, sec_label in [("by_vertical", "Revenue by Vertical"),
                                       ("by_customer", "Revenue by Customer"),
                                       ("by_country", "Revenue by Country")]:
                rows = m.get(sec_key, [])
                if rows:
                    parts.append(f"\n{sec_label}:")
                    for it in sorted(rows, key=lambda x: -x.get("ngn", 0)):
                        parts.append(f"  {it['name']:<36} | {fmt(it.get('ngn', 0)):>14}")

            parts.append("\nBalance-sheet highlights:")
            parts.append(f"  Cash & equivalents: {_ngn('cash')}; Receivables: {_ngn('receivables')}; Total Assets: {_ngn('total_assets')}")
            parts.append(
                "(Segment/customer/country breakdowns exclude Other Income and don't reconcile to Total Revenue. "
                "For full ratios, EVA and the monthly trend, direct the user to the Group Financials dashboard.)"
            )

    parts.append(f"\n=== END OF CONTEXT ===")
    return "\n".join(parts)


# ── Chat context cache (shared across all sessions, cleared on regenerate) ──
@st.cache_resource(show_spinner=False)
def _get_chat_context_cache():
    return {}


def get_chat_context(data_folder):
    """Return the cached context, building it fresh if needed."""
    cache = _get_chat_context_cache()
    if "context" not in cache:
        try:
            cache["context"] = build_chat_context(data_folder)
        except Exception as e:
            cache["context"] = f"Financial data context could not be loaded: {e}"
    return cache["context"]


# ── Claude API call ─────────────────────────────────────────────────
def call_claude(messages, context):
    """
    Call Claude Sonnet via the Anthropic API.
    Uses prompt caching on the system context to reduce cost on subsequent turns.
    """
    try:
        import anthropic
    except ImportError:
        # Return a full 5-tuple — the caller always unpacks 5 values, so a bare
        # string here would raise a ValueError and crash the chat instead of
        # surfacing this message.
        return "⚠️ The `anthropic` package is not installed. Add it to requirements.txt.", 0, 0, 0, 0

    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return (
            "⚠️ `ANTHROPIC_API_KEY` not found in Streamlit secrets. "
            "Add it under Settings → Secrets in Streamlit Cloud.",
            0, 0, 0, 0,
        )

    system_prompt = f"""You are Bobby, a senior financial analyst embedded in the Seamfix Financial Intelligence Suite. You have full access to Seamfix's live financial data, updated weekly.

Your role: answer questions from Seamfix's CEO, executive team, and finance team about the company's cash position, budget performance, revenue pipeline, and expenses. Be direct, precise, and use the actual numbers from the data. When the situation warrants concern, say so clearly.

FINANCIAL DATA:
{context}

GUIDELINES:
- Answer only from the data above. Never invent or estimate numbers not in the context.
- Cross-reference across sections when useful — e.g. link pipeline gap to cash runway, or budget headroom to a hiring decision.
- For individual expense transactions not in the top 50 list, direct the user to the Expense Details tab on the Budget vs Actual dashboard.
- For data that doesn't exist (P&L, accounts receivable, headcount, year-over-year), say so clearly and name what data source would be needed.
- Use ₦ for Naira, $ for USD. Be concise. Use bullet points for lists, prose for explanations.
- If a follow-up question narrows on something from your previous answer, use the prior context — do not re-summarise unnecessarily."""

    try:
        client = anthropic.Anthropic(api_key=api_key)

        # Build message list (role/content pairs only — system handled separately)
        api_messages = [{"role": m["role"], "content": m["content"]} for m in messages]

        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system=[
                {
                    "type": "text",
                    "text": system_prompt,
                    "cache_control": {"type": "ephemeral"},  # Prompt caching — 90% cost reduction on repeated turns
                }
            ],
            messages=api_messages,
        )
        text           = response.content[0].text
        in_tok         = getattr(response.usage, "input_tokens",          0)
        cache_read_tok = getattr(response.usage, "cache_read_input_tokens",    0)
        cache_write_tok= getattr(response.usage, "cache_creation_input_tokens", 0)
        out_tok        = getattr(response.usage, "output_tokens",         0)
        return text, in_tok, cache_read_tok, cache_write_tok, out_tok

    except Exception as e:
        return f"⚠️ API error: {str(e)}", 0, 0, 0, 0


# ── Shared helper: open (or create) a named worksheet ────────────────
def _get_or_create_worksheet(spreadsheet, title, headers):
    """
    Return the worksheet named *title*, creating it (with *headers*) if absent.
    Also renames Sheet1/Sheet2 placeholders on first use.
    Silent on failure — returns None.
    """
    try:
        try:
            ws = spreadsheet.worksheet(title)
        except Exception:
            ws = spreadsheet.add_worksheet(title=title, rows=5000, cols=len(headers))

        # Write headers if the sheet is brand-new
        if not ws.cell(1, 1).value:
            ws.insert_row(headers, index=1)
        return ws
    except Exception:
        return None


# ── Bobby query logger ────────────────────────────────────────────────
def log_bobby_query(user_email, question, response,
                    input_tokens=0, cache_read_tokens=0, cache_write_tokens=0, output_tokens=0):
    """
    Append one row to the 'Bobby Queries' worksheet.

    Token columns explained:
      Non-Cached Input  — new tokens sent this turn (question + conversation history)
      Cache Read        — context tokens served from the prompt cache (~20k, charged at ~10%)
      Cache Write       — tokens written to cache on the very first turn of a session
      Output            — tokens in Bobby's reply
    Silent on any failure — never crash the chat.
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        gc = gspread.authorize(creds)
        spreadsheet = gc.open_by_key(BOBBY_LOG_SHEET_ID)

        headers = [
            "Timestamp", "User Email", "Question",
            "Response (first 300 chars)",
            "Non-Cached Input Tokens", "Cache Read Tokens",
            "Cache Write Tokens", "Output Tokens",
        ]
        ws = _get_or_create_worksheet(spreadsheet, "Bobby Queries", headers)
        if ws is None:
            return

        ws.append_row([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            user_email or "unknown",
            question,
            response[:300] if response else "",
            input_tokens,
            cache_read_tokens,
            cache_write_tokens,
            output_tokens,
        ])
    except Exception:
        pass


# ── Dashboard visit logger ────────────────────────────────────────────
def log_dashboard_visit(user_email):
    """
    Append one row to the 'Dashboard Visits' worksheet.
    Call once per browser session (guard with st.session_state).
    Silent on any failure.
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        gc = gspread.authorize(creds)
        spreadsheet = gc.open_by_key(BOBBY_LOG_SHEET_ID)

        headers = ["Timestamp", "User Email"]
        ws = _get_or_create_worksheet(spreadsheet, "Dashboard Visits", headers)
        if ws is None:
            return

        ws.append_row([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            user_email or "unknown",
        ])
    except Exception:
        pass


# ── Custom OAuth state store (shared across all WebSocket sessions) ──
@st.cache_resource
def _get_oauth_state_store():
    """Global dict of {state_token: creation_timestamp}. Survives page reloads."""
    return {}


def _cleanup_expired_states(store, max_age=600):
    """Remove states older than max_age seconds."""
    now = time.time()
    expired = [k for k, v in store.items() if now - v > max_age]
    for k in expired:
        del store[k]


# ── Helper: Google OAuth check ───────────────────────────────────────
def check_auth():
    if os.getenv("DISABLE_AUTH", "false").lower() == "true":
        return True

    # ── Read custom OAuth config (NOT [auth] — that triggers Streamlit's broken built-in) ──
    oauth_conf = st.secrets.get("google_oauth", {})
    client_id     = oauth_conf.get("client_id", "")
    client_secret = oauth_conf.get("client_secret", "")
    redirect_uri  = oauth_conf.get("redirect_uri", "")
    allowed_emails = [e.lower().strip() for e in oauth_conf.get("allowed_emails", [])]

    if not (client_id and client_secret and redirect_uri):
        return True  # OAuth not configured — allow access

    # ── Already authenticated this session? ──────────────────────────
    if "authenticated_email" in st.session_state:
        email = st.session_state.authenticated_email
        if allowed_emails and email not in allowed_emails:
            st.error(
                f"Access denied. **{email}** is not on the authorized users list.\n\n"
                "Contact the dashboard administrator to request access."
            )
            if st.button("Sign out"):
                del st.session_state["authenticated_email"]
                st.rerun()
            st.stop()
            return False
        return True

    # ── Handle OAuth callback (Google redirected back with ?code=&state=) ──
    params = st.query_params
    code  = params.get("code")
    state = params.get("state")

    if code and state:
        store = _get_oauth_state_store()
        _cleanup_expired_states(store)

        if state not in store:
            st.error("Authentication failed — session expired or invalid state. Please try again.")
            if st.button("🔄 Try again"):
                st.query_params.clear()
                st.rerun()
            st.stop()
            return False

        # State is valid — remove it (one-time use)
        del store[state]

        # Exchange authorization code for tokens
        try:
            token_body = urllib.parse.urlencode({
                "code":          code,
                "client_id":     client_id,
                "client_secret": client_secret,
                "redirect_uri":  redirect_uri,
                "grant_type":    "authorization_code",
            }).encode("utf-8")

            req = urllib.request.Request(
                "https://oauth2.googleapis.com/token",
                data=token_body,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            )
            resp = urllib.request.urlopen(req, timeout=15)
            tokens = json.loads(resp.read())

            # Decode the id_token JWT payload (base64url, no verification needed
            # since we just got it directly from Google over HTTPS)
            id_token = tokens.get("id_token", "")
            parts = id_token.split(".")
            if len(parts) < 2:
                raise ValueError("Invalid id_token format")

            payload_b64 = parts[1]
            # Fix base64url padding
            payload_b64 += "=" * (4 - len(payload_b64) % 4)
            user_info = json.loads(base64.urlsafe_b64decode(payload_b64))
            email = user_info.get("email", "").lower().strip()

            if not email:
                raise ValueError("No email in Google response")

            # Store authenticated email in session
            st.session_state.authenticated_email = email

            # Clear query params and reload cleanly
            st.query_params.clear()
            st.rerun()

        except Exception as e:
            st.error(f"Authentication error: {e}")
            if st.button("🔄 Try again"):
                st.query_params.clear()
                st.rerun()
            st.stop()
            return False

    # ── Show login page ──────────────────────────────────────────────
    st.markdown(
        """<div style="text-align:center;padding:80px 20px">
            <h1 style="color:#00D4AA;margin-bottom:8px">Seamfix Financial Intelligence</h1>
            <p style="color:#94a3b8;margin-bottom:40px">Sign in with your authorized Google account to access the dashboards.</p>
        </div>""",
        unsafe_allow_html=True,
    )

    # Generate a CSRF state token and store it globally
    oauth_state = secrets_mod.token_urlsafe(32)
    store = _get_oauth_state_store()
    _cleanup_expired_states(store)
    store[oauth_state] = time.time()

    auth_url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode({
        "client_id":     client_id,
        "redirect_uri":  redirect_uri,
        "response_type": "code",
        "scope":         "openid email profile",
        "state":         oauth_state,
        "access_type":   "offline",
        "prompt":        "select_account",
    })

    st.link_button("🔐 Sign in with Google", auth_url, use_container_width=True)
    st.stop()
    return False


# ── Helper: fetch Google Sheet as xlsx ───────────────────────────────
@st.cache_data(ttl=300, show_spinner="Fetching live revenue data from Google Sheet...")
def fetch_google_sheet_xlsx(sheet_id):
    errors = []
    try:
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
        req  = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        resp = urllib.request.urlopen(req, timeout=30)
        return resp.read()
    except Exception as e:
        errors.append(f"Public access: {e}")

    try:
        import gspread
        from google.oauth2.service_account import Credentials
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets.readonly",
                "https://www.googleapis.com/auth/drive.readonly",
            ],
        )
        gc          = gspread.authorize(creds)
        spreadsheet = gc.open_by_key(sheet_id)
        return spreadsheet.export(format=gspread.utils.ExportFormat.EXCEL)
    except KeyError:
        errors.append("Service account: No gcp_service_account in secrets")
    except Exception as e:
        errors.append(f"Service account: {e}")

    st.session_state["_gsheet_errors"] = errors
    return None


# ── Helper: fetch a single Google Sheet tab (by gid) as CSV bytes ────
@st.cache_data(ttl=300, show_spinner="Fetching live collections data from Google Sheet...")
def fetch_google_sheet_csv(sheet_id, gid):
    """
    Fetch one specific worksheet tab (pinned by gid) as CSV bytes.
    Tries public export first, then falls back to the service account.
    Pinning to gid guarantees the correct tab even if the workbook has
    several similarly-structured tabs.
    """
    errors = []
    try:
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        req  = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        resp = urllib.request.urlopen(req, timeout=30)
        return resp.read()
    except Exception as e:
        errors.append(f"Public access: {e}")

    try:
        import gspread, csv as _csv
        from google.oauth2.service_account import Credentials
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets.readonly",
                "https://www.googleapis.com/auth/drive.readonly",
            ],
        )
        gc          = gspread.authorize(creds)
        spreadsheet = gc.open_by_key(sheet_id)
        ws          = spreadsheet.get_worksheet_by_id(int(gid))
        rows        = ws.get_all_values()
        buf = io.StringIO()
        _csv.writer(buf).writerows(rows)
        return buf.getvalue().encode("utf-8")
    except KeyError:
        errors.append("Service account: No gcp_service_account in secrets")
    except Exception as e:
        errors.append(f"Service account: {e}")

    st.session_state["_collections_errors"] = errors
    return None


# ── Helper: fetch cash reports from Google Drive folder ──────────────
def fetch_drive_folder_files(folder_id):
    debug_log = []
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            creds_dict, scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        service = build("drive", "v3", credentials=creds)
        debug_log.append("Service account authenticated OK")

        folder_ids = [folder_id]
        subfolder_results = service.files().list(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id, name)", pageSize=50
        ).execute()
        subfolders = subfolder_results.get("files", [])
        for sf in subfolders:
            folder_ids.append(sf["id"])
            debug_log.append(f"Found subfolder: {sf['name']}")

        for sf in subfolders:
            nested_results = service.files().list(
                q=f"'{sf['id']}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
                fields="files(id, name)", pageSize=50
            ).execute()
            for nsf in nested_results.get("files", []):
                folder_ids.append(nsf["id"])
                debug_log.append(f"Found nested subfolder: {sf['name']}/{nsf['name']}")

        XLSX_MIME   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        GSHEET_MIME = "application/vnd.google-apps.spreadsheet"

        xlsx_files, gsheet_files = [], []
        for fid in folder_ids:
            for mime, bucket in [(XLSX_MIME, xlsx_files), (GSHEET_MIME, gsheet_files)]:
                results = service.files().list(
                    q=f"'{fid}' in parents and mimeType='{mime}' and trashed=false",
                    fields="files(id, name, mimeType)", pageSize=50
                ).execute()
                found = results.get("files", [])
                bucket.extend(found)
                if found:
                    debug_log.append(f"Found {len(found)} files (mime={mime}) in folder {fid}")

        all_files = xlsx_files + gsheet_files
        debug_log.append(f"Total: {len(xlsx_files)} xlsx + {len(gsheet_files)} Sheets across {len(folder_ids)} folders")

        if not all_files:
            debug_log.append("No files found. Check folder structure and sharing permissions.")
            return None, debug_log

        downloaded = {}
        for f in all_files:
            fname = f["name"]
            try:
                if f["mimeType"] == GSHEET_MIME:
                    buf = io.BytesIO()
                    downloader = MediaIoBaseDownload(buf, service.files().export_media(fileId=f["id"], mimeType=XLSX_MIME))
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()
                    if not fname.endswith(".xlsx"):
                        fname += ".xlsx"
                else:
                    buf = io.BytesIO()
                    downloader = MediaIoBaseDownload(buf, service.files().get_media(fileId=f["id"]))
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()
                downloaded[fname] = buf.getvalue()
            except Exception as file_err:
                debug_log.append(f"Skipped '{fname}': {file_err}")

        debug_log.append(f"Successfully downloaded {len(downloaded)} files")
        if not downloaded:
            return None, debug_log
        return downloaded, debug_log

    except KeyError:
        return None, ["No gcp_service_account in secrets"]
    except Exception as e:
        debug_log.append(f"Error: {str(e)}")
        return None, debug_log


# ── Helper: prepare data folder ─────────────────────────────────────
def prepare_data_folder():
    data_path = GENERATED_DIR / "data_working"
    data_path.mkdir(exist_ok=True)

    if DATA_DIR.exists():
        for f in DATA_DIR.glob("*.xlsx"):
            dest = data_path / f.name
            if not dest.exists():
                shutil.copy2(f, dest)
        # Group Financials reads a local-only weekly report (not fetched from
        # Drive). Finance drops a fresh "Group Financial Report_<Mon-YY>.xlsx"
        # each week; always overwrite it (and prune older ones) so the working
        # dir never serves a stale month from a persistent container.
        repo_reports = {f.name for f in DATA_DIR.glob("Group Financial Report*.xlsx")}
        for f in DATA_DIR.glob("Group Financial Report*.xlsx"):
            shutil.copy2(f, data_path / f.name)
        for stale in data_path.glob("Group Financial Report*.xlsx"):
            if stale.name not in repo_reports:
                stale.unlink()
        # Budget vs Actual tab reads a committed JSON snapshot of the Netlify
        # budget tracker (no live API). Refresh it each run so repo updates show.
        snap = DATA_DIR / "budget_tracker_snapshot.json"
        if snap.exists():
            shutil.copy2(snap, data_path / snap.name)

    if GOOGLE_DRIVE_FOLDER_ID:
        drive_files, drive_debug = fetch_drive_folder_files(GOOGLE_DRIVE_FOLDER_ID)
        st.session_state["_gdrive_debug"] = drive_debug
        if drive_files:
            for fname, fbytes in drive_files.items():
                (data_path / fname).write_bytes(fbytes)
            st.session_state["cash_source"] = f"Google Drive ({len(drive_files)} files)"
            st.session_state.pop("_gdrive_errors", None)
        else:
            bundled = len(list(data_path.glob("Cash Report*.xlsx")))
            st.session_state["cash_source"] = f"Local files ({bundled} bundled)" if bundled else "Not available"
            st.session_state["_gdrive_errors"] = drive_debug

    if GOOGLE_SHEET_ID:
        xlsx_bytes = fetch_google_sheet_xlsx(GOOGLE_SHEET_ID)
        if xlsx_bytes:
            (data_path / REVENUE_FILENAME).write_bytes(xlsx_bytes)
            st.session_state["revenue_source"] = "Google Sheet (live)"
        else:
            st.session_state["revenue_source"] = (
                "Local file (bundled)" if (data_path / REVENUE_FILENAME).exists() else "Not available"
            )

    if COLLECTIONS_SHEET_ID:
        coll_bytes = fetch_google_sheet_csv(COLLECTIONS_SHEET_ID, COLLECTIONS_GID)
        if coll_bytes:
            (data_path / COLLECTIONS_FILENAME).write_bytes(coll_bytes)
            st.session_state["collections_source"] = "Google Sheet (live)"
        else:
            st.session_state["collections_source"] = (
                "Local file (bundled)" if (data_path / COLLECTIONS_FILENAME).exists() else "Not available"
            )

    if "uploaded_files" in st.session_state:
        for uploaded in st.session_state.uploaded_files:
            (data_path / uploaded.name).write_bytes(uploaded.getvalue())

    return str(data_path)


# ── Helper: generate a dashboard ────────────────────────────────────
def generate_dashboard(script_name, data_folder, output_name):
    output_path = GENERATED_DIR / output_name
    script_path = APP_DIR / script_name
    if not script_path.exists():
        return None

    import subprocess
    result = subprocess.run(
        [sys.executable, "-B", str(script_path), data_folder],  # -B: bypass .pyc bytecode cache
        capture_output=True, text=True, timeout=120,
    )

    if result.returncode != 0:
        st.error(f"Generator error: {result.stderr[-500:]}")
        return None

    # Surface any data-quality warnings emitted by validate_reports()
    # These indicate parsing failures that could produce wrong numbers.
    stdout = result.stdout or ""
    if "DATA QUALITY" in stdout:
        # Extract the warning block between the ===...=== lines
        lines = stdout.splitlines()
        in_block = False
        warn_lines = []
        for line in lines:
            if line.startswith("=" * 10):
                in_block = not in_block
                continue
            if in_block:
                warn_lines.append(line.strip())
        if warn_lines:
            is_error = any("[ERROR" in l for l in warn_lines)
            msg = "\n".join(warn_lines)
            if is_error:
                st.error(f"⚠️ **Data quality errors detected** — numbers may be incorrect.\n\n{msg}")
            else:
                st.warning(f"⚠️ **Data quality warnings** — verify source files.\n\n{msg}")

    generated_file = Path(data_folder) / output_name
    if generated_file.exists():
        return generated_file.read_text(encoding="utf-8")

    alt_path = Path(data_folder).parent / "outputs" / output_name
    if alt_path.exists():
        return alt_path.read_text(encoding="utf-8")

    return None


# ── App-level HTML cache (shared across all sessions) ────────────────
@st.cache_resource(show_spinner=False)
def _get_app_html_cache():
    return {}


# ── Helper: fix nav for embedded Streamlit view ──────────────────────
def fix_html_for_streamlit(html_content):
    if not html_content:
        return html_content
    nav_css = """<style>
  .top-nav, .nav, .nav-bar { display: none !important; }
  body { padding-top: 0 !important; }
</style>"""
    html_content = html_content.replace("<head>", "<head>" + nav_css, 1)
    return html_content


# ── Main App ─────────────────────────────────────────────────────────
def main():
    check_auth()

    # ── Log dashboard visit once per browser session ───────────────────
    if "visit_logged" not in st.session_state:
        user_email = st.session_state.get("authenticated_email", "unknown")
        log_dashboard_visit(user_email)
        st.session_state.visit_logged = True

    # Auto-refresh every 24 hours
    REFRESH_INTERVAL = 86400
    if "last_refresh" not in st.session_state:
        st.session_state.last_refresh = time.time()

    if time.time() - st.session_state.last_refresh > REFRESH_INTERVAL:
        _get_app_html_cache().clear()
        _get_chat_context_cache().clear()
        st.cache_data.clear()
        st.session_state.last_refresh = time.time()
        st.rerun()

    # ── Inject floating 💬 button into the parent Streamlit page ──────
    # Uses components.html() so window.parent correctly targets the Streamlit
    # page DOM — allowing us to inject the button element directly and reliably
    # click the sidebar toggle.  st.markdown() can't reach window.parent.
    components.html("""
<script>
(function injectBobby() {
    var doc = window.parent.document;
    if (doc.getElementById('bobby-fab-wrap')) return;  // already injected

    // Styles (injected into parent page head)
    var s = doc.createElement('style');
    s.id = 'bobby-fab-style';
    s.textContent =
        '#bobby-fab-wrap{position:fixed;bottom:100px;right:28px;z-index:99999;display:flex;flex-direction:column;align-items:center;gap:6px;}' +
        '#bobby-fab-label{background:#0f172a;color:#00D4AA;font-size:11px;font-family:-apple-system,sans-serif;padding:3px 10px;border-radius:10px;border:1px solid rgba(0,212,170,0.35);white-space:nowrap;pointer-events:none;}' +
        '#bobby-fab{width:54px;height:54px;background:linear-gradient(135deg,#00D4AA 0%,#0066CC 100%);border-radius:50%;border:none;cursor:pointer;font-size:24px;color:#fff;box-shadow:0 4px 18px rgba(0,212,170,0.45);transition:transform 0.18s,box-shadow 0.18s;}' +
        '#bobby-fab:hover{transform:scale(1.1);box-shadow:0 6px 26px rgba(0,212,170,0.6);}';
    doc.head.appendChild(s);

    // Build elements
    var wrap = doc.createElement('div'); wrap.id = 'bobby-fab-wrap';
    var lbl  = doc.createElement('div'); lbl.id  = 'bobby-fab-label'; lbl.textContent = 'Ask Bobby';
    var btn  = doc.createElement('button'); btn.id = 'bobby-fab'; btn.title = 'Ask Bobby'; btn.textContent = '💬';

    btn.addEventListener('click', function() {
        // Try every known selector for the sidebar open button across Streamlit versions
        var toggle =
            doc.querySelector('[data-testid="stSidebarCollapsedControl"] button') ||
            doc.querySelector('[data-testid="stSidebarCollapsedControl"]')        ||
            doc.querySelector('[data-testid="collapsedControl"]')                 ||
            doc.querySelector('[data-testid="stSidebarToggle"]')                  ||
            doc.querySelector('button[aria-label*="sidebar"]')                    ||
            doc.querySelector('button[aria-label*="Sidebar"]')                    ||
            doc.querySelector('section[data-testid="stSidebar"] ~ div button');
        if (toggle) {
            toggle.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, view: window.parent}));
        }
    });

    wrap.appendChild(lbl);
    wrap.appendChild(btn);
    doc.body.appendChild(wrap);
})();

// ── Theme sync: listen for postMessage from dashboard iframes ──
(function initThemeSync() {
    var doc = window.parent.document;
    if (doc.getElementById('seamfix-theme-style')) return;

    // Inject a style tag we can update dynamically
    var style = doc.createElement('style');
    style.id = 'seamfix-theme-style';
    doc.head.appendChild(style);

    function applyStreamlitTheme(theme) {
        if (theme === 'dark') {
            style.textContent =
                '[data-testid="stAppViewContainer"], [data-testid="stApp"], .main, .block-container { background-color: #0f172a !important; color: #e2e8f0 !important; }' +
                '[data-testid="stSidebar"], [data-testid="stSidebar"] > div { background-color: #1e293b !important; color: #e2e8f0 !important; }' +
                '[data-testid="stHeader"], header[data-testid="stHeader"] { background-color: #0a0f1e !important; color: #e2e8f0 !important; }' +
                '[data-testid="stSidebar"] label, [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #e2e8f0 !important; }' +
                '[data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] small { color: #94a3b8 !important; }' +
                '[data-testid="stSidebar"] button { color: #e2e8f0 !important; border-color: #334155 !important; }' +
                '[data-testid="stSidebar"] input, [data-testid="stSidebar"] textarea { background-color: #0f172a !important; color: #e2e8f0 !important; border-color: #334155 !important; }' +
                '.stTabs [data-baseweb="tab-list"] { background-color: #0f172a !important; }' +
                '.stTabs [data-baseweb="tab"] { color: #94a3b8 !important; }' +
                '.stTabs [aria-selected="true"] { color: #00D4AA !important; }' +
                '#bobby-fab-label { background: #0f172a !important; color: #00D4AA !important; border-color: rgba(0,212,170,0.35) !important; }';
        } else {
            style.textContent =
                '[data-testid="stAppViewContainer"], [data-testid="stApp"], .main, .block-container { background-color: #f8fafc !important; color: #1e293b !important; }' +
                '[data-testid="stSidebar"], [data-testid="stSidebar"] > div { background-color: #ffffff !important; color: #1e293b !important; }' +
                '[data-testid="stHeader"], header[data-testid="stHeader"] { background-color: #ffffff !important; color: #1e293b !important; }' +
                '[data-testid="stSidebar"] label, [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #1e293b !important; }' +
                '[data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] small { color: #64748b !important; }' +
                '[data-testid="stSidebar"] button { color: #1e293b !important; border-color: #e2e8f0 !important; }' +
                '[data-testid="stSidebar"] input, [data-testid="stSidebar"] textarea { background-color: #ffffff !important; color: #1e293b !important; border-color: #e2e8f0 !important; }' +
                '.stTabs [data-baseweb="tab-list"] { background-color: #ffffff !important; }' +
                '.stTabs [data-baseweb="tab"] { color: #64748b !important; }' +
                '.stTabs [aria-selected="true"] { color: #009E7E !important; }' +
                '#bobby-fab-label { background: #ffffff !important; color: #009E7E !important; border-color: rgba(0,158,126,0.35) !important; }';
        }
        try { localStorage.setItem('seamfix-theme', theme); } catch(e) {}
    }

    // Listen for theme changes from dashboard iframes
    window.parent.addEventListener('message', function(e) {
        if (e.data && e.data.seamfixTheme) {
            applyStreamlitTheme(e.data.seamfixTheme);
        }
    });

    // Apply saved theme on load
    var saved = 'light';
    try { saved = localStorage.getItem('seamfix-theme') || 'light'; } catch(e) {}
    applyStreamlitTheme(saved);
})();
</script>
""", height=0)

    # ── Dashboard tabs ────────────────────────────────────────────
    dash_names  = list(DASHBOARDS.keys())
    dash_labels = [f"{DASHBOARDS[d]['icon']}  {d}" for d in dash_names]
    tab_objects = st.tabs(dash_labels)

    # ── Prepare data ──────────────────────────────────────────────
    app_cache   = _get_app_html_cache()
    data_folder = prepare_data_folder()

    # ── Generate all dashboards in parallel (once, shared across all users) ──
    if any(name not in app_cache for name in DASHBOARDS):
        import concurrent.futures

        def _generate_one(dash_name):
            dash = DASHBOARDS[dash_name]
            html = generate_dashboard(dash["script"], data_folder, dash["output"])
            if html:
                html = fix_html_for_streamlit(html)
            return dash_name, html

        # max_workers is intentionally low (2): Streamlit Cloud runs on a heavily
        # throttled/shared CPU, and the cash/expense/budget generators each re-parse
        # every accumulated weekly cash report via openpyxl (CPU + memory heavy).
        # Running all 6 at once starved them so badly they all hit the 60s subprocess
        # timeout. Capping concurrency lets each finish well within its timeout.
        with st.spinner("Loading dashboards — this takes about 20 seconds on first visit..."):
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                for dash_name, html in executor.map(_generate_one, list(DASHBOARDS.keys())):
                    app_cache[dash_name] = html

    # ── Render each tab ───────────────────────────────────────────
    resize_script = """
<script>
function sendHeight() {
    var h = document.body.scrollHeight + 100;
    window.parent.postMessage({type: 'streamlit:setFrameHeight', height: h}, '*');
}
window.addEventListener('load', sendHeight);
setTimeout(sendHeight, 300);
setTimeout(sendHeight, 800);
setTimeout(sendHeight, 1500);
setTimeout(sendHeight, 3000);
setTimeout(sendHeight, 5000);
</script>
"""
    for i, dash_name in enumerate(dash_names):
        with tab_objects[i]:
            html_content = app_cache.get(dash_name)
            if html_content:
                display_html = html_content.replace("</body>", resize_script + "</body>")
                components.html(display_html, height=15000, scrolling=False)
            else:
                st.warning(
                    f"Could not generate the {dash_name} dashboard. "
                    "Check that all required data files are available."
                )

    # ── Sidebar: Chat interface + Data management ─────────────────
    with st.sidebar:

        # ── Bobby chat (fragment — only sidebar reruns on each send/clear) ──
        _bobby_chat_fragment(data_folder)

        st.divider()

        # ── Data management (collapsed) ───────────────────────────
        with st.expander("⚙️  Data & Settings", expanded=False):
            st.markdown("##### Data Management")
            st.caption(
                "⚠️ Only upload files if authorized by the finance team. "
                "Incorrect files may corrupt the dashboard data."
            )

            uploaded = st.file_uploader(
                "Upload new cash reports",
                type=["xlsx"],
                accept_multiple_files=True,
                key="uploaded_files",
                help="Drop new weekly cash report xlsx files here",
            )

            if st.button("🔄 Regenerate Dashboards", use_container_width=True):
                _get_app_html_cache().clear()
                _get_chat_context_cache().clear()  # also clears Bobby's context so picks up new data
                st.cache_data.clear()
                for key in list(st.session_state.keys()):
                    if key.startswith("_g"):
                        del st.session_state[key]
                st.session_state.last_refresh = time.time()
                st.rerun()

            last_ref = st.session_state.get("last_refresh")
            if last_ref:
                st.caption(f"Last refreshed: {datetime.fromtimestamp(last_ref).strftime('%H:%M %d %b')}")
            st.caption("Auto-refreshes daily")

            st.markdown("##### Data Sources")
            rev_source   = st.session_state.get("revenue_source",    "Checking...")
            cash_source  = st.session_state.get("cash_source",       "Checking...")
            coll_source  = st.session_state.get("collections_source","Checking...")
            upload_count = len(uploaded) if uploaded else 0

            st.caption(f"📈 Revenue: {rev_source}")
            st.caption(f"💵 Cash reports: {cash_source}" + (f" + {upload_count} uploaded" if upload_count else ""))
            st.caption(f"📋 Budget: {'Available' if (DATA_DIR / BUDGET_FILENAME).exists() else 'Missing'}")
            st.caption(f"📥 Collections: {coll_source}")

            gsheet_errs = st.session_state.get("_gsheet_errors", [])
            gdrive_errs = st.session_state.get("_gdrive_errors", [])
            if gsheet_errs or gdrive_errs:
                with st.expander("Connection issues"):
                    if gsheet_errs:
                        st.caption("Google Sheet:")
                        for e in gsheet_errs:
                            st.caption(f"  {e}")
                    if gdrive_errs:
                        st.caption("Google Drive:")
                        for e in gdrive_errs:
                            st.caption(f"  {e}")

        # Logged-in user
        auth_email = st.session_state.get("authenticated_email")
        if auth_email:
            st.divider()
            st.caption(f"Signed in as **{auth_email}**")
            if st.button("Sign out"):
                del st.session_state["authenticated_email"]
                st.rerun()


@st.fragment
def _bobby_chat_fragment(data_folder):
    """Bobby chat UI — runs as a fragment so only the sidebar reruns, not the full page."""
    st.markdown(
        """
        <div style="padding:12px 0 4px">
            <div style="font-size:18px;font-weight:700;color:#00D4AA">💬 Ask Bobby</div>
            <div style="font-size:12px;color:#94a3b8;margin-top:2px">
                AI analyst with access to your live financial data
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Initialize chat history
    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = []

    # Display chat history
    if st.session_state.chat_messages:
        for msg in st.session_state.chat_messages:
            with st.chat_message(msg["role"], avatar="🧑" if msg["role"] == "user" else "🤖"):
                st.markdown(msg["content"])
    else:
        st.caption(
            "👋 Ask me anything about Seamfix's finances — "
            "cash position, budget variances, pipeline deals, expense categories, and more."
        )

    # Chat input
    user_input = st.text_area(
        "Your question",
        key="chat_input_field",
        placeholder="e.g. How much runway do we have? Which pipeline deals are At Risk?",
        label_visibility="collapsed",
        height=96,
    )

    col_send, col_clear = st.columns([2, 1])
    with col_send:
        send_clicked = st.button("Ask Bobby →", use_container_width=True, type="primary")
    with col_clear:
        clear_clicked = st.button("Clear", use_container_width=True)

    if clear_clicked:
        st.session_state.chat_messages = []
        st.rerun(scope="fragment")

    if send_clicked and user_input.strip():
        prompt = user_input.strip()
        st.session_state.chat_messages.append({"role": "user", "content": prompt})

        with st.spinner("Bobby is thinking..."):
            context = get_chat_context(data_folder)
            response, in_tok, cache_read_tok, cache_write_tok, out_tok = \
                call_claude(st.session_state.chat_messages, context)

        st.session_state.chat_messages.append({"role": "assistant", "content": response})

        # Log to Google Sheet (silent on failure)
        user_email = st.session_state.get("authenticated_email", "unknown")
        log_bobby_query(user_email, prompt, response,
                        in_tok, cache_read_tok, cache_write_tok, out_tok)

        st.rerun(scope="fragment")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        st.error(f"**App Error:** {type(e).__name__}: {e}")
        st.code(traceback.format_exc(), language="python")
        st.info("Please share the traceback above with the developer.")
