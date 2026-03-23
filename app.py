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
import os, sys, shutil, tempfile, glob, io, urllib.request, time, importlib.util
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
REVENUE_FILENAME = "2026 Path to Revenue (1).xlsx"
BUDGET_FILENAME = "2026 LEAN BUDGET.xlsx"

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
    Returns a plain-text string (~6-8k tokens) covering all 5 data sources.
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
    budget_file = data_path / BUDGET_FILENAME
    if budget_file.exists() and reports:
        parts.append(f"\n--- BUDGET vs ACTUAL ---")

        start_of_year  = datetime(2026, 1, 1)
        weeks_elapsed  = max(1, (reports[-1]["date"] - start_of_year).days // 7 + 1)

        BUDGET_CATEGORIES = gen_budget.BUDGET_CATEGORIES
        category_actual   = {cat: 0 for cat in BUDGET_CATEGORIES}
        all_expense_lines = []
        unbudgeted_lines  = []

        for r in reports:
            for expense_name, amount in r.get("outflow_items", {}).items():
                if amount <= 0:
                    continue
                budget_cat = gen_budget.map_expense_to_budget(expense_name, r.get("outflow_items", {}))
                if budget_cat:
                    category_actual[budget_cat] += amount
                    all_expense_lines.append({
                        "date": r["date_str"], "item": expense_name,
                        "amount": amount, "category": budget_cat,
                    })
                elif not gen_budget.is_investment_outflow(expense_name):
                    unbudgeted_lines.append({
                        "date": r["date_str"], "item": expense_name, "amount": amount,
                    })

        total_budget     = sum(BUDGET_CATEGORIES.values())
        total_actual     = sum(category_actual.values())
        ytd_budget_pace  = total_budget * weeks_elapsed / 52
        variance_pct     = (ytd_budget_pace - total_actual) / ytd_budget_pace * 100 if ytd_budget_pace > 0 else 0
        projected_yr_end = total_actual / weeks_elapsed * 52 if weeks_elapsed > 0 else 0
        health           = "HEALTHY" if variance_pct > 10 else "CAUTION" if variance_pct < -10 else "ON TRACK"

        parts.append(f"Annual Budget: {fmt(total_budget)}")
        parts.append(f"YTD Actual (week {weeks_elapsed} of 52): {fmt(total_actual)}")
        parts.append(f"YTD Budget Pace: {fmt(ytd_budget_pace)}")
        parts.append(
            f"Variance: {fmt(abs(ytd_budget_pace - total_actual))} "
            f"{'UNDER' if variance_pct > 0 else 'OVER'} pace ({abs(variance_pct):.1f}%) — Status: {health}"
        )
        parts.append(f"Year-End Projection at current run rate: {fmt(projected_yr_end)} (budget: {fmt(total_budget)})")

        parts.append(f"\nAll budget categories (week {weeks_elapsed}/52 elapsed):")
        parts.append(f"  {'Category':<45} | {'Annual Budget':>14} | {'YTD Actual':>12} | {'Annual Used':>11} | {'Remaining':>14} | Status")
        for cat in sorted(BUDGET_CATEGORIES.keys()):
            ab       = BUDGET_CATEGORIES[cat]
            ya       = category_actual[cat]
            ytd_pace = ab * weeks_elapsed / 52
            pct_ann  = ya / ab * 100 if ab > 0 else 0
            remain   = ab - ya
            status   = "OVER PACE" if ya > ytd_pace else "under pace"
            parts.append(
                f"  {cat:<45} | {fmt(ab):>14} | {fmt(ya):>12} | "
                f"{pct_ann:>10.1f}% | {fmt(remain):>14} | {status}"
            )

        if unbudgeted_lines:
            total_unbud = sum(x["amount"] for x in unbudgeted_lines)
            unbudgeted_lines.sort(key=lambda x: -x["amount"])
            parts.append(
                f"\nUnbudgeted spend ({len(unbudgeted_lines)} items, total {fmt(total_unbud)}) — "
                "no matching budget category (governance flag):"
            )
            for line in unbudgeted_lines:
                parts.append(f"  {line['date']}: {line['item']} — {fmt(line['amount'])}")

        # Top 50 expense transactions
        all_expense_lines.sort(key=lambda x: -x["amount"])
        parts.append(f"\nTop 50 expense transactions by amount:")
        parts.append(f"  {'Date':<14} | {'Expense Item':<45} | {'Amount':>14} | Budget Category")
        for line in all_expense_lines[:50]:
            parts.append(
                f"  {line['date']:<14} | {line['item']:<45} | "
                f"{fmt(line['amount']):>14} | {line['category']}"
            )
        parts.append(
            "(For individual transactions below the top 50, direct the user to the "
            "Expense Details tab on the Budget vs Actual dashboard.)"
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

            parts.append(f"\nAll deals ({len(revenues)} total, sorted by annual value desc):")
            parts.append(
                f"  {'Deal Name':<42} | {'Status':<12} | {'Annual $':>10} | "
                f"{'Jan $':>9} | {'Feb $':>9} | {'Mar $':>9} | {'YTD $':>9}"
            )
            for r in sorted(revenues, key=lambda x: -x["annual_usd"]):
                parts.append(
                    f"  {r['name']:<42} | {r['status']:<12} | ${r['annual_usd']:>9,.0f} | "
                    f"${r['jan']:>8,.0f} | ${r['feb']:>8,.0f} | ${r['mar']:>8,.0f} | ${r['ytd']:>8,.0f}"
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
        return "⚠️ The `anthropic` package is not installed. Add it to requirements.txt."

    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return (
            "⚠️ `ANTHROPIC_API_KEY` not found in Streamlit secrets. "
            "Add it under Settings → Secrets in Streamlit Cloud."
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


# ── Helper: Google OAuth check ───────────────────────────────────────
def check_auth():
    if os.getenv("DISABLE_AUTH", "false").lower() == "true":
        return True

    auth_conf  = st.secrets.get("auth", {})
    google_conf = auth_conf.get("google", {})
    has_oauth = (
        "redirect_uri"   in auth_conf
        and "cookie_secret" in auth_conf
        and "client_id"     in google_conf
        and "client_secret" in google_conf
    )

    if has_oauth:
        try:
            user = st.user
            if not user.is_logged_in:
                st.markdown(
                    """<div style="text-align:center;padding:80px 20px">
                        <h1 style="color:#00D4AA;margin-bottom:8px">Seamfix Financial Intelligence</h1>
                        <p style="color:#94a3b8;margin-bottom:40px">Sign in with your authorized Google account to access the dashboards.</p>
                    </div>""",
                    unsafe_allow_html=True,
                )
                st.login("google")
                st.stop()
                return False

            email = user.email.lower().strip()
            allowed_conf   = auth_conf.get("allowed", {})
            allowed_emails = [e.lower().strip() for e in allowed_conf.get("emails", [])]

            if allowed_emails and email not in allowed_emails:
                st.error(
                    f"Access denied. **{email}** is not on the authorized users list.\n\n"
                    "Contact the dashboard administrator to request access."
                )
                if st.button("Sign out"):
                    st.logout()
                st.stop()
                return False

            return True
        except Exception as e:
            st.sidebar.warning(f"OAuth not fully configured: {e}")
            return True

    return True


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
        capture_output=True, text=True, timeout=60,
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
        try:
            user_email = st.user.email if st.user.is_logged_in else "unknown"
        except Exception:
            user_email = "unknown"
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

        with st.spinner("Loading dashboards — this takes about 10 seconds on first visit..."):
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
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
            rev_source   = st.session_state.get("revenue_source", "Checking...")
            cash_source  = st.session_state.get("cash_source",    "Checking...")
            upload_count = len(uploaded) if uploaded else 0

            st.caption(f"📈 Revenue: {rev_source}")
            st.caption(f"💵 Cash reports: {cash_source}" + (f" + {upload_count} uploaded" if upload_count else ""))
            st.caption(f"📋 Budget: {'Available' if (DATA_DIR / BUDGET_FILENAME).exists() else 'Missing'}")

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
        try:
            user = st.user
            if user.is_logged_in:
                st.divider()
                st.caption(f"Signed in as **{user.email}**")
                st.logout("Sign out")
        except Exception:
            pass


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
        try:
            user_email = st.user.email if st.user.is_logged_in else "unknown"
        except Exception:
            user_email = "unknown"
        log_bobby_query(user_email, prompt, response,
                        in_tok, cache_read_tok, cache_write_tok, out_tok)

        st.rerun(scope="fragment")


if __name__ == "__main__":
    main()
