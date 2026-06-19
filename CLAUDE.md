# Seamfix Financial Intelligence Suite — Project Documentation

> **Purpose**: This document gives a new developer or AI assistant full context to maintain and extend this project. Read this first.

---

## MANDATORY: Rules for Any AI Assistant Working on This Project

**These rules apply to every Claude instance, every session, every change — no exceptions.**

### 0. Pull Latest Before Starting ANY Work

Before making any changes, ALWAYS run:
```bash
git pull origin main
```
Multiple people (and their Claude instances) may be working on this project. If you skip this step, you risk overwriting someone else's work or creating merge conflicts. **This is the very first thing you do in every session, before reading files or making edits.**

### 1. Document Every Change

After making ANY code change, you MUST:

- **Update this file (CLAUDE.md)** if the change affects architecture, business logic, column mappings, authentication, deployment, or any "how it works" section. Update the relevant section — don't just append.
- **Add an entry to CHANGELOG.md** with the date, what changed, why, and which files were modified. Follow the format already in the file.
- **If you add a new gotcha or lesson learned**, add it to the "Lessons Learned & Gotchas" section below.

### 2. Commit Message Standards

Every commit message must:
- Start with what type of change it is: `Fix:`, `Feature:`, `Update:`, `Refactor:`, or `Docs:`
- Explain WHY, not just what (e.g., "Fix: achievement % using annual target instead of pace — Finance expects 100% for completed deals" not "Fix percentage calculation")

### 3. Test Before Pushing

Before pushing any change:
- Run the affected generator script locally: `python3 generate_*.py ./data`
- Verify the HTML output opens correctly in a browser
- Check that no other dashboard broke (if you changed shared logic like FX_RATE)

### 4. Don't Break the Column Convention

The Google Sheet column layout (N=Jan through Y=Dec, Z=Deficit, AA=Surplus — these shift whenever Finance inserts a column) is the single most fragile part of this system. If Finance changes the sheet structure:
- Update the column mappings in ALL affected generator files
- Update the "Excel Column Mapping" section in this document
- Add a CHANGELOG entry explaining the old vs new mapping

### 5. Keep This File Current

If you notice any section of this document is outdated or wrong, fix it immediately. This file is the single source of truth for anyone (human or AI) working on this project.

---

## Overview

This is a Streamlit web application that serves 5 interactive financial dashboards plus an AI chat assistant ("Bobby") for Seamfix's executive team. It pulls live data from Google Sheets/Drive, processes Excel files, generates standalone HTML dashboards, and embeds them in a tabbed Streamlit interface.

**Live URL**: https://seamfix-executive-dashboard.streamlit.app  
**Repo**: https://github.com/sfxconwurah/seamfix-dashboard  
**Hosting**: Streamlit Community Cloud  
**Python**: 3.12 (set via Streamlit Cloud settings dropdown — `.python-version` and `runtime.txt` are ignored by Streamlit Cloud)

---

## Repository & Data Policy

**This repository is the single source of truth.** There is exactly one copy of every generator and `app.py` — the ones in this folder. (Historically the project was duplicated inside an outer `financial-dashboards` repo and synced by hand; that caused drift and has been retired. If you find another copy anywhere, this one wins.)

- **Canonical remote (for now)**: `sfxconwurah/seamfix-dashboard` (personal). A future migration to the `seamfix` org repo is planned but blocked on CTO access — do not assume the org repo is current yet.
- **Financial data is local-only going forward**: `*.xlsx` is gitignored. The live app reads fresh data from Google Drive/Sheets, so committed data is not required for production. The existing `data/*.xlsx` files remain tracked as an offline fallback; to refresh that baseline you must force-add (`git add -f <file>.xlsx`). Otherwise, just drop xlsx into `data/` locally and they stay on your machine.
- **Secrets** (`.streamlit/secrets.toml`, any `*.rtf` secrets export) are gitignored and must never be committed.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  Streamlit Cloud (app.py)                                   │
│  ┌───────────────┐  ┌───────────────────────────────────┐  │
│  │ Custom OAuth  │  │ Dashboard Tabs                     │  │
│  │ (Google)      │  │  • Cash Overview                   │  │
│  └───────────────┘  │  • Expense & Vendor Analysis       │  │
│                      │  • Budget vs Actual                │  │
│  ┌───────────────┐  │  • Revenue & Fundability           │  │
│  │ Bobby Chat    │  │  • Pipeline Intelligence           │  │
│  │ (Claude API)  │  └───────────────────────────────────┘  │
│  └───────────────┘                                          │
└─────────────────────────────────────────────────────────────┘
         │                          │
         ▼                          ▼
┌─────────────────┐    ┌──────────────────────────────┐
│ Google Sheet    │    │ Generator Scripts             │
│ (Revenue data)  │    │  generate_dashboard.py        │
│                 │    │  generate_expense_dashboard.py │
│ Google Drive    │    │  generate_budget_dashboard.py  │
│ (Cash reports)  │    │  generate_revenue_dashboard.py │
│                 │    │  generate_pipeline_dashboard.py│
└─────────────────┘    └──────────────────────────────┘
                                    │
                                    ▼
                        ┌──────────────────────┐
                        │ Standalone HTML files │
                        │ (Chart.js + CSS)      │
                        │ Embedded via           │
                        │ components.html()      │
                        └──────────────────────┘
```

### Data Flow

1. **On page load**, `app.py` calls `prepare_data_folder()` which:
   - Copies bundled `data/*.xlsx` files to a working directory
   - Fetches the live Google Sheet (revenue/pipeline) via export URL or service account
   - Fetches all cash report xlsx files from a Google Drive folder
   - Merges any user-uploaded files

2. **Generator scripts** are invoked as subprocesses. Each reads xlsx files from the data folder and outputs a self-contained HTML dashboard file.

3. **HTML dashboards** are embedded in Streamlit tabs using `components.html()` with a height of 15000px.

4. **Bobby** (the chat assistant) builds a text context from ALL financial data, sends it to the Claude API with prompt caching, and displays responses in the sidebar.

---

## File Structure

```
seamfix-dashboard/
├── app.py                          # Main Streamlit app (auth, data prep, tabs, chat)
├── generate_dashboard.py           # Cash Overview dashboard generator
├── generate_expense_dashboard.py   # Expense & Vendor Analysis generator
├── generate_budget_dashboard.py    # Budget vs Actual generator
├── generate_revenue_dashboard.py   # Revenue & Fundability generator
├── generate_pipeline_dashboard.py  # Pipeline Intelligence generator
├── generate_collections_dashboard.py  # Collections Tracker generator
├── generate_financial_report_dashboard.py  # Group Financials (consolidated P&L) generator
├── requirements.txt                # Python dependencies
├── runtime.txt                     # Python version hint (ignored by Streamlit Cloud)
├── .python-version                 # Python version hint (ignored by Streamlit Cloud)
├── .streamlit/
│   ├── config.toml                 # Theme + server config
│   └── secrets.toml.example        # Template for secrets
├── data/                           # Bundled xlsx files (fallback if Drive/Sheets unavailable)
│   ├── Cash Report as at *.xlsx    # Weekly cash position reports
│   ├── 2026 Path to Revenue (1).xlsx   # Revenue & pipeline tracker
│   ├── 2026 LEAN BUDGET.xlsx       # Annual budget breakdown (legacy — no longer read by Budget tab)
│   ├── budget_tracker_snapshot.json   # Budget vs Actual data (committed snapshot of Netlify tracker)
│   └── Group Financial Report_*.xlsx  # Consolidated P&L (LOCAL-ONLY — not fetched online)
└── generated/                      # Runtime working directory (gitignored)
    └── data_working/               # Merged data + generated HTML (ephemeral)
```

---

## Excel Column Mapping (Critical Reference)

### Revenue Sheet (`2026 Path to Revenue (1).xlsx` → "Revenues" tab)

| Column | Content |
|--------|---------|
| A | S/N (serial number — parent deals have this) |
| B | Deal name |
| C | Rail / category |
| D | Recurring / Not Recurring (added June 2026) |
| E | Revenue Start Date |
| F | 2026 Annual Revenue (USD) |
| G–K | Monthly/Weekly/Daily Revenue, Unit Price, No. of Daily Applications (not read by generators) |
| L | Status (On Track / At Risk / Off Track / Closed) |
| M | Comments (free text from finance) |
| N | January actual (USD) |
| O | February actual (USD) |
| P | March actual (USD) |
| Q–Y | April–December actual (USD) |
| Z | Deficit (USD) |
| AA | Surplus (USD) |
| AB | Gap (USD) |

**IMPORTANT — column history**: In March 2026 Finance added monthly columns Apr–Dec. In **June 2026 Finance inserted a new column D ("Recurring/Not Recurring")**, which shifted *every* column after C one to the right: Annual Revenue E→**F**, Start Date D→**E**, Status K→**L**, Comments L→**M**, monthly actuals M–X→**N–Y**, Deficit Y→**Z**, Surplus Z→**AA**. If Finance shifts columns again, update `MONTH_COLUMNS` plus `cells.get('F')` (annual), `cells.get('E')` (start), `cells.get('L')` (status), `cells.get('D')` (recurring), `cells.get('Z')`/`cells.get('AA')` (deficit/surplus) in `generate_revenue_dashboard.py`, and the matching `d.get('F'/'L'/'M')` refs in `generate_pipeline_dashboard.py`.

**ARR cards**: The Revenue & Fundability dashboard shows two ARR KPIs (in `generate_revenue_dashboard.py`): **ARR (Annual Recurring Revenue)** = sum of column F for deals where column D == "Recurring" (`arr_usd`); and **ARR As At <month>** = `recurring_ytd_usd / ytd_actual_revenue_usd` (`arr_asat_pct`) — the recurring share of revenue *actually earned* to date, measured against a **50% target** (`ARR_TARGET_PCT`, flagged red when below). The "ARR As At" card replaced the old "YTD Achievement" KPI (which had become a duplicate of "Annual Progress" once total pipeline equalled the $8M target). A plan-based "ARR Mix (% of Revenue)" card was briefly added then removed at Finance's request — keep only the actuals-based "ARR As At" %.

### Cash Report Files (`Cash Report as at *.xlsx`)

Each weekly cash report has a fixed structure parsed by `generate_dashboard.py`:
- NGN closing balances
- USD closing balances
- Inflow items (categorized)
- Outflow items (categorized)
- Investment portfolio positions

The date is extracted from the filename using regex: `(\d+)\w*\s+(Month)\s+(\d{4})`.

### Budget vs Actual — Budget Tracker snapshot (`data/budget_tracker_snapshot.json`)

**As of 2026-06-11 the Budget vs Actual tab is driven by a committed JSON snapshot of the external Seamfix Budget Tracker (https://seamfix-budget-tracker.netlify.app/), NOT by `2026 LEAN BUDGET.xlsx` + cash-outflow fuzzy matching (that old approach was retired).**

- **Why a snapshot, not a live fetch:** the Netlify tracker is a fully client-side page — its data lives in hardcoded JS objects (`BUDGET_DATA`, `C`, `ACT`, `EXC`), refreshed manually per "Run". There is no API. So we extract the data into `data/budget_tracker_snapshot.json` and commit it. It is a plain `.json` (not gitignored) so it ships to Streamlit Cloud. `app.py`'s `prepare_data_folder()` copies it into the working dir each run.
- **To refresh:** re-extract the tracker's `BUDGET_DATA` (lean mode) + `C` (meta/FX) objects into the snapshot (the current snapshot was built from `Run_004`, 2026-05-30, actuals through Apr-2026). Bump `runId`/`runDate`/`lastActualsMonth`/`elapsedMonths` accordingly.
- **Mode = lean.** The tracker has two modes: **lean** (Acumatica-loaded budget that carries actuals) and **full** (approved budget, **no actuals**). Only lean can drive a budget-vs-actual view, so the snapshot stores lean only.
- **Currency = NGN, bottom-up.** Built so **Group = Σ entities = Σ departments**. NG dept budgets are already NGN; **UK budgets are GBP and UAE budgets are USD**, FX-converted to NGN using the tracker's lean FX (`GBP_NGN=2000`, `USD_NGN=1500`). **Actuals are already NGN for every entity** (do NOT FX-convert actuals). FY rollup reconciles exactly to the tracker's `groupFY_NGN` (₦5.10B). YTD actual reconciles bottom-up to ₦1.45B (the tracker's own headline `groupYTD_A_NGN` ₦1.47B is ~1.4% higher because it includes mapped transactions not allocated to a department head — we deliberately use the reconciling bottom-up figure).
- **Three levels rendered:** GROUP KPIs + health; COMPANY-wide entity cards/table (NG/UK/UAE); DEPARTMENT-wide sortable table (11 depts: BGI=Commercial, OPS=Admin & Operations, SOL=Solutions, HRM=PPC, PDM=Products, FIN=Finance, LEG=Legal, MGT=Directors & Mgt, CAP=CAPEX, UK, UAE) with click-to-expand budget-head drill-down. Plus dept bar chart + group monthly budget-vs-actual line chart, and executive takeaways (over-pace / underspending depts, largest spend lines).
- **Snapshot schema:** top-level meta (`fiscalYear`, `runId`, `runDate`, `lastActualsMonth`, `elapsedMonths`, `months[]`, `leanFX`, revenue context) + `departments[]` (each: `dept_code`, `dept_name`, `entity`, `currency`, `annual_total`, `months{Jan..Dec}` budget, `budget_heads[]` with `name`/`annual`/`months`/`actuals{jan..dec lowercase, NGN}`).
- Generator **fails safe**: missing/unreadable snapshot or empty departments → placeholder HTML + exit 0.
- `2026 LEAN BUDGET.xlsx` (the old source) is no longer read by this tab. `generate_budget_dashboard.py`'s `main()` now reads `budget_tracker_snapshot.json` and outputs `budget_dashboard.html`.

### Group Financial Report (`Group Financial Report_*.xlsx` → "Summary" tab)

Consolidated P&L / balance sheet produced by Finance (drives the **Group Financials** tab via `generate_financial_report_dashboard.py`). The `Summary` tab holds two side-by-side blocks: **NIGERIA** (cols C–H) and **GROUP** (cols K–P). We read the **GROUP** block:

| Column | Content |
|--------|---------|
| K | Line-item label (e.g. "Total Revenue", "Gross Profit", "EBITDA", "Profit After Tax", "Gross Profit Margin", "ARR") |
| L | Current-period YTD value (NGN) |
| M | Prior-period YTD value (NGN) |
| N | Current-period YTD value (USD — report's own period-average FX) |
| O | Prior-period YTD value (USD) |
| P | Variance % |

**Parsing is label-driven** — the generator scans column K for line-item names rather than hardcoding rows, so it survives row insertions. It reads the income-statement region (everything above the `GROUP BALANCE SHEET` header), the ratios block (Gross/EBITDA/Net margins, OpEx/Payroll/Marketing % of revenue, ARR %), the three revenue breakdowns ("Revenue by Vertical/Customer/Country", each terminated by its own `Total` row), and a few balance-sheet highlights. USD figures come from the report's own N/O columns, **not** `FX_RATE=1450`.

**Executive targets (hardcoded in the generator):** Net Profit Margin `NET_MARGIN_TARGET = 10` %, Gross Profit Margin `GROSS_MARGIN_TARGET = 70` %, `WACC_PCT = 37` % (Finance's board-approved cost-of-capital / hurdle rate, used for EVA — see below), and `TAX_RATE = 30` % (Nigeria statutory CIT, used to normalise NOPAT when no tax is booked).

**Layout (top → bottom):** 5 KPI cards (Total Revenue, Gross Profit, EBITDA, PAT, **ARR**); two profitability gauge cards (net & gross margin vs target); **Key Financial Ratios** grid (13 cards: gross/EBITDA/operating/net margin, annualised ROA, OpEx-/payroll-/marketing-to-revenue, effective tax rate, interest coverage, current & cash ratio, ARR % of revenue); **Economic Value Added (EVA)** card; Critical Insights; Income Statement; Expense Analysis + Revenue-by-Vertical chart; revenue breakdowns; Top Customers; Balance Sheet & Liquidity.

**EVA:** `EVA = NOPAT − (Invested Capital × WACC)`, where NOPAT = EBIT × (1 − effective tax rate), **invested capital = Total Current Assets − Total Current Liabilities + Total Non-Current (net fixed) Assets** (≡ Total Assets − Current Liabilities; ₦6.34B for May-26, parsed from the Summary balance-sheet rows), and the capital charge is **prorated to the YTD period** (`year_frac = months_elapsed/12`) so it's comparable with period NOPAT. NOPAT's tax: the report books ₦0 tax YTD, so a derived rate would be a misleading 0% — the generator uses the **30% Nigeria statutory CIT** (`TAX_RATE`) as a fallback, and only switches to the report's booked effective rate once a real tax charge appears (`eff_tax = eff_tax_booked if eff_tax_booked > 0 else TAX_RATE`). WACC is the `WACC_PCT` board-approved hurdle rate (37%) — adjust it if Finance changes the rate. The rate + capital proxy are stated in the card text so it's transparent. Cost lines (cogs/opex/da/interest/tax) are stored **negative** on the Summary tab; subtotals (ebit/ebitda/pbt/pat) positive — ratio code uses `abs()` on costs accordingly.

**Auto-latest selection:** Finance drops a new `Group Financial Report_<Mon-YY>.xlsx` into `data/` each week. `find_file()` returns the **latest** report, not `glob()[0]`: `_report_period_key()` parses `Mon-YY` from the filename (+ optional `_vN` version suffix) and ranks by (year, month, version, mtime); unparseable names fall back to mtime so they never outrank a dated report. The dashboard regenerates on each load / "Regenerate Dashboards" / 24h refresh, so a newer-dated file is picked up automatically — no hook/watcher needed. Period label and all figures come from inside the chosen workbook, so they update with the file.

**Working-dir refresh (critical):** `prepare_data_folder()` copies most `*.xlsx` only `if not dest.exists()`, but `generated/data_working/` **persists within a running Streamlit Cloud container**, so a once-copied report would never refresh and "Regenerate Dashboards" would keep serving a stale month. Therefore `prepare_data_folder()` **always overwrites** `Group Financial Report*.xlsx` and **prunes** any report not in the repo, each run. This is why deploying a new month is a **code-adjacent data change** — after committing the new `.xlsx`, **Reboot** the app (not just Regenerate) so the container pulls the new repo file.

**Gotchas:**
- The report `.xlsx` is **local-only** (gitignored) and is **NOT** fetched from Google Drive/Sheets. On Streamlit Cloud the tab shows a placeholder until the file is force-added to the repo or wired into the Drive fetch. It is sensitive consolidated financial data — do not commit without sign-off.
- The generator **fails safe**: missing file or unrecognisable Summary tab → writes a placeholder HTML and exits 0 (never crashes the tab).
- YoY uses **magnitude growth** and flags loss↔profit sign flips as "↺ turnaround" (so a doubled cost line reads "▲100%" red, not a misleading "-100%" green).
- Segment/customer/country breakdown totals (~₦1.48B) are **less than** P&L Total Revenue (~₦2.37B) — the breakdowns exclude Other Income and some items. Presented as their own tables; do not assume they reconcile to total revenue.

---

## Key Business Logic

### Revenue Target
- **$8M** is the official company annual revenue target (hardcoded as `LANDING_ZONE` in pipeline and `annual_revenue_target_usd` in revenue dashboard)
- The deal bucket sums to ~$10M (optimistic internal target), but projections/gap calculations use $8M

### Pipeline Status Weights (Projection Model)
```python
STATUS_WEIGHTS = {
    'On Track':  1.00,   # 100% weighted
    'Closed':    1.00,   # 100% weighted
    'At Risk':   0.50,   # 50% weighted
    'Off Track': 0.10,   # 10% weighted
    'Unknown':   0.70,   # 70% weighted
}
```

### Momentum Detection (Pipeline)
Compares the two most recent months to classify deal momentum:
- `growing`: latest > previous × 1.1
- `stalled`: latest = 0 but previous > 0, or both recent months = 0 after earlier activity
- `new`: only the most recent month has revenue
- `steady`: everything else

### YTD Calculation (Fully Dynamic)
**No code changes needed when Finance adds new months.** Both the revenue and pipeline dashboards dynamically detect which months have data by scanning monthly totals across all deals:
```python
monthly_totals = [sum(r['monthly'][i] for r in revenues) for i in range(12)]
months_with_data = [i for i, v in enumerate(monthly_totals) if v > 0]
last_data_month = max(months_with_data)  # e.g., 4 = May
num_months = last_data_month + 1         # e.g., 5
```
YTD labels, run rates, momentum, and all calculations automatically adjust.

### Achievement Percentage (Revenue Dashboard)
**Uses full annual target, NOT pace-adjusted target.** This was changed in June 2026 because Finance expects deals that earned their full annual target to show 100%, not 240%.
```python
# CORRECT: simple progress toward annual goal
achievement_pct = ytd_actual / annual_usd * 100

# WRONG (old calculation — DO NOT USE):
# achievement_pct = ytd_actual / (annual_usd * months_active / 12) * 100
```

### Run Rate (Pipeline Dashboard)
```
complete_months = all months before last_data_month
complete_months_avg = sum(complete_months) / len(complete_months)
last_month_scaled = last_month_actual * (30 / days_elapsed)
monthly_run_rate = max(complete_months_avg, total_with_scaled / num_months)
annual_run_rate = monthly_run_rate * 12
```

### FX Rate
Hardcoded: `FX_RATE = 1450` ($1 = ₦1,450). Used for dual-currency display. Update if Naira rate changes significantly.

### Investment Detection
Cash dashboard separates operational flows from investment movements using keyword matching:
- Investment outflows: categories containing "investment in" or "funding"
- Investment inflows: categories containing "investment withdrawal" or "investment liquidation"

### Cash KPI Cards (Currency Split & Runway)
The Cash Overview KPI row shows three position cards:
- **Total Position (incl. Investments)** = `total_cash_ngn` (NGN equivalent of everything).
- **NGN Balance (incl. Investments)** = `ngn_closing + investment_ngn`, displayed in ₦ via `fmt_naira()`.
- **USD Balance (incl. Investments)** = `usd_raw + investment_usd_raw`, displayed in actual dollars via `fmt_usd()` (NOT NGN-equivalent).

NGN balance + (USD balance × report FX) reconciles to Total Position (GBP, usually ~0, is the only other component). **Operational Runway** = `total position / average operational burn`. The numerator includes investments, so the sub-label reads "at operational burn rate" — do NOT re-add "excl. investments" (the burn rate excludes investment *outflows*, but the runway position does not exclude investments).

---

## Authentication

The app uses a **custom Google OAuth flow** (NOT Streamlit's built-in `st.login()` which is broken on Streamlit Cloud as of May 2026).

### How It Works
1. User visits the app → sees "Sign in with Google" button
2. Click generates a CSRF state token stored in `st.cache_resource` (shared global dict)
3. User is redirected to Google's OAuth consent screen
4. Google redirects back to the app's base URL with `?code=...&state=...`
5. App validates state against the global store, exchanges code for tokens, decodes the JWT id_token to get the user's email
6. Email is checked against `allowed_emails` list in secrets
7. Authenticated email stored in `st.session_state`

### Secrets Structure (Streamlit Cloud → Settings → Secrets)

```toml
[google_oauth]
client_id = "your-google-oauth-client-id"
client_secret = "your-google-oauth-client-secret"
redirect_uri = "https://seamfix-executive-dashboard.streamlit.app/"
allowed_emails = [
    "conwurah@gmail.com",
    "conwurah@seamfix.com",
    "fatube@seamfix.com",
    "lwilfred@seamfix.com",
    "obolade@seamfix.com",
    "cemewulu@seamfix.com",
]

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "seamfix-dashboard@your-project.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."

ANTHROPIC_API_KEY = "sk-ant-..."
```

### Adding/Removing Users
Edit the `allowed_emails` list in Streamlit Cloud secrets. No code change needed.

### Google Cloud Console Setup
- Project must have Google Sheets API and Google Drive API enabled
- OAuth consent screen configured (External or Internal)
- OAuth 2.0 Client ID (Web application type)
- Authorized redirect URI: `https://seamfix-executive-dashboard.streamlit.app/`
- Service account with Viewer access to the Google Sheet and Drive folder

---

## Bobby (AI Chat Assistant)

Bobby is an AI financial analyst powered by Claude (Anthropic API). It lives in the sidebar.

### How It Works
1. On first query, `build_chat_context()` compiles ALL financial data into a ~6-8K token text summary
2. This context is sent as the system prompt with `cache_control: ephemeral` (prompt caching — 90% cost reduction on follow-up messages)
3. Uses `claude-sonnet-4-6` model
4. All queries and responses are logged to a separate Google Sheet (`BOBBY_LOG_SHEET_ID`)

### Usage Logging
Bobby logs to Google Sheet `1c7QMZuV-YNDsmn1XYLJtx8pRyYi6g_wwdAHJ-D0cgtk`:
- **Bobby Queries** worksheet: timestamp, user email, question, response preview, token usage
- **Dashboard Visits** worksheet: timestamp, user email (logged once per session)

---

## Deployment & Operations

### Streamlit Cloud Settings
- **Python version**: 3.12 (select from dropdown in Advanced Settings — do NOT rely on `.python-version` or `runtime.txt`)
- **Main file**: `app.py`
- **Branch**: `main`

### Updating Data

**Revenue/Pipeline data**: Finance updates the Google Sheet directly. The dashboard auto-fetches it on each load (5-minute cache via `@st.cache_data(ttl=300)`).

**Cash reports**: Finance uploads new weekly reports to the shared Google Drive folder. The app fetches all xlsx files from the folder and its subfolders.

**Budget file**: Currently bundled in `data/`. To update, replace `data/2026 LEAN BUDGET.xlsx`. Note: `*.xlsx` is gitignored (see "Repository & Data Policy" below), so committing a refreshed baseline requires `git add -f "data/2026 LEAN BUDGET.xlsx"` before pushing.

### Regenerating Dashboards
Users can click "🔄 Regenerate Dashboards" in the sidebar (under ⚙️ Data & Settings) to force a fresh build. The app also auto-refreshes every 24 hours.

### Adding a New Month of Revenue Data

**No code changes needed.** The revenue and pipeline dashboards dynamically detect which months have data by scanning columns M through X (Jan through Dec). When Finance adds May data to column Q in the Google Sheet, the dashboards will automatically:
- Include May in YTD calculations
- Update "Data as of" labels (e.g., "Jan – May 2026 (May partial)")
- Add May to the momentum chart
- Recalculate run rates, achievement percentages, and trend labels
- Update Bobby's chat context with the new month

Just click "Regenerate Dashboards" in the sidebar or wait for the 24-hour auto-refresh.

### If Finance Adds More Columns to the Excel
The monthly revenue columns (N=Jan through Y=Dec) are read dynamically *within* their range, but the range itself plus all single-column refs (annual=F, start=E, status=L, comments=M, recurring=D, Deficit=Z, Surplus=AA) are hardcoded positional references. If Finance inserts/shifts columns, update `MONTH_COLUMNS` and those `cells.get(...)`/`d.get(...)` refs in both `generate_revenue_dashboard.py` and `generate_pipeline_dashboard.py`. (June 2026: a new col D "Recurring/Not Recurring" shifted everything after C right by one — see Excel Column Mapping above.)

---

## Lessons Learned & Gotchas

### Streamlit Cloud Quirks
1. **Python version**: Only controllable via the settings dropdown, NOT via `.python-version` or `runtime.txt` files
2. **Streamlit version**: Cannot be pinned — Streamlit Cloud always installs the latest version regardless of `requirements.txt`
3. **Built-in OAuth is broken** (as of May 2026): `st.login()` / `st.user` / `[auth]` section causes `MismatchingStateError` crashes. That's why we use a custom OAuth flow with `[google_oauth]` secrets section.
4. **"Regenerate Dashboards" ≠ code reload**: The sidebar button only clears the HTML/data caches. It does NOT reload Python source code. After pushing code changes, you MUST reboot the app from share.streamlit.io (three dots → Reboot).
5. **Deploy keys**: After deleting/recreating the repo, you need to re-establish the deploy key in Streamlit Cloud
6. **Secrets are preserved** across reboots — they persist independently of the repo
7. **Secrets structure**: Must use `[google_oauth]` section (NOT `[auth]` or `[auth.google]`). The old format triggers Streamlit's broken built-in OAuth. See Authentication section for the correct format.

### Collections Tracker — Tab Pinning
- The Collections Tracker workbook has **several near-identical tabs**: the live one is "2026 CRITICAL REVENUE INFLOWS" (gid `1584269897`), but "Revenue Bridge - Data" is a structural mirror, and "Closed" / "Potential Revenue Tracker - IAM" also carry an `S/N` header. **Do not scan by header to pick the tab** — you will sometimes land on the mirror. `app.py` fetches the tab as CSV pinned to `COLLECTIONS_GID`, which is robust to tab renaming/reordering. If Finance ever recreates that tab (deletes + re-adds), its gid changes — update `COLLECTIONS_GID` (find the new gid in the URL: `...#gid=NNNN`).
- The header row is **row 3** (rows 1–2 hold the title and USD rate). The generator's `find_header_row` locates it by scanning for `S/N`. Weekly update columns (`Update - 2nd Jan` … `Update 29th May`) are detected dynamically — adding a new week needs no code change.

### Excel Parsing Fragility
- Column positions are hardcoded (A, B, C, D, E, F, L, M, N–Y, Z, AA). If Finance restructures the sheet, all generators break. (A new column was inserted at D in June 2026 — see Excel Column Mapping.)
- The `TOTAL` row is used as a stop marker. If someone removes it or adds data below it, unexpected rows get parsed.
- Section headers (`ANCHOR DEALS`, `EXISTING CUSTOMERS`, etc.) are skipped by exact match. Case matters.
- Date parsing from cash report filenames uses regex — non-standard filenames will be skipped silently.

### Performance
- First load takes ~20 seconds (data fetch + 6 dashboards generated in waves of 2)
- Subsequent loads use `st.cache_resource` (shared HTML cache across all users)
- Bobby's first query takes longer (cache write); follow-up queries use prompt caching (~90% cheaper)
- Auto-refresh every 24 hours clears all caches
- **Generator concurrency is capped at `max_workers=2`** in `app.py` (was 6). Streamlit Cloud's CPU is heavily throttled/shared; the cash/expense/budget generators each re-parse every accumulated weekly cash report via openpyxl. Running all 6 at once starved them so they all hit the (then 60s) subprocess timeout. Per-subprocess timeout is now 120s. **Do NOT raise `max_workers` back up** — it will reintroduce the timeout as more weekly reports accumulate. If first-load ever needs to be faster, optimize the generators (e.g. parse each cash report once and share), not the worker count.

### Common Issues

| Symptom | Cause | Fix |
|---------|-------|-----|
| Dashboard shows stale data | Cached HTML | Click "Regenerate Dashboards" or wait 24h |
| Code changes not taking effect | "Regenerate Dashboards" only refreshes data/HTML cache, NOT code | **Reboot the app** from share.streamlit.io (three dots → Reboot) |
| Revenue numbers wrong | Column shift in Excel | Check column letters in `generate_revenue_dashboard.py` and `generate_pipeline_dashboard.py` |
| Achievement % over 100% | Calculation using pace-adjusted target instead of annual | Should use `ytd_actual / annual_usd` — see "Achievement Percentage" section above |
| "Data quality errors" banner | Cash report parsing failure | Check filename format of new cash reports |
| OAuth error after reboot | Session state cleared | Normal — user just signs in again |
| Bobby says "API error" | Invalid or expired Anthropic key | Update `ANTHROPIC_API_KEY` in secrets |
| No cash data | Drive folder not shared with service account | Re-share folder with service account email |
| "Gap to Target" shows wrong sign | Projection > target but code shows gap | Check `realistic_gap <= 0` conditional in pipeline generator |

---

## Google IDs & Constants

| Resource | ID | Location in Code |
|----------|-----|-----------------|
| Revenue Google Sheet | `1XKIE9eRP8H1AWpuMAJA0U8bM7pQ9o1jvoQobc6aUn5s` | `app.py` → `GOOGLE_SHEET_ID` |
| Collections Tracker Sheet | `17KE1n5_SOeDXaX96Xsa1JfAjNs_OZX8xu-wYDt4LpU8` | `app.py` → `COLLECTIONS_SHEET_ID` |
| Google Drive folder (cash reports) | `1vLq8m030d1ifL6nAVuo9LT5N9NSeGs9U` | `app.py` → `GOOGLE_DRIVE_FOLDER_ID` |
| Bobby usage log sheet | `1c7QMZuV-YNDsmn1XYLJtx8pRyYi6g_wwdAHJ-D0cgtk` | `app.py` → `BOBBY_LOG_SHEET_ID` |
| Collections Tracker tab gid | `1584269897` ("2026 CRITICAL REVENUE INFLOWS" tab) | `app.py` → `COLLECTIONS_GID` |
| Revenue filename | `2026 Path to Revenue (1).xlsx` | `app.py` → `REVENUE_FILENAME` |
| Budget filename | `2026 LEAN BUDGET.xlsx` | `app.py` → `BUDGET_FILENAME` |
| Collections filename | `2026 Collections Tracker.csv` | `app.py` → `COLLECTIONS_FILENAME` |
| FX Rate | $1 = ₦1,450 | All generator files → `FX_RATE` |
| Annual target | $8,000,000 | Pipeline → `LANDING_ZONE`, Revenue → `annual_revenue_target_usd` |

---

## Development Workflow

### Local Development
```bash
pip install -r requirements.txt
# Create .streamlit/secrets.toml from the example
streamlit run app.py
```

### Pushing Changes
```bash
git add <files>
git commit -m "description"
git push origin main
```
Streamlit Cloud auto-deploys on push to `main`. The app reboots within ~30 seconds.

**Push policy (updated 2026-06-02)**: Claude is authorized to commit AND push directly to `origin/main` on Chibuzor's machine — no need to hand the push back. Still follow the rules above: stage only intentional files, write a proper `Fix:`/`Feature:`/`Update:`/`Docs:` commit message explaining *why*, add a CHANGELOG entry, and test generators locally first. After pushing a **code** change, the app needs a reboot from share.streamlit.io to pick it up (docs/config-only changes don't).

### Testing a Generator Locally
Each generator can be run standalone:
```bash
python3 generate_pipeline_dashboard.py ./data
# Outputs: ./data/pipeline_dashboard.html (open in browser to preview)
```

---

## Contact & Handover Notes

- **Original developer**: Chibuzor (conwurah@gmail.com)
- **Built with**: Claude (Anthropic AI) assistance
- **Primary stakeholder**: CEO / Finance team
- **Access control**: Managed via `allowed_emails` in Streamlit secrets

When making changes, always test the generator scripts locally first (`python3 generate_*.py ./data`) before pushing. The HTML output can be opened directly in a browser for visual verification without needing to run the full Streamlit app.
