# Seamfix Financial Intelligence Suite — Project Documentation

> **Purpose**: This document gives a new developer or AI assistant full context to maintain and extend this project. Read this first.

---

## MANDATORY: Rules for Any AI Assistant Working on This Project

**These rules apply to every Claude instance, every session, every change — no exceptions.**

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

The Google Sheet column layout (M=Jan through X=Dec, Y=Deficit, Z=Surplus) is the single most fragile part of this system. If Finance changes the sheet structure:
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
├── requirements.txt                # Python dependencies
├── runtime.txt                     # Python version hint (ignored by Streamlit Cloud)
├── .python-version                 # Python version hint (ignored by Streamlit Cloud)
├── .streamlit/
│   ├── config.toml                 # Theme + server config
│   └── secrets.toml.example        # Template for secrets
├── data/                           # Bundled xlsx files (fallback if Drive/Sheets unavailable)
│   ├── Cash Report as at *.xlsx    # Weekly cash position reports
│   ├── 2026 Path to Revenue (1).xlsx   # Revenue & pipeline tracker
│   └── 2026 LEAN BUDGET.xlsx       # Annual budget breakdown
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
| D | Start date |
| E | Annual revenue target (USD) |
| K | Status (On Track / At Risk / Off Track / Closed) |
| L | Comments (free text from finance) |
| M | January actual (USD) |
| N | February actual (USD) |
| O | March actual (USD) |
| P | April actual (USD) |
| Q–X | May–December actual (USD) — added by Finance in April 2026 |
| Y | Deficit (USD) |
| Z | Surplus (USD) |

**IMPORTANT**: In March 2026, Finance added monthly columns for Apr–Dec (Q through X). This shifted the Deficit column from P→Y and Surplus from Q→Z. If Finance adds more columns in future, these will shift again — search for `cells.get('Y')` and `cells.get('Z')` to update.

### Cash Report Files (`Cash Report as at *.xlsx`)

Each weekly cash report has a fixed structure parsed by `generate_dashboard.py`:
- NGN closing balances
- USD closing balances
- Inflow items (categorized)
- Outflow items (categorized)
- Investment portfolio positions

The date is extracted from the filename using regex: `(\d+)\w*\s+(Month)\s+(\d{4})`.

### Budget File (`2026 LEAN BUDGET.xlsx` → "Budget Summary" tab)

Contains the annual ₦5.1B budget broken into categories. Mapped against actual cash outflows using fuzzy matching in `generate_budget_dashboard.py`.

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
The monthly revenue columns (M=Jan through X=Dec) are read dynamically. However, Deficit (column Y) and Surplus (column Z) are still hardcoded positional references. If Finance inserts columns that shift Y/Z, search for `cells.get('Y')` and `cells.get('Z')` in both `generate_revenue_dashboard.py` and `generate_pipeline_dashboard.py` to update them.

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
- Column positions are hardcoded (A, B, C, E, K, L, M, N, O, P, Y, Z). If Finance restructures the sheet, all generators break.
- The `TOTAL` row is used as a stop marker. If someone removes it or adds data below it, unexpected rows get parsed.
- Section headers (`ANCHOR DEALS`, `EXISTING CUSTOMERS`, etc.) are skipped by exact match. Case matters.
- Date parsing from cash report filenames uses regex — non-standard filenames will be skipped silently.

### Performance
- First load takes ~10 seconds (data fetch + 5 dashboards generated in parallel)
- Subsequent loads use `st.cache_resource` (shared HTML cache across all users)
- Bobby's first query takes longer (cache write); follow-up queries use prompt caching (~90% cheaper)
- Auto-refresh every 24 hours clears all caches

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
