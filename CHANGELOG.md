# Changelog — Seamfix Financial Dashboard

> **Rule**: Every code change must add an entry here. Format: date, type, description, files changed.
> **Types**: `Fix`, `Feature`, `Update`, `Refactor`, `Docs`

---

## 2026-06-17 — Update: Group Financials auto-selects the latest weekly report

**Why:** Finance now drops a fresh `Group Financial Report_<Mon-YY>.xlsx` into `data/` weekly (May-26 added). `find_file()` returned `glob()[0]` (arbitrary order), so with multiple reports accumulating the tab could render a stale month. The dashboard already regenerates on each load / "Regenerate Dashboards" / 24h auto-refresh — the only gap was picking the newest report.

**What:** `find_file()` now returns the **latest** matching report. New `_report_period_key()` parses `Mon-YY` from the filename (plus an optional `_vN` version suffix), ranking by (year, month, version, mtime); files with no recognisable period fall back to mtime so they never outrank a dated report. Verified May-26_v2 wins over Apr-26 and May-26 (v2 beats v1). No hook/watcher needed — dropping a newer-dated file is detected automatically on the next regenerate.

**Note:** report `.xlsx` remains **local-only/gitignored** — to surface a new month on Streamlit Cloud it must still be force-added (`git add -f`) and the app rebooted/regenerated.

**Files**: `generate_financial_report_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-11 — Feature: Budget vs Actual rebuilt from Budget Tracker (group / company / department)

**Why:** Finance maintains a richer, Acumatica-loaded budget-vs-actual view in the external **Seamfix Budget Tracker** (https://seamfix-budget-tracker.netlify.app/). The exec team wanted that surfaced in the dashboard's **Budget vs Actual** tab, broken out **group-wide, company-wide (per entity), and department-wide**. The old approach (fuzzy-matching cash-report outflows against `2026 LEAN BUDGET.xlsx`) was approximate and entity-blind; it has been retired.

**What:** Rewrote `generate_budget_dashboard.py` to render three levels from a committed JSON snapshot:
- **Group** KPIs (annual ₦5.10B, YTD budget ₦2.17B, YTD actual ₦1.45B, remaining, projected year-end) + budget-health takeaway.
- **Company-wide** cards + table per entity — Nigeria, United Kingdom, United Arab Emirates — with % of annual vs the time-elapsed pace marker.
- **Department-wide** sortable/searchable table (11 depts) with click-to-expand budget-head drill-down, a YTD-budget-vs-actual bar chart, a group monthly budget-vs-actual line chart, and executive takeaways (over-pace depts, underspenders, largest spend lines).

**Key design notes:**
- **Snapshot, not live fetch:** the Netlify tracker is fully client-side (hardcoded JS data, manual "Run" refresh, no API). Data extracted into `data/budget_tracker_snapshot.json` (plain `.json`, not gitignored → ships to Streamlit Cloud). `app.py`'s `prepare_data_folder()` now copies it into the working dir each run. Current snapshot = `Run_004` (2026-05-30), actuals through Apr-2026.
- **Lean mode only:** the tracker's "lean" (Acumatica-loaded) budget carries actuals; "full" (approved) has none, so lean is the only basis for budget-vs-actual.
- **NGN, bottom-up:** Group = Σ entities = Σ departments. NG budgets already NGN; **UK (GBP) and UAE (USD) budgets FX-converted** at lean FX (GBP=2000, USD=1500); **actuals are already NGN** for all entities. FY reconciles exactly to the tracker's ₦5.10B; YTD actual reconciles bottom-up to ₦1.45B (tracker headline ₦1.47B is ~1.4% higher — includes mapped txns not allocated to a dept head; we use the reconciling figure).
- Generator **fails safe**: missing/unreadable snapshot → placeholder HTML, exit 0.

**Files**: `generate_budget_dashboard.py` (rewritten), `data/budget_tracker_snapshot.json` (new), `app.py` (copy snapshot to working dir), `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-08 — Feature: Group Financials dashboard (consolidated P&L + profitability vs targets)

**Why:** Finance dropped the consolidated **Group Financial Report** (`Group Financial Report_Apr-26.xlsx`) into `data/` and the exec team needs it surfaced as a dashboard tab — highlighting **Net Profit Margin vs the 10% target** and **Gross Profit Margin vs the 70% target**, plus expense analysis and the other metrics leadership tracks.

**What:** New generator `generate_financial_report_dashboard.py` + new **"Group Financials"** tab (🏦) in `app.py`'s `DASHBOARDS` dict (output `financial_report_dashboard.html`). Parses the **GROUP INCOME STATEMENT** block on the report's **`Summary`** tab. Renders:
- KPI row: Total Revenue (+154% YoY), Gross Profit, EBITDA, Profit After Tax (loss→profit turnaround).
- Two target cards with gauges: **Net Profit Margin 23.9% vs 10% target** (✓ +13.9pts) and **Gross Profit Margin 78.2% vs 70% target** (✓ +8.2pts).
- Auto-generated Critical Insights (margins vs target, turnaround, ARR contraction, NIMC customer concentration ~71%, OpEx discipline).
- Income Statement table (current vs prior YTD, NGN + USD, YoY), Expense Analysis (Payroll/COGS/Marketing/D&A/Other as % of revenue), Revenue by Vertical (doughnut + table)/Country/Top Customers, ARR + balance-sheet highlights.

**Key design notes:**
- Parsing is **label-driven** (scans column **K** of the `Summary` tab for line-item names) so it survives row insertions — same robustness lesson as the revenue sheet. GROUP columns: **K**=label, **L**=current NGN, **M**=prior NGN, **N**=current USD, **O**=prior USD.
- Uses the **report's own USD columns** (period-average FX), NOT the dashboard-wide `FX_RATE=1450`.
- **YoY uses magnitude growth** and flags loss↔profit sign flips as "↺ turnaround" — so cost lines that double read "▲100%" (red) instead of a misleading "-100%" (green).
- Generator **fails safe**: if the file is missing or the Summary tab isn't recognisable, it writes a friendly placeholder HTML and exits 0 (never crashes the tab).
- **DEPLOYMENT CAVEAT:** the report `.xlsx` is **local-only** (gitignored) and is **not** fetched from Google Drive/Sheets, so on Streamlit Cloud this tab will show the placeholder until the file is either (a) force-added to the repo or (b) wired into the Drive/Sheets fetch. Sensitive consolidated financials — do not commit without sign-off.

**Files**: `generate_financial_report_dashboard.py` (new), `app.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-08 — Update: Remove "ARR Mix (% of Revenue)" card, keep only "ARR As At <month>"

**Why:** Finance wanted only the actuals-based ARR percentage. The plan-based "ARR Mix (% of Revenue)" card (recurring share of annual *targets*, 36%) was redundant alongside "ARR As At <month>" (recurring share of revenue *actually earned* YTD, 19% vs 50% target).

**What:** Removed the "ARR Mix (% of Revenue)" KPI card and its now-unused `arr_pct` computation from `generate_revenue_dashboard.py`. Kept "ARR As At <month>". `ARR_TARGET_PCT` (50) is still used by the remaining card.

**Files**: `generate_revenue_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-08 — Update: Replace duplicated "YTD Achievement" card with "ARR As At <month>"

**Why:** Once the revenue column-shift fix made total pipeline (`total_stream_annual`) equal the $8M company target, the **"Annual Progress"** and **"YTD Achievement"** KPI cards showed the identical 26% / "$2.12M of $8.00M" — a visible duplicate.

**What:** Replaced the redundant "YTD Achievement" card with **"ARR As At <month>"** = `recurring_ytd_usd / ytd_actual_revenue_usd` (`arr_asat_pct`) — the recurring share of revenue *actually earned* year-to-date, measured against the same **50% target** (red + "▼ N pts below" when under). Currently **19%** ($394.8K recurring of $2.12M earned, Jan–Jun 2026). This is distinct from the existing plan-based "ARR Mix (% of Revenue)" card (36%): ARR Mix uses annual targets (col F), ARR As At uses realized monthly actuals (cols N–Y). "Annual Progress" is kept; `ytd_achievement_rate`/`active_streams` remain used in the health summary + table.

**Files**: `generate_revenue_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-08 — Feature: ARR Mix % card (recurring share of revenue vs 50% target)

**Why:** Finance wants to see the current ARR as a percentage against a 50% target.

**What:** Added an **"ARR Mix (% of Revenue)"** KPI to the Revenue & Fundability dashboard = `arr_usd / total_stream_annual × 100` (`arr_pct`), shown against a `ARR_TARGET_PCT = 50` target. Card turns red and shows "▼ N pts below" when under target. Currently **36%** ($2.86M recurring of $8.00M total revenue) — 14 pts below the 50% target. Note: total annual revenue currently equals the $8M company target, so the percentage is the same whether measured against total revenue or the company target.

**Files**: `generate_revenue_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-08 — Fix + Feature: Revenue sheet column shift (new col D) + ARR card

**Why:** Finance inserted a new column **D ("Recurring/Not Recurring")** in the live "2026 Path to Revenue" sheet, which shifted every column after C one to the right. The generators were still reading the old positions, so the live Revenue & Fundability and Pipeline Intelligence dashboards were silently reading the **wrong columns** (annual revenue, status, monthly actuals, deficit/surplus all off by one). This corrects the mapping and adds the requested ARR card.

**Column remap (old → new):** Annual Revenue E→**F**, Start Date D→**E**, Status K→**L**, Comments L→**M**, monthly actuals Jan–Dec M–X→**N–Y**, Deficit Y→**Z**, Surplus Z→**AA**. New col **D = Recurring/Not Recurring**.

**Feature — ARR card:** Added an **ARR (Annual Recurring Revenue)** KPI to the Revenue & Fundability dashboard = sum of column F (2026 annual revenue) for deals flagged "Recurring" in column D (excludes Not-Recurring/one-time deals). Current value ≈ **$2.86M** across 17 recurring streams.

**Tested:** Ran both generators locally against the refreshed sheet — 43 streams parsed; ARR $2.86M, YTD actual $2.12M, annual progress 26%, statuses parse correctly (14 On Track / 6 At Risk / 4 Off Track), data range "Jan – Jun 2026 (Jun partial)". Bobby's context is unaffected since it reuses `gen_pipe.extract_revenue_data()`.

**Files**: `generate_revenue_dashboard.py`, `generate_pipeline_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-08 — Fix: Dashboards timing out on load (all 6 hit 60s subprocess timeout)

**Symptom:** The live app showed `TimeoutExpired: Command [... generate_dashboard.py ...] timed out after 60 seconds`, and every tab failed to render.

**Root cause:** `app.py` launched all 6 generator subprocesses at once (`ThreadPoolExecutor(max_workers=6)`), each with a 60s wall-clock timeout. Streamlit Community Cloud runs on a heavily throttled/shared CPU, and the cash, expense, and budget generators each re-parse *every* accumulated weekly cash report via openpyxl (CPU + memory heavy). As more weekly reports piled up through the year, 6 simultaneous openpyxl processes starved each other badly enough that all of them blew past 60s together. Locally (12 reports, unthrottled CPU) each generator finishes in ≤7s, so it was invisible in local testing. This was a latent scaling issue, not caused by the recent NGN/USD or theme changes — `max_workers` and the 60s timeout predate both.

**Fix:**
- Reduced generator concurrency from `max_workers=6` to `max_workers=2` so the heavy generators no longer contend for the throttled CPU.
- Raised the per-subprocess timeout from 60s to 120s for headroom.

**Trade-off:** First-load is a little slower (generators now run ~3 waves of 2 instead of all at once) but completes reliably. Subsequent loads use the shared HTML cache.

**Files**: `app.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-08 — Update: Cash Overview — split NGN/USD balances + clarify runway label

**Why:** Finance wanted the single "Total Position" figure broken out by currency so the NGN-denominated and USD-denominated holdings are visible at a glance, while still keeping the overall total in NGN. They also flagged that the runway KPI's "excl. investments" note was misleading — the runway numerator is the *total* position (which includes investments), so the parenthetical wrongly implied investments were excluded.

**What changed:**
- Added two KPI cards next to **Total Position (incl. Investments)**:
  - **NGN Balance (incl. Investments)** = `ngn_closing + investment_ngn`, shown in ₦.
  - **USD Balance (incl. Investments)** = `usd_raw + investment_usd_raw`, shown in $ (actual dollars, not NGN-equivalent). Each card shows week-over-week % change.
- Added a `fmt_usd()` helper (mirrors `fmt_naira()`) for dollar-suffixed formatting.
- Runway KPI sub-label changed from "(N months at op. burn excl. investments)" to "(N months at operational burn rate)". Logic unchanged — runway is still `total position / operational burn`.

**Note:** NGN + (USD × report FX) reconciles to Total Position (GBP, usually ~0, is the only other component). Verified locally against the 12 bundled cash reports: Total ₦2.42B = NGN ₦228.2M + USD $1.6M.

**Files**: `generate_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-02 — Feature: Light/dark theme toggle across all 6 dashboards

Added a shared theme system (`theme.py`) with CSS custom properties for light and dark modes. **Light mode is now the default** — the dashboards were previously dark-only. Users can toggle between light and dark mode via a button in the top navigation bar. Preference is saved to localStorage and persists across sessions. Chart.js charts also update their grid/text colors on toggle.

**Files**: `theme.py` (new), `generate_revenue_dashboard.py`, `generate_pipeline_dashboard.py`, `generate_dashboard.py`, `generate_expense_dashboard.py`, `generate_budget_dashboard.py`, `generate_collections_dashboard.py`
**Author**: Chibuzor + Claude

---

## 2026-06-02 — Docs: Add "pull latest first" rule to CLAUDE.md

Added Rule #0 requiring every Claude instance to run `git pull origin main` before starting any work, to prevent overwriting changes from other contributors.

**Files**: `CLAUDE.md`
**Author**: Chibuzor + Claude

---

## 2026-06-02 — Feature: Collections Intelligence dashboard (6th tab)

Added a new **Collections Tracker** dashboard tab so the executive team can track critical revenue inflows in one place, see what moved week-over-week, and act on stalled or at-risk collections.

**What it shows:**
- **KPI row**: Total tracked ($2.67M / ₦3.88B), fully collected, in-progress, and payment-pending amounts with item counts and % of portfolio.
- **Critical Actions panel**: Automatically surfaces items needing urgent action (high-value pending items, "at risk" flags, overdue deadlines, specific deliverables assigned), sorted by USD value. Colour-coded red (critical) and amber (follow-up needed).
- **What Moved This Week**: Compares the last two weekly update columns for every item and highlights deals where something changed — green for positive movement (payment received, deal signed), red for negative flags (payment at risk, extension granted, no response), blue for general updates.
- **Full tracker table**: All 20 items with payment status, deal status, predictability, accountable party, latest update, and actionable deliverable. Filterable by status.

**Data source**: Google Sheet `17KE1n5_SOeDXaX96Xsa1JfAjNs_OZX8xu-wYDt4LpU8`, specifically the **"2026 CRITICAL REVENUE INFLOWS" tab (gid `1584269897`)** — fetched live each page load (5-min cache, same pattern as the revenue sheet). The sheet must be shared with the service account or set to "anyone with the link can view".

**Weekly updates auto-reflect** with zero code changes (same as the 2026 Path to Revenue dashboard): the generator detects every `Update - <date>` column dynamically, so when Finance adds next week's column it automatically becomes the new "latest", and the "What Moved This Week" comparison shifts forward. Verified by simulating an added "Update 5th Jun" column.

**How it works**: `app.py` fetches **only that tab as CSV pinned to its gid** (`fetch_google_sheet_csv`) and saves it to the data_working folder as `2026 Collections Tracker.csv`. **Why CSV-by-gid, not whole-workbook xlsx**: the workbook contains several near-identical tabs ("Revenue Bridge - Data" is a structural mirror, plus "Closed" and "IAM" tabs that also carry an S/N header). Scanning by header would risk landing on the wrong tab; pinning to the gid is robust to tab renaming and reordering. The generator (`generate_collections_dashboard.py`) reads the CSV (header in row 3, after the title/USD-rate rows), maps columns by name (not position), detects the 22 weekly-update columns, runs movement/urgency analysis, and outputs a self-contained HTML dashboard. Tested locally against live data — 20 items, $2.67M total.

**Files**: `generate_collections_dashboard.py` (new), `app.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-02 — Docs: Add "Can Claude commit and push for me?" subsection to onboarding

Clarifies what a non-technical collaborator needs in order to push:
- Explains the **commit (local) vs push (remote)** distinction — Claude can always commit; pushing needs repo access.
- Lists the three one-time requirements: accepted collaborator invite, SSH key, and git identity.
- Notes that on a **personal** GitHub repo, a collaborator automatically gets push access — there is no Read/Write dropdown (that only appears on org repos). This answers the "I can't see read/write options" question directly.
- Adds a "permission denied" troubleshooting note.
- Applied to both `ONBOARDING.md` and the Word `.docx`.

**Files**: `ONBOARDING.md`, `Seamfix Dashboard - Onboarding Guide.docx`, `CHANGELOG.md`
**Author**: Chibuzor + Claude

---

## 2026-06-02 — Docs: Add plain-English basics section to onboarding; switch tool guidance to Claude Code

Ahead of onboarding the Head of Finance (a non-technical user):
- **Added a "Brand-New to All This? Read This First" section** to the top of the onboarding guide — plain-English explanations of GitHub, repo, clone, Streamlit, the terminal, git/commit/push, `CLAUDE.md`, and the data → code → website big picture. Lets a complete newcomer understand the system before the setup steps.
- **Switched the recommended Claude tool from Cowork to Claude Code.** Both live in the Claude desktop app; Claude Code is preferred because its usage limits are more generous (Cowork hits limits faster). Updated all setup steps and references accordingly, with a short note explaining the choice.
- Applied the same changes to both `ONBOARDING.md` and the Word version (`Seamfix Dashboard - Onboarding Guide.docx`) via in-place XML edits so formatting is preserved. (Repacked manually as a zip because the local Python is 3.9 and the docx skill's pack script needs 3.10+; XML well-formedness and zip integrity were verified.)

**Files**: `ONBOARDING.md`, `Seamfix Dashboard - Onboarding Guide.docx`, `CHANGELOG.md`
**Author**: Chibuzor + Claude

---

## 2026-06-02 — Docs: Flatten folder layout, fix onboarding repo refs, allow Claude to push

Follow-up to the consolidation below, ahead of onboarding the Head of Finance:
- **Folder layout flattened** to `seamfix/seamfix-dashboard` (removed the outer `financial dashboards` wrapper folder). No code paths were hardcoded, so nothing broke. Loose reference material (PDF exports, planning docx, a superseded older `app.py`, the OAuth patch, the secrets `.rtf`) was gathered into a gitignored `docs/` folder; duplicate copies of tracked files were removed.
- **ONBOARDING.md corrected**: it pointed at `seamfix/finance-dashboard` on branch `snapshot-dev` (the stale org repo). Updated to the live `sfxconwurah/seamfix-dashboard` on `main`, with a note that the org migration is pending CTO access. Added a tip about the local `UPDATE_DASHBOARD.command` preview.
- **Push policy**: Claude may now commit and push directly to `origin/main` (see CLAUDE.md → Development Workflow).
- Removed an inaccurate "Reported by: Lilian Wilfred" attribution on the achievement-% fix below (she had not yet seen the dashboards).
- `.gitignore`: added `docs/`.

**Files**: `ONBOARDING.md`, `CLAUDE.md`, `CHANGELOG.md`, `.gitignore`
**Author**: Chibuzor + Claude

Also updated `Seamfix Dashboard - Onboarding Guide.docx` (the Word version of the onboarding guide) with the same repo-reference corrections, via in-place XML edits so formatting is preserved.

---

## 2026-06-02 — Update: Consolidate to a single repository; keep financial data local-only

Established this repo (`seamfix-dashboard`) as the **single source of truth**. Previously the project was duplicated inside an outer working folder/repo (`sfxconwurah/financial-dashboards`) that held a second copy of every generator + an old `app.py`, kept in sync by a manual `cp` step. That copy had drifted (the outer copies were stale relative to deployment), which was a recurring source of confusion and bugs. The outer folder has been retired; all work now happens here.

Also added `*.xlsx` and `*.rtf` to `.gitignore`: financial data is now local-only going forward (the live app reads from Google Drive/Sheets). The existing `data/*.xlsx` baseline remains tracked as the offline fallback; new/updated xlsx will not be committed unless force-added (`git add -f`).

**Files**: `.gitignore`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Chibuzor + Claude

---

## 2026-06-02 — Docs: Add project rules, onboarding guide, and changelog

Added mandatory documentation rules to CLAUDE.md (auto-documentation, commit standards, testing requirements). Created ONBOARDING.md and Word document guide for finance team onboarding with Claude Cowork. Created this CHANGELOG.md.

**Files**: `CLAUDE.md`, `ONBOARDING.md`, `CHANGELOG.md`, `Seamfix Dashboard - Onboarding Guide.docx`  
**Author**: Chibuzor + Claude

---

## 2026-06-02 — Fix: Achievement % showing 200-240% for completed deals

Changed `achievement_pct` calculation in the Revenue & Fundability dashboard from pace-adjusted (`ytd_actual / ytd_target_pace`) to simple annual progress (`ytd_actual / annual_usd`). Finance expects deals that earned their full annual target to show 100%, not 240%. Also updated the gap calculation, underperformer thresholds, and KPI labels to match.

**Root cause**: The old formula divided YTD actual by a pro-rated target (annual × months/12). A deal earning $100K of a $100K target by May showed: $100K / ($100K × 5/12) = 240%.

**Files**: `generate_revenue_dashboard.py`  
**Author**: Chibuzor + Claude

---

## 2026-06-02 — Fix: Re-enable Google OAuth authentication

Updated Streamlit Cloud secrets from old `[auth]`/`[auth.google]`/`[auth.disabled]` format to new `[google_oauth]` format. Added base URL redirect URI (`https://seamfix-executive-dashboard.streamlit.app/`) to Google Cloud Console. Authentication now working with custom OAuth flow.

**Files**: Streamlit Cloud secrets (no code change — auth code was already updated in May)  
**Author**: Chibuzor + Claude

---

## 2026-05-09 — Feature: Dynamic month detection (no more monthly code updates)

Replaced all hardcoded month references (e.g., "Jan–Apr", `months_active = 4`) with dynamic detection. Both revenue and pipeline dashboards now scan columns M through X to determine which months have data. YTD labels, run rates, momentum, chart datasets, and table headers all adjust automatically.

**Before**: Every month required code changes to add the new column, update labels, and recalculate.  
**After**: Finance adds data to Google Sheet → click "Regenerate Dashboards" → done.

**Files**: `generate_revenue_dashboard.py`, `generate_pipeline_dashboard.py`, `app.py` (Bobby context)  
**Author**: Chibuzor + Claude

---

## 2026-05-09 — Feature: Custom Google OAuth flow

Replaced Streamlit's broken built-in `st.login("google")` with a custom OAuth2 flow. Streamlit 1.57's internal Authlib integration causes `MismatchingStateError` crashes on the OAuth callback. The custom flow uses `urllib.request` to exchange codes and `st.cache_resource` for CSRF state storage.

**Files**: `app.py`, `requirements.txt` (removed Authlib dependency), `.streamlit/secrets.toml.example`  
**Author**: Chibuzor + Claude

---

## 2026-05-09 — Docs: Created CLAUDE.md project documentation

Comprehensive technical documentation covering architecture, data flow, Excel column mappings, business logic, authentication, deployment, and common issues. Designed to give any new developer or AI assistant full context to maintain the project.

**Files**: `CLAUDE.md`  
**Author**: Chibuzor + Claude

---

## 2026-04-30 — Feature: Initial 5-dashboard suite launch

Launched the complete Seamfix Financial Intelligence Suite with 5 interactive dashboards (Cash Overview, Expense & Vendor Analysis, Budget vs Actual, Revenue & Fundability, Pipeline Intelligence) plus Bobby AI chat assistant. Deployed on Streamlit Community Cloud with Google Sheets/Drive integration.

**Files**: All files (initial release)  
**Author**: Chibuzor + Claude
