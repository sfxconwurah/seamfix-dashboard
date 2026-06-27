# Changelog — Seamfix Financial Dashboard

> **Rule**: Every code change must add an entry here. Format: date, type, description, files changed.
> **Types**: `Fix`, `Feature`, `Update`, `Refactor`, `Docs`

---

## 2026-06-27 — Feature: Cash Overview now consolidates UK & UAE balances (group liquidity)

**Why:** Starting with the 26-Jun-2026 cash report, Finance added two new sheets — `UK & UAE` and `Cash UK & UAE` — that consolidate every entity's cash into one ₦ view with an executive **encumbered vs. available** liquidity split. Previously the Cash Overview tab was Nigeria-only, so the group's UK & UAE cash (~₦79.8M) was invisible and the Total Position understated the group.

**What** (`generate_dashboard.py`):
- New `extract_group_balances(wb)` parser (after `extract_cash_summary`): label-driven read of the `Cash UK & UAE` sheet (col B=label, C=local ccy, D=FX, E=Amount ₦, F=Encumbered ₦, G=Available ₦). Returns per-entity rows, the `GRAND TOTAL` gross/encumbered/available figures, the AIF portfolio (₦ + $), and `uk_uae_ngn` (sum of UK & UAE entity NGN only). Returns `None` when the sheet is absent (older reports).
- `extract_report` attaches `rec['group']` + `rec['uk_uae_ngn']` and **adds `uk_uae_ngn` to `total_cash_ngn`** (Nigeria cash is already counted via the legacy Cash Report tab, so only the incremental UK & UAE entity cash is added). `main()`'s total recompute updated to match. Total Position for 26-Jun reconciles to Finance's GRAND TOTAL ₦2.078B within 0.01% (FX/rounding noise between the two sheets).
- UI: new **Available Liquidity (Group)** KPI card (conditional) showing available ₦ with an encumbered-of-gross sub-label; new **Group Cash by Entity** section — a per-entity table (Amount/Encumbered/Available ₦) with a GRAND TOTAL row mirroring Finance's exact headline figures, plus an AIF portfolio note.
- **Backward compatible:** reports without the group sheet → `group=None`, `uk_uae_ngn=0`, no new cards — the Nigeria-only view is unchanged.

---

## 2026-06-23 — Fix: Group Financials trend chart showed negative EBITDA above positive Net Profit

**Why:** The Monthly Performance Trend chart plotted EBITDA below zero while Net Profit was positive — impossible, since Net Profit sits below EBITDA in the P&L. Root cause: `extract_monthly_trend()` read the `MoM` tab's `EBITDA` row directly, but that row is corrupt in the source workbook (wrong sign/magnitude — e.g. Jan-26 showed −₦193.7M when true EBITDA is +₦188.5M).

**What** (`generate_financial_report_dashboard.py`, `extract_monthly_trend`):
- Stopped reading the broken MoM `EBITDA` row. EBITDA is now **derived using the Summary tab's own formula**: `EBITDA = Gross Profit + Other Income − Total Operating Expenses` (reads the MoM `Gross Margin`, `Other Income` and `TOTAL OPERATING EXPENSES` rows). This reconciles exactly to PAT every month (Jan +188.5 → PAT +181.2, Feb +281.2 → +290.9, … May −53.7 → −154.9).
- Note: the Summary YTD EBITDA (₦783.1M, Jan–Jun) does not equal the MoM monthly sum (₦493.4M, Jan–May) because the MoM tab lags one month — both are correct for their period.

---

## 2026-06-23 — Feature: Glossary & Definitions tab + GLOSSARY.md (single source)

**Why:** Finance wanted a plain-language reference for every metric across the suite — and how each one applies to Seamfix — to read alongside the dashboards each week so the numbers are interpreted consistently.

**What** (`generate_glossary_dashboard.py`, new):
- New generator holds all glossary content as one `GLOSSARY` data structure (8 areas: General & Cross-Cutting, Cash Overview, Expense & Vendor Analysis, Budget vs Actual, Revenue & Fundability, Pipeline Intelligence, Collections Tracker, Group Financials; 66 terms). Each term has a plain **Definition** plus an **At Seamfix** note.
- Emits **two** outputs from the one source so they never drift: `glossary_dashboard.html` (the in-app tab — themed via `theme.py`, with a live search box, category nav chips, and a responsive term-card grid) and `GLOSSARY.md` (the repo doc, written with `--markdown`). No new dependency (markdown rendered by hand).
- Wired a **"Glossary & Definitions"** tab (📖) into `app.py`'s `DASHBOARDS` dict, last. Takes no live data, so it always renders.

**Files:** `generate_glossary_dashboard.py` (new), `GLOSSARY.md` (new), `app.py` (DASHBOARDS dict), `CLAUDE.md`, `CHANGELOG.md`. Tested locally: HTML + markdown both generate; previewed light/dark + search filter.

---

## 2026-06-23 — Fix: Bobby chat crashed ("no attribute 'BUDGET_CATEGORIES'") + now covers every dashboard

**Why:** Ask Bobby returned a "Data Load Failure — module 'generate_budget_dashboard' has no attribute 'BUDGET_CATEGORIES'" message and could not answer any question. The Budget vs Actual dashboard was rewritten on 2026-06-11 to read the committed JSON snapshot, which removed `BUDGET_CATEGORIES` / `map_expense_to_budget` / `is_investment_outflow` — but `build_chat_context()` still referenced them, so building Bobby's context raised and the error went into the context itself. Finance also asked that Bobby be able to speak to the *entire* executive dashboard.

**What** (`app.py`):
- **Fix:** rewrote the Budget vs Actual section of `build_chat_context()` to read `budget_tracker_snapshot.json` via `generate_budget_dashboard.compute()` (group/entity/department YTD budget-vs-actual, all NGN) — the same source the dashboard uses. No more reference to the retired `BUDGET_CATEGORIES` API.
- **Feature:** added three previously-missing sections so Bobby covers all dashboards — **Expense & Vendor** (YTD spend, weekly burn, spend-by-category, top vendors via `process_all_files`/`calculate_kpis`), **Collections Tracker** (`extract_collections` — tracked deals, booked vs expected, payment status), and **Group Financials** (`extract_financials` — consolidated P&L, margins vs target, revenue breakdowns, balance-sheet highlights). The three new generators are imported defensively (`_try_import`), and every section guards on its own data source so a missing file degrades only that section, never the whole context.
- **Fix (robustness):** `call_claude()`'s two early returns (anthropic not installed / no API key) now return a full 5-tuple — the caller always unpacks 5 values, so a bare string there would have raised a `ValueError` and crashed the chat instead of showing the warning.

**Files:** `app.py`. Tested locally: context builds with Cash, Budget, Expense and Group Financials sections (Pipeline/Collections only render where their live data is present); reconciles to known figures (budget group ₦5.10B annual / ₦1.45B YTD; Jun-26 P&L revenue ₦3.01B, net margin 20.1%).

---

## 2026-06-22 — Update: Cash Overview balances shown to 2 decimals + separate MTN Shares card

**Why:** Finance wanted the position balances displayed in full to the kobo/cent (not abbreviated B/M/K), and the MTN equity holding surfaced as its own KPI rather than hidden.

**What** (`generate_dashboard.py`):
- New `fmt_naira_precise()` / `fmt_usd_precise()` — same B/M/K suffix as `fmt_naira()`/`fmt_usd()` but **2 decimals across all magnitudes** so the exact figure shows. The three balance cards now use them: **Total Position** (₦2.12B), **NGN Balance** (₦100.50M), **USD Balance** ($1.49M). Flow cards (inflow/outflow/net/forecast) keep the original compact format (1 dp on M, 0 dp on K).
- New **"MTN Shares (Market Value)"** KPI card (19-Jun ₦96.00M, with WoW change), placed after the USD Balance card. `extract_cash_summary()` now also reads the **"MARKET VALUE OF MTN SHARES"** section of the Cash Balance Summary sheet (latest "Market value as at …" row; figure column shifts F↔G across layouts, so it takes the rightmost numeric). MTN is **excluded from Total Position / cash positions** — it's listed equity, not a cash-equivalent — and the card label says so. Card only renders when a value is present (fails safe).

**Files:** `generate_dashboard.py`.

---

## 2026-06-22 — Fix: Cash Overview USD investment omitted the "Other Dollar Investment" block

**Why:** The "USD Balance (incl. Investments)" KPI's investment component read the working **"Cash Report"** tab's `TOTAL INVESTMENT (USD)` line (col H), which only covers the AIF mutual-fund block. It **omitted the separate "Other Dollar Investment"** (e.g. the $41,652.57 FBNQuest holding earmarked for MinIo/Glo licenses) that Finance itemises on the **"Cash Balance Summary"** sheet. The Cash Report tab's AIF total also disagrees with the summary's.

**What** (`generate_dashboard.py`):
- New `extract_cash_summary(wb)` parses the authoritative **"Cash Balance Summary"** sheet (col B=label, C=Principal/Naira, D=Interest/USD, E=Total). It sums each investment block by section header — **AIF Investment Fund** (USD) + **Other Dollar Investment** (USD) → `usd_invest`; **Other Naira Investments** → `ngn_invest` — skipping MTN shares (listed equity, not cash-equivalent). Fails safe to `None` if the sheet is absent.
- `extract_report()` now **overrides `investment_usd_raw`** with the summary's `usd_invest` when present, so the USD investment = AIF + Other Dollar (19-Jun: $1,314,034.38 + $41,652.57 = **$1,355,686.95**, vs the old $1,411,914.93 which both omitted the Other Dollar line and overstated AIF).
- **NGN investment is deliberately left on the Cash Report tab.** The summary's "Other Naira" block is incomplete in older reports (it omits the large NGN fixed deposits the Cash Report tab carries), so changing it would regress the NGN figure / trend.
- Source-of-truth chosen by Finance: use the Cash Balance Summary sheet for the USD investment.

**Files:** `generate_dashboard.py`.

---

## 2026-06-22 — Fix: Cash Overview "USD Balance" showed the OPENING USD position, not the closing

**Why:** The "USD Balance (incl. Investments)" KPI was reading column **D** of the "TOTAL CASH (USD)" row, which is the **OPENING** balance (header "OPENING BALANCE" — equal to the prior week's closing), not the current closing. So the card always lagged a week behind the actual USD cash position.

**What** (`generate_dashboard.py`):
- The USD cash row's col D is now stored as `usd_opening_raw` (documented as the opening balance), and `usd_raw` (the closing USD position that feeds the KPI) is now **derived from the closing NGN (col J) at the report's own FX rate**: `usd_raw = usd_closing_ngn / fx_rate`.
- This exactly reconciles to the report's authoritative **"CASH BALANCE" summary block** USD row in all 12 reports (e.g. 20-Mar: derived = **$357,336.67** = summary value, vs the old col-D opening $354,445.97).
- Investment USD (col H of the "TOTAL INVESTMENT (USD)" row) was already correct — unchanged.

**Files:** `generate_dashboard.py`.

---

## 2026-06-22 — Feature: Group Financials monthly trend chart + "Profit After Tax" renamed to "Net Profit"

**Why:** Execs wanted to see performance momentum across previous months (not just YTD vs prior-year), and preferred the simpler "Net Profit" nomenclature over "Profit After Tax".

**What** (`generate_financial_report_dashboard.py`):
- **Monthly Performance Trend chart.** New `extract_monthly_trend()` reads the report's **"MoM"** tab (month-end date headers in row 1, line-item labels in col C) and builds a month-by-month series for **Revenue, Gross Profit, EBITDA, and Net Profit** (NGN). Parsing is label-driven (matches `total revenue` / `gross margin` / `ebitda` / `pat`) and date-driven (only real-date columns, so the `FY 25`/`FY 26` summary columns are skipped); months with no revenue are dropped so future/empty months don't show. Rendered as a Chart.js line chart in a new section placed right after the margin gauges, before Key Financial Ratios. **Limited to the current fiscal year** (the latest year with revenue data — auto-advances each year), so the current report shows Jan 26 → May 26. `main()` stashes it on `metrics['trend']`; fails safe to no chart if the MoM tab is absent.
- **"Profit After Tax" → "Net Profit"** across all display surfaces: the KPI card, the income-statement row, the profitability insight, and the ROA ratio sub-label ("Net Profit ÷ total assets"). The parser key `get('profit after tax')` (which matches the Summary tab's own label) is unchanged.

**Files**: `generate_financial_report_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-20 — Fix: Expense & Vendor "Total YTD Expenses" now sums full outflows, not just payment batches

**Why:** The "Total YTD Expenses" KPI was understated. `calculate_kpis()` computed `total_ytd` by summing only the **"BREAKDOWN OF PAYMENT BATCH"** vendor amounts (col H) — the disbursement detail — which omits all the other OUTFLOWS categories (salaries, taxes, software, bank charges, etc.). The headline therefore showed far less than the true cash that left the business.

**What** (`generate_expense_dashboard.py`):
- `calculate_kpis()` now takes `all_categories_by_week` and computes `total_ytd` as the sum of **all OUTFLOWS categories across all weeks, excluding "Investment Outflows"** (asset transfers, not operating spend — same exclusion already used for the pie-chart / takeaways denominator). This is the full operational cash outflow.
- `avg_burn` (Avg Weekly Burn Rate) now derives from this corrected total, so it reflects true weekly operating burn.
- Updated the `calculate_kpis(...)` call site to pass `all_categories_by_week`.
- Local 11-week sample now reports **Total YTD Expenses ₦953.3M** (was the vendor-batch-only subset). Note: local `data/` only carries cash reports through March 2026; the live figure tracks newer Apr–Jun reports fetched from Google Drive.

**Files**: `generate_expense_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-19 — Update: Group Financials ARR mirrors the "ARR As At" card (actual recurring % vs 50% target)

**Why:** Clarified exec intent — the Group Financials ARR should mirror the Revenue & Fundability **"ARR As At"** card (the recurring share of revenue *actually earned* YTD vs the 50% target, e.g. 39% live), not the planned/contracted $2.86M annual figure.

**What** (`generate_financial_report_dashboard.py`):
- ARR KPI card now shows `pct = Σ ytd_actual(recurring) / Σ ytd_actual(all) × 100` (identical formula to `arr_asat_pct` in `generate_revenue_dashboard`), labelled **"ARR As At <month> 2026"**, red when below the **50% target**, with `Target 50% · ▼ Npts below` and `$X recurring of $Y earned YTD`. `metrics['arr_ext']` now carries `pct`, `target_pct`, `last_month`, and the recurring/total YTD USD.
- By construction this equals whatever the Revenue & Fundability tab shows (same source + formula). Local fallback xlsx shows 19%; live Google-Sheet data shows ~39%.

**Files**: `generate_financial_report_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-19 — Update: Group Financials ARR now mirrors Revenue & Fundability; ARR ratio card removed

**Why:** Exec request — the ARR shown on Group Financials should be the *same* number as the Revenue & Fundability dashboard (one source of truth for ARR), and the duplicate "ARR % of Revenue" card in the ratios grid should go.

**What** (`generate_financial_report_dashboard.py`):
- **ARR KPI card now sourced from Revenue & Fundability.** `main()` imports `generate_revenue_dashboard.extract_revenue_data()`, finds the Path to Revenue file in the same data folder, and computes ARR = Σ annual USD (col F) of deals flagged "Recurring" (col D); NGN = ARR × `FX_RATE` (1450). Stashed on `metrics['arr_ext']`. Card now reads **$2.86M · 17 recurring of 43 streams** (was the report's own ARR line). The report's `ARR`/`ARR (%)` Summary lines remain a **fallback** if the Path to Revenue file is missing.
- **Removed the "ARR % of Revenue" card** from the Key Financial Ratios grid (now 13 cards).
- **Guarded the report-based "Recurring revenue contracted" YoY insight** so it only fires in the fallback path — avoids showing a second, conflicting ARR figure when the synced ARR is in use.

**Files**: `generate_financial_report_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-19 — Deploy: Group Financial Report 18-Jun-26; fix period parser to anchor on month token

**Why:** New weekly report dropped (`Group Financial Report 18th Jun-26.xlsx`, YTD as at 30 Jun 2026). Its natural filename exposed a bug in `_report_period_key()`: the old regex grabbed the first 3-letters+digits ("Report 18"), failed the month check, and fell back to mtime — so it would have *lost* to a properly-named May file if both were present.

**What:**
- `_report_period_key()` now anchors on a real month abbreviation (`jan…dec`) found anywhere in the name, so `18th Jun-26`, `Jun-26`, and `Jun 2026` all resolve to (2026, 6).
- Deployed the June report (force-added; `*.xlsx` gitignored) and removed the superseded `Group Financial Report_May-26_v2.xlsx` from the repo so the auto-latest picker stays unambiguous.
- Jun-26 figures: Revenue ₦3.01B ($2.19M, +YoY), gross margin 78.0%, net margin 20.1%, PAT ₦606.0M. EVA ≈ −₦734M (NOPAT ₦504.3M − ₦1.24B charge on ₦6.69B invested capital, 6mo). ROIC 15.1% (< 37% WACC).

**Action:** Reboot the live app (not just Regenerate) so the container pulls the new repo file.

**Files**: `generate_financial_report_dashboard.py`, `data/Group Financial Report 18th Jun-26.xlsx` (added), `data/Group Financial Report_May-26_v2.xlsx` (removed), `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-19 — Feature: Group Financials — add Return on Invested Capital (ROIC) ratio card

**Why:** Exec request — surface ROIC alongside the other return metrics so value creation vs cost of capital is visible at a glance.

**What** (`generate_financial_report_dashboard.py`): added a "Return on Invested Capital" card to the Key Financial Ratios grid (now 14 cards). `ROIC = annualised NOPAT ÷ invested capital`, flagged green/red against the 37% WACC (consistent with EVA sign). May-26: 14.1% (< 37% WACC → red, value-eroding YTD).

**Files**: `generate_financial_report_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-19 — Update: EVA invested capital = CA − CL + net fixed assets (was Total Assets)

**Why:** Finance's invested-capital definition is *Current Assets − Current Liabilities + Net Fixed Assets* (net working capital + net fixed assets), not gross Total Assets. The old Total-Assets proxy overstated the capital base (and the capital charge) by the amount of current liabilities.

**What** (`generate_financial_report_dashboard.py`):
- `extract_financials()` now parses `Total Current Asset`, `Total Current Liabilities`, and `Total Non Current Asset` from the Summary balance-sheet block.
- EVA capital charge now uses `inv_cap = current_assets − current_liabilities + net_fixed_assets` (≡ Total Assets − Current Liabilities). May-26: ₦6.34B (was ₦8.18B). Capital charge ₦977.5M; **EVA ≈ −₦606M YTD** (was −₦885M). ROA still uses Total Assets (by definition). EVA card text states the formula.

**Files**: `generate_financial_report_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-19 — Update: EVA NOPAT uses 30% Nigeria statutory tax; WACC 37%

**Why:** The May report books ₦0 tax YTD, so deriving the effective tax rate gave 0% — overstating NOPAT/EVA. Finance: use Nigeria's 30% statutory CIT. Also corrected WACC from the 20% placeholder to Finance's 37% board hurdle rate.

**What** (`generate_financial_report_dashboard.py`):
- `TAX_RATE = 30.0` constant. NOPAT now uses the booked effective rate **only if tax is actually booked**, else falls back to 30% (`eff_tax = eff_tax_booked if eff_tax_booked > 0 else TAX_RATE`). "Effective Tax Rate" ratio card and EVA card text state the basis ("Nigeria statutory CIT" vs "booked effective").
- `WACC_PCT = 37.0` (was 20.0).
- Net effect (May-26): NOPAT ₦371.5M (EBIT ₦530.7M × 70%); capital charge ₦1.26B (₦8.18B × 37% × 5/12); EVA ≈ −₦885M YTD.

**Files**: `generate_financial_report_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-18 — Feature: Group Financials — ARR promoted to top KPI, Key Financial Ratios + EVA added

**Why:** Exec request — ARR should sit among the headline KPI cards (not buried at the bottom), and the tab should surface the financial ratios critical to the business plus an **Economic Value Added (EVA)** measure.

**What** (`generate_financial_report_dashboard.py`):
- **ARR is now the 5th top KPI card** (USD headline, NGN + % of revenue, YoY badge). Removed the old bottom "Annual Recurring Revenue & Balance Sheet" block; that section is now a clean **Balance Sheet & Liquidity** table.
- **New "Key Financial Ratios" section** — a 13-card grid grouped by theme: Profitability (gross/EBITDA/operating/net margin), Returns (annualised ROA), Efficiency (OpEx-, payroll-, marketing-to-revenue, effective tax rate), Liquidity (interest coverage, current ratio, cash ratio), Recurring (ARR % of revenue). Each card shows current value, prior/target, and is green/red/neutral vs its benchmark.
- **New Economic Value Added (EVA) card** — `EVA = NOPAT − (Invested Capital × WACC)`. NOPAT = EBIT × (1 − effective tax rate); invested capital proxied by Total Assets; capital charge prorated to the YTD period. `WACC_PCT = 37.0` — Finance's board-approved hurdle rate. Added a matching EVA insight card. May-26 EVA ≈ −₦730M YTD (NOPAT ₦530.7M < ₦1.26B annualised capital charge on ₦8.18B assets) — returns not yet above cost of capital.

**Note:** EVA depends on the `WACC_PCT` rate (37%) and a Total-Assets capital proxy (the Summary tab has no clean debt+equity line) — both surfaced in the card text. Adjust `WACC_PCT` if the board hurdle rate changes.

**Files**: `generate_financial_report_dashboard.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

---

## 2026-06-17 — Fix: Group Financial report always refreshed in working dir (was serving stale month)

**Why:** After committing the May report, the live tab still showed April. Root cause: `prepare_data_folder()` copies each `*.xlsx` only `if not dest.exists()`, and Streamlit Cloud's `generated/data_working/` persists within a running container. "Regenerate Dashboards" re-runs generators but does NOT re-copy already-present files or prune deleted ones, so the container kept serving the previously-copied report.

**What:** `prepare_data_folder()` now **always overwrites** `Group Financial Report*.xlsx` into the working dir each run (like the budget snapshot) and **prunes** any report no longer in the repo. Combined with `find_file()`'s latest-period picker, the working dir can never serve a stale or retired month.

**Note:** this is a code change → the live app must be **Rebooted** (share.streamlit.io → ⋮ → Reboot), not just "Regenerate Dashboards", to load it and pull the newly-committed report.

**Files**: `app.py`, `CLAUDE.md`, `CHANGELOG.md`
**Author**: Lilian Wilfred + Claude

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
