# Changelog — Seamfix Financial Dashboard

> **Rule**: Every code change must add an entry here. Format: date, type, description, files changed.  
> **Types**: `Fix`, `Feature`, `Update`, `Refactor`, `Docs`

---

## 2026-06-02 — Feature: Collections Intelligence dashboard (6th tab)

Added a new **Collections Tracker** dashboard tab so the executive team can track critical revenue inflows in one place, see what moved week-over-week, and act on stalled or at-risk collections.

**What it shows:**
- **KPI row**: Total tracked ($2.67M / ₦3.88B), fully collected, in-progress, and payment-pending amounts with item counts and % of portfolio.
- **Critical Actions panel**: Automatically surfaces items needing urgent action (high-value pending items, "at risk" flags, overdue deadlines, specific deliverables assigned), sorted by USD value. Colour-coded red (critical) and amber (follow-up needed).
- **What Moved This Week**: Compares the last two weekly update columns for every item and highlights deals where something changed — green for positive movement (payment received, deal signed), red for negative flags (payment at risk, extension granted, no response), blue for general updates.
- **Full tracker table**: All 20 items with payment status, deal status, predictability, accountable party, latest update, and actionable deliverable. Filterable by status.

**Data source**: Google Sheet `17KE1n5_SOeDXaX96Xsa1JfAjNs_OZX8xu-wYDt4LpU8` — fetched live each page load (5-min cache, same pattern as the revenue sheet). The sheet must be shared with the service account or set to "anyone with the link can view".

**How it works**: `app.py` fetches the sheet as xlsx and saves it to the data_working folder as `2026 Collections Tracker.xlsx`. The generator (`generate_collections_dashboard.py`) scans all tabs for the one containing the S/N header, dynamically maps columns by name (not position), detects date-labelled weekly-update columns, runs movement/urgency analysis, and outputs a self-contained HTML dashboard.

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
