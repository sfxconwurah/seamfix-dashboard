# Changelog — Seamfix Financial Dashboard

> **Rule**: Every code change must add an entry here. Format: date, type, description, files changed.  
> **Types**: `Fix`, `Feature`, `Update`, `Refactor`, `Docs`

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
**Reported by**: Lilian Wilfred (Finance)

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
