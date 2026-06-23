"""
Glossary & Definitions generator for the Seamfix Financial Intelligence Suite.

This is the SINGLE SOURCE of the glossary content (the GLOSSARY list below). It emits
two artifacts from that one source so the in-app tab and the committed doc never drift:

  1. glossary_dashboard.html  — themed page embedded as the "Glossary & Definitions"
     dashboard tab (run as `python3 generate_glossary_dashboard.py <data_folder>`).
  2. GLOSSARY.md              — versioned reference doc in the repo
     (run `python3 generate_glossary_dashboard.py --markdown` to (re)write it).

To add or change a term: edit the GLOSSARY list, then regenerate BOTH
(`python3 generate_glossary_dashboard.py --markdown` for the doc; the app rebuilds the
HTML on load). Keep definitions in sync with the generators' actual logic and CLAUDE.md.
"""

import os, sys, html as _html
from datetime import datetime
from theme import get_base_css, get_toggle_html, get_theme_js

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

INTRO = (
    "A plain-language reference for every metric on the Seamfix Financial Intelligence "
    "Suite, with how each one applies to Seamfix. Read it alongside the dashboards each "
    "week so the numbers are interpreted consistently. Definitions track the live "
    "calculations; if a generator's logic changes, update this glossary too."
)

# Key constants stated once so the definitions stay consistent with the generators.
FX_RATE = 1450
ANNUAL_TARGET_USD = 8_000_000

# ── SINGLE SOURCE OF TRUTH ────────────────────────────────────────────────────
# Each section: {category, icon, intro, terms:[{term, definition, seamfix}]}
GLOSSARY = [
    {
        "category": "General & Cross-Cutting",
        "icon": "\U0001F310",  # globe
        "intro": "Conventions that apply across all tabs.",
        "terms": [
            {"term": "NGN (\u20A6) / USD ($)",
             "definition": "Nigerian Naira and US Dollars. Dashboards show one or both depending on the metric.",
             "seamfix": "Seamfix earns and spends in both currencies; cash, budget and most expenses are Naira, while the revenue pipeline and collections are tracked in USD."},
            {"term": "FX Rate",
             "definition": "The exchange rate used to convert between Naira and Dollars for dual-currency display.",
             "seamfix": f"A flat $1 = \u20A6{FX_RATE:,} is hardcoded across the dashboards for comparability. The Cash and Group Financials tabs additionally use each report's own period rate where the source provides one. Update the constant if the rate moves materially."},
            {"term": "YTD (Year-to-Date)",
             "definition": "The cumulative figure from the start of the fiscal year up to the latest data point.",
             "seamfix": "Detected dynamically \u2014 the dashboards scan which months/weeks actually have data, so YTD advances automatically as Finance adds a new month or weekly report."},
            {"term": "FY (Fiscal Year)",
             "definition": "The 12-month accounting period used for budgeting and reporting.",
             "seamfix": "Seamfix's fiscal year runs January\u2013December. The current FY is 2026."},
            {"term": "Pace",
             "definition": "The portion of an annual figure that 'should' have happened by now if activity were spread evenly across the year (annual \u00D7 elapsed period \u00F7 full period).",
             "seamfix": "Used to judge whether spend or revenue is ahead of or behind schedule \u2014 e.g. a budget at 4 of 12 months should be ~33% spent."},
            {"term": "Run Rate",
             "definition": "A projection of the full-year total by extrapolating the average achieved so far.",
             "seamfix": "Used to forecast year-end revenue and spend from partial-year actuals; the pipeline run rate scales the latest partial month up to a full month before averaging."},
            {"term": "Annualised",
             "definition": "A part-period figure scaled up to a 12-month equivalent so it can be compared to annual targets or ratios.",
             "seamfix": "Profitability returns like ROA and ROIC on the Group Financials tab are annualised because the P&L is only a few months of the year."},
            {"term": "YoY (Year-over-Year)",
             "definition": "Comparison of a figure to the same period one year earlier.",
             "seamfix": "On Group Financials, YoY uses magnitude growth and flags a swing between profit and loss as a 'turnaround' rather than a misleading negative percentage."},
        ],
    },
    {
        "category": "Cash Overview",
        "icon": "\U0001F4B0",  # money bag
        "intro": "Weekly cash position from the Finance team's cash reports.",
        "terms": [
            {"term": "Total Position (incl. Investments)",
             "definition": "The grand total of all cash and cash-equivalent holdings, expressed in Naira.",
             "seamfix": "Sums Naira cash, Dollar cash (converted to Naira), Naira investments and Dollar investments (converted). It excludes the MTN share holding, which is listed equity, not cash."},
            {"term": "NGN Balance (incl. Investments)",
             "definition": "Closing Naira cash plus Naira investments.",
             "seamfix": "The Naira slice of the Total Position."},
            {"term": "USD Balance (incl. Investments)",
             "definition": "Closing Dollar cash plus Dollar investments, shown in actual dollars (not Naira-equivalent).",
             "seamfix": "Closing USD cash is derived from the report's USD closing Naira value \u00F7 its own FX rate (not the opening balance). USD investments come from the authoritative Cash Balance Summary sheet (AIF Investment Fund + Other Dollar Investment)."},
            {"term": "MTN Shares (Market Value)",
             "definition": "The current market value of the company's MTN equity holding.",
             "seamfix": "Shown as its own card and deliberately excluded from the cash Total Position \u2014 it is listed equity that can fluctuate, not spendable cash."},
            {"term": "Investments (AIF, Other Dollar, Naira)",
             "definition": "Funds placed in interest-bearing or managed instruments rather than held as operating cash.",
             "seamfix": "The AIF Investment Fund and Other Dollar Investment (USD) plus Naira fixed deposits. Counted in the cash position because they are liquid, but distinguished from operating cash."},
            {"term": "Weekly Inflow / Outflow",
             "definition": "Total money received / paid out during the week, by category.",
             "seamfix": "Operating flows only; investment movements (placing or liquidating funds) are separated out so they don't distort the operating picture."},
            {"term": "Net Cash Flow",
             "definition": "Weekly inflow minus weekly outflow \u2014 whether cash grew or shrank that week.",
             "seamfix": "A negative net flow week means the company drew down cash to operate."},
            {"term": "Operational Runway",
             "definition": "How long current cash would last at the recent average operating burn rate.",
             "seamfix": "Total Position (including investments) \u00F7 average operational burn. The burn rate excludes investment outflows, but the runway numerator still includes investments \u2014 so it is measured 'at the operational burn rate'."},
            {"term": "4-Week Forecast",
             "definition": "A short-term projection of the cash position based on recent net flows.",
             "seamfix": "An early-warning indicator for upcoming cash pressure."},
        ],
    },
    {
        "category": "Expense & Vendor Analysis",
        "icon": "\U0001F4CA",  # bar chart
        "intro": "Operating spend and vendor concentration from the weekly cash reports.",
        "terms": [
            {"term": "Total YTD Expenses",
             "definition": "All operating cash outflows for the year to date.",
             "seamfix": "Summed from the full OUTFLOWS section of every weekly report (salaries, taxes, software, bank charges, vendor payments, etc.), EXCLUDING investment outflows. It is NOT just the payment-batch vendor list, which would undercount the true total."},
            {"term": "Average Weekly Burn Rate",
             "definition": "Total YTD expenses divided by the number of weekly reports.",
             "seamfix": "The typical weekly operating cash consumption \u2014 the denominator behind the cash runway."},
            {"term": "Investment Outflows (excluded)",
             "definition": "Cash moved into investments or funding, not an operating expense.",
             "seamfix": "Asset transfers, not spend \u2014 excluded from Total YTD Expenses and the spend mix so they don't inflate the burn rate."},
            {"term": "Vendor / Payment Batch",
             "definition": "The disbursement-level detail of who was paid.",
             "seamfix": "A subset of total outflows (it omits salaries, taxes, etc.). Drives the vendor ledger, concentration and largest-payment metrics, but not the headline expense total."},
            {"term": "Recurring Vendor",
             "definition": "A vendor that appears in three or more weekly reports.",
             "seamfix": "Signals ongoing commitments worth reviewing for contract terms and consolidation."},
            {"term": "Vendor Concentration",
             "definition": "How much of total spend goes to the top few vendors.",
             "seamfix": "High concentration is a dependency/negotiation-leverage flag."},
        ],
    },
    {
        "category": "Budget vs Actual",
        "icon": "\U0001F4CB",  # clipboard
        "intro": "Group/entity/department spend against the approved budget (Budget Tracker snapshot, lean mode, all NGN).",
        "terms": [
            {"term": "Annual Budget",
             "definition": "The approved full-year spending plan.",
             "seamfix": "Built bottom-up so Group = sum of entities = sum of departments (~\u20A65.10B FY2026). UK budgets are in GBP and UAE in USD, FX-converted to Naira; actuals are already Naira."},
            {"term": "YTD Budget (Pace to Date)",
             "definition": "The portion of the annual budget allocated to the months elapsed so far.",
             "seamfix": "The benchmark a unit's YTD actual is compared against."},
            {"term": "YTD Actual",
             "definition": "Actual spend recorded for the year to date.",
             "seamfix": "Loaded from Acumatica into the Budget Tracker (lean mode). Actuals are Naira for every entity."},
            {"term": "Variance",
             "definition": "YTD budget pace minus YTD actual \u2014 the over- or under-spend.",
             "seamfix": "Positive = under pace (spending less than planned); negative = over pace."},
            {"term": "% of Pace",
             "definition": "YTD actual \u00F7 YTD budget pace, as a percentage.",
             "seamfix": "Over 105% = Over Budget (red); under 70% = Under Budget (green/under-spending); in between = On Track."},
            {"term": "% of Annual",
             "definition": "YTD actual \u00F7 full annual budget.",
             "seamfix": "How much of the whole-year envelope has been used."},
            {"term": "Year-End Projection",
             "definition": "Run-rate forecast of full-year spend from YTD actuals.",
             "seamfix": "Compared to the annual budget to flag likely overspend or savings."},
            {"term": "Entity / Department / Budget Head",
             "definition": "The three drill-down levels: legal entity (NG/UK/UAE), department, and individual budget line.",
             "seamfix": "Eleven departments (e.g. Commercial, Solutions, Finance, CAPEX, UK, UAE). Each department expands to its budget heads."},
            {"term": "Lean vs Full mode",
             "definition": "Two budget views in the tracker: lean (Acumatica-loaded, carries actuals) and full budget.",
             "seamfix": "The dashboard uses the lean budget to drive a budget-vs-actual comparison."},
        ],
    },
    {
        "category": "Revenue & Fundability",
        "icon": "\U0001F680",  # rocket
        "intro": "Revenue performance and recurring-revenue quality against the company target.",
        "terms": [
            {"term": "Annual Revenue Target",
             "definition": "The official company revenue goal for the year.",
             "seamfix": f"${ANNUAL_TARGET_USD:,.0f} ($8M) is the official target used for gap and projection maths, even though the deal bucket sums to a more optimistic ~$10M."},
            {"term": "ARR (Annual Recurring Revenue)",
             "definition": "The annualised value of revenue that recurs (subscriptions/contracted recurring deals).",
             "seamfix": "Sum of annual revenue for deals flagged 'Recurring' (~$2.86M across recurring deals in 2026). A measure of predictable, fundable revenue."},
            {"term": "ARR As At <month>",
             "definition": "The recurring share of revenue actually EARNED year-to-date, versus target.",
             "seamfix": "Recurring YTD \u00F7 total YTD revenue, measured against a 50% target (red when below). This is the actuals-based recurring quality, not the contracted plan figure. The Group Financials tab mirrors this exact number."},
            {"term": "Achievement %",
             "definition": "Progress of a deal or the portfolio toward its full annual target.",
             "seamfix": "YTD actual \u00F7 annual target (NOT pace-adjusted) \u2014 a deal that earned its full annual target shows 100%."},
            {"term": "Recurring vs Not Recurring",
             "definition": "Whether a deal's revenue repeats or is one-off.",
             "seamfix": "Tagged in column D of the revenue sheet; drives the ARR calculations."},
            {"term": "Deficit / Surplus / Gap",
             "definition": "How actual revenue compares to the planned amount for a deal.",
             "seamfix": "Deficit = behind plan, Surplus = ahead, Gap = the outstanding shortfall to target."},
        ],
    },
    {
        "category": "Pipeline Intelligence",
        "icon": "\U0001F3AF",  # target
        "intro": "Deal-level momentum and a probability-weighted view of landing the target.",
        "terms": [
            {"term": "Deal Status (On Track / At Risk / Off Track / Closed)",
             "definition": "Finance's assessment of each deal's health.",
             "seamfix": "Drives the weighted projection. 'Closed' means won/booked."},
            {"term": "Status Weights",
             "definition": "Probability factors applied to each deal's value to build a realistic projection.",
             "seamfix": "On Track 100%, Closed 100%, At Risk 50%, Off Track 10%, Unknown 70%."},
            {"term": "Realistic vs Conservative Projection",
             "definition": "Two weighted forecasts of full-year revenue.",
             "seamfix": "Realistic applies the status weights above; Conservative excludes At Risk and Off Track entirely \u2014 a downside scenario."},
            {"term": "Gap to Target",
             "definition": "Target minus projected revenue.",
             "seamfix": "Positive = shortfall to close; zero or negative = target is projected to be met."},
            {"term": "Momentum (growing / stalled / new / steady)",
             "definition": "Trend signal comparing a deal's two most recent months.",
             "seamfix": "Growing = latest > 1.1\u00D7 prior; stalled = activity dropped to zero; new = only the latest month has revenue; steady = otherwise. Surfaces deals losing traction."},
        ],
    },
    {
        "category": "Collections Tracker",
        "icon": "\U0001F4E5",  # inbox tray
        "intro": "Critical revenue inflows and the status of collecting cash already won.",
        "terms": [
            {"term": "Booked",
             "definition": "Whether the revenue has been formally recognised/contracted.",
             "seamfix": "'YES' means it counts as committed revenue; tracked against total expected."},
            {"term": "Predictability",
             "definition": "Confidence that the expected cash will actually land.",
             "seamfix": "Helps prioritise follow-up on lower-confidence inflows."},
            {"term": "Closure Period",
             "definition": "When the deal/payment is expected to close.",
             "seamfix": "Used with predictability to classify collection urgency."},
            {"term": "Payment Status (Pending / In-Progress / Closed)",
             "definition": "Where the inflow sits in the collection cycle.",
             "seamfix": "'Closed' = collected; pending/in-progress are the focus of weekly collection actions."},
            {"term": "Deal Status",
             "definition": "The commercial state of the underlying deal.",
             "seamfix": "Distinct from payment status \u2014 a closed deal can still have pending payment."},
        ],
    },
    {
        "category": "Group Financials (Consolidated P&L)",
        "icon": "\U0001F3E6",  # bank
        "intro": "The audited-style consolidated income statement, ratios and value metrics.",
        "terms": [
            {"term": "Total Revenue",
             "definition": "All income earned from operations in the period.",
             "seamfix": "Reported for the GROUP block; shown in Naira with the report's own USD equivalent."},
            {"term": "Cost of Sales (COGS)",
             "definition": "Direct costs of delivering the product/service.",
             "seamfix": "Subtracted from revenue to get gross profit."},
            {"term": "Gross Profit / Gross Margin",
             "definition": "Revenue minus cost of sales; margin is that as a % of revenue.",
             "seamfix": "Benchmarked against a 70% gross-margin target on a gauge card."},
            {"term": "Operating Expenses (OpEx)",
             "definition": "Running costs not tied directly to delivery (payroll, marketing, admin, etc.).",
             "seamfix": "Broken down with payroll % and marketing % of revenue ratios."},
            {"term": "EBITDA / EBITDA Margin",
             "definition": "Earnings Before Interest, Tax, Depreciation & Amortisation \u2014 operating profitability before financing and accounting charges.",
             "seamfix": "A proxy for core operating cash generation."},
            {"term": "EBIT",
             "definition": "Earnings Before Interest and Tax (EBITDA less depreciation & amortisation).",
             "seamfix": "The operating profit used as the base for NOPAT in the EVA calculation."},
            {"term": "Net Profit (Profit After Tax) / Net Margin",
             "definition": "The bottom-line profit after all costs and tax; margin is that as a % of revenue.",
             "seamfix": "Displayed as 'Net Profit' everywhere (the source label is 'Profit After Tax'). Benchmarked against a 10% net-margin target."},
            {"term": "ROA (Return on Assets)",
             "definition": "Annualised net profit as a % of total assets \u2014 how efficiently assets generate profit.",
             "seamfix": "One of the profitability ratio cards; annualised because the P&L is a partial year."},
            {"term": "ROIC (Return on Invested Capital)",
             "definition": "Annualised NOPAT \u00F7 invested capital \u2014 return on the capital actually put to work.",
             "seamfix": "Flagged red when below WACC (the company isn't yet earning its cost of capital)."},
            {"term": "WACC (Cost of Capital)",
             "definition": "The blended return investors/lenders require \u2014 the hurdle rate.",
             "seamfix": "Set to the pre-approved 37%. Returns below this destroy value; used in EVA and the ROIC red flag."},
            {"term": "NOPAT",
             "definition": "Net Operating Profit After Tax = EBIT \u00D7 (1 \u2212 effective tax rate).",
             "seamfix": "Because the report books \u20A60 tax YTD, a 30% Nigeria statutory rate is used as a fallback so NOPAT isn't overstated."},
            {"term": "Invested Capital",
             "definition": "The capital tied up in the business: current assets \u2212 current liabilities + net fixed assets.",
             "seamfix": "Parsed from the balance-sheet rows; the base for ROIC and the EVA capital charge."},
            {"term": "EVA (Economic Value Added)",
             "definition": "NOPAT minus a capital charge (invested capital \u00D7 WACC) \u2014 value created above the cost of capital.",
             "seamfix": "The capital charge is prorated to the YTD period so it compares fairly with period NOPAT. Negative EVA means returns haven't yet cleared the 37% hurdle."},
            {"term": "Effective Tax Rate",
             "definition": "Tax expense as a % of pre-tax profit.",
             "seamfix": "Currently ~0% booked YTD; the 30% statutory rate is substituted for NOPAT/EVA until real tax appears."},
            {"term": "Interest Coverage",
             "definition": "How many times operating earnings cover interest expense.",
             "seamfix": "A solvency/financing-risk indicator on the ratios grid."},
            {"term": "Current Ratio / Cash Ratio",
             "definition": "Liquidity ratios: current assets (or cash) \u00F7 current liabilities.",
             "seamfix": "Above 1.0 indicates short-term obligations are comfortably covered."},
            {"term": "Revenue by Vertical / Customer / Country",
             "definition": "Breakdowns of where revenue comes from.",
             "seamfix": "Highlight concentration (e.g. a single large customer/segment). These breakdowns exclude Other Income."},
            {"term": "Monthly Performance Trend",
             "definition": "Month-by-month chart of Revenue, Gross Profit, EBITDA and Net Profit.",
             "seamfix": "Read from the report's MoM tab and limited to the current fiscal year (auto-advances each year)."},
        ],
    },
]


def esc(s):
    return _html.escape(str(s))


# ── HTML output ────────────────────────────────────────────────────────────
def build_html():
    now = datetime.now().strftime('%d %b %Y')
    theme_css = get_base_css()
    toggle_html = get_toggle_html()
    theme_js = get_theme_js()

    total_terms = sum(len(s["terms"]) for s in GLOSSARY)

    css = """
""" + theme_css + """
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: var(--bg-nav);
    color: var(--text-primary);
    padding: 28px 28px 80px;
    font-size: 14px;
    line-height: 1.55;
}
.page-header { margin-bottom: 6px; }
.page-title  { font-size: 26px; font-weight: 700; color: var(--accent); }
.page-sub    { font-size: 13px; color: var(--text-tertiary); margin-top: 3px; }
.intro {
    background: var(--bg-card); border: 1px solid var(--border-main); border-left: 4px solid var(--accent);
    border-radius: 10px; padding: 16px 18px; margin: 20px 0 18px; color: var(--text-secondary);
    font-size: 13px; max-width: 920px;
}
.search-wrap { margin: 0 0 22px; }
#glossarySearch {
    width: 100%; max-width: 520px; padding: 11px 14px; font-size: 14px; font-family: inherit;
    background: var(--bg-input); color: var(--text-primary);
    border: 1px solid var(--border-main); border-radius: 9px; outline: none;
}
#glossarySearch:focus { border-color: var(--accent); }
.no-results { color: var(--text-tertiary); font-size: 13px; font-style: italic; display: none; margin-top: 10px; }

/* Category nav */
.cat-nav { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 26px; }
.cat-chip {
    display: inline-block; padding: 6px 13px; border-radius: 20px; font-size: 12px; font-weight: 600;
    background: var(--accent-bg); color: var(--accent); border: 1px solid var(--border-accent);
    text-decoration: none; cursor: pointer; transition: all .15s;
}
.cat-chip:hover { background: var(--accent); color: var(--text-on-accent); }

/* Section headers */
.sec-hdr {
    display: flex; align-items: baseline; gap: 12px;
    margin: 34px 0 6px; padding-bottom: 12px; border-bottom: 1px solid var(--border-main);
}
.sec-title { font-size: 18px; font-weight: 700; color: var(--text-primary); }
.sec-count { font-size: 12px; color: var(--text-tertiary); }
.sec-intro { font-size: 12px; color: var(--text-tertiary); margin: 8px 0 16px; }

/* Term cards */
.term-grid {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(420px, 1fr));
    gap: 14px; margin-bottom: 8px;
}
.term-card {
    background: var(--bg-card); border: 1px solid var(--border-main); border-radius: 10px;
    padding: 16px 18px;
}
.term-name { font-size: 14px; font-weight: 700; color: var(--accent); margin-bottom: 8px; }
.term-def  { font-size: 13px; color: var(--text-primary); margin-bottom: 10px; }
.term-sfx-lbl {
    font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .6px;
    color: var(--text-tertiary); margin-bottom: 3px;
}
.term-sfx {
    font-size: 12.5px; color: var(--text-secondary);
    background: var(--success-bg); border-radius: 6px; padding: 9px 11px;
}
@media print {
    #themeToggle, .cat-nav, .search-wrap { display: none !important; }
    .term-grid { grid-template-columns: 1fr; }
    body { padding: 12px; }
}
"""

    def anchor(cat):
        return "cat-" + "".join(ch.lower() if ch.isalnum() else "-" for ch in cat)

    # Category nav chips
    nav = "".join(
        '<a class="cat-chip" href="#' + anchor(s["category"]) + '">'
        + esc(s["icon"]) + " " + esc(s["category"]) + "</a>"
        for s in GLOSSARY
    )

    # Sections
    sections = []
    for s in GLOSSARY:
        cards = []
        for t in s["terms"]:
            search_key = (t["term"] + " " + t["definition"] + " " + t["seamfix"]).lower()
            cards.append(
                '<div class="term-card" data-search="' + esc(search_key) + '">'
                + '<div class="term-name">' + esc(t["term"]) + '</div>'
                + '<div class="term-def">' + esc(t["definition"]) + '</div>'
                + '<div class="term-sfx-lbl">At Seamfix</div>'
                + '<div class="term-sfx">' + esc(t["seamfix"]) + '</div>'
                + '</div>'
            )
        sections.append(
            '<div class="glossary-section" id="' + anchor(s["category"]) + '">'
            + '<div class="sec-hdr"><span style="font-size:20px">' + esc(s["icon"]) + '</span>'
            + '<span class="sec-title">' + esc(s["category"]) + '</span>'
            + '<span class="sec-count">' + str(len(s["terms"])) + ' terms</span></div>'
            + '<div class="sec-intro">' + esc(s["intro"]) + '</div>'
            + '<div class="term-grid">' + "".join(cards) + '</div>'
            + '</div>'
        )

    search_js = """
function filterGlossary() {
    var q = (document.getElementById('glossarySearch').value || '').toLowerCase().trim();
    var anyVisible = false;
    document.querySelectorAll('.glossary-section').forEach(function(sec) {
        var secVisible = false;
        sec.querySelectorAll('.term-card').forEach(function(card) {
            var match = !q || (card.getAttribute('data-search') || '').indexOf(q) !== -1;
            card.style.display = match ? '' : 'none';
            if (match) secVisible = true;
        });
        sec.style.display = secVisible ? '' : 'none';
        if (secVisible) anyVisible = true;
    });
    document.getElementById('noResults').style.display = anyVisible ? 'none' : 'block';
}
"""

    html_doc = (
        "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n"
        "<meta charset=\"UTF-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n"
        "<title>Glossary &amp; Definitions &mdash; Seamfix 2026</title>\n"
        "<style>" + css + "</style>\n</head>\n<body>\n"
        + toggle_html + "\n"
        + '<div class="page-header">'
        + '<div class="page-title">\U0001F4D6 Glossary &amp; Definitions</div>'
        + '<div class="page-sub">How to read the Seamfix Financial Intelligence Suite &nbsp;\u00B7&nbsp; '
        + str(total_terms) + ' terms across ' + str(len(GLOSSARY)) + ' areas &nbsp;\u00B7&nbsp; Updated ' + now + '</div>'
        + '</div>'
        + '<div class="intro">' + esc(INTRO) + '</div>'
        + '<div class="search-wrap"><input id="glossarySearch" type="text" placeholder="Search terms, definitions or Seamfix notes\u2026" oninput="filterGlossary()" autocomplete="off">'
        + '<div class="no-results" id="noResults">No matching terms.</div></div>'
        + '<div class="cat-nav">' + nav + '</div>'
        + "".join(sections)
        + "\n<script>\n" + search_js + "\n" + theme_js + "\n</script>\n"
        + "</body>\n</html>"
    )
    return html_doc


# ── Markdown output ──────────────────────────────────────────────────────────
def build_markdown():
    now = datetime.now().strftime('%d %b %Y')
    lines = []
    lines.append("# Seamfix Financial Intelligence Suite \u2014 Glossary & Definitions")
    lines.append("")
    lines.append("> **Generated from `generate_glossary_dashboard.py` (the `GLOSSARY` list).** "
                 "Do not hand-edit this file \u2014 edit the generator and re-run "
                 "`python3 generate_glossary_dashboard.py --markdown`. The same source also "
                 "renders the in-app **Glossary & Definitions** tab.")
    lines.append("")
    lines.append("_" + INTRO + "_")
    lines.append("")
    lines.append("_Last updated: " + now + "._")
    lines.append("")
    lines.append("---")
    lines.append("")
    # Table of contents
    lines.append("## Contents")
    lines.append("")
    for s in GLOSSARY:
        anchor = s["category"].lower().replace(" ", "-")
        for ch in "&()/,.":
            anchor = anchor.replace(ch, "")
        lines.append("- [" + s["category"] + "](#" + anchor + ")")
    lines.append("")
    lines.append("---")
    lines.append("")
    for s in GLOSSARY:
        lines.append("## " + s["category"])
        lines.append("")
        lines.append("_" + s["intro"] + "_")
        lines.append("")
        for t in s["terms"]:
            lines.append("### " + t["term"])
            lines.append("")
            lines.append("**Definition:** " + t["definition"])
            lines.append("")
            lines.append("**At Seamfix:** " + t["seamfix"])
            lines.append("")
        lines.append("---")
        lines.append("")
    return "\n".join(lines)


def main():
    args = sys.argv[1:]
    if "--markdown" in args or "--md" in args:
        out = os.path.join(SCRIPT_DIR, "GLOSSARY.md")
        with open(out, "w", encoding="utf-8") as f:
            f.write(build_markdown())
        print("Wrote " + out)
        return

    folder = args[0] if args else "."
    out = os.path.join(folder, "glossary_dashboard.html")
    try:
        with open(out, "w", encoding="utf-8") as f:
            f.write(build_html())
        print("Glossary dashboard generated: " + out)
    except Exception as e:
        import traceback
        print("Error generating glossary dashboard: " + str(e), file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
