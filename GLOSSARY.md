# Seamfix Financial Intelligence Suite — Glossary & Definitions

> **Generated from `generate_glossary_dashboard.py` (the `GLOSSARY` list).** Do not hand-edit this file — edit the generator and re-run `python3 generate_glossary_dashboard.py --markdown`. The same source also renders the in-app **Glossary & Definitions** tab.

_A plain-language reference for every metric on the Seamfix Financial Intelligence Suite, with how each one applies to Seamfix. Read it alongside the dashboards each week so the numbers are interpreted consistently. Definitions track the live calculations; if a generator's logic changes, update this glossary too._

_Last updated: 23 Jun 2026._

---

## Contents

- [General & Cross-Cutting](#general--cross-cutting)
- [Cash Overview](#cash-overview)
- [Expense & Vendor Analysis](#expense--vendor-analysis)
- [Budget vs Actual](#budget-vs-actual)
- [Revenue & Fundability](#revenue--fundability)
- [Pipeline Intelligence](#pipeline-intelligence)
- [Collections Tracker](#collections-tracker)
- [Group Financials (Consolidated P&L)](#group-financials-consolidated-pl)

---

## General & Cross-Cutting

_Conventions that apply across all tabs._

### NGN (₦) / USD ($)

**Definition:** Nigerian Naira and US Dollars. Dashboards show one or both depending on the metric.

**At Seamfix:** Seamfix earns and spends in both currencies; cash, budget and most expenses are Naira, while the revenue pipeline and collections are tracked in USD.

### FX Rate

**Definition:** The exchange rate used to convert between Naira and Dollars for dual-currency display.

**At Seamfix:** A flat $1 = ₦1,450 is hardcoded across the dashboards for comparability. The Cash and Group Financials tabs additionally use each report's own period rate where the source provides one. Update the constant if the rate moves materially.

### YTD (Year-to-Date)

**Definition:** The cumulative figure from the start of the fiscal year up to the latest data point.

**At Seamfix:** Detected dynamically — the dashboards scan which months/weeks actually have data, so YTD advances automatically as Finance adds a new month or weekly report.

### FY (Fiscal Year)

**Definition:** The 12-month accounting period used for budgeting and reporting.

**At Seamfix:** Seamfix's fiscal year runs January–December. The current FY is 2026.

### Pace

**Definition:** The portion of an annual figure that 'should' have happened by now if activity were spread evenly across the year (annual × elapsed period ÷ full period).

**At Seamfix:** Used to judge whether spend or revenue is ahead of or behind schedule — e.g. a budget at 4 of 12 months should be ~33% spent.

### Run Rate

**Definition:** A projection of the full-year total by extrapolating the average achieved so far.

**At Seamfix:** Used to forecast year-end revenue and spend from partial-year actuals; the pipeline run rate scales the latest partial month up to a full month before averaging.

### Annualised

**Definition:** A part-period figure scaled up to a 12-month equivalent so it can be compared to annual targets or ratios.

**At Seamfix:** Profitability returns like ROA and ROIC on the Group Financials tab are annualised because the P&L is only a few months of the year.

### YoY (Year-over-Year)

**Definition:** Comparison of a figure to the same period one year earlier.

**At Seamfix:** On Group Financials, YoY uses magnitude growth and flags a swing between profit and loss as a 'turnaround' rather than a misleading negative percentage.

---

## Cash Overview

_Weekly cash position from the Finance team's cash reports._

### Total Position (incl. Investments)

**Definition:** The grand total of all cash and cash-equivalent holdings, expressed in Naira.

**At Seamfix:** Sums Naira cash, Dollar cash (converted to Naira), Naira investments and Dollar investments (converted). It excludes the MTN share holding, which is listed equity, not cash.

### NGN Balance (incl. Investments)

**Definition:** Closing Naira cash plus Naira investments.

**At Seamfix:** The Naira slice of the Total Position.

### USD Balance (incl. Investments)

**Definition:** Closing Dollar cash plus Dollar investments, shown in actual dollars (not Naira-equivalent).

**At Seamfix:** Closing USD cash is derived from the report's USD closing Naira value ÷ its own FX rate (not the opening balance). USD investments come from the authoritative Cash Balance Summary sheet (AIF Investment Fund + Other Dollar Investment).

### MTN Shares (Market Value)

**Definition:** The current market value of the company's MTN equity holding.

**At Seamfix:** Shown as its own card and deliberately excluded from the cash Total Position — it is listed equity that can fluctuate, not spendable cash.

### Investments (AIF, Other Dollar, Naira)

**Definition:** Funds placed in interest-bearing or managed instruments rather than held as operating cash.

**At Seamfix:** The AIF Investment Fund and Other Dollar Investment (USD) plus Naira fixed deposits. Counted in the cash position because they are liquid, but distinguished from operating cash.

### Weekly Inflow / Outflow

**Definition:** Total money received / paid out during the week, by category.

**At Seamfix:** Operating flows only; investment movements (placing or liquidating funds) are separated out so they don't distort the operating picture.

### Net Cash Flow

**Definition:** Weekly inflow minus weekly outflow — whether cash grew or shrank that week.

**At Seamfix:** A negative net flow week means the company drew down cash to operate.

### Operational Runway

**Definition:** How long current cash would last at the recent average operating burn rate.

**At Seamfix:** Total Position (including investments) ÷ average operational burn. The burn rate excludes investment outflows, but the runway numerator still includes investments — so it is measured 'at the operational burn rate'.

### 4-Week Forecast

**Definition:** A short-term projection of the cash position based on recent net flows.

**At Seamfix:** An early-warning indicator for upcoming cash pressure.

---

## Expense & Vendor Analysis

_Operating spend and vendor concentration from the weekly cash reports._

### Total YTD Expenses

**Definition:** All operating cash outflows for the year to date.

**At Seamfix:** Summed from the full OUTFLOWS section of every weekly report (salaries, taxes, software, bank charges, vendor payments, etc.), EXCLUDING investment outflows. It is NOT just the payment-batch vendor list, which would undercount the true total.

### Average Weekly Burn Rate

**Definition:** Total YTD expenses divided by the number of weekly reports.

**At Seamfix:** The typical weekly operating cash consumption — the denominator behind the cash runway.

### Investment Outflows (excluded)

**Definition:** Cash moved into investments or funding, not an operating expense.

**At Seamfix:** Asset transfers, not spend — excluded from Total YTD Expenses and the spend mix so they don't inflate the burn rate.

### Vendor / Payment Batch

**Definition:** The disbursement-level detail of who was paid.

**At Seamfix:** A subset of total outflows (it omits salaries, taxes, etc.). Drives the vendor ledger, concentration and largest-payment metrics, but not the headline expense total.

### Recurring Vendor

**Definition:** A vendor that appears in three or more weekly reports.

**At Seamfix:** Signals ongoing commitments worth reviewing for contract terms and consolidation.

### Vendor Concentration

**Definition:** How much of total spend goes to the top few vendors.

**At Seamfix:** High concentration is a dependency/negotiation-leverage flag.

---

## Budget vs Actual

_Group/entity/department spend against the approved budget (Budget Tracker snapshot, lean mode, all NGN)._

### Annual Budget

**Definition:** The approved full-year spending plan.

**At Seamfix:** Built bottom-up so Group = sum of entities = sum of departments (~₦5.10B FY2026). UK budgets are in GBP and UAE in USD, FX-converted to Naira; actuals are already Naira.

### YTD Budget (Pace to Date)

**Definition:** The portion of the annual budget allocated to the months elapsed so far.

**At Seamfix:** The benchmark a unit's YTD actual is compared against.

### YTD Actual

**Definition:** Actual spend recorded for the year to date.

**At Seamfix:** Loaded from Acumatica into the Budget Tracker (lean mode). Actuals are Naira for every entity.

### Variance

**Definition:** YTD budget pace minus YTD actual — the over- or under-spend.

**At Seamfix:** Positive = under pace (spending less than planned); negative = over pace.

### % of Pace

**Definition:** YTD actual ÷ YTD budget pace, as a percentage.

**At Seamfix:** Over 105% = Over Budget (red); under 70% = Under Budget (green/under-spending); in between = On Track.

### % of Annual

**Definition:** YTD actual ÷ full annual budget.

**At Seamfix:** How much of the whole-year envelope has been used.

### Year-End Projection

**Definition:** Run-rate forecast of full-year spend from YTD actuals.

**At Seamfix:** Compared to the annual budget to flag likely overspend or savings.

### Entity / Department / Budget Head

**Definition:** The three drill-down levels: legal entity (NG/UK/UAE), department, and individual budget line.

**At Seamfix:** Eleven departments (e.g. Commercial, Solutions, Finance, CAPEX, UK, UAE). Each department expands to its budget heads.

### Lean vs Full mode

**Definition:** Two budget views in the tracker: lean (Acumatica-loaded, carries actuals) and full budget.

**At Seamfix:** The dashboard uses the lean budget to drive a budget-vs-actual comparison.

---

## Revenue & Fundability

_Revenue performance and recurring-revenue quality against the company target._

### Annual Revenue Target

**Definition:** The official company revenue goal for the year.

**At Seamfix:** $8,000,000 ($8M) is the official target used for gap and projection maths, even though the deal bucket sums to a more optimistic ~$10M.

### ARR (Annual Recurring Revenue)

**Definition:** The annualised value of revenue that recurs (subscriptions/contracted recurring deals).

**At Seamfix:** Sum of annual revenue for deals flagged 'Recurring' (~$2.86M across recurring deals in 2026). A measure of predictable, fundable revenue.

### ARR As At <month>

**Definition:** The recurring share of revenue actually EARNED year-to-date, versus target.

**At Seamfix:** Recurring YTD ÷ total YTD revenue, measured against a 50% target (red when below). This is the actuals-based recurring quality, not the contracted plan figure. The Group Financials tab mirrors this exact number.

### Achievement %

**Definition:** Progress of a deal or the portfolio toward its full annual target.

**At Seamfix:** YTD actual ÷ annual target (NOT pace-adjusted) — a deal that earned its full annual target shows 100%.

### Recurring vs Not Recurring

**Definition:** Whether a deal's revenue repeats or is one-off.

**At Seamfix:** Tagged in column D of the revenue sheet; drives the ARR calculations.

### Deficit / Surplus / Gap

**Definition:** How actual revenue compares to the planned amount for a deal.

**At Seamfix:** Deficit = behind plan, Surplus = ahead, Gap = the outstanding shortfall to target.

---

## Pipeline Intelligence

_Deal-level momentum and a probability-weighted view of landing the target._

### Deal Status (On Track / At Risk / Off Track / Closed)

**Definition:** Finance's assessment of each deal's health.

**At Seamfix:** Drives the weighted projection. 'Closed' means won/booked.

### Status Weights

**Definition:** Probability factors applied to each deal's value to build a realistic projection.

**At Seamfix:** On Track 100%, Closed 100%, At Risk 50%, Off Track 10%, Unknown 70%.

### Realistic vs Conservative Projection

**Definition:** Two weighted forecasts of full-year revenue.

**At Seamfix:** Realistic applies the status weights above; Conservative excludes At Risk and Off Track entirely — a downside scenario.

### Gap to Target

**Definition:** Target minus projected revenue.

**At Seamfix:** Positive = shortfall to close; zero or negative = target is projected to be met.

### Momentum (growing / stalled / new / steady)

**Definition:** Trend signal comparing a deal's two most recent months.

**At Seamfix:** Growing = latest > 1.1× prior; stalled = activity dropped to zero; new = only the latest month has revenue; steady = otherwise. Surfaces deals losing traction.

---

## Collections Tracker

_Critical revenue inflows and the status of collecting cash already won._

### Booked

**Definition:** Whether the revenue has been formally recognised/contracted.

**At Seamfix:** 'YES' means it counts as committed revenue; tracked against total expected.

### Predictability

**Definition:** Confidence that the expected cash will actually land.

**At Seamfix:** Helps prioritise follow-up on lower-confidence inflows.

### Closure Period

**Definition:** When the deal/payment is expected to close.

**At Seamfix:** Used with predictability to classify collection urgency.

### Payment Status (Pending / In-Progress / Closed)

**Definition:** Where the inflow sits in the collection cycle.

**At Seamfix:** 'Closed' = collected; pending/in-progress are the focus of weekly collection actions.

### Deal Status

**Definition:** The commercial state of the underlying deal.

**At Seamfix:** Distinct from payment status — a closed deal can still have pending payment.

---

## Group Financials (Consolidated P&L)

_The audited-style consolidated income statement, ratios and value metrics._

### Total Revenue

**Definition:** All income earned from operations in the period.

**At Seamfix:** Reported for the GROUP block; shown in Naira with the report's own USD equivalent.

### Cost of Sales (COGS)

**Definition:** Direct costs of delivering the product/service.

**At Seamfix:** Subtracted from revenue to get gross profit.

### Gross Profit / Gross Margin

**Definition:** Revenue minus cost of sales; margin is that as a % of revenue.

**At Seamfix:** Benchmarked against a 70% gross-margin target on a gauge card.

### Operating Expenses (OpEx)

**Definition:** Running costs not tied directly to delivery (payroll, marketing, admin, etc.).

**At Seamfix:** Broken down with payroll % and marketing % of revenue ratios.

### EBITDA / EBITDA Margin

**Definition:** Earnings Before Interest, Tax, Depreciation & Amortisation — operating profitability before financing and accounting charges.

**At Seamfix:** A proxy for core operating cash generation.

### EBIT

**Definition:** Earnings Before Interest and Tax (EBITDA less depreciation & amortisation).

**At Seamfix:** The operating profit used as the base for NOPAT in the EVA calculation.

### Net Profit (Profit After Tax) / Net Margin

**Definition:** The bottom-line profit after all costs and tax; margin is that as a % of revenue.

**At Seamfix:** Displayed as 'Net Profit' everywhere (the source label is 'Profit After Tax'). Benchmarked against a 10% net-margin target.

### ROA (Return on Assets)

**Definition:** Annualised net profit as a % of total assets — how efficiently assets generate profit.

**At Seamfix:** One of the profitability ratio cards; annualised because the P&L is a partial year.

### ROIC (Return on Invested Capital)

**Definition:** Annualised NOPAT ÷ invested capital — return on the capital actually put to work.

**At Seamfix:** Flagged red when below WACC (the company isn't yet earning its cost of capital).

### WACC (Cost of Capital)

**Definition:** The blended return investors/lenders require — the hurdle rate.

**At Seamfix:** Set to the pre-approved 37%. Returns below this destroy value; used in EVA and the ROIC red flag.

### NOPAT

**Definition:** Net Operating Profit After Tax = EBIT × (1 − effective tax rate).

**At Seamfix:** Because the report books ₦0 tax YTD, a 30% Nigeria statutory rate is used as a fallback so NOPAT isn't overstated.

### Invested Capital

**Definition:** The capital tied up in the business: current assets − current liabilities + net fixed assets.

**At Seamfix:** Parsed from the balance-sheet rows; the base for ROIC and the EVA capital charge.

### EVA (Economic Value Added)

**Definition:** NOPAT minus a capital charge (invested capital × WACC) — value created above the cost of capital.

**At Seamfix:** The capital charge is prorated to the YTD period so it compares fairly with period NOPAT. Negative EVA means returns haven't yet cleared the 37% hurdle.

### Effective Tax Rate

**Definition:** Tax expense as a % of pre-tax profit.

**At Seamfix:** Currently ~0% booked YTD; the 30% statutory rate is substituted for NOPAT/EVA until real tax appears.

### Interest Coverage

**Definition:** How many times operating earnings cover interest expense.

**At Seamfix:** A solvency/financing-risk indicator on the ratios grid.

### Current Ratio / Cash Ratio

**Definition:** Liquidity ratios: current assets (or cash) ÷ current liabilities.

**At Seamfix:** Above 1.0 indicates short-term obligations are comfortably covered.

### Revenue by Vertical / Customer / Country

**Definition:** Breakdowns of where revenue comes from.

**At Seamfix:** Highlight concentration (e.g. a single large customer/segment). These breakdowns exclude Other Income.

### Monthly Performance Trend

**Definition:** Month-by-month chart of Revenue, Gross Profit, EBITDA and Net Profit.

**At Seamfix:** Read from the report's MoM tab and limited to the current fiscal year (auto-advances each year).

---
