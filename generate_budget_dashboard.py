#!/usr/bin/env python3
"""
Seamfix Budget vs Actual Dashboard Generator
=============================================
Renders budget-vs-actual at three levels — GROUP, COMPANY (entity: NG / UK / UAE),
and DEPARTMENT — from a committed JSON snapshot of the Seamfix Budget Tracker
(https://seamfix-budget-tracker.netlify.app/).

WHY a snapshot: the Netlify tracker is a fully client-side page (its data lives in
hardcoded JS objects, refreshed manually per "Run"). There is no API to fetch live,
so we bake the data into data/budget_tracker_snapshot.json and read it here. To
refresh, re-extract the tracker's `BUDGET_DATA`/`C` objects into that file.

MODE: we use the tracker's "lean" mode — the Acumatica-loaded budget that carries
actuals. The "full" (approved) mode has NO actuals, so it cannot drive a
budget-vs-actual view.

CURRENCY: everything rolls up in NGN. NG department budgets are already NGN; UK
budgets are GBP and UAE budgets are USD, so they are FX-converted to NGN using the
tracker's lean FX rates. Actuals are already stored in NGN for every entity.

The dashboard is built bottom-up, so GROUP = sum(COMPANY) = sum(DEPARTMENT) and the
numbers reconcile at every level.

Usage: python3 generate_budget_dashboard.py [folder_path]
Reads:  <folder>/budget_tracker_snapshot.json
Writes: <folder>/budget_dashboard.html
"""

import os, sys, json
from datetime import datetime
from theme import get_base_css, get_toggle_html, get_theme_js

SNAPSHOT_NAME = "budget_tracker_snapshot.json"
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

ENTITY_LABEL = {'NG': 'Nigeria', 'UK': 'United Kingdom', 'UAE': 'United Arab Emirates'}
ENTITY_COLOR = {'NG': '#10b981', 'UK': '#3b82f6', 'UAE': '#f59e0b'}


def sf(v):
    if v is None:
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def fmt_naira(val):
    v = abs(val)
    sign = '-' if val < 0 else ''
    if v >= 1_000_000_000:
        return sign + "\u20A6" + format(v / 1_000_000_000, ".2f") + "B"
    if v >= 1_000_000:
        return sign + "\u20A6" + format(v / 1_000_000, ".1f") + "M"
    if v >= 1_000:
        return sign + "\u20A6" + format(v / 1_000, ".0f") + "K"
    return sign + "\u20A6" + format(v, ",.0f")


def fmt_pct(val):
    return format(val, ".1f") + "%"


def fx_rate(entity, lean_fx):
    """NGN conversion factor for a department's *budget* (actuals are already NGN)."""
    if entity == 'UK':
        return sf(lean_fx.get('GBP_NGN', 2000))
    if entity == 'UAE':
        return sf(lean_fx.get('USD_NGN', 1500))
    return 1.0


def ytd_sum(month_map, elapsed, lower=False):
    total = 0.0
    for i in range(min(elapsed, 12)):
        key = MONTHS[i].lower() if lower else MONTHS[i]
        total += sf(month_map.get(key, 0))
    return total


def compute(snapshot):
    """Build per-department, per-entity and group budget-vs-actual metrics (all NGN)."""
    lean_fx = snapshot.get('leanFX', {})
    elapsed = int(snapshot.get('elapsedMonths', 4) or 4)

    departments = []
    for dept in snapshot.get('departments', []):
        entity = dept.get('entity', 'NG')
        rate = fx_rate(entity, lean_fx)

        annual_budget = sf(dept.get('annual_total', 0)) * rate
        dept_months = dept.get('months', {}) or {}
        budget_month_ngn = [sf(dept_months.get(m, 0)) * rate for m in MONTHS]
        ytd_budget = sum(budget_month_ngn[:elapsed])

        # Actuals come from the budget heads (already NGN).
        actual_month_ngn = [0.0] * 12
        heads = []
        for bh in dept.get('budget_heads', []):
            h_actuals = bh.get('actuals', {}) or {}
            h_months = bh.get('months', {}) or {}
            for i, m in enumerate(MONTHS):
                actual_month_ngn[i] += sf(h_actuals.get(m.lower(), 0))
            h_annual = sf(bh.get('annual', 0)) * rate
            h_ytd_budget = ytd_sum(h_months, elapsed) * rate
            h_ytd_actual = ytd_sum(h_actuals, elapsed, lower=True)
            heads.append({
                'name': bh.get('name', ''),
                'annual_budget': h_annual,
                'ytd_budget': h_ytd_budget,
                'ytd_actual': h_ytd_actual,
                'variance': h_ytd_budget - h_ytd_actual,
            })
        ytd_actual = sum(actual_month_ngn[:elapsed])

        departments.append({
            'code': dept.get('dept_code', ''),
            'name': dept.get('dept_name', ''),
            'entity': entity,
            'native_currency': dept.get('currency', 'NGN'),
            'annual_budget': annual_budget,
            'ytd_budget': ytd_budget,
            'ytd_actual': ytd_actual,
            'variance': ytd_budget - ytd_actual,
            'pct_of_annual': (ytd_actual / annual_budget * 100) if annual_budget > 0 else 0,
            'pct_of_pace': (ytd_actual / ytd_budget * 100) if ytd_budget > 0 else 0,
            'projected': (ytd_actual / elapsed * 12) if elapsed > 0 else 0,
            'budget_month': budget_month_ngn,
            'actual_month': actual_month_ngn,
            'heads': heads,
        })

    # Entity rollups
    entities = {}
    for d in departments:
        e = entities.setdefault(d['entity'], {
            'entity': d['entity'], 'annual_budget': 0.0, 'ytd_budget': 0.0,
            'ytd_actual': 0.0, 'budget_month': [0.0] * 12, 'actual_month': [0.0] * 12,
            'dept_count': 0,
        })
        e['annual_budget'] += d['annual_budget']
        e['ytd_budget'] += d['ytd_budget']
        e['ytd_actual'] += d['ytd_actual']
        e['dept_count'] += 1
        for i in range(12):
            e['budget_month'][i] += d['budget_month'][i]
            e['actual_month'][i] += d['actual_month'][i]
    for e in entities.values():
        e['variance'] = e['ytd_budget'] - e['ytd_actual']
        e['pct_of_annual'] = (e['ytd_actual'] / e['annual_budget'] * 100) if e['annual_budget'] > 0 else 0
        e['pct_of_pace'] = (e['ytd_actual'] / e['ytd_budget'] * 100) if e['ytd_budget'] > 0 else 0
        e['projected'] = (e['ytd_actual'] / elapsed * 12) if elapsed > 0 else 0

    # Group rollup
    group = {
        'annual_budget': sum(d['annual_budget'] for d in departments),
        'ytd_budget': sum(d['ytd_budget'] for d in departments),
        'ytd_actual': sum(d['ytd_actual'] for d in departments),
        'budget_month': [sum(d['budget_month'][i] for d in departments) for i in range(12)],
        'actual_month': [sum(d['actual_month'][i] for d in departments) for i in range(12)],
    }
    group['variance'] = group['ytd_budget'] - group['ytd_actual']
    group['remaining'] = group['annual_budget'] - group['ytd_actual']
    group['pct_of_annual'] = (group['ytd_actual'] / group['annual_budget'] * 100) if group['annual_budget'] > 0 else 0
    group['pct_of_pace'] = (group['ytd_actual'] / group['ytd_budget'] * 100) if group['ytd_budget'] > 0 else 0
    group['projected'] = (group['ytd_actual'] / elapsed * 12) if elapsed > 0 else 0

    return departments, entities, group, elapsed


def status_for(pct_of_pace):
    """Classify a unit by how its YTD actual compares to its YTD budget pace."""
    if pct_of_pace > 105:
        return 'over', '\U0001F534', 'Over Budget'
    if pct_of_pace < 70:
        return 'under', '\U0001F7E2', 'Under Budget'
    return 'ontrack', '\U0001F7E1', 'On Track'


def build_html(snapshot, departments, entities, group, elapsed, out_path):
    generated_at = datetime.now().strftime('%d %b %Y %H:%M')
    last_actuals = snapshot.get('lastActualsMonth', 'Apr-2026')
    run_id = snapshot.get('runId', '')
    run_date = snapshot.get('runDate', '')
    pace_pct = elapsed / 12 * 100
    elapsed_label = ' '.join([MONTHS[i] for i in range(min(elapsed, 12))][:1]) + " \u2013 " + MONTHS[min(elapsed, 12) - 1]

    # --- KPI cards (group) ---
    var = group['variance']
    var_cls = 'positive' if var >= 0 else 'negative'
    var_word = 'under YTD budget' if var >= 0 else 'over YTD budget'
    proj_over = group['projected'] > group['annual_budget']
    proj_cls = 'negative' if proj_over else 'positive'
    proj_word = 'projected overspend' if proj_over else 'projected headroom'
    proj_gap = abs(group['annual_budget'] - group['projected'])

    kpi_html = (
        '<div class="kpi-card"><div class="kpi-label">Annual Budget (Group)</div>'
        '<div class="kpi-value">' + fmt_naira(group['annual_budget']) + '</div>'
        '<div class="kpi-change neutral">FY ' + str(snapshot.get('fiscalYear', '')) + ' \u00b7 ' + str(len(departments)) + ' departments \u00b7 3 entities</div></div>'

        '<div class="kpi-card"><div class="kpi-label">YTD Actual Spend</div>'
        '<div class="kpi-value">' + fmt_naira(group['ytd_actual']) + '</div>'
        '<div class="kpi-change neutral">' + fmt_pct(group['pct_of_annual']) + ' of annual \u00b7 vs ' + fmt_pct(pace_pct) + ' time elapsed</div></div>'

        '<div class="kpi-card"><div class="kpi-label">YTD Budget (' + elapsed_label + ')</div>'
        '<div class="kpi-value">' + fmt_naira(group['ytd_budget']) + '</div>'
        '<div class="kpi-change ' + var_cls + '">' + fmt_naira(abs(var)) + ' ' + var_word + ' (' + fmt_pct(abs(group['pct_of_pace'] - 100)) + ')</div></div>'

        '<div class="kpi-card"><div class="kpi-label">Annual Budget Remaining</div>'
        '<div class="kpi-value">' + fmt_naira(group['remaining']) + '</div>'
        '<div class="kpi-change neutral">' + fmt_pct(100 - group['pct_of_annual']) + ' of envelope unspent</div></div>'

        '<div class="kpi-card"><div class="kpi-label">Projected Year-End</div>'
        '<div class="kpi-value ' + ('negative' if proj_over else '') + '">' + fmt_naira(group['projected']) + '</div>'
        '<div class="kpi-change ' + proj_cls + '">' + fmt_naira(proj_gap) + ' ' + proj_word + ' at current run-rate</div></div>'
    )

    # --- Group health takeaway ---
    if group['pct_of_pace'] > 105:
        h_status, h_color, h_text = 'OVER PACE', 'FF6B6B', 'Group YTD spend is ahead of the budgeted pace. Watch the departments flagged red below.'
    elif group['pct_of_pace'] < 85:
        h_status, h_color, h_text = 'UNDER PACE', '00D4AA', 'Group YTD spend is comfortably below the budgeted pace \u2014 healthy headroom against the annual envelope.'
    else:
        h_status, h_color, h_text = 'ON TRACK', '4ECDC4', 'Group YTD spend is closely aligned with the budgeted pace.'

    # --- Company (entity) cards + table ---
    ordered_entities = [entities[k] for k in ['NG', 'UK', 'UAE'] if k in entities]
    entity_cards = ""
    for e in ordered_entities:
        st_key, st_icon, st_txt = status_for(e['pct_of_pace'])
        bar_pct = min(100, e['pct_of_annual'])
        col = ENTITY_COLOR.get(e['entity'], '#10b981')
        entity_cards += (
            '<div class="entity-card">'
            '<div class="entity-head"><span class="entity-name">' + ENTITY_LABEL.get(e['entity'], e['entity']) + '</span>'
            '<span class="entity-badge" style="background:' + col + '22;color:' + col + '">' + e['entity'] + '</span></div>'
            '<div class="entity-figs"><div><span class="ef-label">YTD Actual</span><span class="ef-val">' + fmt_naira(e['ytd_actual']) + '</span></div>'
            '<div><span class="ef-label">YTD Budget</span><span class="ef-val">' + fmt_naira(e['ytd_budget']) + '</span></div></div>'
            '<div class="entity-bar-track"><div class="entity-bar-fill" style="width:' + format(bar_pct, ".1f") + '%;background:' + col + '"></div>'
            '<div class="entity-bar-pace" style="left:' + format(pace_pct, ".1f") + '%"></div></div>'
            '<div class="entity-foot"><span>' + fmt_pct(e['pct_of_annual']) + ' of ' + fmt_naira(e['annual_budget']) + ' annual</span>'
            '<span>' + st_icon + ' ' + st_txt + '</span></div></div>'
        )

    entity_rows = ""
    for e in ordered_entities:
        st_key, st_icon, st_txt = status_for(e['pct_of_pace'])
        v_cls = 'positive' if e['variance'] >= 0 else 'negative'
        entity_rows += (
            '<tr><td><strong>' + ENTITY_LABEL.get(e['entity'], e['entity']) + '</strong> (' + e['entity'] + ')</td>'
            '<td>' + str(e['dept_count']) + '</td>'
            '<td>' + fmt_naira(e['annual_budget']) + '</td>'
            '<td>' + fmt_naira(e['ytd_budget']) + '</td>'
            '<td>' + fmt_naira(e['ytd_actual']) + '</td>'
            '<td class="' + v_cls + '">' + fmt_naira(abs(e['variance'])) + '</td>'
            '<td>' + fmt_pct(e['pct_of_annual']) + '</td>'
            '<td>' + st_icon + ' ' + st_txt + '</td></tr>'
        )

    # --- Department table (sortable) + drill-down rows ---
    dept_sorted = sorted(departments, key=lambda d: -d['annual_budget'])
    dept_rows = ""
    for idx, d in enumerate(dept_sorted):
        st_key, st_icon, st_txt = status_for(d['pct_of_pace'])
        v_cls = 'positive' if d['variance'] >= 0 else 'negative'
        col = ENTITY_COLOR.get(d['entity'], '#10b981')
        ent_badge = '<span class="ent-badge" style="background:' + col + '22;color:' + col + '">' + d['entity'] + '</span>'
        dept_rows += (
            '<tr class="dept-row" onclick="toggleHeads(' + str(idx) + ')" '
            'data-name="' + (d['name'] + ' ' + d['code'] + ' ' + d['entity']).lower() + '" '
            'data-annual="' + format(d['annual_budget'], ".0f") + '" data-actual="' + format(d['ytd_actual'], ".0f") + '" '
            'data-variance="' + format(d['variance'], ".0f") + '" data-pct="' + format(d['pct_of_annual'], ".2f") + '">'
            '<td><span class="caret" id="caret-' + str(idx) + '">\u25b8</span> ' + d['name'] + ' ' + ent_badge + '</td>'
            '<td>' + fmt_naira(d['annual_budget']) + '</td>'
            '<td>' + fmt_naira(d['ytd_budget']) + '</td>'
            '<td>' + fmt_naira(d['ytd_actual']) + '</td>'
            '<td class="' + v_cls + '">' + fmt_naira(abs(d['variance'])) + '</td>'
            '<td>' + fmt_pct(d['pct_of_annual']) + '</td>'
            '<td>' + st_icon + ' ' + st_txt + '</td></tr>'
        )
        # hidden detail row with budget heads
        head_sorted = sorted(d['heads'], key=lambda h: -max(h['ytd_actual'], h['annual_budget']))
        sub = ""
        for h in head_sorted:
            if h['annual_budget'] == 0 and h['ytd_actual'] == 0:
                continue
            hv_cls = 'positive' if h['variance'] >= 0 else 'negative'
            sub += (
                '<tr class="head-row"><td class="head-name">' + h['name'] + '</td>'
                '<td>' + fmt_naira(h['annual_budget']) + '</td>'
                '<td>' + fmt_naira(h['ytd_budget']) + '</td>'
                '<td>' + fmt_naira(h['ytd_actual']) + '</td>'
                '<td class="' + hv_cls + '">' + fmt_naira(abs(h['variance'])) + '</td>'
                '<td></td><td></td></tr>'
            )
        if not sub:
            sub = '<tr class="head-row"><td colspan="7" style="color:var(--text-tertiary)">No budget-head detail.</td></tr>'
        dept_rows += '<tr class="heads-wrap" id="heads-' + str(idx) + '" style="display:none"><td colspan="7" style="padding:0"><table class="heads-table">' + sub + '</table></td></tr>'

    # --- Chart data ---
    chart_dept_labels = json.dumps([d['name'] for d in dept_sorted])
    chart_dept_budget = json.dumps([round(d['ytd_budget']) for d in dept_sorted])
    chart_dept_actual = json.dumps([round(d['ytd_actual']) for d in dept_sorted])
    months_lbl = json.dumps(MONTHS)
    grp_budget_month = json.dumps([round(x) for x in group['budget_month']])
    # only show actuals for elapsed months (null afterwards so the line stops)
    actual_series = [round(group['actual_month'][i]) if i < elapsed else None for i in range(12)]
    grp_actual_month = json.dumps(actual_series)

    # --- Executive takeaways ---
    over = sorted([d for d in departments if d['pct_of_pace'] > 105 and d['ytd_budget'] > 0],
                  key=lambda d: -(d['ytd_actual'] - d['ytd_budget']))
    under = sorted([d for d in departments if d['pct_of_pace'] < 60 and d['ytd_budget'] > 0],
                   key=lambda d: d['pct_of_pace'])
    # largest single budget heads by YTD actual
    all_heads = []
    for d in departments:
        for h in d['heads']:
            if h['ytd_actual'] > 0:
                all_heads.append((d['name'], h['name'], h['ytd_actual'], h['variance']))
    all_heads.sort(key=lambda x: -x[2])

    def li_dept(d):
        gap = d['ytd_actual'] - d['ytd_budget']
        return '<li><strong>' + d['name'] + '</strong> (' + d['entity'] + '): ' + fmt_naira(d['ytd_actual']) + ' spent vs ' + fmt_naira(d['ytd_budget']) + ' budgeted \u2014 ' + fmt_naira(abs(gap)) + ' over (' + fmt_pct(d['pct_of_pace']) + ' of pace)</li>'

    over_items = ''.join(li_dept(d) for d in over[:5]) or '<li>No department is materially over its YTD budget pace.</li>'
    under_items = ''.join('<li><strong>' + d['name'] + '</strong> (' + d['entity'] + '): only ' + fmt_pct(d['pct_of_pace']) + ' of YTD budget used (' + fmt_naira(d['ytd_actual']) + ' of ' + fmt_naira(d['ytd_budget']) + ')</li>' for d in under[:5]) or '<li>No department is significantly underspending.</li>'
    head_items = ''.join('<li><strong>' + nm + '</strong> \u2013 ' + hn + ': ' + fmt_naira(amt) + ' YTD</li>' for nm, hn, amt, v in all_heads[:6])

    takeaway_html = (
        '<div class="takeaway-card" style="border-left-color:#' + h_color + '">'
        '<div class="tw-header"><span class="tw-icon">\U0001F4CA</span><span class="tw-title">Group Budget Health</span>'
        '<span class="tw-badge" style="background:#' + h_color + '22;color:#' + h_color + '">' + h_status + '</span></div>'
        '<div class="tw-headline">' + fmt_naira(group['ytd_actual']) + ' spent of ' + fmt_naira(group['ytd_budget']) + ' YTD budget (' + fmt_pct(group['pct_of_pace']) + ' of pace)</div>'
        '<div class="tw-body">' + h_text + ' Annual envelope ' + fmt_naira(group['annual_budget']) + '; ' + fmt_pct(group['pct_of_annual']) + ' consumed against ' + fmt_pct(pace_pct) + ' of the year elapsed.</div></div>'

        '<div class="takeaway-card" style="border-left-color:#FF6B6B">'
        '<div class="tw-header"><span class="tw-icon">\U0001F525</span><span class="tw-title">Departments Over Pace</span></div>'
        '<div class="tw-body"><ul>' + over_items + '</ul></div></div>'

        '<div class="takeaway-card" style="border-left-color:#00D4AA">'
        '<div class="tw-header"><span class="tw-icon">\U0001F4B0</span><span class="tw-title">Departments Underspending</span></div>'
        '<div class="tw-body"><ul>' + under_items + '</ul></div></div>'

        '<div class="takeaway-card" style="border-left-color:#3b82f6">'
        '<div class="tw-header"><span class="tw-icon">\U0001F50D</span><span class="tw-title">Largest Spend Lines (YTD)</span></div>'
        '<div class="tw-body"><ul>' + head_items + '</ul></div></div>'
    )

    theme_css = get_base_css()
    toggle_html = get_toggle_html()
    theme_js = get_theme_js()

    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Seamfix Budget vs Actual</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
""" + theme_css + """
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',sans-serif;background:linear-gradient(135deg,var(--bg-body),var(--bg-card));color:var(--text-primary);min-height:100vh}
.container{max-width:1600px;margin:0 auto;padding:24px 28px 40px}
.header{margin-bottom:24px;padding-bottom:18px;border-bottom:2px solid var(--border-accent)}
.header h1{font-size:2.2em;font-weight:700;background:linear-gradient(135deg,var(--accent),var(--accent-secondary));-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.header .sub{color:var(--text-secondary);font-size:0.95em;margin-top:4px}
.header .meta{color:var(--text-tertiary);font-size:0.8em;margin-top:6px}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:16px;margin-bottom:30px}
.kpi-card{background:var(--bg-card);border:1px solid var(--border-accent);border-radius:12px;padding:20px;transition:all .3s}
.kpi-card:hover{border-color:rgba(0,212,170,0.4);transform:translateY(-3px);box-shadow:var(--shadow-hover)}
.kpi-label{font-size:0.74em;color:var(--text-secondary);text-transform:uppercase;letter-spacing:1.1px;margin-bottom:8px;font-weight:600}
.kpi-value{font-size:1.5em;font-weight:700;color:var(--accent);margin-bottom:6px}
.kpi-value.negative{color:var(--danger)}
.kpi-change{font-size:0.8em;display:flex;align-items:center;gap:4px}
.kpi-change.positive{color:var(--accent)}.kpi-change.negative{color:var(--danger)}.kpi-change.neutral{color:var(--text-tertiary)}
.takeaways-section{margin-bottom:32px}
.takeaways-section h2{font-size:1.3em;margin-bottom:16px;color:var(--warning)}
.takeaways-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(440px,1fr));gap:16px}
.takeaway-card{background:var(--bg-card);border:1px solid var(--border-accent);border-left:4px solid var(--text-tertiary);border-radius:10px;padding:18px}
.tw-header{display:flex;align-items:center;gap:8px;margin-bottom:10px}
.tw-icon{font-size:1.2em}.tw-title{font-size:0.8em;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--text-secondary)}
.tw-badge{font-size:0.7em;font-weight:700;padding:3px 10px;border-radius:12px;margin-left:auto}
.tw-headline{font-size:1.02em;font-weight:600;margin-bottom:8px;line-height:1.4;color:var(--text-heading)}
.tw-body{font-size:0.86em;color:var(--text-secondary);line-height:1.6}
.tw-body ul{margin:4px 0 0 16px}.tw-body li{margin-bottom:5px}
.section{background:var(--bg-card);border:1px solid var(--border-accent);border-radius:12px;padding:24px;margin-bottom:24px}
.section h2{font-size:1.25em;margin-bottom:6px;color:var(--accent)}
.section .desc{font-size:0.85em;color:var(--text-tertiary);margin-bottom:18px}
.entity-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px;margin-bottom:22px}
.entity-card{background:var(--bg-body);border:1px solid var(--border-main);border-radius:10px;padding:18px}
.entity-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}
.entity-name{font-weight:700;font-size:1.05em;color:var(--text-heading)}
.entity-badge{font-size:0.7em;font-weight:700;padding:3px 9px;border-radius:10px}
.entity-figs{display:flex;justify-content:space-between;margin-bottom:12px}
.ef-label{display:block;font-size:0.7em;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:2px}
.ef-val{font-weight:600;font-size:1.05em;color:var(--text-heading)}
.entity-bar-track{position:relative;height:9px;background:var(--bg-gauge);border-radius:5px;overflow:visible;margin-bottom:10px}
.entity-bar-fill{height:100%;border-radius:5px}
.entity-bar-pace{position:absolute;top:-3px;width:2px;height:15px;background:var(--text-secondary);opacity:0.7}
.entity-foot{display:flex;justify-content:space-between;font-size:0.76em;color:var(--text-tertiary)}
.charts-grid{display:grid;grid-template-columns:1fr;gap:20px;margin-bottom:24px}
.chart-box{background:var(--bg-card);border:1px solid var(--border-accent);border-radius:12px;padding:22px}
.chart-box h3{font-size:1em;font-weight:600;margin-bottom:16px;color:var(--text-heading)}
.chart-wrap{position:relative;height:360px}
table{width:100%;border-collapse:collapse;font-size:0.85em}
thead{background:var(--bg-table-header)}
th{padding:10px 12px;text-align:left;font-weight:600;color:var(--accent);border-bottom:2px solid var(--border-accent)}
td{padding:9px 12px;border-bottom:1px solid var(--border-light);color:var(--text-heading)}
tbody tr:hover{background:var(--bg-table-hover)}
.positive{color:#00D4AA}.negative{color:#FF6B6B}
.dept-row{cursor:pointer}
.caret{display:inline-block;width:12px;color:var(--accent);transition:transform .15s}
.ent-badge{font-size:0.68em;font-weight:700;padding:2px 7px;border-radius:8px;margin-left:6px}
.heads-table{width:100%;background:var(--bg-body)}
.heads-table td{padding:7px 12px 7px 34px;font-size:0.92em;border-bottom:1px solid var(--border-light);color:var(--text-secondary)}
.head-name{color:var(--text-secondary)}
.search-box{margin-bottom:14px;padding:11px 14px;background:var(--accent-bg);border:1px solid var(--border-accent);border-radius:8px;color:var(--text-primary);font-family:inherit;width:100%;max-width:380px}
.search-box::placeholder{color:var(--text-tertiary)}
th.sortable{cursor:pointer;user-select:none}
th.sortable:hover{color:var(--warning)}
.note{font-size:0.78em;color:var(--text-tertiary);margin-top:14px;line-height:1.5}
@media(max-width:1024px){.kpi-grid{grid-template-columns:repeat(2,1fr)}.takeaways-grid{grid-template-columns:1fr}}
@media(max-width:640px){.kpi-grid{grid-template-columns:1fr}.header h1{font-size:1.6em}}
</style>
</head>
<body>
""" + toggle_html + """
<div class="container">
<div class="header">
<h1>Budget vs Actual</h1>
<div class="sub">Group, company and department spend against the FY """ + str(snapshot.get('fiscalYear', '')) + """ budget</div>
<div class="meta">Source: Seamfix Budget Tracker (""" + str(run_id) + """, """ + str(run_date) + """) \u00b7 actuals through """ + str(last_actuals) + """ \u00b7 lean (Acumatica-loaded) budget \u00b7 all figures NGN \u00b7 generated """ + generated_at + """</div>
</div>

<div class="kpi-grid">""" + kpi_html + """</div>

<div class="takeaways-section">
<h2>\U0001F4CC Executive Summary</h2>
<div class="takeaways-grid">""" + takeaway_html + """</div>
</div>

<div class="section">
<h2>Company-Wide (by Entity)</h2>
<div class="desc">Budget consumed vs annual envelope for each legal entity. The marker on each bar shows the """ + fmt_pct(pace_pct) + """ time-elapsed pace line. UK budgets converted from GBP and UAE from USD at lean FX; actuals are reported in NGN.</div>
<div class="entity-grid">""" + entity_cards + """</div>
<table>
<thead><tr><th>Entity</th><th>Depts</th><th>Annual Budget</th><th>YTD Budget</th><th>YTD Actual</th><th>Variance</th><th>% of Annual</th><th>Status</th></tr></thead>
<tbody>""" + entity_rows + """</tbody>
</table>
</div>

<div class="charts-grid">
<div class="chart-box"><h3>YTD Budget vs Actual by Department (NGN)</h3><div class="chart-wrap"><canvas id="deptChart"></canvas></div></div>
<div class="chart-box"><h3>Group Monthly Budget vs Actual (NGN)</h3><div class="chart-wrap"><canvas id="trendChart"></canvas></div></div>
</div>

<div class="section">
<h2>Department-Wide</h2>
<div class="desc">Click any department to expand its budget heads. Search or sort to drill in.</div>
<input type="text" class="search-box" id="deptSearch" placeholder="Search department / entity..." onkeyup="filterDepts()">
<table id="deptTable">
<thead><tr>
<th class="sortable" onclick="sortDepts('name')">Department</th>
<th class="sortable" onclick="sortDepts('annual')">Annual Budget</th>
<th>YTD Budget</th>
<th class="sortable" onclick="sortDepts('actual')">YTD Actual</th>
<th class="sortable" onclick="sortDepts('variance')">Variance</th>
<th class="sortable" onclick="sortDepts('pct')">% of Annual</th>
<th>Status</th>
</tr></thead>
<tbody id="deptBody">""" + dept_rows + """</tbody>
</table>
<div class="note">Built bottom-up so Group = sum of entities = sum of departments. Variance is YTD budget minus YTD actual (positive = under budget). "% of Annual" measures YTD actual against the full-year budget; compare it to """ + fmt_pct(pace_pct) + """ time elapsed. Lean mode = Acumatica-loaded budget that carries actuals; the approved ("full") budget has no actuals so is not shown here.</div>
</div>

</div>
<script>
const DEPT_LABELS=""" + chart_dept_labels + """;
const DEPT_BUDGET=""" + chart_dept_budget + """;
const DEPT_ACTUAL=""" + chart_dept_actual + """;
const MONTH_LABELS=""" + months_lbl + """;
const GRP_BUDGET=""" + grp_budget_month + """;
const GRP_ACTUAL=""" + grp_actual_month + """;

function fmtN(v){v=Math.abs(v);if(v>=1e9)return '\u20A6'+(v/1e9).toFixed(2)+'B';if(v>=1e6)return '\u20A6'+(v/1e6).toFixed(1)+'M';if(v>=1e3)return '\u20A6'+(v/1e3).toFixed(0)+'K';return '\u20A6'+v.toFixed(0);}

new Chart(document.getElementById('deptChart'),{type:'bar',data:{labels:DEPT_LABELS,datasets:[
{label:'YTD Budget',data:DEPT_BUDGET,backgroundColor:'rgba(59,130,246,0.55)',borderColor:'#3b82f6',borderWidth:1},
{label:'YTD Actual',data:DEPT_ACTUAL,backgroundColor:'rgba(0,212,170,0.6)',borderColor:'#00D4AA',borderWidth:1}]},
options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8'}},tooltip:{callbacks:{label:function(c){return c.dataset.label+': '+fmtN(c.parsed.y);}}}},
scales:{x:{ticks:{color:'#64748b'},grid:{display:false}},y:{ticks:{color:'#64748b',callback:function(v){return fmtN(v);}},grid:{color:'rgba(0,0,0,0.06)'}}}}});

new Chart(document.getElementById('trendChart'),{type:'line',data:{labels:MONTH_LABELS,datasets:[
{label:'Monthly Budget',data:GRP_BUDGET,borderColor:'#3b82f6',backgroundColor:'rgba(59,130,246,0.08)',fill:true,tension:0.3},
{label:'Monthly Actual',data:GRP_ACTUAL,borderColor:'#00D4AA',backgroundColor:'rgba(0,212,170,0.08)',fill:true,tension:0.3,spanGaps:false}]},
options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#94a3b8'}},tooltip:{callbacks:{label:function(c){return c.dataset.label+': '+(c.parsed.y==null?'n/a':fmtN(c.parsed.y));}}}},
scales:{x:{ticks:{color:'#64748b'},grid:{display:false}},y:{ticks:{color:'#64748b',callback:function(v){return fmtN(v);}},grid:{color:'rgba(0,0,0,0.06)'}}}}});

function toggleHeads(i){var r=document.getElementById('heads-'+i);var c=document.getElementById('caret-'+i);
if(r.style.display==='none'){r.style.display='table-row';c.style.transform='rotate(90deg)';}else{r.style.display='none';c.style.transform='';}}

function filterDepts(){var q=document.getElementById('deptSearch').value.toLowerCase();
var rows=document.querySelectorAll('#deptBody tr.dept-row');
rows.forEach(function(r){var show=r.getAttribute('data-name').indexOf(q)>-1;
r.style.display=show?'':'none';
var id=r.getAttribute('onclick').match(/\\d+/)[0];
var h=document.getElementById('heads-'+id);if(h&&!show){h.style.display='none';var c=document.getElementById('caret-'+id);if(c)c.style.transform='';}});}

var sortState={};
function sortDepts(key){var body=document.getElementById('deptBody');
var pairs=[];var rows=Array.from(body.querySelectorAll('tr.dept-row'));
rows.forEach(function(r){var id=r.getAttribute('onclick').match(/\\d+/)[0];pairs.push([r,document.getElementById('heads-'+id)]);});
var asc=!sortState[key];sortState={};sortState[key]=asc;
pairs.sort(function(a,b){var av,bv;
if(key==='name'){av=a[0].getAttribute('data-name');bv=b[0].getAttribute('data-name');return asc?av.localeCompare(bv):bv.localeCompare(av);}
av=parseFloat(a[0].getAttribute('data-'+key));bv=parseFloat(b[0].getAttribute('data-'+key));return asc?av-bv:bv-av;});
pairs.forEach(function(p){body.appendChild(p[0]);if(p[1])body.appendChild(p[1]);});}
""" + theme_js + """
</script>
</body>
</html>"""

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)


def write_placeholder(out_path, message):
    theme_css = get_base_css()
    toggle_html = get_toggle_html()
    theme_js = get_theme_js()
    html = """<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Budget vs Actual</title><style>""" + theme_css + """
body{font-family:'Inter',sans-serif;background:var(--bg-body);color:var(--text-primary);display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}
.card{background:var(--bg-card);border:1px solid var(--border-accent);border-radius:12px;padding:40px;max-width:560px;text-align:center}
h1{color:var(--accent);font-size:1.4em;margin-bottom:12px}p{color:var(--text-secondary);line-height:1.6}</style></head>
<body>""" + toggle_html + """<div class="card"><h1>Budget vs Actual</h1><p>""" + message + """</p></div>
<script>""" + theme_js + """</script></body></html>"""
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    if not os.path.isdir(folder):
        print("Error: " + folder + " not found")
        sys.exit(1)

    out_path = os.path.join(folder, 'budget_dashboard.html')
    snap_path = os.path.join(folder, SNAPSHOT_NAME)

    if not os.path.exists(snap_path):
        print("Snapshot not found: " + snap_path + " (writing placeholder)")
        write_placeholder(out_path, "Budget data snapshot (budget_tracker_snapshot.json) is not available in this environment.")
        sys.exit(0)

    try:
        with open(snap_path, 'r', encoding='utf-8') as f:
            snapshot = json.load(f)
    except (ValueError, OSError) as e:
        print("Could not read snapshot: " + str(e))
        write_placeholder(out_path, "Budget data snapshot could not be read.")
        sys.exit(0)

    departments, entities, group, elapsed = compute(snapshot)
    if not departments:
        write_placeholder(out_path, "Budget snapshot contained no departments.")
        sys.exit(0)

    build_html(snapshot, departments, entities, group, elapsed, out_path)
    print("Dashboard: " + out_path)
    print("Group: annual=" + fmt_naira(group['annual_budget']) + " ytd_budget=" + fmt_naira(group['ytd_budget']) + " ytd_actual=" + fmt_naira(group['ytd_actual']))


if __name__ == '__main__':
    main()
