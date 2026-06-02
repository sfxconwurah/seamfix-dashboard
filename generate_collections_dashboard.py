#!/usr/bin/env python3
"""
Seamfix Collections Intelligence Dashboard Generator
Tracks critical revenue inflows, weekly progress, movement between weeks, and critical actions.
Usage: python3 generate_collections_dashboard.py [folder_path]

Data source: the "2026 CRITICAL REVENUE INFLOWS" tab of the Collections Tracker
Google Sheet, fetched by app.py as CSV pinned to its gid (1584269897) and saved
as "2026 Collections Tracker.csv". CSV is used (not xlsx) because the workbook
holds several near-identical tabs; pinning to the gid guarantees the right one
regardless of tab renaming or reordering.

Weekly updates are fully dynamic: each new "Update - <date>" column Finance adds
is auto-detected (see is_date_like), so no code change is needed week to week —
the same behaviour as the 2026 Path to Revenue dashboard.
"""

import os, sys, csv
from datetime import datetime
from pathlib import Path
from theme import get_base_css, get_toggle_html, get_theme_js

FX_RATE = 1450  # $1 = ₦1,450
COLLECTIONS_FILENAME = "2026 Collections Tracker.csv"


# ── Formatters ────────────────────────────────────────────────────────────────

def sf(val):
    """Safe float conversion: handles $, ₦, commas."""
    if val is None:
        return 0.0
    s = str(val).replace(',', '').replace('$', '').replace('₦', '').replace('%', '').strip()
    try:
        return float(s)
    except Exception:
        return 0.0


def fmt_usd(val):
    v = abs(float(val))
    sign = '-' if float(val) < 0 else ''
    if v >= 1_000_000:
        return f"{sign}${v / 1_000_000:.2f}M"
    if v >= 1_000:
        return f"{sign}${v / 1_000:.1f}K"
    return f"{sign}${v:,.0f}"


def fmt_ngn(val):
    v = abs(float(val))
    sign = '-' if float(val) < 0 else ''
    if v >= 1_000_000_000:
        return f"{sign}₦{v / 1_000_000_000:.2f}B"
    if v >= 1_000_000:
        return f"{sign}₦{v / 1_000_000:.1f}M"
    if v >= 1_000:
        return f"{sign}₦{v / 1_000:.0f}K"
    return f"{sign}₦{v:,.0f}"


def esc(s):
    """HTML-escape a string for safe embedding."""
    return (str(s)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


# ── xlsx parsing ──────────────────────────────────────────────────────────────

def is_date_like(val):
    """Return True if the value looks like a weekly update header, e.g. 'Update - 2nd Jan', '29th May'."""
    if not val:
        return False
    s = str(val).strip().lower()
    months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
    return any(m in s for m in months) and any(c.isdigit() for c in s)


def clean_date_label(val):
    """Turn 'Update - 2nd Jan' / 'Update 29th May' into a tidy '2nd Jan' / '29th May'."""
    s = str(val).strip()
    for prefix in ('Update -', 'Update-', 'Update'):
        if s.lower().startswith(prefix.lower()):
            s = s[len(prefix):]
            break
    return s.lstrip('-').strip() or str(val).strip()


FIELD_KEYWORDS = {
    'sn':             ['S/N', 'SN', 'S / N'],
    'name':           ['REVENUE ITEM'],
    'vertical':       ['VERTICAL'],
    'customer':       ['CUSTOMER'],
    'product':        ['PRODUCT'],
    'usd':            ['AMOUNT IN USD', 'AMOUNT (USD)', 'USD AMOUNT'],
    'ngn':            ['AMOUNT (NAIRA)', 'AMOUNT IN NAIRA', 'NAIRA', 'NGN AMOUNT'],
    'booked':         ['BOOKED REVENUE'],
    'predictability': ['CLOSURE PREDICTABILITY', 'PREDICTABILITY'],
    'closure_period': ['ESTIMATED DEAL CLOSURE', 'CLOSURE PERIOD', 'DEAL CLOSURE PERIOD'],
    'deal_status':    ['DEAL STATUS'],
    'payment_status': ['PAYMENT STATUS'],
    'action':         ['ACTIONABLE DELIVERABLE', 'DELIVERABLE', 'ACTIONABLE'],
    'accountable':    ['ACCOUNTABLE PARTY', 'ACCOUNTABLE'],
}

SKIP_NAMES = {
    'ANCHOR DEALS', 'EXISTING CUSTOMERS', 'NEW BUSINESS', 'DEALS FROM 2025',
    'TOTAL', 'GRAND TOTAL', 'SUB-TOTAL', 'NOTE', 'NOTES',
}


def find_header_row(rows):
    """Scan first 10 rows for the row that contains the S/N column header."""
    for idx, row in enumerate(rows[:10]):
        for cell in row:
            if cell and str(cell).strip().upper() in ('S/N', 'SN', 'S / N'):
                return idx
    return 2  # fallback (title rows usually occupy rows 0-1)


def extract_collections(csv_path):
    """
    Parse the Collections Tracker CSV (a single sheet tab, pinned by gid in app.py).
    Returns a list of deal dicts, each with an 'updates' list of {date, text} dicts.
    """
    with open(csv_path, newline='', encoding='utf-8') as f:
        rows = list(csv.reader(f))

    if not rows:
        return []

    header_idx = find_header_row(rows)
    header = rows[header_idx]

    # Build column map and detect weekly-update columns
    col_map = {}
    update_cols = []  # list of (col_index, clean_date_label)

    for ci, hval in enumerate(header):
        if not hval:
            continue
        key = str(hval).strip().upper()

        if is_date_like(hval):
            update_cols.append((ci, clean_date_label(hval)))
            continue

        for field, keywords in FIELD_KEYWORDS.items():
            if field not in col_map and any(kw in key for kw in keywords):
                col_map[field] = ci
                break

    # Parse data rows
    items = []
    for row in rows[header_idx + 1:]:
        if not any(str(c).strip() for c in row):
            continue  # skip blank rows

        def get(field):
            ci = col_map.get(field)
            if ci is None or ci >= len(row):
                return None
            v = row[ci]
            return str(v).strip() if v is not None and str(v).strip() else None

        name = get('name')
        if not name:
            continue
        n_upper = name.strip().upper()
        if n_upper in SKIP_NAMES:
            continue
        if any(x in n_upper for x in ('TOTAL', 'GRAND', 'SUB-TOTAL', 'USD RATE', 'EXCHANGE RATE')):
            continue

        # Collect weekly updates
        updates = []
        for ci, date_label in update_cols:
            text = str(row[ci]).strip() if ci < len(row) and row[ci] is not None else ''
            updates.append({'date': date_label, 'text': text})

        items.append({
            'sn':             get('sn') or '',
            'name':           name,
            'vertical':       get('vertical') or '',
            'customer':       get('customer') or '',
            'product':        get('product') or '',
            'usd':            sf(get('usd')),
            'ngn':            sf(get('ngn')),
            'booked':         (get('booked') or '').upper(),
            'predictability': (get('predictability') or '').upper(),
            'closure_period': get('closure_period') or '',
            'deal_status':    get('deal_status') or '',
            'payment_status': get('payment_status') or '',
            'action':         get('action') or '',
            'accountable':    get('accountable') or '',
            'updates':        updates,
        })

    return items


# ── Analysis helpers ──────────────────────────────────────────────────────────

def detect_movement(item):
    """
    Compare the last two non-empty weekly updates.
    Returns (status, detail) where status is: 'positive' | 'negative' | 'updated' | 'no_change'
    """
    filled = [u for u in item['updates'] if u['text'].strip()]
    if len(filled) < 2:
        return 'no_change', ''

    curr = filled[-1]['text'].lower()
    prev = filled[-2]['text'].lower()

    if curr.strip() == prev.strip():
        return 'no_change', ''

    positive_words = ['payment received', 'received', 'paid', 'closed', 'completed', 'approved', 'signed', 'invoiced', 'awarded']
    negative_words = ['at risk', 'overdue', 'no response', 'no feedback', 'extension', 'delayed', 'stalled', 'management approval pending']

    new_pos = [w for w in positive_words if w in curr and w not in prev]
    if new_pos:
        return 'positive', new_pos[0]

    new_neg = [w for w in negative_words if w in curr and w not in prev]
    if new_neg:
        return 'negative', new_neg[0]

    return 'updated', ''


CRITICAL_KEYWORDS = [
    'at risk', 'overdue', 'no response', 'no feedback',
    'management approval pending', 'may 31', 'june 11', 'stalled',
]


def classify_urgency(item):
    """Return urgency level: 'collected' | 'critical' | 'high' | 'medium'."""
    payment = item['payment_status'].lower()

    # Fully closed/collected — no action needed
    if payment == 'closed' and 'in-progress' not in payment:
        return 'collected'

    filled = [u for u in item['updates'] if u['text'].strip()]
    latest = filled[-1]['text'].lower() if filled else ''

    # Largest single item with pending payment
    if item['usd'] >= 500_000 and 'pending' in payment:
        return 'critical'

    # Update text contains a critical keyword
    if any(kw in latest for kw in CRITICAL_KEYWORDS):
        return 'critical'

    # High-value deal with pending payment or low predictability
    if item['usd'] >= 50_000 and ('pending' in payment or item['predictability'] == 'LOW'):
        return 'high'

    # Has a specific actionable deliverable (not generic placeholders)
    generic_actions = {'', 'deal closed', 'po issued', 'deal closed and po available for pickup',
                       'deal closed and po collected'}
    if item['action'].strip().lower() not in generic_actions:
        return 'high'

    return 'medium'


# ── HTML generation ───────────────────────────────────────────────────────────

def generate_html(items, update_dates):
    now         = datetime.now().strftime('%d %b %Y')
    latest_date = update_dates[-1] if update_dates else 'Latest'
    prev_date   = update_dates[-2] if len(update_dates) >= 2 else None

    # KPI summary
    total_usd   = sum(i['usd'] for i in items)
    total_ngn   = sum(i['ngn'] for i in items)
    collected   = [i for i in items if i['payment_status'].lower() == 'closed']
    in_progress = [i for i in items if 'in-progress' in i['payment_status'].lower()]
    pending     = [i for i in items if 'pending' in i['payment_status'].lower()]

    coll_usd = sum(i['usd'] for i in collected)
    prog_usd = sum(i['usd'] for i in in_progress)
    pend_usd = sum(i['usd'] for i in pending)

    coll_pct = round(coll_usd / total_usd * 100) if total_usd else 0
    prog_pct = round(prog_usd / total_usd * 100) if total_usd else 0
    pend_pct = round(pend_usd / total_usd * 100) if total_usd else 0

    # Urgency classification
    critical_items  = [i for i in items if classify_urgency(i) == 'critical']
    high_items      = [i for i in items if classify_urgency(i) == 'high']
    critical_items.sort(key=lambda x: -x['usd'])
    high_items.sort(key=lambda x: -x['usd'])

    # What moved
    moved_items = []
    for item in items:
        status, detail = detect_movement(item)
        if status != 'no_change':
            moved_items.append({**item, '_movement': status, '_detail': detail})

    # ── CSS ──────────────────────────────────────────────────────────────────
    theme_css = get_base_css()
    toggle_html = get_toggle_html()
    theme_js = get_theme_js()

    css = f"""
{theme_css}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: var(--bg-nav);
    color: var(--text-primary);
    padding: 28px 28px 60px;
    font-size: 14px;
    line-height: 1.55;
}}

/* Header */
.page-header {{ margin-bottom: 6px; }}
.page-title  {{ font-size: 26px; font-weight: 700; color: var(--accent); }}
.page-sub    {{ font-size: 13px; color: var(--text-tertiary); margin-top: 3px; margin-bottom: 28px; }}

/* Badges */
.badge {{
    display: inline-block; padding: 2px 10px; border-radius: 20px;
    font-size: 11px; font-weight: 600; text-transform: uppercase; white-space: nowrap;
}}
.b-paid     {{ background: rgba(16,185,129,.15); color: #10b981; border: 1px solid rgba(16,185,129,.3); }}
.b-progress {{ background: rgba(59,130,246,.15); color: #60a5fa; border: 1px solid rgba(59,130,246,.3); }}
.b-pending  {{ background: rgba(245,158,11,.15); color: #f59e0b; border: 1px solid rgba(245,158,11,.3); }}
.b-critical {{ background: rgba(239,68,68,.15);  color: #f87171; border: 1px solid rgba(239,68,68,.3);  }}
.b-high     {{ background: rgba(251,146,60,.15); color: #fb923c; border: 1px solid rgba(251,146,60,.3); }}

/* KPI row */
.kpi-row {{
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 16px; margin-bottom: 32px;
}}
.kpi-card {{
    background: var(--bg-card); border: 1px solid var(--border-main); border-radius: 12px;
    padding: 20px 22px;
}}
.kpi-label {{ font-size: 11px; color: var(--text-tertiary); text-transform: uppercase; letter-spacing: .5px; margin-bottom: 8px; }}
.kpi-val   {{ font-size: 24px; font-weight: 700; }}
.kpi-sub   {{ font-size: 12px; color: var(--text-tertiary); margin-top: 4px; }}

/* Section headers */
.sec-hdr {{
    display: flex; align-items: center; gap: 12px;
    margin: 32px 0 16px; padding-bottom: 12px;
    border-bottom: 1px solid var(--border-main);
}}
.sec-title {{ font-size: 16px; font-weight: 700; color: var(--text-primary); }}
.sec-sub   {{ font-size: 12px; color: var(--text-tertiary); }}

/* Card grids */
.card-grid {{
    display: grid; grid-template-columns: repeat(auto-fill, minmax(360px, 1fr));
    gap: 14px; margin-bottom: 8px;
}}
.action-card {{
    background: var(--bg-card); border: 1px solid var(--border-main); border-radius: 10px;
    padding: 16px 18px;
}}
.action-card.critical {{ border-left: 4px solid #ef4444; }}
.action-card.high     {{ border-left: 4px solid #f59e0b; }}
.card-row {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 5px; }}
.card-name {{ font-weight: 600; font-size: 13px; color: var(--text-primary); }}
.card-amount {{ font-size: 13px; font-weight: 700; color: var(--accent); margin-bottom: 5px; }}
.card-who   {{ font-size: 11px; color: var(--text-tertiary); margin-bottom: 10px; }}
.card-action {{
    font-size: 12px; color: #fbbf24; background: rgba(251,191,36,.06);
    padding: 7px 10px; border-radius: 6px; margin-bottom: 7px;
}}
.card-update {{ font-size: 11px; color: var(--text-secondary); font-style: italic; }}

/* Movement cards */
.mv-card {{
    background: var(--bg-card); border: 1px solid var(--border-main); border-radius: 10px;
    padding: 14px 16px;
}}
.mv-card.positive {{ border-left: 4px solid #10b981; }}
.mv-card.updated  {{ border-left: 4px solid #3b82f6; }}
.mv-card.negative {{ border-left: 4px solid #ef4444; }}
.mv-name     {{ font-weight: 600; font-size: 13px; color: var(--text-primary); margin-bottom: 2px; }}
.mv-customer {{ font-size: 11px; color: var(--text-tertiary); margin-bottom: 10px; }}
.mv-lbl      {{ font-size: 10px; font-weight: 600; text-transform: uppercase; color: var(--text-tertiary); margin-bottom: 3px; }}
.mv-prev {{ font-size: 11px; color: var(--text-tertiary); background: var(--bg-body); padding: 7px 9px; border-radius: 6px; margin-bottom: 7px; }}
.mv-curr {{ font-size: 11px; color: var(--text-heading); background: var(--bg-card); padding: 7px 9px; border-radius: 6px; }}

/* Full tracker table */
.tbl-wrap {{ overflow-x: auto; margin-top: 4px; }}
table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
th {{
    background: var(--bg-table-header); color: var(--text-tertiary);
    font-weight: 600; text-transform: uppercase; font-size: 10px; letter-spacing: .5px;
    padding: 10px 12px; text-align: left; border-bottom: 1px solid var(--border-main);
    white-space: nowrap;
}}
td {{ padding: 10px 12px; border-bottom: 1px solid var(--border-light); vertical-align: top; }}
tr:hover td {{ background: var(--bg-table-hover); }}
.t-name {{ font-weight: 600; color: var(--text-primary); }}
.t-cust {{ font-size: 11px; color: var(--text-tertiary); }}
.t-usd  {{ font-weight: 700; color: var(--accent); white-space: nowrap; }}
.t-ngn  {{ font-size: 11px; color: var(--text-tertiary); white-space: nowrap; }}
.t-upd  {{ max-width: 300px; color: var(--text-tertiary); font-size: 11px; }}
.t-act  {{ max-width: 220px; color: #fbbf24; font-size: 11px; }}
.t-acc  {{ color: #818cf8; font-size: 12px; font-weight: 500; }}
.pred-HIGH   {{ color: #10b981; font-size: 11px; font-weight: 700; }}
.pred-MEDIUM {{ color: #f59e0b; font-size: 11px; font-weight: 700; }}
.pred-LOW    {{ color: #ef4444; font-size: 11px; font-weight: 700; }}

/* Filter bar */
.filter-bar {{ display: flex; gap: 8px; flex-wrap: wrap; align-items: center; margin-bottom: 16px; }}
.f-btn {{
    background: var(--bg-card); border: 1px solid var(--border-main); color: var(--text-tertiary);
    padding: 5px 14px; border-radius: 20px; font-size: 12px; cursor: pointer; transition: all .15s;
}}
.f-btn.active, .f-btn:hover {{ background: rgba(0,212,170,.1); border-color: var(--accent); color: var(--accent); }}
/* Theme toggle in header */
.theme-toggle-wrap {{ display: flex; justify-content: flex-end; margin-bottom: 8px; }}
"""

    # ── Build action card HTML ────────────────────────────────────────────────
    def action_card(item, level):
        filled   = [u for u in item['updates'] if u['text'].strip()]
        latest   = filled[-1]['text'] if filled else ''
        action   = item['action'].strip()
        generic  = {'deal closed', 'po issued', 'deal closed and po available for pickup',
                    'deal closed and po collected', ''}
        show_act = action.lower() not in generic

        badge = '<span class="badge b-critical">Critical</span>' if level == 'critical' else \
                '<span class="badge b-high">Follow Up</span>'

        act_html = f'<div class="card-action">&#9889; {esc(action)}</div>' if show_act else ''
        upd_html = f'<div class="card-update">&#128172; {esc(latest[:280])}{"..." if len(latest) > 280 else ""}</div>' \
                    if latest else ''

        return f"""
<div class="action-card {level}">
    <div class="card-row">
        <div class="card-name">{esc(item['name'])}</div>
        {badge}
    </div>
    <div class="card-amount">{fmt_usd(item['usd'])} &nbsp;&#183;&nbsp; {fmt_ngn(item['ngn'])}</div>
    <div class="card-who">&#128100; {esc(item['accountable'] or 'Unassigned')} &nbsp;&#183;&nbsp; {esc(item['customer'])}</div>
    {act_html}
    {upd_html}
</div>"""

    action_cards_html = ''.join(action_card(i, 'critical') for i in critical_items) + \
                        ''.join(action_card(i, 'high')     for i in high_items)

    # ── Build movement card HTML ──────────────────────────────────────────────
    def movement_card(item):
        filled = [u for u in item['updates'] if u['text'].strip()]
        curr   = filled[-1] if filled else {'date': '', 'text': ''}
        prev   = filled[-2] if len(filled) >= 2 else {'date': '', 'text': ''}
        mv     = item['_movement']

        icon = '&#129513;' if mv == 'positive' else ('&#128308;' if mv == 'negative' else '&#128992;')
        prev_html = f'<div class="mv-lbl">Previous ({esc(prev["date"])})</div>' \
                    f'<div class="mv-prev">{esc(prev["text"][:250])}{"..." if len(prev["text"]) > 250 else ""}</div>' \
                    if prev['text'] else ''

        return f"""
<div class="mv-card {mv}">
    <div class="mv-name">{icon} {esc(item['name'])}</div>
    <div class="mv-customer">{esc(item['customer'])} &nbsp;&#183;&nbsp; {fmt_usd(item['usd'])}</div>
    {prev_html}
    <div class="mv-lbl">Latest ({esc(curr['date'])})</div>
    <div class="mv-curr">{esc(curr['text'][:280])}{"..." if len(curr['text']) > 280 else ""}</div>
</div>"""

    movement_html = ''.join(movement_card(i) for i in moved_items) if moved_items else \
        '<p style="color:var(--text-tertiary);font-size:13px">No significant movement detected between the last two weekly updates.</p>'

    # ── Build tracker table rows ──────────────────────────────────────────────
    def payment_badge(s):
        sl = s.lower()
        if sl == 'closed':
            return f'<span class="badge b-paid">{esc(s)}</span>'
        if 'in-progress' in sl:
            return f'<span class="badge b-progress">{esc(s)}</span>'
        return f'<span class="badge b-pending">{esc(s)}</span>'

    def deal_badge(s):
        sl = s.lower()
        if sl == 'closed':
            return f'<span class="badge b-paid">{esc(s)}</span>'
        if 'in-progress' in sl:
            return f'<span class="badge b-progress">{esc(s)}</span>'
        return f'<span class="badge b-pending">{esc(s)}</span>'

    table_rows = []
    for i_num, item in enumerate(items, 1):
        filled  = [u for u in item['updates'] if u['text'].strip()]
        latest  = filled[-1]['text'] if filled else '—'
        pred    = item['predictability']
        p_class = f'pred-{pred}' if pred in ('HIGH', 'MEDIUM', 'LOW') else ''
        data_p  = item['payment_status'].lower()
        action  = item['action'].strip() or '—'

        table_rows.append(f"""
<tr data-payment="{esc(data_p)}">
    <td style="color:var(--text-tertiary);font-size:11px">{esc(item['sn'] or str(i_num))}</td>
    <td>
        <div class="t-name">{esc(item['name'])}</div>
        <div class="t-cust">{esc(item['customer'])}</div>
    </td>
    <td><div class="t-usd">{fmt_usd(item['usd'])}</div></td>
    <td><div class="t-ngn">{fmt_ngn(item['ngn'])}</div></td>
    <td><span class="{p_class}">{esc(pred or '—')}</span></td>
    <td style="color:var(--text-tertiary);font-size:11px;white-space:nowrap">{esc(item['closure_period'])}</td>
    <td>{deal_badge(item['deal_status'])}</td>
    <td>{payment_badge(item['payment_status'])}</td>
    <td class="t-acc">{esc(item['accountable'])}</td>
    <td class="t-upd">{esc(latest[:300])}{"..." if len(latest) > 300 else ""}</td>
    <td class="t-act">{esc(action)}</td>
</tr>""")

    table_rows_html = '\n'.join(table_rows)

    mv_label = f"{prev_date} &#8594; {latest_date}" if prev_date else latest_date

    # ── Assemble final HTML ───────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Collections Intelligence &mdash; Seamfix 2026</title>
<style>{css}</style>
</head>
<body>
{toggle_html}

<!-- Header -->
<div class="page-header">
    <div class="page-title">&#128229; Collections Intelligence</div>
    <div class="page-sub">Critical Revenue Inflows 2026 &nbsp;&#183;&nbsp; Week of {esc(latest_date)} &nbsp;&#183;&nbsp; FX: $1 = &#8358;{FX_RATE:,} &nbsp;&#183;&nbsp; Generated {now}</div>
</div>

<!-- KPI Row -->
<div class="kpi-row">
    <div class="kpi-card">
        <div class="kpi-label">Total Tracked</div>
        <div class="kpi-val" style="color:var(--text-primary)">{fmt_usd(total_usd)}</div>
        <div class="kpi-sub">{fmt_ngn(total_ngn)} &nbsp;&#183;&nbsp; {len(items)} items</div>
    </div>
    <div class="kpi-card" style="border-top:3px solid #10b981">
        <div class="kpi-label">Fully Collected</div>
        <div class="kpi-val" style="color:#10b981">{fmt_usd(coll_usd)}</div>
        <div class="kpi-sub">{len(collected)} items &nbsp;&#183;&nbsp; {coll_pct}% of portfolio</div>
    </div>
    <div class="kpi-card" style="border-top:3px solid #3b82f6">
        <div class="kpi-label">Collection In Progress</div>
        <div class="kpi-val" style="color:#60a5fa">{fmt_usd(prog_usd)}</div>
        <div class="kpi-sub">{len(in_progress)} items &nbsp;&#183;&nbsp; {prog_pct}% of portfolio</div>
    </div>
    <div class="kpi-card" style="border-top:3px solid #f59e0b">
        <div class="kpi-label">Payment Pending</div>
        <div class="kpi-val" style="color:#f59e0b">{fmt_usd(pend_usd)}</div>
        <div class="kpi-sub">{len(pending)} items &nbsp;&#183;&nbsp; {pend_pct}% of portfolio</div>
    </div>
</div>

<!-- Critical Actions -->
<div class="sec-hdr">
    <span style="font-size:20px">&#128680;</span>
    <span class="sec-title">Critical Actions to Accelerate Collections</span>
    <span class="sec-sub">{len(critical_items)} critical &nbsp;&#183;&nbsp; {len(high_items)} need follow-up</span>
</div>
<div class="card-grid">
{action_cards_html}
</div>

<!-- What Moved This Week -->
<div class="sec-hdr">
    <span style="font-size:20px">&#128200;</span>
    <span class="sec-title">What Moved This Week</span>
    <span class="sec-sub">{mv_label}</span>
</div>
<div class="card-grid">
{movement_html}
</div>

<!-- Full Collections Tracker -->
<div class="sec-hdr">
    <span style="font-size:20px">&#128203;</span>
    <span class="sec-title">Full Collections Tracker</span>
    <span class="sec-sub">{len(items)} items &nbsp;&#183;&nbsp; {esc(latest_date)}</span>
</div>

<div class="filter-bar">
    <span style="color:var(--text-tertiary);font-size:12px">Filter by payment status:</span>
    <button class="f-btn active" onclick="filterTable(this,'all')">All ({len(items)})</button>
    <button class="f-btn" onclick="filterTable(this,'pending')">Pending ({len(pending)})</button>
    <button class="f-btn" onclick="filterTable(this,'in-progress')">In Progress ({len(in_progress)})</button>
    <button class="f-btn" onclick="filterTable(this,'closed')">Collected ({len(collected)})</button>
</div>

<div class="tbl-wrap">
<table id="cTable">
<thead>
<tr>
    <th>#</th>
    <th>Deal / Customer</th>
    <th>USD Amount</th>
    <th>NGN Amount</th>
    <th>Predictability</th>
    <th>Est. Closure</th>
    <th>Deal Status</th>
    <th>Payment Status</th>
    <th>Accountable</th>
    <th>Latest Update ({esc(latest_date)})</th>
    <th>Actionable Deliverable</th>
</tr>
</thead>
<tbody>
{table_rows_html}
</tbody>
</table>
</div>

<script>
function filterTable(btn, status) {{
    document.querySelectorAll('.f-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    btn.classList.add('active');
    document.querySelectorAll('#cTable tbody tr').forEach(function(row) {{
        var p = (row.getAttribute('data-payment') || '').toLowerCase();
        if (status === 'all') {{
            row.style.display = '';
        }} else if (status === 'in-progress') {{
            row.style.display = p.includes('in-progress') ? '' : 'none';
        }} else {{
            row.style.display = p.includes(status) ? '' : 'none';
        }}
    }});
}}

{theme_js}
</script>

</body>
</html>"""

    return html


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path('.')
    csv_path = folder / COLLECTIONS_FILENAME

    # Try alternate filename patterns if exact name not found
    if not csv_path.exists():
        for alt in sorted(folder.glob('*ollection*.csv')):
            csv_path = alt
            break

    if not csv_path.exists():
        # Write a placeholder so the dashboard tab doesn't crash
        placeholder = """<!DOCTYPE html><html><head>
<style>
body{background:#0a0f1e;color:#94a3b8;font-family:sans-serif;
     display:flex;align-items:center;justify-content:center;
     min-height:400px;flex-direction:column;gap:16px;text-align:center;}
h2{color:#f59e0b;} p{color:#475569;font-size:13px;}
</style></head><body>
<h2>&#9888; Collections data not yet available</h2>
<p>The Collections Tracker spreadsheet hasn't been fetched yet.</p>
<p>Click <strong>Regenerate Dashboards</strong> in the sidebar to refresh all data.</p>
</body></html>"""
        (folder / 'collections_dashboard.html').write_text(placeholder, encoding='utf-8')
        print("Collections file not found — placeholder written.", file=sys.stderr)
        return

    try:
        items = extract_collections(csv_path)
        update_dates = [u['date'] for u in items[0]['updates']] if items else []
        html = generate_html(items, update_dates)
        out  = folder / 'collections_dashboard.html'
        out.write_text(html, encoding='utf-8')
        print(f"Collections dashboard generated: {out} ({len(items)} items)")
    except Exception as e:
        import traceback
        print(f"Error generating collections dashboard: {e}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
