#!/usr/bin/env python3
"""
TEST_build_weekly_report.py  -  builds the Exodo Hospitality Weekly Report
from the weekly sales log.  Nothing is typed into code: the Excel is the
only source of numbers.

    python TEST_build_weekly_report.py Exodo_weekly_sales_log.xlsx TEST_Weekly-Report_template.html TEST_Weekly-Report.html

Optional 4th argument:  test  (default)  or  final   -> which invoice dashboards the buttons open.

Rules:
  * reads the workbook's SAVED values (what Excel last calculated).  A formula
    cell that Excel has not calculated yet comes out as "no data" and is listed
    in the Checks box - the number is never guessed.
  * a blank cell is "no data" (null), never 0.
  * the Result column is copied exactly as typed.
"""
import sys, json, re, datetime as dt
from pathlib import Path
import openpyxl

# ============================================================================
# SETTINGS  -  the only part you ever change.
#   key      : short id used inside the page (letters only, lowercase)
#   sheet    : the tab name in the Excel (the "<Loc> Food" / "<Loc> Bar" tabs are optional)
#   name     : how the location is shown on the page
#   matilda  : True only for the two-venue layout (Matilda / Clandestino columns)
#   link_test / link_final : what the "Invoice processing" button opens
# To add a restaurant: add one more {...} line.  Nothing else.
# ============================================================================
LOCATIONS = [
    {"key": "akiro",   "sheet": "Akiro Hand Roll Bar",   "name": "Akiro Hand Roll Bar",   "matilda": False,
     "food": "Akiro Food",   "bar": "Akiro Bar",   "button": "Akiro",   "link_test": "TEST_index.html",         "link_final": "akiro/index.html"},
    {"key": "kayao",   "sheet": "Kayao Restaurant",      "name": "Kayao Restaurant",      "matilda": False,
     "food": "Kayao Food",   "bar": "Kayao Bar",   "button": "Kayao",   "link_test": "TEST_Kayao_index.html",   "link_final": "Kayao/index.html"},
    {"key": "matilda", "sheet": "Matilda & Clandestino", "name": "Matilda | Clandestino", "matilda": True,
     "food": "Matilda Food", "bar": "Matilda Bar", "button": "Matilda", "link_test": "TEST_Matilda_index.html", "link_final": "Matilda/index.html"},
    {"key": "ayayay",  "sheet": "Ayayay Mexican Eatery", "name": "Ayayay Mexican Eatery", "matilda": False,
     "food": "Ayayay Food",  "bar": "Ayayay Bar",  "button": "Ayayay",  "link_test": "TEST_Ayayay_index.html",  "link_final": "Ayayay/index.html"},
]
# ============================================================================

MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def num(v):
    """Same reading rule as the page: blank -> None, '$1,234' -> 1234.0"""
    if v is None or v == '':
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r'[$,\s%]', '', str(v))
    try:
        return float(s)
    except ValueError:
        return None

def date_parts(v):
    if v is None or v == '':
        return None
    if isinstance(v, (dt.datetime, dt.date)):
        return (v.year, v.month, v.day)
    if isinstance(v, (int, float)):
        if v > 1:
            d = dt.date(1899, 12, 30) + dt.timedelta(days=int(round(v)))
            return (d.year, d.month, d.day)
        return None
    s = str(v).strip()
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m: return (int(m[1]), int(m[2]), int(m[3]))
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m: return (int(m[3]), int(m[1]), int(m[2]))
    return None

def iso(p): return f"{p[0]:04d}-{p[1]:02d}-{p[2]:02d}"
def div(a, b):
    if a is None or b is None or b == 0: return None
    return a / b
def nsum(*vals):
    vs = [v for v in vals if v is not None]
    return sum(vs) if vs else None

def find_sheet(wb, wanted):
    if wanted in wb.sheetnames: return wb[wanted]
    norm = lambda s: re.sub(r'[^a-z0-9]', '', str(s).lower())
    for n in wb.sheetnames:
        if norm(n) == norm(wanted): return wb[n]
    return None

def sheet_rows(ws):
    rows = []
    for r in ws.iter_rows(values_only=True):
        if all(c is None for c in r): continue
        rows.append(list(r))
    return rows

# --------------------------------------------------------------------------
def parse_location(ws, ws_f, is_matilda, checks, locname):
    rows_v = sheet_rows(ws)                      # saved values
    rows_f = sheet_rows(ws_f)                    # formulas (to spot uncalculated cells)
    if not rows_v: raise ValueError('sheet is empty')
    hdr = [('' if h is None else str(h).strip()) for h in rows_v[0]]
    idx = {}
    for i, h in enumerate(hdr):
        if h and h not in idx: idx[h] = i
    if 'Start date' not in idx: raise ValueError('no "Start date" column found')

    def g(row, *names):
        for n in names:
            if n in idx:
                return num(row[idx[n]]) if idx[n] < len(row) else None
        return None

    weekly = []
    for r in range(1, len(rows_v)):
        row = rows_v[r]
        frow = rows_f[r] if r < len(rows_f) else []
        sp = date_parts(row[idx['Start date']])
        if not sp: continue
        ep = (date_parts(row[idx['End date']]) if 'End date' in idx else None) or sp
        wk = iso(sp)

        # uncalculated formula cells -> listed, never guessed
        for ci, h in enumerate(hdr):
            if ci < len(frow) and isinstance(frow[ci], str) and frow[ci].startswith('=') \
               and (ci >= len(row) or row[ci] is None):
                checks.append(f"{locname} week {wk}: column \"{h}\" is a formula Excel has not calculated - open the workbook in Excel, press F9, save, rebuild.")

        net = g(row, 'Net sales'); fs = g(row, 'Food sales'); bs = g(row, 'Bar sales')
        w = {
            'start': wk, 'end': iso(ep),
            'label': f"{MONTHS[sp[1]-1]} {sp[2]:02d}",
            'month_key': f"{sp[0]:04d}-{sp[1]:02d}",
            'month_label': f"{MONTHS[sp[1]-1]} {sp[0]}",
            'gross_sales': g(row, 'Gross sales'),
            'discounts': g(row, 'Discounts'),
            'discount_pct': g(row, 'Discount %'),
            'refunds': g(row, 'Refunds'),
            'net_sales': net,
            'voids': g(row, 'Voids'),
            'dinner_covers': g(row, 'Dinner covers'),
            'dinner_check': g(row, 'Dinner Check Ave', 'Dinner Ave. Check'),
            'brunch_covers': g(row, 'Brunch covers'),
            'brunch_check': g(row, 'Brunch Check Ave', 'Brunch Ave. Check'),
            'food_sales': fs,
            'food_cost': g(row, 'Food cost'),
            'food_cogs_pct': g(row, 'Food COGS'),
            'bar_sales': bs,
            'bar_cost': g(row, 'Bar cost'),
            'bar_cogs_pct': g(row, 'Bar COGS'),
            'prime_cost': g(row, 'Prime cost'),
            'prime_pct': g(row, '%'),
            'food_sales_pct': div(fs, net),
            'bar_sales_pct': div(bs, net),
        }
        # a % computed from a blank cost or blank sales is meaningless -> "no data", not 0.0%
        for pk, ck, sk in (('food_cogs_pct','food_cost','food_sales'), ('bar_cogs_pct','bar_cost','bar_sales'),
                           ('prime_pct','prime_cost','net_sales'), ('discount_pct','discounts','gross_sales')):
            if w[ck] is None or w[sk] is None: w[pk] = None
        res = row[idx['Result']] if 'Result' in idx and idx['Result'] < len(row) else None
        w['result'] = None if res is None or str(res).strip() == '' else str(res).strip()

        if is_matilda:
            w['matilda_sales']     = g(row, 'Matilda sales')
            w['clandestino_sales'] = g(row, 'Clandestino Sales')
            w['mat_boh_cost'] = g(row, 'MAT BOH cost');      w['mat_boh_pct'] = g(row, 'MAT BOH')
            w['mat_foh_cost'] = g(row, 'MAT FOH cost');      w['mat_foh_pct'] = g(row, 'MAT FOH')
            w['cld_foh_cost'] = g(row, 'CLD FOH cost');      w['cld_foh_pct'] = g(row, 'CLD FOH')
            w['mat_salaried_cost'] = g(row, 'MAT Salaried cost'); w['mat_salaried_pct'] = g(row, 'MAT Salaried')
            w['cld_salaried_cost'] = g(row, 'CLD Salaried cost'); w['cld_salaried_pct'] = g(row, 'CLD Salaried')
            w['boh_cost'] = w['mat_boh_cost'];                      w['boh_pct'] = div(w['boh_cost'], net)
            w['foh_cost'] = nsum(w['mat_foh_cost'], w['cld_foh_cost']); w['foh_pct'] = div(w['foh_cost'], net)
            w['salaried_cost'] = nsum(w['mat_salaried_cost'], w['cld_salaried_cost']); w['salaried_pct'] = div(w['salaried_cost'], net)
        else:
            w['boh_cost'] = g(row, 'BOH cost');           w['boh_pct'] = g(row, 'BOH')
            w['foh_cost'] = g(row, 'FOH cost');           w['foh_pct'] = g(row, 'FOH')
            w['salaried_cost'] = g(row, 'Salaried cost'); w['salaried_pct'] = g(row, 'Salaried')

        # --- sanity checks (reported, never corrected) ---
        core = ['gross_sales', 'net_sales', 'food_sales', 'food_cost', 'bar_sales', 'bar_cost', 'prime_cost']
        missing = [k for k in core if w[k] is None]
        if len(missing) == len(core):
            checks.append(f"{locname} week {wk}: dated row with no numbers - shown as \"no data\".")
        elif missing:
            note = " Prime cost on that row leaves the blank cost out, so it is understated." if any(k in missing for k in ('food_cost','bar_cost')) else ""
            checks.append(f"{locname} week {wk}: blank cells: {', '.join(missing)} - shown as \"no data\".{note}")
        for pk, ck, sk, lab in [('food_cogs_pct','food_cost','food_sales','Food COGS'),
                                ('bar_cogs_pct','bar_cost','bar_sales','Bar COGS'),
                                ('prime_pct','prime_cost','net_sales','Prime %')]:
            typed, calc = w[pk], div(w[ck], w[sk])
            if typed is not None and calc is not None and abs(typed - calc) > 0.0005:
                checks.append(f"{locname} week {wk}: {lab} cell says {typed*100:.1f}% but cost ÷ sales = {calc*100:.1f}% - the formula cell looks stale (F9 in Excel).")
        if w['result'] is None and w['net_sales'] is not None:
            checks.append(f"{locname} week {wk}: Result column is empty.")
        weekly.append(w)

    if not weekly: raise ValueError('no dated rows found')
    weekly.sort(key=lambda w: w['start'])
    # duplicate / gap check
    prev = None
    for w in weekly:
        if prev:
            if w['start'] == prev['start']:
                checks.append(f"{locname}: week {w['start']} appears twice.")
            else:
                exp = dt.date.fromisoformat(prev['end']) + dt.timedelta(days=1)
                if dt.date.fromisoformat(w['start']) != exp:
                    checks.append(f"{locname}: gap between week ending {prev['end']} and week starting {w['start']}.")
        prev = w
    return weekly

SUM_KEYS = ['gross_sales','discounts','net_sales','food_sales','food_cost','bar_sales','bar_cost',
            'boh_cost','foh_cost','salaried_cost','prime_cost','dinner_covers','brunch_covers']
MAT_SUM_KEYS = ['mat_boh_cost','mat_foh_cost','cld_foh_cost','mat_salaried_cost','cld_salaried_cost']

def rollup_months(weekly, is_matilda):
    buckets, order = {}, []
    for w in weekly:
        if w['month_key'] not in buckets: buckets[w['month_key']] = []; order.append(w['month_key'])
        buckets[w['month_key']].append(w)
    out = []
    for key in order:
        ws = buckets[key]
        m = {'month_key': key, 'month_label': ws[0]['month_label'], 'weeks_count': len(ws),
             'weeks_missing': sum(1 for w in ws if w['net_sales'] is None)}
        for k in SUM_KEYS: m[k] = nsum(*[w.get(k) for w in ws])
        if is_matilda:
            for k in MAT_SUM_KEYS: m[k] = nsum(*[w.get(k) for w in ws])
            ms = nsum(*[w.get('matilda_sales') for w in ws]); cs = nsum(*[w.get('clandestino_sales') for w in ws])
            m['mat_boh_pct'] = div(m['mat_boh_cost'], ms); m['mat_foh_pct'] = div(m['mat_foh_cost'], ms)
            m['cld_foh_pct'] = div(m['cld_foh_cost'], cs)
            m['mat_salaried_pct'] = div(m['mat_salaried_cost'], ms); m['cld_salaried_pct'] = div(m['cld_salaried_cost'], cs)
        net = m['net_sales']
        m['discount_pct'] = div(m['discounts'], m['gross_sales'])
        m['food_cogs_pct'] = div(m['food_cost'], m['food_sales']); m['bar_cogs_pct'] = div(m['bar_cost'], m['bar_sales'])
        m['boh_pct'] = div(m['boh_cost'], net); m['foh_pct'] = div(m['foh_cost'], net)
        m['salaried_pct'] = div(m['salaried_cost'], net); m['prime_pct'] = div(m['prime_cost'], net)
        m['food_sales_pct'] = div(m['food_sales'], net); m['bar_sales_pct'] = div(m['bar_sales'], net)
        dc = nsum(*[(w['dinner_covers'] or 0) * w['dinner_check'] for w in ws if w['dinner_check'] is not None])
        bc = nsum(*[(w['brunch_covers'] or 0) * w['brunch_check'] for w in ws if w['brunch_check'] is not None])
        m['dinner_check_avg'] = div(dc, m['dinner_covers']); m['brunch_check_avg'] = div(bc, m['brunch_covers'])
        out.append(m)
    return out

# ---- vendor / category detail tabs (same layout rule as the page) ----------
def detail_columns(hdr):
    items, i = [], 5
    while i < len(hdr):
        h = hdr[i]
        if not h or h == 'Total %' or h.endswith('%'):
            i += 1; continue
        nxt = hdr[i+1] if i+1 < len(hdr) else ''
        if nxt.endswith('Purchases'):
            items.append({'name': h, 'salesIdx': i, 'costIdx': i+1}); i += 3
        else:
            items.append({'name': re.sub(r' Purchases$', '', h), 'salesIdx': None, 'costIdx': i}); i += 2
    return items

def parse_detail(ws, kind):
    rows = sheet_rows(ws)
    if not rows: return None
    hdr = [('' if h is None else str(h).strip()) for h in rows[0]]
    idx = {}
    for i, h in enumerate(hdr):
        if h and h not in idx: idx[h] = i
    if 'Start date' not in idx: return None
    cols = detail_columns(hdr)
    if not cols: return None
    sales_col = 'Food sales' if kind == 'food' else 'Bar sales'
    purch_col = 'Food Purchases' if kind == 'food' else 'Bar Purchases'
    cell = lambda row, i: (num(row[i]) if i is not None and i < len(row) else None)
    out = []
    for row in rows[1:]:
        sp = date_parts(row[idx['Start date']])
        if not sp: continue
        ep = (date_parts(row[idx['End date']]) if 'End date' in idx else None) or sp
        items = []
        for c in cols:
            cost = cell(row, c['costIdx']) or 0.0
            sales = cell(row, c['salesIdx']) if c['salesIdx'] is not None else None
            if cost == 0 and (sales is None or sales == 0): continue
            items.append({'name': c['name'], 'cost': cost, 'sales': sales})
        if not items: continue
        out.append({'start': iso(sp), 'end': iso(ep), 'label': f"{MONTHS[sp[1]-1]} {sp[2]:02d}",
                    'sales': (cell(row, idx[sales_col]) or 0.0) if sales_col in idx else 0.0,
                    'purchases': (cell(row, idx[purch_col]) or 0.0) if purch_col in idx else 0.0,
                    'items': items})
    if not out: return None
    out.sort(key=lambda r: r['start'])
    names = []
    for r in out:
        for it in r['items']:
            if it['name'] not in names: names.append(it['name'])
    return {'kind': kind, 'names': names, 'rows': out}

# --------------------------------------------------------------------------
def main():
    if len(sys.argv) < 4:
        print(__doc__); sys.exit(1)
    xlsx, template, outfile = sys.argv[1], sys.argv[2], sys.argv[3]
    link_mode = sys.argv[4].lower() if len(sys.argv) > 4 else 'test'
    if link_mode not in ('test', 'final'):
        sys.exit("4th argument must be 'test' or 'final'")

    wb_v = openpyxl.load_workbook(xlsx, data_only=True)   # saved values
    wb_f = openpyxl.load_workbook(xlsx, data_only=False)  # formulas
    checks, locations = [], {}
    for cfg in LOCATIONS:
        ws = find_sheet(wb_v, cfg['sheet'])
        if ws is None:
            checks.append(f"Tab \"{cfg['sheet']}\" not found in the workbook - {cfg['name']} left out."); continue
        ws_f = wb_f[ws.title]
        try:
            weekly = parse_location(ws, ws_f, cfg['matilda'], checks, cfg['name'])
        except ValueError as e:
            checks.append(f"Tab \"{cfg['sheet']}\": {e} - {cfg['name']} left out."); continue
        detail = {}
        for tab, kind in ((cfg['food'], 'food'), (cfg['bar'], 'bar')):
            dws = find_sheet(wb_v, tab)
            if dws is None: continue
            parsed = parse_detail(dws, kind)
            if parsed: detail[kind] = parsed
        locations[cfg['key']] = {'name': cfg['name'], 'weekly': weekly,
                                 'monthly': rollup_months(weekly, cfg['matilda']), 'detail': detail}
    if not locations: sys.exit("No location tabs could be read.")

    built_through = max(l['weekly'][-1]['end'] for l in locations.values())
    sheet_map = [{'key': c['key'], 'sheet': c['sheet'], 'name': c['name'], 'matilda': c['matilda'],
                  'food': c['food'], 'bar': c['bar']} for c in LOCATIONS]
    buttons = '\n'.join(
        f'        <a class="invoice-btn" href="{c["link_test" if link_mode=="test" else "link_final"]}">{c["button"]} →</a>'
        for c in LOCATIONS)

    html = Path(template).read_text(encoding='utf-8')
    def put(tag, value):
        nonlocal html
        if html.count(tag) != 1: sys.exit(f"template placeholder {tag} found {html.count(tag)} times (expected 1)")
        html = html.replace(tag, value)
    put('__ALL_DATA__', json.dumps({'locations': locations}, separators=(',', ':')))
    put('__BUILT_FROM__', Path(xlsx).name)
    put('__BUILT_THROUGH__', built_through)
    put('__BUILT_ON__', dt.date.today().isoformat())
    put('__SHEET_MAP__', json.dumps(sheet_map))
    put('__INVOICE_BUTTONS__', buttons)
    put('__BUILD_CHECKS__', json.dumps(checks))
    put('__LINK_MODE__', link_mode)
    Path(outfile).write_text(html, encoding='utf-8')

    print(f"Built {outfile}  (data through {built_through}, links: {link_mode})")
    for k, l in locations.items():
        print(f"  {l['name']}: {len(l['weekly'])} weeks, {l['weekly'][0]['start']} -> {l['weekly'][-1]['end']}"
              + (f", vendor detail to {l['detail']['food']['rows'][-1]['end']}" if 'food' in l['detail'] else ', no vendor detail tabs'))
    print(f"Checks: {len(checks)}")
    for c in checks: print('  - ' + c)

if __name__ == '__main__':
    main()
