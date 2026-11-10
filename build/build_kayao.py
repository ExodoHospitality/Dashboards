"""Rebuild the Kayao dashboard from TEST_Kayao_v2.xlsx.
Usage: python TEST_build_dashboard.py TEST_Kayao_v2.xlsx TEST_dashboard_template.html TEST_Kayao_index.html
Reads ONLY the typed data tables (Lines, Invoices, Payments, Sales, Calendar, Watch, Owed!B2) and recomputes
everything in Python, so it works whether or not Excel has recalculated the formulas."""
import sys, json, datetime as dt, calendar
from decimal import Decimal, ROUND_HALF_UP
def xround(x, n=2):
    return float(Decimal(str(x)).quantize(Decimal(1).scaleb(-n), rounding=ROUND_HALF_UP))
from collections import defaultdict
import openpyxl

src, tpl, out = sys.argv[1:4]
wb = openpyxl.load_workbook(src, data_only=True)

def rows(name, ncols):
    ws = wb[name]
    hdr = [c.value for c in ws[1][:ncols]]
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None and r[1] is None: continue
        yield dict(zip(hdr, r[:ncols]))

def D(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    return None

cal = [r for r in rows('Calendar', 5) if r['Week']]
cal.sort(key=lambda r: D(r['Start']))
def week_of(date):
    for r in cal:
        if D(r['Start']) <= date <= D(r['End']): return r['Week'], r['Month']
    return None, None

lines = list(rows('Lines', 17))
invs = list(rows('Invoices', 14))
pays = list(rows('Payments', 8))
sales = list(rows('Sales', 4))
watch = list(rows('Watch', 6))
asof = D(wb['Owed']['B2'].value) or dt.date.today()

W = [r['Week'] for r in cal]
wi = {w: i for i, w in enumerate(W)}
cats = {c: [0.0] * len(W) for c in ('Food', 'Bar', 'Supply', 'Cleaning')}
SUBMAP = {'Cocktails & Liquor': 'Cocktails & liquor', 'Beer': 'Beer', 'Wine': 'Wine', 'Non-Alcoholic': 'Non-alcoholic'}
subs = {v: [0.0] * len(W) for v in SUBMAP.values()}
vend = defaultdict(lambda: [0.0] * len(W))
prod_by_inv = defaultdict(float); cc_by_inv = defaultdict(float)
for l in lines:
    if not l['Invoice #']: continue
    ext = float((Decimal(str(l['Qty received'] or 0)) * Decimal(str(l['Unit price'] or 0))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
    key = (l['Vendor'], str(l['Invoice #']))
    prod_by_inv[key] += ext
    if l['Charge to'] == 'Cha Chin': cc_by_inv[key] += ext
    w, _ = week_of(D(l['Date']))
    if w is None or l['Charge to'] != 'Kayao': continue
    i = wi[w]
    if l['Category'] in cats: cats[l['Category']][i] += ext
    if l['Category'] == 'Bar' and l['Bar Subcategory'] in SUBMAP: subs[SUBMAP[l['Bar Subcategory']]][i] += ext
    vend[l['Vendor']][i] += ext

paid_by_inv = defaultdict(float); paydate = {}; acct = {}
for p in pays:
    if not p['Invoice #']: continue
    key = (p['Vendor'], str(p['Invoice #']))
    paid_by_inv[key] += p['Amount'] or 0
    if D(p['Date paid']): paydate[key] = max(paydate.get(key, D(p['Date paid'])), D(p['Date paid']))
    if p['Account month']: acct[key] = p['Account month']

fees = [0.0] * len(W); wk_inv = [0] * len(W); out_inv = []
for iv in invs:
    if not iv['Invoice #']: continue
    key = (iv['Vendor'], str(iv['Invoice #']))
    date = D(iv['Date']); w, _ = week_of(date)
    total = round(iv['Printed Total'] or 0, 2); f = (iv['Freight & Fees'] or 0) + (iv['Sales Tax'] or 0)
    prod = round(prod_by_inv.get(key, 0), 2); ccp = round(cc_by_inv.get(key, 0), 2)
    paid = abs(paid_by_inv.get(key, 0) - total) < 0.005
    pd = paydate.get(key)
    pmonth = acct.get(key) or (pd.strftime('%Y-%m') if pd else '')
    if paid and pd is None and key in acct:
        y, m = map(int, acct[key].split('-')); pd = dt.date(y, m, calendar.monthrange(y, m)[1])
    if w is not None:
        fees[wi[w]] += f; wk_inv[wi[w]] += 1
    out_inv.append(dict(vendor=iv['Vendor'], number=str(iv['Invoice #']), date=date.isoformat(), week=w or '',
                        product=round(prod_by_inv.get(key, 0), 2), fees=round(f, 2), total=total,
                        status='paid' if paid else 'open', paidOn=pd.isoformat() if pd else None,
                        due=D(iv['Due date']).isoformat() if D(iv['Due date']) else None,
                        note=(iv['Terms / notes'] or ''), paidBy=(iv['Paid by'] or 'Kayao'), paidMonth=pmonth if paid else '',
                        chargeTo=('Cha Chin' if ccp and abs(ccp - prod) < 0.005 else 'Split' if ccp else 'Kayao'), owedCC=round(total if (ccp and abs(ccp - prod) < 0.005) else ccp, 2)))
out_inv.sort(key=lambda x: (x['date'], x['vendor'], x['number']))

SALEMAP = {'Food': 'Food', 'Cocktails & Liquor': 'Cocktails & liquor', 'Beer': 'Beer', 'Wine': 'Wine', 'Non-Alcoholic': 'Non-alcoholic'}
sal = {v: [0.0] * len(W) for v in SALEMAP.values()}
for s in sales:
    if s['Week'] in wi and s['Category'] in SALEMAP: sal[SALEMAP[s['Category']]][wi[s['Week']]] += s['Net sales'] or 0

vendors = sorted(({'name': k, 'weeks': [round(x, 2) for x in v]} for k, v in vend.items()), key=lambda v: -sum(v['weeks']))
weeks = [dict(id=r['Week'], start=D(r['Start']).strftime('%m/%d'), end=D(r['End']).strftime('%m/%d'), month=r['Month'],
              entered=wk_inv[i] > 0, invoices=wk_inv[i]) for i, r in enumerate(cal)]
months = []
for r in cal:
    if not months or months[-1]['name'] != r['Month']:
        months.append(dict(name=r['Month'], key=r['Month key'], weeks=[], start=D(r['Start']).isoformat()))
    months[-1]['weeks'].append(r['Week']); months[-1]['end'] = D(r['End']).isoformat()
# carry: previous month -> current month (by the as-of date)
cur_key = asof.strftime('%Y-%m')
idx = next((i for i, m in enumerate(months) if m['key'] == cur_key), None)
if idx is None or idx == 0:
    prev = months[-1] if idx is None else None
    cur_name = asof.strftime('%B'); cur_key2 = cur_key
else:
    prev = months[idx - 1]; cur_name = months[idx]['name']; cur_key2 = months[idx]['key']
carry = {}
if prev:
    y, m = map(int, cur_key2.split('-'))
    carry = dict(**{'from': prev['name'], 'to': cur_name, 'start': prev['start'], 'end': prev['end'], 'weeks': prev['weeks'][0] + '\u2013' + prev['weeks'][-1], 'toStart': dt.date(y, m, 1).isoformat()})
# Cha Chin
def amounts(sheet):
    if sheet not in wb.sheetnames: return 0.0
    return sum((r[1] or 0) for r in wb[sheet].iter_rows(min_row=5, values_only=True) if isinstance(r[1], (int, float)))
received = amounts('Cha Chin')
ay_owed = round(sum(i['total'] for i in out_inv if i['paidBy'] == 'Ayayay'), 2)
ay_open = round(sum(i['total'] for i in out_inv if i['paidBy'] == 'Ayayay' and i['status'] == 'open'), 2)
ay_repaid = round(amounts('Ayayay'), 2)
ayayay = dict(owed=ay_owed, vendorOpen=ay_open, repaid=ay_repaid, balance=round(ay_owed - ay_repaid, 2))
owed_cc = round(sum(i['owedCC'] for i in out_inv), 2)
open_owed = round(sum(i['owedCC'] for i in out_inv if i['status'] == 'open'), 2)
pm_key = str(wb['Owed']['B3'].value or '').strip()
try:
    pm_label = dt.date(int(pm_key[:4]), int(pm_key[5:7]), 1).strftime('%B %Y')
except Exception:
    pm_label = pm_key
pay_month = dict(key=pm_key, label=pm_label)
chachin = dict(owed=owed_cc, received=round(received, 2), balance=round(owed_cc - received, 2), openOwed=open_owed)
watch_out = [dict(tag=w['Tag'] or 'Open', title=str(w['Vendor / topic'] or ''), body=str(w['Item'] or ''))
             for w in watch if w['Item'] and str(w['Status']).lower() in ('open', 'to confirm')]
data = dict(asOf=asof.isoformat(), weeks=weeks, categories={k: [round(x, 2) for x in v] for k, v in cats.items()},
            barSub={k: [round(x, 2) for x in v] for k, v in subs.items()}, fees=[round(x, 2) for x in fees], vendors=vendors,
            partial=['A1'], partialNote={'A1': '08/01–08/02 only: July invoices not loaded'}, months=months, carry=carry, chaChin=chachin, ayayay=ayayay, payMonth=pay_month,
            source=src.split('/')[-1].split('\\')[-1], built=dt.datetime.now().strftime('%m/%d/%Y %H:%M'),
            sales={k: [round(x, 2) for x in v] for k, v in sal.items()}, invoices=out_inv, watch=watch_out)
html = open(tpl, encoding='utf-8').read().replace('/*KAYAO_DATA*/', json.dumps(data, ensure_ascii=False))
open(out, 'w', encoding='utf-8').write(html)
print(f'wrote {out}: {len(out_inv)} invoices, {len(watch_out)} watch items, as of {asof}')
