"""TEST_build_dashboard.py - rebuild the Matilda dashboard from TEST_Matilda_v2.xlsx.
Usage: python TEST_build_dashboard.py TEST_Matilda_v2.xlsx TEST_dashboard_template_matilda.html TEST_Matilda_index.html
Reads ONLY the typed tables (Lines, Invoices, Payments, Sales, Calendar, Vendors, Watch) and Owed!B2/B3,
so it does not matter whether Excel has recalculated. The dashboard never contains anything that is not in the workbook."""
import sys, json, datetime, collections
import openpyxl

src, tpl, out = (sys.argv + ['TEST_Matilda_v2.xlsx', 'TEST_dashboard_template_matilda.html', 'TEST_Matilda_index.html'])[1:4]
wb = openpyxl.load_workbook(src, data_only=True)
RESTAURANT = 'Matilda'

def rows(ws, ncols):
    out = []
    for r in ws.iter_rows(min_row=2, max_col=ncols, values_only=True):
        if r[0] is None and r[1] is None: continue
        out.append(r)
    return out

def iso(d): return d.strftime('%Y-%m-%d') if isinstance(d, (datetime.datetime, datetime.date)) else (d or None)
def md(d): return d.strftime('%m/%d') if isinstance(d, (datetime.datetime, datetime.date)) else ''
def num(v): return float(v or 0)

# Calendar
cal = [r for r in rows(wb['Calendar'], 5) if r[0]]
weeks = [dict(id=r[0], startIso=iso(r[1]), endIso=iso(r[2]), month=r[3], key=r[4], start=md(r[1]), end=md(r[2])) for r in cal]
wk_index = {w['id']: i for i, w in enumerate(weeks)}
def week_of(d):
    s = iso(d)
    for w in weeks:
        if w['startIso'] <= s <= w['endIso']: return w['id']
    return None
months_list = []
for r in wb['Calendar'].iter_rows(min_row=2, min_col=9, max_col=10, values_only=True):
    if r[0]: months_list.append((r[0], r[1]))
months = []
for label, key in months_list:
    ws_ = [w for w in weeks if w['key'] == key]
    months.append(dict(key=key, label=label, short=label[:3], start=ws_[0]['start'] if ws_ else '', end=ws_[-1]['end'] if ws_ else '', startIso=ws_[0]['startIso'] if ws_ else '', endIso=ws_[-1]['endIso'] if ws_ else ''))
def month_key_of_date(d):
    w = week_of(d)
    return next((x['key'] for x in weeks if x['id'] == w), None)

# Lines: Paid by, Date, Week, Month, Vendor, Invoice #, Description, Category, Bar Subcategory, Qty ordered, Qty received, Unit, Unit price, Extended(formula), Invoice found?, Notes, Key
lines = []
for r in rows(wb['Lines'], 17):
    if not r[1]: continue
    ext = round(num(r[10]) * num(r[12]), 2)
    lines.append(dict(date=iso(r[1]), week=week_of(r[1]), vendor=r[4], inv=str(r[5]), cat=r[7], sub=r[8], ext=ext))
CATS = ['Food', 'Bar', 'Supply', 'Cleaning', 'Family Meal']; SUBS = ['Cocktails & liquor', 'Beer', 'Wine', 'Non-alcoholic']
nW = len(weeks)
categories = {c: [0.0] * nW for c in CATS}; barSub = {s: [0.0] * nW for s in SUBS}
product_by_key = collections.defaultdict(float); vend = collections.defaultdict(lambda: [0.0] * nW)
for ln in lines:
    i = wk_index.get(ln['week'])
    if i is None: continue
    categories[ln['cat']][i] += ln['ext']; product_by_key[(ln['vendor'], ln['inv'])] += ln['ext']; vend[ln['vendor']][i] += ln['ext']
    if ln['cat'] == 'Bar' and ln['sub'] in barSub: barSub[ln['sub']][i] += ln['ext']
for d in (categories, barSub):
    for k in d: d[k] = [round(x, 2) for x in d[k]]

# Payments: Paid by, Date paid, Vendor, Invoice #, Amount, Method, Evidence, Account month
pay = collections.defaultdict(list)
for r in rows(wb['Payments'], 8):
    if not r[2]: continue
    pay[(r[2], str(r[3]) if r[3] is not None else '')].append(dict(date=iso(r[1]), amount=round(num(r[4]), 2), method=r[5] or '', evidence=r[6] or '', month=r[7] or ''))

# Invoices: Paid by, Vendor, Invoice #, Date, Week, Month, Product, Fees, Tax, InvTotal, Printed, Foots, Terms, Due
invoices = []; fees = [0.0] * nW; inv_count = [0] * nW
for r in rows(wb['Invoices'], 14):
    if not r[1]: continue
    key = (r[1], str(r[2])); wk = week_of(r[3]); i = wk_index.get(wk)
    printed = round(num(r[10]), 2); ps = pay.get(key, [])
    paid_amt = round(sum(p['amount'] for p in ps), 2)
    status = 'open' if not ps else ('paid' if abs(paid_amt - printed) < 0.005 else 'partial')
    paidOn = max((p['date'] for p in ps if p['date']), default=None) if status == 'paid' else None
    acct = next((p['month'] for p in ps if p['month']), '') if status == 'paid' else ''
    paidMonth = acct or (paidOn[:7] if paidOn else '')
    credit = status == 'paid' and any(p['method'] == 'Credit on account' for p in ps)
    ev = '; '.join(sorted(set(p['evidence'] for p in ps if p['evidence'])))
    note = (r[12] or '') + ((' | ' + ev) if ev else '')
    invoices.append(dict(vendor=r[1], number=str(r[2]), date=iso(r[3]), week=wk, product=round(product_by_key.get(key, 0.0), 2), fees=round(num(r[7]) + num(r[8]), 2), total=printed,
                         status=status, paidOn=paidOn, paidMonth=paidMonth, credit=credit, due=iso(r[13]), note=note, method=ps[0]['method'] if ps else ''))
    if i is not None: fees[i] += num(r[7]) + num(r[8]); inv_count[i] += 1
fees = [round(x, 2) for x in fees]

# Sales: Paid by, Week, Venue, Category, Net sales
sales = {}; venueSales = {}
for r in rows(wb['Sales'], 5):
    if not r[1]: continue
    i = wk_index.get(r[1])
    if i is None: continue
    sales.setdefault(r[3], [0.0] * nW)[i] += num(r[4]); venueSales.setdefault(r[2], [0.0] * nW)[i] += num(r[4])
for d in (sales, venueSales):
    for k in d: d[k] = [round(x, 2) for x in d[k]]

# Vendors (ordered by total product spend)
vnames = [r[0] for r in rows(wb['Vendors'], 2) if r[0]]
vendors = sorted([dict(name=v, weeks=[round(x, 2) for x in vend.get(v, [0.0] * nW)]) for v in vnames], key=lambda v: -sum(v['weeks']))

# Watch: Date, Tag, Vendor/topic, Item, Amount, Status
watch = []
for r in rows(wb['Watch'], 6):
    if not r[1] or (r[5] or '').lower() == 'closed': continue
    watch.append(dict(tag=r[1], title=(r[2] or '') + (f' · ${num(r[4]):,.2f}' if r[4] else ''), body=r[3] or ''))

owed = wb['Owed']; asOf = iso(owed['B2'].value); payKey = str(owed['B3'].value)
y, m = int(payKey[:4]), int(payKey[5:7]); ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
labels = {k: l for l, k in months_list}
nextKey = f'{ny}-{nm:02d}'; nextLabel = labels.get(nextKey) or datetime.date(ny, nm, 1).strftime('%B')
payMonth = dict(key=payKey, label=labels.get(payKey, payKey), nextKey=nextKey, nextLabel=nextLabel)

entered_weeks = [dict(id=w['id'], start=w['start'], end=w['end'], month=w['month'], entered=inv_count[i] > 0, invoices=inv_count[i]) for i, w in enumerate(weeks)]
data = dict(asOf=asOf, weeks=entered_weeks, categories=categories, barSub=barSub, fees=fees, vendors=vendors,
            partial=['A1'], partialNote={'A1': '07/31–08/02 only: earlier A1 invoices not loaded'},
            sales=sales, venueSales=venueSales, invoices=invoices, watch=watch, months=months, payMonth=payMonth,
            footNote=f'Built from {src.split("/")[-1]} (TEST design). Weeks run Monday to Sunday by invoice date; A1 (07/27–08/02) covers 07/31–08/02 only and stays in August. A week that crosses a month end belongs to the month with more of its days. "Paid in {payMonth["label"]}" follows the Account month typed on the Payments sheet when the owner decided a payment counts in another month. Credits on account close an invoice without a cash date.')
html = open(tpl, encoding='utf-8').read()
a = html.find('/*DATA*/'); b = html.find('/*END*/') + len('/*END*/')
html = html[:a] + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + html[b:]
open(out, 'w', encoding='utf-8').write(html)
paid_in = round(sum(x['total'] for x in invoices if x['status'] == 'paid' and x['paidMonth'] == payKey), 2)
print(f'{out}: {len(invoices)} invoices, {len(lines)} lines, {len(vendors)} vendors, paid in {payKey} = {paid_in:,.2f}, open = {sum(x["total"] for x in invoices if x["status"] != "paid"):,.2f}, watch {len(watch)}')
