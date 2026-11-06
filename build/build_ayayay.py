"""TEST_build_dashboard.py - rebuilds the Ayayay dashboard from the TEST workbook.

Usage:  python TEST_build_dashboard.py TEST_Ayayay_v2.xlsx TEST_dashboard_template_ayayay.html TEST_Ayayay_index.html

Reads ONLY the typed tables (Lines, Invoices, Payments, Sales, Calendar, Vendors, Watch) plus the three
input cells (Owed!B2 position date, Owed!B3 payments month, Ratios!B2 catering switch) and recomputes
everything itself, so it does not matter whether Excel has recalculated the file.
The dashboard never contains anything that is not in the workbook.
"""
import sys, json, datetime as dt, collections
from openpyxl import load_workbook

src, tpl, out = sys.argv[1], sys.argv[2], sys.argv[3]
wb = load_workbook(src, data_only=True)

def rows(ws, first_col_required=1, start=5):
    r = start
    while True:
        v = ws.cell(r, first_col_required).value
        if v is None or v == "":
            # allow one blank line? no - tables are contiguous
            break
        yield r
        r += 1

def iso(d): return d.strftime("%Y-%m-%d") if isinstance(d, (dt.datetime, dt.date)) else (d or None)
def md(d): return d.strftime("%m/%d") if isinstance(d, (dt.datetime, dt.date)) else ""
def num(v): return float(v or 0)

# ---- Calendar
cal = wb["Calendar"]
weeks = []
for r in rows(cal):
    weeks.append(dict(id=cal.cell(r,1).value, startD=cal.cell(r,2).value, endD=cal.cell(r,3).value, month=cal.cell(r,4).value, key=cal.cell(r,5).value, partial=(str(cal.cell(r,6).value or "").lower()=="yes"), note=cal.cell(r,7).value or ""))
weeks.sort(key=lambda w: w["startD"])
months = []  # from month list I5:J
r = 5
while cal.cell(r, 9).value:
    months.append((cal.cell(r,9).value, cal.cell(r,10).value)); r += 1
def week_of(d):
    for w in weeks:
        if w["startD"] <= d <= w["endD"]: return w["id"]
    return "NO WEEK"
WIX = {w["id"]: i for i, w in enumerate(weeks)}

# ---- Vendors
ven = wb["Vendors"]; vendors = [ven.cell(r,1).value for r in rows(ven)]

# ---- Lines
ln = wb["Lines"]
lines = []
for r in rows(ln, first_col_required=2):
    d = ln.cell(r,2).value
    lines.append(dict(date=d, week=week_of(d), vendor=ln.cell(r,5).value, inv=str(ln.cell(r,6).value), cat=ln.cell(r,8).value, sub=ln.cell(r,9).value, ext=round(num(ln.cell(r,11).value)*num(ln.cell(r,13).value), 2)))

# ---- Payments
pm = wb["Payments"]
pays = collections.defaultdict(list)
cards = collections.defaultdict(float)
for r in rows(pm, first_col_required=3):
    key = f"{pm.cell(r,3).value}|{pm.cell(r,4).value}"
    dpaid = pm.cell(r,2).value; am = pm.cell(r,9).value
    eff = am if am else (dpaid.strftime("%Y-%m") if isinstance(dpaid, dt.datetime) else None)
    pays[key].append(dict(date=dpaid, amount=num(pm.cell(r,5).value), method=pm.cell(r,6).value, card=pm.cell(r,7).value, eff=eff))
    if pm.cell(r,7).value: cards[pm.cell(r,7).value] += num(pm.cell(r,5).value)

# ---- Invoices
iv = wb["Invoices"]
invoices = []
prod_by_key = collections.defaultdict(float)
for l in lines: prod_by_key[f"{l['vendor']}|{l['inv']}"] += l["ext"]
for r in rows(iv, first_col_required=3):
    vendor = iv.cell(r,2).value; number = str(iv.cell(r,3).value); key = f"{vendor}|{number}"
    d = iv.cell(r,4).value; fees = num(iv.cell(r,8).value); tax = num(iv.cell(r,9).value)
    product = round(prod_by_key.get(key, 0.0), 2); total = round(product + fees + tax, 2)
    paid = round(sum(p["amount"] for p in pays.get(key, [])), 2)
    status = "paid" if abs(total - paid) < 0.005 else "open"
    dated = [p["date"] for p in pays.get(key, []) if isinstance(p["date"], dt.datetime)]
    paid_on = max(dated) if dated else None
    paid_month = pays[key][0]["eff"] if pays.get(key) else None
    invoices.append(dict(vendor=vendor, number=number, date=iso(d), week=week_of(d), product=product, fees=round(fees+tax, 2), total=total,
                         status=status, paidOn=iso(paid_on), paidMonth=paid_month, due=iso(iv.cell(r,14).value), note=iv.cell(r,13).value or "",
                         openAmount=round(total - paid, 2), card=(pays[key][0]["card"] if pays.get(key) else None)))
invoices.sort(key=lambda x: (x["date"], x["vendor"], x["number"]))

# ---- Sales
sl = wb["Sales"]
switch = str(wb["Ratios"]["B2"].value or "No").strip().lower() == "yes"
sales_raw = collections.defaultdict(lambda: [0.0]*len(weeks))
for r in rows(sl, first_col_required=2):
    w = sl.cell(r,2).value; cat = sl.cell(r,3).value; v = num(sl.cell(r,4).value)
    if w in WIX: sales_raw[cat][WIX[w]] += v
SALEMAP = {"Food":"Food", "Cocktails & Liquor":"Cocktails & liquor", "Beer":"Beer", "Wine":"Wine", "Non-Alcoholic":"Non-alcoholic", "Open Drink":"Open drink"}
sales = {}
for k, lab in SALEMAP.items(): sales[lab] = [round(x, 2) for x in sales_raw.get(k, [0.0]*len(weeks))]
if switch:
    sales["Food"] = [round(a+b+c, 2) for a, b, c in zip(sales["Food"], sales_raw.get("Catering", [0]*len(weeks)), sales_raw.get("Events", [0]*len(weeks)))]
sales["TOTAL NET SALES"] = [round(sum(sales[k][i] for k in SALEMAP.values()), 2) for i in range(len(weeks))]

# ---- aggregates by week
CATS = ["Food", "Bar", "Supply", "Cleaning"]
SUBMAP = {"Cocktails & Liquor":"Cocktails & liquor", "Beer":"Beer", "Wine":"Wine", "Non-Alcoholic":"Non-alcoholic", "Unallocated":"Unallocated"}
categories = {c: [0.0]*len(weeks) for c in CATS}
barSub = {v: [0.0]*len(weeks) for v in SUBMAP.values()}
vend_weeks = {v: [0.0]*len(weeks) for v in vendors}
for l in lines:
    if l["week"] not in WIX: continue
    i = WIX[l["week"]]
    categories[l["cat"]][i] += l["ext"]
    if l["cat"] == "Bar" and l["sub"] in SUBMAP: barSub[SUBMAP[l["sub"]]][i] += l["ext"]
    if l["vendor"] in vend_weeks: vend_weeks[l["vendor"]][i] += l["ext"]
fees = [0.0]*len(weeks); counts = [0]*len(weeks)
for x in invoices:
    if x["week"] in WIX: fees[WIX[x["week"]]] += x["fees"]; counts[WIX[x["week"]]] += 1
rnd = lambda arr: [round(v, 2) for v in arr]
categories = {k: rnd(v) for k, v in categories.items()}; barSub = {k: rnd(v) for k, v in barSub.items()}

# ---- Watch
wt = wb["Watch"]
watch = []
for r in rows(wt, first_col_required=3):
    status = str(wt.cell(r,6).value or "Open")
    if status.lower() == "closed": continue
    tag = str(wt.cell(r,2).value or "Open")
    amt = wt.cell(r,5).value
    body = str(wt.cell(r,4).value or "")
    if isinstance(amt, (int, float)) and amt: body = f"${amt:,.2f}. " + body
    body += f"  [{status}]" if status.lower() not in ("open",) else ""
    watch.append(dict(tag=tag, title=str(wt.cell(r,3).value), body=body))

# ---- position / carry
pos = wb["Owed"]["B2"].value; paymonth = str(wb["Owed"]["B3"].value)
as_of = iso(pos)
# carry card: the last month on the Month list that has invoices
mk_with_inv = [k for (m, k) in months if any(w["key"] == k and counts[WIX[w["id"]]] for w in weeks)]
cm_key = mk_with_inv[-1] if mk_with_inv else months[-1][1]
cm_name = dict((k, m) for m, k in months)[cm_key]
cm_weeks = [w for w in weeks if w["key"] == cm_key]
nxt = (cm_weeks[-1]["endD"] + dt.timedelta(days=1))
nxt_name = nxt.strftime("%B")
carry = dict(monthKey=cm_key, month=cm_name, next=nxt_name, from_=iso(cm_weeks[0]["startD"]), to=iso(cm_weeks[-1]["endD"]), paidFrom=iso(nxt),
             title=f"{cm_name} carried into {nxt_name}", desc=f"Invoices dated {md(cm_weeks[0]['startD'])} to {md(cm_weeks[-1]['endD'])} that were paid in {nxt_name} (or count in a later month) or are still open.",
             label=f"Carried from {cm_name}", paidLabel=f"Paid in {nxt_name}")
carry["from"] = carry.pop("from_")

partial = [w["id"] for w in weeks if w["partial"]]
partialNote = {w["id"]: (w["note"].split("Partial: ")[-1] if w["note"] else "partial week") for w in weeks if w["partial"]}
footer = (f"Built from {src.split('/')[-1]} on {dt.date.today().strftime('%m/%d/%Y')}; payment position as of {pos.strftime('%m/%d/%Y')}, payments month {paymonth}. "
          "Weeks run Monday to Sunday by invoice date; a week that crosses a month end goes to the month with more of its days, so 09/28-10/04 is O1 (October). A1 (07/27-08/02) covers 08/01-08/02 only. "
          "Bar ratio = bar purchases / cocktails & liquor, beer, non-alcoholic and open-drink sales. "
          + ("Catering and Events sales are INCLUDED in food sales (Ratios!B2 = Yes)." if switch else "Catering and Events sales are excluded from food sales (Ratios!B2 = No)."))

data = dict(asOf=as_of, paymentsMonth=paymonth,
            weeks=[dict(id=w["id"], start=md(w["startD"]), end=md(w["endD"]), month=w["month"], entered=bool(counts[i] or sales["TOTAL NET SALES"][i]), invoices=counts[i]) for i, w in enumerate(weeks)],
            categories=categories, barSub=barSub, fees=rnd(fees),
            vendors=[dict(name=v, weeks=rnd(vend_weeks[v])) for v in vendors if any(vend_weeks[v])],
            partial=partial, partialNote=partialNote, sales=sales,
            invoices=[{k: v for k, v in x.items()} for x in invoices],
            watch=watch, carry=carry, footer=footer,
            cards=[dict(card=k, amount=round(v, 2)) for k, v in sorted(cards.items())])

t = open(tpl, encoding="utf-8").read()
t = t.replace("__AYAYAY_DATA__", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
t = t.replace("__BUILD_COMMENT__", f"Data built by TEST_build_dashboard.py from {src.split('/')[-1]} on {dt.date.today().isoformat()} (position as of {as_of})")
t = t.replace("__MONTH_OPTIONS__", "".join(f"<option>{m}</option>" for m, k in months))
t = t.replace("__CARRY_TITLE__", carry["title"]).replace("__CARRY_DESC__", carry["desc"]).replace("__CARRY_LABEL__", carry["label"]).replace("__CARRY_PAID_LABEL__", carry["paidLabel"])
open(out, "w", encoding="utf-8").write(t)
tot_prod = round(sum(sum(v) for v in categories.values()), 2)
print(f"wrote {out}: {len(invoices)} invoices, {len(lines)} lines, {len(weeks)} weeks, product {tot_prod:,.2f}, billed {sum(x['total'] for x in invoices):,.2f}, open {sum(x['openAmount'] for x in invoices if x['status']=='open'):,.2f}, watch cards {len(watch)}, as of {as_of}")
