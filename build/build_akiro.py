"""TEST_build_dashboard.py - rebuilds the dashboard (TEST_index.html) from the TEST_ workbook.

Usage:  python3 TEST_build_dashboard.py TEST_Akiro_v2.xlsx TEST_dashboard_template.html TEST_index.html

Reads ONLY the typed data tables (Calendar, Vendors, Lines, Invoices, Payments, Sales, Watch log) and
recomputes every figure itself - it does not depend on Excel having recalculated.  Nothing is written
back to the workbook.
"""
import sys, json, re, datetime as dt, calendar, collections
from zoneinfo import ZoneInfo
import openpyxl

xlsx, template, out = sys.argv[1], sys.argv[2], sys.argv[3]
wb = openpyxl.load_workbook(xlsx, data_only=True)

def rows(ws, ncols):
    """typed table rows: header in row 1, stop at first empty key cell (col A)."""
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ncols + 1)]
        if all(v in (None, "") for v in vals[:4]):
            continue
        yield vals

def iso(d): return d.strftime("%Y-%m-%d") if isinstance(d, (dt.datetime, dt.date)) else None
def md(d): return d.strftime("%m/%d") if isinstance(d, (dt.datetime, dt.date)) else ""
def r2(x): return round(float(x or 0) + 1e-9, 2)

# ---------------- Calendar + months
cal = wb["Calendar"]
weeks = []
for r in range(2, cal.max_row + 1):
    w = cal.cell(r, 1).value
    if not w: break
    weeks.append(dict(id=str(w), start_d=cal.cell(r, 2).value, end_d=cal.cell(r, 3).value, month=cal.cell(r, 4).value, key=cal.cell(r, 5).value))
months = []
for r in range(2, cal.max_row + 1):
    m = cal.cell(r, 10).value
    if not m: break
    months.append(dict(name=m, key=cal.cell(r, 11).value))
for m in months:
    y, mo = int(m["key"][:4]), int(m["key"][5:7])
    m["end"] = f"{y}-{mo:02d}-{calendar.monthrange(y, mo)[1]:02d}"
    ny, nm = (y + 1, 1) if mo == 12 else (y, mo + 1)
    m["next"] = calendar.month_name[nm]
    wk = [w for w in weeks if w["month"] == m["name"]]
    m["weeks"] = f"{wk[0]['id']}–{wk[-1]['id']}, dated {md(wk[0]['start_d'])} to {md(wk[-1]['end_d'])}" if wk else ""
widx = {w["id"]: i for i, w in enumerate(weeks)}
def week_of(d):
    for w in weeks:
        if w["start_d"] <= d <= w["end_d"]: return w["id"]
    return None

# ---------------- Vendors (for register flag only)
# ---------------- Lines
lines = []
owner = None
for v in rows(wb["Lines"], 17):
    loc, date, _, _, vendor, inv, desc, cat, sub, qo, qr, unit, price, _, note, _, charge = v
    if date is None: continue
    owner = owner or loc
    ext = r2((qr or 0) * (price or 0))
    lines.append(dict(date=date, week=week_of(date), vendor=str(vendor).strip(), inv=str(inv).strip(), cat=cat, sub=sub, ext=ext,
                      charge=(str(charge).strip() if charge else "")))
# ---------------- Invoices
invoices = []
for v in rows(wb["Invoices"], 23):
    loc, vendor, inv, date, _, _, _, fees, tax, _, printed, _, terms, due = v[:14]
    note = v[22]
    if inv is None: continue
    invoices.append(dict(vendor=str(vendor).strip(), inv=str(inv).strip(), date=date, fees=float(fees or 0), tax=float(tax or 0), printed=float(printed or 0), due=due, note=note or "", week=week_of(date)))
# ---------------- Payments
pays = collections.defaultdict(list); cash = []
for v in rows(wb["Payments"], 12):
    loc, date, vendor, inv, amt, method, ev, ovr = v[:8]
    if vendor is None: continue
    key = int(ovr[:4]) * 100 + int(ovr[5:7]) if ovr else (date.year * 100 + date.month if date else 0)
    p = dict(date=date, amt=float(amt or 0), method=method, ev=ev, key=key)
    if inv in (None, ""): cash.append(dict(vendor=vendor, **p))
    else: pays[(str(vendor).strip(), str(inv).strip())].append(p)
# ---------------- Sales
sales = collections.defaultdict(lambda: [0.0] * len(weeks)); sales_weeks_seen = set()
for v in rows(wb["Sales"], 5):
    loc, wk, cat, net = v[:4]
    if wk is None: continue
    sales_weeks_seen.add(wk)
    if wk in widx: sales[cat][widx[wk]] += float(net or 0)
# ---------------- Watch log
watch = []
for v in rows(wb["Watch log"], 5):
    d, loc, title, status, text = v
    if title is None and text is None: continue
    watch.append(dict(tag=status or "Open", title=title or "", body=text or "", date=iso(d)))

# ---------------- compute
NW = len(weeks)
prod_by_inv = collections.defaultdict(float); other_by_inv = collections.defaultdict(float); other_name = {}
for l in lines:
    prod_by_inv[(l["vendor"], l["inv"])] += l["ext"]
    if l["charge"] != owner:
        other_by_inv[(l["vendor"], l["inv"])] += l["ext"]; other_name[(l["vendor"], l["inv"])] = l["charge"] or "?"
cats = ["Food", "Bar", "Supply", "Cleaning"]; subs = ["Cocktails & liquor", "Beer", "Wine", "Sake", "Non-alcoholic"]
categories = {c: [0.0] * NW for c in cats}; barSub = {c: [0.0] * NW for c in subs}
vend = collections.defaultdict(lambda: [0.0] * NW)
for l in lines:
    if l["week"] is None or l["cat"] not in cats or l["charge"] != owner: continue
    i = widx[l["week"]]
    categories[l["cat"]][i] += l["ext"]; vend[l["vendor"]][i] += l["ext"]
    if l["cat"] == "Bar" and l["sub"] in subs: barSub[l["sub"]][i] += l["ext"]
fees = [0.0] * NW; inv_count = [0] * NW
inv_json = []; problems = []
whole_other = {k for k, v in other_by_inv.items() if abs(v - prod_by_inv[k]) < 0.005 and abs(prod_by_inv[k]) > 0.005}
for x in invoices:
    if x["week"] is not None and (x["vendor"], x["inv"]) not in whole_other:
        i = widx[x["week"]]
        fees[i] += x["fees"] + x["tax"]; inv_count[i] += 1
    product = r2(prod_by_inv.get((x["vendor"], x["inv"]), 0))
    total = r2(product + x["fees"] + x["tax"])
    if abs(total - x["printed"]) > 0.005: problems.append(f"{x['vendor']} {x['inv']} does not foot: calc {total} printed {x['printed']}")
    pl = pays.get((x["vendor"], x["inv"]), [])
    paid = r2(sum(p["amt"] for p in pl))
    status = "paid" if paid >= total - 0.005 else "open"
    paid_on = max((p["date"] for p in pl if p["date"]), default=None) if pl else None
    k = (x["vendor"], x["inv"]); oth = r2(other_by_inv.get(k, 0))
    if oth: x["note"] = f"CHARGED TO {other_name[k].upper()}: ${oth + (x['fees'] + x['tax'] if k in whole_other else 0):,.2f} owed back to {owner} - not in {owner}'s purchases. " + x["note"]
    inv_json.append(dict(vendor=x["vendor"], number=x["inv"], date=iso(x["date"]), week=x["week"], product=product, fees=r2(x["fees"] + x["tax"]), total=total,
                         chargedTo=(other_name[k] if oth else None), otherAmount=(r2(oth + (x["fees"] + x["tax"] if k in whole_other else 0)) if oth else 0),
                         status=status, paidOn=iso(paid_on), due=iso(x["due"]), note=(x["note"] + (f" | PARTIAL: {paid:.2f} paid" if 0 < paid < total - 0.005 else ""))))
inv_json.sort(key=lambda x: (x["date"], x["vendor"].lower(), x["number"]))
vendors = sorted(vend.items(), key=lambda kv: -sum(kv[1]))
# ---- the same checks as the Checks sheet, recomputed here
vendor_names = {str(v[0]).strip() for v in rows(wb["Vendors"], 5) if v[0]}
inv_keys = {(x["vendor"], x["inv"]) for x in invoices}
for l in lines:
    if l["week"] is None: problems.append(f"Line dated {l['date']:%m/%d/%Y} ({l['vendor']} {l['inv']}) is outside every Calendar week")
    if (l["vendor"], l["inv"]) not in inv_keys: problems.append(f"Line for {l['vendor']} {l['inv']} has no invoice on the Invoices table")
    if l["cat"] not in cats: problems.append(f"Line for {l['vendor']} {l['inv']} has unknown category '{l['cat']}'")
    if l["cat"] == "Bar" and not l["sub"]: problems.append(f"Bar line for {l['vendor']} {l['inv']} has no Bar Subcategory")
    if not l["charge"]: problems.append(f"Line for {l['vendor']} {l['inv']} has no Charge to")
    if l["cat"] != "Bar" and l["sub"]: problems.append(f"{l['cat']} line for {l['vendor']} {l['inv']} carries a Bar Subcategory")
for x in invoices:
    if x["week"] is None: problems.append(f"Invoice {x['vendor']} {x['inv']} dated {x['date']:%m/%d/%Y} is outside every Calendar week")
    if x["vendor"] not in vendor_names: problems.append(f"Vendor '{x['vendor']}' (invoice {x['inv']}) is not on the Vendors sheet")
dups = [k for k, n in collections.Counter((x["vendor"], x["inv"]) for x in invoices).items() if n > 1]
for k in dups: problems.append(f"Duplicate invoice number: {k[0]} {k[1]}")
for k in pays:
    if k not in inv_keys: problems.append(f"Payment for {k[0]} {k[1]} points at an invoice that does not exist")
for x in inv_json:
    pl = pays.get((x["vendor"], x["number"]), []); paid = r2(sum(p["amt"] for p in pl))
    if paid > x["total"] + 0.005: problems.append(f"{x['vendor']} {x['number']} is overpaid: {paid:.2f} against {x['total']:.2f}")
for wk in sales_weeks_seen:
    if wk not in widx: problems.append(f"Sales row for week '{wk}' is not on the Calendar")
problems = sorted(set(problems))
tz = ZoneInfo("America/Sao_Paulo"); now = dt.datetime.now(tz)
location = wb["Lines"].cell(2, 1).value or "Akiro"
first_sales = next((i for i in range(NW) if sales["Food"][i] > 0), None)
partial = [w["id"] for i, w in enumerate(weeks) if first_sales is not None and i < first_sales and inv_count[i] > 0]
data = dict(
    location=location + " Chicago" if location == "Akiro" else location,
    asOf=now.strftime("%Y-%m-%d"), updated=now.strftime("%b %-d, %Y %-I:%M %p"),
    weeks=[dict(id=w["id"], start=md(w["start_d"]), end=md(w["end_d"]), endIso=iso(w["end_d"]), month=w["month"], entered=inv_count[i] > 0, invoices=inv_count[i]) for i, w in enumerate(weeks)],
    months=months,
    categories={c: [r2(v) for v in vals] for c, vals in categories.items()},
    barSub={c: [r2(v) for v in vals] for c, vals in barSub.items()},
    fees=[r2(f) for f in fees],
    vendors=[dict(name=n, weeks=[r2(v) for v in vals]) for n, vals in vendors],
    partial=partial, partialNote={p: "no sales export for this week" for p in partial},
    sales={c: [r2(v) for v in vals] for c, vals in sales.items()},
    invoices=inv_json, watch=watch,
    checks=dict(ok=not problems, problems=problems),
    payments=[dict(vendor=v, invoice=i, amount=r2(pp["amt"]), date=iso(pp["date"]), key=pp["key"], method=pp["method"],
                   billedMonth=next((weeks[widx[x["week"]]]["month"] for x in invoices if (x["vendor"], x["inv"]) == (v, i) and x["week"]), None))
              for (v, i), pl in pays.items() for pp in pl]
             + [dict(vendor=c["vendor"], invoice="", amount=r2(c["amt"]), date=iso(c["date"]), key=c["key"], method=c["method"], billedMonth=None) for c in cash],
    cash=[dict(vendor=c["vendor"], amount=c["amt"], month=next((m["name"] for m in months if int(m["key"][:4]) * 100 + int(m["key"][5:7]) == c["key"]), str(c["key"])), note=c["ev"]) for c in cash],
    footer=f"Built from {xlsx.split('/')[-1]} by TEST_build_dashboard.py. Weeks run Monday to Sunday by invoice date (Calendar sheet). A week that crosses a month end belongs to the month with more of its days; {location} exception: A1 (07/27–08/02) stays in August. Cash paid without an invoice is listed on the Owed tab and is not in purchases.",
)
data["sales"]["TOTAL NET SALES"] = [r2(sum(data["sales"][c][i] for c in data["sales"] if c != "TOTAL NET SALES")) for i in range(NW)]
if partial == ["A1"]: data["partialNote"] = {"A1": "08/01–08/02 only: July is not tracked"}

tpl = open(template, encoding="utf-8").read()
js = json.dumps(data, ensure_ascii=False, indent=1)
html = tpl.replace("/*DATA*/{}/*DATA*/", js)
open(out, "w", encoding="utf-8").write(html)
tot_prod = r2(sum(sum(v) for v in categories.values())); tot_billed = r2(sum(x["total"] for x in inv_json))
paid = r2(sum(x["total"] for x in inv_json if x["status"] == "paid")); open_ = r2(sum(x["total"] for x in inv_json if x["status"] != "paid"))
print(f"{out}: {len(inv_json)} invoices | product {tot_prod:,.2f} | billed {tot_billed:,.2f} | PAID {paid:,.2f} | NOT PAID {open_:,.2f} on {sum(1 for x in inv_json if x['status']!='paid')} | watch {len(watch)} | cash {len(cash)}")
for p in problems: print("PROBLEM:", p)
