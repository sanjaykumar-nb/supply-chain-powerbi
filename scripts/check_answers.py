"""Answer key for the SCMS supply chain Power BI project.

Recomputes every headline number in BUILD_GUIDE.md straight from the raw CSV,
using the same rules as the Power Query steps and DAX measures. Build your
dashboard first, then run this and compare. If a number differs, your model
has a bug (usually a filter, a relationship, or line items vs shipments).

Standard library only, so it runs on any Python 3.8+:
    python scripts/check_answers.py
"""
import csv
import datetime as dt
from collections import Counter, defaultdict
from pathlib import Path
from statistics import mean, median

DATA = Path(__file__).resolve().parent.parent / "data" / "SCMS_Delivery_History_Dataset.csv"

# The file is UTF-8 with a byte-order mark; "utf-8-sig" strips it (Power BI: File Origin 65001 UTF-8).
with open(DATA, encoding="utf-8-sig", newline="") as f:
    lines = list(csv.DictReader(f))


def num(text):
    """Number, or None for placeholder text such as 'Freight Included in Commodity Cost'."""
    try:
        return float(text)
    except ValueError:
        return None


def date(text):
    """Both date formats in the file ('28-Sep-06' and '9/11/14'); None for placeholder text."""
    for fmt in ("%d-%b-%y", "%m/%d/%y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def freight_status(text):
    if text.startswith("Freight Included"):
        return "Included in commodity cost"
    if text == "Invoiced Separately":
        return "Invoiced separately"
    if text.startswith("See "):
        return "Recorded on another line"
    return "Recorded" if num(text) is not None else "Other"


def weight_status(text):
    if text == "Weight Captured Separately":
        return "Captured separately"
    if text.startswith("See "):
        return "Recorded on another line"
    return "Recorded" if num(text) is not None else "Other"


def pct(a, b):
    return f"{100 * a / b:.1f}%" if b else "n/a"


def money(x):
    return f"${x:,.0f}"


# ---- Fact_Shipments: one row per ASN/DN (the "first line" carries weight and freight) ----
value_by_asn = defaultdict(float)
for r in lines:
    value_by_asn[r["ASN/DN #"]] += num(r["Line Item Value"]) or 0.0

shipments = []
for r in lines:
    if r["First Line Designation"] != "Yes":
        continue
    sched, deliv, po = date(r["Scheduled Delivery Date"]), date(r["Delivered to Client Date"]), date(r["PO Sent to Vendor Date"])
    freight, weight = num(r["Freight Cost (USD)"]), num(r["Weight (Kilograms)"])
    shipments.append({
        "asn": r["ASN/DN #"], "country": r["Country"], "vendor": r["Vendor"], "fulfill": r["Fulfill Via"],
        "mode": r["Shipment Mode"] if r["Shipment Mode"] != "N/A" else "Not recorded",
        "year": deliv.year, "days_vs_schedule": (deliv - sched).days,
        "lead_days": (deliv - po).days if po else None,
        "freight": freight, "weight": weight, "value": value_by_asn[r["ASN/DN #"]],
        "freight_status": freight_status(r["Freight Cost (USD)"]),
        "weight_status": weight_status(r["Weight (Kilograms)"]),
    })

late = [s for s in shipments if s["days_vs_schedule"] > 0]
both = [s for s in shipments if s["freight"] is not None and s["weight"] is not None and s["weight"] > 0]
priced = [s for s in shipments if s["freight"] is not None]
lead = [s for s in shipments if s["lead_days"] is not None]

print("=" * 72)
print("ANSWER KEY - SCMS supply chain dashboard")
print("=" * 72)

print("\n[Table sizes]")
print(f"  Fact_LineItems rows ............ {len(lines):,}")
print(f"  Fact_Shipments rows ............ {len(shipments):,}   (distinct ASN/DN: {len({r['ASN/DN #'] for r in lines}):,})")
print(f"  Countries / Vendors ............ {len({r['Country'] for r in lines})} / {len({r['Vendor'] for r in lines})}")

print("\n[Page 1 - Overview KPI cards]")
print(f"  Commodity Value (USD) .......... {money(sum(num(r['Line Item Value']) or 0 for r in lines))}")
print(f"  Packs Delivered ................ {sum(num(r['Line Item Quantity']) or 0 for r in lines):,.0f}")
print(f"  Shipments ...................... {len(shipments):,}")
print(f"  % Late vs Schedule ............. {pct(len(late), len(shipments))}   ({len(late):,} late)")
print(f"  Avg Days Late (late only) ...... {mean(s['days_vs_schedule'] for s in late):.1f}")
print(f"  Freight Cost (recorded, USD) ... {money(sum(s['freight'] for s in priced))}")
print(f"  Avg PO-to-Delivery Days ........ {mean(s['lead_days'] for s in lead):.1f}   (median {median(s['lead_days'] for s in lead):.0f}; {len(lead):,} shipments with a PO date)")

print("\n[Shipments and % late by delivery year]")
by_year = defaultdict(list)
for s in shipments:
    by_year[s["year"]].append(s)
for y in sorted(by_year):
    g = by_year[y]
    print(f"  {y}: {len(g):5,} shipments | late {pct(sum(s['days_vs_schedule'] > 0 for s in g), len(g)):>6}")

print("\n[Page 2 - Delivery performance by shipment mode]")
by_mode = defaultdict(list)
for s in shipments:
    by_mode[s["mode"]].append(s)
for m, g in sorted(by_mode.items(), key=lambda kv: -len(kv[1])):
    lt = [s["lead_days"] for s in g if s["lead_days"] is not None]
    lt_txt = f"{mean(lt):6.1f}" if lt else "   n/a"
    print(f"  {m:13} {len(g):5,} shipments | late {pct(sum(s['days_vs_schedule'] > 0 for s in g), len(g)):>6} | avg PO-to-delivery {lt_txt} days")

print("\n[Direct Drop vs From RDC]")
for fv in ("Direct Drop", "From RDC"):
    g = [s for s in shipments if s["fulfill"] == fv]
    print(f"  {fv:12} {len(g):5,} shipments | late {pct(sum(s['days_vs_schedule'] > 0 for s in g), len(g))}")

print("\n[Vendors with the highest % late, minimum 50 shipments]")
by_vendor = defaultdict(list)
for s in shipments:
    by_vendor[s["vendor"]].append(s)
rows = [(v, len(g), sum(s["days_vs_schedule"] > 0 for s in g)) for v, g in by_vendor.items() if len(g) >= 50]
for v, n, k in sorted(rows, key=lambda t: -t[2] / t[1])[:8]:
    print(f"  {v[:40]:40} {n:5,} shipments | late {pct(k, n)}")

print("\n[Top 5 countries by commodity value]")
cv = Counter()
for r in lines:
    cv[r["Country"]] += num(r["Line Item Value"]) or 0
for c, v in cv.most_common(5):
    print(f"  {c:20} {money(v)}")

print("\n[Page 3 - Freight cost by shipment mode (shipments with BOTH freight and weight recorded)]")
print(f"  Overall freight per kg ......... ${sum(s['freight'] for s in both) / sum(s['weight'] for s in both):.2f}   ({len(both):,} shipments)")
for m in sorted(by_mode, key=lambda m: -len(by_mode[m])):
    g = [s for s in both if s["mode"] == m]
    if g:
        print(f"  {m:13} ${sum(s['freight'] for s in g) / sum(s['weight'] for s in g):7.2f} per kg ({len(g):,} shipments)")
print(f"  Freight as % of commodity value  {pct(sum(s['freight'] for s in priced), sum(s['value'] for s in priced))}   (shipments with recorded freight only)")

print("\n[Page 4 - Data quality, Fact_Shipments]")
for label, key in (("Freight status", "freight_status"), ("Weight status", "weight_status")):
    print(f"  {label}:")
    for k, n in Counter(s[key] for s in shipments).most_common():
        print(f"    {k:28} {n:5,}  ({pct(n, len(shipments))})")
print(f"  Shipment mode not recorded ..... {len(by_mode.get('Not recorded', [])):,}")
print(f"  PO date missing (Direct Drop) .. {sum(1 for s in shipments if s['fulfill'] == 'Direct Drop' and s['lead_days'] is None):,}")
print(f"  'Pre-PQ Process' line items .... {sum(r['PQ #'] == 'Pre-PQ Process' for r in lines):,}")
