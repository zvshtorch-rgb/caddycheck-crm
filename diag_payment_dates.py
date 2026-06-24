"""Diagnostic: inspect distribution of paid invoices by payment_date month for a year."""
import sys
import datetime
from collections import defaultdict

import services.supabase_service as svc

SUPABASE_URL = "https://rdoxihpmghrvroddnkdi.supabase.co"
SUPABASE_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"
    ".eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJkb3hpaHBtZ2hydnJvZGRua2RpIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3MzY0NTExNywiZXhwIjoyMDg5MjIxMTE3fQ"
    ".umgghE4z-ClVQ0KY8LQJhJtbG2tYlVh0fY0d9JnYXBA"
)


def _client():
    from supabase import create_client
    return create_client(SUPABASE_URL, SUPABASE_KEY)


svc._get_client = _client
load_invoices = svc.load_invoices

YEAR = int(sys.argv[1]) if len(sys.argv) > 1 else 2025

invoices = load_invoices()
print(f"Total invoices loaded: {len(invoices)}")

paid = [i for i in invoices if i.is_paid()]
print(f"Paid invoices: {len(paid)}")

# Paid invoices whose payment_date falls in YEAR (this is what the blue line uses)
in_year = [i for i in paid if i.payment_date and i.payment_date.year == YEAR]
print(f"Paid invoices with payment_date in {YEAR}: {len(in_year)}")

# Paid invoices marked as year==YEAR but with NO payment_date (invisible to chart)
missing_date = [i for i in paid if i.year == YEAR and not i.payment_date]
print(f"Paid invoices with year=={YEAR} but NO payment_date (excluded from chart): {len(missing_date)}")

print()
print(f"=== Monthly distribution of PAID amount by payment_date month ({YEAR}) ===")
by_month_amt = defaultdict(float)
by_month_cnt = defaultdict(int)
for i in in_year:
    m = i.payment_date.month
    by_month_amt[m] += (i.payment_amount or 0.0)
    by_month_cnt[m] += 1

for m in range(1, 13):
    name = datetime.date(YEAR, m, 1).strftime("%b")
    print(f"  {name} {YEAR}: {by_month_cnt[m]:3d} invoices  amount={by_month_amt[m]:,.0f}")

print()
print(f"=== Distinct payment_date values used in {YEAR} (top by count) ===")
by_date = defaultdict(lambda: [0, 0.0])
for i in in_year:
    d = i.payment_date.date()
    by_date[d][0] += 1
    by_date[d][1] += (i.payment_amount or 0.0)

for d, (cnt, amt) in sorted(by_date.items(), key=lambda kv: kv[1][0], reverse=True)[:20]:
    print(f"  {d}: {cnt:3d} invoices  amount={amt:,.0f}")

print()
print(f"=== Sample of the biggest month rows ===")
if by_month_amt:
    peak_month = max(by_month_amt, key=by_month_amt.get)
    print(f"Peak month: {datetime.date(YEAR, peak_month, 1).strftime('%B %Y')}")
    for i in [x for x in in_year if x.payment_date.month == peak_month][:25]:
        print(f"  {i.payment_date.date()}  {i.project_name[:30]:30s}  {i.maintenance_year:12s}  amt={i.payment_amount:,.0f}")
