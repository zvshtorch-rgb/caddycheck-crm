# Alias proposal (NOT applied): "Plus Landgraaf Arts" / "Plus Landgraaf Aarts" -> "Plus langdraaf Aarts"

Status: **proposal only**. Nothing was added to `config/ask_data_aliases.py`, no ledger row and no project record was changed.
Prepared read-only on 2026-10-20 as part of Ask Data Phase 3.

## The mismatch

| Where | Name as written | Resolves to a project today? |
|---|---|---|
| Project master | `Plus langdraaf Aarts` (Holland, Active, 8 cameras, installed 2017, payment month September, EOP 2027-03-01) | yes (it *is* the project) |
| Ledger, invoice 8713 (inserted 2026-10-19) | `Plus langdraaf Aarts` (1 row, EUR 1,824.00, unpaid) | yes |
| Ledger, 8 historical rows | `Plus Landgraaf Arts` (invoices 8130, 8194, 8292, 8302, 8310, 8434, 8495, 8590; years 2018-2025; EUR 17,508.00, all paid) | **no** |
| PPS payment overview (2 Aug 2025) | "8590 PLUS Landgraaf Arts ... EUR 1,824.00 ... 8th year" | n/a (external document) |
| PPS reconciliation (13 Aug 2026) | "8713 ... Deinze + Zwevegem + Plus Landgraaf Aarts" | n/a (external document) |

The project master spells **Landgraaf** as "langdraaf" (a typo), and the ledger's older rows use "Arts" instead of "Aarts".

## Why these are the same store (evidence)

- Same town and store name in every source; the PPS documents use "Plus Landgraaf Arts/Aarts" for the invoice that the ledger carries under both spellings.
- Camera arithmetic is consistent: 8 cameras x EUR 228 (Y2+ rate) = EUR 1,824 for both invoice 8590 (2025) and invoice 8713 (2026).
- Only one project in the master could be meant: the other Landgraaf project, `Plus Landgraaf Wetzels` (4 cameras), is a different store with its own ledger rows.

## Proposed aliases (for approval)

```python
"Plus Landgraaf Arts":  "Plus langdraaf Aarts",
"Plus Landgraaf Aarts": "Plus langdraaf Aarts",
```

They map spellings to the existing project only; they never rename or modify ledger rows or the project master.

## Effect if approved

- The 8 historical ledger rows (EUR 17,508.00, all paid) join **Holland** in country views of paid and cancelled amounts.
- **Debt is unaffected**: those rows have no unpaid amount, so total outstanding stays EUR 189,998.00 and the country debt split does not change.
- A user typing just "Landgraaf" would become *ambiguous* (Wetzels or Aarts) and be asked to choose. Today it silently resolves to Wetzels, because the Aarts project name is misspelled.
- Tests to add with the change: both aliases resolve; "Landgraaf" is ambiguous with both candidates; Wetzels stays separate; 8 rows / EUR 17,508.00 move to Holland; total outstanding unchanged.

## Not proposed here

- Correcting the typo in the project master (`langdraaf` -> `Landgraaf`) is a data change that also touches invoice 8713 and any external references; it needs its own decision.
