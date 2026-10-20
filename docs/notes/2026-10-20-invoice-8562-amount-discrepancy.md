# Invoice 8562 (CM Essen): recorded amount is EUR 200 below the standard Y1 rate

Status: **open, no action taken**. Found incidentally during Ask Data Phase 2 (read-only). Nothing was changed.
Written as a standalone issue/note so it stays separate from the Ask Data work.

## Facts (from the CRM ledger, read-only)

| Field | Value |
|---|---|
| Invoice | 8562 |
| Project | CM Essen (Belgium) |
| Maintenance year | Y1 (new installation) |
| Cameras | 6 |
| Recorded amount | EUR 4,468.00 |
| Paid | Yes, payment date 2024-10-25 |
| Invoice year | 2024 |

- The standard Y1 rate is EUR 778 per camera, so 6 cameras = **EUR 4,668.00**. The recorded amount is **EUR 200.00 lower**
  (EUR 744.67 per camera).
- Of the 36 Y1 rows with a camera count invoiced in 2024, **35 are exactly EUR 778 per camera**. 8562 is the only exception.
- The project has no per-project rate override (`rate_y1_override` is empty).
- The related complementary row, invoice 8595 (4 extra cameras, unpaid), is exactly 4 x EUR 778 = EUR 3,112.00.
- There is no bank allocation for 8562; it predates the bank payment records.
- No copy of the original invoice PDF was found in the local Downloads folders, so the issued amount is unverified.

## Open questions (not assumed)

1. Is EUR 4,468 a typing error in the ledger, or the amount actually issued (for example an agreed discount)?
2. If the issued invoice says EUR 4,668, was EUR 200 under-collected, or does the ledger simply differ from the PDF?

## Out of scope / constraints

- Row 8562 is paid and must not be edited as part of any Ask Data change.
- Any correction needs the original PDF and an explicit decision first.
