# Report: historical ledger names that match no current project (paid/cancelled country reporting)

Status: **report only**. No names were mapped, no aliases added, no data changed. Generated read-only on 2026-10-20
(Ask Data Phase 3). The full list of all 69 names is in
[unmatched_ledger_names_2026-10-20.csv](./unmatched_ledger_names_2026-10-20.csv) (rows, paid/cancelled/unpaid amounts, years,
and what the user-input resolver would suggest, for information only).

## What is affected

Ledger rows are attached to a country through the project they belong to. A ledger name that matches no current project
(exactly, after case/accent/punctuation folding, or through an approved alias) has no country, so it is left out of every
country figure.

| | Unmatched | Ledger total | Share |
|---|---|---|---|
| Names / rows | 69 names, 221 rows | | |
| Paid amount | EUR 394,061.00 | EUR 1,894,246.00 | **20.8%** |
| Cancelled amount | EUR 10,020.00 | EUR 17,612.00 | **56.9%** |
| **Unpaid (debt)** | **EUR 0.00** | EUR 189,998.00 | **0%** |

- **Debt reporting is not affected.** Every unpaid row matches a project, so the debt-by-country result reconciles to EUR 189,998.00 with nothing unassigned.
- **Paid and cancelled country reporting is materially affected:** about one fifth of historical paid money and more than half of cancelled money cannot be attributed to a country.
- Most of these are old projects that are no longer in the project master: 15 names were last invoiced in 2018, 15 in 2022 and 12 in 2019; only 2 names have a 2026 invoice.

## Names that materially affect the reporting

Materiality rule: combined paid + cancelled of at least EUR 10,000, plus every name that carries cancelled amounts.

| Ledger name | Rows | Paid | Cancelled | Years | What the resolver would suggest (information only, not applied) |
|---|---|---|---|---|---|
| AD Vlijtingen | 4 | 20,468.00 | | 2020-2022 | ambiguous: AD Tielt-Winge, AD Itegem, AD Tienen, AD Vichte (no real match) |
| Plus Landgraaf Arts | 8 | 17,508.00 | | 2018-2025 | Plus langdraaf Aarts (see the separate alias proposal) |
| Jumbo Roden Jansema | 5 | 11,241.00 | | 2016-2021 | none |
| AD Kalmthout | 2 | 10,060.00 | | 2024-2025 | ambiguous: CM Kalmthout (a cancelled project with a different network prefix) |
| AD Ninove | 1 | | 9,336.00 | 2025 | ambiguous: AH Ninove, AD Deinze, AD Knokke |
| AH Merksem | (cancelled) | | 684.00 | | not applied |

The 4 largest paid names total EUR 59,277.00 (14.7% of the unmatched amount); the remainder is spread thinly over the other 65 names
(largest of them: Coop Bert Stuut 9,872.00, AD Grobbendonk 9,872.00, Jumbo Lokerenpassage 9,114.00).

## How the 69 names split (resolver view, informational)

- 3 are near-identical spellings of a current project: `Plus Landgraaf Arts` (EUR 17,508.00, Holland), `CM Kapelle-op-den-bos` -> `CM Kapelle-op-den-bosch` (EUR 8,048.00, Belgium) and `AD Bloemmolens` -> `AD Bloemmolen` (EUR 456.00, Belgium).
- 16 are ambiguous or similar to several different projects. A wrong guess here would put money in the wrong country, which is why ledger matching is never fuzzy.
- 50 have no resemblance to any current project (former Dutch/Luxembourg sites such as Jumbo, Coop, Plus, Cactus ones).

## Suggested decisions (none taken)

1. Approve or reject the Landgraaf alias proposal (separate note).
2. Decide individually on the two other near-identical spellings (Kapelle-op-den-bos, Bloemmolens).
3. For the ~50 former sites, decide whether to keep them as inactive projects with a country in the project master. That is a data change and needs its own approval.
4. In a later phase, show an "Unassigned" bucket for paid and cancelled country views too, the way debt-by-country already does, so the missing 20.8% is always visible instead of silent.
