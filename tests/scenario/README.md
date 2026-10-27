# Test scenario: Fernwood Distribution cloud portfolio

All people, companies and figures in this folder are invented. Any resemblance to real ones is coincidental.

Fernwood Distribution Ltd is a mid-size food distributor. Its sponsor wants to know which of five proposed cloud initiatives to fund first, with no portfolio budget agreed yet. The scenario date is 2026-03-09.

The `documents/` folder holds what the client sent. It is deliberately imperfect, the way real inputs are. A person or an agent following `SKILL.md` should be able to turn these documents into a defensible matrix, and should notice what is wrong with them. The answer key is in `../TEST_PLAN.md`, not here.

| File | Document type (catalog code) | Covers |
|---|---|---|
| `PROPOSAL-P1.md` to `PROPOSAL-P5.md` | PROPOSAL | One initiative each |
| `ESTIMATE-P1.md` to `ESTIMATE-P4.md` | ESTIMATE | P1 to P4 (none exists for P5) |
| `TARGET_STATE.md` | TARGET_STATE | Portfolio |
| `LENS.md` | LENS | Portfolio |
| `FUNDING.md` | FUNDING | Portfolio |
| `COST_BASELINE.md` | COST_BASELINE | P2 and P3 only |
| `INCIDENTS.md` | INCIDENTS | Portfolio |
| `SECURITY.md` | SECURITY | Policy and review status |
| `CAPACITY.md` | CAPACITY | Portfolio |
| `CONTRACTS.md` | CONTRACTS | Portfolio |

`output/portfolio.json` is the config produced by following the skill on these documents. The test runner checks it.
