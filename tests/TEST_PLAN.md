# Test plan: ea-impact-matrix

## 1. Objective

Show that the skill, used as written, turns a realistic and imperfect set of client documents into a matrix that is correct, defensible and honest about what is missing. Two things are tested: the method (can the instructions in `SKILL.md` lead to the right answer on documents with traps?) and the tool (does `build_impact_matrix.py` behave as documented?).

## 2. Scope

In scope: the instructions in `SKILL.md`, the required-document catalog, the generator script and its flags, the workbook it writes, the template and the example, and the hygiene of the public copy.

Out of scope: Excel desktop rendering (LibreOffice is used to recalculate), very large portfolios, and the quality of the judgement a human reviewer would add.

## 3. Requirements

| ID | Requirement |
|---|---|
| R01 | The lens is confirmed by a named person and date before scoring, and the weights match it |
| R02 | The documents the matrix needs are defined, and missing ones are flagged |
| R03 | Every score is a whole number from 1 to 5 with an assessment, an owner and evidence |
| R04 | The weighted score is correct and the fixed cell addresses stay stable |
| R05 | Rank and score per $100k are correct |
| R06 | Cost comes from line items, and a mismatch with a stated total is called out |
| R07 | A saving without a baseline is not claimed |
| R08 | Initiatives with identical score profiles are flagged |
| R09 | Waves are built as approval tranches that fit the stated funding |
| R10 | Cross-dependencies include shared people, commercial links and enablers |
| R11 | A policy or principle breach is reflected in the score and the gates |
| R12 | Risks have an owner, a response and a residual score, and the red threshold is configurable |
| R13 | Decisions are logged with a named decider |
| R14 | Every information gap becomes an open question |
| R15 | An initiative without a business case or estimate is held, not scored |
| R16 | `--check` and `--strict` report problems with the documented exit codes |
| R17 | Errors block the build and no workbook is written |
| R18 | Configs written for the earlier version still build with the same results |
| R19 | Output is deterministic and has no formula errors after recalculation |
| R20 | The template and the example are usable as shipped |
| R21 | The public copy has no client data and no AI-writing tells (em or en dashes) |

## 4. Scenario

Fernwood Distribution Ltd is fictional. Its CFO wants five proposed cloud initiatives ranked for a first funding tranche of up to 150,000 USD, weighted to reduce run cost. The 17 documents in `scenario/documents/` are what the client sent. Every required document type is present at least once, and some are deliberately incomplete.

The answer key below was written before the config was produced. It is the standard the output is judged against.

### 4.1 Planted traps and expected handling

| Trap | Where | Expected handling | Tests |
|---|---|---|---|
| Sponsor states the lens in an email | LENS | Use `run_cost`, record Dana Whitcombe and 2026-03-02, review date 2026-06-01 | A02 |
| No portfolio budget, only a first tranche up to 150,000 | FUNDING | Build waves as tranches, log the open envelope | A07, A11 |
| P4 deck total 205,000, line items 220,000 | PROPOSAL/ESTIMATE-P4 | Use 220,000, note the difference, ask why | A04, A11 |
| P4 stores keys on nodes and moves patching, against policy SEC-07 | PROPOSAL-P4, SECURITY | Low security score, gate 2 not met, risk with owner | A05, A15 |
| P4 contradicts "managed over self-managed" | TARGET_STATE | Low strategic alignment | A05 |
| P4 needs 1,100 architect hours; one architect has 800 in the period | ESTIMATE-P4, CAPACITY | Low delivery confidence, capacity dependency with P2 | A08, A14 |
| P4 claims 30% saving, no baseline | COST_BASELINE | No saving claimed, baseline requested | A09, A12 |
| P3 vendor says "up to 45% lower", billing is by ingest | PROPOSAL-P3, ESTIMATE-P3 | Use the conservative range, not the vendor figure | A09 |
| P3 notice date for tool C is 2026-09-16 | CONTRACTS | Date recorded, decision date asked | A11 |
| P1 has no saving but enables P3 and P4 | PROPOSAL-P1 | Zero saving, argued as an enabler, dependencies recorded | A08, A09 |
| P2 has a hard date (2027-01-31), outages, audit finding, evidenced saving | CONTRACTS, INCIDENTS, COST_BASELINE | Highest score, rank 1 | A06 |
| P5 is an idea with no estimate or case | PROPOSAL-P5 | Held, not scored, questions logged | A03, A12, B08 |
| Scores must differ between initiatives | all | No identical profiles | A10 |

### 4.2 Expected result

Weighted scores (run-cost weights 10, 35, 15, 10, 5, 15, 10 percent): P2 4.10, P1 3.45, P3 3.25, P4 1.95. Ranks P2, P1, P3, P4. Score per $100k: P1 7.19, P3 4.51, P2 4.27, P4 0.89. Wave 1 is P1 and P2 at 144,000. Wave 2 is P3. P4 and P5 are held.

## 5. Test layers and cases

Layer A (acceptance) checks the config produced by following `SKILL.md` on the scenario documents against the answer key: A01 to A16.

Layer B (workbook) builds the workbook, recalculates it in LibreOffice and compares it with an independent recomputation in `run_tests.py` that does not use the generator's code: B01 to B09.

Layer C (tool) mutates configs or calls the script directly: C01 to C17. These cover the example, template, dump-config, catalog and intake, error handling, strict mode, threshold override, residual and response validation, legacy compatibility, determinism, flag documentation, and hygiene of the public copy.

The cases and their results are listed in `TEST_REPORT.md`.

## 6. Pass criteria

All tests pass. A failing test is either a defect in the tool (fix it and rerun) or a defect in the test (fix it and say so in the report). Nothing is marked as passed by weakening an expectation without a reason.

## 7. Limits of this test

The scenario author, the config author and the runner are the same agent. The answer key was written first to reduce bias, but a human architect should still review the scenario and the scores. The documents are shorter and cleaner than real ones, so passing here does not prove the method handles messy real-world files. LibreOffice is the calculation engine, not Excel.

## 8. How to extend

Add a document with a new trap, add its expected handling to section 4.1, then add a test with the next free ID in `run_tests.py`.
