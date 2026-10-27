# Standards crosswalk

How the impact matrix lines up with TOGAF and COBIT 2019, so you can use the client's own vocabulary and show an auditor or an architecture board where each part of the method sits.

This is the author's working crosswalk, not an official mapping. It names sections and objective IDs only and reproduces no text from either standard. Check the current editions before quoting it. TOGAF is a registered trademark of The Open Group. COBIT is a registered trademark of ISACA. This project is not affiliated with or endorsed by either organization, and using the crosswalk is alignment, not certification.

## Vocabulary

| In the matrix | TOGAF term | COBIT 2019 anchor |
|---|---|---|
| Initiative | Work package or project (Phases E and F) | Program or project in the portfolio (APO05, BAI01) |
| Prioritization lens and weights | Prioritization criteria agreed with stakeholders (Phase F) | Target investment mix and selection of programs to fund (EDM02.01, APO05.02) |
| Wave | Transition architecture | Program plan (BAI01.04), architecture implementation (APO03.04) |
| Wave entry condition | Conditions in the implementation and migration plan | Stage gate of a program (BAI01) |
| Cross-project dependencies | Consolidated gaps, solutions and dependencies matrix (Phase E) | Program plan dependencies (BAI01.04) |
| Resource contention | Implementation factors and constraints (Phase E) | Resource optimization (EDM04) |
| Risk register | Risk management technique, initial and residual level of risk | Risk analysis, profile and response (APO12.02, APO12.03, APO12.06), risk optimization (EDM03) |
| Governance gates | Architecture compliance review, architecture contract (Phase G, governance part) | Performance and conformance monitoring (MEA01, MEA02) |
| Open questions | Assumptions and gaps, architecture requirements management | Requirements definition (BAI02), closest fit |
| KPIs, savings case | Cost and benefit assessment (Phase F) | Budget and cost (APO06), monitor value (EDM02.04), benefits achievement (APO05.05) |
| Decision log | Architecture board decisions, architecture contract | Direct value optimization (EDM02.03), report on the portfolio (APO05.03) |
| Next review date | Architecture change management (Phase H) | Maintain portfolios (APO05.04) |
| Required documents and the register | Inputs to Phases E and F (request for architecture work, draft roadmap, implementation factor catalog) | Funds and constraints (APO05.01), requirements definition (BAI02) |

## The seven dimensions

| Dimension | TOGAF concept | COBIT 2019 anchor |
|---|---|---|
| Business value | Business value of the change, business architecture capability | EDM02, APO05 |
| Cost optimization and savings | Cost and benefit assessment | APO06 |
| Time criticality | Business constraints for implementation, drivers | APO05.02 |
| Risk reduction and resilience | Risk management | APO12, EDM03 |
| Security and compliance | Security architecture, compliance | APO13, MEA03 |
| Delivery confidence | Business transformation readiness assessment | BAI01, EDM04 |
| Strategic and architecture alignment | Architecture principles, target architecture | APO03 |

## The method

| Step | TOGAF | COBIT 2019 |
|---|---|---|
| 1 Read the sources | Architecture requirements and baseline | APO05.01 (funds and constraints) |
| 2 Confirm the lens | Agree prioritization criteria | EDM02.01, APO05.02 |
| 3 Score with evidence | Gap analysis, trade-offs | APO05.02 |
| 4 Add the money | Cost and benefit assessment | APO06 |
| 5 Find dependencies | Dependencies matrix | BAI01.04 |
| 6 Build the wave plan | Transition architectures, roadmap | APO03.04 |
| 7 Log what is missing | Requirements management | BAI02 |
| 8 Record the decision | Board decision, architecture contract | EDM02.03, APO05.03 |

## What was added to follow the standards

These fields are optional. Old configs keep working without them.

| Addition | Why |
|---|---|
| Lens approver and date, risk appetite, risk threshold, next review | The agreed criteria are recorded, a score has an accountable owner, and the portfolio is revisited |
| Response and residual score per risk | Risk is described before and after the response, as risk management practice expects |
| Section 9 on each project sheet: capability, principles, gap, benefits owner, benefits review date, standards references | Traceability from the initiative to the capability it changes and to the person who answers for the benefit |
| Gate 10, architecture decision recorded | Every initiative leaves a decision trail |
| Decision log sheet | The scores recommend and the forum decides, so the decision needs its own record |
| Document catalog, register and intake list | The workbook only holds what its sources support, so the required inputs are defined, tracked and checked |
| Dropdowns and the `--check` command | Fewer invalid entries, and the method's own rules enforced before the workbook is shared |

## What was deliberately left out

The dimension model, the 1 to 5 scale, the single JSON config and the fixed cell addresses stay as they are, because they are what keeps the tool quick to use. Capability-level scoring from COBIT, full compliance assessment levels and agile weighted-shortest-job-first scoring were considered and left out. Add them if a client's governance requires them.
