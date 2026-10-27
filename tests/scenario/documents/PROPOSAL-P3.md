# Proposal P3: Consolidate three monitoring tools into one platform

Sponsor: Priya Narang, Head of Infrastructure. Delivery lead: K. Dunmore. Architect: R. Alder.

## Summary
Replace monitoring tools A, B and C with one consumption-billed observability platform. Alerts, dashboards and the on-call rota move to the new platform.

## Scope
In: platform onboarding, agent rollout, alert migration, training.
Out: application performance tracing, log retention redesign.

## Benefits
- One alerting path instead of three. Two incidents in the last six months were missed because an alert fired in one tool only (see INCIDENTS).
- Lower tooling spend. The vendor says "up to 45% lower", depending on ingest volume. The expected range is 10% to 45% of today's 10,000 a month.

## Risks
- The platform bills by data ingested, so cost grows with volume and the estimate does not include volume added by other initiatives.
- Alert migration gaps during the switch.

## Duration
Three months. Cost and effort are in ESTIMATE-P3.

## Questions for the client
- What ingest volume should be assumed?
- Tool C renews on 2026-12-15 with 90 days' notice. Is that the decision date?
