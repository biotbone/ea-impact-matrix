# Proposal P2: Retire the on-premises file servers

Sponsor: Priya Narang, Head of Infrastructure. Delivery lead: K. Dunmore. Architect: R. Alder.

## Summary
Move the six on-premises file servers to a managed file storage service and decommission the hardware. The servers run Windows Server 2012 R2, whose extended support ends on 2027-01-31. A pilot with the finance share completed in January 2026 with no issues over four weeks.

## Scope
In: data migration, permission mapping, parallel run, cutover, decommission.
Out: the document management system, desktop client changes.

## Benefits
- Removes an operating system that will be unsupported after 2027-01-31.
- Closes the open audit finding on unsupported operating systems (see SECURITY).
- Removes the recurring disk array failures (three unplanned outages in six months, see INCIDENTS).
- Lowers run cost: today 9,400 a month, expected 3,900 a month on managed storage.

## Risks
- Data loss or permission errors at cutover.
- Cutover window is short for the warehouse shares.

## Duration
Four months, April to July 2026. Cost and effort are in ESTIMATE-P2.

## Questions for the client
- What is the acceptable cutover window for the warehouse shares?
- How long must the old shares be retained after cutover?
