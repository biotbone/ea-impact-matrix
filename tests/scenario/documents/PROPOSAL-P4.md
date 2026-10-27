# Proposal P4: Containerize the order-management application

Sponsor: Tomas Ellery, Head of Applications. Delivery lead: L. Hartwell. Architect: R. Alder.

## Summary
Move the order-management application from its current managed platform service to containers on a self-managed Kubernetes cluster built on virtual machines, to avoid platform lock-in and reduce cost.

## Scope
In: cluster build, application refactor, migration, training and handover.
Out: the database tier, the customer-facing web front end.

## Benefits claimed
- Run cost down 30%.
- Release lead time down from 14 days to 5 days.
- Improved security through standardisation.

## Design notes
- Service credentials are provisioned as keys stored on the cluster nodes.
- Operating system patching of the nodes is handled monthly by the platform team.
- Today the application uses the managed service's identity features and the provider patches the platform.

## Risks
- Team has not run a production cluster before.
- Scope of the refactor is not fully known.

## Duration
Five months, June to October 2026. Cost and effort are in ESTIMATE-P4.

## Questions for the client
None.
