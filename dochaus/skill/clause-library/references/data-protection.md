# Data protection

**Topic:** Data protection — roles, breach notice, subprocessors, transfers.
**Contract type:** Any agreement under which one party processes personal data
for the other.
**Jurisdiction:** US commercial market scope, but this family is the most
regime-dependent of the twelve — required terms differ materially between the
EU/UK regimes, US state privacy laws, and sectoral rules. Identify which
regimes apply before judging any clause; for GDPR/UK GDPR work use the DPA
playbook. Positions are firm-configurable defaults overridable via
`.dochaus/profile.md`.

## Preferred

Roles are expressly assigned (who determines purposes and means, who
processes on instructions). The processing party processes personal data only
on documented instructions and only to provide the service; it implements
stated technical and organizational security measures, notifies the other
party of a personal-data breach without undue delay, assists with data-subject
requests and compliance obligations, binds subprocessors to equivalent terms
(with a published list and a notice-and-objection mechanism), supports audits
or supplies recognized audit reports, deletes or returns data at the end of
the engagement, and uses a lawful transfer mechanism for any cross-border
transfer that requires one.

```example
Provider shall process Personal Data only on Customer's documented
instructions and solely to provide the Service; implement the technical and
organizational measures described in Annex [Security]; notify Customer
without undue delay after becoming aware of a Personal Data Breach; engage
subprocessors only under written terms no less protective than this DPA,
maintaining a current subprocessor list and giving Customer prior notice of
additions with a right to object; and, at Customer's election on termination,
delete or return all Personal Data except as retention is required by law.
```

**Off-market variants to watch.** Processor-favorable: rights to use personal
data for its own purposes (analytics, product improvement on identifiable data,
marketing) outside instructions; breach notice diluted to a long fixed window,
conditioned on confirmed impact, or omitted; no subprocessor list, flow-down,
or objection right; security stated as "industry standard" with no annex,
certification, or audit mechanism. Controller-favorable: unlimited on-site
audit rights at any time at the provider's cost where a recognized audit report
would serve; breach-notice windows operationally impossible, paired with
uncapped liability for any incident regardless of fault; instructions defined so
broadly any request becomes a contractual processing obligation.

## Fallbacks (ordered)

1. **Offer when** the provider resists on-demand on-site audits: accept audit
   rights satisfied by recognized third-party reports, with on-site audit
   reserved for cause (after a breach or regulator inquiry).
2. **Offer when** the provider needs subprocessor flexibility: accept a general
   subprocessor authorization with a published list, change notice, and an
   objection right whose remedy is termination of the affected service with a
   pro-rata refund.
3. **Offer when** the provider wants usage analytics: accept
   aggregated/de-identified usage-data rights only where re-identification is
   contractually prohibited and the de-identified definition meets the
   applicable regime's standard.

## Walk-away floor

Own-purpose use of identifiable personal data, or no breach-notice obligation
of any kind.

## Escalation trigger

At or below the floor, do not concede. Route to the **supervising lawyer** (or
the firm's privacy/data-protection lead, or the `.dochaus/profile.md`
Escalation Rules contact) — data-protection terms carry regulatory exposure the
liability clause may not cap.

**Rationale:** Data-protection terms carry regulatory exposure that the
liability clause may not cap; roles, breach notice, and transfers are the
items regulators and counterparties test first.
