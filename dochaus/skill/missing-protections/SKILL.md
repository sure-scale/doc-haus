---
name: missing-protections
description: A checklist of protections that should be present in a contract but are absent entirely — the gaps a clause-by-clause review misses because there is no clause to flag. Use when reviewing an NDA, services agreement or MSA, DPA, or employment agreement and asked what is missing, what gaps exist, or whether the document adequately protects a party.
---

Reviewing what a contract says is not the same as reviewing what it omits. A
one-sided clause is visible on the page; a missing limitation of liability is
invisible until it matters. After any clause-level review, walk the checklist
for the contract's type and report every expected protection that is absent.

These checklists are commercial-market reference points, not legal standards —
what is "expected" varies by jurisdiction, deal size, and bargaining position.
Establish the governing law and which party the firm represents first, and
qualify any finding that depends on either.

Method: build the checklist for the contract's type as an explicit list with
the `todowrite` tool — one item per expected protection — so the absence sweep
is exhaustive and no expected protection is skipped (recall is what this sweep
exists to protect). For each item, search the whole document — protections hide
in schedules, definitions, and annexes, not just the obvious section — then mark
the item `completed` with its verdict (present / partial / missing). Report an
item as missing only when no provision anywhere covers it; if a provision
partially covers it, that is a weak-clause finding for the clause review, not
a gap. Cite the absence as `[<Document> — no provision found]` and cite
partial coverage as `[<Document> § <section>]`.

<nda>
- **Return or destruction of confidential information** — without it, the
  recipient may keep disclosed materials indefinitely after the relationship
  ends. Ask for: an obligation to return or destroy on request or termination,
  with certification on request.
- **Term and survival of confidentiality** — with no stated duration, the
  parties dispute how long protection lasts; trade secrets typically need
  protection for as long as they remain secret. Ask for: a defined
  confidentiality period, with trade secrets protected for as long as they
  qualify.
- **Standard exclusions** — without the customary carve-outs (already known,
  independently developed, publicly available, rightfully received from a
  third party), the recipient breaches by using information it already had.
  Ask for: the standard exclusions from the definition of confidential
  information.
- **Compelled-disclosure procedure** — without it, a subpoena forces the
  recipient to choose between breaching the NDA and defying legal process.
  Ask for: permission to disclose when legally required, with prompt notice
  to the discloser where lawful and cooperation on protective measures.
- **No-license / no-obligation clause** — without it, the counterparty may
  argue the disclosure implied IP rights or a duty to transact. Ask for: a
  statement that no license is granted and neither party is obliged to
  proceed with any transaction.
- **Equitable relief acknowledgment** — damages are hard to prove for
  information leaks; without it, stopping an imminent disclosure is slower.
  Ask for: acknowledgment that breach may cause irreparable harm and that the
  discloser may seek injunctive relief (availability is jurisdiction- and
  court-dependent).
</nda>

<services-msa>
- **Limitation of liability** — an MSA with no liability cap exposes each
  party to uncapped damages for any breach; a cap is standard commercial
  practice in most markets. Ask for: a mutual cap (commonly tied to fees
  paid or payable over a defined period) with negotiated carve-outs.
- **Consequential-damages exclusion** — without it, lost profits and other
  indirect losses are in play for ordinary breaches. Ask for: a mutual
  waiver of indirect and consequential damages, subject to the same
  carve-outs as the cap.
- **Indemnification** — with no indemnity, third-party claims caused by one
  party (IP infringement by deliverables, bodily injury, breach of law) land
  on whichever party gets sued. Ask for: indemnities matched to each party's
  risks, at minimum provider IP-infringement indemnity for the deliverables.
- **Intellectual property ownership** — silence on work product leaves
  ownership to default law, which differs by jurisdiction and often surprises
  the paying party. Ask for: an express allocation — assignment or license of
  deliverables, each party retaining its pre-existing IP.
- **Termination rights** — a contract with no termination clause may lock the
  parties in for the full term regardless of nonperformance. Ask for:
  termination for material breach with notice and cure, and the negotiated
  position on termination for convenience.
- **Effect of termination** — without wind-down terms, the parties dispute
  fees, data return, and transition at the worst possible moment. Ask for:
  payment for work performed, return of materials and data, and survival of
  the clauses meant to outlast the contract.
- **Warranties and service standards** — with no performance warranty, the
  customer's remedy for bad work is a general breach claim with no defined
  standard. Ask for: a workmanlike-performance or conformance-to-spec
  warranty with a defined remedy.
- **Insurance** — without required coverage, an indemnity is only as good as
  the counterparty's balance sheet. Ask for: insurance requirements
  proportionate to the engagement's risk.
- **Confidentiality** — an MSA with no confidentiality clause leaves pricing,
  data, and know-how exchanged under it unprotected. Ask for: a mutual
  confidentiality clause or incorporation of an existing NDA.
- **Dispute resolution and governing law** — silence forces a conflict-of-laws
  fight before the merits. Ask for: express governing law and forum (or
  arbitration), chosen deliberately rather than defaulted.
</services-msa>

<dpa>
- **Breach-notification deadline** — "without undue delay" alone gives the
  controller nothing to enforce, while the controller may face fixed
  regulatory deadlines (under GDPR/UK GDPR regimes, the controller generally
  must notify its authority within a short fixed window). Ask for: processor
  notification within a defined short period after becoming aware of a
  personal-data breach, with cooperation and required details.
- **Processing instructions limitation** — without it, the document fails the
  basic controller-processor allocation. Ask for: processing only on the
  controller's documented instructions, with a notice obligation if an
  instruction appears unlawful.
- **Subprocessor controls** — without consent or notice rights, the controller
  cannot manage where its data goes. Ask for: prior authorization (general or
  specific) for subprocessors, an obligation to flow down equivalent terms,
  and processor liability for subprocessor acts.
- **Security measures** — a DPA with no security schedule has no enforceable
  baseline. Ask for: defined technical and organizational measures
  appropriate to the risk, referenced in an annex.
- **International-transfer mechanism** — under GDPR/UK GDPR-style regimes,
  transfers to third countries need a lawful mechanism; silence leaves the
  controller non-compliant. Ask for: an approved transfer mechanism (such as
  standard contractual clauses or an equivalent) where transfers occur.
- **Audit and information rights** — without them, the controller cannot
  demonstrate compliance to its regulator. Ask for: information and audit
  rights, commonly satisfied first by certifications or third-party reports.
- **Data-subject-request assistance** — the controller answers requests on
  statutory clocks but the processor holds the data. Ask for: an obligation
  to assist with data-subject requests and to forward requests received
  directly.
- **Return or deletion on termination** — without it, personal data persists
  after the engagement with no lawful basis. Ask for: deletion or return of
  personal data at the controller's choice on termination, with certification.
- **Confidentiality of processing personnel** — without it, the processor's
  staff are outside the protection chain. Ask for: a commitment that persons
  processing the data are bound by confidentiality obligations.
</dpa>

<employment>
- **Confidentiality and trade-secret protection** — without it, the employer
  relies solely on background law, which varies by jurisdiction and is harder
  to enforce than an express obligation. Ask for: a confidentiality clause
  surviving termination.
- **IP and invention assignment** — without express assignment, ownership of
  employee-created work depends on default rules that differ by jurisdiction
  and work type. Ask for: assignment of work-related inventions and works,
  subject to any statutory employee-invention protections in the governing
  jurisdiction.
- **Termination notice and severance terms** — silence leaves both parties to
  statutory or at-will defaults that may not match expectations. Ask for:
  express notice periods or severance terms (statutory minimums in many
  jurisdictions override anything less).
- **Post-termination restrictive covenants** — if the role justifies them and
  none exist, the employer has no protection against immediate competition or
  solicitation. Ask for: non-solicitation (and non-compete only where the
  governing jurisdiction enforces them — enforceability varies sharply and
  several jurisdictions restrict or void them), drafted to local
  reasonableness standards.
- **Compensation completeness** — an offer silent on bonus terms, equity
  vesting, or benefits invites disputes over what was promised. Ask for:
  express terms for every compensation element referenced or promised,
  including treatment on termination.
- **Dispute-resolution terms** — silence defaults to local courts and rules;
  arbitration clauses and jury waivers are enforceable in some jurisdictions
  and restricted in others. Ask for: a deliberate choice, validated against
  the governing jurisdiction's employment-law limits.
</employment>

<output>
Report gaps separately from clause-level findings, one row per missing
protection:

| Protection | Status | Why it matters here | Suggested request |
|---|---|---|---|

Status is **missing** (no provision anywhere) or **partial** (cite the
section and hand it to the clause review). "Why it matters here" must be
specific to this contract and party, not the generic rationale above. The
suggested request is the one-line ask to send the counterparty. Order rows by
consequence to the represented party, and mark any item whose importance or
enforceability depends on the governing law as jurisdiction-dependent.
</output>
