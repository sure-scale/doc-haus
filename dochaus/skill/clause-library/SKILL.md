---
name: clause-library
description: Reference library of balanced market positions for the twelve clause families reviewed in every commercial contract, each on a fixed per-clause schema (topic, contract-type, jurisdiction tag, Preferred verbatim text, ordered fallbacks with offer-when conditions, an explicit walk-away floor, rationale, and a role-addressed escalation trigger at the floor). Use when judging whether a clause is standard or one-sided, proposing alternative language, building a fallback position, or answering "what is market" for liability caps, indemnities, termination, payment, IP, confidentiality, warranties, remedies, assignment, governing law, force majeure, or data protection.
---

These are US commercial-market reference points, not legal standards — verify
against the matter's jurisdiction pack before relying on any position, and
qualify jurisdiction-sensitive judgments (restrictive covenants, warranty
disclaimers, liquidated damages, data transfers) explicitly.

Each clause family has one reference file on the fixed schema below. The
positions are **market-balanced defaults, not aspirational** — they describe
where a fair two-sided deal lands, not the most favorable position one side
could demand. They are also **firm-configurable defaults**: the firm's
`.dochaus/profile.md` (see the `firm-profile` skill) overrides any Preferred,
fallback, floor, or escalation contact here. When the profile speaks, it wins;
when it is silent, these defaults apply.

<schema>
Every `references/*.md` file states one clause family on these fields, in this
order. The schema is fixed so the redline tool can insert the Preferred text
byte-exact, the deterministic lint can detect deviation, and `document-reviewer`
can band each clause it finds.

- **Topic / Contract type / Jurisdiction** — a one-line tag header: the clause
  family, the contract types it applies to, and the jurisdiction scope (these
  files are US-commercial-market scope).
- **Preferred** — the firm's default opening position, with paste-ready
  verbatim language in a fenced `example` block. The fence is illustrative
  market language for comparison and as a drafting starting point; firm-approved
  replacement text for a specific document type lives in that type's playbook,
  which controls over this file.
- **Fallbacks (ordered)** — the acceptable rungs below Preferred, in descending
  order, each prefixed with an **Offer when** condition stating the
  counterparty posture or trade that justifies dropping to it. Stop at the
  highest rung whose condition is met.
- **Walk-away floor** — the position below which the firm does not concede
  without a human decision. Stated explicitly so the reviewer can tell a
  cosmetic deviation from a floor-breaching one.
- **Escalation trigger** — what to do when a term sits at or below the floor:
  addressed to a **role**, not a named person, so the position outlasts any
  individual assignment, and routed through the firm's `.dochaus/profile.md`
  Escalation Rules contact when one is set.
- **Rationale** — why the family is negotiated where it is and which sub-term
  moves the outcome most.
</schema>

<families>
- **Liability** — caps, carve-outs, super-caps. `references/liability.md`
- **Indemnification** — scope, triggers, defense control. `references/indemnification.md`
- **Termination** — cause, convenience, renewal, wind-down. `references/termination.md`
- **Payment** — escalation, set-off, refunds, taxes. `references/payment.md`
- **Intellectual property** — ownership, licenses, work product, feedback. `references/intellectual-property.md`
- **Confidentiality** — definition, term, exclusions, restrictive covenants. `references/confidentiality.md`
- **Warranties** — express warranties, disclaimers, remedies. `references/warranties.md`
- **Limitation of remedies** — exclusive remedies, damages waivers, liquidated damages. `references/remedies.md`
- **Assignment and change of control** — consent, deemed assignment, carve-outs. `references/assignment.md`
- **Governing law and dispute resolution** — law, venue, arbitration, equitable relief. `references/governing-law.md`
- **Force majeure** — covered events, notice, exit rights. `references/force-majeure.md`
- **Data protection** — roles, breach notice, subprocessors, transfers. `references/data-protection.md`
</families>

The families align one-to-one with the contract-risk-checklist areas; use that
skill to find problems, this one to judge severity and propose positions.

When a clause materially departs from the Preferred position, flag it with a
citation, identify which fallback rung (if any) it matches, and apply the
banding the reviewer uses: at or above the highest met rung is acceptable;
between rungs is a negotiable deviation; at or below the walk-away floor fires
the escalation trigger.

<jurisdiction-risk>
These positions are US commercial-market reference points; their enforceability
is a jurisdiction question, and model confidence does not track correctness. Let
the structural jurisdiction signals — not certainty — decide when to qualify.

- **Name the assumed governing law first.** Before banding any
  jurisdiction-sensitive family (restrictive covenants and confidentiality,
  warranty disclaimers, liquidated damages and remedies, governing law and
  dispute resolution, data protection and transfers), state the governing law
  you are reasoning under and where it comes from — the clause's choice-of-law
  term or the matter's jurisdiction pack. Never accept an implied jurisdiction:
  if a "what is market" question names none, surface the gap and state what you
  assumed rather than defaulting to US-commercial scope.
- **Grade confidence down — qualify, then escalate** — when the governing law is
  state/local, non-US, lower-court-driven or unsettled, multi-jurisdiction, or
  recently changed. A position that is market and enforceable under US law may be
  unenforceable elsewhere (e.g. liquidated-damages-as-penalty rules, broad
  warranty disclaimers, cross-border data transfers).
- **Governing law vs forum, and mandatory-law override.** Read the
  governing-law family against its own forum/seat term; flag misalignment, and
  flag where the forum's mandatory rules or public policy (consumer, competition,
  employee, data-protection) can override the chosen law and the position it
  would otherwise support.
- **Cross-border CISG prompt.** For a cross-border sale-of-goods contract, check
  whether the CISG applies by default and whether the parties intended to
  exclude it; flag an unaddressed CISG position before relying on these
  US-market warranty/remedy reference points.
- **Multi-jurisdiction.** When more than one governing law is in scope, band
  each clause separately and completely under each law, then reconcile —
  flagging where a rung that is acceptable under one law breaches the floor or is
  void under another, and routing divergences that need a decision to escalation.
</jurisdiction-risk>
