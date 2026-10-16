---
name: contract-risk-checklist
description: The canonical 12-area risk checklist for commercial contract review, with the red-flag patterns that mark each area as one-sided or dangerous. Use when reviewing any contract for risk — an NDA, MSA, SaaS agreement, services or consulting agreement, license, vendor terms, or order form — or when asked to spot issues, find red flags, assess a draft, mark up a counterparty's paper, or answer "what's wrong with this contract."
---

Review every contract against all twelve areas below. Before you start, turn the
twelve areas into an explicit checklist with the `todowrite` tool — one item per
area — and answer each item present / absent / deviating rather than reviewing
from memory, because a single skipped area can let an off-market term through.
Mark an area `completed` only once you have located its governing clause (or
confirmed it has none); an unresolved area stays `in_progress`. For each area:
state what the contract says, cite the governing clause as
`[<Document> § <section>]` (or note its absence — a missing clause is a finding,
not a pass), and flag anything matching the red-flag patterns. Establish the governing-law clause
first and qualify jurisdiction-sensitive findings against the matter's
jurisdiction pack. For balanced market positions and fallback ladders per
area, load the clause-library skill. A mapping of CUAD's 41-category clause
taxonomy onto these twelve areas is in `references/cuad-taxonomy.md`.

<checklist>

## 1. Liability

Inspect: caps, exclusions, carve-outs, mutuality, super-caps for sensitive
exposure.

What bad looks like:
- Carve-outs from the cap so broad they swallow it — "any breach of this
  Agreement," all indemnification obligations, or all confidentiality breaches
  uncapped in one direction.
- One party's liability uncapped while the other's is capped.
- A cap far below realistic exposure (one month's fees) with no separate
  super-cap where the deal involves sensitive data or high-value IP.

## 2. Indemnification

Inspect: scope, triggers, who indemnifies whom, defense and settlement
control, relationship to the liability cap.

What bad looks like:
- One-way indemnity in a deal where both sides create risk.
- Indemnity triggered by "any breach of this Agreement," converting every
  contract claim into an uncapped indemnified claim.
- Indemnitee controls defense and may settle at the indemnitor's cost without
  the indemnitor's consent.
- Indemnity covering first-party (direct) claims, not just third-party claims;
  or no IP-infringement indemnity from the provider in a technology deal.

## 3. Termination

Inspect: for cause, for convenience, notice periods, cure rights, renewal
mechanics, effect on fees and data.

What bad looks like:
- Evergreen auto-renewal with a short or early opt-out window (non-renewal
  notice due 90+ days before renewal) that is easy to miss.
- One party may terminate for convenience but the other cannot; or cure
  periods granted to only one side.
- Termination accelerates all remaining committed fees regardless of cause.
- No wind-down, transition assistance, or data-return obligation on exit.

## 4. Payment

Inspect: amounts, schedule, late fees, set-off, escalation, currency, taxes,
refundability.

What bad looks like:
- Unilateral price escalation — fees change on notice with no cap and no
  corresponding termination right.
- All fees non-refundable even where the provider is in breach; or payment
  obligations that survive termination regardless of cause.
- Set-off prohibited for one party while the other keeps broad set-off or
  withholding rights.
- All taxes and gross-up obligations shifted to one party.

## 5. Intellectual property

Inspect: ownership, license scope, work product, feedback and data rights,
residuals.

What bad looks like:
- Assignment of pre-existing IP, or of improvements to it, to the other side.
- Customer-paid work product owned by the provider, or joint ownership with no
  exploitation and accounting rules.
- A license to customer data or feedback broader than needed to operate the
  service (perpetual, irrevocable, "any purpose").
- License grants that are silently perpetual or irrevocable where the
  commercial intent is term-limited.

## 6. Confidentiality

Inspect: definition, duration, exclusions, permitted disclosures,
return/destruction, adjacent restrictive covenants (non-solicit, non-compete,
exclusivity).

What bad looks like:
- One-way obligations where both parties disclose.
- Marking-only definition with no reasonable-person catch-all; or a residuals
  clause permitting use of anything retained in unaided memory.
- Compelled-disclosure permission with no notice or minimization obligation.
- Restrictive covenants smuggled in: no-hire (vs. no-solicit) provisions,
  non-competes, or exclusivity beyond the deal's scope — enforceability is
  highly jurisdiction-dependent; qualify against the matter's pack.

## 7. Warranties

Inspect: express warranties, disclaimers, duration, remedies for breach.

What bad looks like:
- "As is" disclaimer that reaches the core service commitment itself.
- Warranty remedy that is illusory — sole remedy is re-performance at the
  warrantor's discretion with no refund backstop.
- One-sided compliance-with-laws or non-infringement warranties.
- Disclaimers of implied warranties without the conspicuousness or scope
  limits some jurisdictions require — qualify; do not assert a specific rule.

## 8. Limitation of remedies

Inspect: exclusive remedies, consequential-damages waivers, liquidated
damages, who may enforce.

What bad looks like:
- Consequential-damages waiver applied one-way.
- Waiver listing "lost profits" without limiting it to consequential losses —
  lost profits can be direct damages, so this quietly guts direct recovery.
- An exclusive remedy that can fail of its essential purpose with nothing
  behind it.
- Liquidated damages that look punitive rather than a genuine pre-estimate of
  loss — enforceability varies by jurisdiction; qualify.

## 9. Assignment and change of control

Inspect: consent requirements, deemed-assignment language, change-of-control
triggers, affiliate and sale-of-business carve-outs.

What bad looks like:
- Consent withholdable in sole discretion, combined with a deemed-assignment
  clause making any change of control a breach — a hidden M&A blocker.
- One side assigns freely while the other needs consent.
- Termination right triggered by a change of control of one party only.
- No carve-out for assignment to affiliates or in connection with a sale of
  substantially all assets.

## 10. Governing law and dispute resolution

Inspect: governing law, venue, arbitration, jury waiver, fee shifting,
equitable relief.

What bad looks like:
- Counterparty home-court stack: their governing law, exclusive venue in their
  forum, and a jury waiver together.
- Mandatory arbitration with no carve-out to seek injunctive relief for
  confidentiality or IP breaches.
- One-way fee shifting (loser pays only when one side wins).
- Shortened limitations periods or class waivers — enforceability varies by
  jurisdiction; qualify.

## 11. Force majeure

Inspect: covered events, obligations during the event, notice and mitigation,
exit rights.

What bad looks like:
- Coverage of "economic hardship," market changes, or supplier cost increases
  — turning ordinary business risk into an excuse for non-performance.
- Only one party's performance excused.
- No right to terminate after a prolonged event (commonly 30-90 days).
- No notice or mitigation obligation on the party invoking it; or payment
  obligations excused alongside performance.

## 12. Data protection

Inspect: controller/processor roles, security obligations, breach notice,
subprocessors, international transfers.

What bad looks like:
- Processing roles unassigned, or a processor granted rights to use personal
  data for its own purposes.
- Breach notification replaced by a long fixed window or omitted entirely.
- No subprocessor flow-down, approval, or audit mechanism.
- International transfers unaddressed where data will cross borders — the
  required mechanism depends on the jurisdictions involved; check the matter's
  jurisdiction pack rather than assuming one regime.

</checklist>

These red-flag patterns are US commercial-market reference points, not legal
standards; verify jurisdiction-sensitive items against the matter's
jurisdiction pack before presenting them as findings.

<severity>
Grade each area with one of four severities so the review reads as a triaged
register, not a flat list. Use the same ladder for every area; the area's
red-flag patterns above tell you what "off-market" looks like, this tells you how
hard to push:
- **critical** — off-market or missing in a way that exposes the client to
  uncapped, unrecoverable, or deal-breaking risk (a cap-swallowing carve-out, a
  one-way uncapped indemnity, no IP-infringement indemnity in a technology deal).
  Must be fixed before signing.
- **high** — a materially unfavorable term, or a protection a contract of this
  type should always carry but does not (no termination-for-cause cure right, no
  data-return on exit). Fix, or record an explicit business decision to accept it.
- **medium** — a negotiable deviation worth raising that the client can live with
  if traded for something else.
- **low** — cosmetic, market-standard with a minor wrinkle, or favorable to the
  client; logged so coverage of the area is auditable.

When a term sits between two severities, take the higher one — an over-graded
term costs the lawyer seconds to dismiss; an under-graded one goes unread into the
signed agreement. An entirely-absent protection is graded on the exposure its
absence creates, not downgraded for being a gap rather than bad language.
</severity>
