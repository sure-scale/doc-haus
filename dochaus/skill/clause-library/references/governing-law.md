# Governing law and dispute resolution

**Topic:** Governing law and dispute resolution — law, venue, arbitration,
equitable relief.
**Contract type:** All commercial agreements.
**Jurisdiction:** US commercial market; this family is inherently
jurisdiction-specific — never present a forum or law choice as correct without
checking the matter's jurisdictions and where the parties and assets actually
are. Positions are firm-configurable defaults overridable via
`.dochaus/profile.md` (which carries the firm's default governing law and
venue).

## Preferred

A named, neutral governing law without regard to conflict-of-laws principles,
with exclusive venue in that jurisdiction's courts (or institutional
arbitration seated there), and a carve-out preserving either party's right to
seek injunctive relief for confidentiality or IP breaches in any court of
competent jurisdiction. Each party bears its own fees unless a statute or the
tribunal awards them.

```example
This Agreement is governed by the laws of [neutral jurisdiction], without
regard to its conflict-of-laws principles. The parties consent to the
exclusive jurisdiction and venue of the courts located in [venue] for any
dispute arising out of or relating to this Agreement, except that either
party may seek injunctive or other equitable relief in any court of competent
jurisdiction for breach of Section [Confidentiality] or infringement of its
intellectual property rights.
```

**Off-market variants to watch.** Drafter-favorable: the home-court stack
(drafter's home law, exclusive home venue, and a jury waiver together);
mandatory arbitration with no equitable-relief carve-out; one-way fee shifting.
Non-drafter-favorable: floating venue ("courts of the plaintiff's choosing") or
no exclusive forum at all; carve-outs so broad the agreed forum becomes
optional; refusal of any jury or class waiver where those are ordinary for the
deal type (availability and validity of such waivers varies by jurisdiction;
qualify).

## Fallbacks (ordered)

1. **Offer when** the counterparty rejects a fully neutral stack: accept the
   counterparty's home law with venue in a neutral forum, or vice versa — split
   the stack rather than conceding all of it.
2. **Offer when** the counterparty prefers arbitration: accept institutional
   arbitration (named rules, named seat, English language, one or three
   arbitrators) provided the equitable-relief carve-out survives.
3. **Offer when** the client has minimal litigation exposure under the deal:
   accept the counterparty's home law and forum, documented in the rationale.

## Walk-away floor

Arbitration with no equitable-relief exception, or one-way fee shifting. The
equitable-relief carve-out should survive every trade.

## Escalation trigger

At or below the floor, do not concede. Route to the **supervising lawyer** for
the matter (or the `.dochaus/profile.md` Escalation Rules contact). A non-US
governing law, or a law outside the jurisdictions the firm advises in, is
always escalated rather than improvised.

**Rationale:** The dispute clause prices every other clause: a strong
position in an unreachable forum is worth little. The equitable-relief
carve-out is the one term that should survive every trade.
