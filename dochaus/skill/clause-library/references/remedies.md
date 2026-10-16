# Limitation of remedies

**Topic:** Limitation of remedies — exclusive remedies, damages waivers,
liquidated damages.
**Contract type:** All commercial agreements.
**Jurisdiction:** US commercial market; verify against the matter's
jurisdiction pack — the enforceability of damages waivers, liquidated damages,
and shortened limitations periods varies by jurisdiction. Positions are
firm-configurable defaults overridable via `.dochaus/profile.md`.

## Preferred

A mutual waiver of consequential, incidental, special, and punitive damages,
with lost profits excluded only to the extent they are consequential — direct
damages remain recoverable. The waiver carries the same carve-outs as the
liability cap (confidentiality, indemnity obligations, gross negligence,
willful misconduct). Exclusive remedies are used sparingly and always with a
backstop if the remedy fails of its essential purpose. Equitable relief is
preserved for confidentiality and IP breaches.

```example
Except for Excluded Claims, neither party shall be liable to the other for
any consequential, incidental, indirect, special, or punitive damages, or for
lost profits to the extent they constitute consequential damages, arising out
of or relating to this Agreement, even if advised of the possibility of such
damages. Nothing in this Section limits either party's right to seek
injunctive or other equitable relief for breach of Section [Confidentiality]
or infringement of its intellectual property.
```

**Off-market variants to watch.** Provider-favorable: one-way waiver (provider's
consequential damages excluded while the customer's exposure is unlimited in
kind); "lost profits, lost revenue, loss of data" listed without the "to the
extent consequential" qualifier, quietly excluding direct damages; an exclusive
remedy (service credits only) with nothing behind it when the remedy fails of
its essential purpose. Customer-favorable: damages waiver deleted in one
direction or every customer claim carved out; liquidated damages operating as
penalties; third-party-beneficiary or covenant-not-to-sue provisions giving one
side's affiliates enforcement rights with no reciprocity.

## Fallbacks (ordered)

1. **Offer when** the counterparty wants a flat lost-profits exclusion: accept
   an unqualified lost-profits exclusion only if direct damages are expressly
   preserved elsewhere in the section.
2. **Offer when** the provider proposes SLA-credit remedies: accept exclusive
   remedies (e.g., SLA credits for availability) where a termination-plus-refund
   backstop exists for chronic failure.
3. **Offer when** the counterparty proposes liquidated damages: accept them only
   where the amount is a documented, genuine pre-estimate of loss and is the
   exclusive remedy for that breach alone.

## Walk-away floor

A one-way damages waiver, or an exclusive remedy with no backstop when it fails
of its essential purpose.

## Escalation trigger

At or below the floor, do not concede. Route to the **supervising lawyer** for
the matter (or the `.dochaus/profile.md` Escalation Rules contact) before
accepting an asymmetric waiver or an unbacked exclusive remedy.

**Rationale:** Remedy limitations decide what a breach actually costs; the
lost-profits qualifier and the failure-of-essential-purpose backstop are the
two details that change outcomes most.
