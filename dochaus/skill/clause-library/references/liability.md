# Liability

**Topic:** Limitation of liability — caps, carve-outs, super-caps.
**Contract type:** All commercial agreements.
**Jurisdiction:** US commercial market; verify against the matter's
jurisdiction pack. Positions are firm-configurable defaults overridable via
`.dochaus/profile.md`.

## Preferred

Total liability capped at the fees paid or payable in the twelve months
preceding the claim, applied mutually, with a short list of carve-outs that
are uncapped or subject to a higher super-cap: breach of confidentiality,
IP-infringement indemnity, gross negligence or willful misconduct, and (where
the deal involves personal data) data-protection breaches under a super-cap of
a small multiple of the general cap.

```example
Except for Excluded Claims, each party's total aggregate liability arising out
of or relating to this Agreement shall not exceed the fees paid or payable by
Customer under this Agreement in the twelve (12) months preceding the event
giving rise to the claim. "Excluded Claims" means a party's breach of Section
[Confidentiality], its indemnification obligations under Section [Indemnity],
and its gross negligence, fraud, or willful misconduct.
```

**Off-market variants to watch.** Provider-favorable: cap at one to three
months' fees or a low fixed dollar amount unrelated to exposure; customer
payment obligations carved out of the cap while every provider obligation sits
inside it; no super-cap for data or confidentiality breaches in a deal whose
whole point is handling sensitive data. Customer-favorable: provider liability
uncapped while the customer's is capped; carve-outs that swallow the cap ("any
breach," all indemnities, all confidentiality breaches uncapped); cap measured
over total committed value rather than a trailing period.

## Fallbacks (ordered)

1. **Offer when** the counterparty resists the standard carve-out set: mutual
   cap with carve-outs narrowed to gross negligence/willful misconduct and IP
   indemnity; confidentiality and data moved under a super-cap of 2-3x the
   general cap.
2. **Offer when** the counterparty rejects any uncapped exposure: mutual cap at
   total fees paid under the agreement, all carve-outs under a defined
   super-cap, none uncapped except fraud and willful misconduct.
3. **Offer when** the risk is genuinely asymmetric (and only then): asymmetric
   caps, documented in the rationale — never silently.

## Walk-away floor

An uncapped general liability position for one side only, or a cap whose
carve-outs swallow it entirely so the negotiated cap is illusory.

## Escalation trigger

At or below the floor, do not concede. Route to the **supervising lawyer** for
the matter (or the contact named in the firm's `.dochaus/profile.md` Escalation
Rules) for a written exposure decision before any sign-off.

**Rationale:** The cap is the deal's core risk-allocation term; carve-outs are
where caps die. Most negotiations resolve at the carve-out list and the
super-cap multiple, not the headline cap.
