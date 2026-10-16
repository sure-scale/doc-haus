# Payment

**Topic:** Payment — escalation, set-off, refunds, taxes.
**Contract type:** All commercial agreements with recurring or staged fees.
**Jurisdiction:** US commercial market; verify against the matter's
jurisdiction pack. Positions are firm-configurable defaults overridable via
`.dochaus/profile.md`.

## Preferred

Fees fixed for the initial term and invoiced on a stated schedule, payable
net thirty days. Price increases apply only at renewal, on advance written
notice, and are bounded (a stated percentage or an inflation index); a
material increase gives the customer a right not to renew on normal terms.
Late amounts accrue interest at a stated, modest rate or the maximum lawful
rate if lower. Undisputed amounts are paid on time; good-faith disputed
amounts may be withheld pending resolution without triggering suspension.
Each party bears its own taxes on its income; transaction taxes are stated
separately.

```example
Provider may increase fees effective only upon a renewal term, by no more
than [X]% over the fees for the immediately preceding term, and only with at
least sixty (60) days' prior written notice before the non-renewal deadline.
Customer may withhold amounts disputed in good faith, provided it timely pays
all undisputed amounts and the parties work in good faith to resolve the
dispute promptly.
```

**Off-market variants to watch.** Provider-favorable: unilateral mid-term
price escalation with no cap and no exit right; all fees non-refundable in all
circumstances including the provider's own breach; suspension of service for
any unpaid amount including good-faith-disputed ones; one-way set-off ban.
Customer-favorable: net-90+ terms with no late interest; unlimited withholding
for any "disputed" amount with no good-faith or undisputed-portion qualifier;
all taxes including the provider's income taxes shifted via a one-way gross-up.

## Fallbacks (ordered)

1. **Offer when** the counterparty needs inflation protection: accept an
   index-linked annual increase (e.g., CPI-based) with a hard ceiling, in
   exchange for a termination right if exceeded.
2. **Offer when** the counterparty resists full refundability: accept
   non-refundability for convenience termination only — refunds always
   preserved where the customer exits for the provider's breach.
3. **Offer when** the provider needs a non-payment remedy: accept suspension
   for non-payment limited to undisputed amounts, after notice and a cure
   window.

## Walk-away floor

Uncapped mid-term escalation, or forfeiture of prepaids on the provider's own
breach.

## Escalation trigger

At or below the floor, do not concede. Route to the **supervising lawyer** for
the matter (or the `.dochaus/profile.md` Escalation Rules contact) before
accepting an open-ended price-change or no-refund-on-breach term.

**Rationale:** Payment terms convert directly into deal economics; the
escalation mechanism and the dispute/suspension interaction are the two
levers that most often surprise clients after signature.
