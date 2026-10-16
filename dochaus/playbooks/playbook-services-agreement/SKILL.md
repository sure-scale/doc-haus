---
name: playbook-services-agreement
description: "Firm playbook: positions, fallbacks, and approved clause text for US commercial services and consulting agreements, taken from the customer side. Do not apply to agreements governed by non-US law."
---

This is an executable firm playbook for US commercial services and consulting
agreements, reviewed from the customer (services recipient) side. The
contract-type and jurisdiction tags for every position below come from this
header (commercial services/consulting agreement; US governing law).

Each `##` section is one clause type, stated on a fixed schema so the redline
tool can paste the firm text byte-exact, the deterministic lint can detect
deviation, and `document-reviewer` can band each clause:

- **Preferred** — the firm's default opening position. Text inside an
  ```approved fence is the firm-approved replacement text — copy it byte-exact
  into the redline tool's `replacement` arg.
- **Fallbacks** — acceptable rungs below Preferred, each introduced by the
  "when" condition that justifies offering it (the "offer when" trigger). Text
  inside an ```approved-fallback fence is the firm's fallback text, to be used
  only when its stated condition is met.
- **Unacceptable** — the walk-away floor: the position the firm does not concede
  below without a human decision.
- **Rationale** — why the position sits where it does.

These positions are firm-configurable **defaults**, calibrated to be
market-balanced rather than aspirational. The matter's `.dochaus/profile.md`
(see the `firm-profile` skill) overrides any Preferred, fallback, or floor here;
where it is silent, these defaults apply. When a counterparty term sits at or
below the Unacceptable floor, do not concede: escalate to the **supervising
lawyer** for the matter, or to the contact named in the profile's Escalation
Rules. The escalation is addressed to that role, not a named person, so the
position outlasts any individual assignment.

**Jurisdiction guard:** The positions below reflect US commercial market practice
and assume a US governing law. If the agreement's governing law is outside the
United States, do not apply this playbook; flag the mismatch instead of adapting
the clauses. Enforceability of particular terms varies by state — treat the
positions as market reference points, not statements of law.

## Scope of services and statements of work

**Preferred:** All services are performed under written statements of work that
incorporate the agreement, with a written change-order process; neither party is
obligated by, or pays for, work outside an executed SOW or change order.

```approved
Provider shall perform the services described in one or more statements of work
executed by both parties (each, an "SOW"). Each SOW is governed by and
incorporates this Agreement, and in the event of a conflict between an SOW and
this Agreement, this Agreement controls unless the SOW expressly identifies the
conflicting provision and states that it controls. Any change to the scope,
schedule, deliverables, or fees of an SOW is effective only when set out in a
written change order executed by both parties. Customer has no obligation to pay
for services performed, or expenses incurred, outside an executed SOW or change
order.
```

**Unacceptable:** services defined only by a vague recital with no SOW mechanism;
changes to scope or fees effective by email or by the provider's notice alone; an
SOW permitted to override the agreement's liability, indemnity, or IP terms
without expressly saying so.

**Rationale:** SOW discipline is the control surface for the whole agreement: it
bounds fees, defines acceptance, and prevents the provider from converting
informal requests into billable scope. Silent SOW-over-master precedence lets
negotiated protections be undone in later paperwork that legal never sees.

## Payment terms

**Preferred:** Invoices payable net thirty days, fees fixed for the SOW term,
pre-approved expenses only, and a good-faith dispute right that suspends only the
disputed portion.

```approved
Provider shall invoice Customer as set forth in the applicable SOW. Undisputed
amounts are due within thirty (30) days after Customer's receipt of a correct
invoice. Customer may withhold any amount it disputes in good faith, provided it
pays the undisputed portion and notifies Provider of the basis for the dispute,
and the parties shall work in good faith to resolve the dispute promptly. Fees
are fixed for the term of the applicable SOW. Customer shall reimburse only
reasonable, documented out-of-pocket expenses approved by Customer in advance in
writing. Except as stated in an SOW, each party bears its own taxes, and Customer
is not responsible for taxes based on Provider's income.
```

**Fallbacks:** when the provider has genuine working-capital constraints and the
engagement is short.

```approved-fallback
Provider shall invoice Customer as set forth in the applicable SOW. Undisputed
amounts are due within fifteen (15) days after Customer's receipt of a correct
invoice. Customer may withhold any amount it disputes in good faith, provided it
pays the undisputed portion and notifies Provider of the basis for the dispute.
Customer shall reimburse only reasonable, documented out-of-pocket expenses
approved by Customer in advance in writing.
```

**Unacceptable:** payment due on receipt or in fewer than fifteen days; automatic
late fees or interest accruing on disputed amounts; a provider right to suspend
all services for nonpayment of a disputed invoice; uncapped or unapproved expense
pass-through; rate increases during an SOW term.

**Rationale:** Net thirty with a carve-out for good-faith disputes is the US
commercial baseline. The dispute mechanics matter more than the day count: if
disputed amounts accrue interest or trigger suspension, the customer's only
leverage over defective work disappears.

## Intellectual property in deliverables

**Preferred:** Deliverables are owned by the customer on payment as work made for
hire (with a backup assignment); the provider retains its pre-existing materials
and general know-how but grants a broad license to whatever pre-existing
materials are embedded in the deliverables.

```approved
Subject to Provider's receipt of payment for the applicable deliverable, all
deliverables specified in an SOW are works made for hire for Customer to the
extent permitted by applicable law, and Provider hereby assigns to Customer all
right, title, and interest in and to the deliverables, including all intellectual
property rights therein, to the extent they are not works made for hire. Provider
retains all right, title, and interest in and to its pre-existing materials,
tools, methodologies, and generalized knowledge and skills ("Provider
Materials"). To the extent any Provider Materials are incorporated into or
necessary to use a deliverable, Provider grants Customer a perpetual,
irrevocable, worldwide, non-exclusive, royalty-free, sublicensable license to
use, reproduce, modify, and distribute such Provider Materials as part of the
deliverables. Provider shall not incorporate any third-party or open-source
materials into a deliverable without Customer's prior written approval.
```

**Fallbacks:** when the provider's business model depends on reusing deliverable
components and the customer does not need exclusivity.

```approved-fallback
Provider retains ownership of the deliverables and of its pre-existing materials,
tools, and methodologies. Upon Provider's receipt of payment for the applicable
deliverable, Provider grants Customer a perpetual, irrevocable, worldwide,
non-exclusive, royalty-free, fully paid-up license to use, reproduce, modify,
distribute, and create derivative works of the deliverables for Customer's
business purposes, including the right to sublicense to Customer's affiliates and
contractors acting on Customer's behalf. Provider shall not incorporate any
third-party or open-source materials into a deliverable without Customer's prior
written approval.
```

**Unacceptable:** provider ownership of deliverables with only a terminable or
revocable license back; any assignment to the provider of the customer's
pre-existing intellectual property or data; ownership conditioned on payment of
amounts disputed in good faith; silence on embedded third-party or open-source
components.

**Rationale:** The customer pays for the deliverables and should own them; the
provider legitimately keeps its toolkit. The embedded-materials license is the
clause that makes ownership real — owning a deliverable you cannot lawfully run
or modify is worthless. Work-made-for-hire status depends on the work's category
under US copyright law, which is why the express backup assignment is required
rather than optional.

## Warranties

**Preferred:** The provider warrants professional, workmanlike performance
conforming to the SOW, non-infringement, and compliance with applicable law, with
re-performance or refund as the first-line remedy for nonconforming services.

```approved
Provider represents and warrants that: (a) the services will be performed in a
professional and workmanlike manner consistent with generally accepted industry
standards by appropriately qualified personnel; (b) the services and deliverables
will conform in all material respects to the specifications in the applicable
SOW; (c) to Provider's knowledge, the deliverables do not infringe or
misappropriate any third party's intellectual property rights; and (d) Provider
will comply with all laws applicable to its performance of the services. For any
breach of the warranty in clause (a) or (b), Customer's initial remedy is for
Provider to re-perform the nonconforming services or correct the nonconforming
deliverable at no charge within thirty (30) days of notice; if Provider fails to
do so, Customer may terminate the applicable SOW and receive a refund of fees
paid for the nonconforming services or deliverable. This remedy is in addition
to, and not in lieu of, Customer's other rights and remedies under this
Agreement.
```

**Fallbacks:** when the provider insists that re-performance and refund be the
exclusive remedy for the performance warranty.

```approved-fallback
Provider represents and warrants that: (a) the services will be performed in a
professional and workmanlike manner consistent with generally accepted industry
standards by appropriately qualified personnel; (b) the services and deliverables
will conform in all material respects to the specifications in the applicable
SOW; and (c) Provider will comply with all laws applicable to its performance of
the services. For any breach of the warranty in clause (a) or (b), Customer's
exclusive remedy is for Provider to re-perform the nonconforming services or
correct the nonconforming deliverable at no charge within thirty (30) days of
notice and, if Provider fails to do so, termination of the applicable SOW and a
refund of fees paid for the nonconforming services or deliverable. This
exclusivity does not limit Customer's rights under the indemnification or
limitation of liability provisions of this Agreement.
```

**Unacceptable:** services provided entirely "as is"; a warranty period shorter
than thirty days from delivery or acceptance; a disclaimer that also disclaims
the express warranties given in the agreement; an exclusive remedy that swallows
the indemnities.

**Rationale:** Professional-services warranties are about conformance and
competence, not perfection. The re-performance ladder gives the provider a fair
cure path while preserving refund and termination as the backstop; the fallback
concedes exclusivity only for the performance warranty and expressly preserves
the indemnity, which is the concession providers actually need.

## Indemnification

**Preferred:** The provider indemnifies against third-party claims of IP
infringement and claims arising from bodily injury, property damage, or the
provider's gross negligence or willful misconduct, with a standard
control-of-defense procedure and IP mitigation options.

```approved
Provider shall defend, indemnify, and hold harmless Customer and its affiliates,
and their respective officers, directors, and employees, from and against any
third-party claim, and all resulting damages, costs, and reasonable attorneys'
fees, to the extent arising from: (a) an allegation that the services or
deliverables infringe or misappropriate a third party's intellectual property
rights; (b) bodily injury, death, or damage to tangible property caused by
Provider's acts or omissions; or (c) Provider's gross negligence or willful
misconduct. Customer shall give Provider prompt written notice of the claim,
reasonable cooperation at Provider's expense, and sole control of the defense and
settlement, provided that Provider shall not settle any claim in a manner that
imposes liability on or requires an admission by Customer without Customer's
prior written consent. If a deliverable is, or in Provider's reasonable opinion
is likely to become, the subject of an infringement claim, Provider shall, at its
own expense, procure for Customer the right to continue using it, or modify or
replace it so that it is non-infringing without material loss of functionality;
if neither is commercially practicable, Provider may require return of the
deliverable and shall refund the fees paid for it.
```

**Fallbacks:** when the provider requires standard exceptions to the IP
indemnity.

```approved-fallback
Provider's obligations under clause (a) do not apply to the extent the alleged
infringement arises from: (i) Customer's modification of a deliverable other than
by Provider or with Provider's written approval; (ii) Customer's combination of a
deliverable with materials not provided or approved by Provider, where the
deliverable alone would not infringe; or (iii) Customer's use of a deliverable in
breach of this Agreement, in each case where the claim would not have arisen but
for such modification, combination, or use.
```

**Unacceptable:** no IP indemnity at all; an IP indemnity capped at or below the
general liability cap with no separate treatment; a customer obligation to
indemnify the provider for the provider's own performance of the services; a
provider right to settle claims imposing obligations on the customer without
consent.

**Rationale:** The IP indemnity is the customer's only practical protection
against a third party enjoining use of work it has already paid for, which is why
it pairs with the procure-modify-replace-refund ladder. The fallback exceptions
are market-standard and properly limited by the but-for qualifier; broader
exceptions hollow out the indemnity.

## Limitation of liability

**Preferred:** A mutual exclusion of consequential damages and a mutual cap at
the greater of fees paid or payable under the applicable SOW in the twelve months
preceding the claim, with carve-outs for indemnification obligations,
confidentiality breach, IP infringement, gross negligence, and willful
misconduct.

```approved
Except for Excluded Claims, neither party is liable to the other for any
indirect, incidental, special, consequential, or punitive damages, or for lost
profits, revenue, or data, arising out of or relating to this Agreement, even if
advised of the possibility of such damages. Except for Excluded Claims, each
party's total aggregate liability arising out of or relating to this Agreement
and all SOWs shall not exceed the total fees paid or payable by Customer under
the applicable SOW in the twelve (12) months preceding the event giving rise to
the claim. "Excluded Claims" means: (a) a party's indemnification obligations
under this Agreement; (b) a party's breach of its confidentiality obligations;
(c) Provider's infringement or misappropriation of Customer's intellectual
property; (d) a party's gross negligence or willful misconduct; and (e)
Customer's obligation to pay fees due under this Agreement.
```

**Fallbacks:** when the provider will not accept uncapped indemnification and the
deal warrants a super-cap compromise.

```approved-fallback
Except for Excluded Claims, neither party is liable to the other for any
indirect, incidental, special, consequential, or punitive damages, or for lost
profits, revenue, or data, arising out of or relating to this Agreement, even if
advised of the possibility of such damages. Each party's total aggregate
liability arising out of or relating to this Agreement and all SOWs shall not
exceed the total fees paid or payable by Customer under the applicable SOW in the
twelve (12) months preceding the event giving rise to the claim, except that for
a party's indemnification obligations and breach of confidentiality obligations
the cap is three (3) times that amount. "Excluded Claims" means a party's gross
negligence or willful misconduct, Provider's infringement or misappropriation of
Customer's intellectual property, and Customer's obligation to pay fees due under
this Agreement, none of which are subject to any cap under this Agreement.
```

**Unacceptable:** a one-way limitation protecting only the provider; a cap below
fees paid in the preceding twelve months; consequential-damages exclusions or
caps that apply to the provider's indemnification or confidentiality obligations
with no super-cap; an exclusion of liability for gross negligence or willful
misconduct.

**Rationale:** Mutuality and the carve-out list are where these clauses are won
or lost — a twelve-month fee cap that also caps the IP indemnity converts the
indemnity into a refund. The three-times super-cap is the standard middle ground
when fully uncapped indemnities are refused. Whether particular exclusions are
enforceable varies by state, which is a further reason not to rely on aggressive
exclusions surviving review.

## Termination

**Preferred:** Customer may terminate for convenience on thirty days' notice;
either party may terminate for uncured material breach on thirty days' notice or
immediately on the other's insolvency; on termination the customer pays for
conforming work performed through the effective date and receives all work in
progress.

```approved
Customer may terminate this Agreement or any SOW for convenience upon thirty (30)
days' prior written notice to Provider. Either party may terminate this Agreement
or the affected SOW if the other party materially breaches it and fails to cure
the breach within thirty (30) days after written notice describing the breach, or
immediately upon written notice if the other party becomes insolvent or subject
to bankruptcy or receivership proceedings that are not dismissed within sixty
(60) days. Upon termination, Customer shall pay Provider for services performed
in conformance with this Agreement through the effective date of termination and
for non-cancellable expenses approved in advance, and Provider shall promptly
deliver to Customer all completed and in-progress deliverables and all Customer
materials and data in its possession. Provisions that by their nature should
survive termination — including those governing payment, intellectual property,
confidentiality, indemnification, and limitation of liability — survive.
```

**Fallbacks:** when the provider requires compensation for early convenience
termination of a fixed-fee SOW.

```approved-fallback
Customer may terminate this Agreement or any SOW for convenience upon thirty (30)
days' prior written notice to Provider. If Customer terminates a fixed-fee SOW
for convenience, Customer shall pay the portion of the fixed fee proportionate to
the services performed in conformance with this Agreement through the effective
date of termination, as reasonably documented by Provider, in lieu of any
termination charge. Either party may terminate this Agreement or the affected SOW
if the other party materially breaches it and fails to cure the breach within
thirty (30) days after written notice describing the breach. Upon termination,
Provider shall promptly deliver to Customer all completed and in-progress
deliverables and all Customer materials and data in its possession.
```

**Unacceptable:** no customer right to terminate for convenience; termination
charges or acceleration of unearned fees on convenience termination beyond work
actually performed; a cure period shorter than ten days for non-payment-related
breaches; provider retention of customer data or deliverables as leverage after
termination.

**Rationale:** Convenience termination with payment for work performed is the
standard customer exit from a services relationship that is not working; the
provider's protection is compensation for conforming work, not lock-in. Handback
of work in progress and customer data must be unconditional — a withholding right
turns every fee dispute into a hostage situation.

## Governing law

**Preferred:** Governing law of the State of Delaware without regard to
conflict-of-laws principles, exclusive venue in Delaware courts, and a waiver
that does not reach equitable relief.

```approved
This Agreement is governed by and construed in accordance with the laws of the
State of Delaware, without regard to its conflict-of-laws principles. The parties
consent to the exclusive jurisdiction and venue of the state and federal courts
located in Delaware for any dispute arising out of or relating to this Agreement,
except that either party may seek injunctive or other equitable relief in any
court of competent jurisdiction.
```

**Fallbacks:** when the counterparty rejects Delaware and proposes a major
commercial jurisdiction.

```approved-fallback
This Agreement is governed by and construed in accordance with the laws of the
State of New York, without regard to its conflict-of-laws principles. The parties
consent to the exclusive jurisdiction and venue of the state and federal courts
located in New York County, New York for any dispute arising out of or relating
to this Agreement, except that either party may seek injunctive or other
equitable relief in any court of competent jurisdiction.
```

**Unacceptable:** governing law outside the United States (escalate — this
playbook does not apply); the counterparty's home jurisdiction where it is not a
neutral major commercial forum; mandatory arbitration with no equitable-relief
exception; a jury-trial or class-waiver provision added without specific review.

**Rationale:** Delaware is the firm's standard neutral default and New York the
standard concession for US commercial agreements. A non-US governing law changes
the assumptions behind every other position in this playbook, so it must be
escalated rather than redlined.
