---
name: playbook-saas-msa
description: "Firm playbook: positions, fallbacks, and approved clause text for US-law SaaS and cloud master subscription agreements, with customer-side and vendor-side notes per position."
---

This is an executable firm playbook for SaaS and cloud master subscription
agreements governed by US law. The contract-type and jurisdiction tags for
every position below come from this header (SaaS/cloud master subscription
agreement; US governing law).

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

**Jurisdiction guard:** This playbook states US commercial-market positions. Do
not apply it to a matter whose governing law is outside the United States
without flagging the mismatch to the user first; check the matter's
jurisdiction pack before reviewing against it.

**Sides:** Positions below are the firm's balanced baseline. Each section
carries a position note describing how the customer side and the vendor side
typically diverge from that baseline, so the reviewer can calibrate to the
client's position on the deal.

## Availability and service levels

**Preferred:** A stated monthly uptime commitment of at least 99.9%, backed by
defined service credits and a chronic-failure termination right with a pro-rata
refund.

```approved
Provider will use commercially reasonable efforts to make the Cloud Service
available for Customer's use at least 99.9% of the time in each calendar month,
excluding scheduled maintenance announced reasonably in advance and
unavailability caused by factors outside Provider's reasonable control. If
Provider fails to meet this availability commitment, Provider will issue the
service credits set out in the applicable service level terms. If Provider
fails to meet the availability commitment in any three (3) consecutive calendar
months or in any four (4) calendar months within a rolling twelve-month period,
Customer may terminate the affected Order on written notice and receive a
pro-rata refund of any prepaid, unused fees for the terminated portion of the
Subscription Term.
```

**Fallbacks:** when the counterparty rejects a chronic-failure termination
right and the service is not business-critical for the client.

```approved-fallback
Provider will use commercially reasonable efforts to make the Cloud Service
available for Customer's use at least 99.9% of the time in each calendar month,
excluding scheduled maintenance announced reasonably in advance and
unavailability caused by factors outside Provider's reasonable control. If
Provider fails to meet this availability commitment, Provider will issue the
service credits set out in the applicable service level terms as Customer's
sole and exclusive remedy for the failure, provided that nothing in this
Section limits Customer's right to terminate for material breach under this
Agreement.
```

**Unacceptable:** no availability commitment of any kind; an uptime target
disclaimed as aspirational with credits expressly excluded; availability
measured only annually; service credits stated as the sole remedy with the
material-breach termination right also waived.

**Position note:** Customer side: push for the chronic-failure termination
right and credits that escalate with downtime. Vendor side: the
commercially-reasonable-efforts 99.9% formulation with credits as sole remedy
is the published market default (Bonterms Cloud Terms s. 7.2 uses 99.9% as the
fallback when no SLA is attached; Common Paper pairs its Cloud Service
Agreement with a standard SLA), so a vendor refusing any commitment is below
market.

**Rationale:** 99.9% monthly with credits is the documented market baseline; a
credits-only regime without any exit for sustained failure leaves the customer
paying for a service it cannot use.

## Customer data ownership and use

**Preferred:** Customer retains all rights in Customer Data; provider use is
limited to providing the service; any analytics use is restricted to aggregated,
de-identified data that cannot identify the customer or any individual.

```approved
As between the parties, Customer retains all right, title, and interest,
including all intellectual property rights, in and to Customer Data. Provider
will access and use Customer Data solely to provide and maintain the Cloud
Service and related support under this Agreement and as otherwise instructed by
Customer in writing. Provider may use aggregated and de-identified data derived
from Customer's use of the Cloud Service for its internal analytics and service
improvement only if such data cannot reasonably be used to identify Customer,
any Customer end user, or any individual, and Provider will not attempt to
re-identify any such data.
```

**Fallbacks:** when the counterparty requires usage telemetry for product
operations and limits it to technical usage data rather than Customer Data
content.

```approved-fallback
As between the parties, Customer retains all right, title, and interest,
including all intellectual property rights, in and to Customer Data. Provider
will access and use Customer Data solely to provide and maintain the Cloud
Service and related support under this Agreement. Provider may collect and use
technical usage data and telemetry about the operation and performance of the
Cloud Service, provided such data does not include the content of Customer Data
and is not used or disclosed in any form that identifies Customer or any
individual.
```

**Unacceptable:** any license to use Customer Data to train machine-learning
models or improve products in identifiable form without express consent; an
assignment of Customer Data or derivatives to the provider; provider usage
rights that survive termination; ownership of outputs or derived data resting
with the provider.

**Position note:** Customer side: hold the line on the de-identification
restriction and the no-re-identification covenant. Vendor side: aggregated and
de-identified analytics rights are market-accepted; what is not market is
training rights over identifiable customer content, which sophisticated
customers now routinely strike.

**Rationale:** "Provider will access and use Customer Data solely to provide
and maintain the Cloud Service" and "Customer retains all intellectual property
and other rights in Customer Data" is the market formulation (Bonterms Cloud
Terms ss. 5.1, 15.1); broad analytics or training licenses convert a
subscription into an uncompensated data deal.

## Security obligations

**Preferred:** A binding commitment to appropriate technical and organizational
measures, anchored to a specific security exhibit, with a non-diminishment
covenant and prompt incident notice.

```approved
Provider will implement and maintain appropriate technical and organizational
measures designed to protect the security, confidentiality, and integrity of
Customer Data, including measures designed to prevent unauthorized access, use,
alteration, or disclosure of Customer Data. Provider will maintain the specific
security measures identified in this Agreement or the applicable security
exhibit and will not materially diminish the overall protection provided by
those measures during a Subscription Term. Provider will notify Customer
without undue delay after becoming aware of any unauthorized access to or
acquisition or disclosure of Customer Data, and will provide Customer with
timely information about the incident as it becomes known and reasonable
cooperation in Customer's response.
```

**Unacceptable:** security obligations stated only as "industry standard
practices" with no anchor document and no non-diminishment commitment; a
unilateral right to materially reduce security measures during the term; no
incident-notification obligation; notification owed only for incidents the
provider determines in its sole discretion to be material.

**Position note:** Customer side: insist on the security exhibit anchor and the
non-diminishment covenant — an unanchored "appropriate measures" promise is
unverifiable. Vendor side: the fallback formulation "appropriate technical and
organizational measures designed to prevent unauthorized access, use,
alteration or disclosure" is itself the market floor (Bonterms Cloud Terms s.
5.2 applies it whenever no security exhibit is attached), so a vendor offering
less than that floor is below market.

**Rationale:** Security promises only constrain the vendor if they are anchored
to identifiable measures and cannot be quietly weakened mid-term; incident
notice is the precondition for every downstream breach obligation the customer
carries.

## Suspension rights

**Preferred:** Suspension limited to enumerated grounds (sustained non-payment
after notice, material breach of usage rules, risk of material harm), with
notice where practicable, narrow scope, and prompt restoration.

```approved
Provider may suspend Customer's access to the Cloud Service only if: (a)
Customer's account is at least thirty (30) days overdue and Customer fails to
pay the overdue amounts within ten (10) days after written notice; (b) Customer
is in material breach of the usage restrictions in this Agreement; or (c)
Customer's use of the Cloud Service poses a risk of material harm to the Cloud
Service or to others. Where practicable, Provider will give Customer prior
written notice and a reasonable opportunity to cure before suspending, will
limit any suspension to the affected portion of the Cloud Service, and will
promptly restore access once the grounds for suspension are resolved.
```

**Unacceptable:** suspension at the provider's sole discretion or for
convenience; suspension for any alleged breach without a materiality threshold;
no notice or cure mechanic of any kind; no obligation to restore access once
the cause is resolved; continued accrual of fees stated alongside an
unrestricted suspension right.

**Position note:** Customer side: the three enumerated grounds with notice,
narrow scope, and restoration are the negotiated package — resist "including
but not limited to" expansions. Vendor side: the three grounds (30-day
non-payment, usage-rule breach, risk of material harm) mirror the published
market standard (Bonterms Cloud Terms s. 13), so a vendor needs no broader
right.

**Rationale:** Suspension is a self-help remedy that bypasses the dispute
process; unbounded suspension rights let a vendor convert any commercial
disagreement into an outage.

## Limitation of liability: cap

**Preferred:** A mutual aggregate cap of twelve months of fees, with an
enhanced cap of three times that amount for breaches of confidentiality, data
protection, or security obligations.

```approved
Except as otherwise provided in this Section, each party's entire aggregate
liability arising out of or related to this Agreement will not exceed the
amounts paid or payable by Customer to Provider under this Agreement in the
twelve (12) months immediately preceding the first event giving rise to
liability. Each party's entire aggregate liability arising out of or related to
a breach of its confidentiality, data protection, or security obligations under
this Agreement will not exceed three (3) times the foregoing amount.
```

**Fallbacks:** when the counterparty refuses any enhanced cap, the deal value
is modest, and the data processed is low-sensitivity.

```approved-fallback
Each party's entire aggregate liability arising out of or related to this
Agreement will not exceed the amounts paid or payable by Customer to Provider
under this Agreement in the twelve (12) months immediately preceding the first
event giving rise to liability.
```

**Unacceptable:** a cap measured by fees paid in fewer than the trailing twelve
months or by a single month's fees; a cap that applies only to the provider's
liability while the customer's is uncapped; data-protection and security
liability swept into the general cap when the service processes sensitive or
regulated data.

**Position note:** Customer side: the enhanced data/security cap is the
priority ask; the 3x multiplier matches the published market structure
(Bonterms Cloud Terms s. 16.5 sets an Enhanced Cap of three times the General
Cap for security and DPA breaches; Common Paper's Cloud Service Agreement uses
the same general-cap/enhanced-cap architecture). Vendor side: the 12-month
general cap is firmly market and should be conceded without resistance; the
negotiation is the multiplier and scope of the enhanced cap.

**Rationale:** Twelve months of fees is the documented market general cap; the
enhanced cap exists because the realistic loss from a data incident bears no
relationship to subscription fees.

## Limitation of liability: exclusions and carve-outs

**Preferred:** A mutual waiver of indirect and consequential damages, with the
waiver and cap inapplicable to IP indemnification obligations, fraud or willful
misconduct, payment obligations, and infringement of the other party's IP.

```approved
Neither party will be liable for any indirect, special, incidental, reliance,
or consequential damages, or for loss of use, lost profits, or interruption of
business, however caused and under any theory of liability, even if advised of
the possibility of such damages. The exclusions and cap in this Section do not
apply to: (a) a party's obligations to indemnify the other party against
third-party claims of infringement or misappropriation under this Agreement;
(b) a party's fraud or willful misconduct; (c) Customer's obligation to pay
fees due under this Agreement; or (d) a party's infringement or
misappropriation of the other party's intellectual property rights.
```

**Unacceptable:** a one-way limitation protecting only the provider; the cap
applied to the provider's IP indemnification obligation; an exclusion of direct
damages; language characterizing loss or corruption of Customer Data as
per-se excluded consequential damages; a carve-out list that exposes only the
customer (for example, carving out customer confidentiality breaches while
capping the provider's).

**Position note:** Customer side: also seek a carve-out for gross negligence
where the client's risk profile warrants it — note to the reviewer that the
enforceability of limiting liability for gross negligence varies by state, so
frame this as a negotiation ask rather than a legal entitlement. Vendor side:
the mutual indirect-damages waiver mirrors the market formulation (Bonterms
Cloud Terms s. 16.2); symmetrical carve-outs are the price of keeping it.

**Rationale:** Carve-outs are where liability clauses are actually won or lost;
an IP indemnity that is capped at twelve months of fees is not an indemnity,
and asymmetric carve-out lists are the most common hidden one-sidedness in
vendor paper.

## Term, renewal, and termination

**Preferred:** Twelve-month subscription terms with automatic renewal on
30 days' non-renewal notice, mutual termination for uncured material breach,
and a post-termination data-export window before deletion.

```approved
Each Subscription Term will renew automatically for successive periods equal to
the initial Subscription Term unless either party gives written notice of
non-renewal at least thirty (30) days before the end of the then-current
Subscription Term. Either party may terminate this Agreement on written notice
if the other party fails to cure a material breach within thirty (30) days
after receiving written notice of the breach. Upon expiration or termination,
Provider will, at Customer's request made within thirty (30) days after the
effective date of expiration or termination, make Customer Data available to
Customer for export in a commonly used, machine-readable format at no
additional charge, and will thereafter delete Customer Data in its possession
or control, except to the extent retention is required by applicable law,
in which case the retained data remains subject to the confidentiality and
security obligations of this Agreement.
```

**Fallbacks:** when the counterparty's renewal operations require a 60-day
non-renewal notice window and the client accepts the longer runway.

```approved-fallback
Each Subscription Term will renew automatically for successive periods equal to
the initial Subscription Term unless either party gives written notice of
non-renewal at least sixty (60) days before the end of the then-current
Subscription Term. Provider will notify Customer of the upcoming renewal,
including any change to fees for the renewal period, at least thirty (30) days
before the non-renewal notice deadline. Either party may terminate this
Agreement on written notice if the other party fails to cure a material breach
within thirty (30) days after receiving written notice of the breach.
```

**Unacceptable:** a non-renewal notice window longer than sixty days; automatic
renewal combined with unilateral, uncapped price increases taking effect
without advance notice before the non-renewal deadline; provider termination
for convenience without a pro-rata refund of prepaid fees; no post-termination
data-export window before deletion.

**Position note:** Customer side: the export window and the
renewal-price-notice mechanic are the asks that matter operationally. Vendor
side: 12-month initial terms with 30-day non-renewal notice and 30-day cure for
material breach are the published market defaults (Bonterms Cloud Terms ss.
14.1, 14.3); auto-renewal itself is market and not worth contesting.

**Rationale:** Renewal mechanics decide whether the customer can ever leave on
its own terms; the data-export window is the practical exit right, and a
silent renewal paired with a missed notice deadline and a price increase is the
most common subscription trap.

---

Attribution: market-norm baselines are adapted from the
[Common Paper Cloud Service Agreement](https://commonpaper.com/standards/cloud-service-agreement/)
by Common Paper and the
[Bonterms Cloud Terms, Version 1.0](https://github.com/Bonterms/Cloud-Terms) by
Bonterms. Both sources are licensed under
[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). doc.haus changes these
baselines into preferred terms, fallbacks, unacceptable terms and rationale,
with adapted clause wording and customer-side and vendor-side review guidance.
This is a modified playbook, not an unchanged copy of either agreement. The
source material retains its CC BY 4.0 licence. No endorsement by either publisher
is implied. Bonterms section references above refer to Version 1.0.
