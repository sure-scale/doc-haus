---
name: playbook-engagement-letter
description: "Firm playbook: positions, fallbacks, and approved clause text for law-firm client engagement letters governed by the law of England and Wales. Do not apply to engagements governed by any other jurisdiction."
---

This is an executable firm playbook for law-firm engagement letters (letters of
engagement / client care letters) governed by the law of England and Wales. The
contract-type and jurisdiction tags for every position below come from this
header (solicitor engagement letter; England and Wales).

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
market-balanced rather than aspirational, and bounded by the SRA Standards and
Regulations (a regulatory floor cannot be overridden downward). The matter's
`.dochaus/profile.md` (see the `firm-profile` skill) overrides any Preferred,
fallback, or above-regulatory floor here; where it is silent, these defaults
apply. When a term sits at or below the Unacceptable floor, do not concede:
escalate to the **supervising partner** for the engagement (or the firm's COLP
where the floor is a regulatory one, or the contact named in the profile's
Escalation Rules). The escalation is addressed to that role, not a named person,
so the position outlasts any individual assignment.

**Jurisdiction guard:** This playbook assumes a solicitors' practice regulated in
England and Wales. Solicitor engagement terms — in particular limitation of
liability, complaints handling, and client-money provisions — are heavily
regulated and differ materially between jurisdictions. If the matter's governing
law is not England and Wales, or the firm is not regulated by the Solicitors
Regulation Authority, do not apply this playbook; flag the mismatch instead of
adapting the clauses.

## Scope of engagement

**Preferred:** The scope of work is defined positively and bounded by an express
exclusion: anything not listed is out of scope unless separately agreed in
writing. Commercial merits, tax (beyond anything expressly included), and matters
requiring other professional advice are excluded by default.

```approved
Our work for you is limited to the matters described in this letter. We will not
advise on the commercial merits of the transaction, on tax, on accounting or
financial matters, or on the laws of any jurisdiction other than England and
Wales, unless this letter expressly says so or we separately agree to do so in
writing. If you ask us to carry out additional work, we will confirm the revised
scope and any change to our fee estimate in writing before we begin it.
```

**Unacceptable:** an open-ended scope with no exclusion language; a scope that
implies advice on foreign law, tax, or commercial viability without expressly
including it; scope changes effective without written confirmation.

**Rationale:** Scope creep is the leading source of fee disputes and negligence
claims on engagements. A positive scope plus an express exclusion and a
written-variation gate keeps the retainer aligned with the fee estimate and the
firm's insurance position.

## Fees and estimates

**Preferred:** Fees are calculated by reference to time spent at stated hourly
rates, with a written estimate that is expressly not a fixed quotation and a
commitment to warn the client and agree a revision before the estimate is
materially exceeded.

```approved
Our fees are calculated principally by reference to the time spent on your matter
at the hourly rates set out in this letter. Any estimate we give is our best
assessment of the likely fees based on the information available at the time; it
is not a fixed quotation or a cap. If it becomes likely that our fees will exceed
the estimate, we will tell you before further significant costs are incurred and
agree a revised estimate with you. We will tell you in writing if our hourly
rates change, and any change will apply only to work done after we have told you.
```

**Fallbacks:** when the client requires a fixed fee for a defined stage of work.

```approved-fallback
Our fee for the work described in this letter is fixed at the amount stated in
this letter, exclusive of VAT and disbursements, provided the matter proceeds
substantially as described and on the assumptions set out in this letter. If the
matter changes materially, or an assumption proves incorrect, we will tell you
promptly and agree in writing either a revised fixed fee or a move to our hourly
rates for the additional work before it is carried out.
```

**Unacceptable:** an estimate presented as binding with no revision mechanism; a
right to raise hourly rates retrospectively or without notice; fee terms that are
silent on what happens when the estimate is exceeded.

**Rationale:** The estimate-not-quotation framing with an advance-warning duty
reflects standard client-care practice and protects both sides: the client is
never surprised by fees and the firm is never locked to a stale estimate.

## Billing and payment

**Preferred:** Monthly interim bills with a final bill on completion, payable
within fourteen days, with interest on late payment at a stated margin over the
Bank of England base rate.

```approved
We will deliver interim bills monthly as the matter progresses, with a final bill
on completion of the matter. Each bill is payable within fourteen (14) days of
its date. We may charge interest on any bill not paid within that period at four
per cent (4%) per year above the base rate of the Bank of England from time to
time, calculated from the date of the bill until payment. If a bill remains
unpaid, we may suspend work on your matter after giving you written notice.
```

**Fallbacks:** when the client's accounts-payable process cannot meet a
fourteen-day term.

```approved-fallback
We will deliver interim bills monthly as the matter progresses, with a final bill
on completion of the matter. Each bill is payable within thirty (30) days of its
date. We may charge interest on any bill not paid within that period at four per
cent (4%) per year above the base rate of the Bank of England from time to time,
calculated from the date of the bill until payment.
```

**Unacceptable:** payment terms longer than sixty days; no right to interim
billing on a matter expected to run more than three months; a client right to
withhold the whole of a bill pending resolution of a dispute over part of it.

**Rationale:** Regular interim billing keeps fee exposure visible and disputes
small. The interest provision is a firm commercial position, not a statutory
rate; the suspension right must be exercised on notice to remain consistent with
the firm's professional obligations to the client.

## Money on account

**Preferred:** The firm may require payments on account of fees and
disbursements, held in client account and applied against bills, with a right to
ask for top-ups as the matter proceeds.

```approved
We may ask you to make a payment on account of our fees and disbursements before
we begin substantive work and at intervals as the matter proceeds. We will hold
money paid on account in our client account in accordance with the rules of our
regulator and apply it against our bills as they are delivered. If you do not
provide money on account when reasonably requested, we may suspend work on your
matter after giving you written notice.
```

**Unacceptable:** a provision treating money on account as earned on receipt; any
suggestion that client money will be held other than in accordance with the
applicable regulatory rules on client money.

**Rationale:** Money on account is held as client money under the regulator's
rules until billed; the clause should say so expressly. Treating advances as
earned on receipt is inconsistent with how regulated firms in England and Wales
must hold client money.

## Client responsibilities

**Preferred:** The client commits to clear and timely instructions, provision of
requested documents and information, funding as agreed, and prompt notice of
changes, with an express link between client delay and the estimate and
timetable.

```approved
You agree to give us clear and timely instructions, to provide the documents and
information we reasonably request, to provide money on account and pay our bills
as agreed, and to tell us promptly of any change to the transaction, your
circumstances, or your instructions. Delay in any of these may affect our fee
estimate and the timetable for your matter, and we will not be responsible for
the consequences of acting on incomplete or inaccurate information you have
provided.
```

**Unacceptable:** an engagement letter with no client-responsibility provision at
all; responsibility language that purports to make the client liable for the
firm's own errors.

**Rationale:** A clear statement of client responsibilities supports the fee
estimate, the timetable, and the firm's position if the matter goes wrong because
of late or inaccurate instructions. It must not be drafted so widely that it
shifts the firm's professional responsibility onto the client.

## Conflicts of interest and confidentiality

**Preferred:** The firm confirms its conflict check, commits to notify the client
if a conflict arises, and states a continuing duty of confidentiality subject
only to disclosure required by law or the firm's regulator or authorised by the
client.

```approved
We have carried out a check for conflicts of interest and are not aware of any
conflict that prevents us acting for you. If a conflict of interest arises during
the engagement, we will tell you promptly and explain the options available,
which may include our ceasing to act on the matter. We will keep your affairs and
the information you give us confidential, except where disclosure is required by
law or by our regulator, or where you authorise it. Our duty of confidentiality
continues after this engagement ends.
```

**Unacceptable:** a blanket advance waiver of all future conflicts; a
confidentiality carve-out allowing disclosure for the firm's marketing or
commercial purposes without consent; confidentiality expiring when the engagement
ends.

**Rationale:** Conflict handling and confidentiality are professional-conduct
obligations for solicitors in England and Wales, not purely contractual terms.
The clause should track those obligations procedurally; it cannot contract out of
them, so broad waivers and carve-outs are both unacceptable and ineffective.

## Limitation of liability

**Preferred:** A single aggregate monetary cap corresponding to the firm's
professional indemnity insurance cover, an exclusion of indirect and
consequential loss, and an express carve-out preserving all liability that cannot
lawfully be limited.

```approved
Our total aggregate liability to you arising out of or in connection with this
engagement, whether in contract, tort (including negligence), breach of statutory
duty or otherwise, is limited to three million pounds (GBP 3,000,000). We are not
liable for any indirect or consequential loss, or for loss of profit, revenue, or
anticipated savings. Nothing in this letter excludes or limits any liability that
cannot lawfully be excluded or limited, including liability for death or personal
injury caused by negligence, or for fraud or fraudulent misrepresentation.
```

**Fallbacks:** when the client is a substantial commercial party and negotiates a
higher cap that the firm's insurance position supports.

```approved-fallback
Our total aggregate liability to you arising out of or in connection with this
engagement, whether in contract, tort (including negligence), breach of statutory
duty or otherwise, is limited to ten million pounds (GBP 10,000,000). We are not
liable for any indirect or consequential loss, or for loss of profit, revenue, or
anticipated savings. Nothing in this letter excludes or limits any liability that
cannot lawfully be excluded or limited, including liability for death or personal
injury caused by negligence, or for fraud or fraudulent misrepresentation.
```

**Unacceptable:** any cap below the minimum level of professional indemnity cover
the firm is required to maintain; a cap with no carve-out for liability that
cannot lawfully be limited; an attempt to exclude liability for fraud or for
death or personal injury caused by negligence; an exclusion of all liability for
negligence.

**Rationale:** Limitation of solicitor liability in England and Wales is heavily
constrained by professional regulation and general law, and the constraints are
jurisdiction-specific. The GBP 3,000,000 figure is a firm default chosen to
correspond to a typical professional indemnity insurance position — before using
either fence, confirm the figure against the firm's actual insurance cover and
the minimum terms its regulator requires; a cap set below required cover, or one
purporting to limit non-excludable liability, is likely to be ineffective and may
raise regulatory concerns. Do not transplant this clause to engagements outside
England and Wales.

## Termination of engagement

**Preferred:** The client may end the engagement at any time on written notice;
the firm may cease to act only for good reason and on reasonable written notice;
accrued fees and disbursements survive termination.

```approved
You may end this engagement at any time by giving us written notice. We may cease
to act for you only for good reason — for example, a conflict of interest,
non-payment of our bills or failure to provide money on account, a breakdown in
the relationship of trust and confidence, or your failure to give us instructions
— and only after giving you reasonable written notice. If the engagement ends for
any reason, you remain liable for our fees and disbursements incurred up to the
date it ends, and we will cooperate in the orderly transfer of your matter and
papers, subject to any rights we may have in respect of unpaid bills.
```

**Unacceptable:** a firm right to terminate at will without reason or notice; any
restriction on the client's right to end the engagement; termination provisions
silent on accrued fees or on the handover of the client's matter.

**Rationale:** Asymmetric termination — client free to leave, firm constrained to
good reason and notice — reflects the solicitor-client relationship in England
and Wales. A firm-side at-will termination right is inconsistent with the
professional duty not to leave a client in the lurch.

## Complaints and regulation

**Preferred:** A named internal complaints route, confirmation of the firm's
regulator, and a signpost to the Legal Ombudsman and to the client's rights in
respect of bills, framed procedurally and without hard-coding time limits that
the ombudsman publishes and may change.

```approved
We aim to provide a high standard of service. If you are unhappy with our service
or with a bill, please raise it first with the partner responsible for your
matter, and we will respond under our written complaints procedure, a copy of
which is available on request. If we cannot resolve your complaint, you may be
able to refer it to the Legal Ombudsman, subject to the eligibility rules and
time limits the Legal Ombudsman publishes. You may also have rights to challenge
or complain about a bill, including rights under the applicable legislation
governing solicitors' charges. This firm is authorised and regulated by the
Solicitors Regulation Authority.
```

**Unacceptable:** omission of the complaints provision entirely; language
purporting to exclude or restrict the client's right to complain to the Legal
Ombudsman or to challenge a bill; misstating the firm's regulator.

**Rationale:** Complaints signposting is a client-care requirement for
SRA-regulated firms. Time limits and eligibility for the Legal Ombudsman are set
by the ombudsman's published scheme rules and change from time to time, so the
clause points to them rather than restating them.

## Governing law

**Preferred:** The engagement and the firm's terms of business are governed by
the law of England and Wales, with the courts of England and Wales having
exclusive jurisdiction.

```approved
This engagement and our terms of business are governed by the law of England and
Wales, and the courts of England and Wales have exclusive jurisdiction over any
dispute arising out of or in connection with them.
```

**Unacceptable:** any governing law other than England and Wales (this playbook's
positions assume it — a different governing law requires a different playbook,
not a redline); arbitration or foreign-court jurisdiction clauses purporting to
oust the client's regulatory complaint and bill-challenge routes.

**Rationale:** Every other position in this playbook — liability limitation,
client money, complaints, termination — is written against the law and
professional regulation of England and Wales. Changing the governing law silently
invalidates the rest of the playbook, so a mismatch must be escalated rather than
redlined.
