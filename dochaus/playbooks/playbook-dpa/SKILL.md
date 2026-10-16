---
name: playbook-dpa
description: "Firm playbook: positions, fallbacks, and approved clause text for GDPR/UK GDPR data processing agreements — Article 28 processor obligations, international transfers, liability allocation. EU/UK jurisdiction scope."
---

This is an executable firm playbook for data processing agreements under the
GDPR and UK GDPR, reviewed from the controller (customer) side. The
contract-type and jurisdiction tags for every position below come from this
header (GDPR/UK GDPR data processing agreement; EU/UK scope).

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
market-balanced rather than aspirational, and bounded by the Article 28 GDPR
minimums (a statutory floor cannot be overridden downward). The matter's
`.dochaus/profile.md` (see the `firm-profile` skill) overrides any Preferred,
fallback, or above-statutory floor here; where it is silent, these defaults
apply. When a counterparty term sits at or below the Unacceptable floor, do not
concede: escalate to the firm's **privacy / data-protection lead** (or the
contact named in the profile's Escalation Rules). The escalation is addressed to
that role, not a named person, so the position outlasts any individual
assignment.

**Jurisdiction guard:** This playbook applies only to processing subject to the
GDPR or UK GDPR. It pairs with the EU pack and the UK packs (EW, SCT, NIR). Do not apply
it to a matter with no GDPR/UK GDPR nexus, and flag to the user any DPA whose
governing law or processing scope falls outside the EU/EEA and the UK. Where a
position below reflects market drafting practice rather than a statutory
requirement, the Rationale says so — do not present market backstops (such as
fixed-hour breach-notice windows) as legal obligations.

## Scope, roles, and documented instructions

**Preferred:** The processor processes personal data only on the controller's
documented instructions, with the Article 28 GDPR mechanics for legally
compelled processing and for instructions the processor believes are unlawful.

```approved
The Processor shall process Personal Data only on documented instructions from
the Controller, including with regard to transfers of Personal Data to a third
country or an international organisation, unless required to do so by
applicable law to which the Processor is subject; in such a case, the Processor
shall inform the Controller of that legal requirement before processing, unless
that law prohibits such information on important grounds of public interest.
The Agreement, this DPA, and the Controller's use of the service features
constitute the Controller's complete documented instructions at the date of
this DPA; additional instructions require the parties' written agreement. The
Processor shall immediately inform the Controller if, in its opinion, an
instruction infringes applicable data protection law.
```

**Unacceptable:** a right for the processor to process personal data for its
own purposes (including product improvement or model training on identifiable
data) outside the controller's instructions; instructions defined solely by the
processor's privacy policy as amended from time to time; omission of the
duty to inform the controller before legally compelled processing.

**Rationale:** Documented-instructions language tracks Article 28(3)(a) GDPR,
the core of the controller-processor relationship; own-purposes processing
rights make the vendor a controller in its own right, which changes the
parties' regulatory exposure and must be flagged, not redlined silently.

## Confidentiality of processing personnel

**Preferred:** Every person authorized to process the personal data is bound by
confidentiality, contractually or by statute.

```approved
The Processor shall ensure that persons authorised to process the Personal Data
have committed themselves to confidentiality or are under an appropriate
statutory obligation of confidentiality, and that access to Personal Data is
limited to those personnel who need such access to perform the Processor's
obligations under the Agreement and this DPA.
```

**Unacceptable:** confidentiality limited to employees while contractors and
agency staff are unaddressed; confidentiality obligations that expire on
termination of employment or of the DPA.

**Rationale:** The commitment tracks Article 28(3)(b) GDPR; the
need-to-know limitation is standard market hardening on top of the statutory
floor.

## Security measures

**Preferred:** Appropriate technical and organizational measures anchored to a
specific security annex, with a commitment not to materially diminish
protection during the term.

```approved
Taking into account the state of the art, the costs of implementation and the
nature, scope, context and purposes of processing, as well as the risks to the
rights and freedoms of natural persons, the Processor shall implement and
maintain appropriate technical and organisational measures to ensure a level of
security appropriate to the risk, including the specific measures described in
the security annex to this DPA. The Processor may update those measures from
time to time, provided the updates do not materially diminish the overall level
of protection for the Personal Data during the term of the Agreement.
```

**Unacceptable:** security obligations described only by reference to the
processor's website policy as amended at the processor's discretion, with no
non-diminishment commitment; measures qualified by "commercially reasonable"
in place of the risk-appropriateness standard; no security annex or equivalent
identifiable baseline at all.

**Rationale:** The standard tracks Article 32 GDPR; the annex anchor and the
non-diminishment covenant are the market mechanics that make it enforceable
(the Common Paper DPA likewise anchors security to a referenced security
policy rather than leaving it free-floating).

## Subprocessors

**Preferred:** General written authorization tied to a published subprocessor
list, advance written notice of changes with an objection right, full
flow-down of DPA obligations, and processor liability for subprocessor
performance.

```approved
The Controller grants the Processor general written authorisation to engage the
subprocessors identified in the Processor's subprocessor list as at the date of
this DPA. The Processor shall give the Controller prior written notice of the
addition or replacement of any subprocessor at least ten (10) business days in
advance, and the Controller may object on reasonable data-protection grounds
within that period. If the parties cannot resolve an objection within thirty
(30) days, the Controller may terminate the affected services without penalty
and receive a pro-rata refund of prepaid, unused fees. The Processor shall
impose on each subprocessor, by written contract, data protection obligations
no less protective than those set out in this DPA, shall limit each
subprocessor's access to Personal Data to what is required to perform the
obligations subcontracted to it, and shall remain fully liable to the
Controller for the performance of each subprocessor's obligations.
```

**Fallbacks:** when the processor's notice mechanism is a subscription-based
mailing list or feed rather than direct written notice, and the objection and
flow-down mechanics are otherwise intact.

```approved-fallback
The Controller grants the Processor general written authorisation to engage the
subprocessors identified in the Processor's subprocessor list as at the date of
this DPA. The Processor shall maintain a mechanism by which the Controller can
subscribe to notice of subprocessor changes and shall publish notice of the
addition or replacement of any subprocessor at least ten (10) business days in
advance. The Controller may object on reasonable data-protection grounds within
that period, and if the parties cannot resolve the objection within thirty (30)
days, the Controller may terminate the affected services without penalty. The
Processor shall impose on each subprocessor, by written contract, data
protection obligations no less protective than those set out in this DPA and
shall remain fully liable to the Controller for the performance of each
subprocessor's obligations.
```

**Unacceptable:** unrestricted subprocessor engagement with no notice or
objection mechanism; an objection right whose only outcome is that the
controller continues under protest; flow-down stated as "substantially
similar" obligations without the no-less-protective standard; disclaimer of
processor liability for subprocessor acts and omissions.

**Rationale:** Notice, objection, flow-down, and retained liability are the
Article 28(2) and 28(4) GDPR package; the ten-business-day notice period
matches the Common Paper DPA position, and the termination-on-unresolved-
objection mechanic is the market remedy that gives the objection right teeth.

## Assistance with data subject requests

**Preferred:** The processor assists the controller with data subject rights by
appropriate technical and organizational measures, routes any requests it
receives to the controller, and does not respond on its own.

```approved
Taking into account the nature of the processing, the Processor shall assist
the Controller by appropriate technical and organisational measures, insofar as
this is possible, in fulfilling the Controller's obligation to respond to
requests for exercising data subjects' rights under applicable data protection
law. If the Processor receives a request from a data subject relating to
Personal Data processed under this DPA, it shall promptly notify the Controller
and shall not respond to the request except on the Controller's documented
instructions or as required by applicable law.
```

**Unacceptable:** assistance conditioned on uncapped, undisclosed fees set at
the processor's discretion; a right for the processor to respond to data
subjects directly without the controller's instruction; assistance limited to
access requests while erasure, objection, and portability are unaddressed.

**Rationale:** Tracks Article 28(3)(e) GDPR. Reasonable cost-recovery for
extraordinary assistance is market-acceptable, but only the controller is
positioned to decide how a rights request is answered.

## Personal data breach notification

**Preferred:** Notification without undue delay after the processor becomes
aware of a personal data breach, with a 72-hour contractual backstop, required
content, and ongoing cooperation.

```approved
The Processor shall notify the Controller without undue delay, and in any event
within seventy-two (72) hours, after becoming aware of a Personal Data Breach
affecting Personal Data processed under this DPA. The notification shall, to
the extent then known, describe the nature of the breach, the categories and
approximate number of data subjects and Personal Data records concerned, the
likely consequences, and the measures taken or proposed to address the breach
and mitigate its effects, and the Processor shall supplement the notification
as further information becomes available. The Processor shall provide
reasonable assistance and cooperation with the Controller's obligations to
notify supervisory authorities and data subjects. The Processor's notification
of or response to a Personal Data Breach is not an acknowledgement of fault or
liability.
```

**Fallbacks:** when the processor will not commit to a fixed-hour backstop and
the data processed is low-risk.

```approved-fallback
The Processor shall notify the Controller without undue delay after becoming
aware of a Personal Data Breach affecting Personal Data processed under this
DPA, shall provide the Controller with timely information about the breach as
it becomes known, and shall provide reasonable assistance and cooperation with
the Controller's obligations to notify supervisory authorities and data
subjects.
```

**Unacceptable:** notification owed only for breaches the processor determines
in its sole discretion to be material or to pose a risk to data subjects (that
assessment belongs to the controller); a notice window of five days or longer;
notification obligations conditioned on a confidentiality or non-disparagement
commitment from the controller.

**Rationale:** The statutory processor duty is notification without undue delay
(Article 33(2) GDPR); the 72-hour backstop is contractual market practice — the
Common Paper DPA uses it — chosen because the controller's own
supervisory-authority notification window under Article 33(1) is 72 hours, and
the content list tracks Article 33(3).

## Deletion or return of personal data

**Preferred:** At the controller's choice, deletion or return of all personal
data at the end of the services, with retained copies confined to legal
requirements and kept under the DPA's protections.

```approved
At the Controller's choice, the Processor shall delete or return to the
Controller all Personal Data after the end of the provision of services
relating to processing, and shall delete existing copies unless storage of the
Personal Data is required by applicable law. Any Personal Data retained under
such a legal requirement, including copies in routine backup systems pending
scheduled deletion, shall remain subject to the obligations of this DPA for as
long as it is retained and shall be processed only as required by that law. On
the Controller's request, the Processor shall confirm deletion in writing.
```

**Unacceptable:** deletion at the processor's discretion rather than the
controller's choice between deletion and return; an open-ended retention right
untethered to legal requirements; retained copies released from the DPA's
obligations; deletion conditioned on payment disputes being resolved.

**Rationale:** Tracks Article 28(3)(g) GDPR; the backup carve-out is the
standard operational accommodation and is only safe if retained copies stay
under the DPA.

## Audit rights

**Preferred:** The processor makes available all information necessary to
demonstrate compliance and allows and contributes to audits, with market
scaffolding: third-party reports first, inspections on reasonable notice,
frequency limits that lift after a breach or on a supervisory authority's
request.

```approved
The Processor shall make available to the Controller all information necessary
to demonstrate compliance with the obligations set out in this DPA, and shall
allow for and contribute to audits, including inspections, conducted by the
Controller or an auditor mandated by the Controller. The Controller will first
review the Processor's then-current third-party audit reports and
certifications, and may conduct an inspection where those materials are not
sufficient to demonstrate compliance, on reasonable prior written notice,
during normal business hours, no more than once in any twelve-month period,
and subject to reasonable confidentiality and security requirements. These
frequency and notice limits do not apply to audits following a Personal Data
Breach or required by a supervisory authority.
```

**Fallbacks:** when the processor operates a multi-tenant environment and
restricts on-site inspection, but the information and reports regime is
robust.

```approved-fallback
The Processor shall make available to the Controller all information necessary
to demonstrate compliance with the obligations set out in this DPA, including
the Processor's then-current third-party audit reports and certifications and
written responses to the Controller's reasonable security and compliance
questionnaires. Where such materials are not sufficient to demonstrate
compliance, or following a Personal Data Breach or at the request of a
supervisory authority, the Processor shall allow for and contribute to an
audit by an independent third-party auditor mandated by the Controller and
bound by confidentiality, conducted on reasonable prior written notice and in
a manner that does not compromise the security of other customers' data.
```

**Unacceptable:** audit rights replaced entirely by certifications with no
inspection or independent-audit path under any circumstances; audit fees set
at the processor's unrestricted discretion with no disclosure of rates;
information access limited to a marketing security summary; frequency limits
that continue to apply after a personal data breach.

**Rationale:** The information-and-audit obligation tracks Article 28(3)(h)
GDPR; reports-first sequencing and frequency limits are market practice
(reflected in the Common Paper DPA's audit mechanics), acceptable only while
the breach and supervisory-authority exceptions are preserved.

## International transfers

**Preferred:** No transfers outside the EEA or the UK except under a lawful
transfer mechanism, with the EU Standard Contractual Clauses and the UK IDTA or
UK Addendum incorporated by reference where required, and a cooperation duty if
a mechanism is invalidated.

```approved
The Processor shall not transfer Personal Data outside the European Economic
Area or the United Kingdom, and shall not permit any subprocessor to do so,
except in accordance with a lawful transfer mechanism under applicable data
protection law. Where required for a transfer subject to the GDPR, the parties
shall enter into the Standard Contractual Clauses adopted by the European
Commission for transfers of personal data to third countries, selecting the
modules and options appropriate to the parties' roles, which are incorporated
into this DPA by reference. Where required for a transfer subject to the UK
GDPR, the parties shall enter into the UK International Data Transfer
Agreement or the UK International Data Transfer Addendum to the Standard
Contractual Clauses, as applicable, which is incorporated into this DPA by
reference. If a transfer mechanism relied on by the parties is invalidated or
superseded, the Processor shall cooperate with the Controller in good faith to
implement a lawful replacement mechanism without undue delay and, pending such
implementation, shall suspend the affected transfers at the Controller's
request.
```

**Unacceptable:** a right to transfer personal data to any jurisdiction "as
the Processor deems appropriate"; reliance on an unnamed "approved transfer
mechanism" with no commitment to execute the SCCs or the UK IDTA/Addendum
where required; language purporting to amend the substance of the Standard
Contractual Clauses (the SCCs by their own terms may not be modified except
where they expressly permit it); no remedy or cooperation duty if a relied-on
mechanism is invalidated.

**Rationale:** This section references the SCCs and the UK IDTA/Addendum by
name and incorporates them; it deliberately contains no restatement of their
clause text. The execute-and-incorporate structure matches the Common Paper
DPA's transfer mechanics, and the invalidation-cooperation duty is the market
response to transfer mechanisms being struck down or superseded.

## Liability allocation

**Preferred:** DPA liability sits inside the main agreement's limitations, with
a single combined cap, but nothing limits liability under the transfer
mechanisms where their terms do not permit it, or either party's liability to
data subjects under applicable data protection law.

```approved
Each party's liability arising out of or relating to this DPA, whether in
contract, tort, or otherwise, is subject to the exclusions and limitations of
liability set out in the Agreement, and the parties' aggregate liability under
the Agreement and this DPA together shall not exceed the applicable liability
cap in the Agreement. Nothing in this DPA or the Agreement limits either
party's liability under the Standard Contractual Clauses or the UK
International Data Transfer Agreement or Addendum where such a limitation is
not permitted by their terms, and nothing in this DPA or the Agreement affects
any liability of either party to data subjects under applicable data
protection law.
```

**Unacceptable:** a separate, lower cap for DPA claims sitting beneath the
main agreement's cap; processor liability for data protection breaches
excluded entirely; a clause purporting to cap or exclude liability owed to
data subjects or under the incorporated transfer mechanisms; an indemnity from
the controller covering the processor's own non-compliance.

**Rationale:** The single-combined-cap structure with carve-throughs for the
transfer mechanisms and data-subject liability mirrors the Common Paper DPA's
liability position; where the negotiated main agreement carries an enhanced
cap for data and security claims (see playbook-saas-msa), that enhanced cap is
the one that should govern DPA claims.

---

Attribution: positions and mechanics are adapted from the
[Common Paper DPA](https://github.com/CommonPaper/DPA) by Common Paper, licensed
under [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). doc.haus changes
the source into preferred terms, fallbacks, unacceptable terms and rationale,
with adapted clause wording and controller-side GDPR/UK GDPR review guidance.
This is a modified playbook, not the unchanged DPA. The source material retains
its CC BY 4.0 licence. No Common Paper endorsement is implied.
