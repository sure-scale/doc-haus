---
name: playbook-nda
description: "Firm playbook: positions, fallbacks, and approved clause text for US-law mutual NDAs. Apply only to agreements governed by the law of a US state."
version: "1.1"
last-reviewed: "2026-06-11"
---

This is an executable firm playbook for mutual non-disclosure agreements
governed by US law. The contract-type and jurisdiction tags for every position
below come from this header (mutual NDA; US-state governing law).

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

**Jurisdiction guard:** Every position and fence in this playbook states a
US-law position drafted for agreements governed by the law of a US state.
Before applying any position, locate the agreement's governing-law clause. If
the agreement is governed by the law of a non-US jurisdiction, or governing law
is absent or cannot be determined, stop: do not present these positions or
fences as firm-approved redlines. Warn the user that the bound playbook covers
US-law mutual NDAs only and ask how to proceed.

**oneNDA recognition:** oneNDA (https://www.onenda.org/) is a widely adopted
standard-form mutual NDA whose body is fixed and which is negotiated only
through its variable fields (parties, purpose, confidentiality period, term,
governing law and jurisdiction). If the counterparty's paper is an unmodified
oneNDA — identifiable by the oneNDA name and branding on the document and the
fixed body paired with that short variables section — accept the body as-is and
negotiate only the variable fields. Do not redline the oneNDA body text, and
never copy oneNDA text into this playbook or any firm document (it is licensed
CC BY-ND, which prohibits derivatives). If the body has been modified from the
standard form, treat the document as ordinary counterparty paper and apply this
playbook normally, subject to the jurisdiction guard above.

Several sections below adapt language from the Bonterms Playbook for the
Bonterms Mutual NDA (https://bonterms.com/resources/nda-playbook), which
Bonterms publishes as public domain ("free to use, modify, or distribute," with
the condition that use may not imply Bonterms' endorsement); defined terms have
been conformed to this playbook's "Disclosing Party"/"Receiving Party"
vocabulary.

## Definition of confidential information

**Preferred:** Confidential information is defined broadly to cover all
non-public business, technical, and financial information disclosed in any form,
whether or not marked as confidential.

```approved
"Confidential Information" means all non-public information disclosed by or on
behalf of a party (the "Disclosing Party") to the other party (the "Receiving
Party"), in any form and whether or not marked or identified as confidential,
including business, technical, financial, commercial, product, and customer
information, together with the existence and terms of this Agreement and the
fact that discussions between the parties are taking place. Information disclosed
orally or visually need not be summarized in writing to qualify as Confidential
Information.
```

**Fallbacks:** when the counterparty requires marking as a condition of
protection for written disclosures.

```approved-fallback
"Confidential Information" means all non-public information disclosed by or on
behalf of the Disclosing Party to the Receiving Party that is either marked or
identified as confidential at the time of disclosure, or that a reasonable
person would understand to be confidential given the nature of the information
and the circumstances of disclosure. The terms of this Agreement and the
existence of discussions between the parties are Confidential Information
regardless of marking.
```

**Unacceptable:** definition limited to information marked "Confidential" in
writing with no reasonable-person catch-all; exclusion of the terms and
existence of the Agreement from protection.

**Rationale:** A marking-only definition leaves oral disclosures and unmarked
materials unprotected, which is the most common leakage point in diligence.

## Confidentiality term

**Preferred:** Confidentiality obligations survive for three years from the date
of disclosure, with trade secrets protected for as long as they remain trade
secrets under applicable law.

```approved
The Receiving Party's obligations under this Agreement with respect to each item
of Confidential Information shall survive for a period of three (3) years from
the date that item was disclosed. Notwithstanding the foregoing, Confidential
Information that constitutes a trade secret under applicable law shall remain
subject to the obligations of this Agreement for as long as such information
remains a trade secret.
```

**Fallbacks:** when the counterparty insists on a single fixed term and rejects
an open-ended trade-secret carve-out.

```approved-fallback
The Receiving Party's obligations under this Agreement shall survive for a
period of five (5) years from the date of disclosure of each item of
Confidential Information, after which such obligations shall expire.
```

**Unacceptable:** a term shorter than two years; an unqualified perpetual
confidentiality obligation covering all information rather than only trade
secrets.

**Rationale:** Three years matches the clause-library balanced position;
collapsing trade secrets into a fixed term forfeits indefinite statutory
protection, while perpetual blanket terms are administratively unworkable.

## Exclusions

**Preferred:** Standard exclusions for information that is public, independently
developed, already known, or lawfully received from a third party without a duty
of confidentiality.

```approved
Confidential Information does not include information that the Receiving Party
can demonstrate by competent evidence: (a) is or becomes generally available to
the public other than through a breach of this Agreement by the Receiving Party;
(b) was rightfully known to the Receiving Party without restriction prior to its
disclosure by the Disclosing Party; (c) is independently developed by the
Receiving Party without use of or reference to the Confidential Information; or
(d) is rightfully received by the Receiving Party from a third party without a
duty of confidentiality and without breach of this Agreement.
```

**Unacceptable:** an exclusion for information the Receiving Party "believes" to
be public; omission of the "without use of or reference to" qualifier on
independent development; an exclusion for information disclosed to the Receiving
Party's affiliates.

**Rationale:** These four exclusions are the universally accepted set; broadening
them (especially weak independent-development language) lets a receiving party
launder protected information out of scope.

## Return or destruction

**Preferred:** On request or termination, the Receiving Party returns or destroys
Confidential Information and certifies destruction, with a limited carve-out for
archival and legal-hold copies that remain subject to confidentiality.

```approved
Upon the Disclosing Party's written request or upon termination of this
Agreement, the Receiving Party shall promptly return or destroy all Confidential
Information in its possession or control and, upon request, certify such
destruction in writing. The Receiving Party may retain copies of Confidential
Information to the extent required by applicable law, regulation, or bona fide
internal record-retention policy, and copies created by routine automated backup
systems, provided that all such retained copies remain subject to the
confidentiality obligations of this Agreement for so long as they are retained.
```

**Unacceptable:** no destruction-certification obligation; an unlimited retention
right untethered from law or routine backup; retained copies released from
confidentiality obligations.

**Rationale:** The retention carve-out should be narrow and explicitly keep
retained copies under obligation; an open-ended retention right defeats the
return remedy.

## Permitted disclosures

**Preferred:** Disclosure compelled by law or legal process is permitted, subject
to prompt notice to the Disclosing Party (where lawful) and cooperation to seek
protective treatment.

```approved
If the Receiving Party is required by law, regulation, or valid legal process to
disclose any Confidential Information, it may do so provided that, to the extent
legally permitted, it gives the Disclosing Party prompt prior written notice so
that the Disclosing Party may seek a protective order or other appropriate
remedy, and discloses only that portion of the Confidential Information that it
is legally required to disclose. The Receiving Party shall reasonably cooperate,
at the Disclosing Party's expense, with the Disclosing Party's efforts to obtain
confidential treatment for the disclosed information.
```

**Fallbacks:** when the counterparty cannot commit to pre-disclosure notice for
regulatory or supervisory requests.

```approved-fallback
If the Receiving Party is required by law, regulation, or valid legal process to
disclose any Confidential Information, it may do so, and shall, to the extent
legally permitted and reasonably practicable, notify the Disclosing Party of the
required disclosure. No notice is required where the disclosure is made to a bank
regulatory or supervisory authority in the course of its routine examination or
oversight and is not targeted at the Disclosing Party.
```

**Unacceptable:** permitted disclosure with no notice obligation of any kind; a
right to disclose the entire body of Confidential Information rather than only the
portion legally required.

**Rationale:** Notice and minimization let the Disclosing Party defend its
information; a blanket compelled-disclosure right with no notice removes any
opportunity to seek protective treatment.

## Injunctive relief

**Preferred:** The parties acknowledge that breach may cause irreparable harm and
that the non-breaching party may seek injunctive relief without posting a bond, in
addition to other remedies.

```approved
The Receiving Party acknowledges that any breach or threatened breach of this
Agreement may cause the Disclosing Party irreparable harm for which monetary
damages would be an inadequate remedy. Accordingly, in addition to any other
remedies available at law or in equity, the Disclosing Party shall be entitled to
seek injunctive or other equitable relief to prevent or restrain any such breach
or threatened breach, without the necessity of posting a bond or proving actual
damages.
```

**Fallbacks:** when the counterparty's jurisdiction or policy requires that bond
be left to the court's discretion.

```approved-fallback
The Receiving Party acknowledges that any breach or threatened breach of this
Agreement may cause the Disclosing Party irreparable harm for which monetary
damages would be an inadequate remedy, and that the Disclosing Party shall be
entitled to seek injunctive or other equitable relief in addition to any other
remedies available at law or in equity, subject to the requirements of applicable
law regarding the posting of any bond.
```

**Unacceptable:** clause that requires the Disclosing Party to prove actual
damages before obtaining equitable relief; mutual waiver of the right to seek
injunctive relief.

**Rationale:** Injunctive relief is the only meaningful remedy for confidentiality
breach, since damages are hard to quantify; the no-bond formulation is the firm's
preference but is genuinely negotiable.

## Governing law

**Preferred:** Delaware law with exclusive venue in Delaware courts — the
firm's standard neutral jurisdiction; injunctive relief preserved for
confidentiality breaches.

```approved
This Agreement shall be governed by and construed in accordance with the laws of
the State of Delaware, without regard to its conflict-of-laws principles. The
parties consent to the exclusive jurisdiction and venue of the state and federal
courts located in New Castle County, Delaware for any dispute arising out of or
relating to this Agreement, except that either party may seek injunctive relief
in any court of competent jurisdiction.
```

**Fallbacks:** when the counterparty rejects Delaware and a different neutral US
forum is acceptable.

```approved-fallback
This Agreement shall be governed by and construed in accordance with the laws of
the State of New York, without regard to its conflict-of-laws principles. The
parties consent to the exclusive jurisdiction and venue of the state and federal
courts located in New York County, New York for any dispute arising out of or
relating to this Agreement, except that either party may seek injunctive relief
in any court of competent jurisdiction.
```

**Unacceptable:** governing law of the counterparty's home jurisdiction where it
is not neutral; governing law of a non-US jurisdiction (out of scope for this
playbook — stop and warn per the jurisdiction guard); a carve-out stripping the
right to seek injunctive relief in any competent court; mandatory arbitration
with no equitable-relief exception.

**Rationale:** A neutral, named jurisdiction with preserved injunctive relief
matches the clause-library disputes position; Delaware is the firm's standard
neutral choice and New York the accepted neutral alternative.

## Non-solicitation

**Preferred:** A mutual, twelve-month no-solicit limited to employees who became
known through the engagement, with the usual general-advertising and post-
employment carve-outs.

```approved
During the term of this Agreement and for a period of twelve (12) months
thereafter, neither party shall directly solicit for employment any employee of
the other party with whom it had contact or who became known to it in connection
with the discussions contemplated by this Agreement; provided that this
restriction shall not prohibit (a) general solicitations of employment not
specifically directed at such employees, including through job postings or
recruiting agencies, or (b) the hiring of any employee who responds to such a
general solicitation or who approaches a party on his or her own initiative.
```

**Unacceptable:** a non-solicit longer than twenty-four months; a one-sided
restriction binding only one party; a no-hire (as opposed to no-solicit)
provision; omission of the general-advertising carve-out.

**Rationale:** Non-solicitation is not a confidentiality term and should be
narrow, mutual, and time-limited; no-hire provisions and missing
general-advertising carve-outs are routinely unenforceable and overbroad.

## Affiliates

**Preferred:** Either party may extend the agreement to its affiliates on
notice, remaining responsible for affiliate compliance, with a 50%-control
definition of affiliate.

```approved
Upon notice to the other party, a party may allow its Affiliate to act as a
Disclosing Party or Receiving Party under this Agreement, provided that such
party remains responsible for compliance by its Affiliate with the terms of this
Agreement. "Affiliate" means an entity that controls, is controlled by, or is
under common control with a party, where control means at least 50% ownership
or power to direct an entity's management.
```

**Unacceptable:** affiliate access to Confidential Information with no party
responsible for the affiliate's compliance; an affiliate definition with no
control threshold; one-sided affiliate rights.

**Rationale:** Affiliates routinely need access in diligence, but the
contracting party must remain on the hook for them. Adapted from the Bonterms
NDA playbook items (public domain).

## Security safeguards

**Preferred:** The Receiving Party maintains industry-standard administrative,
physical, and technical safeguards, in addition to the general duty of care.

```approved
Without limiting the foregoing, the Receiving Party will implement and maintain
industry-standard administrative, physical, and technical safeguards designed to
prevent unauthorized access, use, alteration, or disclosure of Confidential
Information.
```

**Unacceptable:** a standard of care lower than reasonable care; care limited to
"the same care the Receiving Party uses for its own information" with no
reasonable-care floor.

**Rationale:** A concrete safeguards obligation gives the general duty of care
operational content and a clear breach standard. Adapted from the Bonterms NDA
playbook items (public domain).

## Breach notification

**Preferred:** The Receiving Party notifies the Disclosing Party of any
unauthorized use, disclosure, or loss within a fixed deadline and cooperates in
containment.

```approved
The Receiving Party will notify the Disclosing Party promptly, and in any event
no later than 48 hours after the Receiving Party discovers any unauthorized use,
disclosure, or loss of Confidential Information. The Receiving Party will
cooperate with the Disclosing Party in every reasonable way to help regain
possession of such Confidential Information and prevent its further unauthorized
use or disclosure.
```

**Fallbacks:** when the counterparty cannot commit to a fixed-hour deadline.

```approved-fallback
The Receiving Party will notify the Disclosing Party promptly after discovering
any unauthorized use, disclosure, or loss of Confidential Information, and will
cooperate with the Disclosing Party in every reasonable way to help regain
possession of such Confidential Information and prevent its further unauthorized
use or disclosure.
```

**Unacceptable:** no breach-notice obligation of any kind; notice limited to
breaches the Receiving Party deems material; no cooperation obligation.

**Rationale:** Without a notice obligation the Disclosing Party learns of a leak
only when the damage surfaces; prompt notice and cooperation preserve the
containment and injunctive-relief options. Adapted from the Bonterms NDA
playbook items (public domain).

## Use of AI systems

**Preferred:** Confidential Information may not be used to train or improve any
AI, machine-learning, or large language model system.

```approved
The Receiving Party will not itself or allow any third party to use any
Confidential Information to train or improve any artificial intelligence,
machine learning, or large language model (LLM) system of any kind.
```

**Fallbacks:** when the counterparty uses third-party AI tools in the ordinary
course of business and rejects a flat prohibition; the fallback still requires
the provider to be barred from training on the information.

```approved-fallback
The Receiving Party will not upload or transmit Confidential Information to any
third-party artificial intelligence, machine learning, or large language model
(LLM) system unless the Receiving Party first ensures that the provider of the
system is subject to binding obligations that (i) prohibit the use of the
Confidential Information to train or improve the system and (ii) are otherwise
no less protective of the Disclosing Party than this Agreement.
```

**Unacceptable:** silence on AI use combined with a broad license to use
Confidential Information "for the Purpose" through third-party services; any
express right to use Confidential Information for model training.

**Rationale:** Training ingestion is effectively irreversible disclosure — the
information cannot be returned or destroyed once embedded in a model. Adapted
from the Bonterms NDA playbook items (public domain).

## Residual knowledge

**Preferred:** No residuals clause. Propose deletion of any residuals provision;
there is no approved replacement text because the firm position is removal.

**Fallbacks:** when the counterparty insists on a residuals clause as a
condition of signing.

```approved-fallback
Notwithstanding anything to the contrary in this Agreement, the Receiving
Party's employees may use their unaided memories to retain and use general
knowledge, skills, experience and know-how learned during exposure to
Confidential Information, provided they do not intentionally memorize
Confidential Information for this purpose.
```

**Unacceptable:** residuals extending to information retained in documents,
copies, or other tangible or electronic form; residuals with no
unaided-memory limitation or no intentional-memorization exclusion; a residuals
clause operating as a license to the Disclosing Party's intellectual property.

**Rationale:** A residuals clause is a deliberate leak in the confidentiality
obligation; if one must be conceded, the unaided-memory formulation is the
narrowest market version. Fallback language adapted from the Bonterms NDA
playbook items (public domain).

## Assignment

**Preferred:** No assignment without the other party's prior approval.

```approved
Neither party may assign this Agreement without the prior approval of the other
party.
```

**Fallbacks:** when the counterparty requires a customary carve-out for
corporate transactions.

```approved-fallback
Neither party may assign this Agreement without the prior written consent of the
other party, except that either party may assign this Agreement without consent
to a successor in connection with a merger, acquisition, or sale of all or
substantially all of its assets, provided the successor agrees in writing to be
bound by this Agreement.
```

**Unacceptable:** free assignability by either party; assignment to a direct
competitor of the other party without consent, including under any
corporate-transaction carve-out.

**Rationale:** An NDA assigned freely can put Confidential Information in the
hands of an unvetted or competing successor; consent or a narrow
successor-in-interest carve-out keeps control with the parties. Approved
language adapted from the Bonterms NDA playbook items (public domain).
