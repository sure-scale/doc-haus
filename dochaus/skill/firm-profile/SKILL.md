---
name: firm-profile
description: Defines the practice profile — a plain-English markdown file at .dochaus/profile.md in the matter directory recording the firm's escalation rules, house style, risk calibration, and jurisdiction defaults — and how to create, load, and apply it. Use when setting up a new firm or matter, onboarding, when asked to "use our house style" or record firm preferences, or before any review, draft, or redline that should follow the firm's standing positions.
---

The practice profile is the firm's standing instructions for a matter, written
in plain English with the supervising lawyer. It lives at one fixed location:
`.dochaus/profile.md` inside the matter directory (the session's working
directory). One file, four fixed sections, nothing else.

<loading>
At the start of any review, draft, or redline, `read` `.dochaus/profile.md`.

- **Present** — apply it throughout. The clause-library and the playbooks state
  each clause on a fixed three-position schema — Preferred, ordered fallbacks
  with "offer when" conditions, and an explicit walk-away floor with a
  role-addressed escalation trigger — and those positions are firm-configurable
  **defaults**. Profile positions override those defaults: a Risk Calibration
  non-negotiable raises a clause's floor, a standard concession lowers it, and
  the Escalation Rules contact is the role the schema's floor escalation routes
  to. The profile does not override a bound playbook's approved replacement
  *text* — a playbook is the firm's reviewed wording for that document type —
  but it does set which fallback rung and floor that wording is measured
  against.
- **Conflict** — when an instruction, playbook position, or document term
  conflicts with the profile, flag the conflict explicitly, quoting the profile
  line it crosses. Never silently follow either side; the lawyer resolves it.
- **Anything matching an Escalation Rule** stops with a clear statement of
  which rule was triggered — do not proceed and escalate in the same breath.
- **Absent** — proceed on balanced market-practice defaults, and when the
  moment is natural (new matter, onboarding, "use our preferences") offer to
  create one via the interview in `references/interview.md`.
</loading>

<sections>
The profile has exactly these four `##` sections, in this order.

- **Escalation Rules** — what always goes to a human before it happens: work
  product that may never leave the firm without named sign-off, exposure or
  deal-value thresholds that loop in a partner, topics never handled without
  escalation (privilege calls, regulator contact, conflicts), and who the
  escalation contact is.
- **House Style** — the drafting voice (plain English or traditional, "shall"
  vs "must/will"), defined-term conventions (where terms are defined, how they
  are marked on first use), numbering and heading conventions, date and number
  formats, and boilerplate the firm always includes or never includes. **Default
  voice when the profile is absent or leaves it unspecified: plain language** —
  short sentences, active voice, no needless archaisms or Latin, jargon defined
  on first use. Guardrail: eliminate redundancy only where it is genuinely
  redundant — keep a doublet (e.g. "indemnify and hold harmless") when each word
  carries a distinct legal meaning; plain language never strips a load-bearing
  term.
- **Risk Calibration** — the negotiation posture per contract type, given the
  side the firm usually sits on: **aggressive** (open firm-favorable, concede
  slowly), **balanced** (market-standard two-sided positions), or
  **conservative** (accept reasonable counterparty terms, redline only material
  risk). Plus non-negotiables and standard concessions. Non-negotiables and
  concessions tune the clause-library/playbook three-position ladder for the
  firm: a non-negotiable raises a clause family's walk-away floor (and may move
  its Preferred), a standard concession marks a fallback rung the firm offers
  readily. List a clause family here only to override its market-balanced
  default — silence leaves the shipped default in force.
- **Jurisdiction Defaults** — default governing law and venue for new
  documents, dispute-resolution preference (courts or arbitration, and where),
  the jurisdictions the firm advises in, and what happens when a document falls
  outside them (flag as outside scope, never improvise foreign law).
</sections>

<jurisdiction-risk>
Jurisdiction is a risk-calibration axis, not just a setting. Model confidence is
uninformative — it does not correlate with whether the answer is right — so
abstention and qualification must be driven by the structural jurisdiction
signals below, not by how sure the answer feels.

- **Name the assumed governing law first.** Before any jurisdiction-sensitive
  conclusion, state the governing law you are reasoning under and where it comes
  from: the document's choice-of-law clause, the matter's jurisdiction pack, or
  an explicit instruction. If none of those fixes it, say the governing law is
  undetermined and stop — do not pick one.
- **Never accept an implied jurisdiction premise.** If a question assumes a
  jurisdiction without stating one ("is this non-compete enforceable?"), surface
  the missing premise and name what you assumed before answering; do not silently
  default to a familiar jurisdiction (e.g. U.S. federal or the firm's home law).
- **Grade confidence down — qualify, then escalate — when the governing law is:**
  state/local or sub-national rather than national; non-US (hallucination rises
  sharply for less-prominent and non-US jurisdictions); driven by lower-court or
  unsettled authority rather than apex-court or statute; multi-jurisdiction (the
  worst category for both AI and lawyers); or recently changed (governing law is
  a jurisdiction-and-time pair — confirm the rule is current and was not
  superseded). Any one signal qualifies the conclusion; two or more route it to
  the Escalation Rules contact.
- **Governing law vs forum.** Treat the choice-of-law clause and the
  forum/venue/arbitration-seat clause as separate. Flag misalignment (one law,
  another forum) and note that the forum's own mandatory law and public policy
  may displace the chosen law.
- **Public-policy / mandatory-law override.** Even with a valid choice of law,
  flag where the forum's mandatory rules or public policy can override it
  (employee protections, consumer law, competition, data protection); a chosen
  law does not guarantee the chosen result.
- **Cross-border CISG prompt.** For a cross-border sale-of-goods contract
  between parties in different countries, check whether the UN Convention on
  Contracts for the International Sale of Goods (CISG) applies by default and
  whether the parties intended to exclude it; flag an unaddressed CISG position.
- **Multi-jurisdiction workflow.** When more than one governing law is in scope,
  do not blend them. Handle each jurisdiction procedurally and completely on its
  own — the same enumerated checks per jurisdiction, without fatigue or
  shortcutting later ones — then reconcile in a final pass: note where the
  jurisdictions agree, where they diverge, and which divergences need a human
  decision.
</jurisdiction-risk>

<skeleton>
A new profile starts from this skeleton. Keep entries as short plain-English
lines a non-lawyer could follow; write `Not specified` rather than guessing an
answer the lawyer never gave.

```markdown
# Practice Profile

## Escalation Rules
- Nothing leaves the firm to a counterparty, court, or regulator without
  sign-off from [name/role].
- Loop in [name/role] when exposure or deal value exceeds [amount].
- Always escalate: [topics].

## House Style
- Voice: [plain English (default) / traditional]; obligations use ["must" / "shall"].
- Defined terms: [convention, e.g. bold-quoted on first use, Definitions
  section up front].
- Formats: dates as [format]; numbering [convention].
- Always include: [boilerplate]. Never include: [boilerplate].

## Risk Calibration
- [Contract type] (we are usually the [side]): [aggressive / balanced /
  conservative]. Non-negotiable: [terms]. Standard concessions: [terms].

## Jurisdiction Defaults
- Governing law and venue for new documents: [jurisdiction], unless [exception].
- Disputes: [courts / arbitration under [rules], seated in [place]].
- We advise on [jurisdictions] law only; anything else is flagged as outside
  scope.
```
</skeleton>

<creating>
To create or update a profile, run the interview in `references/interview.md`:
pre-fill every answer you can from the matter's existing documents, ask only
the gaps in small batches, confirm a compact summary, then write the file to
`.dochaus/profile.md` (create the `.dochaus/` directory if needed). If you
cannot write files, output the complete file in one fenced block labeled with
that path for the lawyer to save. Updates edit the existing file section by
section — never drop sections the lawyer did not ask to change.
</creating>
