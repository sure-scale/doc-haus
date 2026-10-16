---
name: redline-conventions
description: House conventions for editing matter documents — choosing exact-text anchors, picking redline vs tracked-changes, attributing changes to the firm, and applying playbook approved language byte-exact. Use whenever editing, redlining, amending, or applying tracked changes to a matter document.
---

Every edit to a matter document is a proposal: the tools record a pending
redline the lawyer accepts or rejects in the application, and the canonical
.docx is not modified until they accept. Your job is to locate the right text,
propose the right replacement, and explain it.

<anchors>
Both editing tools locate text by exact-text match, never by position or
offset. The anchor IS the address.

- **Copy anchors verbatim** from a verified source — a search-document citation
  excerpt, or text returned by read-document. Never type an anchor from memory
  or normalize its punctuation, capitalization, or spacing: a paraphrased
  anchor will not be found.
- **Make the anchor unique before calling.** Short phrases ("the Receiving
  Party", "30 days") recur throughout a contract; the tool anchors to the
  FIRST occurrence, which may be the wrong clause. Extend the anchor with
  surrounding words from the same sentence until it can only match the
  intended passage.
- **If the result reports more than one occurrence**, treat the anchor as
  ambiguous: the proposal may have landed on the wrong instance. Re-anchor with
  a longer, unique string and let the new proposal supersede the old one
  (the tools retire conflicting pending proposals automatically).
- **If the text is not found**, the tool says so and nothing is recorded.
  Re-read the passage (read-document / search-document) and anchor to text the
  document actually contains — never guess at a nearby clause or retry with an
  invented variant.
</anchors>

<tool-choice>
- `tracked-changes` — a surgical word or phrase swap: `find` exact text,
  `replace` it. Use for the smallest change that satisfies the instruction.
- `redline` — rewrite a whole clause: `clause` locates the paragraph (pass the
  citation excerpt or a sentence within it) and `replacement` replaces the
  paragraph's entire text. Use when the instruction reworks a clause rather
  than swapping a term — including when applying a playbook's approved
  replacement text.
- `word-integration` (action `replace`) — a silent, untracked edit. Use ONLY
  when the lawyer explicitly asks for a clean change with no tracked revision.
  Never choose it on your own judgment.
</tool-choice>

<granularity>
Edit at the smallest unit that achieves the position: a word before a phrase, a
phrase before a sentence, a sentence before the whole clause. Small redlines are
faster for the counterparty's lawyer to review and signal that the rest of the
clause is acceptable — a wholesale rewrite invites a wholesale counter.

- Rewrite an entire clause only when the surgical edits would leave it less
  readable than a clean replacement — and when you do, say so and why in your
  explanation of the change.
- Playbook ```approved``` fences are the deliberate exception: the fence is the
  firm-approved unit, applied byte-exact as a whole-clause replacement (see
  <playbook-language>). Never decompose a fence into smaller edits.
</granularity>

<no-silent-edits>
- Every change is tracked and every change is explained. After each proposal,
  state the document, the clause affected, the old wording, the new wording,
  and why — and that it is pending the lawyer's accept/reject review.
- Make exactly the change instructed. Do not also "improve" neighbouring
  wording, fix typos you were not asked about, or fold several instructions
  into one anchor.
- One instruction, one proposal. If a change spans several clauses, make one
  tool call per clause so each can be accepted or rejected independently.
</no-silent-edits>

<markup-deliverable>
When the lawyer asks for a markup or redline deliverable, preserve all three
review layers. They serve different purposes and none substitutes for another:

1. **Prioritized cover summary.** Lead the response with the material changes,
   ordered by legal or negotiating priority. For each, name the document and
   section, state the deficiency or purpose in one line, identify the specific
   source or cross-reference supporting it, and give a short strategy note.
2. **Visible change.** Record every proposed insertion and deletion through
   `redline` or `tracked-changes`; never silently replace the canonical text.
   The pending change shown in the application is the reviewable markup layer.
3. **Attorney review comment.** Pass a concise `comment` with every proposal.
   It must explain the deficiency or risk, why the proposed text addresses it,
   and the supporting matter source, playbook position, or cross-reference.
   The comment is stored beside the pending change and is non-operative. Never
   insert a literal comment marker or drafting rationale into document text.

Do not inflate the cover summary with immaterial edits. A single-change markup
may have one compact summary entry, one visible proposal, and one comment; a
multi-change markup keeps the same structure for every material change. If a
change lacks support, flag the gap instead of inventing a rationale or source.
</markup-deliverable>

<authorship>
Pass `author` on every redline and tracked-changes call, set to the firm or
the supervising lawyer's name as the matter uses it. Never leave it to default
and never use a tool or product name: the author field travels in the .docx to
opposing counsel, and a tool name in the revision history discloses how the
document was prepared. If you do not know the right name, ask before editing.
</authorship>

<playbook-language>
When a playbook provides text in an ```approved``` or ```approved-fallback```
fence, copy it into `replacement` byte-exact — no rewording, no "fitting it to
the contract's style", no dropped or added sentences. The fence is
firm-approved language; any deviation is unapproved language presented as
approved. Use a fallback fence only when its stated "when" condition is met,
and say which condition triggered it. If the fence contains a bracketed
`[insert ...]` placeholder you cannot resolve from the matter, stop and ask —
never send a placeholder into a client document.
</playbook-language>
