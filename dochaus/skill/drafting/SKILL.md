---
name: drafting
description: How to draft a new document into the matter — template-first via list-templates/draft-document, from scratch only when no template fits. Use whenever asked to create, draft, or prepare a new document.
---

Drafting creates a NEW .docx in the matter. Changing an existing document is the
redline workflow, not this one.

<workflow>
1. **Classify and gather first.** Identify the document family, intended use,
   controlling sources, and required inputs at runtime. Take the inputs from the
   conversation and the matter's existing documents (`search-document`); ask the
   lawyer only for what is genuinely missing. A working draft may retain an
   explicit placeholder while inputs are collected, but it is not final merely
   because it renders. Preserve a visible blank or clearly identified unresolved
   statement in a deliverable only where the runtime-classified document family
   and intended use conventionally permit it; otherwise stop and ask for the
   missing input or decision. Do not assume a form, relationship, or legal
   posture the inputs do not establish.
2. **Template first.** Call `list-templates`. If a template matches the request,
   run the guided interview (below) to collect every placeholder value and every
   optional-section decision, then call `draft-document` ONCE with `template` +
   `fills` + `omit`. Placeholder text must match the list-templates output
   exactly, brackets included.
3. **From scratch only when no template fits.** Write the complete body as
   markdown and call `draft-document` with `content`. Use the headings, sequence,
   and formatting conventional for the runtime-classified document family, with
   one blank line between blocks. Load the clause-library as a position baseline
   only for contracts or other negotiated instruments for which that skill is
   applicable; for other families, use the matching supplied authority,
   checklist, or playbook. Mark anything you cannot know with a unique bracketed
   instruction placeholder. For lawyer-facing unresolved drafting
   questions, use `[open item: ...]` so the rendered DOCX can highlight the item.
   Make every placeholder unique because identical placeholders can only ever be
   filled with the same value. A placeholder is a visible working-draft marker,
   never permission to replace an unknown fact, authority, or legal effect with
   plausible complete-looking text.
4. **Space the document like the finished Word file.** A blank line in the
   markdown becomes a paragraph break — anything you want on its own line
   must be its own block. Give standalone structural elements and input fields
   the spacing their document family requires. Lay the body out the way the
   document type is conventionally formatted, with readable paragraph breaks and
   no walls of text.
5. **Completeness self-audit before finalizing.** A deliverable fails on a
   single missed operative fact, instruction, required component, or internal
   reference, so re-check the draft
   against an explicit checklist rather than trusting recall. Use the
   `todowrite` tool as that checklist, and run it as a **two-way
   reconciliation** between the controlling sources and the draft:
   - **Enumerate** one item per atomic source element or required family
     component the draft must carry, plus every internal reference the draft
     makes. Each item is specific and checkable, not a category.
   - **Source → draft (nothing operative dropped).** Every operative source
     fact, instruction, value, qualification, and relationship must appear in
     the draft or be explicitly routed as an open item. Re-check
     item-by-item by reading the draft back (`read-document`) and mark each item
     `completed` only once you confirm that element is present and correct. An
     item you cannot confirm stays `in_progress` and becomes an Open item; mark
     `cancelled` only for elements a keep-or-omit decision dropped.
   - **Draft → source (nothing concrete invented).** Every matter-specific fact,
     instruction, value, and legal consequence in the draft must trace to a
     source, an applicable approved position, or a clearly bracketed working-draft placeholder.
     Content that traces to none is fabricated: remove or bracket it and record
     it as an Open item.
   - **Second-pass absence sweep:** before drafting, write down the standard
     required components for the runtime-classified document type. Derive them
     from the applicable source, authority, checklist, or playbook rather than
     from a fixed cross-family list. Add one `todowrite` item per component and
     confirm each is present in the form its family requires. A
     required component entirely absent is a finding to report, never a silent
     pass.
6. **Fill out omitted mechanics by document family.** Before finalizing, derive
   the mechanics required to make each promised outcome or required component
   work from the runtime-classified family, controlling sources, and applicable
   authority. For contracts or other negotiated instruments, use the applicable
   contract checklist and clause library; for other families, use only their
   matching authority, checklist, or playbook. For each source silence, either
   draft a supported conventional mechanism flagged drafted-for-review or record
   the gap as an Open item. If authority or source support is missing, do not
   hide that gap with neutral-sounding language. Never import mechanics from one
   family into another or ship a promised outcome with no supported implementation
   behind it.
7. **Completeness principles for operative content.** Beyond the source
   reconciliation above, audit each operative part against the
   `<completeness-principles>` below when the runtime document family makes a
   principle relevant. Content can reproduce every source value correctly and
   still fail to achieve its intended legal or procedural function.
8. **Report a structured change-memo.** Name the document you created and emit
   the memo per the `<change-memo>` format below — BLUF, issues table, the
   dual-entry open-items taxonomy, and the redline appendix — so the lawyer can
   tell a deliberate decision from an unfilled gap at a glance.
</workflow>

<completeness-principles>
Content can be textually complete yet functionally incomplete. After the
self-audit and fill-out passes, apply each principle only when supported by the
runtime-classified document family, controlling sources, or applicable
authority. Do not presume that every document contains obligations, defined
terms, elections, adverse sides, or a conflicts analysis.

1. **Terminology and scope completeness.** When the document's operation turns
   on a term that the applicable legal or procedural regime gives a specific meaning
   broader or narrower than its ordinary-language meaning, define that term to
   its governing-regime scope rather than leaving it to lay reading. Otherwise
   the document may silently fail to reach, or exceed, what the regime intends.
   Do not add a definition when the family convention or authority does not call
   for one.

2. **Mechanism completeness.** When a required outcome depends on a process,
   confirm that the draft includes the actors, sequence, conditions, and evidence
   the applicable family or authority requires. Include alternatives only when a
   controlling source, approved family position, or governing regime recognizes
   them; never invent optional paths from a cross-family checklist.

3. **Role and interest completeness.** When the classified document family
   requires analysis of roles, capacity, interests, or relationships, cover every
   person and entity within the scope set by the controlling sources or
   authority. Do not assume such an analysis is required, and do not stop at one
   organizational level when the governing scope reaches further.

When the requirements of a document type or applicable authority are uncertain,
derive them at runtime from authoritative sources via the `legal-research` skill.
Research must derive generally applicable requirements, not reverse-engineer a
desired conclusion; a derived requirement must apply to every document of that
classified type, not just the matter at hand.
</completeness-principles>

<change-memo>
Every draft built from a template or from source documents ships with a
structured change-memo, not a prose summary. The memo makes the drafting
reasoning auditable and separates a placeholder/gap from a decision. Use this
exact structure:

1. **BLUF (bottom line up front).** Two or three sentences: what document was
   created, the controlling source(s) it was built from, and the single most
   important thing the lawyer must decide or supply before its intended use.
2. **Issues table (two columns).** One row per material change or judgment call,
   in document order. The `Location / component` column stores the affected location, section,
   or component using the document's own locator. The right column records the
   change made AND the rationale for it — what the
   draft does, how it deviates from the template or source, and why. Show every
   substantive change; a change with no recorded rationale is not done.

   | Location / component | Change and rationale |
   | -------------------- | -------------------- |
   | Document locator | What changed; why. |

3. **Open items (dual-entry taxonomy).** List every unresolved item, each tagged
   with exactly one of these four types so its nature is unambiguous:
   - **Unfilled placeholder** — a `[insert ...]` the lawyer or client must
     supply (a fact only they know).
   - **Unresolved source conflict** — two source documents disagree and the
     draft could not pick a controlling term without instruction.
   - **Promised term, no mechanics** — the source requires an outcome or
     component but omits the implementation needed to make it work; flag exactly
     what is missing.
   - **Escalation hold** — an item that needs a person/role to decide before the
     draft is final (a position that breaches a firm floor, a jurisdiction
     question).
   Every unresolved item that remains visible in the document must appear here
   with the same location and status. The memo cannot describe an item as resolved
   when the body leaves it open, and memo disclosure cannot make invented body
   text acceptable.
4. **Redline appendix.** The list of every deviation from the template or source
   text in document order as the final section, so the body of the memo stays
   readable and the full change record is still available in one place.
</change-memo>

<guided-interview>
Template drafting is a short interview, not a form. The template's variables ARE
its placeholders — the wording of each tells you what kind of answer it wants.
The template's `optionalClauses` are keep-or-omit decisions, regardless of
whether the classified document family calls them clauses, sections, or
components.

1. **List the variables.** From the list-templates entry: every placeholder plus
   every optional component.
2. **Prefill before asking.** Take values from the conversation and the matter's
   existing documents (`search-document`). Never ask for something already known.
3. **Ask the rest in small logical batches** organized by the document family's
   own dependencies. Ask each optional component as keep-or-omit with a one-line
   plain-English description of what it does. Never one long questionnaire and
   never one question per turn when a batch reads naturally.
4. **Confirm a compact input schedule** before drafting: every variable with its
   value, every optional component marked keep or omit. A variable the lawyer
   leaves open stays a placeholder in the working draft — say so and do not call
   the document final or ready for external use.
5. **Draft ONCE.** One `draft-document` call with `template`, all `fills`, and
   the declined components in `omit`. Never assemble the document piecemeal.
</guided-interview>

<template-creation>
Creating a reusable template (`create-template`) adds a drafting base to the
firm's global library, shared across every matter — distinct from drafting a
document into one matter.

**When to create one.** The user asks to create or save a template directly, asks
to turn an existing matter document into a template, or you are drafting a
document type that `list-templates` shows no match for and a template would help.

**Body conventions.** Use the same family-specific markdown and structure as
from-scratch drafting. The one difference is total: a template carries no real
client or matter data. Replace every matter-specific value or identifier with a
unique descriptive `[insert ...]` placeholder, never a bare placeholder and
never the same placeholder twice (draft-document fills by exact text, so
duplicates collapse to one value).

**Optional components.** A section or component the lawyer may keep or drop
appends `[optional: short-name]` to the end of its heading. The short-name is
unique per template, lowercase, and hyphenated. Its extent is the heading plus
everything under it up to the next same-or-higher heading; on a non-heading
paragraph the marker makes just that paragraph optional. At draft time the
component is kept (marker stripped) unless its name is passed to
`draft-document`'s `omit`.

**Turning an existing document into a template (PII scrub).** Call
`read-document` to load the full text, then rewrite the whole body replacing every
client-specific detail with a placeholder. Before calling `create-template`, scrub
all personal, organizational, matter-specific, and identifying data. None may
survive into the template — only `[insert ...]` placeholders.

**Review before creating.** A template is drafted from many times, so it must be
right before it enters the library. Spawn the document-reviewer subagent (task
tool) over the proposed body once: structure must match the document type and
title, the type's required components must be present, and authority-specific
language must not be baked in unless the template is intentionally limited to
that authority. Apply Must-fix
findings to the body before calling `create-template`; one round, do not loop.

**Report.** After `create-template`, name the template and list the placeholders
it exposes, so the user knows what every future draft will be asked to fill.
</template-creation>

<rules>
- Pick a descriptive file name derived from the document's subject and type.
- Never overwrite: if the name is taken, pick another unique name.
- After drafting, refinements go through the normal editing tools
  (tracked-changes / redline) on the new document.
</rules>
