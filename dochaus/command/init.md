---
description: Orient on this matter — read its intake record and documents, then write a concise matter profile.
agent: summarizer
---

Produce a matter orientation for the legal matter in the working directory. This
is the legal equivalent of onboarding to a project: understand what the matter
is, what documents it contains, and what stage it is at, so any future session
starts informed.

Do this in order:

1. Read `matter.json` and report the matter name, id, jurisdiction (if recorded),
   intake notes, and conflicts-check status.
2. Read `source-index.json` (if present) and list the documents by name with
   their character counts, noting the largest few. If `source-map.md` exists,
   summarize the section structure it shows.
3. Scan for durable work-product artifacts already present (`research.md`,
   drafts, memos, redlines, chronology) and note which exist.

Then write `MATTER.md` summarizing, in terse bullets:
- Matter objective (one sentence, from the intake record or inferred from the
  documents — mark inferred).
- Jurisdiction and conflicts status.
- Documents (names, types, relative sizes).
- Current stage and open items.
- A short "first step" suggestion for a new session.

Do not paste document bodies or long excerpts into MATTER.md — reference
documents by name only. If the matter appears empty (no matter.json and no
documents), say so plainly and stop.
