---
name: verify-work-product
description: Run the six deterministic verification gates over a deliverable before filing, sending, or exporting — citation anchoring, value grounding, research-source audit, render readiness, redline anchor integrity, privilege/export gate. Use when asked to verify, pre-flight, sanity-check, or gate-check a document before it leaves the workspace; before filing, sending, or exporting; or at any drafting, review, or redline stage boundary.
---

Every claim a deliverable makes about a matter is a load-bearing assertion.
Verification is the discipline of checking each claim mechanically against
the matter's indexed documents and source-term artifacts before the
deliverable leaves the workspace. Call the `verify` tool to run the six
gates in one pass; do not call the underlying primitives individually
unless the lawyer asks to re-run a single gate after a fix.

<procedure>
1. **Identify the deliverable.** If the lawyer names a document, pass it as
   the `deliverable` argument. Otherwise let the tool auto-discover (it
   looks for the assembled or primary deliverable markdown). Memos are not
   auto-discovered — they are usually meta-files (change memos, research
   memos, coverage reports) rather than the primary deliverable — so to
   verify a memo, pass its filename explicitly via `deliverable`.
2. **Run `verify`.** One call, no arguments beyond the optional deliverable,
   runs all six gates and writes `verification-report.md` to the matter.
3. **Read the report's overall verdict.** Three states:
   - `PASS` — every applicable gate passed. Some gates may be
     not-applicable (no research, no redlines, no export queued); those are
     not failures.
   - `PARTIAL` — at least one gate failed. List the failed gates and their
     findings; the deliverable is not ready to leave the workspace until
     every FAIL is fixed or explicitly signed off.
   - `BLOCKED` — a gate threw (runtime/transport failure, e.g. a fetch
     failed). Treat as a hard stop; report the blocking gate and stop.
4. **For each FAIL, fix the underlying deliverable, not the report.** A
   FAIL on `value-grounding` means a number traces to no source: re-anchor
   the number to a source term or, in a working draft, replace it with an
   explicit unresolved marker. A FAIL on `citation-anchors` means a quote is not verbatim
   in the indexed documents: re-quote from `search-document` or `cite`.
   Re-run `verify` after the fix.
5. **Apply the semantic unresolved-content check.** The deterministic
   `render-readiness` gate distinguishes recognized descriptive markers from
   malformed blanks; it does not decide whether an unresolved item is acceptable
   for the requested use. Before final delivery, locate every remaining marker,
   blank, or unresolved statement. Keep one only when it is conventional for the
   runtime-classified document family and intended use, remains visibly
   unresolved, and is mirrored in the memo or report with its location and next
   action. Otherwise resolve it from grounded support or stop. Never turn missing
   facts or authority into neutral complete-looking prose merely to obtain PASS.
6. **Terminal state.** The deliverable is safe to file, send, or export only
   when the semantic unresolved-content check above is satisfied and either the
   report reads `PASS` or every remaining FAIL has a recorded lawyer sign-off in
   the matter (named counsel, date, rationale).
   Never soften a FAIL into "probably fine" — either fix it or escalate it.
</procedure>

<gates>
| Gate | Checks |
|---|---|
| `citation-anchors` | Every quoted passage in the deliverable appears verbatim in the matter's indexed documents. |
| `value-grounding` | Every dollar amount, percentage, ISO date, and year in the deliverable either appears in a source-term artifact (`binding-decisions.md`, `key-dates.md`) or sits inside an `[insert ...]` placeholder. |
| `research-sources` | URLs in `research.md` resolve and match their claimed titles; authorities are fetched, not asserted from memory. |
| `render-readiness` | No bare `[___]` or empty brackets remain. A recognized descriptive marker may be syntactically valid for a working draft, but does not by itself make a final deliverable ready. |
| `redline-anchors` | Every pending redline in `pending-redlines.json` still anchors to text present in its target document. |
| `privilege-export` | No entry in `redaction-log.md` is flagged both `privileged` AND `export-requested`. |
</gates>

<non-goals>
- Does not replace lawyer judgment. Gates raise findings; the lawyer
  decides whether a finding is material.
- Does not invent thresholds. No dollar amounts, percentages, statute
  citations, jurisdiction codes, or doctrine names anywhere in the skill or
  the tool. Every threshold is the matter's own.
- Does not run the renderer. The `render-readiness` gate checks the seam;
  the drafting path does the actual render and has its own gates.
- Does not cross matters. Every gate is scoped to `ctx.directory`. A gate
  in matter A never reads matter B.
- Does not edit the deliverable. Verification only writes
  `verification-report.md`; fixes land through the normal editing tools
  (`redline`, `tracked-changes`).
</non-goals>

The verdict is final for this gate run only. A subsequent edit to the
deliverable invalidates the report; re-run `verify` after every material
change.
