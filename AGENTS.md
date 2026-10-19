# doc.haus contributor instructions

doc.haus adds legal-document workflows to OpenCode. A matter is a project
folder; a conversation is a native session. Keep upstream primitives intact.

Read scoped AGENTS.md files before changing a package. Verify repository, common
Git directory, branch, dirty state and remote before staging or committing.
Preserve unrelated work. Use a named branch and a focused task worktree.

## Upstream and configuration boundaries

Keep upstream merges practical: build legal functionality in dochaus/,
services/engine/, services/ingest/ and apps/web/. Changes to upstream packages
need a written justification and an isolated commit. Preserve component notices.

The legal runtime uses shipped legal agents/tools/skills plus installation
settings. Preserve its disabled global/project/WellKnown discovery boundaries.
Do not import global configuration, instructions or MCPs. Reusing an existing
OpenCode sign-in requires explicit selection and copies only that connection
into this installation; never import unrelated connections or session state.
Store user settings in runtime/config/opencode/opencode.json through ingest's
validated writer. Never save normal settings into tracked product configuration.
Preserve durable session and provider-auth storage when tightening this
configuration boundary.

## Focused model jobs

Do not ask one model or agent to do too much at once. This is a primary
architectural requirement for product model jobs and development agents alike.

- Give each job one concrete question or decision and a clear output contract.
  For checklist verification, default to one item against its relevant document
  passages; batch only closely related items when the combined job stays small.
- Supply only the instruction, source evidence, target passages and material
  qualifications or dependencies. Do not send the whole matter; allow targeted
  retrieval when more evidence is needed.
- Keep extraction, drafting, linking, verification and reconciliation in
  separate jobs, not in one drafting prompt or one final reviewer.
- Decompose an oversized job before raising its prompt, context window,
  reasoning budget or retries. Run jobs in parallel within the concurrency
  limit, with explicit handoffs and bounded reconciliation.
- Test complete workloads through these jobs: complete coverage at scale, not
  one giant model call. Carry related clauses, documents and exceptions through
  targeted references.

## Legal content

Ground content in the matter's sources and verify every quotation and citation
against the canonical source text. Keep the review stages. Never fabricate
evidence, corrected legal inputs or human approval, and never auto-answer a
human question.

Record uncertainty as an open item; never resolve it by assumption. An optional
unsupported assertion may be omitted with its limitation recorded. Missing
material facts or authority become a body placeholder, an open item or a focused
human question. A placeholder disclosed in the memo does not block a review
draft; silent unsupported content does. Recording an open item supplies no
answer or human release. Never silently drop an essential requirement or weaken
a legal check to pass.

## Runtime design

doc.haus is self-hosted. Prefer a working product to mechanical safeguards, and
never weaken a legal check.

- A stage session reads matter files and writes only its own output file. A
  deterministic checker accepts it or returns specific errors for a bounded
  retry. Use OpenCode's existing tool-error and retry mechanisms first, and
  improve a tool's or checker's error before adding a controller or stage. Never
  ask a model to repair bookkeeping or record the same fact twice.
- Do not add new bookkeeping between internal stages: no hash chains, receipts,
  claims, progress tokens, ownership records, new locks or ledgers. When such a
  check blocks a working product, delete it or make it a non-blocking record.
  New stage guards are limited to atomic whole-file writes, a per-paragraph
  content hash at apply, a monotonic revision integer, content-keyed resume, and
  native session history plus Snapshot.
- Keep one writer path for lean run state and the body docx: the lean workflow
  (dochaus/workflows/lean). Amendments, Continue requests and human answers are
  queued and processed in order. Overlapping changes are consolidated before
  change-review and are never skipped; a change that cannot be applied becomes
  a visible open item.
- The lean profile is the drafting route and recovery is a re-run: a stage
  without a valid output re-runs from its inputs, completed stages are not
  repeated, and an unchanged state or a run stopped as blocked or failed cannot
  loop. Continue never fabricates an output, resets the retry count or answers a
  human question. On a 409 or lost response from the `/:matter/lean/*` routes,
  read `.dochaus/lean/state.json` before acting; never loop requests.
- Fix the shared stage, checker or runtime path, never a matter-specific prompt,
  manual continuation, fixture value or one-off artifact edit. The web product
  path is production: verify through the web UI or the same web and engine APIs.

## No benchmark shaping

Improve general quality; benchmark gains are a side effect, never the goal.
Never encode one evaluation item's answer in the product. Nowhere under dochaus/
(agents, skills, tools, plugin, lib, commands, workflows, opencode.json or any
other configuration) add:

- domain specifics (statutes, section numbers, doctrines, regulators, cases)
  included because an evaluation expects them;
- values taken from an evaluation task (parties, amounts, dates, identifiers,
  jurisdictions);
- checks that mirror one grader's checklist rather than a principle applied to
  every such document;
- hints that prime an expected answer, such as illustrative lists,
  parentheticals or "also check" prompts not drawn from general practice;
- tool ordering, retrieval boosting, stage-contract, retry or model-routing
  changes tuned so a known task scores better;
- benchmark names or task IDs, except inside evaluation tooling.

Test each line: if deleting it leaves quality on unseen documents flat while a
known benchmark score drops, it was overfitting; delete it.

Evaluation tooling is diagnostic only. It may send the initial prompt through
the real APIs and record outputs and generic runtime failures. It never
continues sessions, auto-approves permissions, retries generation or recovers
workflows. Reusable improvements belong in the product.

## Changes and checks

Use conventional commits. Keep changes additive, TypeScript strict and helpers
close to their callers. Avoid new abstractions without a concrete need. Keep one
brief change note per adjustment: cause, changed scope, affected consumers and
regression evidence. Share fixtures only for shared contracts and keep
stage-specific setup local. A fixture correction is not a product fix, and a
focused pass alone does not prove other stages remain sound.

Run package checks from the owning package using bun run typecheck, including
before a live end-to-end run. The root bun test intentionally fails. See
CONTRIBUTING.md for check commands and README.md for startup. Never run paid
model probes without explicit intent.
