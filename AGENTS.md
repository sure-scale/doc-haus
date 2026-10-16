# doc.haus contributor instructions

doc.haus adds legal-document workflows to OpenCode. A matter is a project
folder; a conversation is a native session. Keep upstream primitives intact.

Read scoped AGENTS.md files before changing a package. Verify repository, common
Git directory, branch, dirty state and remote before staging or committing.
Preserve unrelated work. Use a named branch and a focused task worktree.

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

Split substantial work into focused agents with one concrete question and a
small relevant context. Keep extraction, drafting, linking, verification and
reconciliation separate. Test complete coverage through those focused jobs.
Ground legal assertions in sources, verify quotations and citations, and record
missing facts as visible open items. Never fabricate evidence or human approval.

Use conventional commits. Keep changes additive, TypeScript strict and helpers
close to their callers. Avoid new abstractions without a concrete need.

Run package checks from the owning package using bun run typecheck. The root
bun test intentionally fails. See CONTRIBUTING.md for check commands and
README.md for startup. Never run paid model probes without explicit intent.
