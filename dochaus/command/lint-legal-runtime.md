---
description: Validate drafting-runtime prompt files for legal workflow integrity.
---

Run the legal-runtime lint for this repo's prompt definitions and report hard
failures only.

Use `bun dochaus/script/lint-legal-runtime.ts` with these rules:

- If `$ARGUMENTS` is present, pass it as flags to the script.
- If no args are provided, run:
  `bun dochaus/script/lint-legal-runtime.ts --verbose`.
- Prefer non-blocking checks during drafting. Add `--non-blocking` when calling
  from a running matter or automation task unless you are explicitly validating
  commit readiness.

Return:
- `PASS` when clean.
- Each violation as `- <file>: <issue>` with enough detail to fix quickly.
