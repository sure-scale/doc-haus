# Contributing

Use Bun 1.4.0 or newer and read AGENTS.md plus the instructions inside the package you
change. CI checks both the minimum version and the latest stable Bun. Keep legal
functionality additive so OpenCode merges remain practical.
Use synthetic documents in tests and exclude credentials, documents, runtime
state, captures and generated output from commits.

Install both dependency graphs:

```sh
bun install --frozen-lockfile
bun install --cwd services/ingest --frozen-lockfile
```

Run checks from the owning packages:

```sh
bun run --cwd dochaus typecheck
bun run --cwd services/engine typecheck
bun run --cwd services/ingest typecheck
bun run --cwd apps/web typecheck
bun run --cwd apps/web build
bun run lint
(cd scripts && bun test test/start-runtime.test.ts test/bun-runtime.test.ts test/install-dependencies.test.ts test/updater.test.ts --timeout 30000)
(cd apps/web && bun test test)
```

Run the focused ordinary tests relevant to your change, including configuration,
preferences, routing and research consent. Engine integration tests use isolated
state and local scripted providers. Do not run `bun test` at repository root.
Preserve the WASM assets emitted by the web build. Include browser evidence for
UI changes and state clearly when external authentication was mocked.

Use conventional commit messages and explain the behavior, reason and relevant
verification. Do not include unrelated work in the change.
