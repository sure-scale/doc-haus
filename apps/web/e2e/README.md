Run the setup browser regression against a separately started Vite or preview server:

```sh
PLAYWRIGHT_BASE_URL=http://127.0.0.1:5189 \
DOCHAUS_EVIDENCE_DIR=/tmp/doc-haus-setup-evidence \
bun apps/web/e2e/release-setup.ts
```

The runner uses `@playwright/test` and an already installed Chromium. When that
package is installed in another workspace, set `PLAYWRIGHT_MODULE` to its
`index.mjs` entrypoint. Set `PLAYWRIGHT_CHROMIUM_EXECUTABLE` if the installed browser version differs
from the Playwright package cache. It never installs packages or starts a backend.

All backend/auth traffic is intercepted with local fixtures; unexpected requests
are blocked and fail the run. Browser contexts are isolated per scenario. The
URL can target either private or public code; screenshots and the JSON result
are written only beneath `DOCHAUS_EVIDENCE_DIR`.

Run the focused desktop Settings keyboard regression with the same environment:

```sh
PLAYWRIGHT_BASE_URL=http://127.0.0.1:5189 \
DOCHAUS_EVIDENCE_DIR=/tmp/doc-haus-settings-evidence \
bun apps/web/e2e/settings-focus.ts
```

It covers initial focus, native modal Tab containment, Escape/Close/backdrop
dismissal, opener restoration, nested jurisdiction search and unsaved model
choices. Its Save check writes only to the intercepted in-memory config fixture.

Run the working Word review acceptance against a Vite server whose `/ingest`
proxy targets a free port and whose `/engine` proxy targets an unused port:

```sh
(cd apps/web && VITE_INGEST_URL=/ingest VITE_OPENCODE_URL=/engine \
  DOC_HAUS_INGEST_PORT=4801 DOC_HAUS_ENGINE_PORT=4397 \
  bun run dev -- --host 127.0.0.1 --port 5474 --strictPort)

PLAYWRIGHT_BASE_URL=http://127.0.0.1:5474 \
DOCHAUS_E2E_INGEST_PORT=4801 \
DOCHAUS_EVIDENCE_DIR=/tmp/doc-haus-word-evidence \
bun apps/web/e2e/working-word-review.ts
```

Unlike the checks above, this runner starts the real ingest service itself on
`DOCHAUS_E2E_INGEST_PORT`, with a temporary data root and settings directory,
and restarts it mid-run. Its legal generation reaches a scripted local engine
fixture, never a provider. Browser engine traffic (`ready`, session and agent
lists, the event stream) and the release update check are intercepted. It
seeds a synthetic Word deliverable with a numbered list and a table, then
covers Prepare, single acceptance, stale proposals, API refusals, a redline
export that leaves stale proposals out, restart persistence, rejection and the
downloaded package XML.
The temporary data root is removed afterwards.

Run the matter chat layout regression with the same environment as the
setup and settings checks:

```sh
PLAYWRIGHT_BASE_URL=http://127.0.0.1:5189 \
DOCHAUS_EVIDENCE_DIR=/tmp/doc-haus-composer-evidence \
bun apps/web/e2e/matter-composer.ts
```

It opens a blocked drafting run with its requirement-check form at
1512×806, 1440×900, 1280×720 and 400×860, then checks that Continue drafting,
Check the full draft and Details are in view and under their own centres, that
arrowing through the requirements keeps the focused one inside the scrolling
list, that the composer stays inside the chat card and clear of the matter
status bar, that Send is the element under its own centre and takes a real
pointer click, and that the narrow window has no horizontal overflow. The click's request is answered by the
fixture; nothing reaches a live run.
