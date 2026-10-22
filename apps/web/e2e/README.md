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

Run startup provider health and reconnection acceptance with the same environment:

```sh
PLAYWRIGHT_BASE_URL=http://127.0.0.1:5189 \
DOCHAUS_EVIDENCE_DIR=/tmp/doc-haus-provider-health-evidence \
bun apps/web/e2e/provider-health.ts
```

It covers launch/setup gating, one check per distinct selected model, verified,
expired-key, model-access and transient states, native scoped event failures,
and recovery after replacing a synthetic key. Desktop and 390px screenshots
also verify pointer targets and horizontal overflow. No real credentials or
provider traffic are used.

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
downloaded package XML. Prepare's line must count only the batch's proposals
still fresh after each accept. Its open items must stay listed, including after
the restart and reload, and a second seeded deliverable must show none of
them until the first is opened again.
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

Run the completed-run Open items regression with the same environment:

```sh
PLAYWRIGHT_BASE_URL=http://127.0.0.1:5189 \
DOCHAUS_EVIDENCE_DIR=/tmp/doc-haus-open-items-evidence \
bun apps/web/e2e/run-open-items.ts
```

It opens Details on a completed run and on a run paused at the final check,
each with 21 synthetic open items, at 1440×1000, 1280×720 and 400×860, shows
them all, then reads the Details scroller just past the Open items summary and
to its end. At each position the summary must stay in view at the scroller's
top edge, be the element under its own centre and have an opaque background in
the run header's colour. A real pointer click just past the summary must fold
the list. It sends no writes.

Run the page layout regression with the same environment:

```sh
PLAYWRIGHT_BASE_URL=http://127.0.0.1:5189 \
DOCHAUS_EVIDENCE_DIR=/tmp/doc-haus-page-layout-evidence \
bun apps/web/e2e/page-layout.ts
```

At 1440×1000 and 400×860 it sets every Page layout control in Settings →
Drafting and checks the agreement page preview and the saved body: flat keys,
with no new nested object. The first-page and firm-name checkboxes must hide
while page numbers or the draft line are off and keep their value when shown
again. In onboarding's details step it checks that default
jurisdictions only mark suggested legends, that nothing is chosen for the
lawyer, that the preview follows the firm name and legend, and that Save sends
the two legends as flat keys, empty when left alone. The clock is fixed so the
draft line's date is stable. Saves reach only the in-memory fixture.

Run the product design regression with the same preview environment:

```sh
PLAYWRIGHT_BASE_URL=http://127.0.0.1:5189 \
DOCHAUS_EVIDENCE_DIR=/tmp/doc-haus-design-evidence \
bun apps/web/e2e/material-expressive.ts
```

It captures Matters, creation/editing, all five libraries, saved conversation,
document drawer/reader, review, Settings/Appearance and onboarding at
375×812, 768×1024 and 1440×1000; covers empty/error/loading states, short
desktop workspaces, compact/large preferences, reduced motion, collapse hit
areas and unchanged matter request bodies. The 720×500 layout checks reflow
at the CSS viewport equivalent of a 1440×1000 window at 200% browser zoom.
All provider and generation traffic is intercepted or refused.

Set `AXE_SCRIPT` to a local axe-core browser bundle to assert zero automated
WCAG 2.2 AA-tag violations in each capture. There is no shipped axe dependency.
Set `DOCHAUS_BASELINE=1` to capture an older preview with new-design assertions
skipped. Screenshots/results remain outside public source. Run focused native
control regressions with `bun test/material-primitives.browser.ts` and
`bun test/jurisdiction-select.browser.ts`; use `PLAYWRIGHT_CHROMIUM_PATH` if the
installed Chromium differs from the package's default.
