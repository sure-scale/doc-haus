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
