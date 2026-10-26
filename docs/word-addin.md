# Open in Word

Use **Open in Word** on a Word document in a matter, then open the downloaded
copy in desktop Microsoft Word. The doc.haus pane identifies its matter and
document. Choose **Save to doc.haus** and wait for confirmation before closing.
**Done in Word** saves, confirms the final document, clears its open badge and
closes that document. If Word cannot close it, the pane keeps the saved
confirmation visible so you can close it manually.

The pane also checks saved content in the background. This is best effort:
rapid saves may become one version, and closing Word can interrupt an upload.
Reopening a linked copy can recover its last saved state. Native Word Save alone
does not confirm that doc.haus received it. When offline, keep the file and use
Save to doc.haus after reconnecting.

## Install on a desktop

The doc.haus installer offers setup when Word is installed. Existing users can
run this from their doc.haus folder, with the same data-root settings as startup:

```sh
bun scripts/setup-word-addin.ts
```

Setup registers the XML add-in for the current user: Word's WEF folder on macOS,
or its Office 16.0 developer sideload registry entry on Windows. It does not open
Word. Restart Word after setup. This sideload route requires Office to allow
developer add-ins; firms can instead deploy the generated manifest through
Microsoft 365's admin centre. Desktop Word on Windows and Mac is the target;
Word for the web is outside this feature.

For a local installation, setup creates an installation-specific localhost TLS
certificate and private key under the private runtime directory. It asks the
operating system to trust that certificate in the current user's account.
The certificate is a TLS server leaf, not a certificate authority, and is valid
only for localhost/loopback. Setup never disables certificate verification.
The private key stays with the installation. Re-run setup to renew an expired
certificate, then restart doc.haus and Word. Do not copy runtime keys into Git.

The default pane address is `https://localhost:4511`. Set `DOCHAUS_WORD_ORIGIN`
before setup to choose another localhost port. Multiple installations need
different ports and their own private data roots. Only one manifest with the
doc.haus add-in ID can be registered per Word user; setup selects that instance.

Configuration is saved in `runtime/config/word-addin.json`, outside source.
For an already trusted certificate, configure `DOCHAUS_WORD_ORIGIN`,
`DOCHAUS_WORD_TLS_CERT` and `DOCHAUS_WORD_TLS_KEY`. Startup serves the pane and
its limited receiver over HTTPS, alongside the normal doc.haus services.

## Firm, Docker or remote hosting

Use an HTTPS origin reachable and trusted by each Word desktop. Configure the
server with:

```sh
DOCHAUS_WORD_ORIGIN=https://word.example.org
DOCHAUS_WORD_EXTERNAL_HTTPS=1
```

Terminate TLS at your existing reverse proxy. Forward `/word-addin/*` and only
POST/OPTIONS `/word/saves`, `/word/heartbeat`, `/word/done` to ingest. Do not
expose the rest of the ingest API through this Word receiver. In Docker, build
the pane with `bun run --cwd apps/word-addin build` and include its `dist` assets
with the server. The normal start scripts build these assets automatically.

Download the manifest from `https://word.example.org/word-addin/manifest.xml`
and deploy it to Word users. Alternatively, on each desktop run setup with the
same origin and `DOCHAUS_WORD_EXTERNAL_HTTPS=1`; this registers the manifest
without creating or trusting a local certificate. The server origin must remain
stable for existing linked files. This configuration supports remote hosting;
cross-machine Office acceptance still needs validation in the target deployment.

## Versions and access

Word uploads use the existing document revision writer. They retain tracked
changes and comments, attribute the author from Word's last-modified property
(or `unknown`), refresh source search content and appear as **Word** versions.
The parent records the version edited in Word. If doc.haus or another Word copy
changed the document meanwhile, both versions remain and the UI warns about the
overlap. A clean Word revision becomes the accepted working base for subsequent
drafting; unresolved tracked changes retain the existing review guard.

An active pane sends heartbeats. The matter warns before edits, acceptance or a
drafting run while Word has a document open. The badge expires if the pane stops
responding; it cannot prove a file is closed. Done in Word clears it promptly.

Each linked file contains a bearer capability that can add versions only to its
own document and report its editing status. Treat a linked copy as able to save
back to that document. Use **Disconnect Word** in doc.haus to revoke its links.
Deleting and recreating a same-named document does not restore old access.
Ordinary **Download** copies have the doc.haus link and token removed, while
preserving unrelated add-ins. Tokens cannot read or delete documents or call
the normal API. The receiver accepts only its configured pane origin.

## Implementation note

Office exposes saved state but no reliable every-save/close notification. This
feature therefore confirms explicit uploads and supplements them with polling,
rather than promising one version for every native save. It adds a static pane,
restricted HTTPS receiver, existing-version-writer integration and web controls.
Focused regressions cover serialized saves, reconnect, malformed packages,
revocation, document identity, overlap parents, source indexing, clean downloads
and live-view refresh. Native Mac and Windows checks use isolated installations
and synthetic documents. Firm deployment still needs acceptance in the target
environment.
