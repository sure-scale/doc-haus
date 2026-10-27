# Open in Word

Use **Open in Word** on a Word document in a matter, then open the downloaded
copy in desktop Microsoft Word. The doc.haus pane identifies its matter and
document. Choose **Save to doc.haus** and wait for confirmation before closing.
**Done in Word** saves, confirms the final document, clears its open badge and
closes that document. If Word cannot close it, the pane keeps the saved
confirmation visible so you can close it manually.

Opening a linked copy checks its version without changing doc.haus or taking an
editing lock. The first **Save to doc.haus** claims editing if the copy is current.
The pane then checks saved content in the background. This is best effort:
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
POST/OPTIONS `/word/lease`, `/word/saves`, `/word/heartbeat`, `/word/done` and
`/word/desk/sources`, `/word/desk/source`, `/word/desk/versions`, `/word/desk/compare` to ingest. Do not
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
The parent records the version edited in Word. A stale copy cannot overwrite a
newer version. A clean Word revision becomes the accepted working base for subsequent
drafting; unresolved tracked changes retain the existing review guard.

The pane shows this copy's base version and the current version in doc.haus. It
checks these before saving and after reconnecting. If Word reloads its pane, it
can recover a confirmation only when the saved content matches exactly; it never
assumes an older or forwarded copy is the newest version. Done updates the link
in this physical copy before closing. If a copy is stale, keep its local file,
return to the matter and compare its versions before opening a fresh linked copy.
Keep any unsent local edits separately before replacing your Word copy.

One Word editor can hold the document's editing lock. While held, doc.haus
refuses replacement, acceptance, rejection, restore, redaction and other writes
to that document; viewing, comparing and downloading remain available. The pane
renews its lock while connected. Done releases it; an abandoned lock expires
after two minutes. **Release Word lock** and **Take over editing** in doc.haus
compare the displayed version and owner before acting. Taking over disconnects
earlier linked copies, whose local files remain intact. Expired or replaced
owners must check their version and explicitly claim editing again.

The lock protects writes to doc.haus. Word still allows local editing and native
Save; the add-in cannot promise to disable those commands. Its Save and Done
actions pause when the copy is stale or another editor owns the document.
The separate open badge is a contact hint and cannot prove a file is closed.

Each linked file contains a bearer capability that can add versions only to its
own document, report its editing status, and read confirmed matter sources and
outputs through Source Desk. Saved-version comparison is limited to the linked
document. Treat a linked copy as able to save back to that document and read
those matter previews. Use **Disconnect Word** in doc.haus to revoke its links.
Deleting and recreating a same-named document does not restore old access.
Ordinary **Download** copies have the doc.haus link and token removed, while
preserving unrelated add-ins. Tokens cannot delete documents or call the normal
API. Desk reads use a separate bounded route set; the receiver accepts only its
configured pane origin.

## Source Desk

Choose **Open Source Desk** in the Word pane to keep a captured Word passage
beside saved matter sources. The movable modeless window uses the pane for
document access. Keep the pane open while using it. Hosts without secure dialog
messaging open the same workspace inside the pane. **Use Source Desk in this pane**
also closes a floating desk and opens a fresh workspace below the save controls.
Opening, refreshing or
closing the desk does not claim editing ownership or change the save queue.

Choose **Open a source** to read it immediately. **Keep open** adds a source
to the strip for switching between useful documents; up to eight can be kept.
The filename and saved version stay visible, with version IDs and content hashes
in **Source details**. Open previews stay unchanged when you select another
Word passage and choose **Use current selection**. Sources with
unconfirmed external changes are unavailable until doc.haus records them; a
changed source is never silently substituted. Previews and source lists have
visible size limits.

Word text is labelled as captured saved or unsaved context, separate from a
confirmed matter version. **Show in Word** selects an exact, unique
passage only if the document's body text still matches the capture. Repeated passages,
multiline selections, Word find codes and selections over 255 characters cannot
be navigated in this first slice; select a short unique passage and refresh.
Body text changes invalidate navigation; formatting and markup are not checked. Refreshes
never select text or move focus in Word. **Find** (Cmd/Ctrl+F inside the desk)
highlights exact text in the selection and source; it does not assess legal
support. Escape closes Find and returns focus to its button.

**Compare versions** uses the existing comparison engine on two explicitly
chosen retained versions of the linked document. Choose an **Earlier version**
and a **Later version**, then **Compare**. Their applied version numbers stay
visible as Previous/Next moves through changed passages. The existing doc.haus
redline marks insertions blue and underlined, deletions red and struck through,
and moves green. Separate earlier/later text and technical details are available
in expandable sections. The open
Word document is never substituted for either saved version. This view does
not invoke native Word Compare or promise annotation-author fidelity.
The version list is a snapshot: **latest at refresh** refers to the last list
refresh. Choose **Refresh versions** after saving to load newer entries.

## Implementation note

Office exposes saved state but no reliable every-save/close notification. This
feature therefore confirms explicit uploads and supplements them with polling,
rather than promising one version for every native save. It adds a static pane,
restricted HTTPS receiver, existing-version-writer integration and web controls.
Focused regressions cover serialized saves, reconnect, malformed packages,
revocation, document identity, stale versions, ownership recovery, source indexing, clean downloads
and live-view refresh. Native Mac and Windows checks use isolated installations
and synthetic documents. Firm deployment still needs acceptance in the target
environment.
