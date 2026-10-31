# Open in Word

Use **Open in Word** on a Word document in a matter. Allow the browser to open
doc.haus if asked; the installed desktop helper opens a linked, editable working
copy in Microsoft Word. The doc.haus pane identifies its matter and
document. Choose **Save** and wait for confirmation before closing in Word.

Opening a linked copy checks its version without changing doc.haus or taking an
editing lock. The first **Save** claims editing if the copy is current.
The pane then checks saved content in the background. This is best effort:
rapid saves may become one version, and closing Word can interrupt an upload.
Reopening a linked copy can recover its last saved state. Native Word Save alone
does not confirm that doc.haus received it. When offline, keep the file and use
**Save** after reconnecting.

## Install on a desktop

Onboarding's final Details step selects **Set up Microsoft Word** when desktop
Word is installed. **Save and finish** installs it and asks you to approve the
local connection. Clear that choice to set it up later.

Existing users can open **Settings → Add-ins → Microsoft Word**. Choose
**Set up Word** or **Repair Word connection**. The status checks this
installation's active connection, certificate trust and installed manifest.
Setup takes effect without restarting doc.haus; restart Word if instructed.
Keep doc.haus running while using the add-in. An ordinary Word document is not
linked automatically: add it to a matter, then use **Open in Word**.

The desktop launcher registers `doc.haus://launch` and the narrowly scoped
Word-open handoff for this installation. Startup and Word setup refresh an
already installed launcher. The helper fetches the linked DOCX over verified
HTTPS using a short-lived, one-use ticket, then stores it in this desktop's
private runtime `word/checkout` directory. Repeated opens of the same linked
version reuse its file without overwriting local edits. A new version or link
may create a separate working copy; old copies are kept for recovery. The cache
is deliberately persistent rather than `/tmp`, so an interrupted session does
not discard unsent work. It is not the canonical matter file.

If the browser blocks the handoff, Word is unavailable or the desktop helper is
not installed, the app offers **Download Word copy**. Open that file manually.
Launch confirmation means the operating system accepted the open command;
**Saved** in the Word pane confirms that doc.haus received the document.

The add-in's initial page is hosted at `https://doc.haus/word-addin/bootstrap`.
It contains no document content or save capability, and checks this instance's
limited health endpoint before loading its pane. If a local instance is stopped,
it offers **Open doc.haus** and checks automatically every second until the
instance is ready. A remote instance instead asks the user to check the server.
This recovery page requires network access; it does
not promise first-load availability without an internet connection.

For command-line installations, run this from the doc.haus folder, with the
same data-root settings as startup:

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
The private key stays with the installation. Use **Repair Word connection**
to renew a certificate or select this installation's add-in again. The CLI
route requires a doc.haus restart. Do not copy runtime keys into Git.

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

Terminate TLS at your existing reverse proxy. Forward `/word-addin/*`, exact
GET `/word/open/<ticket>` and POST `/word/open/<ticket>/result`, and only
POST/OPTIONS `/word/lease`, `/word/saves`, `/word/heartbeat`, `/word/done` and
`/word/desk/sources`, `/word/desk/source`, `/word/desk/versions`, `/word/desk/compare`,
`/word/desk/open`, and the exact `/word/chat/` commands `conversations`,
`conversation`, `send`, `poll`, `cancel`, `materials`, `upload`, `suggest` and
`prepare` to ingest. Do not
expose the rest of the ingest API through this Word receiver. In Docker, build
the pane with `bun run --cwd apps/word-addin build` and include its `dist` assets
with the server. The normal start scripts build these assets automatically.

Download the manifest from `https://word.example.org/word-addin/manifest.xml`
and deploy it to Word users. Alternatively, on each desktop run setup with the
same origin and `DOCHAUS_WORD_EXTERNAL_HTTPS=1`; this registers the manifest
without creating or trusting a local certificate. The server origin must remain
stable for existing linked files. This configuration supports remote hosting;
cross-machine Office acceptance still needs validation in the target deployment.
For automatic opening, each desktop also needs the doc.haus desktop helper;
the remote server cannot launch Word on another computer. Its URI carries the
configured HTTPS origin, so it does not assume the server is localhost. Desktops
without the helper can use the download fallback.

## Versions and access

Word uploads use the existing document revision writer. They retain tracked
changes and comments, attribute the author from Word's last-modified property
(or `unknown`), refresh source search content and appear as **Word** versions.
The parent records the version edited in Word. A stale copy cannot overwrite a
newer version. A clean Word revision becomes the accepted working base for subsequent
drafting; unresolved tracked changes retain the existing review guard.

The compact status shows the current version in doc.haus; different copy and
matter versions appear only when they disagree. **Ready** describes a connected
copy without claiming its content is saved. **Saved** requires confirmed matching
content. The pane checks versions before saving and after reconnecting. If Word reloads its pane, it
can recover a confirmation only when the saved content matches exactly; it never
assumes an older or forwarded copy is the newest version. If a copy is stale, keep its local file,
return to the matter and compare its versions before opening a fresh linked copy.
Keep any unsent local edits separately before replacing your Word copy.

In a current desktop Word, **Load latest version** replaces the open document
after confirmation. Use Save As first to keep local work. Loading does not
publish a version; choose **Save** when ready. **Open a separate copy** remains
available for older Word versions or documents too large to refresh in place.

One Word editor can hold the document's editing lock. While held, doc.haus
refuses replacement, acceptance, rejection, restore, redaction and other writes
to that document; viewing, comparing and downloading remain available. The pane
renews its lock while connected; an abandoned lock expires
after two minutes. **Release Word lock** and **Take over editing** in doc.haus
compare the displayed version and owner before acting. Taking over disconnects
earlier linked copies, whose local files remain intact. Expired or replaced
owners must check their version and explicitly claim editing again.

The lock protects writes to doc.haus. Word still allows local editing and native
Save; the add-in cannot promise to disable those commands. Its **Save** action
pauses when the copy is stale or another editor owns the document.
The separate open badge is a contact hint and cannot prove a file is closed.

Each linked file contains a bearer capability that can add versions only to its
own document, report its editing status, and read confirmed documents and outputs already saved
in the matter through the document viewer. Saved-version comparison is limited to the linked
document. Treat a linked copy as able to save back to that document and read
those matter previews. Taking over editing revokes earlier copies' save links.
Deleting and recreating a same-named document does not restore old access.
Ordinary **Download** copies have the doc.haus link and token removed, while
preserving unrelated add-ins. Tokens cannot delete documents or call the normal
API. Chat commands use the linked matter and document and never proxy arbitrary
engine sessions or tools. The receiver accepts only its configured pane origin.

## Chat beside the document

Ask a question or request a change in **Chat**. A selection in Word automatically
adds its exact passage; remove its chip to ask without it. **Attach**
adds a saved matter document, a playbook or a supported upload. Uploading for
the chat leaves the matter's files alone; **Add to this matter** uses normal
source ingestion. Referencing a playbook does not change the matter's default.

Default chat settings are in doc.haus **Settings → Word add-in**.
The composer’s **This document** toggle uses the current Word text
plus deliberately attached material. **Whole matter** can also retrieve saved
matter sources and its playbook. Earlier broader answers remain visible but
cannot supply evidence after switching to the restricted scope.

Suggestions are on by default, with three clickable prompts in an empty chat.
Preparation uses sampled document passages together with any current selection.
Ideas never apply themselves. Pause stops new preparation; ordinary chat remains
available. Suggestions are not a complete document review.

A proposed amendment shows its before/after wording, formatting or table-row
change and any open items. **Refine** keeps the original target. **Apply** uses
Word-native tracked edits for supported operations, including adding a row,
deleting an eligible existing row, or joining two compatible ordinary paragraphs
with one space. A join preserves the wording and tracks the changed boundary.
A paragraph split can replace an existing space with a paragraph break while
retaining the source character formatting and paragraph properties, including
qualified paragraphs inside a table cell. The new boundary is tracked too.
A changed or unsupported passage requires a
fresh review. Used action buttons disappear after application.
Apply leaves the matter's version history alone: choose **Save** to synchronize
Word's resulting document and wait for the saved-version acknowledgement.
Reviewed application requires current Microsoft 365 Word with WordApi 1.6;
unsupported hosts retain chat and the existing save flow.

## Matter documents and version comparison

The **Chat**, **Documents** and **Compare** tabs switch the pane’s main view.
Chat opens on first use; the pane remembers the last tab deliberately selected
on this installation. Switching tabs keeps the chat draft and review in place.
Documents lists files already saved in this matter, such as
client instructions, heads of terms or correspondence. Choose a file to read it;
no Word selection is needed or captured automatically. The ↗ button opens the
current view in a wider movable window. Keep the Word pane open for that window's
connection. If a wider window is unavailable, continue inside the pane.
Opening, refreshing or closing these views does not claim editing ownership or
change the save queue. The ribbon's **doc.haus** command opens the pane.

Each document opens in a tab; **+ Open document** opens another, and × closes a
tab. Up to eight exact saved snapshots can stay open. The reader shows its saved
version, date and read-only status. Real DOCX headings are displayed where they
can be matched safely to the extracted text; other documents use plain text.
**Details** contains version IDs and content hashes. **Refresh document list**
updates the choices without replacing open snapshots. Documents with unconfirmed
external changes are unavailable until doc.haus records them; a changed file is
never silently substituted. Previews and document lists have visible size limits.

Optionally select a passage in Word and turn on **Word selection**. The passage
appears beside the saved document, with its capture time. ↻ captures the selection
again; turning the toggle off returns to the full-width reader. These actions
preserve open document tabs. Captured text is separate from confirmed matter
versions; Word's native saved state does not prove doc.haus has received it.

**Show in Word** selects an exact, unique passage only if the document's body
text still matches the capture. Repeated passages, multiline selections, Word
find codes and selections over 255 characters cannot be navigated; select a
short unique passage and refresh the selection. Body text changes invalidate
navigation; formatting and markup are not checked. Capturing or refreshing does
not select text or move focus in Word. **Find** (Cmd/Ctrl+F inside the reader)
highlights exact text in the open document and, when shown, the captured Word
text; it does not assess legal support. Use the arrows to step through matches.
Escape closes Find and returns focus to its button.

**Compare** opens two saved versions of this linked document. Choose **Earlier**
and **Later** versions. Known pairs totalling at most 2 MiB compare automatically,
starting with the latest saved version and its predecessor; larger or unknown-size
pairs show a **Compare** button. One comparison runs at a time. Its applied
version numbers stay visible as ‹ and › move through changed passages. The
redline marks insertions blue and underlined, deletions red and struck through,
and moves green. Untracked edits are labelled. **Details** includes separate
earlier/later text and technical metadata.

The open Word document is never substituted for either saved version. Comparison
navigation stays within the saved-version redline; **Show in Word** belongs to
the separately captured selection. This view does not invoke native Word Compare
or promise annotation-author fidelity. **Refresh version list** loads newer
choices without replacing the applied comparison. “Latest known” refers to that
list's last refresh; choose the newer version to update the comparison.

The ⋯ save menu contains **Check status**, the alternative save action and
**Stop editing here** when applicable. Checking status never saves. **Save**
claims editing when available and confirms a version. **Open matter in doc.haus**
returns to the matter.

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
