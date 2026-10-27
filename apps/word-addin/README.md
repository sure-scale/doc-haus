# doc.haus Word add-in

An additive static pane for desktop Word. The instance HTTPS bridge serves
`dist/` under `/word-addin/` and renders `/word-addin/manifest.xml` with
`renderManifest(configuredHttpsOrigin)` from `src/manifest.ts`.

Run `bun run typecheck`, `bun run test`, and `bun run build` in this package.
`WORD_ADDIN_ORIGIN=https://your-instance.example bun run build` also writes a
manifest for that origin. Use a certificate trusted by Office; HTTP loopback is
not sufficient. The pane uses Microsoft's hosted Office.js and the existing
doc.haus icon copied at build time. No document or credential data belongs in
this package.

The linked DOCX supplies `dochaus.link`. Its HTTPS receiver must match the pane
origin. The bearer capability is sent only to that receiver, with redirects and
ambient credentials disabled. It is never shown, logged, or added to a URL.

Matter documents and version comparison add `desk.html`/`desk.js`/`desk.css` to the static build. The pane
keeps the capability and brokers only validated read/selection requests from a
same-origin Office dialog with DialogApi 1.2 and DialogOrigin 1.1. Missing support
falls back to an inline document view. Explicit target origins, channel identity and
bounded messages are required; a reloaded child establishes a fresh channel.
Desk requests never call the save queue or editing lease. See
[`docs/word-addin.md`](../../docs/word-addin.md#matter-documents-and-version-comparison) for scope and limits.

The manifest, linked handout and pane share one add-in version. Both manifest
pane URLs include that version, and a missing or different URL version stops
startup before Office access. Token requests also carry the version; ingest
rejects old clients with HTTP 426 before recording contact or accepting bytes.
After updating this version, run Word setup again, restart Word and open a fresh
handout. File equality alone does not prove Office adopted its cached manifest.
This compatibility gate does not elect a writer or adopt another copy's latest
parent. The server checks versions and a recoverable editing lease separately.

Opening or reconstructing the pane only checks the current version and editing
status. It does not claim editing ownership, send a heartbeat or upload a file.
Choose **Save to doc.haus** or **Done in Word** to claim a current, available
document before saving. **Start saving here** can also claim it explicitly.
Save and Done are disabled for stale versions or another editor's lease. A new
pane cannot inherit another pane's lease from a copied link.
Background polling checks status without claiming. **Check status** also does
not claim merely because another editor became unavailable during its request.

**Take over saving here** is a deliberate operation for a current copy held by
another session. It sends the displayed current version and lease generation;
a concurrent change refuses the takeover. **Release editing** leaves Word and
its local changes open. Heartbeats renew only the existing owner. Expiry, a
takeover or release pauses saving until an explicit eligible claim. Word still
allows native local editing; this is an enforced doc.haus writing lease, not a
native Word edit lock.

Stale, removed or externally changed versions cannot save or finish. Keep the
local file with Save As, return to the matter in doc.haus, compare **Versions**
or use **Open in Word** for a fresh copy. The pane never silently replaces its
base with the server's latest ID. Matching local acknowledgement fingerprints
can recover an earlier confirmed parent, but a different forwarded copy keeps
its embedded base and must resolve the conflict.

**Save to doc.haus** saves in Word, exports the DOCX in 1 MiB slices, closes the
Office file handle, and waits for an acknowledged version. **Done in Word** does
the same before releasing editing ownership and closing the current document.
Closing requires WordApi 1.5; the pane still supports WordApi 1.1 and provides a
manual-close confirmation on older hosts or when closing fails. It always uses
Word's `Save` close behavior. New changes detected before closing or after a
close error keep the document open and need another save. After release,
explicitly start saving again before sending new changes. See Microsoft's
[Document.close documentation](<https://learn.microsoft.com/en-us/javascript/api/word/word.document?view=word-js-1.5#word-word-document-close-member(1)>).
The final fingerprint check, settings persistence and native close are not
atomic. A later edit is preserved locally by `Save` and can be recovered by
reopening the file.
Save, export, retry, background sync,
and Done run in a single queue. Failed uploads keep that saved snapshot in memory
and retry the same bytes/base/lease generation, at most twice per operation,
before allowing the
user to retry. Only a server acknowledgement advances the parent; a lost
response is retried under its original generation before checking a potentially
advanced server version. A fenced or expired retry is refused. If the server
proves that original base is still current and available, an explicit Save can
claim a new generation, then save/export the current Word document anew. A
successful deliberate takeover can replace an old pending generation only when
the same base remains current. Neither path replays old bytes under a new fence
or adopts a newer parent. No document
bytes are placed in browser storage.

Once editing is deliberately claimed, a five-second poll checks
`document.saved` and hashes
sorted uncompressed DOCX parts to ignore ZIP timestamp/order/compression churn.
Some desktop Word exports also rotate revision-session identifiers without an
edit. For comparison only, the fingerprint parses XML and removes the known
WordprocessingML revision-session attributes and the direct `rsids` settings
list. Actual text, formatting, comments and tracked changes remain significant;
foreign/unknown attributes remain significant too. Uploaded DOCX bytes are never
rewritten by this comparison step.
The standard application-properties `TotalTime` numeric leaf is also normalized
for comparison, since it records editing duration (minutes by default), which
can advance while Word is open without a content edit. The four cached counts
`Words`, `Characters`, `Paragraphs` and `CharactersWithSpaces` are normalized
only as unique direct standard numeric leaves without attributes. Word can
recount these on reopening without changing content. Other application
properties, including `Pages` and `Lines`, and custom XML remain significant.
Uploaded DOCX bytes remain unchanged. See Microsoft's
[TotalTime documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.extendedproperties.totaltime?view=openxml-3.0.1)
and [cached word count definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.extendedproperties.words?view=openxml-3.0.1).
Word can also add `proofState` after background spelling and grammar checking
finishes. Comparison ignores only the direct standard settings leaf with known
`spelling`/`grammar` completion flags. Unknown attributes, nested data and
language preferences remain significant. See Microsoft's
[ProofState documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.proofstate?view=openxml-3.0.1).
Proofing can also add standard empty `proofErr` anchors and split otherwise
identical text runs. Comparison removes only the four known spelling/grammar
anchor types without extension data, then joins adjacent simple text runs in
standard Word paragraphs. Runs must have no attributes and identical complete
properties from the supported character-format subset (bold, italic, explicit
fonts, language, size and colour). Unknown or revision properties, fields,
comments, bookmarks, hyperlinks and other boundaries prevent joining. Text is
kept exactly; unpreserved boundary whitespace, tabs and line breaks prevent
joining too. Theme fonts, right-to-left properties and other unsupported run
layouts remain significant and may still produce an extra version after Word
fragments them. This is comparison only; the original proofing annotations,
runs and bytes are uploaded unchanged.
See Microsoft's [proofing-anchor documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.prooferror?view=openxml-3.0.1),
[run structure](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-open-and-add-text-to-a-word-processing-document)
and [whitespace rules](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.texttype.space?view=openxml-3.0.1).
The standard core-properties `dcterms:modified` full timestamp is normalized for
comparison because Word advances it on a native save after reopening, even
without an edit. Only the direct, typed timestamp leaf is handled; creator,
last-modified author, creation/printing dates, extensions and custom XML remain
significant. The standard core `revision` counter is normalized only as a
unique direct numeric leaf without attributes; extended or nonnumeric values
remain significant. Uploaded metadata is unchanged. See Microsoft's
[package modified-time definition](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.packageproperties.modified?view=netframework-4.8.1)
and [package revision property](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.packageproperties).
On startup, an initial export can recover a matching acknowledgement even when
Word reports unsaved; it stays Not saved and is never uploaded automatically.
Unsaved edits are never automatically uploaded. This is best-effort sync: several native saves
may become one version, and closing Word can stop a final upload. There is no
native close-save guarantee. Use an explicit action and wait for Saved before
closing.

Acknowledgement metadata (version, fingerprint, receive time and overlap) is
stored locally under a SHA-256 key derived from the link identity, original base
and token. The token and document bytes are not stored there. Background acknowledgements
do not write Word settings, avoiding document dirtying and save loops. Only
Done, after the upload and Done responses are acknowledged and the current
content matches, updates the existing validated link's `baseVersion`, awaits
settings persistence, explicitly saves the Word document, and closes with
`Save` only after Word reports saved. This separate save persists the settings
change before close. The original link identity and
other fields are preserved. Other physical copies are not modified. On reopen,
stored metadata is used only if
the current exported fingerprint matches. An edited/forwarded copy otherwise
starts from its embedded base; the server blocks stale saves instead of silently
treating it as the latest copy. Storage may be unavailable
or cleared; the saved file is still recoverable on reopen.
A changed comparison algorithm can invalidate an older stored fingerprint. One
new acknowledged save establishes the current checkpoint; an unmatched old
checkpoint still uses the embedded base and preserves stale-copy protection.
Office can reconstruct the pane before Done. A new client whose initial content
differs from its cached acknowledgement retains the embedded base. This can
conservatively block saving within one physical copy; Done-only persistence
does not cover earlier pane reconstructions. Preserve the file and compare or
reopen the current version rather than bypassing this check.
Never replace this check with token-wide latest-version adoption, which could
hide a real forwarded-copy fork.

A save is confirmed only after a fresh saved-file fingerprint matches the
acknowledged upload. Done checks again after its server response and keeps the
document open if
editing continued. The pane keeps a light saved-state watch after Done, with
periodic content checks; later edits require an explicit editing claim before
saving again. Network deadlines cover the complete save response,
including its JSON body.
