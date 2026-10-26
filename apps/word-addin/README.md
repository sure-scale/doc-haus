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

**Save to doc.haus** saves in Word, exports the DOCX in 1 MiB slices, closes the
Office file handle, and waits for an acknowledged version. **Done in Word** does
the same before ending the live heartbeat and closing the current document.
Closing requires WordApi 1.5; the pane still supports WordApi 1.1 and provides a
manual-close confirmation on older hosts or when closing fails. It always uses
Word's `Save` close behavior. New changes detected before closing or after a
close error restore the heartbeat and need another save. See Microsoft's
[Document.close documentation](https://learn.microsoft.com/en-us/javascript/api/word/word.document?view=word-js-1.5#word-word-document-close-member(1)).
The final fingerprint check, settings persistence and native close are not
atomic. A later edit is preserved locally by `Save` and can be recovered by
reopening the file.
Save, export, retry, background sync,
and Done run in a single queue. Failed uploads keep that saved snapshot in memory
and retry the same bytes/base, at most twice per operation, before allowing the
user to retry. No document bytes are placed in browser storage.

While the pane stays open, a five-second poll checks `document.saved` and hashes
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
starts from its embedded base; the server can disclose overlapping changes
instead of silently treating it as the latest copy. Storage may be unavailable
or cleared; the saved file is still recoverable on reopen.
A changed comparison algorithm can invalidate an older stored fingerprint. One
new acknowledged save establishes the current checkpoint; an unmatched old
checkpoint still uses the embedded base and preserves overlap reporting.

A save is confirmed only after a fresh saved-file fingerprint matches the
acknowledged upload. Done checks again after its server response and restores
the heartbeat if editing continued. The pane keeps a light saved-state watch
after Done, with periodic content checks, so later edits resume the connection;
Save also remains available. Network deadlines cover the complete save response,
including its JSON body.
