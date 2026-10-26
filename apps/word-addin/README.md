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
the same before ending the live heartbeat. Save, export, retry, background sync,
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
can advance while Word is open without a content edit. Other application
properties and custom XML remain significant. Uploaded DOCX bytes remain
unchanged. See Microsoft's [TotalTime documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.extendedproperties.totaltime?view=openxml-3.0.1).
Word can also add `proofState` after background spelling and grammar checking
finishes. Comparison ignores only the direct standard settings leaf with known
`spelling`/`grammar` completion flags. Unknown attributes, nested data, proofing
annotations and language preferences remain significant. See Microsoft's
[ProofState documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.proofstate?view=openxml-3.0.1).
It also captures the initial saved file on reopen. Unsaved edits are never
automatically exported/uploaded. This is best-effort sync: several native saves
may become one version, and closing Word can stop a final upload. There is no
native close-save guarantee. Use an explicit action and wait for Saved before
closing.

Acknowledgement metadata (version, fingerprint, receive time and overlap) is
stored locally under a SHA-256 key derived from the link identity, original base
and token. The token and document bytes are not stored there. Writing back to
Word settings after each acknowledgement would dirty the document and risk save
loops, so the add-in does not do it. On reopen, stored metadata is used only if
the current exported fingerprint matches. An edited/forwarded copy otherwise
starts from its embedded base; the server can disclose overlapping changes
instead of silently treating it as the latest copy. Storage may be unavailable
or cleared; the saved file is still recoverable on reopen.

A save is confirmed only after a fresh saved-file fingerprint matches the
acknowledged upload. Done checks again after its server response and restores
the heartbeat if editing continued. The pane keeps a light saved-state watch
after Done, with periodic content checks, so later edits resume the connection;
Save also remains available. Network deadlines cover the complete save response,
including its JSON body.
