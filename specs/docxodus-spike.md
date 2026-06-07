# Spike #1 — Docxodus under Bun + offset alignment

**Decision: GO.** Docxodus is validated as the DOCX engine for the redline track
(#2 word-integration, #3 tracked-changes, #4 redline, #5 viewer). Replaces the
FUTURE.md §Retrieval/Custom-tools validation gate.

Backed by scratch scripts in `specs/docxodus-spike/` (`make-sample.mjs`,
`spike.mjs`, `spike2.mjs`), run under Bun 1.3.14 against `docxodus@6.4.0`.

## What was tested

`docxodus@6.4.0` ships a .NET 8 (`dotnet.js`) OOXML engine compiled to WASM
(`DocumentFormat.OpenXml`), self-contained (no npm deps), ~17.7 MB unpacked.

| Question | Answer |
|----------|--------|
| Loads **headless under Bun** (not just browser)? | **Yes.** `initialize(wasmBasePath)` succeeds; cold start **154–223 ms**; `getVersion() → { dotnetVersion: "8.0.27", platform: "browser-wasm" }`. The `MONO_WASM: Error loading symbol file` line is a non-fatal symbols warning. |
| Native `w:ins`/`w:del` that open in Word? | **Yes.** A session opened with `{ trackedChanges, revisionAuthor }` emits valid `w:ins`/`w:del` on `replaceText` (`edited.docx`: `w:ins=1 w:del=1`). `compareDocuments(original, modified)` produces the same independently. |
| Edit targetable to a precise location → tracked change? | **Yes**, via **anchors**, not global offsets. `findByText` / `findAllByText` / `grep` return a stable block anchor id (e.g. `p:body:<unid>`); `replaceTextRange(anchorId, find, replace)` / `replaceText` then mutate that block. |
| Do Docxodus offsets align with our citation offsets? | **No, and we don't need them to.** Our citations index `mammoth.extractRawText()` plaintext (UTF-16, `\n\n` between paragraphs); Docxodus offsets index its own markdown projection. On the sample they differ by 4 (`searchTextOffsets` start 224 vs mammoth 228). **Bridge by re-anchoring on the cited text** (we already store each chunk's `text`), not by mapping offsets. Zero mapping cost. |

## Architectural consequences for #2–#5

- **Address edits by text, not offsets.** Take a citation's stored `text`, call
  `findByText`/`findAllByText`/`grep` to get the anchor, then edit. No
  offset-unit translation between mammoth and Docxodus.
- **Plain vs tracked is one setting.** Session `{}` → plain edit (`plain.docx`:
  `w:ins=0 w:del=0`). Session `{ trackedChanges, revisionAuthor }` → tracked.
  #2 uses plain; #3 uses tracked.
- **Two redline producers:** in-session tracked edits, or `compareDocuments` of
  two whole docs. #4 can use either.
- **WASM hosting:** the runtime lives at `docxodus/dist/wasm/`; resolve it from
  `node_modules` and pass to `initialize()` — no separate hosting needed
  server-side.

## Fallback (unused)

Not needed. Had it been no-go: `docx-redline-js` (MIT, text-match targeting
only; not on npm). SuperDoc rejected — AGPLv3, incompatible with distribution.
