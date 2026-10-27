import { tool, type LegalToolContext } from "./legacy-definition"
import { existsSync, lstatSync, realpathSync } from "node:fs"
import path from "node:path"
import { createHash } from "node:crypto"
import { docxodus } from "../lib/docxodus"
import { listRedactions, recordRedaction, scrubDocxPackage, docxPackageResidue, type RedactionRow } from "../lib/redactions"
import { supersedeMatchingRedlines } from "../lib/redlines"
import { assertWordEditAllowed } from "../lib/word-edit-lease"
import { canonicalArtifactIdentity, withArtifactMutationLocksAsync } from "../capabilities/artifacts/mutation-lock"
import { combinedParts, readLeanDocuments } from "../workflows/lean/deliverables"
import { readLeanState } from "../workflows/lean/state"

// doc.haus redact tool. TRUE redaction: removes the sensitive text from the
// .docx itself — every body/header/footer/footnote/endnote/comment occurrence
// via Docxodus, plus what Docxodus cannot address (docProps, attribute values
// such as alt text, link targets and document variables, other XML parts) via a
// package walk — and re-scans every readable value of the package before
// writing, naming the parts it cannot read. This is the opposite of an overlay
// "redaction" (a black box drawn over intact text), the recurring data-breach
// vector this tool exists to prevent.
//
// Removal is destructive and irreversible, so it is all-or-nothing: any
// occurrence the engine cannot remove (text spanning a paragraph boundary,
// residue surviving the scrub, such as a value split across runs outside the
// Docxodus projection or a copy inside a binary part) aborts the whole
// operation before the file is touched, and the user approves via permission
// "redact" before anything is removed. Every redaction is recorded in the
// matter's redaction log after the existing revision writer confirms the save
// and re-index. A later log failure is reported without hiding that save.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Permanently remove sensitive text from the matter's Word (.docx) document — true redaction, not an overlay. Finds every occurrence of the exact text across body, headers, footers, footnotes, endnotes, and comments, replaces each with a redaction label, scrubs the text from document metadata and attribute values (title, author, comment authors, image alt text, link targets, document variables), re-scans every readable part of the file and names any binary parts it cannot read, and records the redaction in the matter's redaction log. The user approves before anything is removed. THIS IS IRREVERSIBLE — the text is gone from the file. Use search-document or read-document first to confirm the exact text, and give a reason a privilege or redaction log can cite.",
  args: {
    document: tool.schema.string().describe("Document file name within the matter (the docPath from a citation)"),
    text: tool.schema.string().describe("The exact sensitive text to remove everywhere it appears (case-sensitive)"),
    reason: tool.schema
      .string()
      .describe('Why this is being redacted, for the redaction log — e.g. "SSN — PII", "attorney-client privileged"'),
    author: tool.schema.string().optional().describe("Who is performing the redaction, for the redaction log (default: doc.haus)"),
    label: tool.schema
      .string()
      .optional()
      .describe('Replacement marker written where the text was (default: "[REDACTED]")'),
  },
  async execute(args, ctx) {
    const file = path.isAbsolute(args.document) ? args.document : path.join(ctx.directory, args.document)
    if (!matterDocument(ctx.directory, file)) return `Document not found in this matter: ${args.document}`
    const label = args.label ?? "[REDACTED]"
    if (label.includes(args.text)) return "The redaction label must not contain the text being redacted."

    const dx = await docxodus()
    const bytes = await Bun.file(file).bytes()

    // Tracked changes store deleted text in w:del elements that survive every
    // Docxodus save, so a document with pending revisions cannot be truly
    // redacted — the "removed" text would still ship inside the file. Refuse and
    // have the user resolve revisions first.
    const docMeta = await dx.getDocumentMetadata(bytes)
    if (docMeta.hasTrackedChanges)
      return `${path.basename(file)} has unresolved tracked changes. Revision history retains inserted and deleted text inside the file, so redaction cannot guarantee removal. Accept or reject all revisions (in Word, or via the doc.haus redline review) and try again.`

    const session = dx.openDocxSession(bytes, {})
    const pattern = args.text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    const matches = session.grep(pattern, { scope: dx.ProjectionScopes.All })

    // A match spanning a paragraph boundary cannot be replaced by the span-
    // addressed engine, and a partial redaction must never look like a complete
    // one — abort before anything is asked or touched.
    const crossBlock = session.grepCrossBlock(pattern, { scope: dx.ProjectionScopes.All }).filter((m) => m.slices.length > 1)
    if (crossBlock.length) {
      session.close()
      return `Cannot redact: ${crossBlock.length} occurrence(s) of the text span a paragraph boundary, which this tool cannot remove in one operation. Redact the text in per-paragraph pieces instead.`
    }

    // The needle may also live only in package metadata (a creator name, a
    // keyword) with no visible occurrence — surface those parts so the user
    // approves the real scope of the removal. Parts grep already covers are
    // excluded; what remains is docProps and other unprojected parts.
    const allParts = [...new Set(docxPackageResidue(bytes, args.text).hits.map((hit) => hit.part))]
    const metadataParts = allParts.filter(
      (p) => !/^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/.test(p),
    )
    if (!matches.length && !allParts.length) {
      session.close()
      // This redaction was made, but its save to doc.haus failed: the text is
      // gone from the file and logged, so running it again sends that save.
      const logged = attempt(() => listRedactions(ctx.directory).find((row) => row.doc_name === path.basename(file) && row.redacted_text === args.text))
      if (logged && !(logged instanceof Error)) return await saveAgain(file, args.text, logged, ctx)
      return `Text not found in ${path.basename(file)} (content or metadata): ${JSON.stringify(args.text)}`
    }
    session.close()

    const author = args.author ?? "doc.haus"
    // Removal is irreversible, so it is gated on the matter owner's approval
    // (permission "redact" in opencode.json). ctx.ask blocks until they reply and
    // throws on reject — asked before any mutation so a rejection changes nothing.
    // "Always" approves future redactions of this document only.
    await ctx.ask({
      permission: "redact",
      patterns: [file],
      always: [file],
      metadata: {
        document: path.basename(file),
        text: args.text,
        occurrences: matches.length,
        metadataParts,
        reason: args.reason,
        label,
        author,
      },
    })

    const combined = redactionCombined(ctx.directory, file)
    const locked = [file, ...(combined ? [combined.file, combined.body, combined.notes] : []), ".dochaus/source-corpus"]
    const redaction = await withArtifactMutationLocksAsync(ctx.directory, locked, async () => {
      assertWordEditAllowed(ctx.directory, file)
      const currentCombined = redactionCombined(ctx.directory, file)
      if (currentCombined && currentCombined.file !== combined?.file)
        throw new Error("The combined document changed before redaction. Check the current documents and try again.")
      if (currentCombined) assertWordEditAllowed(ctx.directory, currentCombined.file)
      if (!existsSync(file)) return { error: `Document not found in this matter: ${args.document}` }
      const currentBytes = await Bun.file(file).bytes()
      const currentMeta = await dx.getDocumentMetadata(currentBytes)
      if (currentMeta.hasTrackedChanges)
        return { error: `${path.basename(file)} gained unresolved tracked changes before approval. Resolve them and try again.` }

      const currentSession = dx.openDocxSession(currentBytes, {})
      const currentMatches = currentSession.grep(pattern, { scope: dx.ProjectionScopes.All })
      const currentCrossBlock = currentSession
        .grepCrossBlock(pattern, { scope: dx.ProjectionScopes.All })
        .filter((match) => match.slices.length > 1)
      if (currentCrossBlock.length) {
        currentSession.close()
        return { error: `Redaction aborted, document unchanged: ${currentCrossBlock.length} occurrence(s) now span a paragraph boundary.` }
      }
      const currentParts = docxPackageResidue(currentBytes, args.text).hits
      if (!currentMatches.length && !currentParts.length) {
        currentSession.close()
        return { error: `Text is no longer present in ${path.basename(file)} after approval.` }
      }

      // Replace span-addressed matches in reverse span order within each anchor
      // so earlier spans stay valid as later ones shrink the text.
      const ordered = [...currentMatches].sort((a, b) =>
        a.enclosingAnchor.id === b.enclosingAnchor.id ? b.span.start - a.span.start : 0,
      )
      for (const match of ordered) {
        const result = currentSession.replaceMatch(match, label)
        if (!result.success) {
          currentSession.close()
          return { error: `Redaction aborted, document unchanged: replacing an occurrence in ${path.basename(file)} failed (${result.error?.message ?? "unknown error"}).` }
        }
      }
      const replaced = currentSession.save()
      currentSession.close()
      const scrub = scrubDocxPackage(replaced, args.text, label)
      const residue = docxPackageResidue(scrub.bytes, args.text)
      if (residue.hits.length)
        return {
          error: `Redaction aborted, document unchanged: the text still appears after removal where this tool cannot safely remove it: ${residue.hits.map((hit) => `${hit.part} (${hit.where})`).join("; ")}. The document needs manual review.`,
        }

      return { scrub, residue, occurrences: currentMatches.length, priorSha256: sha256(currentBytes) }
    })
    if ("error" in redaction) return redaction.error

    // Publication owns its locks and checks this exact prepared base again.
    // Do not hold a tool-process lock while calling the revision service.
    const saved = await saveRedaction(file, redaction.scrub.bytes, redaction.priorSha256, ctx)
    if (!saved.confirmed) return {
      title: `Redaction not confirmed: ${path.basename(file)}`,
      output: `The redaction save was not confirmed. ${saved.error} No redaction log entry was recorded and no pending redlines were retired. Reload the current document in doc.haus before trying again; an interrupted response may have followed a successful save.`,
      metadata: { document: file, redaction: null, reingested: false, supersededRedlines: null },
    }
    const { id, stale } = await withArtifactMutationLocksAsync(ctx.directory, [file], async () => ({
      stale: attempt(() => supersedeMatchingRedlines(ctx.directory, file,
        (row) => row.find_text.includes(args.text) || row.new_text.includes(args.text))),
      id: attempt(() => recordRedaction(ctx.directory, {
        docPath: file, docName: path.basename(file), redactedText: args.text, label,
        reason: args.reason, author, occurrences: redaction.occurrences,
        metadataHits: Object.values(redaction.scrub.hits).reduce((a, b) => a + b, 0),
      })),
    }))
    const warnings = [
      id instanceof Error &&
        `WARNING: the redaction log entry was not recorded (${id.message}). The matter's redaction log has no entry for this removal; record it manually.`,
      stale instanceof Error &&
        `WARNING: pending redlines that reference the removed text were not retired (${stale.message}). Reject any that still propose or anchor to the removed text in the redline review.`,
      ...saveWarnings(saved),
      redaction.residue.unread.length &&
        `NOTE: ${redaction.residue.unread.length} part(s) are not readable text (${redaction.residue.unread.join(", ")}). The scan searched them only for a literal copy of the text and found none; it cannot see compressed, encoded or image content, so review them manually if they could contain or depict it.`,
      !(stale instanceof Error) && stale.length && `Retired ${stale.length} pending redline(s) that referenced the removed text.`,
    ].filter(Boolean)

    return {
      title: `Redacted ${path.basename(file)}`,
      output: [
        `Permanently removed ${JSON.stringify(args.text)} from ${path.basename(file)}: ${redaction.occurrences} text occurrence(s) replaced with ${JSON.stringify(label)}, ${Object.keys(redaction.scrub.hits).length} package part(s) scrubbed. Verified zero residue in the text the scan reads: every text node, attribute value and run-joined paragraph in ${redaction.residue.read.length} XML part(s)${redaction.residue.unread.length ? `; ${redaction.residue.unread.length} part(s) not read (see note)` : ""}.${id instanceof Error ? "" : ` Recorded as redaction log entry #${id} (reason: ${args.reason}, by ${author}).`}`,
        ...warnings,
      ].join("\n"),
      metadata: {
        document: file,
        occurrences: redaction.occurrences,
        scrubbedParts: redaction.scrub.hits,
        unreadParts: redaction.residue.unread,
        redaction: id instanceof Error ? null : id,
        reason: args.reason,
        author,
        reingested: saved.confirmed,
        supersededRedlines: stale instanceof Error ? null : stale.map((row) => row.id),
      },
    }
  },
})

// A log write's result, or the error it threw.
function attempt<T>(write: () => T): T | Error {
  try {
    return write()
  } catch (error) {
    return error instanceof Error ? error : new Error(String(error))
  }
}

// Ingest registers documents by name in the matter folder itself: only a file
// saved directly there, by its real path and not through a link, is a matter
// document. doc.haus's own records (.dochaus/: kept versions, the registry)
// never are, under any spelling of their path.
function matterDocument(directory: string, file: string) {
  try {
    return !lstatSync(file).isSymbolicLink() && path.dirname(realpathSync(file)) === realpathSync(directory)
  } catch {
    return false
  }
}

function redactionCombined(directory: string, document: string) {
  const summary = readLeanDocuments(directory)
  if (readLeanState(directory)?.status !== "completed" || summary?.status !== "completed") return
  const combined = combinedParts(directory, summary.documents)
  if (!combined) return
  const target = canonicalArtifactIdentity(directory, document)
  return [combined.body, combined.notes].some(name => canonicalArtifactIdentity(directory, name) === target) ? combined : undefined
}

// Only the existing role-preserving revision writer may publish a candidate.
// Both bounded sends carry the same prior hash: a lost acknowledgement cannot
// authorize overwriting a later version or produce an unconfirmed log entry.
function sha256(bytes: Uint8Array) { return createHash("sha256").update(bytes).digest("hex") }
type RedactionSave = { confirmed: true; purge_failed?: unknown } | { confirmed: false; error: string }
async function saveRedaction(file: string, bytes: Uint8Array<ArrayBuffer>, priorSha256: string, ctx: LegalToolContext): Promise<RedactionSave> {
  let error = "doc.haus could not be reached."
  for (let send = 0; send < 2; send++) {
    const response = await fetch(`${ingestUrl}/matters/${path.basename(ctx.directory)}/documents/revisions`, {
      method: "POST",
      headers: { "x-dochaus-origin": "redaction", "x-dochaus-actor": `agent:${ctx.agent}`, "x-dochaus-run": ctx.sessionID,
        "x-document-prior-sha256": priorSha256 },
      body: (() => {
        const form = new FormData()
        form.append("file", new File([bytes], path.basename(file)))
        return form
      })(),
    }).catch(() => null)
    if (response?.ok) return { ...(await response.json().catch(() => ({})) as { purge_failed?: unknown }), confirmed: true }
    if (response) {
      error = response.status === 423 ? "The document or its combined document is being edited in Word."
        : response.status === 409 ? "The document changed after the redaction was prepared."
        : `doc.haus refused the save (HTTP ${response.status}).`
      if (response.status < 500) break
    }
  }
  return { confirmed: false, error }
}

function saveWarnings(saved: RedactionSave) {
  if (!saved.confirmed) return ["WARNING: the redacted document's save was not confirmed. Reload its current copy before trying again; an interrupted response may have followed a successful save."]
  return typeof saved.purge_failed === "string" ? [`WARNING: ${saved.purge_failed}`] : []
}

// Running a logged redaction again, once the text is gone from the file: its
// save is sent again, on the matter owner's approval as the redaction was,
// since it deletes doc.haus's earlier versions of the document.
async function saveAgain(file: string, text: string, logged: RedactionRow, ctx: LegalToolContext) {
  await ctx.ask({
    permission: "redact",
    patterns: [file],
    always: [file],
    metadata: { document: path.basename(file), text, occurrences: 0, resave: true, metadataParts: [], reason: logged.reason, label: logged.label, author: logged.author },
  })
  const bytes = await Bun.file(file).bytes()
  const saved = await saveRedaction(file, bytes, sha256(bytes), ctx)
  return {
    title: `Redacted ${path.basename(file)}`,
    output: [
      `${JSON.stringify(text)} was already removed from ${path.basename(file)} (redaction log entry #${logged.id}).${saved.confirmed ? ` Saved the redacted document to doc.haus again${saved.confirmed && typeof saved.purge_failed === "string" ? "" : ", so its earlier versions, which held the text, are deleted"}.` : ""}`,
      ...saveWarnings(saved),
    ].join("\n"),
    metadata: { document: file, occurrences: 0, redaction: logged.id, reason: logged.reason, author: logged.author, reingested: saved.confirmed },
  }
}
