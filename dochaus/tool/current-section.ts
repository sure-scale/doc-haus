import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { buildChains, chainNames, resolveCurrentSection, verbatimSpan, type Chain, type SectionOpRow } from "../lib/amendment-chain"
import { sectionPassages } from "../lib/chunk-text"
import { provenanceFooter, sourceProvenance } from "../lib/source-provenance"

// doc.haus current-section resolver. For a section label, walks the matter's
// amendment chains newest-first and returns the verbatim operative wording of
// whichever amender last touches that section (the stored op span), the text
// that follows it in that amender located by position only, the code-point
// position of both, and the full lineage so the answer is auditable. It never
// presents an amender's own clause as the base section's current text: the
// resulting text is not recorded, and the output says so. Falls back honestly
// when no chain exists, no amender touches the section, or the base document
// was not matched.
//
// Differs from get-section, which returns one document's verbatim section text
// without walking the chain. Use current-section for "what does §X currently
// say?"; use get-section for "what does document Y say at §X?".
//
// Cap on returned text mirrors get-section so one oversized section cannot
// flood the context window. Read-only; all writes happen in the ingest service.

const TEXT_CAP = 20000
// An operative statement is one sentence; clip a quote that is not.
const QUOTE_CAP = 1000

type Chunk = { doc_path: string; doc_name: string; section: string; text: string; char_start: number }

// Section text by label. `exclude` drops documents by path — amending
// documents, whose own numbering is not the base document's.
function fallbackSectionText(db: Database, section: string, document: string | undefined, exclude = new Set<string>()) {
  const where = document ? "doc_name = ? AND section = ?" : "section = ?"
  const params = document ? [document, section] : [section]
  const rows = db
    .query(`SELECT doc_path, doc_name, section, text, char_start FROM chunks WHERE ${where} ORDER BY doc_name, char_start, id`)
    .all(...params) as Chunk[]
  return rows.filter((row) => !exclude.has(row.doc_path))
}

function quote(text: string) {
  return text.length > QUOTE_CAP
    ? `"${text.slice(0, QUOTE_CAP)}" [quote clipped at ${QUOTE_CAP} characters]`
    : `"${text}"`
}

function renderChunks(rows: Chunk[]) {
  return sectionPassages(rows, TEXT_CAP)
    .passages.map(
      (passage) =>
        `=== ${passage.doc_name} — § ${passage.section} (chars ${passage.char_start}–${passage.char_start + passage.text.length}) ===\n${passage.text}`,
    )
    .join("\n\n")
}

function conjugate(operation: "amend" | "restate" | "delete" | "add") {
  return operation + (operation.endsWith("e") ? "d" : "ed")
}

// `quotedAbove` is the op whose wording the answer already quotes; its lineage
// line refers back instead of repeating the quote.
function renderLineage(db: Database, chain: Chain, label: string, quotedAbove?: SectionOpRow) {
  const lines: string[] = []
  let seenNewest = false
  for (let i = chain.entries.length - 1; i >= 0; i--) {
    const entry = chain.entries[i]!
    const matching = entry.ops.filter((op) => op.section_label === label)
    if (matching.length === 0) continue
    const op = matching[matching.length - 1]!
    const superseded = seenNewest
    seenNewest = true
    // Quote the operative sentence at the displayed op offsets — not the
    // doc-level recital, which lives at different offsets. The chars shown
    // on this line (op.char_start..op.char_end) must match the quoted text.
    const opVerbatim = op === quotedAbove ? null : verbatimSpan(db, op.doc_path, op.char_start, op.char_end)
    const verbatim =
      op === quotedAbove
        ? "operative wording quoted above"
        : opVerbatim
          ? quote(opVerbatim)
          : "operative text not recoverable from the index at these offsets"
    const tag = superseded ? " (now superseded)" : ""
    lines.push(`- ${entry.rel.doc_name} §${label} (chars ${op.char_start}..${op.char_end}): ${conjugate(op.operation)}${tag} — ${verbatim}`)
  }
  if (chain.base.kind === "resolved") {
    lines.push(`- ${chain.base.doc.name} §${label}: base text${lines.length > 0 ? " (superseded)" : ""}`)
  }
  return `Lineage (newest first):\n${lines.join("\n")}`
}

export default tool({
  description:
    'Trace a contract section across the matter\'s amendment chain — e.g. "what does Section 4 currently say?". Walks amendments newest-first and returns the latest amendment\'s verbatim operative wording for that section (amend/restate/add) — the amendment\'s statement, not the resulting section text — plus the text that follows it in the amendment (by position, unverified) and where to read both, or reports the section deleted. Says when the current text is not resolved. Always cites the full lineage. Use get-section for raw text from one document without chain-walking.',
  args: {
    section: tool.schema.string().describe('Section label, e.g. "4" or "7.2(a)" or "EXHIBIT A".'),
    document: tool.schema
      .string()
      .optional()
      .describe(
        "Base document name to resolve from (or a base's stated name when it matched no indexed document), when the matter has more than one amendment chain. Omit to use the matter's longest chain.",
      ),
  },
  async execute(args, ctx) {
    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) return "No documents have been indexed for this matter yet."
    const db = new Database(dbPath, { readonly: true })
    // Child documents (attachment / zip entry) named in the answer get their
    // provenance chain once, after it, so every line keeps the plain name.
    const provenance = sourceProvenance(ctx.directory)
    const withProvenance = (text: string, names: Iterable<string>) => text + provenanceFooter(names, provenance)

    const chains = buildChains(db)
    if (chains.length === 0) {
      const rows = fallbackSectionText(db, args.section, args.document)
      db.close()
      if (!rows.length) return missingNote(args.section, args.document, "no amendment chain in this matter")
      return withProvenance(`${renderChunks(rows)}\n\nNote: no amendment chain in this matter.`, rows.map((row) => row.doc_name))
    }

    // A base is named by its document, or by its stated name when the base
    // matched no indexed document.
    const baseLabel = (c: Chain) => (c.base.kind === "resolved" ? c.base.doc.name : c.base.targetName)
    const candidates = args.document
      ? chains.filter((c) => baseLabel(c).toLowerCase() === args.document!.toLowerCase())
      : chains

    if (candidates.length === 0) {
      // Document arg didn't match any chain base. Fall back to that document's
      // own section text — it may be a standalone doc the agent asked about, or
      // the base of a chain whose stated name matched no indexed document.
      const unmatched = chains.flatMap((c) => (c.base.kind === "unresolved" ? [`"${c.base.targetName}"`] : []))
      const reason = unmatched.length
        ? `"${args.document}" was not matched to any amendment chain base, but it may be one of the stated bases not matched to an indexed document: ${unmatched.join(", ")}. ` +
          "Pass a stated name as the document argument to trace that chain"
        : undefined
      const rows = fallbackSectionText(db, args.section, args.document)
      db.close()
      if (!rows.length) return missingNote(args.section, args.document, reason ?? "no amendment chain involves this document")
      return withProvenance(`${renderChunks(rows)}\n\nNote: ${reason ?? `no amendment chain involves "${args.document}"`}.`, rows.map((row) => row.doc_name))
    }

    if (candidates.length > 1) {
      const bases = candidates.map(baseLabel).join(", ")
      db.close()
      return withProvenance(
        `Multiple amendment chains could contain §${args.section}. ` +
          `Disambiguate by passing the document argument. Candidate base documents: ${bases}.`,
        candidates.flatMap((c) => (c.base.kind === "resolved" ? [c.base.doc.name] : [])),
      )
    }

    const chain = candidates[0]!
    const baseName = chain.base.kind === "resolved" ? chain.base.doc.name : `"${chain.base.targetName}" (unresolved)`
    const resolved = resolveCurrentSection(db, chain, args.section)

    if (!resolved) {
      // No amender touches the section; return base text with a clear note.
      if (chain.base.kind === "resolved") {
        const rows = fallbackSectionText(db, args.section, chain.base.doc.name)
        db.close()
        if (!rows.length) return missingNote(args.section, args.document, `no amendment modifies §${args.section}`)
        return withProvenance(`${renderChunks(rows)}\n\nNote: no amendment modifies §${args.section}. Text below is from the base document.`, rows.map((row) => row.doc_name))
      }
      // The stated base matched no indexed document. Never offer an amending
      // document's own clause as the base text; list the other documents'
      // sections with that label and say the base is unknown.
      const amenders = new Set(chains.flatMap((c) => c.entries.map((e) => e.rel.doc_path)))
      const rows = fallbackSectionText(db, args.section, undefined, amenders)
      db.close()
      const unmatched = `the base document "${chain.base.targetName}" was not matched to an indexed document`
      // The lookup spans the matter; a stated name is not a document to scope by.
      if (!rows.length) return missingNote(args.section, undefined, `no amendment modifies §${args.section}, and ${unmatched}`)
      return withProvenance(
        `Note: no amendment modifies §${args.section}, but ${unmatched}. ` +
          `Sections labelled §${args.section} in the matter's non-amending documents (none is confirmed as the base):\n\n${renderChunks(rows)}`,
        rows.map((row) => row.doc_name),
      )
    }

    if (resolved.operation === "delete") {
      const lineage = renderLineage(db, chain, args.section)
      db.close()
      return withProvenance(
        `=== Current state of §${args.section} (chain root: ${baseName}) ===\n` +
          `Deleted by ${resolved.amender.rel.doc_name} — no current text.\n\n${lineage}`,
        chainNames(chain),
      )
    }

    // The index records where the amendment operates (the op span), not the
    // resulting section text. Quote the amendment's own operative wording and
    // the text after it by position only, labelled as such, and say the
    // current text is not resolved.
    const { op, operativeText, followingText, read } = resolved
    const lineage = renderLineage(db, chain, args.section, operativeText ? op : undefined)
    db.close()
    const amender = resolved.amender.rel.doc_name
    const span = `chars ${op.char_start}..${op.char_end}`
    // Tool-neutral: names the document and position without assuming a tool.
    const position =
      read.offset === null || read.limit === null
        ? null
        : `from code-point offset ${read.offset} (${read.limit} characters)`
    const unavailable = `around ${span}; its position is unavailable because the indexed text is incomplete.`
    const header =
      `=== Current state of §${args.section} (chain root: ${baseName}) ===\n` +
      `Latest operative source: ${amender} — ${conjugate(resolved.operation)} §${args.section}.\n\n`
    if (!operativeText) {
      return withProvenance(
        header +
          `Note: the current text of §${args.section} is not resolved — the amendment's operative wording at ${span} is not recoverable from the index.\n` +
          `Next step: read ${amender} ${position ? `${position} to see the amendment's text at that position.` : unavailable}\n\n${lineage}`,
        chainNames(chain),
      )
    }
    const following = followingText
      ? `Text following the amendment's statement in ${amender} (verbatim; may contain the restated wording; not verified as the resulting text of §${args.section}):\n${followingText}\n\n`
      : ""
    return withProvenance(
      header +
      `Operative wording of ${amender} (verbatim, ${span}). This is the amendment's statement, not the resulting text of §${args.section}:\n` +
      `${quote(operativeText)}\n\n` +
      following +
      `Note: the current text of §${args.section} is not resolved — the index records where the amendment operates, not the resulting section text.\n` +
      `Next step: read ${amender} ${position ? `${position} — the statement and the text after it — and confirm the resulting wording of §${args.section} there before quoting it as current.` : unavailable}\n\n${lineage}`,
      chainNames(chain),
    )
  },
})

function missingNote(section: string, document: string | undefined, reason: string) {
  const scope = document ? `in "${document}"` : "in this matter"
  return (
    `No section "${section}" ${scope}. ${reason.charAt(0).toUpperCase() + reason.slice(1)}. ` +
    "Call get-section with no arguments to see the matter outline, or use search-document to locate the passage by content."
  )
}
