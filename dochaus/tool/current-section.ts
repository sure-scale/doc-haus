import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { buildChains, resolveCurrentSection, verbatimSpan, type Chain } from "../lib/amendment-chain"

// doc.haus current-section resolver. For a section label, walks the matter's
// amendment chains newest-first and returns the operative text from whichever
// amender last touches that section, plus the full lineage so the answer is
// auditable. Falls back honestly when no chain exists, no amender touches the
// section, or the amender's replacement text is not separately indexed.
//
// Differs from get-section, which returns one document's verbatim section text
// without walking the chain. Use current-section for "what does §X currently
// say?"; use get-section for "what does document Y say at §X?".
//
// Cap on returned text mirrors get-section so one oversized section cannot
// flood the context window. Read-only; all writes happen in the ingest service.

const TEXT_CAP = 20000

type Chunk = { doc_name: string; section: string; text: string; char_start: number; char_end: number }

function fallbackSectionText(db: Database, section: string, document: string | undefined) {
  const where = document ? "doc_name = ? AND section = ?" : "section = ?"
  const params = document ? [document, section] : [section]
  const rows = db
    .query(`SELECT doc_name, section, text, char_start, char_end FROM chunks WHERE ${where} ORDER BY doc_name, char_start`)
    .all(...params) as Chunk[]
  return rows
}

function renderChunks(rows: Chunk[]) {
  const groups = new Map<string, Chunk[]>()
  for (const row of rows) {
    const key = `${row.doc_name}\x00${row.section}`
    groups.set(key, [...(groups.get(key) ?? []), row])
  }
  const parts: string[] = []
  let used = 0
  for (const rows of groups.values()) {
    const kept: Chunk[] = []
    for (const row of rows) {
      if (used + row.text.length > TEXT_CAP && parts.length + kept.length > 0) break
      kept.push(row)
      used += row.text.length
    }
    if (kept.length) {
      const first = kept[0]!
      const header = `=== ${first.doc_name} — § ${first.section} (chars ${first.char_start}–${kept[kept.length - 1]!.char_end}) ===`
      parts.push(`${header}\n${kept.map((row) => row.text).join("")}`)
    }
  }
  return parts.join("\n\n")
}

function conjugate(operation: "amend" | "restate" | "delete" | "add") {
  return operation + (operation.endsWith("e") ? "d" : "ed")
}

function renderLineage(db: Database, chain: Chain, label: string) {
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
    const opVerbatim = verbatimSpan(db, op.doc_path, op.char_start, op.char_end)
    const verbatim = opVerbatim ? `"${opVerbatim}"` : "operative text not recoverable from the index at these offsets"
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
    'Resolve the current operative text of a contract section across the matter\'s amendment chain — e.g. "what does Section 4 currently say?". Walks amendments newest-first and returns the latest replacement text (amend/restate/add) or reports the section deleted. Always cites the full lineage. Use get-section for raw text from one document without chain-walking.',
  args: {
    section: tool.schema.string().describe('Section label, e.g. "4" or "7.2(a)" or "EXHIBIT A".'),
    document: tool.schema
      .string()
      .optional()
      .describe(
        "Base document name to resolve from, when the matter has more than one amendment chain. Omit to use the matter's longest chain.",
      ),
  },
  async execute(args, ctx) {
    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) return "No documents have been indexed for this matter yet."
    const db = new Database(dbPath, { readonly: true })

    const chains = buildChains(db)
    if (chains.length === 0) {
      const rows = fallbackSectionText(db, args.section, args.document)
      db.close()
      if (!rows.length) return missingNote(args.section, args.document, "no amendment chain in this matter")
      return `${renderChunks(rows)}\n\nNote: no amendment chain in this matter.`
    }

    const candidates = args.document
      ? chains.filter(
          (c) => c.base.kind === "resolved" && c.base.doc.name.toLowerCase() === args.document!.toLowerCase(),
        )
      : chains

    if (candidates.length === 0) {
      // Document arg didn't match any chain base. Fall back to that document's
      // own section text — it may be a standalone doc the agent asked about.
      const rows = fallbackSectionText(db, args.section, args.document)
      db.close()
      if (!rows.length) return missingNote(args.section, args.document, "no amendment chain involves this document")
      return `${renderChunks(rows)}\n\nNote: no amendment chain involves "${args.document}".`
    }

    if (candidates.length > 1) {
      const bases = candidates.map((c) => (c.base.kind === "resolved" ? c.base.doc.name : c.base.targetName)).join(", ")
      db.close()
      return (
        `Multiple amendment chains could contain §${args.section}. ` +
        `Disambiguate by passing the document argument. Candidate base documents: ${bases}.`
      )
    }

    const chain = candidates[0]!
    const baseName = chain.base.kind === "resolved" ? chain.base.doc.name : `"${chain.base.targetName}" (unresolved)`
    const resolved = resolveCurrentSection(db, chain, args.section)

    if (!resolved) {
      // No amender touches the section; return base text with a clear note.
      const rows = fallbackSectionText(db, args.section, chain.base.kind === "resolved" ? chain.base.doc.name : undefined)
      db.close()
      if (!rows.length) return missingNote(args.section, args.document, `no amendment modifies §${args.section}`)
      return `${renderChunks(rows)}\n\nNote: no amendment modifies §${args.section}. Text below is from the base document.`
    }

    const lineage = renderLineage(db, chain, args.section)
    if (resolved.operation === "delete") {
      db.close()
      return (
        `=== Current state of §${args.section} (chain root: ${baseName}) ===\n` +
        `Deleted by ${resolved.amender.rel.doc_name} — no current text.\n\n${lineage}`
      )
    }

    if (!resolved.replacementText) {
      const op = resolved.amender.ops.filter((o) => o.section_label === args.section).slice(-1)[0]!
      const verb = conjugate(resolved.operation)
      db.close()
      return (
        `=== Current state of §${args.section} (chain root: ${baseName}) ===\n` +
        `Latest operative source: ${resolved.amender.rel.doc_name} — ${verb} §${args.section} (operative span chars ${op.char_start}..${op.char_end}).\n` +
        `Note: replacement text not separately indexed in ${resolved.amender.rel.doc_name} (no chunk labeled "${args.section}"). ` +
        `Use read-document on "${resolved.amender.rel.doc_name}" at chars ${op.char_start}..${op.char_end} to view verbatim.\n\n${lineage}`
      )
    }

    const text = resolved.replacementText.slice(0, TEXT_CAP)
    const truncation = resolved.replacementText.length > TEXT_CAP
      ? `\n[Truncated at the ${TEXT_CAP}-character cap. Use get-section with document: "${resolved.amender.rel.doc_name}" and section: "${args.section}" to fetch the rest.]`
      : ""
    db.close()
    const opVerb = conjugate(resolved.operation)
    return (
      `=== Current text of §${args.section} (chain root: ${baseName}) ===\n` +
      `Latest operative source: ${resolved.amender.rel.doc_name} — ${opVerb} §${args.section}.\n\n` +
      `Current text (verbatim from ${resolved.amender.rel.doc_name}, chunks labeled §${args.section}):\n${text}${truncation}\n\n${lineage}`
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
