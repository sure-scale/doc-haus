import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { buildChains, type Chain } from "../lib/amendment-chain"

// doc.haus amendment-chain lookup. Reads doc_relations and section_ops rows
// (both written by services/ingest/src/structure.ts) and reports which matter
// documents amend which, the verbatim recital stating it, and which sections
// each amender operates on. Use before relying on any agreement that may have
// been amended: the agent reads the operative version first instead of an
// amended base. Read-only; all writes happen in the ingest service.

export default tool({
  description:
    "Show amendment relationships between this matter's documents: which document amends which, the verbatim recital stating it, which sections each amender operates on (amend/restate/delete/add), and which document to read first. Use this before relying on any agreement that may have been amended. Returns only extracted facts; if a relationship is not listed, verify with search-document.",
  args: {
    document: tool.schema
      .string()
      .optional()
      .describe(
        "Exact name of an indexed document to trace from (as amender or as amended target). Omit to list every amendment relationship in the matter.",
      ),
  },
  async execute(args, ctx) {
    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) {
      return "No documents have been indexed for this matter yet."
    }

    const db = new Database(dbPath, { readonly: true })
    const chains = buildChains(db)
    if (chains.length === 0) {
      const count = (db.query("SELECT COUNT(*) AS n FROM doc_relations").get() as { n: number }).n
      db.close()
      return count === 0
        ? "No amendment relationships were extracted for this matter. " +
            "Extraction only covers formulaic amendment recitals (\"Amendment No. _ to ...\", \"amends and restates ...\"); " +
            "if you suspect an amendment phrased differently, use the search-document tool to look for it."
        : `${count} amendment relationship(s) were extracted for this matter, but no chain root could be derived from them (every amended target is itself an amending document). ` +
            "Use the search-document tool to inspect the amendment recitals directly."
    }

    const matchesArg = (chain: Chain) =>
      !args.document ||
      chain.entries.some((e) => e.rel.doc_name.toLowerCase() === args.document!.toLowerCase()) ||
      chain.entries.some((e) => e.rel.target_name.toLowerCase() === args.document!.toLowerCase()) ||
      (chain.base.kind === "resolved" && chain.base.doc.name.toLowerCase() === args.document!.toLowerCase())

    const selected = chains.filter(matchesArg)
    // A document may be the base of a chain whose stated name matched no
    // indexed document; show those chains instead of reporting none.
    const unmatched = chains.filter((c) => c.base.kind === "unresolved")
    if (selected.length === 0 && args.document && unmatched.length > 0) {
      db.close()
      return (
        `No extracted amendment relationship names "${args.document}", but the stated base of these chains was not matched to any indexed document, ` +
        `so "${args.document}" may be one of them. Compare the recitals below with it before relying on it as unamended.\n\n` +
        unmatched.map(renderChain).join("\n\n")
      )
    }
    if (selected.length === 0) {
      const total = chains.reduce((sum, c) => sum + c.entries.length, 0)
      db.close()
      return args.document
        ? `No extracted amendment relationship involves "${args.document}" ` +
            `(the matter has ${total} extracted amendment relationship(s) on other documents). ` +
            "Extraction only covers formulaic amendment recitals; use the search-document tool to check for amendment language phrased differently."
        : `${total} amendment relationship(s) were extracted for this matter, but no chain root could be derived. ` +
            "Use the search-document tool to inspect the amendment recitals directly."
    }

    const blocks = selected.map(renderChain)
    db.close()
    return blocks.join("\n\n")
  },
})

function renderChain(chain: Chain) {
  const ordered = chain.entries
  const base =
    chain.base.kind === "resolved"
      ? chain.base.doc.name
      : `"${chain.base.targetName}" (as stated; not matched to any indexed document)`
  const lines = ordered.map((entry) => {
    const target =
      entry.resolution.kind === "resolved"
        ? `resolved to ${entry.resolution.doc.name}`
        : entry.resolution.kind === "ambiguous"
          ? `not matched to any indexed document (ambiguous between: ${entry.resolution.names.join(", ")})`
          : "not matched to any indexed document"
    const recital = entry.recital
      ? `"${entry.recital}"`
      : "recital text not recoverable from the index at these offsets"
    const opsLine =
      entry.ops.length === 0
        ? ""
        : `\n  Operates on: ${entry.ops
            .map((op) => `§${op.section_label} (${op.operation})`)
            .join(", ")}`
    return (
      `- ${entry.rel.doc_name} ${entry.rel.relation} "${entry.rel.target_name}" — ${target}\n` +
      `  Recital (chars ${entry.rel.char_start}..${entry.rel.char_end} of ${entry.rel.doc_name}): ${recital}` +
      opsLine
    )
  })
  const latest = ordered[ordered.length - 1]!
  // Unmatched amendments are grouped by the name each states, so this chain
  // may not hold every amendment of the real base.
  const readFirst =
    chain.base.kind === "resolved"
      ? `Read ${latest.rel.doc_name} first — it is the most recently indexed amending document. `
      : `Among the documents listed, ${latest.rel.doc_name} is the most recently indexed amending document. ` +
        "The base was not matched to an indexed document, so an amendment that states it differently may be listed under another chain. "
  return (
    `Chain: ${base} is amended (and possibly superseded) by:\n` +
    lines.join("\n") +
    `\n${readFirst}` +
    "Whether each amendment restates the base in its entirety or merely modifies specific provisions must be confirmed in its recitals, quoted verbatim above."
  )
}
