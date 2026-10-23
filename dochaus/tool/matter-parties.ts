import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { provenanceFooter, sourceProvenance } from "../lib/source-provenance"

// doc.haus party lookup tool. Reads the `parties` table that
// `services/ingest` populates deterministically (structure.ts): each row is a
// party name regex-extracted verbatim from a document's preamble, with offsets
// into that document's extracted text. Nothing here is generated — the tool
// only echoes stored rows, so it cannot fabricate a party. A parse miss is an
// absent row, and the tool says so and points at search-document instead of
// guessing. This tool is read-only; all writes happen in the ingest service.

type PartyRow = {
  doc_name: string
  name: string
  role: string | null
  char_start: number
  char_end: number
}

export default tool({
  description:
    "List the parties to each document in the current matter (or to one document), as extracted verbatim from document preambles. Use this to answer who the parties to an agreement are; the listing may be incomplete, so confirm with search-document when a party you expect is missing.",
  args: {
    document: tool.schema
      .string()
      .optional()
      .describe("Restrict to a single document by its exact name. Omit to list parties for the whole matter."),
  },
  async execute(args, ctx) {
    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) {
      return "No documents have been indexed for this matter yet."
    }

    const limit = 500
    const db = new Database(dbPath, { readonly: true })
    const sql =
      "SELECT doc_name, name, role, char_start, char_end FROM parties" +
      (args.document ? " WHERE doc_name = ?" : "") +
      ` ORDER BY doc_name, char_start LIMIT ${limit}`
    const rows = (args.document ? db.query(sql).all(args.document) : db.query(sql).all()) as PartyRow[]
    const indexed = args.document
      ? db.query("SELECT 1 FROM documents WHERE name = ?").get(args.document) !== null
      : true
    db.close()

    if (args.document && !indexed) {
      return `"${args.document}" is not an indexed document in this matter. Use the search-document tool to find the document you mean.`
    }

    if (rows.length === 0) {
      const scope = args.document ? `"${args.document}"` : "this matter"
      return `No parties were extracted from the preamble of ${scope}. Use the search-document tool with the party name (or with a query like "by and between") to find the parties directly in the document text.`
    }

    const byDoc = Map.groupBy(rows, (row) => row.doc_name)

    const body = [...byDoc.entries()]
      .map(
        ([docName, parties]) =>
          `${docName}\n` +
          parties
            .map(
              (p) =>
                `- ${p.name} ${p.role ? `(${p.role})` : "(role unstated)"} [chars ${p.char_start}-${p.char_end}]`,
            )
            .join("\n"),
      )
      .join("\n\n")

    const truncation =
      rows.length === limit
        ? `\n\nNote: the listing was truncated at ${limit} parties. Pass the document argument to narrow the scope.`
        : ""

    return (
      body +
      provenanceFooter(byDoc.keys(), sourceProvenance(ctx.directory)) +
      truncation +
      "\n\nNote: parties are extracted verbatim from each document's preamble and the list may be incomplete. Use the search-document tool to confirm any party."
    )
  },
})
