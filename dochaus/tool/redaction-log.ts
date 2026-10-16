import { tool } from "./legacy-definition"
import path from "node:path"
import { listRedactions } from "../lib/redactions"

// doc.haus redaction-log tool. Reads the matter's redaction log — the record the
// redact tool writes for every removal (what was removed, by whom, why, when).
// This is the work product a lawyer attaches to a production: the redacted
// document plus the log justifying each removal.

export default tool({
  description:
    "Read the matter's redaction log: every redaction performed in this matter — the text removed, the replacement label, the reason, who performed it, how many occurrences, and when. Optionally filter to one document. Use this to produce a redaction log for a production set or to check what has already been removed.",
  args: {
    document: tool.schema.string().optional().describe("Limit the log to one document file name within the matter"),
  },
  async execute(args, ctx) {
    const docPath = args.document
      ? path.isAbsolute(args.document)
        ? args.document
        : path.join(ctx.directory, args.document)
      : undefined
    const rows = listRedactions(ctx.directory, docPath)
    const lines =
      rows.length === 0
        ? "No redactions recorded in this matter."
        : rows
            .map(
              (r) =>
                `#${r.id} ${new Date(r.created_at).toISOString().slice(0, 16).replace("T", " ")} | ${r.doc_name} | removed ${JSON.stringify(r.redacted_text)} → ${JSON.stringify(r.label)} (${r.occurrences} occurrence(s), ${r.metadata_hits} metadata hit(s)) | reason: ${r.reason} | by ${r.author}`,
            )
            .join("\n")
    return {
      title: `${rows.length} redaction(s)${args.document ? ` in ${path.basename(args.document)}` : ""}`,
      output: lines,
      metadata: { redactions: rows.map((r) => r.id) },
    }
  },
})
