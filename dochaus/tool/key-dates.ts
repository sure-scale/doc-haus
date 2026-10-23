import { tool } from "./legacy-definition"
import path from "node:path"
import { listKeyDates } from "../lib/key-dates"
import { provenanceFooter, sourceProvenance } from "../lib/source-provenance"

// doc.haus key-dates tool. Reads the matter's key-dates register — the rows the
// record-key-date tool writes during reviews (renewals, termination-notice
// windows, expiries, payment milestones). Read-only: the tool only echoes
// stored rows, soonest calendar date first, formula-only rows last. An empty
// register returns an explicit all-clear, because silence is indistinguishable
// from failure.

export default tool({
  description:
    "Read the matter's key-dates register: every renewal, termination-notice window, expiry, payment milestone, and deliverable deadline recorded in this matter, soonest first. Optionally filter to one document. Use this to build a compliance calendar or check what deadlines a prior review already recorded.",
  args: {
    document: tool.schema
      .string()
      .optional()
      .describe("Limit the register to one document file name within the matter"),
  },
  async execute(args, ctx) {
    const rows = listKeyDates(ctx.directory, args.document ? path.basename(args.document) : undefined)
    const lines =
      rows.length === 0
        ? "No key dates recorded in this matter."
        : rows
            .map(
              (r) =>
                `#${r.id} ${r.date_iso ?? "(no calendar date)"} | ${r.kind} | ${r.doc_name}${r.party ? ` | binds ${r.party}` : ""}${r.notice_days != null ? ` | ${r.notice_days}-day notice` : ""}${r.cancel_by_iso ? ` | act by ${r.cancel_by_iso}` : ""}${r.date_formula ? ` | formula: ${JSON.stringify(r.date_formula)}` : ""} | ${r.reason} | ${JSON.stringify(r.quote)} | by ${r.author}`,
            )
            .join("\n")
    return {
      title: `${rows.length} key date(s)${args.document ? ` in ${path.basename(args.document)}` : ""}`,
      output: lines + provenanceFooter(rows.map((r) => r.doc_name), sourceProvenance(ctx.directory)),
      metadata: { keyDates: rows.map((r) => r.id) },
    }
  },
})
