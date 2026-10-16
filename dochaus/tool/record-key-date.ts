import { tool } from "./legacy-definition"
import path from "node:path"
import { KEY_DATE_KINDS, recordKeyDate } from "../lib/key-dates"

// doc.haus record-key-date tool. Writes one typed row into the matter's key-dates
// register (.dochaus/legal.db) so renewals and deadlines found during a review
// survive the chat session. Reviews call this as a side-effect — even when the
// clause is unremarkable — and the key-dates tool reads the register back
// deterministically.

export default tool({
  description:
    "Record one key date from a document into the matter's persistent key-dates register: a renewal, termination-notice window, expiry, payment milestone, deliverable deadline, or other dated obligation. Record dates exactly as the document states them: pass dateIso only when the document fixes a calendar date (or one is computable from an anchor date the document itself states); pass dateFormula with the verbatim relative formula when no calendar date is resolvable. Leave unresolved fields out and NAME them in your reply rather than guessing — never invent a date, notice window, or party. quote is the verbatim clause text fixing the date.",
  args: {
    document: tool.schema.string().describe("Document file name within the matter the date comes from"),
    kind: tool.schema.enum(KEY_DATE_KINDS).describe("What kind of date this is"),
    dateIso: tool.schema
      .string()
      .optional()
      .describe("The calendar date in ISO form (YYYY-MM-DD), only when the document resolves to one"),
    dateFormula: tool.schema
      .string()
      .optional()
      .describe(
        'The verbatim relative formula when no calendar date is resolvable, e.g. "90 days prior to the end of the then-current term"',
      ),
    noticeDays: tool.schema.number().int().optional().describe("Notice window in days, when the document states one"),
    cancelByIso: tool.schema
      .string()
      .optional()
      .describe("Last calendar date (YYYY-MM-DD) to act (e.g. send non-renewal notice), only when resolvable"),
    party: tool.schema.string().optional().describe("The party the date binds, as the document names it"),
    quote: tool.schema.string().describe("Verbatim clause text fixing the date"),
    reason: tool.schema.string().describe("Why this date matters, in one line for the register"),
    author: tool.schema.string().describe("Who is recording the date (the reviewing agent or lawyer)"),
  },
  async execute(args, ctx) {
    if (!args.dateIso && !args.dateFormula)
      return "Nothing to record: pass dateIso (a resolved calendar date) or dateFormula (the document's verbatim relative formula). Never guess either — if both are unresolvable, the clause fixes no date."
    // The register sorts lexicographically on date_iso, so only strict
    // YYYY-MM-DD may land in it — a loose format silently corrupts the order.
    const malformed = [args.dateIso, args.cancelByIso].filter(
      (d): d is string => d !== undefined && (!/^\d{4}-\d{2}-\d{2}$/.test(d) || Number.isNaN(Date.parse(d))),
    )
    if (malformed.length)
      return `Not recorded: ${malformed.map((d) => JSON.stringify(d)).join(" and ")} is not a valid YYYY-MM-DD calendar date. dateIso and cancelByIso must be strict ISO dates; if the document fixes no calendar date, pass dateFormula with the verbatim relative formula instead.`
    const id = recordKeyDate(ctx.directory, {
      docName: path.basename(args.document),
      kind: args.kind,
      dateIso: args.dateIso,
      dateFormula: args.dateFormula,
      noticeDays: args.noticeDays,
      cancelByIso: args.cancelByIso,
      party: args.party,
      quote: args.quote,
      reason: args.reason,
      author: args.author,
    })
    return {
      title: `Recorded ${args.kind} date for ${path.basename(args.document)}`,
      output: `Recorded key date #${id}: ${args.kind} in ${path.basename(args.document)} — ${args.dateIso ?? `formula: ${JSON.stringify(args.dateFormula)}`}${args.cancelByIso ? `, act by ${args.cancelByIso}` : ""}${args.noticeDays != null ? `, ${args.noticeDays}-day notice` : ""}. The register persists in this matter; read it with the key-dates tool.`,
      metadata: { keyDate: id, kind: args.kind, document: path.basename(args.document) },
    }
  },
})
