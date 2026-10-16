import { tool } from "./legacy-definition"
import { formatCitationCheckSummary, writeCitationCheckReport } from "../lib/citation-check"

// doc.haus check-citations tool. Runs the deterministic citation checker over
// matter artifacts without the full verification pipeline (no network fetches,
// no quote-anchor verification, no value grounding). Produces a JSON report and
// a concise user-facing summary the agent can surface during review or export.
// The verify tool's citation-anchors gate calls the same library; this tool
// exposes it standalone so the agent can check citations at any point without
// running all six verification gates.

export default tool({
  description:
    'Run the deterministic citation checker over the matter\'s generated artifacts (drafts, memos, research, review packets, issue ledgers). Inventories every legal citation, normalizes it, links it to durable source material (authority ledger, source artifacts, research rows, official fetches), and reports each as verified, unresolved, or unsupported. Writes citation-check-report.json and returns a concise summary. Use this before export, at review boundaries, or when the lawyer asks to check citations — it is faster than the full verify tool because it does no network fetches.',
  args: {
    artifacts: tool.schema
      .array(tool.schema.string())
      .optional()
      .describe('Specific artifact filenames to scan (default: auto-discover draft/agreement/memo/research/review/issue-ledger files).'),
    output: tool.schema
      .string()
      .optional()
      .describe('Report filename (default "citation-check-report.json").'),
  },
  async execute(args, ctx) {
    const report = writeCitationCheckReport({
      directory: ctx.directory,
      artifacts: args.artifacts,
      output: args.output,
    })
    const summary = formatCitationCheckSummary(report)
    const status = report.summary.unsupported > 0
      ? "unsupported citations found"
      : report.summary.unresolved > 0
        ? "unresolved citations remain"
        : "all citations verified"
    return {
      title: `Citation check: ${report.summary.verified}/${report.summary.total} verified (${report.summary.unsupported} unsupported)`,
      output: summary,
      metadata: {
        status,
        total: report.summary.total,
        verified: report.summary.verified,
        unresolved: report.summary.unresolved,
        unsupported: report.summary.unsupported,
        groups: report.summary.groups,
        artifacts: report.artifacts,
        reportFile: report.report_file,
        pinpoints: report.rows
          .filter((row) => row.pinpoint)
          .map((row) => ({
            citation: row.citation_text,
            artifact: `${row.artifact.file}:${row.artifact.line}`,
            reason: row.pinpoint!.reason,
          })),
      },
    }
  },
})
