import { tool } from "./legacy-definition"
import { ALL_GATES, runVerificationGates, type GateId } from "../lib/verification-loop"

// doc.haus verify tool. Runs the six deterministic verification gates over a
// deliverable and writes verification-report.md. Lawyer- and agent-facing; the
// verify-work-product skill states when to invoke.

export default tool({
  description:
    'Run deterministic verification gates over a matter deliverable before filing, sending, or exporting. Checks citation anchoring, value grounding, research-source audit, render readiness, redline anchor integrity, and privilege/export gates. Writes verification-report.md and returns PASS / PARTIAL / BLOCKED. Use before any export, file, or send; at drafting/review/redline stage boundaries; or when the lawyer asks to pre-flight, sanity-check, or gate-check a deliverable.',
  args: {
    deliverable: tool.schema
      .string()
      .optional()
      .describe('Deliverable filename to verify (default: auto-discover the assembled markdown).'),
    gates: tool.schema
      .array(tool.schema.enum(ALL_GATES as [GateId, ...GateId[]]))
      .optional()
      .describe('Subset of gates to run (default: all). One of citation-anchors, value-grounding, research-sources, render-readiness, redline-anchors, privilege-export.'),
    output: tool.schema
      .string()
      .optional()
      .describe('Report filename (default "verification-report.md").'),
  },
  async execute(args, ctx) {
    const report = await runVerificationGates({
      directory: ctx.directory,
      deliverable: args.deliverable,
      gates: args.gates,
      output: args.output,
    })
    return {
      title: `Verification: ${report.status}`,
      output: `${report.status}: ${report.deliverable ?? "no deliverable discovered"}. Wrote ${report.reportFile}.`,
      metadata: {
        status: report.status,
        deliverable: report.deliverable,
        gates: Object.fromEntries(
          Object.entries(report.gates).map(([id, r]) => [id, r.status]),
        ),
        reportFile: report.reportFile,
      },
    }
  },
})
