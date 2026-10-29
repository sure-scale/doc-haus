import { tool } from "./legacy-definition"
import { chronologyRequest } from "../lib/chronology-client"
import { UNTRUSTED_RECORD_NOTICE } from "../lib/untrusted"

const date = tool.schema.object({
  original: tool.schema.string(),
  precision: tool.schema.enum(["exact", "approximate", "month", "year", "ambiguous", "relative", "undated"]),
  start: tool.schema.string().optional(),
  end: tool.schema.string().optional(),
  rangeKind: tool.schema.enum(["duration", "uncertainty"]).optional(),
  time: tool.schema.string().optional(),
  timezone: tool.schema.string().optional(),
  qualifier: tool.schema.string().optional(),
  documentDate: tool.schema.string().optional(),
  relative: tool.schema
    .object({
      expression: tool.schema.string(),
      anchorDocument: tool.schema.string(),
      anchorDate: tool.schema.string().optional(),
      offsetDays: tool.schema.number().optional(),
      calculation: tool.schema.string().optional(),
      interpretation: tool.schema.enum(["document-stated", "legal-interpretation"]).optional(),
    })
    .optional(),
})
const evidence = tool.schema.object({
  id: tool.schema.string(),
  documentName: tool.schema.string(),
  sourceId: tool.schema.string(),
  version: tool.schema.string(),
  excerpt: tool.schema.string().min(10).max(600),
  section: tool.schema.string(),
  charStart: tool.schema.number().int().min(0),
  charEnd: tool.schema.number().int().min(1),
  verified: tool.schema.literal(true),
  freshness: tool.schema.literal("current"),
  provenance: tool.schema.enum(["source", "supplied-chronology"]),
  sourceKind: tool.schema.string().optional(),
})
const account = tool.schema.object({
  id: tool.schema.string(),
  title: tool.schema.string(),
  detail: tool.schema.string(),
  date,
  evidence: tool.schema.array(evidence).min(1),
  attribution: tool.schema.string().optional(),
  privilege: tool.schema.enum(["cleared", "uncertain", "privileged", "not-privileged"]),
  significance: tool.schema.enum(["key", "relevant", "background"]),
  people: tool.schema.array(tool.schema.string()),
  status: tool.schema.enum(["reported", "scheduled", "required"]),
  obligationId: tool.schema.string().optional(),
  gaps: tool.schema.array(tool.schema.string()),
})

export default tool({
  description:
    "Stage exact sourced changes to the saved chronology for the lawyer's confirmation. Never applies changes, marks review, or runs extraction. First read-chronology for current revision/event IDs/source IDs/versions and verify each quote with cite; preserve competing accounts, corrections and hidden status. Use a stable unique proposal id for retrying a lost response. Changes to existing events must include current targetIds; merge/split retain source accounts. Use accountIds to retain existing target accounts, including stale or unavailable historical evidence, without copying quotations; accounts supplies new current verified evidence. Explain the exact proposal and evidence in chat and link the lawyer to its workbench review; only the application's Accept/Keep/Skip controls decide it. Never claim a staged proposal is saved chronology or human approval.",
  args: {
    expectedRevision: tool.schema.number().int().min(0),
    proposals: tool.schema
      .array(
        tool.schema.object({
          id: tool.schema.string().min(1).max(120),
          kind: tool.schema.enum(["addition", "amendment", "conflict", "removal", "merge", "split"]),
          targetIds: tool.schema.array(tool.schema.string()),
          events: tool.schema.array(
            tool.schema.object({
              id: tool.schema.string().optional(),
              title: tool.schema.string(),
              detail: tool.schema.string(),
              date,
              accounts: tool.schema.array(account).optional(),
              accountIds: tool.schema.array(tool.schema.string()).optional(),
            }),
          ),
          reason: tool.schema.string().min(1),
          recommendation: tool.schema.string().min(1),
        }),
      )
      .min(1)
      .max(5),
  },
  async execute(args, ctx) {
    const { matterId, result } = await chronologyRequest(ctx.directory, ctx.abort, {
      ...args,
      proposals: args.proposals.map((proposal) => ({ ...proposal, origin: "chat" })),
    })
    const link = `/matter/${encodeURIComponent(matterId)}?view=chronology&proposal=${encodeURIComponent(args.proposals[0].id)}`
    return {
      title: "Chronology changes awaiting confirmation",
      output: `Staged for review; the saved chronology has not changed.\n\n${JSON.stringify(args.proposals)}\n\n[Review these exact proposals](${link})\n\nChoose Accept, Keep current or Skip for later in the workbench. ${UNTRUSTED_RECORD_NOTICE}`,
      metadata: {
        chronologyMatterId: matterId,
        chronologyProposalIds: args.proposals.map((p) => p.id),
        chronologyRevision: result.revision,
        chronologyLink: link,
      },
    }
  },
})
