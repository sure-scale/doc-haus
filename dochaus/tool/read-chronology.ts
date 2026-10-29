import { tool } from "./legacy-definition"
import { chronologyRequest } from "../lib/chronology-client"
import type { Chronology } from "../lib/chronology"
import { UNTRUSTED_RECORD_NOTICE } from "../lib/untrusted"

export default tool({
  description:
    "Read this matter's saved chronology, including sourced accounts, uncertainty, privilege flags, hidden status, freshness and lawyer review state together. Verified quotations prove a text match, not that an event occurred. Read view=events for stable event IDs, event for paged accounts, evidence for paged quotations from an event, sources for coverage/source versions, or proposals for changes awaiting the lawyer. Hidden events are included and explicitly labelled. This never starts extraction or changes review decisions. Follow next_offset; use revision to detect changes while paging.",
  args: {
    view: tool.schema.enum(["events", "event", "evidence", "sources", "proposals"]).optional(),
    event_id: tool.schema.string().optional(),
    query: tool.schema.string().max(300).optional(),
    offset: tool.schema.number().int().min(0).optional(),
    limit: tool.schema.number().int().min(1).max(10).optional(),
  },
  async execute(args, ctx) {
    const { matterId, result } = await chronologyRequest(ctx.directory, ctx.abort)
    const state = result as unknown as Chronology
    const view = args.view ?? "events"
    const query = args.query?.trim().toLocaleLowerCase()
    const offset = args.offset ?? 0
    const limit = args.limit ?? 5
    let context: Record<string, unknown> = {}
    let items: unknown[]
    if (view === "event" || view === "evidence") {
      const event = state.events.find((e) => e.id === args.event_id)
      if (!event) return "That saved event is not in this matter. Read view=events to obtain its current ID."
      const { accounts, ...entry } = event
      context = { event: entry }
      items =
        view === "evidence"
          ? accounts.flatMap(({ evidence, ...account }) => evidence.map((source) => ({ account, source })))
          : accounts.map(({ evidence, ...account }) => ({
              ...account,
              evidence: evidence.map(({ excerpt, ...source }) => ({ ...source, quotationLength: excerpt.length })),
            }))
    } else if (view === "sources")
      items = state.coverage.map(({ completedKeys, ...source }) => ({
        ...source,
        completedSectionCount: completedKeys.length,
      }))
    else if (view === "proposals")
      items = state.proposals.filter((p) => p.status === "pending" || p.status === "deferred")
    else
      items = state.events
        .filter((e) => !query || JSON.stringify(e).toLocaleLowerCase().includes(query))
        .map((e) => ({
          id: e.id,
          title: e.title,
          detail: e.detail,
          date: e.date,
          origin: e.origin,
          hidden: e.hidden,
          reviewed: e.reviewed,
          disputed: e.disputed,
          dispute: e.dispute,
          note: e.note,
          key: e.key,
          lawyerEdited: e.lawyerEdited,
          previousIds: e.previousIds,
          supersededBy: e.supersededBy,
          accounts: e.accounts.map((a) => ({
            id: a.id,
            date: a.date,
            status: a.status,
            privilege: a.privilege,
            people: a.people,
            nameAssociations: a.nameAssociations,
            evidence: a.evidence.map((s) => ({
              id: s.id,
              sourceId: s.sourceId,
              documentName: s.documentName,
              verified: s.verified,
              freshness: s.freshness,
              issue: s.issue,
            })),
          })),
        }))
    if (offset > items.length) return "Offset is beyond this view. Reload at offset 0; the chronology may have changed."
    const selected: unknown[] = []
    const page = () => ({
      matterId,
      revision: state.revision,
      view,
      framing: state.framing,
      ...context,
      offset,
      total: items.length,
      next_offset: offset + selected.length < items.length ? offset + selected.length : null,
      items: selected,
      link: `/matter/${encodeURIComponent(matterId)}?view=chronology${args.event_id ? `&event=${encodeURIComponent(args.event_id)}` : ""}`,
      notice: UNTRUSTED_RECORD_NOTICE,
    })
    for (const item of items.slice(offset, offset + limit)) {
      selected.push(item)
      if (JSON.stringify(page()).length <= 18000) continue
      selected.pop()
      break
    }
    // Large entries can be inspected through read-chronology's focused event
    // view; a single exceptionally large account is reported rather than clipped.
    if (!selected.length && offset < items.length)
      return "This item is too large for one chronology response. Use view=event or view=evidence with its event_id, or inspect it in the Chronology workbench. No excerpt has been clipped."
    return {
      title: "Saved chronology",
      output: JSON.stringify(page()),
      metadata: {
        chronologyMatterId: matterId,
        chronologyRevision: state.revision,
        truncated: page().next_offset !== null,
      },
    }
  },
})
