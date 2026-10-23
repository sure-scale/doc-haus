import type { Turn } from "./turn-model"

// Auto's acknowledgement: the router's one line restating the request ("On
// it — I'll check the lease for the break clause."), shown under the lawyer's
// message while the chosen assistant starts. It is display only. It lives on
// the user turn's `ack`, never in any turn's text, so it is never routing
// history, assistant context, quick-draft text, an upgrade's working draft or
// drafting-run input. The router session that wrote it is deleted, so the
// line is kept here by user message id for as long as the page is open:
// settling and remounting reload turns from the engine, which has no copy.

const acks = new Map<string, string>()

// The chat's own line when the router wrote none it could show.
export const FALLBACK_ACK = {
  answer: "Looking into that now.",
  draft: "Starting on that draft.",
  offer: "Happy to draft that.",
} as const

/** The pending user turn for `text` with its acknowledgement, unless it has
 * one already: the first line shown stays. */
export function withAck(turns: Turn[], text: string, ack: string) {
  const last = turns.at(-1)
  if (last?.role !== "user" || last.id || last.text !== text || last.ack) return turns
  return [...turns.slice(0, -1), { ...last, ack }]
}

/** The pending user turn for `text` without its acknowledgement (routing or
 * the send failed, so nothing is being handed over). */
export function withoutAck(turns: Turn[], text: string) {
  const last = turns.at(-1)
  if (last?.role !== "user" || last.id || last.text !== text || !last.ack) return turns
  const { ack: _ack, ...rest } = last
  return [...turns.slice(0, -1), rest]
}

/** Stored turns replacing the shown ones keep the shown acknowledgements: a
 * pending turn's passes to the newest stored user message with its text. */
export function carryAcks(shown: Turn[], stored: Turn[]) {
  for (const turn of shown) {
    if (turn.role !== "user" || !turn.ack) continue
    if (turn.id) {
      acks.set(turn.id, turn.ack)
      continue
    }
    const match = stored.findLast((t) => t.role === "user" && t.id && t.text === turn.text && !acks.has(t.id))
    if (match?.id) acks.set(match.id, turn.ack)
  }
  return withAcks(stored)
}

/** Stored turns with the acknowledgements shown for them on this page. */
export function withAcks(turns: Turn[]) {
  return turns.map((turn) => {
    const ack = turn.role === "user" && turn.id ? acks.get(turn.id) : undefined
    return ack ? { ...turn, ack } : turn
  })
}
