import type { RevisionListEntry } from "docxodus"
import { docxodus } from "./docxodus"

// A lawyer's decisions on a counterparty's tracked changes (#200), applied to
// one version's bytes in one Docxodus session. Ids are the reader's
// (dochaus/lib/docx-markup.ts): one per Docxodus unit, in document order, and
// only `supported` units can be decided. A listed id without a decision is held.
//
// - Docxodus derives an id from its unit's content, so resolving a revision
//   nested in another changes the enclosing one's id. Each step therefore finds
//   its unit again by the native carriers (constituentKeys) it was listed with.
// - A unit with no carrier left was resolved by another decision, such as a
//   deletion inside an insertion that was rejected: `consumed`, whether it was
//   decided or held.
// - Holds are not touched: they stay pending with their own author and date.
// - Any refusal refuses the whole set; nothing is returned to write. The session
//   writes no anchor ids into the counterparty's file.
export type MarkupDecision = "accept" | "reject" | "hold"
export type MarkupOutcome = "applied" | "consumed" | "held"
export type MarkupDecisionErrorCode = "unknown_revision" | "unsupported_revision" | "decision_failed"

export class MarkupDecisionError extends Error {
  constructor(readonly code: MarkupDecisionErrorCode, message: string, readonly revision: string) {
    super(message)
  }
}

export async function applyMarkupDecisions(bytes: Uint8Array, decisions: ReadonlyMap<string, MarkupDecision>) {
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try {
    const listed = session.listRevisions()
    for (const id of decisions.keys())
      if (!listed.some(entry => entry.id === id))
        throw new MarkupDecisionError("unknown_revision", "A decided change is not in this version of the document. Refresh the list and decide again.", id)
    const decision = (entry: RevisionListEntry) => decisions.get(entry.id) ?? "hold"
    const unsupported = listed.find(entry => decision(entry) !== "hold" && entry.resolutionStatus !== "supported")
    if (unsupported)
      throw new MarkupDecisionError("unsupported_revision", "One of the decided changes cannot be resolved here. Hold it and resolve it in Word.", unsupported.id)
    const outcomes = new Map<string, MarkupOutcome>()
    for (const entry of listed) {
      const chosen = decision(entry)
      if (chosen === "hold") continue
      const live = liveUnit(session.listRevisions(), entry, true)
      if (!live) {
        outcomes.set(entry.id, "consumed")
        continue
      }
      const result = chosen === "accept" ? session.acceptRevision(live.id) : session.rejectRevision(live.id)
      if (!result.success)
        throw new MarkupDecisionError("decision_failed", `A change could not be ${chosen}ed (${result.error?.code ?? "refused"}). Nothing was changed; resolve it in Word.`, entry.id)
      outcomes.set(entry.id, "applied")
    }
    const remaining = session.listRevisions()
    for (const entry of listed)
      if (decision(entry) === "hold") outcomes.set(entry.id, liveUnit(remaining, entry, false) ? "held" : "consumed")
    return { bytes: session.save(), outcomes: listed.map(entry => ({ id: entry.id, decision: decision(entry), outcome: outcomes.get(entry.id)! })) }
  } finally {
    session.close()
  }
}

/** The live unit holding `entry`'s carriers now; undefined when none is left.
 * With `resolving`, a unit that has also taken another unit's carriers (Docxodus
 * merges adjacent same-author insertions once a deletion between them goes)
 * refuses: resolving it would resolve a held or differently decided change too. */
function liveUnit(live: RevisionListEntry[], entry: RevisionListEntry, resolving: boolean) {
  const keys = new Set(entry.constituentKeys)
  const found = live.filter(row => row.partUri === entry.partUri && row.constituentKeys.some(key => keys.has(key)))
  if (resolving && (found.length > 1 || found.some(row => row.constituentKeys.some(key => !keys.has(key)))))
    throw new MarkupDecisionError("decision_failed", "Two changes became one when another change was resolved. Decide them the same way, or resolve them in Word.", entry.id)
  return found[0]
}
