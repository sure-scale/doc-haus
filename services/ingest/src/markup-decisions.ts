import type { RevisionListEntry } from "docxodus"
import { docxodus } from "./docxodus"

// A lawyer's decisions on a counterparty's tracked changes (#200), applied to
// one version's bytes in one Docxodus session. Ids are the reader's
// (dochaus/lib/docx-markup.ts): one per Docxodus unit, in document order, and
// only `supported` units can be decided. A listed id without a decision is held.
//
// - Docxodus 13 derives a unit's id from its part and native carriers
//   (constituentKeys), not its text, so a listed id that still resolves names
//   exactly the listed unit. Each change is decided by that id. One whose id is
//   gone (a carrier consumed, or carriers merged or split) is found after the
//   others by the carriers it was listed with. A listing costs dozens of
//   decisions, so there is none per change.
// - A unit with no carrier left was resolved by another decision, such as a
//   deletion inside an insertion that was rejected: `consumed`, whether it was
//   decided or held.
// - Holds are not touched: they stay pending with their own author and date.
// - Resolving one change can join two others into one (Docxodus merges adjacent
//   same-author changes of one type once the text between them goes). A decided
//   change caught that way would take a held or differently decided one with it,
//   so it is decided before the rest in a fresh session from the same bytes. One
//   such promotion per change: a promoted change caught again refuses.
// - Any refusal refuses the whole set; nothing is returned to write. The session
//   writes no anchor ids into the counterparty's file.
export type MarkupDecision = "accept" | "reject" | "hold"
export type MarkupOutcome = "applied" | "consumed" | "held"
export type MarkupDecisionErrorCode = "unknown_revision" | "unsupported_revision" | "decision_failed"

export class MarkupDecisionError extends Error {
  /** `mergedWith`: the listed changes the refused one became one with. */
  constructor(readonly code: MarkupDecisionErrorCode, message: string, readonly revision: string, readonly mergedWith: readonly string[] = []) {
    super(message)
  }
}

type Docxodus = Awaited<ReturnType<typeof docxodus>>
type Applied = { bytes: Uint8Array; outcomes: { id: string; decision: MarkupDecision; outcome: MarkupOutcome }[] }

export async function applyMarkupDecisions(bytes: Uint8Array, decisions: ReadonlyMap<string, MarkupDecision>) {
  const dx = await docxodus()
  let first: string[] = []
  for (;;) {
    const result = decideInOrder(dx, bytes, decisions, first)
    if (!("merged" in result)) return result
    first = [...result.merged, ...first]
  }
}

/** One pass over the decisions, `first` ahead of the rest in document order.
 * Returns the decided changes that merged with another, to promote, unless one
 * of them was promoted already. */
function decideInOrder(dx: Docxodus, bytes: Uint8Array, decisions: ReadonlyMap<string, MarkupDecision>, first: readonly string[]): Applied | { merged: string[] } {
  const session = dx.openDocxSession(bytes, { persistAnchorIds: false })
  try {
    const listed = session.listRevisions()
    const byId = new Map(listed.map(entry => [entry.id, entry]))
    for (const id of decisions.keys())
      if (!byId.has(id))
        throw new MarkupDecisionError("unknown_revision", "A decided change is not in this version of the document. Refresh the list and decide again.", id)
    const decision = (entry: RevisionListEntry) => decisions.get(entry.id) ?? "hold"
    const unsupported = listed.find(entry => decision(entry) !== "hold" && entry.resolutionStatus !== "supported")
    if (unsupported)
      throw new MarkupDecisionError("unsupported_revision", "One of the decided changes cannot be resolved here. Hold it and resolve it in Word.", unsupported.id)
    const promoted = new Set(first)
    const order = [...first.map(id => byId.get(id)!), ...listed.filter(entry => !promoted.has(entry.id))]
    const outcomes = new Map<string, MarkupOutcome>()
    const resolve = (entry: RevisionListEntry, id: string) => decision(entry) === "accept" ? session.acceptRevision(id) : session.rejectRevision(id)
    const failed = (entry: RevisionListEntry, code = "refused") =>
      new MarkupDecisionError("decision_failed", `A change could not be ${decision(entry)}ed (${code}). Nothing was changed; resolve it in Word.`, entry.id)
    // Each by its listed id; a change whose id is gone waits until the others are done.
    const moved: RevisionListEntry[] = []
    for (const entry of order) {
      if (decision(entry) === "hold") continue
      const result = resolve(entry, entry.id)
      if (result.success) outcomes.set(entry.id, "applied")
      else if (result.error?.code === "revision_not_found") moved.push(entry)
      else throw failed(entry, result.error?.code)
    }
    // The session's listing, kept until its next change: a carrier gone never
    // comes back, so consumed changes share one listing.
    let live: RevisionListEntry[] | undefined
    const merged: string[] = []
    for (const entry of moved) {
      const unit = liveUnit(live ??= session.listRevisions(), entry)
      if (!unit) outcomes.set(entry.id, "consumed")
      else if ("merged" in unit) {
        if (promoted.has(entry.id))
          throw new MarkupDecisionError("decision_failed", "Two changes became one when another change was resolved. Decide them the same way, or resolve them in Word.", entry.id, joinedWith(listed, entry, unit.merged))
        merged.push(entry.id)
      } else {
        const result = resolve(entry, unit.id)
        if (!result.success) throw failed(entry, result.error?.code)
        outcomes.set(entry.id, "applied")
        live = undefined
      }
    }
    if (merged.length) return { merged }
    for (const entry of listed)
      if (decision(entry) === "hold") outcomes.set(entry.id, liveUnit(live ??= session.listRevisions(), entry) ? "held" : "consumed")
    return { bytes: session.save(), outcomes: listed.map(entry => ({ id: entry.id, decision: decision(entry), outcome: outcomes.get(entry.id)! })) }
  } finally {
    session.close()
  }
}

/** The live unit holding `entry`'s carriers now; undefined when none is left.
 * Carriers that another unit's have joined, or that split across units, are a
 * merge: resolving them would resolve another change too. */
function liveUnit(live: RevisionListEntry[], entry: RevisionListEntry) {
  const keys = new Set(entry.constituentKeys)
  const found = live.filter(row => row.partUri === entry.partUri && row.constituentKeys.some(key => keys.has(key)))
  if (found.length > 1 || found.some(row => row.constituentKeys.some(key => !keys.has(key)))) return { merged: found }
  return found[0]
}

/** The other listed changes whose carriers are in the `merged` units. */
function joinedWith(listed: RevisionListEntry[], entry: RevisionListEntry, merged: RevisionListEntry[]) {
  const keys = new Set(merged.flatMap(row => row.constituentKeys))
  return listed.filter(other => other.id !== entry.id && other.partUri === entry.partUri && other.constituentKeys.some(key => keys.has(key))).map(other => other.id)
}
