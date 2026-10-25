import type { CounterpartyItem } from "../../lib/docx-markup"
import { readChangePlan, type ChangeGroup } from "./change-pipeline"
import { readChangeList, type ApplyResult } from "./checks"
import { readLeanPlan, type DocID } from "./plan"
import { REQUEST_SOURCE, matterSourceIDs, type SourceQuote } from "./quotes"
import { approvedResearch } from "./research"
import { documentOrder, findingRoute, heldBackReason, holdsBeyondMarket, raisedFindings, type FindingOrigin, type FindingTarget } from "./review-findings"
import { readLeanJson, readQuestions, type ReviewStageID } from "./state"

// The issue register: a view over the run's own records, built beside the
// final gate. Each review finding (review-findings.ts) is listed with the
// document it concerns, where it was raised, its origin and basis, and the
// disposition its change plan gave it. An amendment resolves the findings of
// the group whose paragraph it amends; apply.json gives its outcome. Nothing
// is recorded for the register itself, and it carries no text into delivery.

export type RegisterDisposition = "amended" | "not_amended" | "deferred" | "recommendation" | "rejected" | "question" | "earlier_document" | "memo" | "held_back" | "unadjudicated"
export type RegisterEntry = {
  id: string
  document: FindingTarget
  raised_in: DocID
  stage: ReviewStageID | "counterparty"
  pids?: string[]
  origin: FindingOrigin | "counterparty"
  basis: "authority" | "source" | "practice"
  finding: string
  instead?: string
  /** For the other side's change or comment, the item it is, by source version and markup id. */
  counterparty?: CounterpartyItem
  /** amended when an amendment for it was applied; not_amended when its group's
   * amendments were withheld, unapplied or never made; earlier_document when it
   * concerns a document already final, disclosed in the notes; held_back when
   * the run's Posture held it back as beyond market, disclosed in the notes. */
  disposition: RegisterDisposition
  /** For a finding a group resolves, each amendment of the group's paragraphs and what apply did with it. */
  amendments: { id: string; pid: string; outcome: "applied" | "unapplied" | "withheld" | "pending" }[]
  /** For a finding a question resolves, the question and whether it is answered. */
  question?: { id: string; status: string }
  /** For a finding a disposition resolves, its reason. */
  reason?: string
}

export function issueRegister(directory: string): { entries: RegisterEntry[] } {
  const plans = new Map((readLeanPlan(directory)?.documents ?? []).map(row => [row.id as string, readChangePlan(directory, row.id)]))
  const questions = new Map(readQuestions(directory).map(row => [row.id, row]))
  const order = documentOrder(directory)
  const holds = holdsBeyondMarket(directory)
  return { entries: raisedFindings(directory).map(row => {
    const base = { id: row.id, document: row.document, raised_in: row.raised_in, stage: row.stage, ...(row.pids ? { pids: row.pids } : {}),
      origin: row.origin, basis: findingBasis(directory, row.sources), finding: row.finding, ...(row.instead ? { instead: row.instead } : {}),
      ...(row.counterparty ? { counterparty: row.counterparty } : {}), amendments: [] }
    const route = findingRoute(order, row, holds)
    if (route === "held") return { ...base, disposition: "held_back" as const, reason: heldBackReason(directory, row) }
    if (route === "memo") return { ...base, disposition: "memo" as const }
    if (route === "earlier") return { ...base, disposition: "earlier_document" as const }
    const target = row.document as DocID
    const adjudication = plans.get(target)
    const resolves = (entry: { resolves?: string[] }) => entry.resolves?.includes(row.id)
    const group = [...(adjudication?.groups ?? []), ...(adjudication?.settled_groups ?? [])].find(resolves)
    if (group) {
      const amendments = groupAmendments(directory, target, group)
      return { ...base, disposition: amendments.some(entry => entry.outcome === "applied") ? "amended" as const : "not_amended" as const, amendments }
    }
    const question = adjudication?.questions.find(resolves)
    if (question) return { ...base, disposition: "question" as const, question: { id: question.id, status: questions.get(question.id)?.status ?? question.status } }
    const disposition = adjudication?.findings?.find(resolves)
    if (disposition) return { ...base, disposition: disposition.disposition, reason: disposition.reason }
    return { ...base, disposition: "unadjudicated" as const }
  }) }
}

/** What a finding rests on, from its sources: approved authority, a matter
 * source or the lawyer's words, or otherwise general practice. */
function findingBasis(directory: string, sources: SourceQuote[]): RegisterEntry["basis"] {
  const authority = new Set(approvedResearch(directory).map(row => row.url))
  if (sources.some(source => authority.has(source.source_id))) return "authority"
  const matter = new Set([...matterSourceIDs(directory), REQUEST_SOURCE])
  return sources.some(source => matter.has(source.source_id)) ? "source" : "practice"
}

/** The amendments of a group's paragraphs, with their outcome at apply. */
function groupAmendments(directory: string, document: DocID, group: ChangeGroup): RegisterEntry["amendments"] {
  const pids = new Set(group.pids)
  const changes = readChangeList(directory, document)
  const applied = readLeanJson(directory, `${document}/apply.json`) as ApplyResult | undefined
  const outcome = (id: string) => applied?.applied.some(row => row.id === id) ? "applied" as const : applied?.unapplied.some(row => row.id === id) ? "unapplied" as const : "pending" as const
  return [
    ...(changes?.amendments ?? []).filter(row => pids.has(row.pid)).map(row => ({ id: row.id, pid: row.pid, outcome: outcome(row.id) })),
    ...(changes?.withheld ?? []).filter(row => pids.has(row.amendment.pid)).map(row => ({ id: row.amendment.id, pid: row.amendment.pid, outcome: "withheld" as const })),
  ]
}
