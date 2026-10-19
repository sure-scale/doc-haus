import type { LeanDraftingStatus } from "../../api/ingest/lean"
import {
  afterCheckChange,
  attentionItems,
  checklistFindingsCount,
  delegableQuestion,
  findingsConsequence,
  leanContinueAvailable,
  leanRunStopped,
  notesWording,
  openQuestions,
  unresolvedFindingsPhrase,
} from "../drafting/model"
import { stoppedByTechnicalProblem } from "./activity"

// What the lawyer needs to know once the run settles or stalls, in one short
// line, and the honest checks behind a completed run. Built only from the
// drafting status. It never answers a question or continues the run; the chat's
// header and questions card are where the lawyer acts.

export type RunOutcome = { tone: "attention" | "done"; text: string }

export function runOutcome(status: LeanDraftingStatus): RunOutcome | undefined {
  const questions = openQuestions(status).length
  const asks = questions === 1 ? "a question" : `${questions} questions`
  if (status.status === "completed") {
    const pending = questions ? ` ${questions === 1 ? "One question still awaits" : `${questions} questions still await`} your answer.` : ""
    // Delivered, but with checklist requirements the draft does not satisfy.
    const findings = checklistFindingsCount(status)
    if (findings !== undefined)
      return {
        tone: "done",
        text: `Drafting delivered a review draft with ${unresolvedFindingsPhrase(findings)}. ${findingsConsequence(findings)}${pending}`,
      }
    return { tone: "done", text: `Drafting is complete.${pending}` }
  }
  if (leanRunStopped(status))
    return { tone: "attention", text: "You stopped this draft. Its work so far is kept; a new drafting request starts a new draft." }
  if (questions) return { tone: "attention", text: `Drafting is waiting on ${asks}${waitingChoice(status, questions)}` }
  const resume = leanContinueAvailable(status) ? " Continue drafting (above) picks up from the last finished step." : ""
  if (status.status === "blocked") {
    if (status.paused_at)
      return { tone: "attention", text: "Paused at your request. Continue drafting (above) picks up from the step that was in progress." }
    const atGate = status.stages.gate?.status === "blocked"
    const blocking = (atGate ? status.gate?.blocking.length : undefined) ?? status.open_items.filter((item) => item.status === "open").length
    const items = `${blocking} open item${blocking === 1 ? "" : "s"}`
    if (atGate)
      return {
        tone: "attention",
        text: `Paused at the final check with ${items}. Continue drafting (above) runs a repair round on them.`,
      }
    return { tone: "attention", text: `Paused with ${items}.${resume}` }
  }
  if (stoppedByTechnicalProblem(status))
    return {
      tone: "attention",
      text: "Paused after a technical problem. The work so far is kept; Continue drafting (above) picks up from the last finished step.",
    }
  if (status.status === "failed") return { tone: "attention", text: `Drafting stopped before it finished.${resume}` }
  if (!status.in_flight) return { tone: "attention", text: `Paused: nothing is working on this run right now.${resume}` }
  return undefined
}

/** How the lawyer moves a question on: answer it, or, where Decide for me is
 * offered, accept the recommendation. */
function waitingChoice(status: LeanDraftingStatus, questions: number) {
  const delegable = openQuestions(status).filter(delegableQuestion).length
  if (!delegable) return " only you can answer."
  if (questions === 1) return ": answer it, or use Decide for me to accept the recommendation."
  const some = delegable === questions ? "any of them" : delegable === 1 ? "one of them" : `${delegable} of them`
  return `: answer each, or use Decide for me to accept the recommendation on ${some}.`
}

export type ResultCheck = { tone: "pass" | "warn" | "neutral"; text: string }

/** The checks a completed run can honestly claim, from its status alone: the
 * final check, quotations verified against the sources (every misquote the
 * checks find becomes a quote item), and what remains for the lawyer. A draft
 * changed or missing after the final check is not what it checked, so that
 * pass is not claimed; nor is it for a draft delivered with unresolved
 * checklist findings. */
export function resultChecks(status: LeanDraftingStatus): ResultCheck[] {
  const items = attentionItems(status).filter((item) => item.kind !== "question")
  const misses = items.filter((item) => item.kind === "quote_miss").length
  const placeholders = items.filter((item) => item.kind === "placeholder").length
  const unverified = items.filter((item) => item.kind === "checklist_open").length
  const unresolved = items.filter((item) => item.kind === "checklist_unresolved").length
  const other = items.length - misses - placeholders - unverified - unresolved
  const findings = checklistFindingsCount(status)
  const after = afterCheckChange(status)
  const checks: Array<ResultCheck | undefined> = [
    status.gate?.status === "pass"
      ? after
        ? { tone: "warn", text: `${after.label}: ${after.detail}` }
        : findings === undefined
          ? { tone: "pass", text: "Final workflow check passed" }
          : undefined
      : undefined,
    findings !== undefined
      ? {
          tone: "warn",
          text: findings ? `${findings} checklist requirement${findings === 1 ? "" : "s"} not satisfied` : "Checklist requirements not satisfied",
        }
      : undefined,
    status.gate && !misses ? { tone: "pass", text: "Quotes checked against sources" } : undefined,
    misses
      ? { tone: "warn", text: `${misses} quote${misses === 1 ? "" : "s"} not found word for word` }
      : undefined,
    placeholders
      ? { tone: "neutral", text: `${placeholders} placeholder${placeholders === 1 ? "" : "s"} to fill` }
      : undefined,
    unverified
      ? { tone: "warn", text: `${unverified} requirement${unverified === 1 ? "" : "s"} left open by the checklist check` }
      : undefined,
    other
      ? { tone: "neutral", text: `${other} other item${other === 1 ? "" : "s"} in ${notesWording(status.memo).phrase}` }
      : undefined,
  ]
  return checks.filter((check) => check !== undefined)
}
