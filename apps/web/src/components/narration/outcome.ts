import type { LeanDraftingStatus } from "../../api/ingest/lean"
import { attentionItems, leanContinueAvailable, openQuestions } from "../drafting/model"
import { stoppedByTechnicalProblem } from "./activity"

// What the lawyer needs to know once the run settles or stalls, in one short
// line, and the honest checks behind a completed run. Built only from the
// drafting status. It never answers a question or continues the run; the chat's
// header and questions card are where the lawyer acts.

export type RunOutcome = { tone: "attention" | "done"; text: string }

export function runOutcome(status: LeanDraftingStatus): RunOutcome | undefined {
  const questions = openQuestions(status).length
  const asks = questions === 1 ? "a question" : `${questions} questions`
  if (status.status === "completed")
    return {
      tone: "done",
      text: `Drafting is complete.${questions ? ` ${questions === 1 ? "One question still awaits" : `${questions} questions still await`} your answer.` : ""}`,
    }
  if (questions) return { tone: "attention", text: `Drafting is waiting on ${asks} only you can answer.` }
  const resume = leanContinueAvailable(status) ? " Continue drafting (above) picks up from the last finished step." : ""
  if (status.status === "blocked") {
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

export type ResultCheck = { tone: "pass" | "warn" | "neutral"; text: string }

/** The checks a completed run can honestly claim, from its status alone: the
 * final check, quotations verified against the sources (every misquote the
 * checks find becomes a quote item), and what remains for the lawyer. */
export function resultChecks(status: LeanDraftingStatus): ResultCheck[] {
  const items = attentionItems(status).filter((item) => item.kind !== "question")
  const misses = items.filter((item) => item.kind === "quote_miss").length
  const placeholders = items.filter((item) => item.kind === "placeholder").length
  const other = items.length - misses - placeholders
  const checks: Array<ResultCheck | undefined> = [
    status.gate?.status === "pass" ? { tone: "pass", text: "Final workflow check passed" } : undefined,
    status.gate && !misses ? { tone: "pass", text: "Quotes checked against sources" } : undefined,
    misses
      ? { tone: "warn", text: `${misses} quote${misses === 1 ? "" : "s"} not found word for word` }
      : undefined,
    placeholders
      ? { tone: "neutral", text: `${placeholders} placeholder${placeholders === 1 ? "" : "s"} to fill` }
      : undefined,
    other
      ? { tone: "neutral", text: `${other} other item${other === 1 ? "" : "s"} in the drafting notes` }
      : undefined,
  ]
  return checks.filter((check) => check !== undefined)
}
