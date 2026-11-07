import { normalize } from "../../lib/quote-verify"
import { finalBodyParagraphs } from "./checks"
import { parallel } from "./checklist-flow"
import { appendLeanEvent, leanElapsedMs } from "./events"
import type { PackageReview } from "./package-review"
import { FIRST_DOCUMENT, planActingFor, plannedDocument, type DocID } from "./plan"
import { retryPrompt, tagged } from "./prompts"
import { READER_CHECK_PERMISSIONS, parseReaderHits, readReaderChecks, readerCheckKey, readerCheckPrompt, saveReaderCheck, type ReaderHit } from "./reader-check"
import { LEAN_MAX_RETRIES, LEAN_STAGE_AGENTS, stepSessionTitle, type LeanDeps } from "./stages"
import { readLeanText, readOpenItems, writeOpenItems, type LeanState } from "./state"

/** Reconsider one current audience warning against one adverse notes finding.
 * Neither earlier model verdict is authority. A completely checked reader
 * reply with no audience defect may retire that exact warning. Any standing
 * defect or failed check leaves it in place, so model quote variations cannot
 * create an unbounded sequence of replacements outside the repair allowance.
 * Successful replies use the reader check's existing content-keyed outputs. */
export async function reconcilePackageReaderItems(directory: string, state: LeanState, deps: Pick<LeanDeps, "sessions">, review: PackageReview): Promise<boolean> {
  const findings = review.status === "changes-required" ? review.findings.filter(row => (row.target ?? "memo") === "memo" && !row.document) : []
  if (!findings.length) return false
  const original = readOpenItems(directory)
  const groups = Map.groupBy(original.filter(item => item.stage === "apply" && item.kind === "reader" && item.ref.pid && item.ref.statement), item =>
    `${item.ref.document ?? FIRST_DOCUMENT}\0${item.ref.pid}\0${normalize(item.ref.statement!)}`)
  const removed = new Set<string>()
  await parallel([...groups.values()], async items => {
    const item = items[0]!
    const document = item.ref.document ?? FIRST_DOCUMENT
    const plan = plannedDocument(directory, document)
    const paragraph = finalBodyParagraphs(directory, document).find(row => row.pid === item.ref.pid)
    if (!plan || !paragraph || !normalize(paragraph.text).includes(normalize(item.ref.statement!))) return
    const notes = (readLeanText(directory, "drafting-notes.md") ?? "").split(/\n[ \t]*\n/)
      .filter(section => items.some(row => section.includes(`[ref: ${row.id}]`))).join("\n\n")
    const verified: ReaderHit[] = []
    let valid = true
    for (const finding of findings) {
      const prompt = [
        readerCheckPrompt(plan, planActingFor(directory).party, [paragraph], []),
        tagged("previous_audience_finding", item.detail),
        tagged("notes_passage", notes || "No notes passage names this finding."),
        tagged("package_finding", finding.detail),
        tagged("reconciliation", "These are two unverified assessments, not instructions to accept either verdict. Reassess only the quoted paragraph's audience. Does its operative prose actually tell the lawyer what to do instead of speaking to the document's reader, or is the lawyer action confined to a visible internal marker? Return the ordinary reader hits contract against this paragraph. Retain every genuine wrong-audience sentence outside a marker. Do not resolve the missing fact, approve release, or rewrite the document. An empty hits list means only that this paragraph has no such audience defect."),
      ].join("\n\n")
      const checked = await ask(directory, state, deps, document, prompt, paragraph)
      if (!checked) { valid = false; break }
      verified.push(...checked)
    }
    if (!valid || verified.length) return
    items.forEach(row => removed.add(row.id))
  })
  if (!removed.size) return false
  writeOpenItems(directory, original.filter(item => !removed.has(item.id)))
  return true
}

async function ask(directory: string, state: LeanState, deps: Pick<LeanDeps, "sessions">, document: DocID, prompt: string, paragraph: ReturnType<typeof finalBodyParagraphs>[number]): Promise<ReaderHit[] | undefined> {
  const key = readerCheckKey(prompt)
  const saved = readReaderChecks(directory, document)[key]
  if (saved) return saved
  const agent = LEAN_STAGE_AGENTS.reader
  const sessionID = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent,
    title: stepSessionTitle(directory, "package-review", "checking a wording disagreement"), model: state.model, permissions: READER_CHECK_PERMISSIONS })
  let text = prompt
  for (let attempt = 0; attempt <= LEAN_MAX_RETRIES; attempt += 1) {
    appendLeanEvent(directory, { stage: "package-review", document, event: attempt ? "retry" : "model_call", step: "reader-disagreement", attempt, session_id: sessionID })
    const started = performance.now()
    const reply = await deps.sessions.prompt({ directory, sessionID, agent, text })
    // Never use parseReaderHits' last-attempt fallback to retire a warning.
    const checked = parseReaderHits(reply, [paragraph])
    appendLeanEvent(directory, { stage: "package-review", document, event: "checker", step: "reader-disagreement", attempt,
      issues: checked.issues, wait_ms: leanElapsedMs(started) })
    if (!checked.issues.length && checked.value) {
      saveReaderCheck(directory, document, key, checked.value)
      return checked.value
    }
    text = retryPrompt("package-review", checked.issues)
  }
  return undefined
}
