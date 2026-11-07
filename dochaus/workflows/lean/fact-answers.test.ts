import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import path from "node:path"
import { finalGate } from "./gate"
import { checkAnswerClass, classifyAnswers, refreshFactAnswers } from "./answer-class"
import { factContext, factDisposition, factOpenItems, waitingFactQuestions, withFactOptions } from "./fact-answers"
import { writeLeanPlan } from "./plan"
import { openItem } from "./quotes"
import { answerQuestion, answeredQuestionsText, initLeanState, questionContent, readQuestions, unclassifiedAnswer, writeLeanText, writeOpenItems, writeQuestions, type Question } from "./state"
import { cleanupMatters, matter, scriptedModel, AGENT, SOURCE } from "./test-support"

afterEach(cleanupMatters)

const FACT: Question = { id: "q-inspection", stage: "intake", kind: "fact", status: "answered", question: "What was the inspection date?", context: "The notice needs the inspection date.", answer: "Continue with the date marked unknown.", fact_status: "provisional" }

test("an answered fact left provisional does not resolve its uncertainty at delivery", () => {
  const directory = matter()
  writeQuestions(directory, [FACT])
  const item = openItem("intake", "question", "Inspection date remains unknown.", { question_id: FACT.id, placeholder: "[open item: confirm the inspection date]" })
  writeOpenItems(directory, [item])
  writeLeanText(directory, "drafting-notes.md", `# Notes\n\nConfirm the inspection date [ref: ${item.id}].\n`)
  const found = () => finalGate(directory).items.find(row => row.id === item.id)
  expect(found()).toMatchObject({ status: "open" })
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p-inspection", sha256: "0", text: "Inspection date: [open item: confirm the inspection date]." }]))
  expect(found()).toMatchObject({ status: "disclosed" })
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect(answeredQuestionsText(directory)).toContain("fact remains unknown")
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p-inspection", sha256: "0", text: "Inspection date: [open item: confirm the unrelated address]." }]))
  expect(found()).toMatchObject({ status: "open" })
  writeQuestions(directory, [{ ...FACT, fact_status: "wait" }])
  expect(found()).toMatchObject({ status: "open" })
  writeQuestions(directory, [{ ...FACT, answer: "The inspection was on 8 August.", fact_status: "supplied" }])
  expect(found()).toMatchObject({ status: "resolved" })
})

test("fact choices govern workflow treatment without pretending to supply a fact", () => {
  const directory = matter()
  const question = withFactOptions({ ...FACT, status: "open", answer: undefined, fact_status: undefined })
  expect(question.options?.map(option => option.fact_action)).toEqual(["provisional", "wait"])
  writeQuestions(directory, [question])
  const provisional = answerQuestion(directory, question.id, { option_id: "o1" })
  expect(factDisposition(provisional)).toBe("provisional")
  expect(unclassifiedAnswer(provisional)).toBe(false)
  expect(waitingFactQuestions(directory)).toEqual([])
  expect(factOpenItems(directory)).toMatchObject([{ kind: "question", stage: "intake", ref: { question_id: question.id, placeholder: `[open item: ${question.question}]` } }])
  expect(questionContent(provisional)).toHaveProperty("fact_status", "provisional")
  const waiting = answerQuestion(directory, question.id, { option_id: "o2" })
  expect(factDisposition(waiting)).toBe("wait")
  expect(waitingFactQuestions(directory).map(row => row.id)).toEqual([question.id])
  expect(answeredQuestionsText(directory)).toContain("do not draft or ask a status question")
  const replacement = answerQuestion(directory, question.id, { text: "The inspection was on 8 August." })
  expect(replacement).not.toHaveProperty("fact_status")
  expect(unclassifiedAnswer(replacement)).toBe(true)
})

test("existing option identities and non-fact questions remain unchanged", () => {
  const existing: Question = { ...FACT, advice: { options: [{ id: "o1", text: "An earlier offered choice" }], unavailable: { reason: "Check the evidence." } } }
  expect(withFactOptions(existing)).toBe(existing)
  const judgment = { ...FACT, kind: "judgment" as const }
  expect(withFactOptions(judgment)).toBe(judgment)
})

test("fact classification needs quoted human evidence and never promotes delegation into a supplied fact", () => {
  const supplied = { ...FACT, answer: "The inspection was on 8 August." }
  expect(checkAnswerClass(JSON.stringify({ class: "answered", fact_status: "supplied", fact_quote: supplied.answer }), supplied)).toEqual({ issues: [], answer_class: "answered", fact_status: "supplied" })
  expect(checkAnswerClass(JSON.stringify({ class: "answered", fact_status: "supplied", fact_quote: "A different date" }), supplied).issues).not.toEqual([])
  expect(checkAnswerClass(JSON.stringify({ class: "delegated", fact_status: "supplied", fact_quote: supplied.answer }), supplied).issues).not.toEqual([])
  expect(checkAnswerClass(JSON.stringify({ class: "answered" }), supplied).issues).not.toEqual([])
})

test("legacy external-task answers stay unknown until the lawyer authorises treatment or supplies evidence", async () => {
  const directory = matter()
  const state = initLeanState(directory, { request: "Draft a notice.", ownerSessionID: "ses_owner", model: null })
  const old: Question = { ...FACT, answer: "Ask the project manager to obtain the inspection report.", fact_status: undefined, answer_kind: "option", option_id: "o1" }
  writeQuestions(directory, [old])
  const model = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ class: "answered", fact_status: "unknown" }) })
  await classifyAnswers(directory, state, { sessions: model.port })
  expect(readQuestions(directory)[0]).toMatchObject({ answer: old.answer, answer_kind: "option", fact_status: "unknown" })
  expect(waitingFactQuestions(directory)).toHaveLength(1)
  const count = model.prompts.length
  await classifyAnswers(directory, state, { sessions: model.port })
  expect(model.prompts).toHaveLength(count)
  answerQuestion(directory, old.id, { text: "The inspection was on 8 August." })
  const confirmed = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ class: "answered", fact_status: "supplied", fact_quote: "The inspection was on 8 August." }) })
  await classifyAnswers(directory, state, { sessions: confirmed.port })
  expect(factOpenItems(directory)).toEqual([])
  expect(waitingFactQuestions(directory)).toEqual([])
})

test("unreadable fact classification preserves the unknown instead of granting drafting permission", async () => {
  const directory = matter()
  const state = initLeanState(directory, { request: "Draft a notice.", ownerSessionID: "ses_owner", model: null })
  writeQuestions(directory, [{ ...FACT, fact_status: undefined, answer: "Ask for the report." }])
  const unreadable = scriptedModel({ [AGENT.answer]: () => "invalid" })
  await classifyAnswers(directory, state, { sessions: unreadable.port })
  expect(readQuestions(directory)[0]).toMatchObject({ fact_status: "unknown" })
  expect(waitingFactQuestions(directory)).toHaveLength(1)
})

test("a runtime error before classification cannot let an unread fact pass as supplied", async () => {
  const directory = matter()
  const state = initLeanState(directory, { request: "Draft a notice.", ownerSessionID: "ses_owner", model: null })
  writeQuestions(directory, [{ ...FACT, fact_status: undefined, answer: "Ask for the report." }])
  const interrupted = scriptedModel({ [AGENT.answer]: () => { throw new Error("The answer session stopped.") } })
  await classifyAnswers(directory, state, { sessions: interrupted.port })
  expect(factDisposition(readQuestions(directory)[0]!)).toBeUndefined()
  expect(unclassifiedAnswer(readQuestions(directory)[0]!)).toBe(true)
  expect(waitingFactQuestions(directory)).toHaveLength(1)
  const items = factOpenItems(directory)
  expect(items).toHaveLength(1)
  writeOpenItems(directory, items)
  writeLeanText(directory, "drafting-notes.md", `# Notes\n\nConfirm the fact [ref: ${items[0]!.id}].\n`)
  expect(finalGate(directory).items.find(item => item.id === items[0]!.id)).toMatchObject({ status: "open" })
})

function waitingMatter() {
  const directory = matter()
  const state = initLeanState(directory, { request: "Draft a notice.", ownerSessionID: "ses_owner", model: null })
  writeQuestions(directory, [withFactOptions({ ...FACT, status: "open", answer: undefined, fact_status: undefined })])
  answerQuestion(directory, FACT.id, { option_id: "o2" }, factContext(directory))
  return { directory, state }
}

function updateAfterAnswer(directory: string, text: string, offset = 1) {
  const at = Date.parse(readQuestions(directory)[0]!.answered_at!) + offset
  writeLeanText(directory, "updated-notes.md", `<!-- update:${String(at).padStart(15, "0")}-000001-synthetic -->\n${text}\n`)
}

test("a later express drafting instruction reopens a wait once, preserving the original answer", async () => {
  const { directory, state } = waitingMatter()
  const instruction = "You may now prepare the review draft with the inspection date visibly marked as unknown."
  updateAfterAnswer(directory, instruction)
  const model = scriptedModel({ [AGENT.answer]: prompt => {
    expect(prompt).toContain("Reconsider ONE pending factual uncertainty")
    expect(prompt).toContain(instruction)
    return JSON.stringify({ fact_status: "provisional", reason: "The later instruction expressly permits a marked review draft.", basis: [{ source_id: "updated-notes", quote: instruction }] })
  } })
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(waitingFactQuestions(directory)).toEqual([])
  expect(readQuestions(directory)[0]).toMatchObject({ fact_status: "provisional", answer: expect.stringContaining("Wait for me"), fact_resolution: { reason: expect.stringContaining("later instruction") } })
  expect(answeredQuestionsText(directory)).toContain(instruction)
  expect(factOpenItems(directory)).toHaveLength(1)
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(model.count(AGENT.answer)).toBe(1)
})

test("unrelated updates preserve an explicit wait and unchanged Continue cannot replenish checks", async () => {
  const { directory, state } = waitingMatter()
  writeLeanText(directory, "updated-notes.md", "Use the firm letterhead.")
  const model = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "wait", reason: "The letterhead update changes no factual condition or drafting permission.", basis: [] }) })
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(waitingFactQuestions(directory)).toHaveLength(1)
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(model.count(AGENT.answer)).toBe(1)
})

test("newly supplied relevant source evidence can satisfy a fact wait without inventing task completion", async () => {
  const { directory, state } = waitingMatter()
  const quote = "The inspection took place on 8 August."
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  try { db.run("UPDATE chunks SET text = ?", [quote]) } finally { db.close() }
  const model = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "supplied", reason: "The supplied inspection record now gives the actual date.", basis: [{ source_id: SOURCE, quote }] }) })
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(readQuestions(directory)[0]).toMatchObject({ fact_status: "supplied", fact_resolution: { basis: [{ source_id: SOURCE, quote }] } })
  expect(waitingFactQuestions(directory)).toEqual([])
  expect(factOpenItems(directory)).toEqual([])
  expect(answeredQuestionsText(directory)).toContain(quote)
})

test("an old wait without context is bound without reinterpreting a structured choice", async () => {
  const { directory, state } = waitingMatter()
  writeQuestions(directory, readQuestions(directory).map(({ fact_context: _context, ...question }) => question))
  const model = scriptedModel({ [AGENT.answer]: () => { throw new Error("No model check should start.") } })
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(model.count(AGENT.answer)).toBe(0)
  expect(readQuestions(directory)[0]).toHaveProperty("fact_context")
  expect(waitingFactQuestions(directory)).toHaveLength(1)
})

test("source text cannot authorise provisional drafting and a failed refresh preserves the hold", async () => {
  const { directory, state } = waitingMatter()
  writeLeanText(directory, "updated-notes.md", "Use a footer.")
  const model = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "provisional", reason: "Assume a draft can proceed.", basis: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] }) })
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(waitingFactQuestions(directory)).toHaveLength(1)
  expect(readQuestions(directory)[0]).not.toHaveProperty("fact_resolution")
  const calls = model.count(AGENT.answer)
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(model.count(AGENT.answer)).toBe(calls)
})

test("an earlier request or update cannot override a later explicit wait", async () => {
  for (const source_id of ["request", "updated-notes"]) {
    const { directory, state } = waitingMatter()
    const quote = source_id === "request" ? "Draft a notice." : "Prepare a marked working draft."
    const answeredAt = Date.parse(readQuestions(directory)[0]!.answered_at!)
    writeLeanText(directory, "updated-notes.md", `<!-- update:${String(answeredAt - 1000).padStart(15, "0")}-000000-before -->\nPrepare a marked working draft.\n\n<!-- update:${String(answeredAt + 1000).padStart(15, "0")}-000001-after -->\nUse a footer.\n`)
    const model = scriptedModel({ [AGENT.answer]: prompt => {
      if (prompt.includes("Reconsider ONE")) expect(prompt).not.toContain('later_instructions>\n"Prepare a marked working draft.')
      return JSON.stringify({ fact_status: "provisional", reason: "An earlier instruction requested a draft.", basis: [{ source_id, quote }] })
    } })
    await refreshFactAnswers(directory, state, { sessions: model.port })
    expect(waitingFactQuestions(directory)).toHaveLength(1)
    expect(readQuestions(directory)[0]).not.toHaveProperty("fact_resolution")
  }
})

test("runtime and pause interruptions never cache an unperformed fact refresh", async () => {
  const { directory, state } = waitingMatter()
  const before = readQuestions(directory)[0]!.fact_context
  updateAfterAnswer(directory, "The inspection evidence is still being obtained.")
  const interrupted = scriptedModel({ [AGENT.answer]: () => { throw new Error("The provider connection stopped.") } })
  await expect(refreshFactAnswers(directory, state, { sessions: interrupted.port })).rejects.toThrow("provider connection stopped")
  expect(readQuestions(directory)[0]!.fact_context).toBe(before)
  const controller = new AbortController()
  const paused = scriptedModel({ [AGENT.answer]: () => {
    controller.abort()
    return JSON.stringify({ fact_status: "wait", reason: "No evidence supplied.", basis: [] })
  } })
  await refreshFactAnswers(directory, state, { sessions: paused.port, pause: controller.signal })
  expect(readQuestions(directory)[0]!.fact_context).toBe(before)
  expect(waitingFactQuestions(directory)).toHaveLength(1)
})

test("an unscoped fact needed only in a later planned document is disclosed there; scoped facts stay scoped", () => {
  const directory = matter()
  writeLeanPlan(directory, [
    { title: "Cover note", purpose: "Covering note", file: "cover.docx", parts: [{ heading: "Cover", scope: "Covering note" }] },
    { title: "Inspection notice", purpose: "Record the inspection", file: "notice.docx", parts: [{ heading: "Inspection", scope: "Inspection terms" }] },
  ])
  writeQuestions(directory, [FACT])
  const [item] = factOpenItems(directory)
  writeOpenItems(directory, [item!])
  writeLeanText(directory, "drafting-notes.md", `# Notes\n\nInspection date requires confirmation [ref: ${item!.id}].`)
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p-cover", sha256: "0", text: "Please review the attached notice." }]))
  const body = (text: string) => writeLeanText(directory, "d2/paragraphs-applied.json", JSON.stringify([{ pid: "p-inspection", sha256: "0", text }]))
  body(`Inspection date: ${item!.ref.placeholder}`)
  const outcome = () => finalGate(directory).items.find(row => row.id === item!.id)
  expect(outcome()).toMatchObject({ status: "disclosed" })
  body("Inspection date: not stated.")
  expect(outcome()).toMatchObject({ status: "open" })
  body(`Inspection date: ${item!.ref.placeholder}`)
  writeQuestions(directory, [{ ...FACT, document: "d1" }])
  expect(outcome()).toMatchObject({ status: "open" })
})

test("provisional treatment follows new actual evidence or a later hold while unrelated updates preserve it", async () => {
  for (const change of ["unrelated", "source", "hold"] as const) {
    const { directory, state } = waitingMatter()
    answerQuestion(directory, FACT.id, { option_id: "o1" }, factContext(directory))
    const quote = change === "hold" ? "Stop drafting until I confirm the inspection date." : "The inspection took place on 8 August."
    if (change === "source") {
      const db = new Database(path.join(directory, ".dochaus", "legal.db"))
      try { db.run("UPDATE chunks SET text = ?", [quote]) } finally { db.close() }
    } else updateAfterAnswer(directory, change === "hold" ? quote : "Use a footer.")
    const model = scriptedModel({ [AGENT.answer]: () => JSON.stringify(change === "source"
      ? { fact_status: "supplied", reason: "The supplied inspection record gives the date.", basis: [{ source_id: SOURCE, quote }] }
      : change === "hold" ? { fact_status: "wait", reason: "The lawyer now expressly stops further drafting.", basis: [{ source_id: "updated-notes", quote }] }
      : { fact_status: "provisional", reason: "The footer instruction does not change permission to draft provisionally.", basis: [] }) })
    await refreshFactAnswers(directory, state, { sessions: model.port })
    expect(factDisposition(readQuestions(directory)[0]!)).toBe(change === "source" ? "supplied" : change === "hold" ? "wait" : "provisional")
    const calls = model.count(AGENT.answer)
    await refreshFactAnswers(directory, state, { sessions: model.port })
    expect(model.count(AGENT.answer)).toBe(calls)
  }
})

test("answer receipt order preserves a same-millisecond later update and classification cannot swallow it", async () => {
  const directory = matter()
  const state = initLeanState(directory, { request: "Draft a notice.", ownerSessionID: "ses_owner", model: null })
  const before = factContext(directory)
  const at = Date.now() - 10000
  const prefix = String(at).padStart(15, "0")
  const id = `${prefix}-000001-answer`
  const answer = "Wait until I provide the inspection report."
  writeQuestions(directory, [withFactOptions({ ...FACT, status: "open", answer: undefined, fact_status: undefined })])
  answerQuestion(directory, FACT.id, { text: answer }, before, { id, at: new Date(at).toISOString() })
  const instruction = "Continue with the inspection date visibly marked unknown in the review draft."
  writeLeanText(directory, "updated-notes.md", `<!-- update:${prefix}-000002-update -->\n${instruction}\n`)
  const model = scriptedModel({ [AGENT.answer]: prompt => prompt.includes("Reconsider ONE")
    ? JSON.stringify({ fact_status: "provisional", reason: "The later received update permits a marked draft.", basis: [{ source_id: "updated-notes", quote: instruction }] })
    : JSON.stringify({ class: "answered", fact_status: "wait", fact_quote: answer }) })
  await classifyAnswers(directory, state, { sessions: model.port })
  expect(readQuestions(directory)[0]).toMatchObject({ fact_status: "wait", fact_context: before, answer_request_id: id, answered_at: new Date(at).toISOString() })
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(factDisposition(readQuestions(directory)[0]!)).toBe("provisional")
  expect(waitingFactQuestions(directory)).toEqual([])
})

test("removed or replaced supporting evidence cannot preserve a stale source-derived supplied fact", async () => {
  for (const change of ["missing", "malformed", "replacement"] as const) {
    const { directory, state } = waitingMatter()
    const original = "The inspection took place on 8 August."
    const replace = (text: string) => {
      const db = new Database(path.join(directory, ".dochaus", "legal.db"))
      try { db.run("UPDATE chunks SET text = ?", [text]) } finally { db.close() }
    }
    replace(original)
    const supplied = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "supplied", reason: "The supplied report states the date.", basis: [{ source_id: SOURCE, quote: original }] }) })
    await refreshFactAnswers(directory, state, { sessions: supplied.port })
    expect(factDisposition(readQuestions(directory)[0]!)).toBe("supplied")
    const replacement = change === "replacement" ? "The corrected inspection date is 9 August." : "The report contains no inspection date."
    replace(replacement)
    expect(waitingFactQuestions(directory)).toHaveLength(1)
    expect(factOpenItems(directory)).toHaveLength(1)
    const review = scriptedModel({ [AGENT.answer]: () => change === "malformed" ? "invalid" : JSON.stringify(change === "replacement"
      ? { fact_status: "supplied", reason: "The corrected supplied report gives the revised actual date.", basis: [{ source_id: SOURCE, quote: replacement }] }
      : { fact_status: "wait", reason: "The supporting date has been removed and the original evidence condition is not satisfied.", basis: [] }) })
    await refreshFactAnswers(directory, state, { sessions: review.port })
    expect(factDisposition(readQuestions(directory)[0]!)).toBe(change === "replacement" ? "supplied" : "wait")
    if (change !== "replacement") expect(readQuestions(directory)[0]).not.toHaveProperty("fact_resolution")
    else expect(answeredQuestionsText(directory)).toContain(replacement)
    const count = review.count(AGENT.answer)
    await refreshFactAnswers(directory, state, { sessions: review.port })
    expect(review.count(AGENT.answer)).toBe(count)
  }
})

test("a human-supplied factual assertion does not become source-derived because unrelated source text changes", async () => {
  const directory = matter()
  const state = initLeanState(directory, { request: "Draft a notice.", ownerSessionID: "ses_owner", model: null })
  writeQuestions(directory, [{ ...FACT, answer: "I confirm the inspection was on 8 August.", fact_status: "supplied", fact_context: factContext(directory) }])
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  try { db.run("UPDATE chunks SET text = ?", ["New supply terms contain no inspection date."]) } finally { db.close() }
  const model = scriptedModel({ [AGENT.answer]: () => { throw new Error("No source revalidation applies to a human assertion.") } })
  await refreshFactAnswers(directory, state, { sessions: model.port })
  expect(model.count(AGENT.answer)).toBe(0)
  expect(factDisposition(readQuestions(directory)[0]!)).toBe("supplied")
})

test("withdrawing supplied evidence preserves the original permission to draft provisionally", async () => {
  const { directory, state } = waitingMatter()
  answerQuestion(directory, FACT.id, { option_id: "o1" }, factContext(directory))
  const replace = (text: string) => {
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    try { db.run("UPDATE chunks SET text = ?", [text]) } finally { db.close() }
  }
  const quote = "The inspection took place on 8 August."
  replace(quote)
  const supplied = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "supplied", reason: "The supplied report gives the inspection date.", basis: [{ source_id: SOURCE, quote }] }) })
  await refreshFactAnswers(directory, state, { sessions: supplied.port })
  expect(factDisposition(readQuestions(directory)[0]!)).toBe("supplied")
  replace("The supplied material no longer records an inspection date.")
  expect(waitingFactQuestions(directory)).toEqual([])
  expect(factOpenItems(directory)).toHaveLength(1)
  const restored = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "provisional", reason: "The missing date stays marked under the original drafting permission.", basis: [] }) })
  await refreshFactAnswers(directory, state, { sessions: restored.port })
  expect(factDisposition(readQuestions(directory)[0]!)).toBe("provisional")
  expect(readQuestions(directory)[0]).not.toHaveProperty("fact_resolution")
  expect(waitingFactQuestions(directory)).toEqual([])
  expect(factOpenItems(directory)).toHaveLength(1)
})

test("an original provisional choice cannot be revived over a later hold when source refresh fails", async () => {
  const { directory, state } = waitingMatter()
  answerQuestion(directory, FACT.id, { option_id: "o1" }, factContext(directory))
  const hold = "Wait until the signed inspection report confirms the date before further drafting."
  updateAfterAnswer(directory, hold)
  const stopped = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "wait", reason: "The later instruction imposes a new evidence condition.", basis: [{ source_id: "updated-notes", quote: hold }] }) })
  await refreshFactAnswers(directory, state, { sessions: stopped.port })
  const replace = (text: string) => {
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    try { db.run("UPDATE chunks SET text = ?", [text]) } finally { db.close() }
  }
  const quote = "The signed inspection report confirms that the inspection was on 8 August."
  replace(quote)
  const supplied = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "supplied", reason: "The signed report satisfies the stated evidence condition.", basis: [{ source_id: SOURCE, quote }] }) })
  await refreshFactAnswers(directory, state, { sessions: supplied.port })
  expect(factDisposition(readQuestions(directory)[0]!)).toBe("supplied")
  replace("The signed inspection report has been withdrawn.")
  expect(waitingFactQuestions(directory)).toHaveLength(1)
  const failed = scriptedModel({ [AGENT.answer]: () => "invalid" })
  await refreshFactAnswers(directory, state, { sessions: failed.port })
  expect(factDisposition(readQuestions(directory)[0]!)).toBe("unknown")
  expect(waitingFactQuestions(directory)).toHaveLength(1)
})

test("focused refresh can preserve an original typed or picked provisional approval after unrelated updates", async () => {
  for (const typed of [false, true]) {
    const { directory, state } = waitingMatter()
    const permission = "Please continue with the date visibly marked unknown in a review draft."
    answerQuestion(directory, FACT.id, typed ? { text: permission } : { option_id: "o1" }, factContext(directory))
    if (typed) {
      const classified = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ class: "answered", fact_status: "provisional", fact_quote: permission }) })
      await classifyAnswers(directory, state, { sessions: classified.port })
    }
    const original = readQuestions(directory)[0]!.answer!
    const replace = (text: string) => {
      const db = new Database(path.join(directory, ".dochaus", "legal.db"))
      try { db.run("UPDATE chunks SET text = ?", [text]) } finally { db.close() }
    }
    const quote = "The inspection was on 8 August."
    replace(quote)
    const supplied = scriptedModel({ [AGENT.answer]: () => JSON.stringify({ fact_status: "supplied", reason: "The supplied report gives the date.", basis: [{ source_id: SOURCE, quote }] }) })
    await refreshFactAnswers(directory, state, { sessions: supplied.port })
    updateAfterAnswer(directory, "Use a footer.")
    replace("The inspection date is no longer established by the report.")
    expect(waitingFactQuestions(directory)).toHaveLength(1)
    const reconsider = scriptedModel({ [AGENT.answer]: prompt => {
      expect(prompt).toContain("an earlier permission never overrides a later hold")
      return JSON.stringify({ fact_status: "provisional", reason: "The original permission still governs because the only later instruction concerns the footer.", basis: [{ source_id: "original-answer", quote: original }] })
    } })
    await refreshFactAnswers(directory, state, { sessions: reconsider.port })
    expect(factDisposition(readQuestions(directory)[0]!)).toBe("provisional")
    expect(waitingFactQuestions(directory)).toEqual([])
    expect(factOpenItems(directory)).toHaveLength(1)
  }
})
