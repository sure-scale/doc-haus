import { afterEach, describe, expect, test } from "bun:test"
import { writeFileSync } from "node:fs"
import path from "node:path"
import { ADVISOR_FAILED, adviseQuestions, advisePrompt, checkAdvice } from "./advise"
import { questionIssues } from "./change-checks"
import { openItem } from "./quotes"
import { runLean, stageInputSha } from "./runner"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { keepAskedQuestions, readLeanState, readQuestions, writeLeanText, writeQuestions, type Question, type QuestionAdvice } from "./state"
import { FIXED_PLAN, writeLeanPlan } from "./plan"
import { confirmationQuestions, writeInstructionsFile } from "./instructions"
import { AGENT, SOURCE, SOURCE_TEXT, cleanupMatters, matter, scriptedModel } from "./test-support"
import { withWorkflowLimit } from "./model-concurrency-test-support"

afterEach(cleanupMatters)

const QUESTION: Question = { id: "q-delivery", stage: "intake", question: "Which delivery period should the note state?", context: "The period decides when the supplier is late.", status: "open", kind: "judgment" }
const FILES: Question = { id: "requested-files-1", stage: "draft", question: "Your words ask for one file, but the plan delivers two. Do you want it delivered as your words ask, or as planned?", context: "Drafting waits for your answer.", status: "open", kind: "file_name",
  options: [{ id: "o1", text: "As my words ask" }, { id: "o2", text: "As planned" }] }
const BASIS = [{ source_id: SOURCE, quote: "within thirty days of the order date", relevance: "The supplied delivery terms state the period this note must reproduce." }]
const OPTIONS = [{ text: "Thirty days from the order date" }, { text: "Fourteen days from the order date" }]
const VALID = { options: OPTIONS, recommendation: { option: "Thirty days from the order date", rationale: "The supply terms set thirty days.", basis: BASIS } }
const UNAVAILABLE = { unavailable: { reason: "The sources do not settle the period." } }
const check = (reply: unknown, question = QUESTION) => checkAdvice(matter(), question, typeof reply === "string" ? reply : JSON.stringify(reply))
const asks = (...questions: { question: string; kind?: string; context: string }[]) => () => JSON.stringify({ questions })
const LAW = { question: "Which governing law applies?", kind: "instruction", context: "The request names none." }
const SIGNATORY = { question: "Who signs for the supplier?", kind: "fact", context: "The supply terms name no signatory." }

describe("advice checker", () => {
  test("a supported recommendation is kept with code-assigned option ids and its verified passages", () => {
    expect(check(VALID)).toEqual({ issues: [], advice: {
      options: [{ id: "o1", text: "Thirty days from the order date" }, { id: "o2", text: "Fourteen days from the order date" }],
      recommendation: { option: "o1", rationale: "The supply terms set thirty days.", basis: BASIS },
    } })
  })

  test("no supported option is kept as unavailable, with no options or with two to four", () => {
    expect(check(UNAVAILABLE)).toEqual({ issues: [], advice: UNAVAILABLE })
    expect(check({ options: OPTIONS, ...UNAVAILABLE }).advice).toEqual({ options: [{ id: "o1", text: OPTIONS[0]!.text }, { id: "o2", text: OPTIONS[1]!.text }], ...UNAVAILABLE })
  })

  test("a question with its own options is recommended by their index, without options of the advisor's", () => {
    expect(check({ recommendation: { option: 1, rationale: "The request leaves the files to the plan.", basis: [{ source_id: "request", quote: "short supply note" }] } }, FILES))
      .toEqual({ issues: [], advice: { recommendation: { option: "o2", rationale: "The request leaves the files to the plan.", basis: [{ source_id: "request", quote: "short supply note" }] } } })
  })

  test("a generated-options recommendation resolves by its option text, not a slipped index", () => {
    const options = [{ text: "Seven days" }, { text: "Fourteen days" }, { text: "Thirty days" }, { text: "Sixty days" }]
    const advised = (option: unknown) => check({ options, recommendation: { ...VALID.recommendation, option } })
    // Copied with other case, spacing and a trailing full stop, the text still names its option.
    expect(advised("  thirty   DAYS. ").advice).toMatchObject({ recommendation: { option: "o3" } })
    // A 1-based index for "Thirty days" would have named "Sixty days": an index is no longer read.
    expect(advised(3).advice).toBeUndefined()
    expect(advised(3).issues.join(" ")).toContain('"recommendation.option" must be the text of one of your options')
  })

  test("an unmatched recommendation text is a retry issue", () => {
    const checked = check({ ...VALID, recommendation: { ...VALID.recommendation, option: "Thirty days from delivery" } })
    expect(checked.advice).toBeUndefined()
    expect(checked.issues).toEqual(['"recommendation.option" must be the text of one of your options, copied exactly: "Thirty days from the order date", "Fourteen days from the order date".'])
  })

  test("own-option questions keep the index", () => {
    const basis = [{ source_id: "request", quote: "short supply note" }]
    expect(check({ recommendation: { option: 1, rationale: "The request leaves the files to the plan.", basis } }, FILES).advice).toMatchObject({ recommendation: { option: "o2" } })
    const byText = check({ recommendation: { option: "As planned", rationale: "The request leaves the files to the plan.", basis } }, FILES)
    expect(byText.advice).toBeUndefined()
    expect(byText.issues).toEqual(['"recommendation.option" must be the index, from 0, of one of the 2 options.'])
    expect(advisePrompt(matter(), FILES)).toContain('where "option" is the index of one of the options above')
    expect(advisePrompt(matter(), QUESTION)).toContain('"option" the text of the one recommended, copied exactly from your options')
  })

  const SCOPE: Question = { id: "q-scope", stage: "draft", question: 'The document plan leaves this instructed act without a document executed by the parties who must take it: "Confirm the delivery period", directed by "within thirty days of the order date" [source: supply-terms.docx]; to be taken by "Supplier" and "Buyer"; no planned document carries it. Should the package add a document, or change a planned one, so that those parties take it, or is it outside this request?', context: "Drafting waits for your answer.", status: "open", kind: "scope" }
  const scopeReply = (carrier: string) => ({ options: [{ text: carrier }, { text: "A new letter executed by the supplier and the buyer" }, { text: "Outside this request" }],
    recommendation: { option: "Outside this request", rationale: "The supply terms already set the period.", basis: BASIS } })

  test("a scope option naming an uploaded source as carrier is a retry issue", () => {
    const named = check(scopeReply("Add it to supply-terms.docx, executed by the supplier and the buyer"), SCOPE)
    expect(named.advice).toBeUndefined()
    expect(named.issues).toEqual(['Option "Add it to supply-terms.docx, executed by the supplier and the buyer" puts the act in "supply-terms.docx", an uploaded matter source the package does not write: offer a planned document, a new document, or "outside this request".'])
    expect(check(scopeReply("Add it to supply-terms.docx, signed by both"), { ...SCOPE, kind: "judgment" }).issues).toEqual([])
    // A source's title alone is prose, such as an amendment to it.
    expect(check(scopeReply("A new amendment to the supply-terms, signed by both"), SCOPE).issues).toEqual([])
  })

  test("a scope option naming a planned document whose title resembles a source passes", () => {
    const labelled = { ...SCOPE, question: SCOPE.question.replace("no planned document carries it", 'planned in d1 "Variation of supply-terms", executed by "Supplier"') }
    expect(check(scopeReply("Variation of supply-terms, executed by the supplier and the buyer"), labelled).advice).toMatchObject({ recommendation: { option: "o3" } })
    const directory = matter()
    writeLeanPlan(directory, [{ title: "Supply-terms side letter", purpose: "Record the delivery period", parts: [{ heading: "Delivery", scope: "Delivery period" }] }])
    expect(checkAdvice(directory, SCOPE, JSON.stringify(scopeReply("Put it in the supply-terms side letter, signed by both"))).issues).toEqual([])
  })

  test("a scope option on a saved plan names a planned document or puts the item outside this request; the plan step's unsaved question still offers a new document", () => {
    const directory = matter()
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record the delivery terms", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
    // The gate's question, asked once the plan is saved, as the gate asks it.
    const gate: Question = { id: "package-scope-gate", stage: "gate", status: "open", kind: "scope", mixed: true, context: "Delivery is not held for it.",
      question: "The checklist check left this instruction open, so the package may not carry it:\nState when payment is due. Open decisions: whether payment is due on delivery.\nAnswer the open decisions above (supply any missing fact or approval they name). If the package should carry it, say which planned document carries it; otherwise say it is outside this request." }
    const reply = (...texts: string[]) => JSON.stringify({ options: texts.map(text => ({ text })), recommendation: { option: texts[0], rationale: "The supply terms set payment on delivery.", basis: BASIS } })
    const held = checkAdvice(directory, gate, reply("Hold the note in draft until the payment terms are confirmed", "Outside this request"))
    expect(held.advice).toBeUndefined()
    expect(held.issues).toEqual(['Option "Hold the note in draft until the payment terms are confirmed" neither names a planned document nor says the item is outside this request: the run can only carry it in a planned document or leave it out. Name the planned document that carries it, by its title or file ("Supply note", "supply-note.docx"), or say it is outside this request.'])
    expect(checkAdvice(directory, gate, reply("Carry it in the supply note: payment is due on delivery", "Treat it as outside the scope of this request")).advice).toMatchObject({ recommendation: { option: "o1" } })
    expect(checkAdvice(directory, gate, reply("Put it in supply-note.docx", "Outside this request")).issues).toEqual([])
    // The plan step's recheck of a saved plan offers no new document either; the documents it labels count as planned.
    const saved = { ...SCOPE, question: SCOPE.question.replace(/no planned document carries it\. Should .*$/, 'planned in d2 "Side letter", executed by "Supplier". Which planned document, executed by those parties, should carry it, or is it outside this request?') }
    expect(checkAdvice(directory, saved, reply("The side letter, executed by the supplier and the buyer", "Outside this request")).issues).toEqual([])
    expect(checkAdvice(directory, saved, reply("A new letter executed by the supplier and the buyer", "Outside this request")).issues).toHaveLength(1)
    // Before its plan is saved, the plan step's question offers a new document, which no list names.
    expect(checkAdvice(matter(), SCOPE, reply("A new letter executed by the supplier and the buyer", "Outside this request")).issues).toEqual([])
  })

  const rejected: [string, unknown, string, Question?][] = [
    ["a reply that is not JSON", "Thirty days, I think.", "JSON"],
    ["both a recommendation and unavailable", { ...VALID, ...UNAVAILABLE }, "exactly one"],
    ["neither a recommendation nor unavailable", { options: OPTIONS }, "exactly one"],
    ["one option with a recommendation", { ...VALID, options: [OPTIONS[0]] }, "2 to 4 options"],
    ["five options", { ...VALID, options: ["a", "b", "c", "d", "e"].map(text => ({ text })) }, "2 to 4 options"],
    ["one option with unavailable", { options: [OPTIONS[0]], ...UNAVAILABLE }, "no options, or 2 to 4"],
    ["a blank option", { ...VALID, options: [OPTIONS[0], { text: " " }] }, "non-empty text"],
    ["options that are the same once normalised", { ...VALID, options: [{ text: "Thirty days" }, { text: "  thirty   DAYS " }] }, "differ"],
    ["options given for a question with its own", { ...VALID, recommendation: { ...VALID.recommendation, basis: [{ source_id: "request", quote: "supply note" }] } }, "its own options", FILES],
    ["a recommendation of no option", { ...VALID, recommendation: { ...VALID.recommendation, option: 2 } }, "recommendation.option"],
    ["a blank rationale", { ...VALID, recommendation: { ...VALID.recommendation, rationale: " " } }, "rationale"],
    ["a blank reason", { unavailable: { reason: "" } }, "unavailable.reason"],
    ["no supporting passage", { ...VALID, recommendation: { ...VALID.recommendation, basis: [] } }, "at least one"],
    ["a passage from no allowed source", { ...VALID, recommendation: { ...VALID.recommendation, basis: [{ source_id: "memory", quote: "thirty days" }] } }, "source_id"],
    ["a blank quote", { ...VALID, recommendation: { ...VALID.recommendation, basis: [{ source_id: "request", quote: "  " }] } }, "non-empty"],
    ["a quote its source does not hold", { ...VALID, recommendation: { ...VALID.recommendation, basis: [{ source_id: SOURCE, quote: "within fourteen days" }] } }, "within fourteen days"],
  ]
  for (const [name, reply, issue, question] of rejected)
    test(`rejects ${name}`, () => {
      const checked = check(reply, question)
      expect(checked.advice).toBeUndefined()
      expect(checked.issues.join(" ")).toContain(issue)
    })
})

describe("advising asked questions", () => {
  test("each question an intake step asks is advised in its own fresh read-only session before anything from the step is recorded", async () => {
    const directory = matter()
    const seen: { status?: string; recorded: number }[] = []
    const model = scriptedModel({ [AGENT.intake]: asks(LAW, SIGNATORY), [AGENT.advise]: () => {
      seen.push({ status: readLeanState(directory)?.status, recorded: readQuestions(directory).length })
      return JSON.stringify(UNAVAILABLE)
    } })
    const state = await runLean(directory, { sessions: model.port })
    expect(state.status).toBe("blocked")
    expect(state.stages.intake?.status).toBe("blocked")
    expect(seen).toEqual([{ status: "running", recorded: 0 }, { status: "running", recorded: 0 }])
    const sessions = model.sessions.filter(row => row.agent === AGENT.advise)
    expect(sessions).toHaveLength(2)
    expect(sessions.every(row => row.permissions === LEAN_CHILD_PERMISSIONS && row.parentID === "ses_owner")).toBe(true)
    const prompts = model.prompts.filter(row => row.agent === AGENT.advise)
    expect(new Set(prompts.map(row => row.sessionID)).size).toBe(2)
    expect(prompts.find(row => row.text.includes(LAW.question))!.text).toContain("Kind: instruction\nWhich governing law applies?")
    expect(prompts.find(row => row.text.includes(LAW.question))!.text).toContain("The request names none.")
    expect(prompts.find(row => row.text.includes(LAW.question))!.text).toContain("Draft a short supply note from the supply terms.")
    expect(prompts.find(row => row.text.includes(LAW.question))!.text).toContain(SOURCE)
    expect(prompts.find(row => row.text.includes(LAW.question))!.text).toContain("Whom the lawyer acts for is not recorded yet")
    expect(prompts.find(row => row.text.includes(LAW.question))!.text).toContain("An instruction does not decide whether it can validly be carried out as given")
    expect(readQuestions(directory).map(row => [row.question, row.kind, row.advice])).toEqual([
      [LAW.question, "instruction", UNAVAILABLE], [SIGNATORY.question, "fact", UNAVAILABLE]])
  }, 120_000)

  test("an answered question is not advised", async () => {
    const directory = matter()
    writeQuestions(directory, [{ ...QUESTION, id: "q-done", status: "answered", answer: "Thirty days." }])
    const model = scriptedModel({ [AGENT.advise]: () => JSON.stringify(VALID) })
    const advised = await adviseQuestions(directory, readLeanState(directory)!, "intake", { sessions: model.port }, [{ ...QUESTION, id: "q-done" }, QUESTION])
    expect(model.count(AGENT.advise)).toBe(1)
    expect(advised[0]).toEqual({ ...QUESTION, id: "q-done" })
    expect(advised[1]!.advice).toMatchObject({ recommendation: { option: "o1" } })
  })

  test("an open question asked again with the same words and options keeps its advice; any change is advised again", async () => {
    const directory = matter()
    const kept = { options: [{ id: "o1", text: "Thirty days from the order date" }, { id: "o2", text: "Fourteen days from the order date" }],
      recommendation: { option: "o1", rationale: "The supply terms set thirty days.", basis: BASIS } }
    writeQuestions(directory, [{ ...QUESTION, advice: kept }, { ...FILES, advice: UNAVAILABLE }])
    // Advised afresh, the options would read otherwise under the same ids.
    const model = scriptedModel({ [AGENT.advise]: prompt => JSON.stringify(prompt.includes("As my words ask")
      ? { recommendation: { option: 1, rationale: "The request names no file.", basis: [{ source_id: "request", quote: "short supply note", relevance: "The request names the supplied note, not a file name." }] } }
      : { options: [{ text: "Seven days" }, { text: "Ten days" }], recommendation: { ...VALID.recommendation, option: "Seven days" } }) })
    const advise = (...questions: Question[]) => adviseQuestions(directory, readLeanState(directory)!, "intake", { sessions: model.port }, questions)
    expect(await advise(QUESTION, FILES)).toEqual([{ ...QUESTION, advice: kept }, { ...FILES, advice: UNAVAILABLE }])
    expect(model.count(AGENT.advise)).toBe(0)
    const [reworded] = await advise({ ...QUESTION, question: "Which delivery period applies?" })
    expect(model.count(AGENT.advise)).toBe(1)
    expect(reworded!.advice).toMatchObject({ options: [{ id: "o1", text: "Seven days" }, { id: "o2", text: "Ten days" }] })
    const [reoffered] = await advise({ ...FILES, options: [{ id: "o1", text: "As planned" }, { id: "o2", text: "As my words ask" }] })
    expect(model.count(AGENT.advise)).toBe(2)
    expect(reoffered!.advice).toMatchObject({ recommendation: { option: "o2" } })
  })

  test("a question with its own options shows them to the advisor and keeps them", async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.advise]: () => JSON.stringify({ recommendation: { option: 1, rationale: "The request names no file.", basis: [{ source_id: "request", quote: "short supply note", relevance: "The request names the supplied note, not a file name." }] } }) })
    const [advised] = await adviseQuestions(directory, readLeanState(directory)!, "d1.draft", { sessions: model.port }, [FILES])
    expect(model.prompts[0]!.text).toContain("0: As my words ask\n1: As planned")
    expect(advised).toEqual({ ...FILES, advice: { recommendation: { option: "o2", rationale: "The request names no file.", basis: [{ source_id: "request", quote: "short supply note", relevance: "The request names the supplied note, not a file name." }] } } })
  })

  test("a question asked before the outline is advised for the client recorded on the matter", async () => {
    const directory = matter()
    const reply = { recommendation: { option: 1, rationale: "The request names no file.", basis: [{ source_id: "request", quote: "short supply note", relevance: "The request names the supplied note." }] } }
    const model = scriptedModel({ [AGENT.advise]: () => JSON.stringify(reply) })
    await adviseQuestions(directory, readLeanState(directory)!, "intake", { sessions: model.port }, [FILES])
    expect(model.prompts[0]!.text).toContain("<acting_for>\nWhom the lawyer acts for is not recorded yet")
    // The ingest service writes matter.json; a synthetic one stands in for it.
    writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "matter-fixture", name: "Supply note", client: "Northwind Supplies Ltd" }))
    await adviseQuestions(directory, readLeanState(directory)!, "intake", { sessions: model.port }, [{ ...FILES, id: "q-client" }])
    expect(model.prompts.at(-1)!.text).toContain("<acting_for>\nThe lawyer acts for: Northwind Supplies Ltd (the client recorded on the matter), unless")
  })

  test("a document's confirmation is advised from the document's own opening, the request and the answers, without the matter's other context", async () => {
    const directory = matter()
    writeLeanText(directory, "instructions.json", JSON.stringify({ items: [{ id: "i1", text: "State delivery within thirty days.", source_id: SOURCE, quote: "within thirty days of the order date" }], documents: [{ source_id: SOURCE, from: "the client", quote: "within thirty days of the order date" }] }))
    writeInstructionsFile(directory)
    const [confirmation] = confirmationQuestions(directory)
    const reply = { recommendation: { option: 0, rationale: "The client wrote the terms for this work.", basis: [{ source_id: SOURCE, quote: "The supplier shall deliver", relevance: "Shows the client's terms for this work." }] } }
    const model = scriptedModel({ [AGENT.advise]: () => JSON.stringify(reply) })
    const [advised] = await adviseQuestions(directory, readLeanState(directory)!, "intake", { sessions: model.port }, [confirmation!, QUESTION])
    const [asked, other] = [model.prompts.find(row => row.text.includes(confirmation!.question))!.text, model.prompts.find(row => row.text.includes(QUESTION.question))!.text]
    expect(asked).toContain(`<document_opening>\n${SOURCE_TEXT}\n</document_opening>`)
    expect(asked).toContain("<request>\nDraft a short supply note from the supply terms.")
    expect(asked).toContain("that alone is no reason to treat the document as evidence")
    for (const tag of ["<instructions>", "<acting_for>", "<legal_research>", "<existing_topics>", "Matter source names"]) {
      expect(asked).not.toContain(tag)
      expect(other).toContain(tag)
    }
    expect(advised!.advice).toMatchObject({ recommendation: { option: "o1" } })
  })

  test("a changed question scope or missing supporting quote clears cached recommendation while retaining option ids", async () => {
    const directory = matter()
    const advice = { options: [{ id: "o1", text: "Keep the supplied note" }, { id: "o2", text: "Seek new instructions" }],
      recommendation: { option: "o1", rationale: "The request asks for the note.", basis: [{ source_id: "request", quote: "short supply note", relevance: "The original request names this document." }] } }
    writeQuestions(directory, [{ ...QUESTION, advice }])
    const model = scriptedModel({})
    const changed = await adviseQuestions(directory, readLeanState(directory)!, "intake", { sessions: model.port }, [{ ...QUESTION, context: "A different document is now requested.", document: "d2" }])
    expect(changed[0]!.advice).toEqual({ options: advice.options, unavailable: { reason: expect.stringContaining("changed") } })
    writeLeanText(directory, "request.md", "Prepare a chronology.")
    const sourceChanged = await adviseQuestions(directory, readLeanState(directory)!, "intake", { sessions: model.port }, [QUESTION])
    expect(sourceChanged[0]!.advice).toHaveProperty("unavailable")
    expect(sourceChanged[0]!.advice?.options).toEqual(advice.options)
    expect(model.count(AGENT.advise)).toBe(0)
  })

  test("when the advisor's retries run out the question is recorded as advisor_failed and intake stays blocked", async () => {
    // An unchecked reply is retried in its session; a runtime error ends the advice at once.
    const cases: [() => string, number][] = [[() => "I would pick thirty days.", 1 + LEAN_MAX_RETRIES], [() => { throw new Error("provider unavailable") }, 1]]
    for (const [advise, calls] of cases) {
      const directory = matter()
      const model = scriptedModel({ [AGENT.intake]: asks(LAW), [AGENT.advise]: advise })
      const state = await runLean(directory, { sessions: model.port })
      expect(state.status).toBe("blocked")
      expect(state.stages.intake?.status).toBe("blocked")
      expect(model.count(AGENT.advise)).toBe(calls)
      expect(readQuestions(directory)).toEqual([expect.objectContaining({ question: LAW.question, status: "open", kind: "instruction",
        advice: { unavailable: { reason: expect.any(String), reason_code: "advisor_failed" } } })])
    }
  }, 120_000)

  test("a preflight review whose advisor failed is advised again; a checked one is reused", async () => {
    const directory = matter()
    const key = (question: Question) => JSON.stringify([question.question, question.context, question.kind, question.document, question.pid])
    const other: Question = { ...QUESTION, id: "q-notice", question: "Which notice period should the note state?" }
    const reviewed = new Map<string, { advice?: QuestionAdvice; group?: string }>([
      [key(QUESTION), { advice: { unavailable: { reason: ADVISOR_FAILED, reason_code: "advisor_failed" } } }],
      [key(other), { advice: UNAVAILABLE, group: "Notice period" }],
    ])
    const model = scriptedModel({ [AGENT.advise]: () => JSON.stringify(VALID) })
    const [retried, reused] = await adviseQuestions(directory, readLeanState(directory)!, "d1.draft", { sessions: model.port }, [QUESTION, other], reviewed)
    expect(model.count(AGENT.advise)).toBe(1)
    expect(retried!.advice).toMatchObject({ recommendation: { option: "o1", basis: BASIS } })
    expect(reused).toMatchObject({ advice: UNAVAILABLE, group: "Notice period" })
  })

  test("six questions respect the configured concurrency limit", async () => withWorkflowLimit(3, async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.intake]: asks(...["one", "two", "three", "four", "five", "six"].map(n => ({ ...LAW, question: `Question ${n}?` }))) })
    const live = { now: 0, max: 0 }
    const port: StageSessionPort = { createSession: model.port.createSession, async prompt(input) {
      if (input.agent !== AGENT.advise) return model.port.prompt(input)
      live.now += 1
      live.max = Math.max(live.max, live.now)
      await new Promise(resolve => setTimeout(resolve, 20))
      live.now -= 1
      return model.port.prompt(input)
    } }
    expect((await runLean(directory, { sessions: port })).status).toBe("blocked")
    expect(model.count(AGENT.advise)).toBe(6)
    expect(live.max).toBe(3)
    expect(readQuestions(directory).every(row => row.advice)).toBe(true)
  }), 120_000)

  test("a question without its kind is retried", async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.intake]: (_prompt, index) => asks(index ? LAW : { question: LAW.question, context: LAW.context })() })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
    const intake = model.prompts.filter(row => row.agent === AGENT.intake)
    expect(intake).toHaveLength(2)
    expect(intake[1]!.text).toContain('questions[0].kind must be "fact", "instruction" or "judgment"')
    expect(readQuestions(directory).map(row => row.kind)).toEqual(["instruction"])
    // The change stages' questions are checked the same way; code-set kinds are not a model's to give.
    expect(questionIssues({ question: "Which law?" }, 0, new Map(), new Map()).join(" ")).toContain("questions[0].kind must be")
    expect(questionIssues({ question: "Which law?", kind: "scope" }, 0, new Map(), new Map()).join(" ")).toContain("questions[0].kind must be")
    expect(questionIssues({ question: "Which law?", kind: "judgment" }, 0, new Map(), new Map())).toEqual([])
  }, 120_000)

  test("a stage key is unchanged by a question's kind, options, advice or answer time", () => {
    const directory = matter()
    const asked: Question[] = [{ ...QUESTION, id: "gate-q", stage: "gate" }, { ...QUESTION, id: "q-answered", status: "answered", answer: "Thirty days." }]
    const steps = ["finalize", "gate", "d1.change-review"] as const
    writeQuestions(directory, asked.map(({ kind: _kind, ...question }) => question))
    const before = steps.map(step => stageInputSha(directory, step))
    writeQuestions(directory, asked.map(question => ({ ...question, options: FILES.options, advice: { unavailable: { reason: "None.", reason_code: "advisor_failed" as const } },
      ...(question.status === "answered" ? { answered_at: "2026-10-19T00:00:00.000Z" } : {}) })))
    expect(steps.map(step => stageInputSha(directory, step))).toEqual(before)
  })

  test("an open question stored before advice is asked again, advised; an answer is kept", async () => {
    const directory = matter()
    const answered = { id: "q-kept", stage: "intake", question: "Who signs for the supplier?", context: "", status: "answered", answer: "The director." }
    writeLeanText(directory, "questions.json", JSON.stringify({ questions: [{ id: "q-old", stage: "intake", question: LAW.question, context: LAW.context, status: "open" }, answered] }))
    const model = scriptedModel({ [AGENT.intake]: asks(LAW) })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
    expect(model.count(AGENT.intake)).toBe(1)
    // The answer is kept, read once as the lawyer's own.
    expect(readQuestions(directory)).toEqual([{ ...answered, answer_class: "answered" }, expect.objectContaining({ question: LAW.question, status: "open", kind: "instruction", advice: expect.anything() })])
  }, 120_000)

  test("a question passed to the open item still asking it keeps its kind, options and advice", () => {
    const advice = { unavailable: { reason: "None." } }
    const kept = keepAskedQuestions([{ ...FILES, advice }], [openItem("changes", "question", FILES.question, { question_id: FILES.id, pid: "p000000000001", document: "d2" })], () => false)
    expect(kept).toEqual([{ ...FILES, stage: "changes", pid: "p000000000001", document: "d2", advice }])
  })
})

test("advice on how far the drafting qualifies names the points it applies to", () => {
  expect(advisePrompt(matter(), QUESTION)).toContain("An option on how far the drafting qualifies or hedges names the points it applies to: a document's register is not a choice between blanket modes.")
})

test("advice reads the matter's dates, so a deadline option does not run from the drafting day", () => {
  const directory = matter()
  expect(advisePrompt(directory, QUESTION)).not.toContain("<matter_dates>")
  writeLeanText(directory, "intake.json", JSON.stringify({ ready: true, dates: { document_date: { date: "2026-03-01", basis: "request", source_id: "request", quote: "dated 1 March 2026" }, operative_dates: [] } }))
  expect(advisePrompt(directory, QUESTION)).toContain("<matter_dates>\nDate the document 2026-03-01")
})

// An option to discard a planned document and prepare another was recommended,
// adopted and asked again until the repair cap: nothing can carry it out once
// the plan is saved.
test("once the plan is saved, advice on any but a package-scope question keeps the planned documents", () => {
  const directory = matter()
  expect(advisePrompt(directory, QUESTION)).not.toContain(FIXED_PLAN)
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record the delivery terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
  expect(advisePrompt(directory, { ...QUESTION, stage: "change-plan", document: "d1", kind: "instruction" })).toContain(FIXED_PLAN)
  expect(advisePrompt(directory, { ...QUESTION, kind: "scope" })).not.toContain(FIXED_PLAN)
})
