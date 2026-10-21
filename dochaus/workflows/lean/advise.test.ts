import { afterEach, describe, expect, test } from "bun:test"
import { adviseQuestions, checkAdvice } from "./advise"
import { questionIssues } from "./change-checks"
import { openItem } from "./quotes"
import { runLean, stageInputSha } from "./runner"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { keepAskedQuestions, readLeanState, readQuestions, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, SOURCE, cleanupMatters, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const QUESTION: Question = { id: "q-delivery", stage: "intake", question: "Which delivery period should the note state?", context: "The period decides when the supplier is late.", status: "open", kind: "judgment" }
const FILES: Question = { id: "requested-files-1", stage: "draft", question: "Your words ask for one file, but the plan delivers two. Do you want it delivered as your words ask, or as planned?", context: "Drafting waits for your answer.", status: "open", kind: "file_name",
  options: [{ id: "o1", text: "As my words ask" }, { id: "o2", text: "As planned" }] }
const BASIS = [{ source_id: SOURCE, quote: "within thirty days of the order date" }]
const OPTIONS = [{ text: "Thirty days from the order date" }, { text: "Fourteen days from the order date" }]
const VALID = { options: OPTIONS, recommendation: { option: 0, rationale: "The supply terms set thirty days.", basis: BASIS } }
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
    ["a recommendation not by index", { ...VALID, recommendation: { ...VALID.recommendation, option: "Thirty days from the order date" } }, "recommendation.option"],
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
    expect(prompts[0]!.text).toContain("Kind: instruction\nWhich governing law applies?")
    expect(prompts[0]!.text).toContain("The request names none.")
    expect(prompts[0]!.text).toContain("Draft a short supply note from the supply terms.")
    expect(prompts[0]!.text).toContain(SOURCE)
    expect(prompts[0]!.text).toContain("The request does not say whom the lawyer acts for")
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
      ? { recommendation: { option: 1, rationale: "The request names no file.", basis: [{ source_id: "request", quote: "short supply note" }] } }
      : { ...VALID, options: [{ text: "Seven days" }, { text: "Ten days" }] }) })
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
    const model = scriptedModel({ [AGENT.advise]: () => JSON.stringify({ recommendation: { option: 1, rationale: "The request names no file.", basis: [{ source_id: "request", quote: "short supply note" }] } }) })
    const [advised] = await adviseQuestions(directory, readLeanState(directory)!, "d1.draft", { sessions: model.port }, [FILES])
    expect(model.prompts[0]!.text).toContain("0: As my words ask\n1: As planned")
    expect(advised).toEqual({ ...FILES, advice: { recommendation: { option: "o2", rationale: "The request names no file.", basis: [{ source_id: "request", quote: "short supply note" }] } } })
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

  test("six questions are advised at most four at a time", async () => {
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
    expect(live.max).toBe(4)
    expect(readQuestions(directory).every(row => row.advice)).toBe(true)
  }, 120_000)

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
