import { afterEach, describe, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { checkChanges } from "./change-checks"
import { finalBodyItems, readChangeList, type ParagraphRow } from "./checks"
import type { GateResult } from "./gate"
import type { StepID } from "./plan"
import { runLean } from "./runner"
import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { hashStageInputs, readLeanJson, readLeanState, readOpenItems, writeLeanState, writeQuestions } from "./state"
import { AGENT, SUPPORT, cleanupMatters, draftingNotes, happyPath, matter, pidOf, sameSession, scriptedModel, type Reply } from "./test-support"

// Amendments that rest on drafting judgment rather than a source passage. Real
// checker, coordinator, applier and gate; only the model is scripted.

afterEach(cleanupMatters)

const DELIVERY = "The supplier shall deliver the goods."
const JUDGED = "The supplier shall deliver the goods in good condition [open item: confirm the condition standard chosen]."
const UNMARKED = "The supplier shall deliver the goods in good condition."
const P1 = "p000000000001"
const NO_SOURCE = (index: number, pid = P1) =>
  `amendments[${index}] ([${pid}]) has no source passage. Quote the passage of a matter document that supports the change itself; ` +
  "where the change rests on drafting judgment or a view of the law, add a new visible [open item: confirm ...] marker saying what the lawyer is to confirm. " +
  "A marker, ___ field or placeholder already in the paragraph does not count."
const MARKER_ADDS_NOTHING = (index: number, pid = P1) =>
  `amendments[${index}] ([${pid}]) only adds an open-item marker to a paragraph that already shows one. Drop it, or change the drafting.`
const MARKER_ONLY = (index: number, pid = P1) =>
  `amendments[${index}] ([${pid}]) changes only the paragraph's open-item markers. Change the drafting and keep the marker, or put the point in "questions" with the pid "${pid}".`
const SOURCES_SHAPE = (index: number) => `amendments[${index}].sources must be a list of {source_id, quote}.`

const row = (pid: string, text: string): ParagraphRow => ({ pid, sha256: paragraphSha256(text), text })
const PARAGRAPHS = [row("p000000000001", DELIVERY), row("p000000000002", "Payment is due on delivery.")]
const reply = (amendments: object[], questions: object[] = []) => JSON.stringify({ amendments, questions })
const paragraph = (text: string) => [row(P1, text), PARAGRAPHS[1]!]
const gateOf = (directory: string) => readLeanJson(directory, "gate.json") as GateResult

describe("the changes checker admits unsourced judgment amendments only when marked", () => {
  test("empty sources with a visible confirm marker in the updated paragraph are accepted as they are", () => {
    const check = checkChanges(matter(), reply([{ pid: "p000000000001", updated: JUDGED, reason: "Drafting judgment.", sources: [] }]), PARAGRAPHS)
    expect(check.issues).toEqual([])
    expect(check.changes?.amendments).toEqual([expect.objectContaining({ id: "a1", pid: "p000000000001", updated: JUDGED, sources: [] })])
  })

  test("empty sources without a marker are an actionable retry issue", () => {
    const check = checkChanges(matter(), reply([{ pid: P1, updated: UNMARKED, reason: "Drafting judgment.", sources: [] }]), PARAGRAPHS)
    expect(check.issues).toEqual([NO_SOURCE(0)])
    expect(check.changes).toBeUndefined()
  })

  test("missing or non-list sources are a shape issue even when the paragraph is marked, and never reach the filler", () => {
    for (const sources of [undefined, null, "none", { source_id: "x" }]) {
      const check = checkChanges(matter(), reply([{ pid: P1, updated: JUDGED, reason: "Drafting judgment.", ...(sources === undefined ? {} : { sources }) }]), PARAGRAPHS)
      expect(check.issues).toEqual([SOURCES_SHAPE(0)])
      expect(check.changes).toBeUndefined()
    }
  })

  test("a sourced amendment is checked and kept as before", () => {
    const updated = "The supplier shall deliver the goods within thirty days of the order date."
    const check = checkChanges(matter(), reply([{ pid: "p000000000001", updated, reason: "Period.", sources: SUPPORT }]), PARAGRAPHS)
    expect(check.issues).toEqual([])
    expect(check.changes?.amendments).toEqual([expect.objectContaining({ updated, sources: SUPPORT })])
  })
})

describe("the marker must be new to the amendment", () => {
  const CONFIRMED = "The supplier shall deliver the goods [open item: confirm the delivery period chosen]."
  const unsourced = (current: string, updated: string) =>
    checkChanges(matter(), reply([{ pid: P1, updated, reason: "Drafting judgment.", sources: [] }]), paragraph(current))

  test("an open item already in the paragraph does not cover a new unsourced change", () => {
    expect(unsourced(CONFIRMED, "The supplier shall deliver the goods [open item: confirm the delivery period chosen] and shall insure them in transit.").issues)
      .toEqual([NO_SOURCE(0)])
  })

  test("a ___ field already in the paragraph does not cover a new unsourced change", () => {
    expect(unsourced("The supplier shall deliver the goods to ______.", "The supplier shall deliver the goods to ______ and gives no warranty of fitness.").issues)
      .toEqual([NO_SOURCE(0)])
  })

  test("replacing the paragraph with a ___ field is a hidden deletion and an unsourced change without a marker", () => {
    const issues = unsourced(DELIVERY, "______").issues
    expect(issues).toHaveLength(2)
    expect(issues[0]).toContain("replaces the paragraph's wording with markers only")
    expect(issues[1]).toBe(NO_SOURCE(0))
  })

  test("a drafting placeholder, already there or newly added, is not an open item for the change", () => {
    expect(unsourced("The supplier shall deliver the goods to [insert delivery address].",
      "The supplier shall deliver the goods to [insert delivery address] and shall insure them in transit.").issues).toEqual([NO_SOURCE(0)])
    expect(unsourced(DELIVERY, "The supplier shall deliver the goods to [insert delivery address].").issues).toEqual([NO_SOURCE(0)])
  })

  test("a new marker on a paragraph that already had one is accepted", () => {
    const updated = "The supplier shall deliver the goods [open item: confirm the delivery period chosen] and shall insure them in transit [open item: confirm the transit cover chosen]."
    const check = unsourced(CONFIRMED, updated)
    expect(check.issues).toEqual([])
    expect(check.changes?.amendments).toEqual([expect.objectContaining({ pid: P1, updated, sources: [] })])
  })
})

describe("an amendment that changes only markers", () => {
  const CONFIRMED = "The supplier shall deliver the goods [open item: confirm the delivery period chosen]."
  const QUESTION = { question: "Is delivery to be made in instalments?", kind: "judgment", pid: P1 }

  test("adding a marker to a paragraph that already shows one is an issue, even with a question for it", () => {
    const updated = "The supplier shall deliver the goods [open item: confirm the delivery period chosen] [open item: whether delivery is in instalments]."
    const check = checkChanges(matter(), reply([{ pid: P1, updated, reason: "Instalments are open.", sources: [] }], [QUESTION]), paragraph(CONFIRMED))
    expect(check.issues).toEqual([MARKER_ADDS_NOTHING(0)])
  })

  test("marking an unmarked paragraph with no question for it is an issue", () => {
    const check = checkChanges(matter(), reply([{ pid: P1, updated: CONFIRMED, reason: "The period is open.", sources: [] }]), PARAGRAPHS)
    expect(check.issues).toEqual([MARKER_ONLY(0)])
  })

  test("marking an unmarked paragraph for a question about it is accepted", () => {
    const check = checkChanges(matter(), reply([{ pid: P1, updated: CONFIRMED, reason: "The period is open.", sources: [] }], [QUESTION]), PARAGRAPHS)
    expect(check.issues).toEqual([])
    expect(check.changes?.amendments).toEqual([expect.objectContaining({ pid: P1, updated: CONFIRMED })])
    expect(check.changes?.questions).toEqual([expect.objectContaining({ question: QUESTION.question, pid: P1 })])
  })

  test("removing a marker with a supporting source is a visible change and is accepted", () => {
    const check = checkChanges(matter(), reply([{ pid: P1, updated: DELIVERY, reason: "The period is settled.", sources: SUPPORT }]), paragraph(CONFIRMED))
    expect(check.issues).toEqual([])
  })
})

describe("a split that keeps the paragraph as it stands", () => {
  const CONFIRMED = "The supplier shall deliver the goods [open item: confirm the delivery period chosen]."
  const ADDED = "The supplier shall insure the goods in transit [open item: confirm the transit cover chosen]."
  const split = (current: string, updated: string[], sources: object[] = []) =>
    checkChanges(matter(), reply([{ pid: P1, operation: "split", updated, reason: "Adds a paragraph.", sources }]), paragraph(current))

  test("the kept paragraph before the new one changes nothing; the new one is checked as before", () => {
    const check = split(DELIVERY, [DELIVERY, ADDED])
    expect(check.issues).toEqual([])
    expect(check.changes?.amendments).toEqual([expect.objectContaining({ pid: P1, operation: "split", updated: [DELIVERY, ADDED] })])
    expect(split(DELIVERY, [DELIVERY, "The supplier shall insure the goods in transit."]).issues).toEqual([NO_SOURCE(0)])
  })

  test("the kept paragraph after the new one, with any open item it shows, changes nothing", () => {
    expect(split(DELIVERY, [ADDED, DELIVERY]).issues).toEqual([])
    expect(split(CONFIRMED, [ADDED, CONFIRMED]).issues).toEqual([])
  })

  test("a title kept as it stands with new introductory text changes nothing in the title", () => {
    expect(split("Delivery", ["Delivery", "Delivery is due within thirty days of the order date."], SUPPORT).issues).toEqual([])
  })

  test("a replacement that only adds a marker still changes only markers", () => {
    const check = checkChanges(matter(), reply([{ pid: P1, updated: CONFIRMED, reason: "The period is open.", sources: [] }]), PARAGRAPHS)
    expect(check.issues).toEqual([MARKER_ONLY(0)])
  })

  test("a split part that only adds a marker to the paragraph still changes only markers", () => {
    expect(split(DELIVERY, [CONFIRMED, ADDED]).issues).toEqual([MARKER_ONLY(0)])
  })
})

describe("markers are read as the final body will show them", () => {
  const SPACING = (marker: string) => `amendments[0] ([${P1}]) writes the marker "${marker}" with spacing that will not show as an open item. ` +
    "Write it as [open item: ...] with ordinary single spaces."
  const EMPTY = (marker: string) => `amendments[0] ([${P1}]) adds the marker "${marker}", which does not say what the lawyer is to confirm or supply. Name that point in the marker.`
  const unsourced = (updated: string) => checkChanges(matter(), reply([{ pid: P1, updated, reason: "Drafting judgment.", sources: [] }]), PARAGRAPHS)

  test("a marker written with a non-breaking space or a double space is not a new marker, and the retry says why", () => {
    for (const gap of ["\u00A0", "  "]) {
      const check = unsourced(`${UNMARKED} [open${gap}item: confirm the condition standard chosen]`)
      expect(check.issues).toEqual([SPACING("[open item: confirm the condition standard chosen]"), NO_SOURCE(0)])
    }
    expect(unsourced(`${UNMARKED} [open item:  confirm the condition standard chosen]`).issues).toEqual([])
  })

  test("a marker that names nothing to confirm is an issue", () => {
    expect(unsourced(`${UNMARKED} [open item: confirm]`).issues).toEqual([EMPTY("[open item: confirm]")])
    expect(unsourced(`${UNMARKED} [open item: confirm ...]`).issues).toEqual([EMPTY("[open item: confirm ...]")])
    expect(unsourced(`${UNMARKED} [open item: confirm the standard]`).issues).toEqual([])
  })

  test("the same marker in two paragraphs is an open item in each; repeated ___ fields stay one item", () => {
    const marker = "[open item: confirm the standard]"
    const rows = [row(P1, `Goods are delivered in good condition ${marker}.`), row("p000000000002", `Goods are packed to the same standard ${marker}.`),
      row("p000000000003", "Signed for the supplier: ____"), row("p000000000004", "Signed for the buyer: ____")]
    const items = finalBodyItems(matter(), rows)
    const markers = items.filter(item => item.ref.placeholder === marker)
    expect(markers.map(item => item.ref.pid)).toEqual([P1, "p000000000002"])
    expect(new Set(markers.map(item => item.id)).size).toBe(2)
    expect(items.filter(item => item.ref.placeholder === "____").map(item => item.ref.pid)).toEqual(["p000000000003"])
  })
})

describe("on the last allowed attempt", () => {
  test("an unmarked unsourced amendment is withheld and recorded; the others are kept", () => {
    const sourced = { pid: P1, updated: "The supplier shall deliver the goods within thirty days of the order date.", reason: "Period.", sources: SUPPORT }
    const unmarked = { pid: "p000000000002", updated: "Payment is due within ten days of delivery.", reason: "Drafting judgment.", sources: [] }
    const check = checkChanges(matter(), reply([sourced, unmarked]), PARAGRAPHS, { withholdUnverified: true })
    expect(check.issues).toEqual([])
    expect(check.changes?.amendments).toEqual([expect.objectContaining({ id: "a1", pid: P1 })])
    expect(check.changes?.withheld).toEqual([{ amendment: expect.objectContaining({ id: "a2", pid: "p000000000002" }), issues: [NO_SOURCE(1, "p000000000002")] }])
    expect(check.openItems).toEqual([expect.objectContaining({ kind: "unapplied_amendment", ref: expect.objectContaining({ amendment_id: "a2", pid: "p000000000002" }) })])
  })

  test("through the run: the withheld amendment is disclosed and not applied, the sourced one is applied, and the run completes", async () => {
    const directory = matter()
    const changes = sameSession(prompt => reply([
      { pid: pidOf(prompt, DELIVERY), updated: "The supplier shall deliver the goods within thirty days of the order date.", reason: "Period.", sources: SUPPORT },
      { pid: pidOf(prompt, "Payment is due on delivery."), updated: "Payment is due within ten days of delivery.", reason: "Drafting judgment.", sources: [] },
    ]))
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: changes })
    const state = await runLean(directory, { sessions: model.port })

    expect(model.count(AGENT.changes)).toBe(4)
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [{ id: "a1" }], unapplied: [] })
    expect(readChangeList(directory, "d1")?.withheld).toEqual([expect.objectContaining({ amendment: expect.objectContaining({ id: "a2" }) })])
    const item = readOpenItems(directory).find(row => row.kind === "unapplied_amendment")!
    const finalize = model.prompts.find(prompt => prompt.agent === AGENT.finalize)!.text
    expect(finalize).toContain("- a2 [")
    expect(finalize).toContain("(withheld, not applied)")
    expect(finalize).toContain(item.id)
    expect(state.status).toBe("completed")
    expect(gateOf(directory)).toMatchObject({ status: "pass", delivery: "review_draft", disclosed: expect.arrayContaining([expect.objectContaining({ id: item.id })]) })
  }, 120_000)
})

describe("a paragraph is never deleted silently", () => {
  test("wording replaced by a marker alone is a retry issue, and on the last attempt is withheld so the paragraph keeps its text", () => {
    const content = reply([{ pid: P1, updated: "[open item: confirm whether delivery terms are needed]", reason: "Delete.", sources: [] }])
    const issue = "amendments[0] ([p000000000001]) replaces the paragraph's wording with markers only, which would delete it without showing a deletion. " +
      "Keep the wording and add an [open item: ...] marker saying why the lawyer should consider removing it."
    expect(checkChanges(matter(), content, PARAGRAPHS).issues).toEqual([issue])
    const last = checkChanges(matter(), content, PARAGRAPHS, { withholdUnverified: true })
    expect(last.changes?.amendments).toEqual([])
    expect(last.openItems).toEqual([expect.objectContaining({ kind: "unapplied_amendment", ref: expect.objectContaining({ amendment_id: "a1", pid: P1 }) })])
  })

  test("a marker beside kept wording is not a deletion", () => {
    expect(checkChanges(matter(), reply([{ pid: P1, updated: JUDGED, reason: "Drafting judgment.", sources: [] }]), PARAGRAPHS).issues).toEqual([])
  })
})

describe("the request and the lawyer's answers are a quotable source", () => {
  const BY_REQUEST = (quote: string) => [{ source_id: "request", quote }]
  const answered = () => {
    const directory = matter()
    writeQuestions(directory, [{ id: "q1", stage: "changes", question: "Which delivery period applies?", context: "Silent.", status: "answered", answer: "Delivery within fourteen days." }])
    return directory
  }
  const check = (directory: string, quote: string) =>
    checkChanges(directory, reply([{ pid: P1, updated: UNMARKED, reason: "Per the request.", sources: BY_REQUEST(quote) }]), PARAGRAPHS)

  test("words of the request, with spacing folded, support a change without a marker", () => {
    const result = check(matter(), "Draft a  short supply note")
    expect(result.issues).toEqual([])
    expect(result.changes?.amendments).toEqual([expect.objectContaining({ sources: BY_REQUEST("Draft a  short supply note") })])
  })

  test("words of an answer are quotable; the question, a misquote, or words running across request and answer are not", () => {
    const directory = answered()
    expect(check(directory, "within fourteen days").issues).toEqual([])
    for (const quote of ["Which delivery period applies", "Draft a long supply note", "supply terms. Delivery within"]) {
      const issues = check(directory, quote).issues
      expect(issues).toHaveLength(1)
      expect(issues[0]).toContain("are not in the request, user updates or any one of the lawyer's answers")
    }
  })

  test("through the run, a change resting only on the request is tagged for the drafting notes", async () => {
    const directory = matter()
    const changes = sameSession(prompt => reply([{ pid: pidOf(prompt, DELIVERY), updated: UNMARKED, reason: "Per the request.", sources: BY_REQUEST("a short supply note") }]))
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: changes })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    expect(model.prompts.find(prompt => prompt.agent === AGENT.changes)!.text).toContain('use the source name "request"')
    const evidence = JSON.parse(/<applied_changes>\n([\s\S]*?)\n<\/applied_changes>/.exec(model.prompts.find(prompt => prompt.agent === AGENT.finalize)!.text)![1]!)
    expect(evidence).toMatchObject([{ id: "a1", application_mode: "code", evidence_basis: "source-passages-provided", sources: BY_REQUEST("a short supply note") }])
  }, 120_000)
})

describe("a question whose marker only a withheld amendment carried", () => {
  const MISQUOTED = [{ source_id: SUPPORT[0]!.source_id, quote: "words that appear in no matter document" }]
  const QUESTION = "Which condition standard applies to delivered goods?"

  test("is still asked, refers to the withheld amendment, and is no retry issue on the last attempt", () => {
    const content = reply([{ pid: P1, updated: JUDGED, reason: "Drafting judgment.", sources: MISQUOTED }], [{ pid: P1, question: QUESTION, kind: "judgment" }])
    expect(checkChanges(matter(), content, PARAGRAPHS).issues.length).toBeGreaterThan(0)
    const check = checkChanges(matter(), content, PARAGRAPHS, { withholdUnverified: true })
    expect(check.issues).toEqual([])
    expect(check.changes?.amendments).toEqual([])
    expect(check.openItems).toEqual([
      expect.objectContaining({ kind: "unapplied_amendment", ref: expect.objectContaining({ amendment_id: "a1", pid: P1 }) }),
      expect.objectContaining({ kind: "question", ref: expect.objectContaining({ pid: P1, amendment_id: "a1" }) }),
    ])
  })

  test("through the run: an unresolved question pauses before applying or delivering", async () => {
    const directory = matter()
    const changes = sameSession(prompt => reply(
      [{ pid: pidOf(prompt, DELIVERY), updated: JUDGED, reason: "Drafting judgment.", sources: MISQUOTED }],
      [{ pid: pidOf(prompt, DELIVERY), question: QUESTION, kind: "judgment" }]))
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: changes })
    const state = await runLean(directory, { sessions: model.port })

    const question = readOpenItems(directory).find(item => item.kind === "question")!
    expect(question.ref.amendment_id).toBe("a1")
    expect(state.status).toBe("blocked")
    expect(state.stages["d1.apply"]?.status).toBe("pending")
    expect(gateOf(directory)).toBeUndefined()
  }, 120_000)
})

describe("a marked judgment amendment through the run", () => {
  test("an unmarked one is retried with the issue; the marked retry is applied, its marker is disclosed, and the drafting notes must name it", async () => {
    const directory = matter()
    const changes = sameSession((prompt, index) =>
      reply([{ pid: pidOf(prompt, DELIVERY), updated: index === 0 ? UNMARKED : JUDGED, reason: "Drafting judgment.", sources: [] }]))
    // The retry does not repeat the open items, so the notes keep reading the first prompt.
    const first = { prompt: "" }
    const notes: Reply = (prompt, index) => {
      first.prompt ||= prompt
      return index === 0 ? "# Drafting notes\n\nThe draft follows the supply terms." : draftingNotes(first.prompt, index)
    }
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: changes, [AGENT.finalize]: notes })
    const state = await runLean(directory, { sessions: model.port })

    const changesPrompts = model.prompts.filter(prompt => prompt.agent === AGENT.changes)
    expect(changesPrompts).toHaveLength(2)
    expect(changesPrompts[1]!.text).toContain(NO_SOURCE(0, pidOf(changesPrompts[0]!.text, DELIVERY)))
    expect(readChangeList(directory, "d1")?.amendments).toEqual([expect.objectContaining({ updated: JUDGED, sources: [] })])
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [{ id: "a1", mode: "code" }], unapplied: [] })

    const marker = readOpenItems(directory).find(item => item.kind === "placeholder" && item.stage === "apply")!
    expect(marker.ref).toMatchObject({ placeholder: "[open item: confirm the condition standard chosen]", document: "d1" })
    const finalize = model.prompts.filter(prompt => prompt.agent === AGENT.finalize)
    expect(finalize).toHaveLength(2)
    expect(finalize[0]!.text).toContain(marker.id)
    const evidence = JSON.parse(/<applied_changes>\n([\s\S]*?)\n<\/applied_changes>/.exec(finalize[0]!.text)![1]!)
    expect(evidence).toMatchObject([{ id: "a1", application_mode: "code", evidence_basis: "no-source-passage", reason: "Drafting judgment.", sources: [] }])
    expect(finalize[1]!.text).toContain(`The drafting notes do not mention these open items: ${marker.id}.`)

    expect(state.status).toBe("completed")
    const gate = gateOf(directory)
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gate.disclosed).toEqual([expect.objectContaining({ id: marker.id, kind: "placeholder", status: "disclosed" })])
  }, 120_000)
})

describe("prompt version", () => {
  test("the contract change is lean-18-advised-questions, and a run keyed under lean-17-advice-scope re-runs every model step", async () => {
    expect(LEAN_PROMPT_VERSION).toBe("lean-18-advised-questions")
    const directory = matter()
    expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
    const idle = scriptedModel({})
    expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
    expect(idle.prompts).toHaveLength(0)

    const state = readLeanState(directory)!
    writeLeanState(directory, {
      ...state,
      stages: Object.fromEntries(Object.entries(state.stages).map(([step, row]) =>
        [step, { ...row!, input_sha: hashStageInputs(step as StepID, "lean-17-advice-scope", stageInputs(directory, step as StepID)) }])),
    })
    const rerun = scriptedModel(happyPath)
    expect((await runLean(directory, { sessions: rerun.port })).status).toBe("completed")
    expect(rerun.counts().map(count => count > 0)).toEqual([true, true, true, true])
  }, 120_000)
})
