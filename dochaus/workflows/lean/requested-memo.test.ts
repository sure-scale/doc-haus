import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { checkDraftingNotes } from "./checks"
import { memoNotesFile, readLeanDocuments, requestedMemo } from "./deliverables"
import { checkDeliverySelections, selectedDeliveryPlan } from "./delivery-selection"
import { runDraftStage } from "./draft-sections"
import type { GateResult } from "./gate"
import { operativePlan, parsePlan, readLeanPlan, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { openItem } from "./quotes"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES, stageInputs } from "./stages"
import { readLeanJson, readLeanState, readLeanText, readOpenItems, writeLeanText } from "./state"
import { AGENT, DRAFT, cleanupMatters, drafter, happyPath, matter, pidOf, registering, sameSession, scriptedModel, type Reply } from "./test-support"

// A memo the request asks for in place of the drafting notes: the outline names
// it on a "Memo:" line in the request's words, the plan names the notes file
// for it, and finalize writes the notes as that memo, each open item's id in a
// trailing "[ref: <id>]". Real coordinator, checker and gate; only the model is
// scripted.

afterEach(cleanupMatters)

const REQUEST = "Draft a short supply note from the supply terms, with a cover memo on the open points."
const DOCUMENT = { title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }
const DELIVERY = "The supplier shall deliver the goods."
const JUDGED = "The supplier shall deliver the goods in good condition [open item: confirm the condition standard chosen]."
const MEMO_OPENING = 'The request asks for a memo to accompany the drafted documents, "cover memo". Write that memo.'
const MEMO_ORDER = "Lead with the discrepancies found and the choices made, then give the recommendations and the open points, in plain language."
const NOTES_OPENING = "This is a lean drafting run. Write the drafting notes that go to the lawyer with the drafted documents."
const MARKER_RULE = 'end that statement with the item\'s exact id as a trailing "[ref: <id>]"'
const FINAL_NOT_PROCESS = "Describe the final documents and decisions, not how this run produced them."
/** How plain notes describe review changes and the redline's coverage. */
const REVIEW_CHANGE_RULES = ["Identify any requested comparison not covered by that baseline as outstanding.", "Use change_origin to distinguish source-to-initial-draft revisions from later review amendments"]
const NEUTRAL_ACTOR = ["Say what each open item is and what needs to be done about it", "say what was chosen and that the choice needs confirmation"]

const outline = (line: string) => `Acting for: not stated\n${line}\n# Supply note\n## Delivery`
/** The scripted drafter whose outline reply carries this Memo line. */
const memoDrafter = (line: string): Reply => {
  const draft = drafter(DRAFT)
  return (prompt, index) => draft(prompt, index).replace("Acting for: not stated", `Acting for: not stated\n${line}`)
}
/** A memo naming every listed open item in a trailing ref marker. */
const memoReply: Reply = prompt => {
  const ids = [...prompt.matchAll(/<open_items>\n([\s\S]*?)\n<\/open_items>/g)].flatMap(match => [...match[1]!.matchAll(/^- (\S+):/gm)].map(row => row[1]!))
  return ["# Cover memo", "", "The supply terms set the delivery period, and the draft states it.", "",
    ...ids.map(id => `- The condition standard is a drafting choice for the lawyer to confirm. [ref: ${id}]`)].join("\n")
}
const notesRow = (directory: string) => readLeanDocuments(directory)?.documents.find(row => row.kind === "drafting_notes")
const finalizePrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.finalize).map(prompt => prompt.text)

describe("the outline's Memo line", () => {
  test("is read before the first title in the request's words, and a non-answer is no memo", () => {
    for (const [line, memo] of [["Memo: cover memo", "cover memo"], ["Memo: a cover memo", "cover memo"], ["- **Memo:** `Cover Memo`", "Cover Memo"],
      ['Memo: "cover memo on the open points"', "cover memo on the open points"], ["Memo: cover memo.", "cover memo"]])
      expect(parsePlan(outline(line), REQUEST)).toEqual({ issues: [], value: expect.objectContaining({ memo }) })
    for (const line of ["Memo: none", "Memo: not requested", "Memo: N/A", "Memo: no memo", "Memo: nothing", "Memo: no", "Memo:", ""]) {
      const parsed = parsePlan(outline(line), REQUEST)
      expect(parsed.issues).toEqual([])
      expect(parsed.value).not.toHaveProperty("memo")
    }
    // Under a title the line is that document's purpose, not the run's memo.
    const under = parsePlan("Acting for: not stated\n# Supply note\nMemo: cover memo\n## Delivery", REQUEST)
    expect(under.value).not.toHaveProperty("memo")
    expect(under.value?.documents[0]?.purpose).toBe("Memo: cover memo")
  })

  test("a name is a non-answer only as a whole, the first line wins, and without the request it is not grounded", () => {
    expect(parsePlan(outline("Memo: Unclear points memo"), "Draft a short supply note, with an unclear points memo for the lawyer."))
      .toEqual({ issues: [], value: expect.objectContaining({ memo: "Unclear points memo" }) })
    expect(parsePlan("Acting for: not stated\nMemo: cover memo\nMemo: open points\n# Supply note\n## Delivery", REQUEST).value?.memo).toBe("cover memo")
    expect(parsePlan(outline("Memo: advice letter"))).toEqual({ issues: [], value: expect.objectContaining({ memo: "advice letter" }) })
  })

  test("a name fits in 60 characters and carries no character a file name cannot", () => {
    // Whole words the request uses: a name is never named by a word it cuts.
    const name = `cover memo ${"x".repeat(49)}`
    const long = `${name}x`
    const request = `${REQUEST} Title it ${name}.`
    expect(parsePlan(outline(`Memo: ${name}`), request)).toEqual({ issues: [], value: expect.objectContaining({ memo: name }) })
    expect(parsePlan(outline(`Memo: ${name.slice(0, -1)}`), request).issues.join(" ")).toContain("the request does not use those words")
    for (const line of [`Memo: ${long}`, ...["\\", ":", "*", "?", '"', "<", ">", "|"].map(mark => `Memo: cover${mark}memo`)]) {
      const parsed = parsePlan(outline(line))
      expect(parsed.issues.join(" ")).toContain("cannot be part of a file name")
      expect(parsed.value).not.toHaveProperty("memo")
    }
  })

  test("a name the request does not use, or one a file name cannot carry, is retry feedback and is dropped", () => {
    const ungrounded = parsePlan(outline("Memo: advice letter"), REQUEST)
    expect(ungrounded.issues.join(" ")).toContain('The "Memo:" line names "advice letter", but the request does not use those words.')
    expect(ungrounded.value).not.toHaveProperty("memo")
    for (const line of ["Memo: cover memo/notes", "Memo: cover-memo.docx", `Memo: ${"cover memo ".repeat(7)}`]) {
      const parsed = parsePlan(outline(line), REQUEST)
      expect(parsed.issues.join(" ")).toContain("cannot be part of a file name")
      expect(parsed.value).not.toHaveProperty("memo")
    }
  })
})

describe("the requested memo's notes file", () => {
  test("is the first document's delivery stem and the memo name, which the prompt recovers", () => {
    expect(memoNotesFile([DOCUMENT], "cover memo")).toBe("Supply note - cover memo.docx")
    expect(memoNotesFile([{ ...DOCUMENT, file: "supply-note-final.docx" }, { ...DOCUMENT, title: "Covering letter" }], "cover memo")).toBe("supply-note-final - cover memo.docx")
    // A long stem is shortened so the memo name survives the name limit.
    const long = memoNotesFile([{ ...DOCUMENT, title: "Supply note ".repeat(12) }], "cover memo")
    expect(long).toBe(`${"Supply note ".repeat(9).trim()} - cover memo.docx`)

    const directory = matter()
    for (const [title, file, name] of [["Supply note", "Supply note - cover memo.docx", "cover memo"], ["Supply note - Annex", "Supply note - Annex - cover memo.docx", "cover memo"],
      ["Supply note ".repeat(12), long, "cover memo"], ["Supply note", "supply-advice.docx", "supply-advice"]]) {
      writeLeanPlan(directory, [{ ...DOCUMENT, title: title!.trim() }], undefined, undefined, { name: name!, file: file! })
      expect(requestedMemo(directory)).toEqual({ file: file!, name: name! })
    }
  })

  test("without a combined document, is named after the first document with a requested file name, else the first document", () => {
    const letter = { ...DOCUMENT, title: "Covering letter" }
    expect(memoNotesFile([DOCUMENT, { ...letter, file: "covering-letter.docx" }], "cover memo")).toBe("covering-letter - cover memo.docx")
    expect(memoNotesFile([DOCUMENT, letter], "cover memo")).toBe("Supply note - cover memo.docx")
    expect(memoNotesFile([], "cover memo")).toBe("Draft - cover memo.docx")
  })

  test("passes the delivery selection's file-name check, and selection keeps the memo, its file following the first document unless it selects a notes name", () => {
    const directory = matter()
    const file = memoNotesFile([DOCUMENT], "cover memo")
    const instruction = `Deliver the note as Supply Final.docx, the advice as supply-advice.docx and the memo as ${file}`
    writeLeanText(directory, "request.md", `${instruction}.`)
    const notes = { document: null, role: "drafting_notes" as const, instruction_quote: instruction }
    expect(checkDeliverySelections(directory, [{ ...notes, file }])).toEqual([])

    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo", file })
    expect(selectedDeliveryPlan(directory, [])).toBeUndefined()
    const body = selectedDeliveryPlan(directory, [{ document: "d1", role: "body", file: "Supply Final.docx", instruction_quote: instruction }])!
    expect(body.documents[0]?.file).toBe("Supply Final.docx")
    expect(body).toMatchObject({ memo: "cover memo", notes_file: "Supply Final - cover memo.docx" })
    expect(selectedDeliveryPlan(directory, [{ ...notes, file: "supply-advice.docx" }])?.notes_file).toBe("supply-advice.docx")
  })

  test("plan.json keeps the memo's name beside its notes file, outside the operative plan, and a memo that is not a name is not read", () => {
    const directory = matter()
    const file = memoNotesFile([DOCUMENT], "cover memo")
    const plan = writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo", file })
    expect(readLeanText(directory, "plan.json")).toBe(`${JSON.stringify({ schema: plan.schema, documents: plan.documents, acting_for: undefined, basis: undefined, memo: "cover memo", notes_file: file }, null, 2)}\n`)
    expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", notes_file: file })
    expect(operativePlan(readLeanPlan(directory))).toEqual(operativePlan(writeLeanPlan(matter(), [DOCUMENT])))
    for (const memo of ["", 5]) {
      writeLeanText(directory, "plan.json", JSON.stringify({ ...plan, memo }))
      expect(readLeanPlan(directory)).toBeUndefined()
    }
  })
})

describe("finalize in memo mode", () => {
  test("writes the notes as the memo; with or without a memo the notes carry ref markers, and without one the plan and finalize inputs are as before", () => {
    const directory = matter()
    const plan = writeLeanPlan(directory, [DOCUMENT])
    expect(readLeanText(directory, "plan.json")).toBe(`${JSON.stringify({ schema: plan.schema, documents: plan.documents, acting_for: undefined, basis: undefined }, null, 2)}\n`)
    const notes = stagePrompt(directory, "finalize")
    const notesInputs = stageInputs(directory, "finalize").map(row => row.name)
    expect(notes).toContain(NOTES_OPENING)
    expect(notes).toContain(MARKER_RULE)
    expect(notesInputs).not.toContain("requested-memo-contract")
    expect(stagePrompt(directory, "d1.draft", { outline: true })).toContain("Write one line \"Memo: <its name in the request's own words>\" before the first \"# \" title only when the request asks for a memo or note to accompany the documents")

    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo", file: memoNotesFile([DOCUMENT], "cover memo") })
    const memo = stagePrompt(directory, "finalize")
    expect(memo).toContain(MEMO_OPENING)
    expect(memo).toContain(MEMO_ORDER)
    expect(memo).toContain(MARKER_RULE)
    // Only the opening, the notes file's planned name, the plain notes' review-change
    // rules and the reply's name differ; every other contract line, the ref-marker
    // rule included, stays.
    const memoLines = new Set(memo.split("\n"))
    expect(notes.split("\n").filter(line => !memoLines.has(line))).toEqual([NOTES_OPENING, expect.stringContaining('"planned_filename": "Supply note - drafting notes.docx"'),
      ...REVIEW_CHANGE_RULES.map(rule => expect.stringContaining(rule)), expect.stringContaining("Reply with the drafting notes"), expect.stringContaining("saved as the drafting notes")])
    expect(stageInputs(directory, "finalize").map(row => row.name)).toEqual([...notesInputs.slice(0, notesInputs.indexOf("drafting-notes-contract") + 1), "requested-memo-contract", ...notesInputs.slice(notesInputs.indexOf("drafting-notes-contract") + 1)])
  })

  test("a memo describes the final documents, not how this run produced them; plain notes keep their review-change rules; neither names who acts", () => {
    const directory = matter()
    writeLeanPlan(directory, [DOCUMENT])
    const notes = stagePrompt(directory, "finalize")
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo", file: memoNotesFile([DOCUMENT], "cover memo") })
    const memo = stagePrompt(directory, "finalize")
    expect(memo).toContain(FINAL_NOT_PROCESS)
    expect(notes).not.toContain(FINAL_NOT_PROCESS)
    for (const rule of REVIEW_CHANGE_RULES) {
      expect(notes).toContain(rule)
      expect(memo).not.toContain(rule)
    }
    expect(memo).not.toContain("drafting notes")
    // The delivery map names each planned file, the reply's own included, and
    // says nothing of its status or registration.
    const delivery = (prompt: string) => JSON.parse(/<delivery_artifacts>\n([\s\S]*?)\n<\/delivery_artifacts>/.exec(prompt)![1]!) as Record<string, unknown>[]
    expect(delivery(memo)).toContainEqual({ kind: "drafting_notes", document: null, planned_filename: memoNotesFile([DOCUMENT], "cover memo") })
    for (const prompt of [notes, memo]) {
      for (const row of delivery(prompt)) expect(Object.keys(row)).toEqual(["kind", "document", "planned_filename"])
      for (const words of ["pending rendering", "not yet written", "registration"]) expect(prompt).not.toContain(words)
    }
    const agent = readFileSync(new URL("../../agent/lean-drafting-notes.md", import.meta.url), "utf8")
    for (const text of [notes, memo, agent])
      for (const words of ["the lawyer must confirm", "what the lawyer needs to do", "lawyer confirmation"]) expect(text).not.toContain(words)
    for (const prompt of [notes, memo]) for (const rule of NEUTRAL_ACTOR) expect(prompt).toContain(rule)
  })

  test("with or without a memo the notes check accepts each id in a ref marker and asks for one in its retry", () => {
    const directory = matter()
    const item = openItem("apply", "placeholder", "Confirm the condition standard chosen.", { document: "d1" })
    // The retry names the reply as the prompt does.
    const retry = `mention these open items: ${item.id}. Say what each is and what needs to be done about it, ending that statement with its exact id as "[ref: <id>]".`
    writeLeanPlan(directory, [DOCUMENT])
    expect(checkDraftingNotes(directory, "# Drafting notes\n\nThe draft states the delivery period.\n", [item]).issues.join(" ")).toContain(`The drafting notes do not ${retry}`)
    expect(checkDraftingNotes(directory, `# Drafting notes\n\n- The condition standard is for the lawyer to confirm. [ref: ${item.id}]\n`, [item]).issues).toEqual([])
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo", file: memoNotesFile([DOCUMENT], "cover memo") })
    expect(checkDraftingNotes(directory, "# Cover memo\n\nThe draft states the delivery period.\n", [item]).issues.join(" ")).toContain(`The memo does not ${retry}`)
    expect(checkDraftingNotes(directory, `# Cover memo\n\n- The condition standard is for the lawyer to confirm. [ref: ${item.id}]\n`, [item]).issues).toEqual([])
  })
})

describe("a requested memo through the run", () => {
  test("a grounded Memo line names the notes at plan time, finalize writes the memo with ref markers, and the gate settles each id", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", REQUEST)
    const first = { prompt: "" }
    const notes: Reply = (prompt, index) => {
      first.prompt ||= prompt
      return index === 0 ? "# Cover memo\n\nThe draft follows the supply terms." : memoReply(first.prompt, index)
    }
    const changes = sameSession(prompt => JSON.stringify({ amendments: [{ pid: pidOf(prompt, DELIVERY), updated: JUDGED, reason: "Drafting judgment.", sources: [] }], questions: [] }))
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter("Memo: cover memo"), [AGENT.changes]: changes, [AGENT.finalize]: notes })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", notes_file: "Supply note - cover memo.docx" })
    const marker = readOpenItems(directory).find(item => item.kind === "placeholder" && item.stage === "apply")!
    const [prompt, retry] = finalizePrompts(model)
    expect(prompt).toContain(MEMO_OPENING)
    expect(prompt).toContain(MARKER_RULE)
    expect(retry).toContain(`The memo does not mention these open items: ${marker.id}. Say what each is`)
    expect(readLeanText(directory, "drafting-notes.md")).toContain(`[ref: ${marker.id}]`)
    expect(readLeanJson(directory, "gate.json") as GateResult).toMatchObject({ status: "pass", blocking: [] })
    expect((readLeanJson(directory, "gate.json") as GateResult).disclosed).toEqual([expect.objectContaining({ id: marker.id, status: "disclosed" })])
    expect(notesRow(directory)?.deliverable).toBe("Supply note - cover memo.docx")
    expect(deliverables.names).toContain("Supply note - cover memo.docx")
  }, 120_000)

  test("the plan names the notes after the combined file, not an earlier document without a file", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a delivery schedule, a supply note from the supply terms and a cover memo, and send the note and memo together as supply-pack.docx.")
    const draft = drafter("# Delivery schedule\n\n## Dates\n\nThe delivery dates.\n", DRAFT)
    const outlined: Reply = (prompt, index) => {
      const reply = draft(prompt, index)
      return reply.startsWith("Acting for:")
        ? reply.replace("Acting for: not stated", "Acting for: not stated\nMemo: cover memo").replace("# Supply note", "# Supply note\nFile: supply-pack.docx\nDeliverable: combined")
        : reply
    }
    const model = scriptedModel({ [AGENT.draft]: outlined })
    expect((await runDraftStage(directory, readLeanState(directory)!, { sessions: model.port }, () => {}, "d1", () => {})).status).toBe("done")
    expect(readLeanPlan(directory)?.documents.map(row => [row.title, row.file, row.deliverable])).toEqual([["Delivery schedule", undefined, undefined], ["Supply note", "supply-pack.docx", "combined"]])
    expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", notes_file: "supply-pack - cover memo.docx" })
  }, 60_000)

  test("a Memo line the request does not support is retried, then dropped, and the run keeps its drafting notes", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms.")
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter("Memo: cover memo") })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes('"Memo:" line names "cover memo"'))).toHaveLength(LEAN_MAX_RETRIES)
    expect(readLeanPlan(directory)).not.toHaveProperty("notes_file")
    expect(readLeanPlan(directory)).not.toHaveProperty("memo")
    const [prompt] = finalizePrompts(model)
    expect(prompt).toContain(NOTES_OPENING)
    expect(prompt).toContain(MARKER_RULE)
    expect(notesRow(directory)?.deliverable).toBe("Supply note - drafting notes.docx")
  }, 120_000)
})

describe("memo mode keys on the plan's memo name, not its notes file", () => {
  const INSTRUCTION = "Deliver the note as Supply Final.docx and the notes as supply-advice.docx"
  /** Writes the plan package review selects, as the run saves it. */
  const saveSelected = (directory: string, selections: Parameters<typeof selectedDeliveryPlan>[1]) => {
    const selected = selectedDeliveryPlan(directory, selections)!
    writeLeanText(directory, "plan.json", `${JSON.stringify(selected, null, 2)}\n`)
    return selected
  }

  test("a notes file a delivery selection names, without a memo, leaves finalize as without a memo", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [DOCUMENT])
    const notes = stagePrompt(directory, "finalize")
    const notesInputs = stageInputs(directory, "finalize").map(row => row.name)

    expect(saveSelected(directory, [{ document: null, role: "drafting_notes", file: "supply-advice.docx", instruction_quote: INSTRUCTION }])).not.toHaveProperty("memo")
    expect(readLeanPlan(directory)).toMatchObject({ notes_file: "supply-advice.docx" })
    expect(requestedMemo(directory)).toBeUndefined()
    // Only the notes file's planned name differs from the run without one.
    expect(stagePrompt(directory, "finalize")).toBe(notes.replace('"planned_filename": "Supply note - drafting notes.docx"', '"planned_filename": "supply-advice.docx"'))
    expect(stageInputs(directory, "finalize").map(row => row.name)).toEqual(notesInputs)
    expect(stageInputs(directory, "finalize").map(row => row.name)).not.toContain("requested-memo-contract")
  })

  test("a memo plan adds the memo contract to finalize only", () => {
    const directory = matter()
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo", file: memoNotesFile([DOCUMENT], "cover memo") })
    expect(stageInputs(directory, "d1.apply").map(row => row.name)).not.toContain("requested-memo-contract")
    expect(stageInputs(directory, "finalize").map(row => row.name)).toContain("requested-memo-contract")
  })

  test("a memo name containing ' - ' is the plan's name exactly", () => {
    const directory = matter()
    const file = memoNotesFile([DOCUMENT], "cover - memo")
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover - memo", file })
    expect(requestedMemo(directory)).toEqual({ file: "Supply note - cover - memo.docx", name: "cover - memo" })
  })

  test("a selection renaming the first document moves the memo's file with it; a selected notes name replaces it and keeps the memo", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    const rename = { document: "d1" as const, role: "body" as const, file: "Supply Final.docx", instruction_quote: INSTRUCTION }
    const notes = { document: null, role: "drafting_notes" as const, file: "supply-advice.docx", instruction_quote: INSTRUCTION }
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo", file: memoNotesFile([DOCUMENT], "cover memo") })
    expect(selectedDeliveryPlan(directory, [rename])).toMatchObject({ memo: "cover memo", notes_file: "Supply Final - cover memo.docx" })
    expect(selectedDeliveryPlan(directory, [notes])).toMatchObject({ memo: "cover memo", notes_file: "supply-advice.docx" })
    expect(selectedDeliveryPlan(directory, [rename, notes])).toMatchObject({ memo: "cover memo", notes_file: "supply-advice.docx" })

    // Once the memo's file is a selected name, renaming the first document leaves it.
    saveSelected(directory, [notes])
    expect(selectedDeliveryPlan(directory, [rename])).toMatchObject({ memo: "cover memo", notes_file: "supply-advice.docx" })
    expect(requestedMemo(directory)).toEqual({ file: "supply-advice.docx", name: "cover memo" })
  })

  test("a selection renaming the document the memo is named after, not the first, moves the memo's file with it", () => {
    const directory = matter()
    const documents = [DOCUMENT, { ...DOCUMENT, title: "Covering letter", file: "covering-letter.docx" }]
    writeLeanPlan(directory, documents, undefined, undefined, { name: "cover memo", file: memoNotesFile(documents, "cover memo") })
    const rename = { document: "d2" as const, role: "body" as const, file: "Covering Final.docx", instruction_quote: INSTRUCTION }
    expect(selectedDeliveryPlan(directory, [rename])).toMatchObject({ memo: "cover memo", notes_file: "Covering Final - cover memo.docx" })
  })
})
