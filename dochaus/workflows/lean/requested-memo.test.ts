import { afterEach, describe, expect, test } from "bun:test"
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
import { AGENT, COVER_MEMO, DRAFT, changesFirst, cleanupMatters, drafter, happyPath, listed, matter, noChanges, pidOf, registering, reviewFirst, sameSession, scriptedModel, type Reply } from "./test-support"

// A memo the request asks for: the outline names it on a "Memo:" line in the
// request's words and plans it as a document, the notes file is named for it,
// and finalize writes the notes as that memo, each open item's id in a trailing
// "[ref: <id>]". Real coordinator, checker and gate; only the model is scripted.

afterEach(cleanupMatters)

const REQUEST = "Draft a short supply note from the supply terms, with a cover memo on the open points."
const DOCUMENT = { title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }
const DELIVERY = "The supplier shall deliver the goods."
const JUDGED = "The supplier shall deliver the goods in good condition [open item: confirm the condition standard chosen]."
const MEMO_OPENING = 'The request asks for a memo to accompany the drafted documents, "cover memo". Write the drafting notes as that memo.'
const MEMO_ORDER = "Lead with the discrepancies found and the choices made, then give the recommendations and the open points, in plain language."
const NOTES_OPENING = "This is a lean drafting run. Write the drafting notes that go to the lawyer with the drafted documents."
const MARKER_RULE = 'end that statement with the item\'s marker copied whole, exactly as open_items lists it, as a trailing "[ref: <id>]"'

/** An outline with this Memo line, planning the memo titled `memo` when given. */
const outline = (line: string, memo?: string) => `Acting for: not stated\n${line}\n# Supply note\nForm: other\n## Delivery${memo ? `\n# ${memo}\nForm: advice` : ""}`
const MEMO_DOCUMENT = { title: "Cover memo", purpose: "The cover memo", parts: [{ heading: "Decisions", scope: "Decisions taken" }], form: "advice" as const, executed_by: [] }
/** The scripted drafter whose outline reply carries this Memo line, and plans
 * the cover memo unless `planned` is false. */
const memoDrafter = (line: string, planned = true): Reply => {
  const draft = planned ? drafter(DRAFT, COVER_MEMO) : drafter(DRAFT)
  return (prompt, index) => draft(prompt, index).replace("Acting for: not stated", `Acting for: not stated\n${line}`)
}
/** Focused prose names the actual open point; code owns the memo title and
 * references. Direct legacy prompt fixtures retain their full-note reply. */
const memoReply: Reply = prompt => {
  if (prompt.includes("<section_kind>")) return prompt.includes("<section_kind>\nopen-item\n")
    ? "The condition standard is a drafting choice for the lawyer to confirm."
    : "The supply terms set the delivery period, and the draft states it."
  const ids = [...prompt.matchAll(/<open_items>\n([\s\S]*?)\n<\/open_items>/g)].flatMap(match => [...match[1]!.matchAll(/^- \[ref: (\S+)\]/gm)].map(row => row[1]!))
  return ["# Cover memo", "", "The supply terms set the delivery period, and the draft states it.", "",
    ...ids.map(id => `- The condition standard is a drafting choice for the lawyer to confirm. [ref: ${id}]`)].join("\n")
}
const notesRow = (directory: string) => readLeanDocuments(directory)?.documents.find(row => row.kind === "drafting_notes")
const finalizePrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.finalize).map(prompt => prompt.text)

describe("the outline's Memo line", () => {
  test("is read in the request's words wherever it stands, and a non-answer is no memo", () => {
    for (const [line, memo] of [["Memo: cover memo", "cover memo"], ["Memo: a cover memo", "cover memo"], ["- **Memo:** `Cover Memo`", "Cover Memo"],
      ['Memo: "cover memo on the open points"', "cover memo on the open points"], ["Memo: cover memo.", "cover memo"]])
      expect(parsePlan(outline(line, memo), REQUEST)).toEqual({ issues: [], value: expect.objectContaining({ memo }) })
    for (const line of ["Memo: none", "Memo: not requested", "Memo: N/A", "Memo: no memo", "Memo: nothing", "Memo: no", "Memo:", ""]) {
      const parsed = parsePlan(outline(line), REQUEST)
      expect(parsed.issues).toEqual([])
      expect(parsed.value).not.toHaveProperty("memo")
    }
    // Under a title, or after the parts, the line is still the run's memo, held
    // to the memo rules: here no title is the memo, so it is retry feedback and
    // the memo is added as its own document.
    for (const late of ["Acting for: not stated\n# Supply note\nForm: other\nMemo: cover memo\n## Delivery", "Acting for: not stated\n# Supply note\nForm: other\n## Delivery\nMemo: cover memo"]) {
      const parsed = parsePlan(late, REQUEST)
      expect(parsed.issues).toEqual([expect.stringContaining('The "Memo:" line names "cover memo", but no planned title is that memo')])
      expect(parsed.value).toMatchObject({ memo: "cover memo", documents: [{ title: "Supply note", purpose: "", parts: [{ heading: "Delivery", scope: "" }] }, { title: "Cover memo", form: "advice" }] })
    }
  })

  test("a Memo line after the first title plans the memo: the title naming it is the memo, else the one advice document is adopted", () => {
    const named = parsePlan("Acting for: not stated\n# Supply note\nForm: other\n## Delivery\n# Cover memo\nForm: advice\n## Decisions\nMemo: cover memo", REQUEST)
    expect(named.issues).toEqual([])
    expect(named.value).toMatchObject({ memo: "cover memo", documents: [{ title: "Supply note" }, { title: "Cover memo", form: "advice", parts: [{ heading: "Decisions", scope: "" }] }] })
    // No title is the memo: retry feedback, and on the last attempt the one document planned as advice is the memo.
    const adopted = parsePlan("Acting for: not stated\n# Partner note\nForm: advice\n## Points\n# Supply note\nForm: other\nMemo: cover memo\n## Delivery", REQUEST)
    expect(adopted.issues).toEqual([expect.stringContaining('The "Memo:" line names "cover memo", but no planned title is that memo; the planned titles are "Partner note", "Supply note".')])
    expect(adopted.value).toMatchObject({ memo: "cover memo", documents: [{ title: "Supply note" }, { title: "Partner note", form: "advice", executed_by: [] }] })
  })

  test("a name is a non-answer only as a whole, the first line wins, and without the request it is not grounded", () => {
    expect(parsePlan(outline("Memo: Unclear points memo", "Unclear points memo"), "Draft a short supply note, with an unclear points memo for the lawyer."))
      .toEqual({ issues: [], value: expect.objectContaining({ memo: "Unclear points memo" }) })
    expect(parsePlan("Acting for: not stated\nMemo: cover memo\nMemo: open points\n# Supply note\nForm: other\n## Delivery\n# Cover memo\nForm: advice", REQUEST).value?.memo).toBe("cover memo")
    expect(parsePlan(outline("Memo: advice letter", "Advice letter"))).toEqual({ issues: [], value: expect.objectContaining({ memo: "advice letter" }) })
  })

  test("a name is a title: held to whole words the request uses, whatever its length, and carrying no character a file name cannot", () => {
    // Whole words the request uses: a name is never named by a word it cuts.
    const name = `cover memo ${"x".repeat(109)}`
    const request = `${REQUEST} Title it ${name}.`
    expect(parsePlan(outline(`Memo: ${name}`, name), request)).toEqual({ issues: [], value: expect.objectContaining({ memo: name }) })
    expect(parsePlan(outline(`Memo: ${name.slice(0, -1)}`), request).issues.join(" ")).toContain("the request does not use those words")
    for (const line of ["\\", ":", "*", "?", '"', "<", ">", "|"].map(mark => `Memo: cover${mark}memo`)) {
      const parsed = parsePlan(outline(line))
      expect(parsed.issues.join(" ")).toContain("cannot be part of a file name")
      expect(parsed.value).not.toHaveProperty("memo")
    }
  })

  test("a name the request does not use, or one a file name cannot carry, is retry feedback and is dropped", () => {
    const ungrounded = parsePlan(outline("Memo: advice letter"), REQUEST)
    expect(ungrounded.issues.join(" ")).toContain('The "Memo:" line names "advice letter", but the request does not use those words.')
    // A rejected name is renamed, not a reason to drop the memo the request asks for.
    expect(ungrounded.issues.join(" ")).toContain('keep the memo planned as its "# " document. Leave the line and that document out only when the request does not ask for a memo or note.')
    expect(ungrounded.issues.join(" ")).not.toContain("otherwise leave the line out")
    expect(ungrounded.value).not.toHaveProperty("memo")
    for (const line of ["Memo: cover memo/notes", "Memo: cover-memo.docx"]) {
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
    for (const [title, file] of [["Supply note", "Supply note - cover memo.docx"], ["Supply note - Annex", "Supply note - Annex - cover memo.docx"], ["Supply note ".repeat(12), long]]) {
      writeLeanPlan(directory, [{ ...DOCUMENT, title: title!.trim() }], undefined, undefined, { name: "cover memo" })
      expect(requestedMemo(directory)).toEqual({ file: file!, name: "cover memo" })
    }
    // The memo document is not what the notes are named after.
    writeLeanPlan(directory, [DOCUMENT, MEMO_DOCUMENT], undefined, undefined, { name: "cover memo", document: "d2" })
    expect(requestedMemo(directory)).toEqual({ file: "Supply note - cover memo.docx", name: "cover memo" })
  })

  test("without a combined document, is named after the first document with a requested file name, else the first document", () => {
    const letter = { ...DOCUMENT, title: "Covering letter" }
    expect(memoNotesFile([DOCUMENT, { ...letter, file: "covering-letter.docx" }], "cover memo")).toBe("covering-letter - cover memo.docx")
    expect(memoNotesFile([DOCUMENT, letter], "cover memo")).toBe("Supply note - cover memo.docx")
    expect(memoNotesFile([], "cover memo")).toBe("Draft - cover memo.docx")
  })

  test("passes the delivery selection's file-name check, and selection keeps the memo, its file following the first document; a notes name names the record", () => {
    const directory = matter()
    const file = memoNotesFile([DOCUMENT], "cover memo")
    const instruction = `Deliver the note as Supply Final.docx, the advice as supply-advice.docx and the memo as ${file}`
    writeLeanText(directory, "request.md", `${instruction}.`)
    const notes = { document: null, role: "drafting_notes" as const, instruction_quote: instruction }
    expect(checkDeliverySelections(directory, [{ ...notes, file }])).toEqual([])

    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo" })
    expect(selectedDeliveryPlan(directory, [])).toBeUndefined()
    const body = selectedDeliveryPlan(directory, [{ document: "d1", role: "body", file: "Supply Final.docx", instruction_quote: instruction }])!
    expect(body.documents[0]?.file).toBe("Supply Final.docx")
    expect(body.memo).toBe("cover memo")
    expect(requestedMemo(directory, body)?.file).toBe("Supply Final - cover memo.docx")
    expect(selectedDeliveryPlan(directory, [{ ...notes, file: "supply-advice.docx" }])).toMatchObject({ memo: "cover memo", record_file: "supply-advice.docx" })
  })

  test("plan.json keeps the memo's name outside the operative plan, beside the memo's document, and a memo that is not a name is not read", () => {
    const directory = matter()
    const plan = writeLeanPlan(directory, [DOCUMENT, MEMO_DOCUMENT], undefined, undefined, { name: "cover memo", document: "d2" })
    expect(readLeanText(directory, "plan.json")).toBe(`${JSON.stringify({ schema: plan.schema, documents: plan.documents, acting_for: undefined, basis: undefined, memo: "cover memo", memo_document: "d2" }, null, 2)}\n`)
    expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", memo_document: "d2" })
    const { memo: _, ...unnamed } = plan
    expect(operativePlan(readLeanPlan(directory))).toEqual(operativePlan(unnamed))
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
    expect(stagePrompt(directory, "d1.draft", { outline: true })).toContain("When the request asks for a memo or note alongside the documents, whatever its subject, write one line \"Memo: <its name in the request's own words>\" before the first \"# \" title and plan that memo as a \"# \" document with that title and \"Form: advice\"")

    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo" })
    const memo = stagePrompt(directory, "finalize")
    expect(memo).toContain(MEMO_OPENING)
    expect(memo).toContain(MEMO_ORDER)
    expect(memo).toContain(MARKER_RULE)
    // Only the opening and the notes file's planned name differ; every other
    // contract line, the ref-marker rule included, stays.
    const memoLines = new Set(memo.split("\n"))
    expect(notes.split("\n").filter(line => !memoLines.has(line))).toEqual([NOTES_OPENING, expect.stringContaining('"planned_filename": "Supply note - drafting notes.docx"')])
    expect(stageInputs(directory, "finalize").map(row => row.name)).toEqual([...notesInputs.slice(0, notesInputs.indexOf("drafting-notes-contract") + 1), "requested-memo-contract", ...notesInputs.slice(notesInputs.indexOf("drafting-notes-contract") + 1)])
  })

  test("with or without a memo the notes check accepts each id in a ref marker and asks for one in its retry", () => {
    const directory = matter()
    const item = openItem("apply", "placeholder", "Confirm the condition standard chosen.", { document: "d1" })
    const retry = `do not mention these open items: ${item.id}. Say what each is and what the lawyer needs to do about it, ending that statement with its exact id as "[ref: <id>]".`
    writeLeanPlan(directory, [DOCUMENT])
    expect(checkDraftingNotes(directory, "# Drafting notes\n\nThe draft states the delivery period.\n", [item]).issues.join(" ")).toContain(retry)
    expect(checkDraftingNotes(directory, `# Drafting notes\n\n- The condition standard is for the lawyer to confirm. [ref: ${item.id}]\n`, [item]).issues).toEqual([])
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo" })
    expect(checkDraftingNotes(directory, "# Cover memo\n\nThe draft states the delivery period.\n", [item]).issues.join(" ")).toContain(retry)
    expect(checkDraftingNotes(directory, `# Cover memo\n\n- The condition standard is for the lawyer to confirm. [ref: ${item.id}]\n`, [item]).issues).toEqual([])
  })
})

describe("a requested memo through the run", () => {
  test("a grounded Memo line names the notes at plan time, finalize writes the memo with ref markers, and the gate settles each id", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", REQUEST)
    const first = new Map<string | undefined, string>()
    const notes: Reply = (prompt, index, sessionID) => {
      if (!first.has(sessionID)) {
        first.set(sessionID, prompt)
        if (prompt.includes("<section_kind>\nopen-item\n"))
          return "The condition standard is a drafting choice for the lawyer to confirm. [ref: model-invented]"
      }
      return memoReply(first.get(sessionID)!, index)
    }
    // The memo document has no supply paragraph to review or change.
    const changes = sameSession((prompt, index) => listed(prompt, DELIVERY)
      ? JSON.stringify({ amendments: [{ pid: pidOf(prompt, DELIVERY), updated: JUDGED, reason: "Drafting judgment.", sources: [] }], questions: [] }) : noChanges(prompt, index))
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter("Memo: cover memo"), [AGENT.review]: reviewFirst, [AGENT.changes]: changes, [AGENT.finalize]: notes })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", memo_document: "d2" })
    expect(readLeanPlan(directory)?.documents.at(-1)).toMatchObject({ title: "Cover memo", form: "advice", executed_by: [] })
    expect(requestedMemo(directory)?.file).toBe("Supply note - cover memo.docx")
    const marker = readOpenItems(directory).find(item => item.kind === "placeholder" && item.stage === "apply")!
    const prompts = finalizePrompts(model)
    for (const prompt of prompts.filter(prompt => prompt.includes("<section_kind>"))) {
      expect(prompt).toContain("<notes_role>\nThese sections form the requested accompanying memo, titled cover memo.")
      expect(prompt).toContain(`<current_request>\n${REQUEST}`)
      expect(prompt).not.toContain(MARKER_RULE)
    }
    const retry = model.prompts.find(prompt => prompt.agent === AGENT.finalize && prompt.text.includes("Correct only this section's errors"))!
    expect(retry.text).toContain("The coordinator attaches its reference markers")
    const sameSection = model.prompts.filter(prompt => prompt.agent === AGENT.finalize && prompt.sessionID === retry.sessionID)
    expect(sameSection).toHaveLength(2)
    expect(sameSection[0]!.text).toContain("<section_kind>\nopen-item\n")
    expect(sameSection[0]!.text).not.toContain(marker.id)
    const saved = readLeanText(directory, "drafting-notes.md")!
    expect(saved).toStartWith("# cover memo\n")
    expect(saved.split(`[ref: ${marker.id}]`)).toHaveLength(2)
    expect(saved).toContain(`The condition standard is a drafting choice for the lawyer to confirm. [ref: ${marker.id}]`)
    expect(saved).not.toContain("model-invented")
    expect(readLeanJson(directory, "gate.json") as GateResult).toMatchObject({ status: "pass", blocking: [] })
    expect((readLeanJson(directory, "gate.json") as GateResult).disclosed).toEqual([expect.objectContaining({ id: marker.id, status: "disclosed" })])
    expect(notesRow(directory)?.deliverable).toBe("Supply note - cover memo.docx")
    expect(deliverables.names).toContain("Supply note - cover memo.docx")
  }, 120_000)

  test("the plan names the notes after the combined file, not an earlier document without a file", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a delivery schedule, a supply note from the supply terms and a cover memo, and send the note and memo together as supply-pack.docx.")
    const draft = drafter("# Delivery schedule\n\n## Dates\n\nThe delivery dates.\n", DRAFT, COVER_MEMO)
    const outlined: Reply = (prompt, index) => {
      const reply = draft(prompt, index)
      return reply.startsWith("Acting for:")
        ? reply.replace("Acting for: not stated", "Acting for: not stated\nMemo: cover memo").replace("# Supply note", "# Supply note\nFile: supply-pack.docx\nDeliverable: combined")
        : reply
    }
    const model = scriptedModel({ [AGENT.draft]: outlined })
    expect((await runDraftStage(directory, readLeanState(directory)!, { sessions: model.port }, () => {}, "d1", () => {})).status).toBe("done")
    expect(readLeanPlan(directory)?.documents.map(row => [row.title, row.file, row.deliverable])).toEqual([["Delivery schedule", undefined, undefined], ["Supply note", "supply-pack.docx", "combined"], ["Cover memo", undefined, undefined]])
    expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", memo_document: "d3" })
    expect(requestedMemo(directory)?.file).toBe("supply-pack - cover memo.docx")
  }, 60_000)

  test("a requested memo whose name the request does not use is renamed on retry and stays planned", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", REQUEST)
    const [joined, renamed] = [memoDrafter("Memo: cover memo for the client"), memoDrafter("Memo: cover memo")]
    const outlined: Reply = (prompt, index) => (prompt.includes("does not use those words") ? renamed : joined)(prompt, index)
    // The memo document has no supply paragraph to review or change.
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: outlined, [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst, [AGENT.finalize]: memoReply })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: registering().registerDeliverable })).status).toBe("completed")

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes("does not use those words"))).toHaveLength(1)
    expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", memo_document: "d2" })
    expect(readLeanPlan(directory)?.documents.at(-1)).toMatchObject({ title: "Cover memo", form: "advice", executed_by: [] })
  }, 120_000)

  test("a Memo line the request does not support is retried, then dropped, and the run keeps its drafting notes", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms.")
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: memoDrafter("Memo: cover memo", false) })
    expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")

    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes('"Memo:" line names "cover memo"'))).toHaveLength(LEAN_MAX_RETRIES)
    expect(readLeanPlan(directory)).not.toHaveProperty("memo_document")
    expect(readLeanPlan(directory)).not.toHaveProperty("memo")
    const [prompt] = finalizePrompts(model)
    expect(prompt).toContain("<notes_role>\nThese sections form the drafting notes for the lawyer accompanying the drafted documents.")
    expect(prompt).not.toContain("These sections form the requested accompanying memo")
    expect(readLeanText(directory, "drafting-notes.md")).toStartWith("# Drafting notes\n")
    expect(notesRow(directory)?.deliverable).toBe("Supply note - drafting notes.docx")
  }, 120_000)
})

describe("memo mode keys on the plan's memo name, not its notes file", () => {
  const INSTRUCTION = "Deliver the note as Supply Final.docx and the notes as supply-advice.docx"
  test("a notes file a delivery selection names, without a memo, names the record and leaves finalize as without a memo", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [DOCUMENT])
    const notes = stagePrompt(directory, "finalize")
    const notesInputs = stageInputs(directory, "finalize").map(row => row.name)

    const selection = { document: null, role: "drafting_notes" as const, file: "supply-advice.docx", instruction_quote: INSTRUCTION }
    const saved = selectedDeliveryPlan(directory, [selection])!
    expect(saved).not.toHaveProperty("memo")
    writeLeanText(directory, "plan.json", `${JSON.stringify(saved, null, 2)}\n`)
    expect(readLeanPlan(directory)).toMatchObject({ record_file: "supply-advice.docx" })
    expect(selectedDeliveryPlan(directory, [selection])).toBeUndefined()
    expect(requestedMemo(directory)).toBeUndefined()
    // Only the notes file's planned name differs from the run without one.
    expect(stagePrompt(directory, "finalize")).toBe(notes.replace('"planned_filename": "Supply note - drafting notes.docx"', '"planned_filename": "supply-advice.docx"'))
    expect(stageInputs(directory, "finalize").map(row => row.name)).toEqual(notesInputs)
    expect(notesInputs).not.toContain("requested-memo-contract")
  })

  test("a memo plan adds the memo contract to finalize only", () => {
    const directory = matter()
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo" })
    expect(stageInputs(directory, "d1.apply").map(row => row.name)).not.toContain("requested-memo-contract")
    expect(stageInputs(directory, "finalize").map(row => row.name)).toContain("requested-memo-contract")
  })

  test("a memo name containing ' - ' is the plan's name exactly", () => {
    const directory = matter()
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover - memo" })
    expect(requestedMemo(directory)).toEqual({ file: "Supply note - cover - memo.docx", name: "cover - memo" })
  })

  test("a selection renaming the first document moves the memo's file with it; a selected notes name names the record and keeps the memo", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    const rename = { document: "d1" as const, role: "body" as const, file: "Supply Final.docx", instruction_quote: INSTRUCTION }
    const notes = { document: null, role: "drafting_notes" as const, file: "supply-advice.docx", instruction_quote: INSTRUCTION }
    writeLeanPlan(directory, [DOCUMENT], undefined, undefined, { name: "cover memo" })
    const renamed = selectedDeliveryPlan(directory, [rename])!
    expect(renamed.memo).toBe("cover memo")
    expect(requestedMemo(directory, renamed)?.file).toBe("Supply Final - cover memo.docx")
    expect(selectedDeliveryPlan(directory, [notes])).toMatchObject({ memo: "cover memo", record_file: "supply-advice.docx" })
    expect(selectedDeliveryPlan(directory, [rename, notes])).toEqual({ ...renamed, record_file: "supply-advice.docx" })
  })

  test("a selection renaming the document the memo is named after, not the first, moves the memo's file with it", () => {
    const directory = matter()
    const documents = [DOCUMENT, { ...DOCUMENT, title: "Covering letter", file: "covering-letter.docx" }]
    writeLeanPlan(directory, documents, undefined, undefined, { name: "cover memo" })
    const rename = { document: "d2" as const, role: "body" as const, file: "Covering Final.docx", instruction_quote: INSTRUCTION }
    expect(requestedMemo(directory, selectedDeliveryPlan(directory, [rename]))?.file).toBe("Covering Final - cover memo.docx")
  })
})
