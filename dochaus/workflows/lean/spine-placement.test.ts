import { afterEach, describe, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import { LEAN_ACTS_SCHEMA } from "./act-record"
import { runDraftStage, sectionPrompt } from "./draft-sections"
import { readLeanEvents } from "./events"
import { readLeanPlan, type PlannedDocument } from "./plan"
import { stagePrompt } from "./prompts"
import {
  PLAN_PLACEMENT_REVISIONS, SPINE_ATOMS, documentGrounds, groundLine, groundsBlock, parseEntryLevel, parseEntryPlacement, parsePartEntries, partEntriesPrompt, placementItems, placementPlanHash, planPlacementCurrent,
  planPlacementIssues, readLeanPlacement, readPartEntries, readSpineLevels, spineEntries, spineLevelKey, spineLevelPrompt, spinePlacementPrompt, withPartEntries, writePlacement,
  type EntryPlacement,
} from "./spine-placement"
import { LEAN_MAX_RETRIES, disclosableItems } from "./stages"
import { leanFile, readLeanState, readLeanText, writeLeanText } from "./state"
import { writeLeanPlan } from "./plan"
import { AGENT, cleanupMatters, labelledPart, matter, scriptedModel } from "./test-support"

// Spine placement: for each drafting-spine entry, one focused job says whether
// it binds some part's text or the package, and another, beside it, names the
// planned parts that rest on it, which bind it at either level, and each
// planned part's reverse job names the entries it rests on, which only add to
// those parts; a text
// entry of any kind placed in no part goes back to the outline once, what
// is left is recorded unplaced and disclosed, and each part prompt is told the
// entries placed in it.
// Synthetic matter only.

afterEach(cleanupMatters)

const REVISE = "Your reply did not pass the draft checks:"
const ONE = "Acting for: not stated\n# Supply note\nForm: other\n## Payment\nWhen payment falls due."
const TWO = "Acting for: not stated\n# Supply note\nForm: other\n## Delivery\nWhen the goods are delivered.\n## Payment\nWhen payment falls due."
const DELIVERY = "The supplier must deliver within thirty days of the order date."
const PAYMENT = "Payment falls due on delivery."
const SPINE = JSON.stringify({ entries: [
  { kind: "requirement", detail: DELIVERY, evidence: [{ source_id: "supply-terms.docx", quote: "within thirty days of the order date" }], limitations: [] },
  { kind: "fact", detail: PAYMENT, evidence: [{ source_id: "supply-terms.docx", quote: "Payment is due on delivery." }], limitations: [] },
], limitations: [] })

/** A limitation on the payment entry, where the spine says what the cover note must flag. */
const FLAG = "The cover note must flag that the sources give no invoice date."
const LIMITED = JSON.stringify({ ...JSON.parse(SPINE), entries: [JSON.parse(SPINE).entries[0], { ...JSON.parse(SPINE).entries[1], limitations: [FLAG] }] })

const documents: PlannedDocument[] = [{ id: "d1", title: "Supply note", purpose: "Records the supply terms.", parts: [
  { heading: "Delivery", scope: "When the goods are delivered." }, { heading: "Payment", scope: "When payment falls due." },
] }]

/** The draft agent: `first` answers the outline, `revised` any revision request, each part one paragraph. */
const outliner = (first: string, revised = first) => (text: string) =>
  /^Draft part \d+ of \d+ of "/m.test(text) ? labelledPart(text, "The terms are recorded.") : text.startsWith(REVISE) ? revised : first

/** A part's reverse job, as the placement agent answers it: the part and every entry, without quotes. */
const reverseJob = (text: string) => text.includes("<part>\n")
const tag = (text: string, name: string) => text.split(`<${name}>\n`)[1]!.split(`\n</${name}>`)[0]!.split("\n").slice(1).join("\n")

/** The placement agent: each entry rests on the planned part of the heading its
 * detail names, if planned; `reverse` answers each part's reverse job (placer's
 * names no entry). */
const placerWith = (reverse: (part: { ref: string; heading: string }, entries: { index: number; detail: string }[]) => number[]) => (text: string) => {
  if (reverseJob(text)) return JSON.stringify({ entries: reverse(JSON.parse(tag(text, "part")), tag(text, "entries").split("\n").map(line => JSON.parse(line))) })
  const planned = JSON.parse(text.split("<documents>\n")[1]!.split("\n</documents>")[0]!.split("\n").slice(1).join("\n")) as { parts: { ref: string; heading: string }[] }[]
  const entry = JSON.parse(text.split("<entry>\n")[1]!.split("\n</entry>")[0]!.split("\n").slice(1).join("\n")) as { detail: string }
  const heading = entry.detail.startsWith("The supplier must deliver") ? "Delivery" : "Payment"
  return JSON.stringify({ parts: planned.flatMap(document => document.parts.filter(part => part.heading === heading).map(part => part.ref)) })
}
const placer = placerWith(() => [])

/** A matter whose spine's entries are each atomic: the atoms are the entries. */
function prepared(spine = SPINE) {
  const directory = matter()
  writeLeanText(directory, "acts.json", JSON.stringify({ schema: LEAN_ACTS_SCHEMA, status: "none", acts: [], verify: null, deliverables: [], issues: [] }))
  writeLeanText(directory, "spine.json", spine)
  writeLeanText(directory, SPINE_ATOMS, spine)
  return { directory, state: readLeanState(directory)! }
}

/** The level agent: the delivery entry binds the package, every other entry some part's text. */
const packageDelivery = (text: string) => JSON.stringify({ level: text.includes(DELIVERY) ? "package" : "text" })

/** A spine of both audiences: a direction to the drafter, then a statement to the reader. */
const DIRECTION = "Prepare the note on the working assumption that every order is placed in writing."
const AUDIENCE = JSON.stringify({ entries: [
  { kind: "requirement", detail: DIRECTION, evidence: [], limitations: [] },
  JSON.parse(SPINE).entries[1],
], limitations: [] })
/** The level agent, answering by audience: the direction to the drafter is package, the statement to the reader text. */
const byAudience = (text: string) => JSON.stringify({ level: text.includes(DIRECTION) ? "package" : "text" })

const run = (directory: string, model: ReturnType<typeof scriptedModel>, document: "d1" | "d2" = "d1") =>
  runDraftStage(directory, readLeanState(directory)!, { sessions: model.port }, () => {}, document, () => {})
/** The sessions of the parts' reverse jobs. */
const reverseSessions = (model: ReturnType<typeof scriptedModel>) => new Set(model.sessions.flatMap((session, at) => session.title?.includes("placement: part ") ? [`ses_child${at + 1}`] : []))
const placementPrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.placement && !reverseSessions(model).has(prompt.sessionID))
const reversePrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.placement && reverseSessions(model).has(prompt.sessionID))
const levelPrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.level)
const revisions = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.startsWith(REVISE))
const partPrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.draft && /^Draft part \d+ of /m.test(prompt.text))

describe("spine entry level", () => {
  test("a reply must be {\"level\": \"text\"} or {\"level\": \"package\"}", () => {
    expect(parseEntryLevel(JSON.stringify({ level: "text" }))).toEqual({ issues: [], value: "text" })
    expect(parseEntryLevel(JSON.stringify({ level: "package" }))).toEqual({ issues: [], value: "package" })
    const shape = 'Reply with the JSON object only: {"level":"text"} or {"level":"package"}.'
    expect(parseEntryLevel("text")).toEqual({ issues: [`${shape} No JSON object was found in the reply.`] })
    for (const reply of [JSON.stringify({ level: "part" }), JSON.stringify({ parts: [] })])
      expect(parseEntryLevel(reply)).toEqual({ issues: [shape] })
  })

  test("one entry's level prompt holds that entry and the question, and no document", () => {
    const { directory } = prepared()
    const [delivery] = spineEntries(directory)
    const prompt = spineLevelPrompt(delivery!)
    expect(prompt).toStartWith("<entry>\n")
    expect(prompt.indexOf("<entry>")).toBeLessThan(prompt.indexOf("<question>"))
    expect(prompt).toContain(DELIVERY)
    expect(prompt).not.toContain(PAYMENT)
    expect(prompt).not.toContain("<documents>")
    // The entry reads as it does in its placement job.
    expect(spinePlacementPrompt(documents, delivery!)).toContain(prompt.split("\n\n<question>")[0]!)
  })

  test("an entry's limitations reach its level and placement jobs and the drafter's ground line", () => {
    const { directory } = prepared(LIMITED)
    const [delivery, payment] = spineEntries(directory)
    expect(payment!.limitations).toEqual([FLAG])
    expect(JSON.parse(tag(spineLevelPrompt(payment!), "entry"))).toEqual({ kind: "fact", detail: PAYMENT, limitations: [FLAG], sources: ["supply-terms.docx"], quote: "Payment is due on delivery." })
    expect(tag(spinePlacementPrompt(documents, payment!), "entry")).toBe(tag(spineLevelPrompt(payment!), "entry"))
    expect(groundLine(payment!)).toBe(`fact: ${PAYMENT} [source: supply-terms.docx] [limitations: ${FLAG}]`)
    expect(groundsBlock([{ ...payment!, level: "text" }])).toContain(`- fact: ${PAYMENT} [source: supply-terms.docx] [limitations: ${FLAG}] (evidence: "Payment is due on delivery."; retrieve the full words before quoting them)`)
    // An entry without limitations reads as before.
    expect(spineLevelPrompt(delivery!)).not.toContain("limitations")
    expect(groundLine(delivery!)).toBe(`requirement: ${DELIVERY} [source: supply-terms.docx]`)
  })

  test("the question asks by audience: a statement to the reader is text, a direction to the drafter package, with no tie-break to text", () => {
    const { directory } = prepared(AUDIENCE)
    for (const entry of spineEntries(directory)) {
      const question = spineLevelPrompt(entry).split("<question>\n")[1]!
      expect(question).toContain('- "text": it is what the document says to its reader.')
      expect(question).toContain('- "package": it is what the drafter does or works under in producing the documents. A direction to the drafter is "package", even where it shapes what the text says.')
      expect(question).not.toContain("is both")
    }
    expect(parseEntryLevel(byAudience(spineLevelPrompt(spineEntries(directory)[0]!)))).toEqual({ issues: [], value: "package" })
    expect(parseEntryLevel(byAudience(spineLevelPrompt(spineEntries(directory)[1]!)))).toEqual({ issues: [], value: "text" })
  })

  test("a direction to the drafter at the package level, placed in no part, raises no outline defect and is not disclosed; the statement to the reader is placed", async () => {
    const { directory } = prepared(AUDIENCE)
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.level]: byAudience, [AGENT.placement]: (text: string) => !reverseJob(text) && text.includes(DIRECTION) ? JSON.stringify({ parts: [] }) : placer(text) })
    expect((await run(directory, model)).status).toBe("done")
    expect(revisions(model)).toEqual([])
    const record = readLeanPlacement(directory)!
    expect(record).toMatchObject({ status: "placed", issues: [] })
    expect(record.entries).toEqual([{ entry: 0, level: "package", parts: [] }, { entry: 1, level: "text", parts: ["d1/1"] }])
    expect(planPlacementIssues(spineEntries(directory), record.entries)).toEqual([])
    expect(placementItems(directory)).toEqual([])
    // At the text level the same direction, placed nowhere, would go back to the outline.
    expect(planPlacementIssues(spineEntries(directory), [{ ...record.entries[0]!, level: "text" }])).toEqual([
      `No part covers requirement: ${DIRECTION}. Plan the part, or extend the scope of the part that should.`,
    ])
  })
})

describe("the outline's part scopes", () => {
  test("a part's scope says what the part tells the reader, and a requested marking goes in its document's Legend: line", () => {
    const prompt = stagePrompt(prepared().directory, "d1.draft", { outline: true })
    expect(prompt).toContain('A part\'s scope line says what the part tells the reader; a marking the lawyer asked for goes in its document\'s "Legend:" line, never in a part. ' +
      "Put internal source mappings in scope notes, not in visible headings or operative wording.")
    expect(prompt).not.toContain("planning explanations")
  })
})

describe("spine placement parser", () => {
  test("a reply must be {\"parts\": [...]}; known refs are kept once, in plan order", () => {
    expect(parseEntryPlacement("Delivery", documents).issues[0]).toContain('{"parts":["d1/3"]}')
    expect(parseEntryPlacement(JSON.stringify({ parts: "d1/1" }), documents).value).toBeUndefined()
    expect(parseEntryPlacement(JSON.stringify({ parts: ["d1/2", "d1/1", "d1/2"] }), documents)).toEqual({ issues: [], value: ["d1/1", "d1/2"] })
    expect(parseEntryPlacement(JSON.stringify({ parts: [] }), documents)).toEqual({ issues: [], value: [] })
  })

  test("an unknown part is retried with an issue naming its ref, and dropped on the last attempt", () => {
    const retried = parseEntryPlacement(JSON.stringify({ parts: ["d1/1", "d1/3", "d2/1"] }), documents)
    expect(retried.value).toBeUndefined()
    expect(retried.issues).toEqual([
      '"d1/3" is not a planned part. Name parts by their refs in <documents> ("d1/1", "d1/2").',
      '"d2/1" is not a planned part. Name parts by their refs in <documents> ("d1/1", "d1/2").',
    ])
    expect(parseEntryPlacement(JSON.stringify({ parts: ["d1/1", "d1/3"] }), documents, { final: true }).value).toEqual(["d1/1"])
  })

  test("a text entry of any kind placed in no part is a gap for the outline, a fact too; a package entry or an unchecked entry is not", () => {
    const entries = [
      { index: 0, kind: "requirement", detail: DELIVERY, evidence: [{ source_id: "supply-terms.docx", quote: "q" }], limitations: [] },
      { index: 1, kind: "fact", detail: PAYMENT, evidence: [], limitations: [] },
      { index: 2, kind: "decision", detail: "Deliver to the buyer's yard.", evidence: [], limitations: [] },
      { index: 3, kind: "requirement", detail: "Return the note as one file.", evidence: [], limitations: [] },
    ]
    expect(planPlacementIssues(entries, [
      { entry: 0, level: "text", parts: [] }, { entry: 1, level: "text", parts: [] }, { entry: 2, level: "text", parts: [], unchecked: true }, { entry: 3, level: "package", parts: [] },
    ])).toEqual([
      `No part covers requirement: ${DELIVERY}. Plan the part, or extend the scope of the part that should.`,
      `No part covers fact: ${PAYMENT}. Plan the part, or extend the scope of the part that should.`,
    ])
    expect(planPlacementIssues(entries, [
      { entry: 0, level: "text", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/2"] }, { entry: 2, level: "text", parts: ["d1/1"] }, { entry: 3, level: "package", parts: [] },
    ])).toEqual([])
  })

  test("a placed package entry grounds its parts; one placed nowhere raises nothing, and a text one placed nowhere is a gap", async () => {
    const third = { kind: "decision", detail: "Deliver to the buyer's yard", evidence: [], limitations: [] }
    const { directory } = prepared(JSON.stringify({ ...JSON.parse(SPINE), entries: [...JSON.parse(SPINE).entries, third] }))
    await run(directory, scriptedModel({ [AGENT.draft]: outliner(TWO) }))
    const plan = readLeanPlan(directory)!.documents
    const rows: EntryPlacement[] = [{ entry: 0, level: "package", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/2"] }, { entry: 2, level: "package", parts: [] }]
    expect(planPlacementIssues(spineEntries(directory), rows)).toEqual([])
    writePlacement(directory, plan, { entries: rows, parts: [], issues: [] })
    expect(documentGrounds(directory, "d1").map(grounds => grounds.map(entry => entry.index))).toEqual([[0], [1]])
    expect(placementItems(directory)).toEqual([])
    // The same decision at the text level, and the fact, placed nowhere, go back to the outline and are disclosed.
    const gap = rows.map(row => row.entry === 2 ? { ...row, level: "text" as const } : row.entry === 1 ? { ...row, parts: [] } : row)
    const issues = planPlacementIssues(spineEntries(directory), gap)
    expect(issues).toEqual([
      `No part covers fact: ${PAYMENT}. Plan the part, or extend the scope of the part that should.`,
      "No part covers decision: Deliver to the buyer's yard. Plan the part, or extend the scope of the part that should.",
    ])
    writePlacement(directory, plan, { entries: gap, parts: [], issues })
    expect(placementItems(directory).map(item => item.detail)).toEqual([
      `Document plan: no planned part sets out or reasons from this fact: ${PAYMENT} [source: supply-terms.docx]. The lawyer should confirm the drafts deal with it.`,
      "Document plan: no planned part sets out or reasons from this decision: Deliver to the buyer's yard. The lawyer should confirm the drafts deal with it.",
    ])
  })

  test("one entry's prompt: the shared documents first, then that entry, then the question", () => {
    const { directory } = prepared()
    const [delivery, payment] = spineEntries(directory)
    const first = spinePlacementPrompt(documents, delivery!)
    const second = spinePlacementPrompt(documents, payment!)
    expect(first).toStartWith("<documents>\n")
    expect(first.indexOf("<documents>")).toBeLessThan(first.indexOf("<entry>"))
    expect(first.indexOf("<entry>")).toBeLessThan(first.indexOf("<question>"))
    expect(first.split("<entry>")[0]).toBe(second.split("<entry>")[0]!)
    expect(first).toContain('"ref": "d1/2"')
    expect(first).toContain(DELIVERY)
    expect(first).not.toContain(PAYMENT)
  })

  test("the placement jobs read each document's form, and the record's key moves with it", () => {
    const { directory } = prepared()
    const [delivery] = spineEntries(directory)
    const formed: PlannedDocument[] = [{ ...documents[0]!, form: "agreement" }]
    expect(JSON.parse(tag(spinePlacementPrompt(formed, delivery!), "documents"))[0]).toMatchObject({ id: "d1", purpose: "Records the supply terms.", form: "agreement" })
    expect(spinePlacementPrompt(documents, delivery!)).not.toContain('"form"')
    expect(placementPlanHash(directory, formed)).not.toBe(placementPlanHash(directory, documents))
  })

  test("the record's key moves with the atoms' bytes and the documents' placement view, not with spine.json", () => {
    const { directory } = prepared()
    const key = placementPlanHash(directory, documents)
    expect(placementPlanHash(directory, [{ ...documents[0]!, file: "note.docx" }])).toBe(key)
    expect(placementPlanHash(directory, [{ ...documents[0]!, parts: documents[0]!.parts.slice(1) }])).not.toBe(key)
    writeLeanText(directory, "spine.json", SPINE.replace("thirty", "forty"))
    expect(placementPlanHash(directory, documents)).toBe(key)
    writeLeanText(directory, SPINE_ATOMS, SPINE.replace("thirty", "forty"))
    expect(placementPlanHash(directory, documents)).not.toBe(key)
    // A record of the earlier schema, without the parts' reverse jobs, is read as no record.
    writePlacement(directory, documents, { entries: [{ entry: 0, level: "text", parts: ["d1/1"] }], parts: [], issues: [] })
    expect(readLeanPlacement(directory)).toBeDefined()
    const { parts: _parts, ...earlier } = JSON.parse(readLeanText(directory, "placement.json")!)
    writeLeanText(directory, "placement.json", JSON.stringify({ ...earlier, schema: "dochaus.lean.placement/3" }))
    expect(readLeanPlacement(directory)).toBeUndefined()
  })

  test("the level and placement jobs read the atoms, never the entries as the spine wrote them", () => {
    const { directory } = prepared()
    writeLeanText(directory, "spine.json", JSON.stringify({ entries: [{ kind: "requirement", detail: `${DELIVERY} ${PAYMENT}`, evidence: [], limitations: ["Both terms."] }], limitations: [] }))
    expect(spineEntries(directory).map(entry => [entry.index, entry.detail])).toEqual([[0, DELIVERY], [1, PAYMENT]])
    writeLeanText(directory, SPINE_ATOMS, "")
    expect(spineEntries(directory)).toEqual([])
  })
})

describe("reverse placement, one job per planned part", () => {
  test("one part's prompt: every entry, without quotes, first, then its document, the part and the question", () => {
    const { directory } = prepared(LIMITED)
    const entries = spineEntries(directory)
    const formed: PlannedDocument = { ...documents[0]!, form: "agreement" }
    const [first, second] = [0, 1].map(index => partEntriesPrompt(formed, index, entries)) as [string, string]
    expect(first).toStartWith("<entries>\n")
    expect(first.indexOf("<entries>")).toBeLessThan(first.indexOf("<document>"))
    expect(first.indexOf("<document>")).toBeLessThan(first.indexOf("<part>"))
    expect(first.indexOf("<part>")).toBeLessThan(first.indexOf("<question>"))
    // The entries are shared, so the prefix caches across the parts' jobs.
    expect(first.split("<document>")[0]).toBe(second.split("<document>")[0]!)
    expect(tag(first, "entries").split("\n").map(line => JSON.parse(line))).toEqual([
      { index: 0, kind: "requirement", detail: DELIVERY }, { index: 1, kind: "fact", detail: PAYMENT, limitations: [FLAG] },
    ])
    // No quote or source: the payment entry's quote is not its detail.
    for (const word of ['"quote"', "Payment is due on delivery.", "supply-terms.docx"]) expect(first).not.toContain(word)
    expect(JSON.parse(tag(first, "document"))).toEqual({ id: "d1", title: "Supply note", purpose: "Records the supply terms.", form: "agreement" })
    expect(JSON.parse(tag(first, "part"))).toEqual({ ref: "d1/1", heading: "Delivery", scope: "When the goods are delivered." })
    expect(JSON.parse(tag(second, "part"))).toEqual({ ref: "d1/2", heading: "Payment", scope: "When payment falls due." })
    expect(first).not.toContain("When payment falls due.")
    expect(first).toContain("Which entries must this part set out or reason from?")
  })

  test("a reply must be {\"entries\": [...]}; known indexes are kept once, in spine order, and an empty list stands", () => {
    const entries = spineEntries(prepared().directory)
    const shape = 'Reply with the JSON object only: {"entries":[3,7]}, each an index from <entries>, or {"entries":[]} when the part rests on none.'
    expect(parsePartEntries("Entry 1", entries)).toEqual({ issues: [`${shape} No JSON object was found in the reply.`] })
    for (const reply of [{ entries: 1 }, { entries: ["1"] }, { entries: [0.5] }, { parts: ["d1/1"] }])
      expect(parsePartEntries(JSON.stringify(reply), entries)).toEqual({ issues: [shape] })
    expect(parsePartEntries(JSON.stringify({ entries: [1, 0, 1] }), entries)).toEqual({ issues: [], value: [0, 1] })
    expect(parsePartEntries(JSON.stringify({ entries: [] }), entries)).toEqual({ issues: [], value: [] })
  })

  test("an unknown index is retry feedback naming it, and dropped on the last attempt", () => {
    const entries = spineEntries(prepared().directory)
    const unknown = (index: number) => `${index} is not an entry index. Name entries by the "index" each has in <entries>.`
    expect(parsePartEntries(JSON.stringify({ entries: [1, 2, 7] }), entries)).toEqual({ issues: [unknown(2), unknown(7)] })
    expect(parsePartEntries(JSON.stringify({ entries: [1, 7] }), entries, { final: true })).toEqual({ issues: [unknown(7)], value: [1] })
  })

  test("a reverse hit adds its part to the entry's placement, in plan order; a reverse miss never removes one", () => {
    const two: PlannedDocument[] = [...documents, { id: "d2", title: "Cover note", purpose: "Flags the open points.", parts: [{ heading: "Open points", scope: "What the lawyer must confirm." }] }]
    const rows: EntryPlacement[] = [{ entry: 0, level: "text", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/2"] }, { entry: 2, level: "package", parts: [] }]
    expect(withPartEntries(two, rows, [{ part: "d1/1", entries: [] }, { part: "d1/2", entries: [1] }, { part: "d2/1", entries: [1, 2] }])).toEqual([
      { entry: 0, level: "text", parts: ["d1/1"] },
      { entry: 1, level: "text", parts: ["d1/2", "d2/1"], added: ["d2/1"] },
      { entry: 2, level: "package", parts: ["d2/1"], added: ["d2/1"] },
    ])
    expect(withPartEntries(two, [{ entry: 1, level: "text", parts: ["d2/1"] }], [{ part: "d1/1", entries: [1] }])).toEqual([{ entry: 1, level: "text", parts: ["d1/1", "d2/1"], added: ["d1/1"] }])
    expect(withPartEntries(two, rows, [{ part: "d1/1", entries: [], unchecked: true }, { part: "d1/2", entries: [] }])).toEqual(rows)
  })

  test("a part's reverse hit gives its entry a second home, whose part prompt carries it with its limitations", async () => {
    const { directory } = prepared(LIMITED)
    // The delivery part also reasons from the payment entry.
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placerWith(part => part.heading === "Delivery" ? [1] : []) })
    expect((await run(directory, model)).status).toBe("done")
    const record = readLeanPlacement(directory)!
    expect(record.parts).toEqual([{ part: "d1/1", entries: [1] }, { part: "d1/2", entries: [] }])
    expect(record.entries).toEqual([{ entry: 0, level: "text", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/1", "d1/2"], added: ["d1/1"] }])
    expect(documentGrounds(directory, "d1").map(grounds => grounds.map(entry => entry.index))).toEqual([[0, 1], [1]])
    const [delivery] = partPrompts(model).map(prompt => prompt.text)
    expect(delivery!.split("<rests_on>\n")[1]).toContain(`- fact: ${PAYMENT} [source: supply-terms.docx] [limitations: ${FLAG}]`)
  })

  test("a reverse hit fills the gap the entry's own job left, so the outline is not revised", async () => {
    const { directory } = prepared()
    // The plan has no delivery part, but its payment part reasons from the delivery requirement.
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.placement]: placerWith(() => [0]) })
    expect((await run(directory, model)).status).toBe("done")
    expect(revisions(model)).toEqual([])
    const record = readLeanPlacement(directory)!
    expect(record).toMatchObject({ status: "placed", issues: [] })
    expect(record.entries).toEqual([{ entry: 0, level: "text", parts: ["d1/1"], added: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/1"] }])
  })

  test("an unknown index is retried in the part's session; a reply that never passes adds nothing and is not saved", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: (text, _index, sessionID) => {
      if (!reverseSessions(model).has(sessionID!)) return placer(text)
      const part = JSON.parse(tag(model.prompts.find(prompt => prompt.sessionID === sessionID)!.text, "part")) as { ref: string }
      if (part.ref === "d1/2") return "The payment entry."
      return JSON.stringify({ entries: text.startsWith("Your reply") ? [1] : [9] })
    } })
    expect((await run(directory, model)).status).toBe("done")
    const asked = (ref: string) => reversePrompts(model).filter(prompt => prompt.sessionID === reversePrompts(model).find(first => first.text.includes(`"ref": "${ref}"`))!.sessionID)
    expect(asked("d1/1")).toHaveLength(2)
    expect(asked("d1/1")[1]!.text).toContain('9 is not an entry index. Name entries by the "index" each has in <entries>.')
    expect(asked("d1/2")).toHaveLength(1 + LEAN_MAX_RETRIES)
    const record = readLeanPlacement(directory)!
    expect(record.parts).toEqual([{ part: "d1/1", entries: [1] }, { part: "d1/2", entries: [], unchecked: true }])
    expect(record.entries).toEqual([{ entry: 0, level: "text", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/1", "d1/2"], added: ["d1/1"] }])
    expect(Object.values(readPartEntries(directory))).toEqual([[1]])
  })

  test("an unchanged part is not asked again, in the revised plan or after the lawyer edits another part", async () => {
    const { directory } = prepared()
    const revised = `${ONE}\n## Delivery\nWhen the goods are delivered.`
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE, revised), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    // The delivery requirement had no part, so the outline was revised once; its payment part was asked only for the first plan.
    expect(revisions(model)).toHaveLength(1)
    expect(reversePrompts(model).map(prompt => JSON.parse(tag(prompt.text, "part")).heading)).toEqual(["Payment", "Delivery"])
    expect(readLeanPlacement(directory)!.parts).toEqual([{ part: "d1/1", entries: [] }, { part: "d1/2", entries: [] }])
    // The lawyer edits the delivery part's scope: only that part's reverse job runs again, with every entry's own job.
    const plan = readLeanPlan(directory)!
    writeLeanPlan(directory, plan.documents.map(({ id: _id, ...document }) => ({ ...document, parts: document.parts.map(part => part.heading === "Delivery" ? { ...part, scope: "When and where the goods are delivered." } : part) })),
      plan.basis, plan.acting_for, undefined, plan.posture, plan.dates)
    expect(planPlacementCurrent(directory)).toBe(false)
    const placed = placementPrompts(model).length
    expect((await run(directory, model)).status).toBe("done")
    expect(reversePrompts(model).slice(2).map(prompt => JSON.parse(tag(prompt.text, "part")).scope)).toEqual(["When and where the goods are delivered."])
    expect(placementPrompts(model)).toHaveLength(placed + 2)
    expect(Object.keys(readPartEntries(directory))).toHaveLength(3)
    expect(planPlacementCurrent(directory)).toBe(true)
  })
})

describe("spine placement at the outline step", () => {
  test("each entry is placed by a job in a session of its own and the record names its parts", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    for (const prompts of [levelPrompts(model), placementPrompts(model)]) {
      expect(prompts.map(prompt => prompt.sessionID).toSorted()).toEqual(
        model.sessions.flatMap((session, at) => session.agent === prompts[0]!.agent && !reverseSessions(model).has(`ses_child${at + 1}`) ? [`ses_child${at + 1}`] : []).toSorted())
      expect(new Set(prompts.map(prompt => prompt.sessionID)).size).toBe(2)
      expect(prompts.filter(prompt => prompt.text.includes(DELIVERY) && prompt.text.includes(PAYMENT))).toEqual([])
    }
    const record = readLeanPlacement(directory)!
    expect(record.status).toBe("placed")
    expect(record.entries).toEqual([{ entry: 0, level: "text", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/2"] }])
    // Each planned part's reverse job, in a session of its own, is recorded too.
    expect(record.parts).toEqual([{ part: "d1/1", entries: [] }, { part: "d1/2", entries: [] }])
    expect(reversePrompts(model)).toHaveLength(2)
    expect(new Set(reversePrompts(model).map(prompt => prompt.sessionID)).size).toBe(2)
    expect(planPlacementCurrent(directory)).toBe(true)
    expect(readLeanEvents(directory).filter(event => event.event === "placement").map(event => event.outcome)).toEqual(["placed"])
  })

  test("a requirement placed in no part goes back to the outline, whose revised plan is covered and placed again", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE, TWO), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    expect(revisions(model)).toHaveLength(1)
    expect(revisions(model)[0]!.text).toContain(`- No part covers requirement: ${DELIVERY}. Plan the part, or extend the scope of the part that should.\n`)
    expect(revisions(model)[0]!.text).not.toContain("widen")
    expect(placementPrompts(model)).toHaveLength(4)
    // Levels read the entry alone, so the revised plan asks none again.
    expect(levelPrompts(model)).toHaveLength(2)
    expect(readLeanPlan(directory)!.documents[0]!.parts.map(part => part.heading)).toEqual(["Delivery", "Payment"])
    expect(readLeanPlacement(directory)!.status).toBe("placed")
    expect(disclosableItems(directory).filter(item => item.detail.includes(DELIVERY))).toEqual([])
  })

  test("a requirement still unplaced after the one revision is recorded unplaced and disclosed, and the step drafts", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    expect(PLAN_PLACEMENT_REVISIONS).toBe(1)
    expect(revisions(model)).toHaveLength(1)
    expect(placementPrompts(model)).toHaveLength(4)
    const record = readLeanPlacement(directory)!
    expect(record.status).toBe("unplaced")
    expect(record.issues).toHaveLength(1)
    expect(existsSync(leanFile(directory, "d1/draft.md"))).toBe(true)
    expect(disclosableItems(directory).filter(item => item.detail.includes(DELIVERY)).map(item => item.detail)).toEqual([
      `Document plan: no planned part sets out or reasons from this requirement: ${DELIVERY} [source: supply-terms.docx]. The lawyer should confirm the drafts deal with it.`,
    ])
  })

  test("an entry whose reply never passes is recorded unchecked and raises no revision", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: text => text.includes("<entry>") && text.includes(DELIVERY) ? "Delivery" : text.startsWith("Your reply") ? "Delivery" : placer(text) })
    expect((await run(directory, model)).status).toBe("done")
    expect(revisions(model)).toEqual([])
    expect(readLeanPlacement(directory)!.entries.find(row => row.entry === 0)).toEqual({ entry: 0, level: "text", parts: [], unchecked: true })
    expect(readLeanPlacement(directory)!.status).toBe("placed")
  })

  test("a stale record is placed again before the first document is drafted, asking only a changed entry's level", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placer })
    await run(directory, model)
    expect(placementPrompts(model)).toHaveLength(2)
    expect(levelPrompts(model)).toHaveLength(2)
    writeLeanText(directory, SPINE_ATOMS, SPINE.replace("thirty days", "forty days"))
    expect(planPlacementCurrent(directory)).toBe(false)
    await run(directory, model)
    expect(placementPrompts(model)).toHaveLength(4)
    expect(levelPrompts(model).slice(2).map(prompt => prompt.text.includes("forty days"))).toEqual([true])
    expect(planPlacementCurrent(directory)).toBe(true)
    expect(documentGrounds(directory, "d1").map(grounds => grounds.map(entry => entry.index))).toEqual([[0], [1]])
  })

  test("a later document's parts read the record and place nothing", async () => {
    const { directory } = prepared(JSON.stringify({ entries: [{ kind: "requirement", detail: PAYMENT, evidence: [], limitations: ["No source."] }], limitations: [] }))
    const outline = `${TWO}\n# Payment receipt\nForm: other\n## Payment\nThe payment received.`
    const model = scriptedModel({ [AGENT.draft]: outliner(outline), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    expect(readLeanPlacement(directory)!.entries).toEqual([{ entry: 0, level: "text", parts: ["d1/2", "d2/1"] }])
    const placed = placementPrompts(model).length
    expect((await run(directory, model, "d2")).status).toBe("done")
    expect(placementPrompts(model)).toHaveLength(placed)
    const d2 = partPrompts(model).at(-1)!.text
    expect(d2).toContain(`Draft part 1 of 1 of "Payment receipt"`)
    expect(d2).toContain(`<rests_on>\nThis part rests on:\n- requirement: ${PAYMENT} [limitations: No source.]\n`)
  })

  test("a package entry is placed like any other; placed in no part, it raises no revision and is not disclosed", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.level]: packageDelivery, [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    expect(levelPrompts(model)).toHaveLength(2)
    expect(placementPrompts(model).map(prompt => prompt.text.includes(DELIVERY)).toSorted()).toEqual([false, true])
    expect(revisions(model)).toEqual([])
    const record = readLeanPlacement(directory)!
    expect(record).toMatchObject({ status: "placed", issues: [] })
    expect(record.entries).toEqual([{ entry: 0, level: "package", parts: [] }, { entry: 1, level: "text", parts: ["d1/1"] }])
    expect(placementItems(directory)).toEqual([])
    expect(disclosableItems(directory).filter(item => item.detail.includes(DELIVERY))).toEqual([])
  })

  test("a level reply is retried in its session; one that never passes places the entry as text", async () => {
    const { directory } = prepared()
    // Delivery's level job answers in prose once, then passes; payment's never passes.
    const level = (text: string, _index: number, sessionID?: string) => {
      const first = levelPrompts(model).find(prompt => prompt.sessionID === sessionID)!.text
      if (first.includes(PAYMENT)) return "text"
      return text.startsWith("Your reply") ? JSON.stringify({ level: "text" }) : "It binds the text."
    }
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.level]: level, [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    const asked = (session: string) => levelPrompts(model).filter(prompt => prompt.sessionID === session).length
    expect([...new Set(levelPrompts(model).map(prompt => prompt.sessionID))].map(asked).toSorted()).toEqual([2, 1 + LEAN_MAX_RETRIES])
    expect(readLeanPlacement(directory)!.entries).toEqual([{ entry: 0, level: "text", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/2"], level_unchecked: true }])
    // Only the settled level is kept for a resume.
    expect(Object.values(readSpineLevels(directory))).toEqual(["text"])
  })

  test("levels settled before a pause are kept, so the resumed step asks none again", async () => {
    const { directory } = prepared()
    const stop = new Error("pause requested")
    const paused = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.level]: packageDelivery, [AGENT.placement]: () => { throw stop } })
    await expect(run(directory, paused)).rejects.toBe(stop)
    expect(readLeanPlan(directory)).toBeUndefined()
    const [delivery, payment] = spineEntries(directory)
    expect(readSpineLevels(directory)).toEqual({ [spineLevelKey(delivery!)]: "package", [spineLevelKey(payment!)]: "text" })
    const resumed = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placer })
    expect((await run(directory, resumed)).status).toBe("done")
    expect(levelPrompts(resumed)).toEqual([])
    expect(readLeanPlacement(directory)!.entries).toEqual([{ entry: 0, level: "package", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/2"] }])
  })
})

describe("spine placement in the part prompts", () => {
  test("a part prompt carries its grounds in their own tag, never in its labels, the outline list or the draft", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placer })
    await run(directory, model)
    const [delivery, payment] = partPrompts(model).map(prompt => prompt.text)
    const block = delivery!.split("<rests_on>\n")[1]!.split("\n</rests_on>")[0]!
    expect(block).toStartWith(`This part rests on:\n- requirement: ${DELIVERY} [source: supply-terms.docx] (evidence: "within thirty days of the order date"; retrieve the full words before quoting them)`)
    expect(block).toContain("State each here, in this part's own words, and reason from it. A cross-reference to another part adds to that; it never substitutes for it. An entry about method, calculation or handling binds what this part says, not words it repeats.")
    // A text entry is stated; nothing in the block reads as a direction to apply.
    expect(block).not.toContain("This part is drafted under:")
    expect(block).not.toContain(PAYMENT)
    expect(payment).toContain(`- fact: ${PAYMENT}`)
    // Outside the tag, the entry appears nowhere: not in the heading line or the outline list.
    expect(delivery!.replace(/<rests_on>[\s\S]*<\/rests_on>/, "")).not.toContain(DELIVERY)
    expect(readLeanPlan(directory)!.documents[0]!.parts.map(part => [part.heading, part.scope])).toEqual([["Delivery", "When the goods are delivered."], ["Payment", "When payment falls due."]])
    const draft = readLeanText(directory, "d1/draft.md")!
    expect(draft).not.toContain(DELIVERY)
    expect(draft).not.toContain("rests on")
  })

  test("a package entry its placement job places reaches that part's prompt and outline grounds; placed nowhere, it reaches none", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.level]: packageDelivery, [AGENT.placement]: placer })
    await run(directory, model)
    expect(readLeanPlacement(directory)!.entries).toEqual([{ entry: 0, level: "package", parts: ["d1/1"] }, { entry: 1, level: "text", parts: ["d1/2"] }])
    // A package entry is a direction the part applies, not words it states.
    const [delivery, payment] = partPrompts(model).map(prompt => prompt.text)
    const block = delivery!.split("<rests_on>\n")[1]!.split("\n</rests_on>")[0]!
    expect(block).toStartWith(`This part is drafted under:\n- requirement: ${DELIVERY}`)
    expect(block).toContain("Apply each as it directs how this part is drafted or where a point is disclosed: state it only where this part is the place it names; otherwise its words are not document text.")
    expect(block).not.toContain("State each here")
    expect(payment).not.toContain(DELIVERY)
    expect(documentGrounds(directory, "d1").map(grounds => grounds.map(entry => [entry.index, entry.level]))).toEqual([[[0, "package"]], [[1, "text"]]])
    const review = stagePrompt(directory, "d1.review")
    expect(review.split("<drafting_spine>")[1]).toContain(DELIVERY)
    // The reviews' outline names it as the part prompt was told it.
    const outline = JSON.parse(review.split("<document_outline>\n")[1]!.split("\n</document_outline>")[0]!)
    expect(outline.parts.map((part: { rests_on?: string[]; drafted_under?: string[] }) => [part.rests_on, part.drafted_under])).toEqual([
      [undefined, [`requirement: ${DELIVERY} [source: supply-terms.docx]`]], [[`fact: ${PAYMENT} [source: supply-terms.docx]`], undefined],
    ])
    // Placed in no part, it still rides the drafting spine every stage reads, and grounds no part.
    writePlacement(directory, readLeanPlan(directory)!.documents, { entries: [{ entry: 0, level: "package", parts: [] }, { entry: 1, level: "text", parts: ["d1/2"] }], parts: [], issues: [] })
    expect(documentGrounds(directory, "d1").map(grounds => grounds.map(entry => entry.index))).toEqual([[], [1]])
    expect(stagePrompt(directory, "d1.review").split("<document_outline>\n")[1]!.split("\n</document_outline>")[0]).not.toContain(DELIVERY)
  })

  test("the reviews' outline names each part's grounds; with none placed a part prompt has no block", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placer })
    await run(directory, model)
    const outline = JSON.parse(stagePrompt(directory, "d1.review").split("<document_outline>\n")[1]!.split("\n</document_outline>")[0]!)
    expect(outline.parts.map((part: { rests_on?: string[] }) => part.rests_on)).toEqual([
      [`requirement: ${DELIVERY} [source: supply-terms.docx]`], [`fact: ${PAYMENT} [source: supply-terms.docx]`],
    ])
    expect(groundsBlock([])).toBe("")
    expect(sectionPrompt(readLeanPlan(directory)!.documents[0]!, 0)).not.toContain("<rests_on>")
  })

  test("a part's block states its text entries and applies its package entries, each group under its own line", () => {
    const { directory } = prepared()
    const [delivery, payment] = spineEntries(directory)
    const block = groundsBlock([{ ...delivery!, level: "package" }, { ...payment!, level: "text" }])
    expect(block).toStartWith("<rests_on>\nThis part rests on:\n")
    const [stated, applied] = block.split("This part is drafted under:\n")
    expect(stated).toContain(`- fact: ${PAYMENT} [source: supply-terms.docx]`)
    expect(stated).toContain("State each here, in this part's own words, and reason from it.")
    expect(stated).not.toContain(DELIVERY)
    expect(applied).toStartWith(`- requirement: ${DELIVERY} [source: supply-terms.docx]`)
    expect(applied).toContain("Apply each as it directs how this part is drafted or where a point is disclosed: state it only where this part is the place it names; otherwise its words are not document text.")
    expect(applied).not.toContain(PAYMENT)
  })
})
