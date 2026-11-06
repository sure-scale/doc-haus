import { afterEach, describe, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import { LEAN_ACTS_SCHEMA } from "./act-record"
import { runDraftStage, sectionPrompt } from "./draft-sections"
import { readLeanEvents } from "./events"
import { readLeanPlan, type PlannedDocument } from "./plan"
import { stagePrompt } from "./prompts"
import {
  PLAN_PLACEMENT_REVISIONS, documentGrounds, groundsBlock, parseEntryPlacement, placementPlanHash, planPlacementCurrent, planPlacementIssues, readLeanPlacement, spineEntries, spinePlacementPrompt,
} from "./spine-placement"
import { disclosableItems } from "./stages"
import { leanFile, readLeanState, readLeanText, writeLeanText } from "./state"
import { AGENT, cleanupMatters, matter, scriptedModel } from "./test-support"

// Spine placement: one focused job per drafting-spine entry names the planned
// parts that rest on it; a requirement or decision placed in no part goes back
// to the outline, what is left is recorded unplaced and disclosed, and each
// part prompt is told the entries it rests on. Synthetic matter only.

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

const documents: PlannedDocument[] = [{ id: "d1", title: "Supply note", purpose: "Records the supply terms.", parts: [
  { heading: "Delivery", scope: "When the goods are delivered." }, { heading: "Payment", scope: "When payment falls due." },
] }]

/** The draft agent: `first` answers the outline, `revised` any revision request, each part one paragraph. */
const outliner = (first: string, revised = first) => (text: string) =>
  /^Draft part \d+ of \d+ of "/m.test(text) ? "The terms are recorded." : text.startsWith(REVISE) ? revised : first

/** The placement agent: each entry rests on the planned part of the heading its detail names, if planned. */
const placer = (text: string) => {
  const planned = JSON.parse(text.split("<documents>\n")[1]!.split("\n</documents>")[0]!.split("\n").slice(1).join("\n")) as { parts: { ref: string; heading: string }[] }[]
  const entry = JSON.parse(text.split("<entry>\n")[1]!.split("\n</entry>")[0]!.split("\n").slice(1).join("\n")) as { detail: string }
  const heading = entry.detail.startsWith("The supplier must deliver") ? "Delivery" : "Payment"
  return JSON.stringify({ parts: planned.flatMap(document => document.parts.filter(part => part.heading === heading).map(part => part.ref)) })
}

function prepared(spine = SPINE) {
  const directory = matter()
  writeLeanText(directory, "acts.json", JSON.stringify({ schema: LEAN_ACTS_SCHEMA, status: "none", acts: [], verify: null, deliverables: [], issues: [] }))
  writeLeanText(directory, "spine.json", spine)
  return { directory, state: readLeanState(directory)! }
}

const run = (directory: string, model: ReturnType<typeof scriptedModel>, document: "d1" | "d2" = "d1") =>
  runDraftStage(directory, readLeanState(directory)!, { sessions: model.port }, () => {}, document, () => {})
const placementPrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.placement)
const partPrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.draft && /^Draft part \d+ of /m.test(prompt.text))

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

  test("a requirement or decision placed in no part is an outline issue; a fact or an unchecked entry is not", () => {
    const entries = [
      { index: 0, kind: "requirement", detail: DELIVERY, evidence: [{ source_id: "supply-terms.docx", quote: "q" }] },
      { index: 1, kind: "fact", detail: PAYMENT, evidence: [] },
      { index: 2, kind: "decision", detail: "Deliver to the buyer's yard.", evidence: [] },
    ]
    expect(planPlacementIssues(entries, [{ entry: 0, parts: [] }, { entry: 1, parts: [] }, { entry: 2, parts: [], unchecked: true }])).toEqual([
      `No planned part sets out or reasons from the requirement "${DELIVERY}" [source: supply-terms.docx]. Plan the part whose scope states or applies it, or widen the scope of the part that should.`,
    ])
    expect(planPlacementIssues(entries, [{ entry: 0, parts: ["d1/1"] }, { entry: 1, parts: [] }, { entry: 2, parts: ["d1/1"] }])).toEqual([])
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

  test("the record's key moves with the spine's bytes and the documents' coverage view", () => {
    const { directory } = prepared()
    const key = placementPlanHash(directory, documents)
    expect(placementPlanHash(directory, [{ ...documents[0]!, file: "note.docx" }])).toBe(key)
    expect(placementPlanHash(directory, [{ ...documents[0]!, parts: documents[0]!.parts.slice(1) }])).not.toBe(key)
    writeLeanText(directory, "spine.json", SPINE.replace("thirty", "forty"))
    expect(placementPlanHash(directory, documents)).not.toBe(key)
  })
})

describe("spine placement at the outline step", () => {
  test("each entry is placed by a job in a session of its own and the record names its parts", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    expect(placementPrompts(model).map(prompt => prompt.sessionID).toSorted()).toEqual(
      model.sessions.flatMap((session, at) => session.agent === AGENT.placement ? [`ses_child${at + 1}`] : []).toSorted())
    expect(new Set(placementPrompts(model).map(prompt => prompt.sessionID)).size).toBe(2)
    expect(placementPrompts(model).filter(prompt => prompt.text.includes(DELIVERY) && prompt.text.includes(PAYMENT))).toEqual([])
    const record = readLeanPlacement(directory)!
    expect(record.status).toBe("placed")
    expect(record.entries.toSorted((a, b) => a.entry - b.entry)).toEqual([{ entry: 0, parts: ["d1/1"] }, { entry: 1, parts: ["d1/2"] }])
    expect(planPlacementCurrent(directory)).toBe(true)
    expect(readLeanEvents(directory).filter(event => event.event === "placement").map(event => event.outcome)).toEqual(["placed"])
  })

  test("a requirement placed in no part goes back to the outline, whose revised plan is covered and placed again", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE, TWO), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    const revisions = model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.startsWith(REVISE))
    expect(revisions).toHaveLength(1)
    expect(revisions[0]!.text).toContain(`No planned part sets out or reasons from the requirement "${DELIVERY}"`)
    expect(placementPrompts(model)).toHaveLength(4)
    expect(readLeanPlan(directory)!.documents[0]!.parts.map(part => part.heading)).toEqual(["Delivery", "Payment"])
    expect(readLeanPlacement(directory)!.status).toBe("placed")
    expect(disclosableItems(directory).filter(item => item.detail.includes(DELIVERY))).toEqual([])
  })

  test("a requirement still unplaced after the revisions is recorded unplaced and disclosed, and the step drafts", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.startsWith(REVISE))).toHaveLength(PLAN_PLACEMENT_REVISIONS)
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
    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.startsWith(REVISE))).toEqual([])
    expect(readLeanPlacement(directory)!.entries.find(row => row.entry === 0)).toEqual({ entry: 0, parts: [], unchecked: true })
    expect(readLeanPlacement(directory)!.status).toBe("placed")
  })

  test("a stale record is placed again before the first document is drafted", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placer })
    await run(directory, model)
    expect(placementPrompts(model)).toHaveLength(2)
    writeLeanText(directory, "spine.json", SPINE.replace("thirty days", "forty days"))
    expect(planPlacementCurrent(directory)).toBe(false)
    await run(directory, model)
    expect(placementPrompts(model)).toHaveLength(4)
    expect(planPlacementCurrent(directory)).toBe(true)
    expect(documentGrounds(directory, "d1").map(grounds => grounds.map(entry => entry.index))).toEqual([[0], [1]])
  })

  test("a later document's parts read the record and place nothing", async () => {
    const { directory } = prepared(JSON.stringify({ entries: [{ kind: "requirement", detail: PAYMENT, evidence: [], limitations: ["No source."] }], limitations: [] }))
    const outline = `${TWO}\n# Payment receipt\nForm: other\n## Payment\nThe payment received.`
    const model = scriptedModel({ [AGENT.draft]: outliner(outline), [AGENT.placement]: placer })
    expect((await run(directory, model)).status).toBe("done")
    expect(readLeanPlacement(directory)!.entries).toEqual([{ entry: 0, parts: ["d1/2", "d2/1"] }])
    const placed = placementPrompts(model).length
    expect((await run(directory, model, "d2")).status).toBe("done")
    expect(placementPrompts(model)).toHaveLength(placed)
    const d2 = partPrompts(model).at(-1)!.text
    expect(d2).toContain(`Draft part 1 of 1 of "Payment receipt"`)
    expect(d2).toContain(`<rests_on>\nThis part rests on:\n- requirement: ${PAYMENT}\n`)
  })
})

describe("spine placement in the part prompts", () => {
  test("a part prompt carries its grounds in their own tag, never in its labels, the outline list or the draft", async () => {
    const { directory } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.placement]: placer })
    await run(directory, model)
    const [delivery, payment] = partPrompts(model).map(prompt => prompt.text)
    const block = delivery!.split("<rests_on>\n")[1]!.split("\n</rests_on>")[0]!
    expect(block).toContain(`- requirement: ${DELIVERY} [source: supply-terms.docx] (evidence: "within thirty days of the order date"; retrieve the full words before quoting them)`)
    expect(block).toContain("A cross-reference to another part adds to that; it never substitutes for it.")
    expect(block).not.toContain(PAYMENT)
    expect(payment).toContain(`- fact: ${PAYMENT}`)
    // Outside the tag, the entry appears nowhere: not in the heading line or the outline list.
    expect(delivery!.replace(/<rests_on>[\s\S]*<\/rests_on>/, "")).not.toContain(DELIVERY)
    expect(readLeanPlan(directory)!.documents[0]!.parts.map(part => [part.heading, part.scope])).toEqual([["Delivery", "When the goods are delivered."], ["Payment", "When payment falls due."]])
    const draft = readLeanText(directory, "d1/draft.md")!
    expect(draft).not.toContain(DELIVERY)
    expect(draft).not.toContain("rests on")
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
})
