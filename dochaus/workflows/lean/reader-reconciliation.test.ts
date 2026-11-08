import { afterEach, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import type { ParagraphRow } from "./checks"
import { parsePlan, writeLeanPlan } from "./plan"
import { applyReaderItems, checkReaderParts, readReaderChecks, readerCheckKey, readerItems, saveReaderCheck } from "./reader-check"
import { reconcilePackageReaderItems } from "./reader-reconciliation"
import { readLeanState, readOpenItems, writeLeanText, writeOpenItems } from "./state"
import { AGENT, cleanupMatters, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const pid = "p100000000001"
const prose = "Please reply at the agreed address."
const marker = "[open item: verify the reply address]"
const paragraph = `${prose} ${marker}`
const hit = (sentence: string) => ({ pid, sentence, addressed_to: "the lawyer", keep: "a reply is requested" })
const review = { status: "changes-required" as const, findings: [{ target: "memo" as const, detail: "The notes describe a completed recipient-wording edit as outstanding." }], limitations: [], delivery: [] }

function setup(sentence = paragraph, statement = sentence) {
  const directory = matter()
  const plan = parsePlan("Acting for: not stated\n# Notice\nForm: letter\n## Response\nThe reply requested.")
  writeLeanPlan(directory, plan.value!.documents.map(document => ({ ...document, furniture: { legend: "Working review copy" } })), undefined,
    { party: "The sender", basis: "stated" })
  const rows = [{ pid, text: sentence, sha256: paragraphSha256(sentence), draft_part: 0 }]
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(rows))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(rows))
  const [warning] = readerItems("apply", "d1", [hit(statement)])
  const [history] = readerItems("draft", "d1", [hit(statement)])
  const unknown = { id: "apply:placeholder:unknown", stage: "apply" as const, kind: "placeholder" as const, status: "open" as const,
    detail: "The reply address is unknown.", ref: { document: "d1" as const, pid, placeholder: marker } }
  writeOpenItems(directory, [warning!, history!, unknown])
  writeLeanText(directory, "drafting-notes.md", `# Notes\n\nRecipient wording remains to be rewritten. [ref: ${warning!.id}]\n`)
  return { directory, warning: warning!, history: history!, unknown, rows }
}

test("a checked disagreement retires only the current false audience warning and preserves the factual gap and draft history", async () => {
  const { directory, warning, history, unknown } = setup()
  const model = scriptedModel({ [AGENT.reader]: () => '{"hits":[]}' })
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: model.port }, review)).toBe(true)
  expect(readOpenItems(directory)).toEqual([history, unknown])
  expect(model.prompts[0]!.text).toContain(`<package_finding>\n${review.findings[0]!.detail}`)
  expect(model.prompts[0]!.text).toContain(warning.detail)
  expect(model.sessions[0]!.permissions).toEqual([{ action: "*", resource: "*", effect: "deny" }])
  // The actual reader output is content-keyed; no extra receipt is needed.
  writeOpenItems(directory, [warning, history, unknown])
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: model.port }, review)).toBe(true)
  expect(model.prompts).toHaveLength(1)
})

test("a checked retirement survives an ordinary cached apply, keeps other hits and history, and changed text is checked again", async () => {
  const { directory, warning, history, unknown, rows } = setup(paragraph, prose)
  const otherPid = "p100000000002"
  const repeatedPid = "p100000000003"
  const genuine = { ...hit("Lawyer: confirm the dispatch date before sending."), pid: otherPid }
  const distinct = hit("at the agreed address.")
  const repeated = { ...hit(prose), pid: repeatedPid }
  const current: ParagraphRow[] = [...rows,
    { pid: otherPid, text: genuine.sentence, sha256: paragraphSha256(genuine.sentence), draft_part: 0 },
    { pid: repeatedPid, text: prose, sha256: paragraphSha256(prose), draft_part: 0 },
  ]
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(current))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(current))
  let prompt = ""
  const cached = [hit("Please  reply at the agreed address."), distinct, genuine, repeated]
  await checkReaderParts(directory, "d1", [{ part: 0, rows: current }], async (_label, text, check) => {
    prompt = text
    return check(JSON.stringify({ hits: cached }), false)
  })
  const key = readerCheckKey(prompt)
  const historicalKey = readerCheckKey(`${prompt}\nEarlier audience question.`)
  saveReaderCheck(directory, "d1", historicalKey, cached)
  const idle = scriptedModel({})
  const before = await applyReaderItems(directory, readLeanState(directory)!, "d1", current, new Set([pid]), { sessions: idle.port })
  expect(before.map(item => item.ref.statement)).toEqual(cached.map(row => row.sentence))
  expect(idle.prompts).toEqual([])

  const reader = scriptedModel({ [AGENT.reader]: () => '{"hits":[]}' })
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: reader.port }, review)).toBe(true)
  expect(readOpenItems(directory)).toEqual([history, unknown])
  const applied = await applyReaderItems(directory, readLeanState(directory)!, "d1", current, new Set([pid]), { sessions: idle.port })
  expect(applied.map(item => item.ref.statement)).toEqual([distinct, genuine, repeated].map(row => row.sentence))
  expect(applied.some(item => item.id === warning.id)).toBe(false)
  expect(idle.prompts).toEqual([])
  expect(readReaderChecks(directory, "d1")[key]).toEqual([distinct, genuine, repeated])
  expect(readReaderChecks(directory, "d1")[historicalKey]).toEqual(cached)

  const changed = "Lawyer: verify the delivery method before sending."
  const final = current.map(row => row.pid === pid ? { ...row, text: changed, sha256: paragraphSha256(changed) } : row)
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(final))
  const fresh = scriptedModel({ [AGENT.reader]: () => JSON.stringify({ hits: [hit(changed), genuine] }) })
  const checked = await applyReaderItems(directory, readLeanState(directory)!, "d1", final, new Set([pid]), { sessions: fresh.port })
  expect(fresh.prompts).toHaveLength(1)
  expect(fresh.prompts[0]!.text).toContain(changed)
  expect(checked.map(item => item.ref.statement)).toEqual([changed, genuine.sentence])
  expect(readReaderChecks(directory, "d1")[key]).toEqual([distinct, genuine, repeated])
  expect(readReaderChecks(directory, "d1")[historicalKey]).toEqual(cached)
})

test("retirement without an ordinary part cache does not certify its uninspected neighbor", async () => {
  const { directory, history, unknown, rows } = setup(paragraph, prose)
  const genuine = { ...hit("Lawyer: confirm the dispatch date before sending."), pid: "p100000000002" }
  const current = [...rows, { pid: genuine.pid, text: genuine.sentence, sha256: paragraphSha256(genuine.sentence), draft_part: 0 }]
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(current))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(current))
  const historicalKey = readerCheckKey("Earlier part text and audience question.")
  saveReaderCheck(directory, "d1", historicalKey, [hit(prose)])
  const reader = scriptedModel({ [AGENT.reader]: () => '{"hits":[]}' })
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: reader.port }, review)).toBe(true)
  expect(readOpenItems(directory)).toEqual([history, unknown])
  expect(Object.keys(readReaderChecks(directory, "d1")).toSorted())
    .toEqual([historicalKey, readerCheckKey(reader.prompts[0]!.text)].toSorted())
  const fresh = scriptedModel({ [AGENT.reader]: () => JSON.stringify({ hits: [genuine] }) })
  const checked = await applyReaderItems(directory, readLeanState(directory)!, "d1", current, new Set([pid]), { sessions: fresh.port })
  expect(fresh.prompts).toHaveLength(1)
  expect(checked.map(item => item.ref.statement)).toEqual([genuine.sentence])
  expect(readReaderChecks(directory, "d1")[historicalKey]).toEqual([hit(prose)])
})

test("a genuine wrong-audience sentence still in the current body is retained", async () => {
  const sentence = "Lawyer: check the reply address before sending this letter."
  const { directory, warning, history, unknown } = setup(sentence)
  const model = scriptedModel({ [AGENT.reader]: () => JSON.stringify({ hits: [hit(sentence)] }) })
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: model.port }, review)).toBe(false)
  expect(readOpenItems(directory)).toEqual([warning, history, unknown])
})

test("invalid or mixed-span replies cannot retire a warning even on the last retry", async () => {
  const { directory, warning, history, unknown } = setup()
  const model = scriptedModel({ [AGENT.reader]: () => JSON.stringify({ hits: [hit(paragraph)] }) })
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: model.port }, review)).toBe(false)
  expect(readOpenItems(directory)).toEqual([warning, history, unknown])
  expect(model.prompts).toHaveLength(4)
})

test("a differently quoted genuine defect does not create an uncharged replacement loop", async () => {
  const sentence = "Lawyer: verify the reply address. Check the date before sending."
  const { directory, warning, history, unknown } = setup(sentence)
  const model = scriptedModel({ [AGENT.reader]: () => JSON.stringify({ hits: [hit("Check the date before sending.")] }) })
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: model.port }, review)).toBe(false)
  expect(readOpenItems(directory)).toEqual([warning, history, unknown])
})

test("unrelated body corrections and old sentences do not trigger reader retirement", async () => {
  const { directory, warning, history, unknown } = setup()
  const model = scriptedModel({ [AGENT.reader]: () => '{"hits":[]}' })
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: model.port }, { ...review, findings: [{ target: "body", document: "d1", detail: "Change the operative period." }] })).toBe(false)
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid, text: prose, sha256: paragraphSha256(prose) }]))
  expect(await reconcilePackageReaderItems(directory, readLeanState(directory)!, { sessions: model.port }, review)).toBe(false)
  expect(model.prompts).toHaveLength(0)
  expect(readOpenItems(directory)).toEqual([warning, history, unknown])
})
