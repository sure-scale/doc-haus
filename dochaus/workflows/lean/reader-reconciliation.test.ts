import { afterEach, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { parsePlan, writeLeanPlan } from "./plan"
import { readerItems } from "./reader-check"
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

function setup(sentence = paragraph) {
  const directory = matter()
  const plan = parsePlan("Acting for: not stated\n# Notice\nForm: letter\n## Response\nThe reply requested.")
  writeLeanPlan(directory, plan.value!.documents)
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid, text: sentence, sha256: paragraphSha256(sentence) }]))
  const [warning] = readerItems("apply", "d1", [hit(sentence)])
  const [history] = readerItems("draft", "d1", [hit(sentence)])
  const unknown = { id: "apply:placeholder:unknown", stage: "apply" as const, kind: "placeholder" as const, status: "open" as const,
    detail: "The reply address is unknown.", ref: { document: "d1" as const, pid, placeholder: marker } }
  writeOpenItems(directory, [warning!, history!, unknown])
  writeLeanText(directory, "drafting-notes.md", `# Notes\n\nRecipient wording remains to be rewritten. [ref: ${warning!.id}]\n`)
  return { directory, warning: warning!, history: history!, unknown }
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
