import { afterEach, expect, test } from "bun:test"
import { buildChecklistReinspectionPrompt, buildChecklistCheckerPrompt, buildChecklistProducerPrompt, type ChecklistCheckerInput, type ChecklistProducerInput } from "./checklist-contract"
import { CHECKLIST_JOB_AGENTS, CHECKLIST_JOB_PERMISSIONS, runChecklistCheckJob, runChecklistExtractionJob, runChecklistRefinementJob } from "./checklist-job"
import { LEAN_PROMPT_VERSION, stageInputs, LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"

import { cleanupMatters, matter, SOURCE } from "./test-support"
import { readLeanState, LEAN_STATE_SCHEMA, hashStageInputs, writeLeanText, writeLeanBytes } from "./state"
import { docxodus } from "../../lib/docxodus"
import { buildChecklistDocumentContext } from "./checklist-document-context"
import { writeLeanPlan } from "./plan"
import { renderMarkdownToDocx } from "../../lib/render"
import { readDocxCommentParagraphs, exportDocxComments, readDocxRevisionText } from "../../lib/docx-comments"

afterEach(cleanupMatters)

function acceptApplyForTest(directory: string) {
  const state = readLeanState(directory)
  writeLeanText(directory, "state.json", JSON.stringify({ ...state, schema: LEAN_STATE_SCHEMA, profile: "lean", stages: { ...state?.stages,
    "d1.apply": { status: "done", input_sha: hashStageInputs("d1.apply", LEAN_PROMPT_VERSION, stageInputs(directory, "d1.apply")) },
  } }))
}

const source = { document: SOURCE, locator: "paragraph 1", quote: "The proposal remains open." }
const producer: ChecklistProducerInput = { source: { document: SOURCE, role: "client correspondence" }, passages: [source], instructions: [], targets: [{ document: "draft.docx", role: "agreement" }], retrieval: "Use exact cited source coordinates." }
const checker: ChecklistCheckerInput = {
  item_id: "source-1-item-1",
  item: { check: "Check the proposed arrangement is identified as open.", basis: "proposal", source_refs: [source], qualifications: ["Unconfirmed."], target: { document: "draft.docx", role: "agreement", treatment: "open_issue" } },
  current_evidence: [], settled_instructions: [], retrieval: "Retrieve the cited target paragraph when needed.",
}
const finding = { item_id: checker.item_id, status: "indeterminate", reason: "Current evidence needs retrieval.", evidence: [], source_refs: [source], inspection_scope: "Supplied source only.", limitation: "Target not inspected." }

function sessionPort(replies: (string | Error)[]) {
  const created: Parameters<StageSessionPort["createSession"]>[0][] = []
  const prompts: Parameters<StageSessionPort["prompt"]>[0][] = []
  const sessions: StageSessionPort = {
    async createSession(input) { created.push(input); return "session-1" },
    async prompt(input) {
      prompts.push(input)
      const reply = replies[prompts.length - 1]
      if (reply instanceof Error) throw reply
      if (reply === undefined) throw new Error("No scripted response")
      return reply
    },
  }
  return { sessions, created, prompts }
}

test("one typed extraction job uses native readonly boundary and canonical prompt", async () => {
  const directory = matter(source.quote)
  const port = sessionPort([JSON.stringify({ items: [checker.item] })])
  const result = await runChecklistExtractionJob(directory, { sessions: port.sessions, model: { providerID: "provider", modelID: "model" }, parentID: "owner" }, producer)
  expect(result.status).toBe("succeeded")
  expect(port.created).toHaveLength(1)
  expect(port.created[0]).toMatchObject({ directory, parentID: "owner", agent: CHECKLIST_JOB_AGENTS.extract, model: { providerID: "provider", modelID: "model" }, permissions: CHECKLIST_JOB_PERMISSIONS })
  expect(port.prompts[0].text).toBe(buildChecklistProducerPrompt(producer))
  expect(CHECKLIST_JOB_PERMISSIONS.filter(row => row.effect === "allow").map(row => row.action)).toEqual(["read-record", "read-document", "search-document", "verify-quote"])
  if (result.status === "succeeded") expect(result.value.items[0]).toEqual(checker.item)
})

test("malformed JSON and shape receive bounded correction in the same native session", async () => {
  const directory = matter(source.quote)
  const port = sessionPort(["```json\n{}\n```", "{}", JSON.stringify(finding)])
  const before = JSON.stringify(checker)
  const result = await runChecklistCheckJob(directory, { sessions: port.sessions, model: null }, checker)
  expect(result.status).toBe("succeeded")
  expect(result.attempts).toBe(3)
  expect(port.created).toHaveLength(1)
  expect(port.prompts.every(row => row.sessionID === "session-1" && row.agent === CHECKLIST_JOB_AGENTS.check)).toBe(true)
  expect(port.prompts[0].text).toBe(buildChecklistCheckerPrompt(checker))
  expect(port.prompts[1].text).toContain("without Markdown fences")
  expect(port.prompts[2].text).toContain("item_id")
  expect(JSON.stringify(checker)).toBe(before)
})

test("persistent malformed output stops at the existing retry bound without repair", async () => {
  const directory = matter(source.quote)
  const port = sessionPort(Array.from({ length: LEAN_MAX_RETRIES + 1 }, () => "not json"))
  const result = await runChecklistExtractionJob(directory, { sessions: port.sessions, model: null }, producer)
  expect(result.status).toBe("failed")
  expect(result.attempts).toBe(LEAN_MAX_RETRIES + 1)
  expect(result.reply).toBe("not json")
  expect(port.prompts).toHaveLength(LEAN_MAX_RETRIES + 1)
})

test("provider failure and invalid caller input do not trigger output correction", async () => {
  const directory = matter(source.quote)
  const port = sessionPort([new Error("provider unavailable")])
  await expect(runChecklistCheckJob(directory, { sessions: port.sessions, model: null }, checker)).rejects.toThrow("provider unavailable")
  expect(port.prompts).toHaveLength(1)
  const untouched = sessionPort([])
  expect(() => runChecklistExtractionJob(directory, { sessions: untouched.sessions, model: null }, { ...producer, passages: [] })).toThrow()
  expect(untouched.created).toHaveLength(0)
})

test("a changed source quote is rejected then corrected against canonical source bytes", async () => {
  const directory = matter(source.quote)
  const valid = { items: [checker.item] }
  const baseline = await runChecklistExtractionJob(directory, { sessions: sessionPort([JSON.stringify(valid)]).sessions, model: null }, producer)
  expect(baseline.status).toBe("succeeded")
  const changed = { items: [{ ...checker.item, source_refs: [{ ...source, quote: "The proposal is accepted." }] }] }
  // Even the input passage repeating the miss cannot establish its authority.
  const port = sessionPort([JSON.stringify(changed), JSON.stringify(valid)])
  const result = await runChecklistExtractionJob(directory, { sessions: port.sessions, model: null }, { ...producer, passages: changed.items[0].source_refs })
  expect(result.status).toBe("succeeded")
  expect(result.attempts).toBe(2)
  expect(port.prompts[1].text).toContain('quotes "The proposal is accepted."')
  expect(port.prompts[1].text).toContain("verbatim source words")
  if (result.status === "succeeded") expect(result.value).toEqual(valid)
})

test("checker identity mismatch receives correction without rewriting the input item", async () => {
  const directory = matter(source.quote)
  const baseline = await runChecklistCheckJob(directory, { sessions: sessionPort([JSON.stringify(finding)]).sessions, model: null }, checker)
  expect(baseline.status).toBe("succeeded")
  const port = sessionPort([JSON.stringify({ ...finding, item_id: "another-item" }), JSON.stringify(finding)])
  const before = JSON.stringify(checker)
  const result = await runChecklistCheckJob(directory, { sessions: port.sessions, model: null }, checker)
  expect(result.status).toBe("succeeded")
  expect(result.attempts).toBe(2)
  expect(port.prompts[1].text).toContain(`item_id must equal the supplied "${checker.item_id}"`)
  expect(JSON.stringify(checker)).toBe(before)
})

test("producer rejects unknown document identities while allowing descriptive role variation", async () => {
  const directory = matter(source.quote)
  const variedRole = { items: [{ ...checker.item, target: { ...checker.item.target, role: "document recording the parties' arrangement" } }] }
  const baseline = await runChecklistExtractionJob(directory, { sessions: sessionPort([JSON.stringify(variedRole)]).sessions, model: null }, producer)
  expect(baseline.status).toBe("succeeded")
  const port = sessionPort([
    JSON.stringify({ items: [{ ...checker.item, target: { ...variedRole.items[0].target, document: "invented.docx" } }] }),
    JSON.stringify(variedRole),
  ])
  const result = await runChecklistExtractionJob(directory, { sessions: port.sessions, model: null }, producer)
  expect(result.status).toBe("succeeded")
  expect(result.attempts).toBe(2)
  expect(port.prompts[1].text).toContain("declared document identity")
  if (result.status === "succeeded") expect(result.value).toEqual(variedRole)
})

test("request quotes use canonical request or supplied instructions, consumer answers stay separate", async () => {
  const directory = matter(source.quote)
  writeLeanText(directory, "request.md", "Keep unresolved proposals open.")
  const requestRef = { document: "request", locator: "request paragraph 1", quote: "Keep unresolved proposals open." }
  const requestItem = { ...checker.item, source_refs: [requestRef] }
  const canonical = await runChecklistExtractionJob(directory, { sessions: sessionPort([JSON.stringify({ items: [requestItem] })]).sessions, model: null }, producer)
  expect(canonical.status).toBe("succeeded")
  const supplied = { document: "request", locator: "request paragraph 2", quote: "Review only the supplied proposal." }
  const suppliedResult = await runChecklistExtractionJob(directory, { sessions: sessionPort([JSON.stringify({ items: [{ ...checker.item, source_refs: [supplied] }] })]).sessions, model: null }, { ...producer, instructions: [supplied] })
  expect(suppliedResult.status).toBe("succeeded")
  const answer = { document: "later-answers.json", locator: "/answers/0", quote: "The proposal still requires confirmation." }
  const consumer = await runChecklistCheckJob(directory, { sessions: sessionPort([JSON.stringify({ ...finding, source_refs: [answer] })]).sessions, model: null }, { ...checker, settled_instructions: [answer] })
  expect(consumer.status).toBe("succeeded")
  const notSupplied = sessionPort(Array.from({ length: LEAN_MAX_RETRIES + 1 }, () => JSON.stringify({ ...finding, source_refs: [answer] })))
  const rejected = await runChecklistCheckJob(directory, { sessions: notSupplied.sessions, model: null }, checker)
  expect(rejected.status).toBe("failed")
})

test("current body evidence is verified outside initial hints and altered output evidence receives correction", async () => {
  const directory = matter(source.quote)
  writeLeanPlan(directory, [{ title: "Arrangement", purpose: "Record arrangement", file: "draft.docx", parts: [{ heading: "Terms", scope: "Terms" }] }])
  const body = "The supplier must deliver sixty blue widgets by Friday."
  writeLeanBytes(directory, "d1/body-applied.docx", (await renderMarkdownToDocx(`# Arrangement\n\n${body}`)).bytes)
  acceptApplyForTest(directory)
  const evidence = { document: "draft.docx", locator: "accepted paragraph 1", text: body, kind: "body" as const }
  // No initial hint contains this paragraph: it is available by targeted lookup.
  const exact = { ...finding, evidence: [evidence] }
  const baseline = await runChecklistCheckJob(directory, { sessions: sessionPort([JSON.stringify(exact)]).sessions, model: null }, checker)
  expect(baseline.status).toBe("succeeded")
  expect(baseline.attempts).toBe(1)
  const port = sessionPort([JSON.stringify({ ...exact, evidence: [{ ...evidence, text: body.replace("sixty", "ninety") }] }), JSON.stringify({ ...exact, evidence: [{ ...evidence, document: "another.docx" }] }), JSON.stringify(exact)])
  const corrected = await runChecklistCheckJob(directory, { sessions: port.sessions, model: null }, checker)
  expect(corrected.status).toBe("succeeded")
  expect(corrected.attempts).toBe(3)
  expect(port.prompts[1].text).toContain("not verified in current body evidence")
  expect(port.prompts[2].text).toContain('document "another.docx"')
})

test("native comment text is accepted as comment and rejected as operative body or arbitrary artifact claim", async () => {
  const directory = matter(source.quote)
  writeLeanPlan(directory, [{ title: "Arrangement", purpose: "Record arrangement", file: "draft.docx", parts: [{ heading: "Terms", scope: "Terms" }] }])
  const body = (await renderMarkdownToDocx("# Arrangement\n\nThe original arrangement remains pending.")).bytes
  writeLeanBytes(directory, "d1/body-applied.docx", body)
  acceptApplyForTest(directory)
  const paragraph = readDocxCommentParagraphs(body).find(row => row.paragraph_text.includes("original arrangement"))!
  const comment = "A proposed replacement requires further authority."
  const annotated = exportDocxComments(body, [{ id: "review", author: "Reviewer", text: comment, anchor: { paragraph_index: paragraph.paragraph_index, paragraph_text: paragraph.paragraph_text } }])
  expect(annotated.unresolved).toEqual([])
  writeLeanBytes(directory, "d1/redline.docx", annotated.bytes)
  const evidence = { document: "draft.docx", locator: "redline comment 0", text: comment, kind: "comment" as const }
  const exact = { ...finding, evidence: [evidence] }
  expect((await runChecklistCheckJob(directory, { sessions: sessionPort([JSON.stringify(exact)]).sessions, model: null }, checker)).status).toBe("succeeded")
  const port = sessionPort([
    JSON.stringify({ ...exact, evidence: [{ ...evidence, kind: "body" }] }),
    JSON.stringify({ ...exact, evidence: [{ ...evidence, kind: "artifact", text: "The document is fully signed and contains all required changes." }] }),
    JSON.stringify(exact),
  ])
  const corrected = await runChecklistCheckJob(directory, { sessions: port.sessions, model: null }, checker)
  expect(corrected.status).toBe("succeeded")
  expect(corrected.attempts).toBe(3)
  expect(port.prompts[1].text).toContain("not verified in current body evidence")
  expect(port.prompts[2].text).toContain("not verified in current artifact evidence")
})

test("deleted revision evidence stays separate from the current operative body", async () => {
  const directory = matter(source.quote)
  writeLeanPlan(directory, [{ title: "Arrangement", purpose: "Record arrangement", file: "draft.docx", parts: [{ heading: "Terms", scope: "Terms" }] }])
  const old = (await renderMarkdownToDocx("# Arrangement\n\nObsoleteword arrangement.")).bytes
  const current = (await renderMarkdownToDocx("# Arrangement\n\nCurrentword arrangement.")).bytes
  writeLeanBytes(directory, "d1/body-applied.docx", current)
  acceptApplyForTest(directory)
  const redline = await (await docxodus()).compareDocuments(old, current, { authorName: "Reviewer" })
  writeLeanBytes(directory, "d1/redline.docx", redline)
  const deleted = readDocxRevisionText(redline).deleted_text.find(text => text.includes("Obsoleteword"))!
  expect(deleted).toBeDefined()
  const evidence = { document: "draft.docx", locator: "native deleted revision", text: deleted, kind: "deleted_text" as const }
  const exact = { ...finding, evidence: [evidence] }
  const port = sessionPort([JSON.stringify({ ...exact, evidence: [{ ...evidence, kind: "body" }] }), JSON.stringify(exact)])
  const result = await runChecklistCheckJob(directory, { sessions: port.sessions, model: null }, checker)
  expect(result.status).toBe("succeeded")
  expect(result.attempts).toBe(2)
  expect(port.prompts[1].text).toContain("not verified in current body evidence")
})


test("reinspection uses existing native checker and exact validation, without an automatic pass", async () => {
  const directory = matter(source.quote)
  const port = sessionPort([JSON.stringify({ ...finding, item_id: "wrong" }), JSON.stringify(finding)])
  const original = { ...finding, status: "unsatisfied" as const }
  const result = await runChecklistCheckJob(directory, { sessions: port.sessions, model: null }, checker, original)
  expect(port.created).toHaveLength(1)
  expect(port.prompts[0].text).toBe(buildChecklistReinspectionPrompt(checker, original))
  expect(port.prompts[1].text).toContain("item_id must equal")
  expect(result.status).toBe("succeeded")
  if (result.status === "succeeded") expect(result.value.status).toBe("indeterminate")
  expect(original.status).toBe("unsatisfied")
})


test("a refiner cannot certify retargeted children and may preserve source uncertainty instead", async () => {
  const directory = matter(source.quote)
  const input = { parent_id: "parent", parent: checker.item, source_unit: { id: "unit", role: "Original source", passages: [source] }, instructions: [],
    targets: [{ document: "draft.docx", role: "Agreement" }, { document: "memo.docx", role: "Advice" }], retrieval: "Exact source lookup." }
  const child = { ...checker.item, required_action: "propose" as const, target: { document: "memo.docx", role: "Advice", treatment: "advice" as const } }
  const changed = { parent_id: "parent", unit_id: "unit", disposition: "refined" as const, reason: "A proposal is supported.", items: [child], source_refs: [source], limitations: [] }
  const unresolved = { ...changed, disposition: "unresolved" as const, reason: "The source supports a proposal but does not establish the parent's target treatment.", limitations: ["The original parent cannot be certified."] }
  const deps = sessionPort([JSON.stringify(changed), JSON.stringify(unresolved)])
  const result = await runChecklistRefinementJob(directory, { ...deps, model: null }, input)
  expect(result.status).toBe("succeeded")
  expect(result.attempts).toBe(2)
  if (result.status === "succeeded") expect(result.value).toEqual(unresolved)
  expect(deps.prompts[1]!.text).toContain("return unresolved")
  expect(deps.prompts[1]!.text).toContain("Do not invent a duty")
  expect(input.parent).toEqual(checker.item)
})


test("native revision properties and inserted text are quotable artifact evidence but not invented properties or operative text", async () => {
  const directory = matter(source.quote)
  writeLeanPlan(directory, [{ title: "Arrangement", purpose: "Record arrangement", file: "draft.docx", parts: [{ heading: "Terms", scope: "Terms" }] }])
  const prior = (await renderMarkdownToDocx("Obsolete receipt.")).bytes
  const compared = (await renderMarkdownToDocx("Tracked new receipt.")).bytes
  const redline = await (await docxodus()).compareDocuments(prior, compared)
  writeLeanBytes(directory, "d1/redline.docx", redline)
  writeLeanBytes(directory, "d1/body-applied.docx", (await renderMarkdownToDocx("Different current receipt.")).bytes)
  acceptApplyForTest(directory)
  const context = buildChecklistDocumentContext(directory, { item_id: checker.item_id, item: checker.item })
  const annotation = context.artifacts[0]!.value.annotations!.find(row => row.artifact.endsWith("/redline.docx"))!
  const properties = JSON.parse(annotation.properties)
  expect(properties.accepted_body_matches_current).toBe(false)
  expect(properties.revision_elements.ins).toBeGreaterThan(0)
  const propertyEvidence = { document: annotation.artifact, locator: ".dochaus/lean/checklist-context/d1.json /annotations/1/properties", text: annotation.properties, kind: "artifact" as const }
  const insertedEvidence = { document: annotation.artifact, locator: ".dochaus/lean/checklist-context/d1.json /annotations/1/revisions/inserted_text/0", text: annotation.revisions.inserted_text[0]!, kind: "artifact" as const }
  const exact = { ...finding, evidence: [propertyEvidence, insertedEvidence] }
  const accepted = await runChecklistCheckJob(directory, { sessions: sessionPort([JSON.stringify(exact)]).sessions, model: null }, context.input)
  expect(accepted.status).toBe("succeeded")
  if (accepted.status === "succeeded") expect(accepted.value.status).toBe("indeterminate")
  const invented = { ...propertyEvidence, text: JSON.stringify({ ...properties, accepted_body_matches_current: true }) }
  const wrongBody = { ...insertedEvidence, document: "draft.docx", kind: "body" }
  const port = sessionPort([JSON.stringify({ ...finding, evidence: [invented] }), JSON.stringify({ ...finding, evidence: [wrongBody] }), JSON.stringify(exact)])
  const corrected = await runChecklistCheckJob(directory, { sessions: port.sessions, model: null }, context.input)
  expect(corrected.status).toBe("succeeded")
  expect(corrected.attempts).toBe(3)
  expect(port.prompts[1]!.text).toContain("not verified in current artifact evidence")
  expect(port.prompts[2]!.text).toContain("not verified in current body evidence")
})
