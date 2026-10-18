import { afterEach, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { focusedChangeReviewInputs, runFocusedChangeReview, runFocusedChangeReviewJob } from "./focused-change-review"
import { acceptChangePlan } from "./change-pipeline"
import { checkChanges } from "./change-checks"
import { readParagraphs } from "./checks"
import { writeLeanPlan } from "./plan"
import { packageCompositionBlock } from "./prompts"
import { drainLeanRequests, enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { LEAN_CHILD_PERMISSIONS } from "./stages"
import { readLeanJson, readLeanState, readLeanText, writeLeanText } from "./state"
import { cleanupMatters, happyPath, matter, requestMemo, scriptedModel, SOURCE, SUPPORT } from "./test-support"

afterEach(cleanupMatters)
async function fixture() {
  const directory = matter()
  await runLean(directory, { sessions: scriptedModel(happyPath).port })
  const paragraphs = readParagraphs(directory, "d1")
  const delivery = paragraphs.find(row => row.text === "The supplier shall deliver the goods.")!
  const payment = paragraphs.find(row => row.text === "Payment is due on delivery.")!
  const plan = { shared_decisions: [], groups: [{ id: "delivery", pids: [delivery.pid], goal: "State the supplied delivery period, preserving related payment timing.", sources: SUPPORT, limitations: [] }], dismissed: [], questions: [], findings: [] }
  expect(acceptChangePlan(directory, "d1", JSON.stringify(plan)).issues).toEqual([])
  const value = { amendments: [{ pid: delivery.pid, updated: "The supplier shall deliver the goods within thirty days of the order date.", reason: "States the source period.", sources: SUPPORT }], explanations: [{ pid: payment.pid, reason: "The initial draft preserves the source payment term.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] }], questions: [] }
  const checked = checkChanges(directory, JSON.stringify(value), paragraphs, { document: "d1" })
  expect(checked.issues).toEqual([])
  writeLeanText(directory, "d1/changes.json", JSON.stringify(checked.changes))
  return { directory, delivery, payment, plan }
}

test("coverage includes the current group, ungrouped explanation and unresolved disclosure without omnibus history", async () => {
  const { directory, delivery, payment, plan } = await fixture()
  const full = { ...plan, findings: [
    { disposition: "deferred", kind: "information", detail: "Confirm a material delivery exception before execution.", reason: "The exception is unresolved.", documents: ["d1"], sources: SUPPORT, limitations: [] },
    { disposition: "rejected", kind: "suggestion", detail: "UNRELATED_REJECTED_HISTORY", reason: "Not requested.", documents: ["d1"], sources: SUPPORT, limitations: [] },
  ] }
  expect(acceptChangePlan(directory, "d1", JSON.stringify(full)).issues).toEqual([])
  writeLeanText(directory, "d1/review-wording.md", "UNRELATED_RAW_REVIEW_HISTORY".repeat(2000))
  const inputs = focusedChangeReviewInputs(directory, "d1")
  expect(inputs).toHaveLength(3)
  expect(inputs.flatMap(row => row.pids)).toContain(delivery.pid)
  expect(inputs.flatMap(row => row.pids)).toContain(payment.pid)
  expect(inputs.some(row => row.id === "disclosure:0")).toBe(true)
  for (const input of inputs) {
    expect(input.text).toContain("explicit dependency task")
    expect(input.text).toContain("view=proposed_body")
    expect(input.text).toContain("even when this assignment has no inline target PID")
    expect(input.text).toContain("never infer an omission in the pending revision from that older view")
    expect(input.text).toContain("retrieve its actual relevant final clause")
    expect(input.text).toContain("before asserting law")
    expect(input.text).not.toContain("UNRELATED_RAW_REVIEW_HISTORY")
    expect(input.text).not.toContain("UNRELATED_REJECTED_HISTORY")
    expect(input.text.length).toBeLessThan(32_001)
  }
})

test("one supported defect prevents aggregate pass while all focused results and limitations survive", async () => {
  const { directory, delivery } = await fixture()
  const port = scriptedModel({ "lean-change-review": text => text.includes('"id":"group:delivery"')
    ? JSON.stringify({ status: "changes-required", findings: [{ pid: delivery.pid, detail: "The delivery exception conflicts with the dependent payment clause; inspect and conform both." }], limitations: ["Keep the source qualification."] })
    : JSON.stringify({ status: "pass", findings: [], limitations: ["Review draft only."] }) })
  const result = await runFocusedChangeReview(directory, readLeanState(directory)!, { sessions: port.port }, "d1", () => {})
  expect(result.status).toBe("done")
  expect(port.sessions).toHaveLength(2)
  expect(port.sessions.every(row => JSON.stringify(row.permissions) === JSON.stringify(LEAN_CHILD_PERMISSIONS))).toBe(true)
  const review = readLeanJson(directory, "d1/change-review.json") as { status: string; findings: unknown[]; limitations: string[] }
  expect(review.status).toBe("changes-required")
  expect(review.findings).toHaveLength(1)
  expect(review.limitations).toEqual(["Keep the source qualification.", "Review draft only."])
})

test("a missing worker result cannot overwrite the previous review with partial passes", async () => {
  const { directory } = await fixture()
  const before = readLeanText(directory, "d1/change-review.json")
  const port = scriptedModel({})
  const result = await runFocusedChangeReview(directory, readLeanState(directory)!, { sessions: { ...port.port, async prompt(input) {
    if (input.text.includes('"id":"retained:')) throw Error("test transport failure")
    return JSON.stringify({ status: "pass", findings: [], limitations: [] })
  } } }, "d1", () => {})
  expect(result.status).toBe("failed")
  expect(readLeanText(directory, "d1/change-review.json")).toBe(before)
})

test("invalid pass-with-findings is corrected through the native contract without dropping the defect", async () => {
  const { directory, delivery } = await fixture()
  const port = scriptedModel({ "lean-change-review": (text, index) => JSON.stringify({ status: index === 0 ? "pass" : "changes-required", findings: [{ pid: delivery.pid, detail: "A supported delivery condition is missing." }], limitations: [] }) })
  const result = await runFocusedChangeReview(directory, readLeanState(directory)!, { sessions: port.port }, "d1", () => {})
  expect(result.status).toBe("done")
  expect((readLeanJson(directory, "d1/change-review.json") as { status: string }).status).toBe("changes-required")
  expect(port.prompts.some(row => row.text.includes("A pass cannot contain material correction findings"))).toBe(true)
})

test("oversized assignments and conflicting same-PID operations fail before a model admission", async () => {
  const { directory, plan } = await fixture()
  writeLeanText(directory, "request.md", "indispensable instruction ".repeat(2000))
  expect(() => focusedChangeReviewInputs(directory, "d1")).toThrow("exceeds bounded context")
  writeLeanText(directory, "request.md", "Draft the supplied terms.")
  const changes = readLeanJson(directory, "d1/changes.json") as { amendments: unknown[] }
  changes.amendments.push(changes.amendments[0])
  writeLeanText(directory, "d1/changes.json", JSON.stringify(changes))
  expect(() => focusedChangeReviewInputs(directory, "d1")).toThrow("one replacement per paragraph")
  expect(plan.groups).toHaveLength(1)
})

test("an actual input change during review restarts without accepting a stale union", async () => {
  const { directory } = await fixture()
  const before = readLeanText(directory, "d1/change-review.json")
  const port = scriptedModel({})
  const result = await runFocusedChangeReview(directory, readLeanState(directory)!, { sessions: { ...port.port, async prompt() {
    writeLeanText(directory, "request.md", "Changed instruction received during inspection.")
    return JSON.stringify({ status: "pass", findings: [], limitations: [] })
  } } }, "d1", () => {})
  expect(result.status).toBe("restart")
  expect(readLeanText(directory, "d1/change-review.json")).toBe(before)
})

test("an ungrouped retained amendment is reviewed alongside its existing explanation", async () => {
  const { directory, payment } = await fixture()
  const changes = readLeanJson(directory, "d1/changes.json") as { amendments: unknown[] }
  const checked = checkChanges(directory, JSON.stringify({ amendments: [{ pid: payment.pid, updated: "Payment is due on delivery, subject to the supplied conditions.", reason: "Preserves payment timing while disclosing conditions.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] }], questions: [] }), readParagraphs(directory, "d1"), { document: "d1" })
  expect(checked.issues).toEqual([])
  changes.amendments.push(checked.changes!.amendments[0])
  writeLeanText(directory, "d1/changes.json", JSON.stringify(changes))
  const inputs = focusedChangeReviewInputs(directory, "d1")
  expect(inputs).toHaveLength(2)
  const retained = inputs.find(row => row.id === `retained:${payment.pid}`)!
  expect(retained.text).toContain("Payment is due on delivery, subject to the supplied conditions.")
  expect(retained.text).toContain("The initial draft preserves the source payment term.")
})

test("a focused cross-document defect keeps its earlier owner and cannot be redirected to a future body", async () => {
  const { directory, delivery } = await fixture()
  writeLeanPlan(directory, [
    { title: "Operative terms", purpose: "State the supplied terms.", parts: [{ heading: "Terms", scope: "Supplied terms" }] },
    { title: "Advice", purpose: "Explain the operative terms and remaining decisions.", parts: [{ heading: "Advice", scope: "Source issues" }] },
  ])
  for (const name of ["paragraphs.json", "changes.json", "change-plan.json"]) writeLeanText(directory, `d2/${name}`, readLeanText(directory, `d1/${name}`)!)
  const port = scriptedModel({ "lean-change-review": () => JSON.stringify({ status: "changes-required", findings: [{ document: "d1", pid: delivery.pid, detail: "The actual earlier operative clause contradicts this advice's stated delivery period." }], limitations: [] }) })
  const result = await runFocusedChangeReview(directory, readLeanState(directory)!, { sessions: port.port }, "d2", () => {})
  expect(result.status).toBe("done")
  const review = readLeanJson(directory, "d2/change-review.json") as { findings: { document: string }[] }
  expect(review.findings.every(row => row.document === "d1")).toBe(true)
  const future = scriptedModel({ "lean-change-review": () => JSON.stringify({ status: "changes-required", findings: [{ document: "d3", detail: "Invalid future-body target." }], limitations: [] }) })
  expect((await runFocusedChangeReview(directory, readLeanState(directory)!, { sessions: future.port }, "d2", () => {})).status).toBe("failed")
  expect(readLeanJson(directory, "d2/change-review.json")).toEqual(review)
})

test("later included instructions reach every reviewer with exact provenance while queued updates are not authority", async () => {
  const { directory } = await fixture()
  const update = enqueueLeanRequest(directory, { kind: "update", text: "Preserve the agreed thirty-day delivery period; do not substitute a different period." })
  drainLeanRequests(directory)
  enqueueLeanRequest(directory, { kind: "update", text: "QUEUED_NOT_YET_GOVERNING" })
  for (const input of focusedChangeReviewInputs(directory, "d1")) {
    expect(input.text).toContain("Preserve the agreed thirty-day delivery period; do not substitute a different period.")
    expect(input.text).toContain(`requests/done/${update.id}.json /text`)
    expect(input.text).not.toContain("QUEUED_NOT_YET_GOVERNING")
  }
})

test("source-grounded restoration of an ungrouped initial paragraph remains in semantic coverage", async () => {
  const { directory, payment } = await fixture()
  const changes = readLeanJson(directory, "d1/changes.json") as { explanations: unknown[]; retained_initial?: unknown[] }
  changes.explanations = []
  changes.retained_initial = [{ pid: payment.pid, reason: "Withdraw the earlier unsupported payment change and preserve the source-backed initial term.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] }]
  writeLeanText(directory, "d1/changes.json", JSON.stringify(changes))
  const input = focusedChangeReviewInputs(directory, "d1").find(row => row.id === `retained:${payment.pid}`)!
  expect(input).toBeDefined()
  expect(input.text).toContain('"retained_initial":[{"pid":')
  expect(input.text).toContain("Withdraw the earlier unsupported payment change")
  expect(input.text).toContain("Payment is due on delivery.")
  expect(input.text).toContain("Verify each retained_initial withdrawal against its canonical supporting words")
})


test("single focused review entry uses the same native validation without publishing an aggregate or accepting an invalid pass", async () => {
  const { directory, delivery } = await fixture()
  const before = readLeanText(directory, "d1/change-review.json")
  const input = focusedChangeReviewInputs(directory, "d1")[0]!
  let calls = 0
  const native = scriptedModel({ "lean-change-review": () => {
    calls++
    return JSON.stringify({ status: calls === 1 ? "pass" : "changes-required", findings: [{ pid: delivery.pid, detail: "The proposed condition requires source-grounded adjudication before acceptance." }], limitations: [] })
  } })
  const result = await runFocusedChangeReviewJob(directory, readLeanState(directory)!, { sessions: native.port }, "d1", input, () => {})
  expect(calls).toBe(2)
  expect(result.issues).toEqual([])
  expect(result.review?.status).toBe("changes-required")
  expect(result.review?.findings[0]?.pid).toBe(delivery.pid)
  expect(readLeanText(directory, "d1/change-review.json")).toBe(before)
})


async function largeRelatedFixture() {
  const { directory } = await fixture()
  writeLeanPlan(directory, [
    { title: "Terms", purpose: "State supplied duties.", parts: [{ heading: "Terms", scope: "Supplied duties" }] },
    { title: "Advice", purpose: "Explain conditions in the operative terms.", parts: [{ heading: "Advice", scope: "Source issues and dependencies" }] },
  ])
  const paragraphs = Array.from({ length: 18 }, (_, index) => ({ pid: `p${String(index).padStart(12, "0")}`, text: `Related passage ${index}. ` + "Preserve this condition and its dependent exception. ".repeat(16) })).map(row => ({ ...row, sha256: paragraphSha256(row.text) }))
  const group = { id: "related-conditions", pids: paragraphs.map(row => row.pid), goal: "Review the combined conditional effect of every target, preserving the dependency on the operative terms and the unresolved approval qualification.", sources: SUPPORT, limitations: ["Approval remains unknown; this is advice, not an agreed alteration."] }
  writeLeanText(directory, "d2/paragraphs.json", JSON.stringify(paragraphs))
  expect(acceptChangePlan(directory, "d2", JSON.stringify({ shared_decisions: [], groups: [group], dismissed: [], findings: [], questions: [] })).issues).toEqual([])
  const changes = { amendments: paragraphs.slice(0, -2).map(row => ({ pid: row.pid, updated: row.text + "Disclose the source-supported dependency without inferring approval. ".repeat(8), reason: "Explain the supplied condition and preserve its exception.", sources: SUPPORT })),
    explanations: [{ pid: paragraphs.at(-2)!.pid, reason: "The original target already preserves the supplied condition.", sources: SUPPORT }],
    retained_initial: [{ pid: paragraphs.at(-1)!.pid, reason: "Withdraw an unsupported alteration; preserve the supplied initial condition.", sources: SUPPORT }], questions: [] }
  const checked = checkChanges(directory, JSON.stringify(changes), paragraphs, { document: "d2" })
  expect(checked.issues).toEqual([])
  writeLeanText(directory, "d2/changes.json", JSON.stringify({ ...checked.changes, retained_initial: changes.retained_initial }))
  return { directory, paragraphs, group }
}

test("oversized related goals decompose by owned target without losing operations, qualifications or cross-document dependencies", async () => {
  const { directory, paragraphs, group } = await largeRelatedFixture()
  const inputs = focusedChangeReviewInputs(directory, "d2")
  expect(inputs).toHaveLength(paragraphs.length)
  expect(inputs.flatMap(row => row.pids).sort()).toEqual(paragraphs.map(row => row.pid).sort())
  for (const input of inputs) {
    expect(input.pids).toHaveLength(1)
    expect(input.text.length).toBeLessThan(32_001)
    const assignment = JSON.parse(input.text.split("<assignment>")[1]!.split("</assignment>")[0]!)
    expect(assignment.goal).toEqual(group)
    expect(assignment.assigned_pids).toEqual(input.pids)
    const operations = JSON.parse(input.text.split("<cumulative_operations>\n")[1]!.split("\n</cumulative_operations>")[0]!)
    expect([...operations.amendments, ...operations.explanations, ...operations.retained_initial].map(row => row.pid)).toEqual(input.pids)
    expect(input.text).toContain("retrieve their actual proposed passages")
    expect(input.text).toContain("retrieve its actual relevant final clause")
    expect(input.text).toContain("view=proposed_body")
  }
})

test("every decomposed target must settle; a sibling cross-document defect survives the strict union", async () => {
  const { directory, paragraphs } = await largeRelatedFixture()
  const earlier = readParagraphs(directory, "d1")[0]!
  const before = readLeanText(directory, "d2/change-review.json")
  const failing = scriptedModel({ "lean-change-review": text => {
    if (text.includes(`"assigned_pids":["${paragraphs.at(-1)!.pid}"]`)) throw Error("missing one target review")
    return JSON.stringify({ status: "pass", findings: [], limitations: [] })
  } })
  expect((await runFocusedChangeReview(directory, readLeanState(directory)!, { sessions: failing.port }, "d2", () => {})).status).toBe("failed")
  expect(readLeanText(directory, "d2/change-review.json")).toBe(before)
  const complete = scriptedModel({ "lean-change-review": text => text.includes(`"assigned_pids":["${paragraphs.at(-1)!.pid}"]`)
    ? JSON.stringify({ status: "changes-required", findings: [{ document: "d1", pid: earlier.pid, detail: "The advice conflicts with the actual earlier operative condition; reconcile the combined effect." }], limitations: [] })
    : JSON.stringify({ status: "pass", findings: [], limitations: [] }) })
  expect((await runFocusedChangeReview(directory, readLeanState(directory)!, { sessions: complete.port }, "d2", () => {})).status).toBe("done")
  expect(complete.sessions).toHaveLength(paragraphs.length)
  expect(readLeanJson(directory, "d2/change-review.json")).toMatchObject({ status: "changes-required", findings: [{ document: "d1", pid: earlier.pid }] })
})

test("a requested memo reaches every focused review assignment, disclosure included; without one they are unchanged", async () => {
  const { directory, plan } = await fixture()
  expect(acceptChangePlan(directory, "d1", JSON.stringify({ ...plan, findings: [{ disposition: "deferred", kind: "information", detail: "Confirm a material delivery exception before execution.", reason: "The exception is unresolved.", documents: ["d1"], sources: SUPPORT, limitations: [] }] })).issues).toEqual([])
  const texts = () => new Map(focusedChangeReviewInputs(directory, "d1").map(row => [row.id, row.text]))
  const without = texts()
  expect(without.has("disclosure:0")).toBe(true)
  requestMemo(directory)
  const block = packageCompositionBlock(directory, "d1")
  expect(block).toStartWith("<package_composition>\n")
  const memo = texts()
  expect([...memo.keys()]).toEqual([...without.keys()])
  for (const [id, text] of memo) {
    expect([id, text.split(block).length - 1]).toEqual([id, 1])
    expect([id, text.replace(`\n\n${block}`, "")]).toEqual([id, without.get(id)!])
    expect([id, without.get(id)!.includes("package_composition")]).toEqual([id, false])
  }
})
