import { afterAll, beforeAll, expect, test } from "bun:test"
import { Hono } from "hono"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { installWordConsequences, consequenceContextKey, verifiedConsequenceFindings } from "./word-consequences"
import { writeWordChatPreferences } from "./word-chat-preferences"
import type { WordChatAuth } from "./word-chat-context"
import type { WordChatContext, WordChatTurn } from "../../../dochaus/lib/word-chat-contract"
import type { WordConsequenceAssessment } from "../../../dochaus/lib/word-consequence-contract"
type Native = { id: string; directory: string; parentID?: string; metadata: Record<string, unknown>; time: { created: number; updated: number; idle?: number }; outcome?: string }
const context: WordChatContext = { text: "Payment is due in thirty days.\rInterest follows the payment deadline.\rAn exception applies on holidays.", paragraphs: [
  { id: "p1", text: "Payment is due in thirty days." }, { id: "p2", text: "Interest follows the payment deadline." }, { id: "p3", text: "An exception applies on holidays." },
], saved: false, capturedAt: "2026-01-01T00:00:00Z", truncated: false }
let runtime: string, previous: string | undefined
beforeAll(() => { runtime = mkdtempSync(path.join(tmpdir(), "consequence-native-")); previous = process.env.DOCHAUS_RUNTIME_DIR; process.env.DOCHAUS_RUNTIME_DIR = runtime })
afterAll(() => { if (previous === undefined) delete process.env.DOCHAUS_RUNTIME_DIR; else process.env.DOCHAUS_RUNTIME_DIR = previous; rmSync(runtime, { recursive: true, force: true }) })
function harness() {
  writeWordChatPreferences({ consequences: true })
  const auth = { matter: { dir: "/synthetic/matter" }, grant: { document_id: "document", token_hash: "token" } } as WordChatAuth
  const rows = new Map<string, Native>(), outputs = new Map<string, object>(), stopped: string[] = [], submitted: Native[] = [], prompts: string[] = []
  let now = 1, hold = "", withFinding = true, failSubmission = false, onRead: ((route: string) => void) | undefined
  const root: Native = { id: "conversation", directory: auth.matter.dir, metadata: { type: "word-chat", document: "document" }, time: { created: now++, updated: 1 } }; rows.set(root.id, root)
  const turn: Native = { id: "amendment", directory: auth.matter.dir, parentID: root.id, time: { created: now++, updated: 1 }, metadata: { type: "word-chat-turn", kind: "amendment", document: "document", input: { text: "Shorten the payment deadline.", attachments: [] }, sources: [] } }; rows.set(turn.id, turn)
  const reviewed: WordChatTurn = { id: turn.id, conversationId: root.id, requestId: "foreground", kind: "amendment", scope: "document", text: "Shorten the payment deadline.", answer: "Reviewed.", status: "complete", sources: [], attachments: [], createdAt: 1,
    proposal: { id: "proposal", documentId: "document", before: context.paragraphs[0]!.text, after: "Payment is due in ten days.", explanation: "Requested change.", openItems: [], createdAt: context.capturedAt, target: { id: "target", text: context.paragraphs[0]!.text, paragraphs: [context.paragraphs[0]!], startOffset: 0, endOffset: context.paragraphs[0]!.text.length, capturedAt: context.capturedAt, saved: false } } }
  const complete = (s: Native) => {
    const stage = s.metadata.stage
    const output = stage === "meaning" ? { meaning: "The deadline becomes earlier.", uncertainties: [] } : stage === "discovery" ? { candidates: [{ paragraphId: "p2", relatedParagraphIds: ["p3"] }, { paragraphId: "p3", relatedParagraphIds: [] }] } : stage === "dependencies" ? { paragraphIds: ["p3"] } : { findings: withFinding ? [{ title: "Interest starts earlier", explanation: "The deadline change may advance interest, subject to the holiday exception.", kind: "effect", passage: { paragraphId: "p2", quote: "Interest follows the payment deadline." }, related: [{ paragraphId: "p3", quote: "An exception applies on holidays." }] }] : [] }
    if (stage === "judgment" && prompts.find(p => p.startsWith(s.id))?.includes('"targetPassage":{"id":"p3"')) (output as { findings: unknown[] }).findings = []
    outputs.set(s.id, output); s.time.idle = now++; s.outcome = "succeeded"
  }
  let app = new Hono(), dispose = () => {}
  function install() {
    dispose(); app = new Hono()
    dispose = installWordConsequences({
      route: (name, run) => app.post(`/word/chat/${name}`, async c => { try { return c.json(await run(c, auth) as object) } catch (e) { const error = e as Error & { status?: number }; return c.json({ error: error.message }, (error.status ?? 503) as 400) } }),
      body: c => c.req.json(), root: async (_auth, id) => { if (id !== root.id) throw Object.assign(new Error("wrong conversation"), { status: 404 }); return root },
      turns: async () => [...rows.values()].filter(s => s.metadata.type === "word-chat-turn"), reviewed: async () => reviewed,
      model: () => ({ providerID: "test", modelID: "fixture" }), error: (message, status = 400) => Object.assign(new Error(message), { status }),
      stop: async (_auth, id) => { stopped.push(id); rows.get(id)!.outcome = "interrupted" },
      api: async (_auth, route, method = "GET", body) => {
        onRead?.(route)
        if (route === "/api/session" && method === "POST") { const input = body as { id: string; parentID: string; metadata: Native["metadata"] }; if (!rows.has(input.id)) rows.set(input.id, { ...input, directory: auth.matter.dir, time: { created: now++, updated: now } }); return rows.get(input.id) }
        if (route.startsWith("/api/session?")) return [...rows.values()].filter(s => s.parentID === new URL(`http://local${route}`).searchParams.get("parentID"))
        const match = /^\/api\/session\/([^/]+)(.*)$/.exec(route)!, s = rows.get(match[1]!); if (!s) throw new Error("unknown native session")
        if (method === "PATCH") { s.metadata = (body as { metadata: Native["metadata"] }).metadata; return }
        if (match[2] === "/message") return [{ info: { role: "assistant" }, parts: [{ type: "text", text: JSON.stringify(outputs.get(s.id)) }] }]
        return s
      },
      submit: async (_auth, id, _messageID, prompt) => { if (failSubmission) { failSubmission = false; throw new Error("Lost context response before prompt admission") }; const s = rows.get(id)!; prompts.push(`${id}:${prompt}`); if (!submitted.some(job => job.id === id)) submitted.push(s); if (s.metadata.stage !== hold) complete(s) },
    }).dispose
  }
  install()
  const post = async (route: string, body: object) => { const response = await app.request(`/word/chat/${route}`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) }); return { status: response.status, body: await response.json() as WordConsequenceAssessment } }
  const assess = (requestId = "one", current = context) => post("assess", { conversationId: root.id, requestId, turnId: turn.id, context: current })
  const poll = (assessmentId: string, current = context) => post("assessment", { conversationId: root.id, assessmentId, context: current })
  async function settled(id: string, current = context) { let r = await poll(id, current); for (let i = 0; i < 9 && r.body.status === "running"; i++) r = await poll(id, current); return r }
  return { rows, outputs, stopped, submitted, prompts, turn, reviewed, assess, poll, post, settled, complete, install, dispose: () => dispose(), failNextSubmit: () => { failSubmission = true }, hold: (stage: string) => { hold = stage }, finding: (value: boolean) => { withFinding = value }, onRead: (f: (route: string) => void) => { onRead = f } }
}
test("current Word quotation checks reject invented evidence and mark inference", () => {
  const f = { title: "Effect", explanation: "Possible change", kind: "effect", passage: { paragraphId: "p2", quote: "Interest follows" } }
  expect(verifiedConsequenceFindings({ findings: [f] }, context, "key", new Set(["p2"]))[0]?.basis).toBe("inference")
  expect(verifiedConsequenceFindings({ findings: [{ ...f, passage: { paragraphId: "p2", quote: "invented" } }] }, context, "key", new Set(["p2"]))).toEqual([])
  expect(consequenceContextKey({ ...context, capturedAt: "later", saved: true })).toBe(consequenceContextKey(context))
  expect(consequenceContextKey({ ...context, truncated: true })).not.toBe(consequenceContextKey(context))
})
test("split native jobs resume after reconnect, keep foreground independent and expose meaning/coverage", async () => {
  const h = harness(); try { const a = await h.assess(); expect(a.body.status).toBe("running"); expect(h.submitted.map(s => s.metadata.stage)).toEqual(["meaning"]); h.install(); const r = await h.settled(a.body.id); expect(r.body.status).toBe("complete"); expect(r.body.meaning).toBe("The deadline becomes earlier."); expect(r.body.coverage).toEqual({ reviewed: 3, total: 3, candidates: 2, checked: 2 }); expect(r.body.findings[0]?.related?.[0]?.paragraphId).toBe("p3"); expect(h.turn.outcome).toBeUndefined(); expect(h.submitted.every(s => s.parentID === "conversation" && s.metadata.type !== "word-chat-turn")).toBe(true); expect((await h.post("assessment", { conversationId: "conversation", context })).body.id).toBe(a.body.id) } finally { h.dispose() }
})
test("no-finding is a checked bounded result", async () => { const h = harness(); h.finding(false); try { const a = await h.assess(); const r = await h.settled(a.body.id); expect(r.body.status).toBe("complete"); expect(r.body.findings).toEqual([]); expect(r.body.coverage.checked).toBe(2) } finally { h.dispose() } })
test("obsolete output cannot publish after Word or material instruction changes", async () => {
  const h = harness(); h.hold("judgment"); try { const a = await h.assess(); await h.poll(a.body.id); await h.poll(a.body.id); const changed = { ...context, paragraphs: context.paragraphs.map(p => p.id === "p3" ? { ...p, text: "The exception is removed." } : p), text: context.text.replace("An exception applies on holidays.", "The exception is removed.") }; expect((await h.poll(a.body.id, changed)).body.status).toBe("stale"); h.submitted.filter(s => s.metadata.stage === "judgment").forEach(h.complete); expect((await h.poll(a.body.id, changed)).body.findings).toEqual([]); const next = await h.assess("two", changed); h.rows.set("new-instruction", { ...h.turn, id: "new-instruction", time: { created: 100, updated: 100 }, metadata: { ...h.turn.metadata, kind: "chat", input: { text: "Keep the exception even if wording is unchanged.", attachments: [] } } }); expect((await h.poll(next.body.id, changed)).body.status).toBe("stale") } finally { h.dispose() }
})
test("disable stops jobs and admission; cancellation or disable during final publication suppresses output", async () => {
  const h = harness(); h.hold("meaning"); try { const a = await h.assess(); writeWordChatPreferences({ consequences: false }); await Promise.resolve(); expect((await h.poll(a.body.id)).body.status).toBe("disabled"); expect(h.stopped).toContain(h.submitted[0]!.id); expect((await h.assess("disabled")).body.status).toBe("disabled"); expect(h.submitted.length).toBe(1) } finally { h.dispose() }
  for (const reason of ["cancelled", "disabled"] as const) { const live = harness(); try { const a = await live.assess(); await live.settled(a.body.id); let lists = 0, fired = false; live.onRead(route => { if (route.startsWith("/api/session?")) lists++; if (route.startsWith("/api/session?") && lists >= 3 && !fired) { fired = true; if (reason === "disabled") writeWordChatPreferences({ consequences: false }); else live.rows.get(a.body.id)!.metadata.cancelled = true } }); const r = await live.poll(a.body.id); expect(r.body.status).toBe(reason); expect(r.body.findings).toEqual([]); expect(r.body.meaning).toBeUndefined() } finally { live.dispose() } }
})
test("scope and stable identity are enforced", async () => {
  const h = harness(); h.hold("meaning"); try { const a = await h.assess(); expect((await h.assess()).body.id).toBe(a.body.id); expect(h.submitted.length).toBe(1); expect((await h.post("assessment", { conversationId: "wrong", assessmentId: a.body.id })).status).toBe(404); expect((await h.post("assessment-cancel", { conversationId: "conversation", assessmentId: a.body.id })).body.status).toBe("cancelled"); h.rows.get(a.body.id)!.metadata.document = "other"; expect((await h.poll(a.body.id)).status).toBe(404) } finally { h.dispose() }
})
test("unchanged successful jobs are reused and prior native results remain historical", async () => {
  const h = harness(); try { const a = await h.assess(); await h.settled(a.body.id); const count = h.submitted.length; const same = await h.assess("same"); const r = await h.settled(same.body.id); expect(h.submitted.length).toBe(count); expect(r.body.history?.[0]?.status).toBe("stale"); expect(r.body.history?.[0]?.findings.length).toBe(1); const changed = { ...context, text: context.text + "\rUnrelated address.", paragraphs: [...context.paragraphs, { id: "p4", text: "Unrelated address." }] }; const next = await h.assess("added", changed); await h.settled(next.body.id, changed); expect(h.submitted.filter(s => s.metadata.stage === "meaning").length).toBe(1); expect(h.submitted.filter(s => s.metadata.stage === "judgment").length).toBe(3) } finally { h.dispose() }
})
test("invalid candidate IDs and quotations cannot become complete no-finding results", async () => {
  for (const stage of ["discovery", "judgment"]) { const h = harness(); h.hold(stage); try { const a = await h.assess(); await h.settled(a.body.id); const jobs = h.submitted.filter(s => s.metadata.stage === stage); for (const s of jobs) { h.complete(s); h.outputs.set(s.id, stage === "discovery" ? { candidates: [{ paragraphId: "invented", relatedParagraphIds: [] }] } : { findings: [{ title: "Invalid", explanation: "Unsupported", kind: "effect", passage: { paragraphId: "p2", quote: "invented" } }] }) } const r = await h.settled(a.body.id); expect(r.body.status).toBe("partial"); expect(r.body.gaps.length).toBeGreaterThan(0); if (stage === "judgment") expect(r.body.findings).toEqual([]); else expect(r.body.coverage.reviewed).toBe(0) } finally { h.dispose() } }
})
test("narrowed scope does not resurrect earlier matter evidence; application status invalidates old output", async () => {
  const h = harness(); try { h.turn.metadata.sources = [{ id: "old-matter", name: "Earlier matter", provenance: "saved-version", quote: "PRIVATE EARLIER MATERIAL" }]; h.rows.set("new-question", { ...h.turn, id: "new-question", time: { created: 100, updated: 100 }, metadata: { ...h.turn.metadata, kind: "chat", input: { text: "Use only this document.", attachments: [] }, sources: [] } }); const a = await h.assess(); const r = await h.settled(a.body.id); expect(r.body.status).toBe("complete"); expect(h.prompts.join("\n")).not.toContain("PRIVATE EARLIER MATERIAL"); h.reviewed.proposalApplied = true; expect((await h.poll(a.body.id)).body.status).toBe("stale") } finally { h.dispose() }
})

test("lost unprompted native admission resumes with the same durable request identity", async () => {
  const h = harness()
  try {
    h.failNextSubmit()
    expect((await h.assess()).status).toBe(503)
    expect(h.rows.size).toBe(4)
    const retry = await h.assess()
    expect(retry.status).toBe(200)
    expect((await h.settled(retry.body.id)).body.status).toBe("complete")
    expect(h.submitted.filter(s => s.metadata.stage === "meaning").length).toBe(1)
  } finally { h.dispose() }
})
test("truncated capture and changed row target remain explicit partial checks", async () => {
  const h = harness()
  try {
    const truncated = { ...context, truncated: true }
    const a = await h.assess("truncated", truncated)
    expect(a.status).toBe(200)
    const r = await h.settled(a.body.id, truncated)
    expect(r.body.status).toBe("partial")
    expect(r.body.gaps.join(" ")).toContain("outside this capture")
    const old = h.reviewed.proposal!; delete h.reviewed.proposal
    h.reviewed.rowProposal = { kind: "table_row_delete", id: "row", documentId: "document", explanation: "Remove row", openItems: [], createdAt: context.capturedAt, rowIndex: 1, cells: ["Missing target row."], table: { tableIndex: 0, rows: [[context.paragraphs[0]!], [{ id: "missing-row", text: "Missing target row." }]] } }
    const row = await h.assess("row")
    const checked = await h.settled(row.body.id)
    expect(checked.body.status).toBe("partial")
    expect(checked.body.gaps.join(" ")).toContain("target changed")
    h.reviewed.proposal = old
  } finally { h.dispose() }
})
test("disabled reconnect keeps the most recent result as explicitly historical", async () => {
  const h = harness()
  try {
    const a = await h.assess(); await h.settled(a.body.id)
    writeWordChatPreferences({ consequences: false })
    h.install()
    const result = await h.poll(a.body.id)
    expect(result.body.status).toBe("disabled")
    expect(result.body.findings).toEqual([])
    expect(result.body.history?.at(-1)?.status).toBe("disabled")
    expect(result.body.history?.at(-1)?.findings.length).toBe(1)
  } finally { h.dispose() }
})
