import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import type { ParagraphRow } from "./checks"
import { readLeanEvents } from "./events"
import { checkNotesEvidenceRequest, resolveFocusedNotesDependency, type NotesEvidenceRequest } from "./focused-notes-dependencies"
import type { NotesBody } from "./focused-notes-evidence"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { readLeanState } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)
const row = (pid: string, text: string): ParagraphRow => ({ pid, text, sha256: "synthetic", draft_part: 0 })
const body: NotesBody = { document: "d1", initial: [row("p111111111111", "Acceptance is immediate.")], current: [
  row("p222222222222", "The custodian accepts the container."),
  row("p333333333333", "Acceptance takes effect only after inspection."),
], applied: [{ id: "a1", pid: "p111111111111", mode: "code", result_pids: ["p222222222222", "p333333333333"] }] }
const claim = "The acceptance remains conditional."
const prose = `The memo explains the handover. ${claim}`
const request: NotesEvidenceRequest = { document: "d1", claim, question: "What condition governs acceptance in the current handover?" }
const input = { id: "change:d2:synthetic", text: "Explain the memo revision, using the handover only as dependency evidence." }
const selected = { targets: [{ document: "d1", pids: ["p222222222222"] }], limitation: "" }

function locator(replies: unknown[]) {
  const sessions: Parameters<StageSessionPort["createSession"]>[0][] = []
  const prompts: Parameters<StageSessionPort["prompt"]>[0][] = []
  const port: StageSessionPort = {
    async createSession(value) { sessions.push(value); return "synthetic-dependency" },
    async prompt(value) { prompts.push(value); return JSON.stringify(replies[Math.min(prompts.length - 1, replies.length - 1)]) },
  }
  return { sessions, prompts, port }
}

test("dependency request belongs to an exact current prose claim and a planned document", () => {
  expect(checkNotesEvidenceRequest(request, [body], prose)).toEqual({ request, issues: [] })
  for (const invalid of [null, [], {}, { ...request, document: "d9" }, { ...request, claim: "" },
    { ...request, claim: "An earlier reply contained this claim." }, { ...request, question: " " }])
    expect(checkNotesEvidenceRequest(invalid, [body], prose).request).toBeUndefined()
})

test("lookup reconstructs actual accepted text, mapped baseline and qualifications without trusting model quotations", async () => {
  const directory = matter()
  const model = locator([{ ...selected, quotation: "Invented unconditional approval." }])
  const result = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, input, prose, request, [body], () => {})
  expect(result.status).toBe("resolved")
  if (result.status !== "resolved") throw new Error(result.reason)
  expect(result.evidence.map(row => row.name)).toEqual(["dependency_evidence"])
  const packet = JSON.parse(result.evidence[0]!.text)
  expect(packet.role).toContain("not an additional explanation owner")
  expect(packet.context.find((row: { name: string }) => row.name === "generated_initial_evidence").text).toContain("Acceptance is immediate.")
  const current = packet.context.find((row: { name: string }) => row.name === "current_accepted_evidence").text
  expect(current).toContain("The custodian accepts the container.")
  expect(current).toContain("only after inspection")
  expect(JSON.stringify(result)).not.toContain("Invented unconditional approval")
  expect(model.sessions).toHaveLength(1)
  expect(model.sessions[0]!.permissions).toEqual(LEAN_CHILD_PERMISSIONS)
  expect(model.sessions[0]!.permissions![0]).toEqual({ action: "*", resource: "*", effect: "deny" })
  expect(model.sessions[0]!.permissions!.some(row => row.action.includes("write") && row.effect === "allow")).toBe(false)
  expect(model.prompts[0]!.text).toContain("<current_target_lookup>")
  expect(model.prompts[0]!.text).toContain(claim)
  expect(model.prompts[0]!.text).toContain("whether it supports or contradicts")
  expect(result.limitation).toContain("does not establish document-wide absence")
})

test("the shipped notes role recognizes the dependency lookup and only its one claim is inlined", async () => {
  const directory = matter()
  const model = locator([selected])
  const unrelated = "Unrelated section context. ".repeat(1_000)
  const oversized = `${unrelated}${prose}`
  await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, { ...input, text: unrelated }, oversized, request, [body], () => {})
  const prompt = model.prompts[0]!.text
  const role = readFileSync(new URL("../../agent/lean-drafting-notes.md", import.meta.url), "utf8")
  expect(role).toContain("explicitly labelled current_target_lookup")
  expect(prompt).toStartWith("<current_target_lookup>\n")
  expect(prompt).toContain(claim)
  expect(prompt).not.toContain(unrelated)
  expect(prompt).not.toContain("The memo explains the handover.")
  expect(prompt).toContain("targets")
})

test("contradictory current evidence is retained rather than upgraded to approval", async () => {
  const directory = matter()
  const changed = { ...body, current: [row("p222222222222", "The custodian refuses acceptance pending inspection.")] }
  const model = locator([selected])
  const result = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, input, prose, request, [changed], () => {})
  expect(result.status).toBe("resolved")
  if (result.status !== "resolved") throw new Error(result.reason)
  expect(result.evidence[0]!.text).toContain("refuses acceptance")
  expect(result.evidence[0]!.text).toContain("must still pass its ordinary checker")
})

test("unavailable evidence remains an inspection limitation, never missing body content", async () => {
  const directory = matter()
  const model = locator([{ targets: [], limitation: "The requested current clause could not be located." }])
  const result = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, input, prose, request, [body], () => {})
  expect(result.status).toBe("resolved")
  if (result.status !== "resolved") throw new Error(result.reason)
  expect(result.limitation).toBe("The requested current clause could not be located.")
  expect(JSON.parse(result.evidence[0]!.text).context).toEqual([])
})

test("invented, historical or wrong-document targets cannot supply trusted dependency evidence", async () => {
  const directory = matter()
  for (const target of [{ document: "d1", pids: ["p999999999999"] }, { document: "d1", pids: ["p111111111111"] }, { document: "d2", pids: ["p222222222222"] }]) {
    const model = locator([{ targets: [target], limitation: "" }])
    const result = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, input, prose, request, [body], () => {})
    expect(result.status).toBe("failed")
    expect(model.sessions).toHaveLength(1)
    expect(model.prompts).toHaveLength(LEAN_MAX_RETRIES + 1)
  }
  const unused = locator([selected])
  expect((await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: unused.port }, input, prose, { ...request, claim: "not in prose" }, [body], () => {})).status).toBe("failed")
  expect(unused.sessions).toHaveLength(0)
  expect(readLeanEvents(directory).filter(event => event.event === "checker").every(event => !JSON.stringify(event).includes(claim))).toBe(true)
})
