import { afterEach, expect, test } from "bun:test"
import { readdirSync } from "node:fs"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { NOTES_COVERAGE_AGENT } from "./focused-notes-coverage"
import { runFocusedNotes } from "./focused-notes"
import { writeLeanPlan } from "./plan"
import { LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { leanFile, readLeanState, readLeanText, writeLeanText, writeOpenItems, type OpenItem } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)
const point: OpenItem = { id: "apply:omission:seal-design", stage: "apply", kind: "omission", status: "open",
  detail: "The seal design remains unknown; the custodian must supply it.", ref: { document: "d2", pid: "p222222222222" } }
const claim = "The container handover requires inspection before acceptance."
const original = `${claim} ${point.detail}`
const corrected = `Acceptance is immediate. ${point.detail}`
const paragraph = (pid: string, text: string) => ({ pid, text, sha256: paragraphSha256(text), draft_part: 0 })

function fixture(condition: string) {
  const directory = matter("Keep the archive label on the container.", "Prepare the container handover and storage notice.")
  writeLeanPlan(directory, [
    { title: "Container handover", purpose: "Record container acceptance.", parts: [{ heading: "Acceptance", scope: "Acceptance condition." }] },
    { title: "Storage notice", purpose: "Record storage instructions.", parts: [{ heading: "Storage", scope: "Keep the label visible." }] },
  ])
  for (const [document, rows] of [["d1", [paragraph("p111111111111", condition)]], ["d2", [paragraph("p222222222222", "Keep the label visible in storage.")]]] as const) {
    writeLeanText(directory, `${document}/paragraphs.json`, JSON.stringify(rows))
    writeLeanText(directory, `${document}/paragraphs-applied.json`, JSON.stringify(rows))
  }
  writeOpenItems(directory, [point])
  return directory
}

function model(mode: "supports" | "contradicts" | "unanswered", miscopyFirstClaim = false) {
  const initial = new Map<string, string>()
  const writes: string[] = [], checks: string[] = [], lookups: string[] = []
  const checkPackets: { sessionID: string; text: string; initial: string }[] = []
  let ownerWrites = 0
  let sessionCount = 0
  let requests = 0
  const port: StageSessionPort = {
    async createSession() { return `session-${++sessionCount}` },
    async prompt(input) {
      if (!initial.has(input.sessionID)) initial.set(input.sessionID, input.text)
      const first = initial.get(input.sessionID)!
      if (input.agent === NOTES_COVERAGE_AGENT) {
        checks.push(input.text)
        checkPackets.push({ sessionID: input.sessionID, text: input.text, initial: first })
        if (!first.includes(point.detail)) return JSON.stringify({ status: "pass", reason: "The section explains its own assigned point." })
        if (!first.includes("dependency_evidence") || mode === "unanswered") return JSON.stringify({ status: "needs-evidence", reason: "The present claim needs the companion's actual acceptance condition.",
          request: { document: "d1", claim: miscopyFirstClaim && requests++ === 0 ? claim.toLowerCase() : claim, question: "What is the current acceptance condition?" } })
        if (mode === "contradicts" && first.includes(`<proposed_prose>\n${original}`)) return JSON.stringify({ status: "rewrite", reason: "The inspected current companion provides immediate acceptance; the prose incorrectly describes a condition." })
        return JSON.stringify({ status: "pass", reason: "The same prose is supported by the actual companion and preserves the unknown design and required action." })
      }
      if (first.includes("<current_target_lookup>")) {
        lookups.push(input.text)
        return JSON.stringify({ targets: [{ document: "d1", pids: ["p111111111111"] }], limitation: "" })
      }
      writes.push(input.text)
      if (first.includes(point.detail)) { ownerWrites++; return mode === "contradicts" && ownerWrites > 1 ? corrected : original }
      return "The note preserves the container label instruction."
    },
  }
  return { port, writes, checks, checkPackets, lookups, get ownerWrites() { return ownerWrites } }
}

test("an inexact dependency claim is corrected by the same verifier before lookup, without repeating the writer", async () => {
  const directory = fixture("Acceptance requires inspection.")
  const scripted = model("supports", true)
  expect((await runFocusedNotes(directory, readLeanState(directory)!, { sessions: scripted.port }, () => {})).status).toBe("done")
  const ownerChecks = scripted.checkPackets.filter(packet => packet.initial.includes(point.detail))
  expect(ownerChecks).toHaveLength(3)
  expect(ownerChecks[1]!.sessionID).toBe(ownerChecks[0]!.sessionID)
  expect(ownerChecks[1]!.text).toContain("claim")
  expect(ownerChecks[1]!.text).toContain("Correct only")
  expect(ownerChecks[2]!.sessionID).not.toBe(ownerChecks[0]!.sessionID)
  expect(ownerChecks[2]!.initial).toContain("dependency_evidence")
  expect(scripted.lookups).toHaveLength(1)
  expect(scripted.ownerWrites).toBe(1)
  expect(readLeanText(directory, "drafting-notes.md")).toContain(original)
})

for (const mode of ["supports", "contradicts"] as const) test(`dependency inspection ${mode} the same proposed prose before any writer retry or acceptance`, async () => {
  const directory = fixture(mode === "supports" ? "Acceptance requires inspection." : "Acceptance is immediate.")
  const scripted = model(mode)
  const result = await runFocusedNotes(directory, readLeanState(directory)!, { sessions: scripted.port }, () => {})
  expect(result.status).toBe("done")
  const ownerChecks = scripted.checks.filter(text => text.includes(point.detail))
  expect(ownerChecks[0]).toContain(`<proposed_prose>\n${original}`)
  expect(ownerChecks[0]).not.toContain("dependency_evidence")
  expect(ownerChecks[0]).toContain("Container handover")
  expect(ownerChecks[1]).toContain(`<proposed_prose>\n${original}`)
  expect(ownerChecks[1]).toContain("dependency_evidence")
  expect(ownerChecks[1]).toContain(mode === "supports" ? "Acceptance requires inspection." : "Acceptance is immediate.")
  expect(scripted.ownerWrites).toBe(mode === "supports" ? 1 : 2)
  const notes = readLeanText(directory, "drafting-notes.md")!
  expect(notes).toContain(mode === "supports" ? original : corrected)
  expect(notes.split(`[ref: ${point.id}]`)).toHaveLength(2)
  const calls = scripted.writes.length + scripted.checks.length + scripted.lookups.length
  expect((await runFocusedNotes(directory, readLeanState(directory)!, { sessions: scripted.port }, () => {})).status).toBe("done")
  expect(scripted.writes.length + scripted.checks.length + scripted.lookups.length).toBe(calls)
})

test("an unanswered evidence request never falls through to acceptance within bounded writer recovery", async () => {
  const directory = fixture("Acceptance requires inspection.")
  const scripted = model("unanswered")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, { sessions: scripted.port }, () => {})).status).toBe("failed")
  expect(scripted.ownerWrites).toBe(LEAN_MAX_RETRIES + 1)
  expect(scripted.lookups).toHaveLength(LEAN_MAX_RETRIES * (LEAN_MAX_RETRIES + 1))
  expect(readLeanText(directory, "drafting-notes.md")).toBeUndefined()
  const saved = readdirSync(leanFile(directory, "notes-sections")).map(file => readLeanText(directory, `notes-sections/${file}`))
  expect(saved.length).toBeGreaterThan(0)
  expect(saved).not.toContain(`${original}\n`)
})

const globalClaim = "The handover contains no acceptance condition anywhere."
const unsupportedScope = `${globalClaim} ${point.detail}`
const scoped = `The inspected paragraph does not establish the handover's acceptance condition. ${point.detail} Until supplied, the seal cannot be finalized.`

function scopedModel(mode: "scoped" | "essential" | "contradiction") {
  const initial = new Map<string, string>()
  const writes: string[] = [], checks: string[] = []
  let ownerWrites = 0, lookups = 0, sessionCount = 0
  const port: StageSessionPort = {
    async createSession() { return `scope-session-${++sessionCount}` },
    async prompt(input) {
      if (!initial.has(input.sessionID)) initial.set(input.sessionID, input.text)
      const first = initial.get(input.sessionID)!
      if (input.agent === NOTES_COVERAGE_AGENT) {
        if (!first.includes(point.detail)) return JSON.stringify({ status: "pass", reason: "The separate section explains its own duty." })
        checks.push(input.text)
        if (first.includes(`<proposed_prose>\n${unsupportedScope}`)) return JSON.stringify({ status: "needs-evidence", reason: "The selected paragraph cannot establish a document-wide negative.",
          request: { document: "d1", claim: globalClaim, question: "What current evidence supports this document-wide absence claim?" } })
        if (mode === "essential") {
          expect(first).toContain("Explain every governing acceptance condition.")
          return JSON.stringify({ status: "rewrite", reason: "The original duty requires every acceptance condition; a limited inspection cannot replace that essential explanation." })
        }
        if (mode === "contradiction") return JSON.stringify({ status: "rewrite", reason: "The supplied current text expressly requires inspection. An unverified caveat cannot hide that contradiction." })
        expect(first).toContain(scoped)
        expect(first).toContain("Only the labelled paragraph was inspected")
        return JSON.stringify({ status: "pass", reason: "The narrowed claim matches its actual inspection limit and preserves the unknown design, consequence and required action." })
      }
      if (first.includes("<current_target_lookup>")) {
        lookups++
        return JSON.stringify({ targets: [{ document: "d1", pids: ["p111111111111"] }], limitation: "Only the labelled paragraph was inspected; other current passages were not inspected." })
      }
      if (first.includes(point.detail)) {
        writes.push(input.text)
        ownerWrites++
        if (ownerWrites === 1) return unsupportedScope
        expect(input.text).toContain(globalClaim)
        if (ownerWrites === 2) expect(input.text).toContain("unsupported at this inspection scope")
        expect(input.text).toContain("Only the labelled paragraph was inspected")
        return mode === "contradiction" ? `Whether acceptance requires inspection remains unverified. ${point.detail}` : scoped
      }
      return "The separate section preserves the container label instruction."
    },
  }
  return { port, writes, checks, get ownerWrites() { return ownerWrites }, get lookups() { return lookups } }
}

for (const mode of ["scoped", "essential", "contradiction"] as const) test(`completed bounded inspections use checked writer scope recovery (${mode})`, async () => {
  const directory = fixture(mode === "contradiction" ? "Acceptance requires inspection." : "Keep the archive label visible.")
  if (mode === "essential") writeOpenItems(directory, [{ ...point, detail: `${point.detail} Explain every governing acceptance condition.` }])
  const scripted = scopedModel(mode)
  const result = await runFocusedNotes(directory, readLeanState(directory)!, { sessions: scripted.port }, () => {})
  expect(result.status).toBe(mode === "scoped" ? "done" : "failed")
  expect(scripted.lookups).toBe(LEAN_MAX_RETRIES)
  expect(scripted.ownerWrites).toBe(mode === "scoped" ? 2 : LEAN_MAX_RETRIES + 1)
  expect(scripted.checks.slice(0, LEAN_MAX_RETRIES + 1).every(text => text.includes(`<proposed_prose>\n${unsupportedScope}`))).toBe(true)
  const inspected = scripted.checks.slice(1, LEAN_MAX_RETRIES + 1)
  const evidenceText = mode === "contradiction" ? "Acceptance requires inspection." : "Keep the archive label visible."
  expect(inspected.map(text => text.split(evidenceText).length)).toEqual(Array(LEAN_MAX_RETRIES).fill(inspected[0]!.split(evidenceText).length))
  if (mode !== "scoped") {
    expect(readLeanText(directory, "drafting-notes.md")).toBeUndefined()
    return
  }
  const notes = readLeanText(directory, "drafting-notes.md")!
  expect(notes).toContain(scoped)
  expect(notes).not.toContain(globalClaim)
  expect(notes.split(`[ref: ${point.id}]`)).toHaveLength(2)
  const before = scripted.ownerWrites
  expect((await runFocusedNotes(directory, readLeanState(directory)!, { sessions: scripted.port }, () => {})).status).toBe("done")
  expect(scripted.ownerWrites).toBe(before)
})

test("legacy assembled passage prose is not a checked answer to its original question, while accepted neighbors are reused", async () => {
  const directory = fixture("Acceptance requires inspection.")
  const first = model("supports")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, { sessions: first.port }, () => {})).status).toBe("done")
  const files = readdirSync(leanFile(directory, "notes-sections"))
  const owner = files.find(file => readLeanText(directory, `notes-sections/${file}`)?.trim() === original)!
  expect(owner).toBeDefined()
  const legacy = "### Current target passage\n\nThe storage notice keeps the label visible.\n"
  writeLeanText(directory, `notes-sections/${owner}`, legacy)
  const resumed = model("supports")
  expect((await runFocusedNotes(directory, readLeanState(directory)!, { sessions: resumed.port }, () => {})).status).toBe("done")
  expect(resumed.ownerWrites).toBe(1)
  expect(resumed.writes).toHaveLength(1)
  expect(readLeanText(directory, `notes-sections/${owner}`)).toBe(legacy)
  const notes = readLeanText(directory, "drafting-notes.md")!
  expect(notes).toContain(original)
  expect(notes).not.toContain("### Current target passage")
  expect(notes.split(`[ref: ${point.id}]`)).toHaveLength(2)
})
