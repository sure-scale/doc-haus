import { afterEach, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { runFocusedNotes } from "./focused-notes"
import { checkNotesCoverageReply, NOTES_COVERAGE_AGENT } from "./focused-notes-coverage"
import { checkNotesEvidenceRequest, notesDependencyEvidence, resolveFocusedNotesDependency, type NotesEvidenceRequest } from "./focused-notes-dependencies"
import { writeLeanPlan } from "./plan"
import { LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { readLeanState, readLeanText, writeLeanText } from "./state"
import { cleanupMatters, matter, SOURCE, SOURCE_TEXT } from "./test-support"

afterEach(cleanupMatters)
const claim = "The source requires delivery within thirty days of the order date."
const prose = `${claim} The supplied payment term remains payable on delivery.`
const request: NotesEvidenceRequest = { source_id: SOURCE, claim, question: "What delivery timing and qualifications does this source state?" }

function fixture() {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Explain the supplied delivery and payment terms.", parts: [{ heading: "Delivery", scope: "Delivery timing." }] }])
  const text = "The supplier must deliver within thirty days of the order date."
  const rows = [{ pid: "p111111111111", text, sha256: paragraphSha256(text), draft_part: 0 }]
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(rows))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(rows))
  return directory
}

test("source requests bind an exact prose claim to one canonical source, distinct from body targets", () => {
  const reference = { documents: [{ document: "d1" as const }], prose, sources: [SOURCE] }
  const verdict = { status: "needs-evidence" as const, reason: "The source passage was not supplied.", request }
  expect(checkNotesCoverageReply(JSON.stringify(verdict), reference).value).toEqual(verdict)
  for (const invalid of [{ ...request, document: "d1" }, { ...request, source_id: "unknown.docx" }, { ...request, claim: claim.toLowerCase() }]) {
    expect(checkNotesCoverageReply(JSON.stringify({ ...verdict, request: invalid }), reference).value).toBeUndefined()
    expect(checkNotesEvidenceRequest(invalid, reference.documents, prose, [SOURCE]).request).toBeUndefined()
  }
})

test("source selection copies verified canonical context, preserving qualifications instead of model-supplied text", async () => {
  const directory = fixture()
  const prompts: string[] = []
  const port: StageSessionPort = {
    async createSession() { return "source-selection" },
    async prompt(input) {
      prompts.push(input.text)
      return JSON.stringify({ quotes: [{ source_id: SOURCE, quote: "within thirty days of the order date", text: "Invented payment waiver." }], limitation: "" })
    },
  }
  const result = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: port }, { id: "passage:d1:0", text: "One delivery explanation." }, prose, request, [], () => {})
  expect(result.status).toBe("resolved")
  if (result.status !== "resolved") throw new Error(result.reason)
  expect(prompts).toHaveLength(1)
  expect(prompts[0]).toContain("<source_dependency_lookup>")
  const evidence = JSON.stringify(result.evidence)
  expect(evidence).toContain(SOURCE_TEXT)
  expect(evidence).not.toContain("Invented payment waiver")
  expect(result.observation.sources).toEqual([{ source_id: SOURCE, section: expect.any(String), text: SOURCE_TEXT, excerpt: false }])
  expect(result.observation.targets).toEqual([])
  const combined = JSON.parse(notesDependencyEvidence([result.observation, result.observation], [])[0]!.text)
  expect(combined.sources).toHaveLength(1)
  expect(combined.observations).toHaveLength(2)
  expect(combined.observations.every((observation: object) => !("sources" in observation))).toBe(true)
})

test("a long source supplies a bounded canonical excerpt with an explicit incomplete scope", async () => {
  const source = `${"Earlier storage wording. ".repeat(200)}Delivery takes place only after the container is inspected. ${"Later storage wording. ".repeat(200)}`
  const directory = matter(source)
  const port: StageSessionPort = {
    async createSession() { return "long-source-selection" },
    async prompt() { return JSON.stringify({ quotes: [{ source_id: SOURCE, quote: "Delivery takes place only after the container is inspected." }], limitation: "Further storage terms were not inspected." }) },
  }
  const result = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: port }, { id: "passage:d1:0", text: "One delivery explanation." }, prose, request, [], () => {})
  expect(result.status).toBe("resolved")
  if (result.status !== "resolved") throw new Error(result.reason)
  expect(result.observation.sources).toHaveLength(1)
  expect(result.observation.sources![0]!.text.length).toBeLessThanOrEqual(4_000)
  expect(result.observation.sources![0]!.text).toContain("only after the container is inspected")
  expect(result.observation.sources![0]!.excerpt).toBe(true)
  expect(result.limitation).toBe("Further storage terms were not inspected.")
})

test("fabricated source quotations fail bounded selection; unavailable inspection remains an explicit limit", async () => {
  const directory = fixture()
  let calls = 0
  const port: StageSessionPort = {
    async createSession() { return "source-selection" },
    async prompt() { calls++; return JSON.stringify({ quotes: [{ source_id: SOURCE, quote: "The delivery obligation is waived." }], limitation: "" }) },
  }
  const input = { id: "passage:d1:0", text: "One delivery explanation." }
  const failed = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: port }, input, prose, request, [], () => {})
  expect(failed.status).toBe("failed")
  expect(calls).toBe(LEAN_MAX_RETRIES + 1)
  port.prompt = async () => JSON.stringify({ quotes: [], limitation: "The material source qualification could not be inspected." })
  const unavailable = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: port }, input, prose, request, [], () => {})
  expect(unavailable.status).toBe("resolved")
  if (unavailable.status !== "resolved") throw new Error(unavailable.reason)
  expect(unavailable.observation.sources).toEqual([])
  expect(unavailable.limitation).toContain("could not be inspected")
})

test("source lookup reaches the independent checker on unchanged prose before a writer retry", async () => {
  const directory = fixture()
  const packets: string[] = [], sourcePrompts: string[] = []
  let sessions = 0, passageWrites = 0
  const port: StageSessionPort = {
    async createSession() { return `source-flow-${++sessions}` },
    async prompt(input) {
      if (input.agent === NOTES_COVERAGE_AGENT) {
        if (!input.text.includes(`<proposed_prose>\n${prose}`)) return JSON.stringify({ status: "pass", reason: "The scripted section explains its scoped assignment." })
        packets.push(input.text)
        return JSON.stringify(input.text.includes('"sources":[') ? { status: "pass", reason: "The canonical source supports this prose and its qualifications." }
          : { status: "needs-evidence", reason: "The writer's source passage is not in this packet.", request })
      }
      if (input.text.includes("<source_dependency_lookup>")) {
        sourcePrompts.push(input.text)
        return JSON.stringify({ quotes: [{ source_id: SOURCE, quote: "within thirty days of the order date" }], limitation: "" })
      }
      if (input.text.includes("<section_kind>\npassage")) { passageWrites++; return prose }
      return "The note accompanies the supplied delivery terms for the lawyer."
    },
  }
  expect((await runFocusedNotes(directory, readLeanState(directory)!, { sessions: port }, () => {})).status).toBe("done")
  expect(passageWrites).toBe(1)
  expect(sourcePrompts).toHaveLength(1)
  expect(packets).toHaveLength(2)
  expect(packets[0]).not.toContain(SOURCE_TEXT)
  expect(packets[1]).toContain(SOURCE_TEXT)
  expect(packets.every(packet => packet.includes(`<proposed_prose>\n${prose}`))).toBe(true)
  expect(readLeanText(directory, "drafting-notes.md")).toContain(prose)
})
