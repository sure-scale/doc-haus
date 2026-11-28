import { afterEach, describe, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import type { ParagraphRow } from "./checks"
import { readLeanEvents } from "./events"
import { runFocusedNotes } from "./focused-notes"
import { checkNotesCoverageReply, NOTES_COVERAGE_AGENT } from "./focused-notes-coverage"
import { checkNotesEvidenceRequest, notesDependencyEvidence, resolveFocusedNotesDependency, type NotesEvidenceRequest } from "./focused-notes-dependencies"
import type { NotesBody } from "./focused-notes-evidence"
import { writeLeanPlan } from "./plan"
import { LEAN_CHILD_PERMISSIONS, LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"
import { readLeanState, readLeanText, writeLeanText } from "./state"
import { cleanupMatters, matter, SOURCE, SOURCE_TEXT } from "./test-support"

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
const selected = { targets: [{ document: "d1" as const, pids: ["p222222222222"] }], limitation: "" }

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
  expect(checkNotesEvidenceRequest(request, [{ document: "d1" }], prose)).toEqual({ request, issues: [] })
  for (const invalid of [null, [], {}, { ...request, document: "d9" }, { ...request, claim: "" },
    { ...request, claim: "An earlier reply contained this claim." }, { ...request, question: " " }])
    expect(checkNotesEvidenceRequest(invalid, [body], prose).request).toBeUndefined()
  for (const changed of [claim.toLowerCase(), claim.replace(".", "!")]) {
    const checked = checkNotesEvidenceRequest({ ...request, claim: changed }, [{ document: "d1" }], prose)
    expect(checked.request).toBeUndefined()
    expect(checked.issues.join(" ")).toContain("request.claim")
    expect(checked.issues.join(" ")).not.toContain(changed)
  }
  expect(checkNotesEvidenceRequest({ ...request, document: "d9" }, [{ document: "d1" }], prose).issues.join(" ")).toContain('request.document="d9"')
})

test("lookup reconstructs actual accepted text, mapped baseline and qualifications without trusting model quotations", async () => {
  const directory = matter()
  const model = locator([{ ...selected, quotation: "Invented unconditional approval." }])
  const result = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, input, prose, request, [body], () => {})
  expect(result.status).toBe("resolved")
  if (result.status !== "resolved") throw new Error(result.reason)
  expect(result.evidence.map(row => row.name)).toEqual(["dependency_evidence"])
  const packet = JSON.parse(result.evidence[0]!.text)
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
  expect(result.limitation).toContain("does not establish document-wide absence")
})

test("the shipped notes role recognizes the dependency lookup and only its one claim is inlined", async () => {
  const directory = matter()
  const model = locator([selected])
  const unrelated = "Unrelated section context. ".repeat(1_000)
  const oversized = `${unrelated}${prose}`
  await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, { ...input, text: unrelated }, oversized, request, [body], () => {})
  const prompt = model.prompts[0]!.text
  expect(model.sessions[0]!.agent).toBe("lean-notes-evidence")
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
  for (const invalid of [{ ...request, claim: "not in prose" }, { ...request, claim: claim.toLowerCase() }, { ...request, document: "d9" as const }])
    expect((await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: unused.port }, input, prose, invalid, [body], () => {})).status).toBe("failed")
  expect(unused.sessions).toHaveLength(0)
  expect(readLeanEvents(directory).filter(event => event.event === "checker").every(event => !JSON.stringify(event).includes(claim))).toBe(true)
})

test("a resolved lookup retains its actual claim-specific selection and limitation as an observation", async () => {
  const directory = matter()
  const model = locator([selected])
  const result = await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, input, prose, request, [body], () => {})
  if (result.status !== "resolved") throw new Error(result.reason)
  expect(result.observation).toEqual({ request, targets: selected.targets, limitation: "" })
  expect(result.limitation).toContain("does not establish document-wide absence")
})

test("a subsequent lookup receives relevant prior selection and limits without repeating canonical paragraphs", async () => {
  const directory = matter()
  const prior = [{ request: { ...request, question: "Which earlier condition applies?" }, targets: selected.targets, limitation: "The later inspection exception remains uninspected." },
    { request: { ...request, document: "d2" as const, claim: "An unrelated carrier condition." }, targets: [], limitation: "Unrelated record limitation." }]
  const model = locator([selected])
  await resolveFocusedNotesDependency(directory, readLeanState(directory)!, { sessions: model.port }, input, prose, request, [body], () => {}, prior)
  const prompt = model.prompts[0]!.text
  expect(prompt).toContain("<prior_dependency_inspections>")
  expect(prompt).toContain(prior[0]!.limitation)
  expect(prompt).toContain(prior[0]!.request.question)
  expect(prompt).toContain("p222222222222")
  expect(prompt).not.toContain(prior[1]!.limitation)
  for (const paragraph of [...body.initial, ...body.current]) expect(prompt).not.toContain(paragraph.text)
})

test("overlapping dependency observations share canonical nodes while retaining every question and actual limitation", () => {
  const observations = [
    { request, targets: selected.targets, limitation: "The remaining term was not inspected." },
    { request: { ...request, claim: "A different claim remains conditional.", question: "What later condition applies?" }, targets: [{ document: "d1" as const, pids: ["p222222222222", "p333333333333"] }], limitation: "" },
    { request: { ...request, claim: "A missing exception remains unknown.", question: "Is an exception available?" }, targets: [], limitation: "The exception could not be inspected." },
  ]
  const evidence = notesDependencyEvidence(observations, [body])
  expect(evidence.map(row => row.name)).toEqual(["dependency_evidence"])
  const packet = JSON.parse(evidence[0]!.text)
  expect(packet.observations).toEqual(observations)
  const current = packet.context.filter((row: { name: string }) => row.name === "current_accepted_evidence").map((row: { text: string }) => JSON.parse(row.text))
  const initial = packet.context.filter((row: { name: string }) => row.name === "generated_initial_evidence").map((row: { text: string }) => JSON.parse(row.text))
  expect(current).toHaveLength(1)
  expect(current[0].paragraphs.map((row: { pid: string }) => row.pid)).toEqual(["p222222222222", "p333333333333"])
  expect(current[0].paragraphs.map((row: { text: string }) => row.text)).toEqual(body.current.map(row => row.text))
  expect(initial).toHaveLength(1)
  expect(initial[0].paragraphs.map((row: { text: string }) => row.text)).toEqual(body.initial.map(row => row.text))
  expect(current[0].baseline_pairs).toEqual([{ initial: "p111111111111", current: ["p222222222222", "p333333333333"] }])
  expect(current[0].applied_outcomes).toEqual([{ id: "a1", pid: "p111111111111", mode: "code", result_pids: ["p222222222222", "p333333333333"] }])
  expect(packet.role).toContain("not an additional explanation owner")
  expect(packet.scope).toContain("never establishes document-wide absence")
  expect(notesDependencyEvidence([], [body])).toEqual([])
})

describe("focused notes source dependencies", () => {
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
})
