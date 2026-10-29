import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import {
  chronologyPreflight,
  chronologySections,
  startChronologyRun,
  stopChronologyRun,
  awaitChronologyRun,
  parseChronologyDate,
  parseChronologyExtraction,
  resolveRelativeDate,
  assessChronology,
  recentChronologyInstructions,
  chronologySources,
  normaliseChronologyDate,
  type ChronologyGenerationDeps,
} from "./chronology-generation"
import { actOnChronology, decideChronologyProposal, mutateChronology, readChronology } from "./chronology-store"
import { openDb, upsertDocument } from "./db"
import { liveText } from "../../../dochaus/lib/extract"
import type { StageSessionPort } from "../../../dochaus/workflows/lean/stages"
const dirs: string[] = []
const previous = process.env.DOCHAUS_RUNTIME_DIR
afterEach(() => {
  if (previous === undefined) delete process.env.DOCHAUS_RUNTIME_DIR
  else process.env.DOCHAUS_RUNTIME_DIR = previous
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})
function matter() {
  const dir = mkdtempSync(path.join(tmpdir(), "chronology-generation-"))
  dirs.push(dir)
  writeFileSync(
    path.join(dir, "matter.json"),
    JSON.stringify({ title: "Disputed instrument calibration", client: "Laboratory owner" }),
  )
  const runtime = path.join(dir, "runtime")
  mkdirSync(path.join(runtime, "config/opencode"), { recursive: true })
  writeFileSync(path.join(runtime, "config/opencode/opencode.json"), JSON.stringify({ model: "fixture/local" }))
  process.env.DOCHAUS_RUNTIME_DIR = runtime
  return dir
}
function source(dir: string, name: string, body: string, quality?: string, role = "source") {
  const file = path.join(dir, name)
  writeFileSync(file, body)
  const db = openDb(dir)
  try {
    const id = upsertDocument(db, file, name, Date.now(), null, {}, quality ?? null)
    if (role !== "source") db.run("UPDATE documents SET role='deliverable' WHERE id=?", [id])
    return id
  } finally {
    db.close()
  }
}
function event(title: string, quote: string, extra: Record<string, unknown> = {}) {
  return {
    entryKind: "occurrence",
    title,
    detail: `The record reports ${title.toLowerCase()}.`,
    quote,
    date: { original: "6 May 2025", precision: "exact", start: "2025-05-06" },
    people: [],
    status: "reported",
    significance: "background",
    privilege: "uncertain",
    gaps: [],
    ...extra,
  }
}
const complete = (events: ReturnType<typeof event>[]) => JSON.stringify({ complete: true, events, gaps: [] })
function fake(
  handler: (input: Parameters<StageSessionPort["prompt"]>[0]) => string | Promise<string>,
): ChronologyGenerationDeps {
  let id = 0
  return {
    engine: { fetch: (async () => Response.json({ data: [] })) as unknown as typeof fetch },
    sessions: {
      async createSession(input) {
        expect(input.permissions).toEqual([{ action: "*", resource: "*", effect: "deny" }])
        return `job-${++id}`
      },
      async prompt(input) {
        return handler(input)
      },
      async interrupt() {},
    },
  }
}
const sectionText = (prompt: string) =>
  JSON.parse(prompt.match(/<source-evidence>\n([^]*?)\n<\/source-evidence>/)![1]!) as string
function matchFirst(prompt: string, relation = "same") {
  const candidates = JSON.parse(prompt.split("\nCandidates: ")[1]!) as { id: string }[]
  return JSON.stringify({ eventId: candidates[0]?.id ?? null, relation })
}

test("every section is read; supplied-only accounts survive and archive/unreadable/failure outcomes stay honest", async () => {
  const dir = matter(),
    first = "The instrument was collected on 6 May 2025.",
    last = "The calibration certificate was signed on 9 May 2025."
  source(dir, "Long note.md", first + "\n" + "Routine background record. ".repeat(850) + "\n" + last)
  source(
    dir,
    "Unreadable.md",
    "",
    JSON.stringify({ status: "no_readable_text", signals: [{ code: "empty", detail: "No text could be extracted" }] }),
  )
  source(dir, "Failed.pdf", "bad pdf")
  source(dir, "Bundle.zip", "not read")
  source(dir, "Generated.md", "Derived work product", undefined, "deliverable")
  const supplied = "On 2 May 2025 the inspection was requested; origin not identified."
  writeFileSync(path.join(dir, "Chronology.md"), supplied)
  const calls: string[] = []
  const deps = fake((input) => {
    if (input.agent === "chronology-match") return JSON.stringify({ eventId: null, relation: "independent" })
    const body = sectionText(input.text)
    calls.push(body)
    return complete([
      ...(body.includes(first) ? [event("Instrument collected", first)] : []),
      ...(body.includes(last) ? [event("Certificate signed", last)] : []),
      ...(body.includes(supplied) ? [event("Inspection requested", supplied)] : []),
    ])
  })
  deps.readText = (file) =>
    file.endsWith("Failed.pdf") ? Promise.reject(new Error("Unreadable PDF structure")) : liveText(file)
  expect(chronologyPreflight(dir).sources).toHaveLength(5)
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, deps)
  await awaitChronologyRun(dir)
  const saved = readChronology(dir)
  expect(saved.events).toHaveLength(3)
  expect(saved.events.every((e) => !e.reviewed && !e.opened)).toBe(true)
  expect(saved.events.find((e) => e.title === "Inspection requested")?.accounts[0]?.evidence[0]?.provenance).toBe(
    "supplied-chronology",
  )
  expect(saved.coverage.find((c) => c.documentName === "Long note.md")).toMatchObject({
    status: "complete",
    sections: 3,
    processedSections: [0, 1, 2],
  })
  expect(saved.coverage.find((c) => c.documentName === "Unreadable.md")?.status).toBe("unreadable")
  expect(saved.coverage.find((c) => c.documentName === "Failed.pdf")?.status).toBe("failed")
  expect(saved.coverage.find((c) => c.documentName === "Bundle.zip")?.status).toBe("unsupported")
  expect(calls).toHaveLength(4)
  expect(readFileSync(path.join(dir, "Chronology.md"), "utf8")).toBe(supplied)
})

test("failed quote leaves its section retryable; resume preserves usable draft without duplicate accounts", async () => {
  const dir = matter(),
    quote = "The instrument was collected on 6 May 2025.",
    other = "The courier returned an unsigned form."
  source(dir, "Record.md", `${quote}\n${other}`)
  await startChronologyRun(
    dir,
    { expectedRevision: 0, privilege: "all-cleared", mode: "initial" },
    fake(() =>
      complete([
        event("Instrument collected", quote),
        event("Form returned", "This sentence never existed in the source."),
      ]),
    ),
  )
  await awaitChronologyRun(dir)
  let state = readChronology(dir)
  expect(state.events).toHaveLength(1)
  expect(state.coverage[0]?.completedKeys).toHaveLength(0)
  expect(state.run?.status).toBe("failed")
  const retry = fake((input) =>
    input.agent === "chronology-match"
      ? JSON.stringify({ eventId: null, relation: "independent" })
      : complete([event("Instrument collected", quote), event("Form returned", other)]),
  )
  await startChronologyRun(dir, { expectedRevision: state.revision, privilege: "mixed", mode: "resume" }, retry)
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(state.events).toHaveLength(2)
  expect(state.events.reduce((n, e) => n + e.accounts.length, 0)).toBe(2)
  expect(state.coverage[0]).toMatchObject({ status: "complete", events: 2 })
  expect(state.run?.privilege).toBe("mixed")
})

test("refresh groups duplicate new accounts into one proposal and preserves lawyer wording and hidden choices", async () => {
  const dir = matter(),
    first = "The instrument was collected on 6 May 2025."
  source(dir, "Initial.md", first)
  await startChronologyRun(
    dir,
    { expectedRevision: 0, privilege: "mixed", mode: "initial" },
    fake(() => complete([event("Instrument collected", first)])),
  )
  await awaitChronologyRun(dir)
  let state = readChronology(dir),
    old = state.events[0]!
  for (const action of [
    { type: "edit", eventId: old.id, title: "Lawyer collection wording", detail: "Lawyer account", date: old.date },
    { type: "note", eventId: old.id, note: "Keep signed envelope" },
    { type: "key", eventId: old.id, key: true },
    { type: "hide", eventId: old.id, hidden: true },
  ] as const)
    state = actOnChronology(dir, state.revision, action) as typeof state
  const q1 = "The technician completed calibration on 8 May 2025.",
    q2 = "Calibration was completed by the technician on 8 May 2025."
  source(dir, "Technician.md", q1)
  source(dir, "Manager.md", q2)
  const refresh = fake((input) =>
    input.agent === "chronology-match"
      ? matchFirst(input.text)
      : complete([event("Calibration completed", sectionText(input.text).includes(q1) ? q1 : q2)]),
  )
  await startChronologyRun(
    dir,
    { expectedRevision: state.revision, privilege: "all-cleared", mode: "refresh" },
    refresh,
  )
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(state.events).toHaveLength(1)
  expect(state.events[0]).toMatchObject({
    title: "Lawyer collection wording",
    hidden: true,
    key: true,
    note: "Keep signed envelope",
  })
  expect(state.proposals).toHaveLength(1)
  expect(state.proposals[0]?.kind).toBe("addition")
  expect(state.proposals[0]?.events[0]?.accounts).toHaveLength(2)
  state = decideChronologyProposal(dir, state.revision, state.proposals[0]!.id, "accept")
  expect(state.events).toHaveLength(2)
  expect(state.events[1]?.reviewed).toBe(true)
  expect(state.events[0]?.hidden).toBe(true)
})

test("a stopped run ignores a late reply and repeating its admitted request does not dispatch twice", async () => {
  const dir = matter(),
    quote = "The instrument was collected on 6 May 2025."
  source(dir, "Record.md", quote)
  let resolve!: (s: string) => void,
    entered!: () => void,
    calls = 0
  const admitted = new Promise<void>((r) => {
    entered = r
  })
  const deps = fake(async () => {
    calls++
    entered()
    return new Promise<string>((r) => {
      resolve = r
    })
  })
  const request = {
    expectedRevision: 0,
    privilege: "mixed" as const,
    mode: "initial" as const,
    requestId: "request-idempotent-001",
  }
  await startChronologyRun(dir, request, deps)
  await admitted
  expect((await startChronologyRun(dir, request, deps)).run?.id).toBe(request.requestId)
  expect(calls).toBe(1)
  stopChronologyRun(dir, readChronology(dir).revision)
  resolve(complete([event("Instrument collected", quote)]))
  await awaitChronologyRun(dir)
  expect(readChronology(dir).events).toHaveLength(0)
  expect(readChronology(dir).run?.status).toBe("stopped")
})

test("ambiguous repeated quotation is an issue unless the occurrence is identified", async () => {
  const dir = matter(),
    quote = "The instrument was collected on 6 May 2025."
  source(dir, "Repeated.md", `${quote}\nA different speaker repeats this.\n${quote}`)
  await startChronologyRun(
    dir,
    { expectedRevision: 0, privilege: "mixed", mode: "initial" },
    fake(() => complete([event("Instrument collected", quote)])),
  )
  await awaitChronologyRun(dir)
  let state = readChronology(dir)
  expect(state.events).toHaveLength(0)
  expect(state.coverage[0]?.issues.join(" ")).toContain("ambiguous repeated")
  const canonical = await liveText(path.join(dir, "Repeated.md")),
    last = canonical.lastIndexOf(quote)
  await startChronologyRun(
    dir,
    { expectedRevision: state.revision, privilege: "mixed", mode: "resume" },
    fake(() => complete([event("Instrument collected", quote, { quoteStart: last })])),
  )
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(state.events[0]?.accounts[0]?.evidence[0]?.charStart).toBe(last)
})

test("date precision stays honest; only explicit supported calendar-day formulas calculate", () => {
  expect(parseChronologyDate({ original: "May 2025", precision: "month", start: "2025-05" }).start).toBe("2025-05")
  expect(() => parseChronologyDate({ original: "May 2025", precision: "month", start: "2025-05-01" })).toThrow(
    "invent a day",
  )
  expect(() => parseChronologyDate({ original: "03/04/25", precision: "ambiguous", start: "2025-03-04" })).toThrow(
    "silently resolved",
  )
  const date = parseChronologyDate({
    original: "within 3 calendar days after this note",
    precision: "relative",
    relative: {
      expression: "within 3 calendar days after",
      anchorDocument: "Note.md",
      anchorDate: "2025-05-06",
      offsetDays: 3,
      interpretation: "document-stated",
    },
  })
  resolveRelativeDate(date, { documentName: "Note.md", sourceId: "note", version: "current" }, "2025-05-06", {})
  expect(date.start).toBe("2025-05-09")
  const business = structuredClone(date)
  business.relative!.expression = "within 3 business days after"
  resolveRelativeDate(business, { documentName: "Note.md", sourceId: "note", version: "current" }, "2025-05-06", {})
  expect(business.start).toBeUndefined()
  expect(() => parseChronologyExtraction('{"complete":false,"events":[]}')).toThrow("complete section reading")
})

test("discovery uses actual context, caches assessment and preserves Hide without extracting", async () => {
  const dir = matter()
  source(dir, "Note.md", "The instrument arrived late and a dated account would help resolve the dispute.")
  let calls = 0
  const deps = fake((input) => {
    calls++
    expect(input.agent).toBe("chronology-assess")
    expect(input.text).not.toContain("practice_areas")
    return JSON.stringify({ status: "relevant", reason: "The record describes disputed timing." })
  })
  let state = await assessChronology(dir, deps)
  expect(state.suggestion.status).toBe("relevant")
  expect(state.events).toHaveLength(0)
  state = await assessChronology(dir, deps)
  expect(calls).toBe(1)
  state = actOnChronology(dir, state.revision, { type: "invitation", choice: "hide" }) as typeof state
  source(dir, "Later.md", "A further dated note.")
  state = await assessChronology(dir, deps)
  expect(state.suggestion.status).toBe("hidden")
  expect(calls).toBe(1)
})

test("section splitting covers all characters with bounded overlap", () => {
  const body = "abcdef".repeat(6000),
    sections = chronologySections(body)
  expect(sections[0]?.start).toBe(0)
  expect(sections.at(-1)?.end).toBe(body.length)
  for (let i = 1; i < sections.length; i++) expect(sections[i]!.start).toBe(sections[i - 1]!.end - 700)
})

test("scheduled obligation and completion remain separate and linked; conflicting accounts retain both sources", async () => {
  const dir = matter(),
    required = "Calibration must be completed by 8 May 2025.",
    done = "The technician completed calibration on 9 May 2025.",
    disagree = "The manager says calibration was completed on 10 May 2025."
  source(dir, "Requirement.md", required)
  source(dir, "Completion.md", done)
  source(dir, "Disagreement.md", disagree)
  const deps = fake((input) => {
    if (input.agent === "chronology-match") {
      const candidates = JSON.parse(input.text.split("\nCandidates: ")[1]!) as {
        id: string
        accounts: { status: string }[]
      }[]
      const reported = candidates.find((c) => c.accounts.some((a) => a.status === "reported"))
      return JSON.stringify({
        eventId: reported?.id ?? candidates[0]?.id,
        relation: reported ? "conflict" : "completion",
      })
    }
    const body = sectionText(input.text)
    return complete([
      body.includes(required)
        ? event("Calibration completed", required, { status: "required" })
        : body.includes(done)
          ? event("Calibration completed", done)
          : event("Calibration completed", disagree),
    ])
  })
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, deps)
  await awaitChronologyRun(dir)
  const state = readChronology(dir)
  expect(state.events).toHaveLength(2)
  const obligation = state.events.find((e) => e.accounts[0]?.status === "required")!,
    completion = state.events.find((e) => e.accounts[0]?.status === "reported")!
  expect(completion.accounts[0]?.obligationId).toBe(obligation.id)
  expect(completion.accounts).toHaveLength(2)
  expect(completion.disputed).toBe(true)
})

test("withdrawn source creates an individual reversible removal proposal and preserves the saved event", async () => {
  const dir = matter(),
    quote = "The instrument was collected on 6 May 2025."
  source(dir, "Record.md", quote)
  await startChronologyRun(
    dir,
    { expectedRevision: 0, privilege: "mixed", mode: "initial" },
    fake(() => complete([event("Instrument collected", quote)])),
  )
  await awaitChronologyRun(dir)
  let state = readChronology(dir)
  rmSync(path.join(dir, "Record.md"))
  await startChronologyRun(
    dir,
    { expectedRevision: state.revision, privilege: "mixed", mode: "refresh" },
    fake(() => complete([])),
  )
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(state.events).toHaveLength(1)
  expect(state.events[0]?.hidden).toBe(false)
  expect(state.proposals[0]?.kind).toBe("removal")
  state = decideChronologyProposal(dir, state.revision, state.proposals[0]!.id, "accept")
  expect(state.events[0]?.hidden).toBe(true)
})

test("impossible calendar dates and reversed periods remain failures rather than invented precision", () => {
  for (const start of ["2025-02-31", "2025-19-01", "2025-00-03"])
    expect(() => parseChronologyDate({ original: "unclear date", precision: "exact", start })).toThrow("impossible")
  expect(() =>
    parseChronologyDate({
      original: "period",
      precision: "exact",
      start: "2025-05-08",
      end: "2025-05-06",
      rangeKind: "duration",
    }),
  ).toThrow("on or after")
})

test("source reclassification during extraction rejects the late account before quotation verification", async () => {
  const dir = matter(),
    quote = "The instrument was collected on 6 May 2025.",
    id = source(dir, "Record.md", quote)
  let entered!: () => void,
    reply!: (s: string) => void,
    verifications = 0
  const admitted = new Promise<void>((r) => {
    entered = r
  })
  const deps = fake(async () => {
    entered()
    return new Promise<string>((r) => {
      reply = r
    })
  })
  deps.verify = async (citation) => {
    verifications++
    return { ...citation, verified: true }
  }
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, deps)
  await admitted
  const db = openDb(dir)
  try {
    db.run("UPDATE documents SET role='deliverable' WHERE id=?", [id])
  } finally {
    db.close()
  }
  reply(complete([event("Instrument collected", quote)]))
  await awaitChronologyRun(dir)
  const state = readChronology(dir)
  expect(verifications).toBe(0)
  expect(state.events).toHaveLength(0)
  expect(state.coverage[0]?.completedKeys).toHaveLength(0)
  expect(state.coverage[0]?.issues.join(" ")).toContain("no longer authorised")
})

test("source reclassification during matching cannot save previously verified cached evidence", async () => {
  const dir = matter(),
    first = "The instrument was collected on 6 May 2025.",
    second = "The instrument was collected by the courier on 6 May 2025."
  source(dir, "One.md", first)
  const sourceId = source(dir, "Two.md", second)
  let entered!: () => void,
    reply!: (s: string) => void,
    matchPrompt = ""
  const admitted = new Promise<void>((r) => {
    entered = r
  })
  const deps = fake(async (input) => {
    if (input.agent === "chronology-match") {
      matchPrompt = input.text
      entered()
      return new Promise<string>((r) => {
        reply = r
      })
    }
    const body = sectionText(input.text)
    return complete([event("Instrument collected", body.includes(second) ? second : first)])
  })
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, deps)
  await admitted
  const db = openDb(dir)
  try {
    db.run("UPDATE documents SET role='deliverable' WHERE id=?", [sourceId])
  } finally {
    db.close()
  }
  reply(matchFirst(matchPrompt))
  await awaitChronologyRun(dir)
  const state = readChronology(dir)
  expect(state.events).toHaveLength(1)
  expect(state.events[0]?.accounts).toHaveLength(1)
  expect(state.proposals).toHaveLength(0)
  expect(state.coverage.find((c) => c.documentName === "Two.md")?.completedKeys).toHaveLength(0)
})

test("a source-date correction during a model section controls newly returned relative events", async () => {
  const dir = matter(),
    quote = "A response is required within 3 calendar days after this note, dated 2025-05-06."
  source(dir, "Note.md", quote)
  let entered!: () => void, reply!: (s: string) => void
  const admitted = new Promise<void>((r) => {
    entered = r
  })
  const deps = fake(async () => {
    entered()
    return new Promise<string>((r) => {
      reply = r
    })
  })
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, deps)
  await admitted
  actOnChronology(dir, readChronology(dir).revision, {
    type: "source-date",
    documentName: "Note.md",
    date: "2025-05-10",
  })
  reply(
    complete([
      event("Response required", quote, {
        status: "required",
        date: {
          original: "within 3 calendar days after this note",
          precision: "relative",
          relative: {
            expression: "within 3 calendar days after",
            anchorDocument: "Note.md",
            anchorDate: "2025-05-06",
            offsetDays: 3,
            interpretation: "document-stated",
          },
        },
      }),
    ]),
  )
  await awaitChronologyRun(dir)
  const state = readChronology(dir)
  expect(state.events[0]?.date.start).toBe("2025-05-13")
  expect(state.events[0]?.accounts[0]?.date.relative?.anchorDate).toBe("2025-05-10")
  expect(readFileSync(path.join(dir, "Note.md"), "utf8")).toBe(quote)
})

test("refresh reconciles against the current merged record rather than superseded source records", async () => {
  const dir = matter(),
    quote = "The instrument was collected on 6 May 2025."
  source(dir, "Carrier.md", quote)
  await startChronologyRun(
    dir,
    { expectedRevision: 0, privilege: "all-cleared", mode: "initial" },
    fake(() => complete([event("Instrument collected", quote)])),
  )
  await awaitChronologyRun(dir)
  let state = readChronology(dir)
  const previousId = state.events[0]!.id
  state = actOnChronology(dir, state.revision, {
    type: "add",
    title: "Collection context",
    detail: "Lawyer note",
    date: state.events[0]!.date,
  }) as typeof state
  state = actOnChronology(dir, state.revision, {
    type: "merge",
    eventIds: state.events.map((item) => item.id),
    title: "Instrument collected",
    detail: "Collection with lawyer context",
    date: state.events[0]!.date,
  }) as typeof state
  const currentId = state.events.find((item) => !item.supersededBy?.length)!.id
  source(dir, "Depot.md", "The depot confirms that the instrument was collected on 6 May 2025.")
  let candidateIds: string[] = []
  await startChronologyRun(
    dir,
    { expectedRevision: state.revision, privilege: "mixed", mode: "refresh" },
    fake((input) => {
      if (input.agent === "chronology-match") {
        candidateIds = (JSON.parse(input.text.split("\nCandidates: ")[1]!) as { id: string }[]).map((item) => item.id)
        return JSON.stringify({ eventId: currentId, relation: "same" })
      }
      return complete([
        event("Instrument collected", "The depot confirms that the instrument was collected on 6 May 2025."),
      ])
    }),
  )
  await awaitChronologyRun(dir)
  expect(candidateIds).toContain(currentId)
  expect(candidateIds).not.toContain(previousId)
  expect(readChronology(dir).proposals[0]!.targetIds).toEqual([currentId])
})

function nativeChat(directory: string, readWords: () => string, observed: string[] = [], moved = false) {
  const main = { id: "main-chat", location: { directory }, agent: "draft", time: { created: 1, updated: 10 } }
  return (async (input: Parameters<typeof fetch>[0]) => {
    const url = new URL(String(input))
    observed.push(url.pathname + url.search)
    if (url.pathname === "/api/session")
      return Response.json({
        data: [
          { ...main, id: "internal", metadata: { dochaus_role: "chronology" }, agent: "chronology-extract" },
          { ...main, id: "child", parentID: main.id },
          { ...main, id: "lean", agent: "lean-intake" },
          main,
        ],
      })
    if (url.pathname === "/api/session/main-chat")
      return Response.json({
        data: moved ? { ...main, location: { directory: directory + "-different-matter" } } : main,
      })
    if (url.pathname === "/api/session/main-chat/message")
      return Response.json({
        data: [
          {
            id: "assistant",
            type: "assistant",
            time: { created: 5 },
            content: [{ type: "text", text: "Assistant guesses must never become lawyer instructions." }],
          },
          {
            id: "synthetic",
            type: "synthetic",
            time: { created: 4 },
            text: "Internal continuation must never become lawyer instructions.",
          },
          { id: "lawyer", type: "user", time: { created: 3 }, text: readWords() },
        ],
      })
    throw new Error("Unexpected native route: " + url.pathname)
  }) as unknown as typeof fetch
}

test("recent matter instructions contain only real main-chat user admissions from bounded native reads", async () => {
  const dir = matter(),
    observed: string[] = []
  const result = await recentChronologyInstructions(dir, {
    fetch: nativeChat(dir, () => "Resolve the disputed shipment dates before the hearing.", observed),
  })
  expect(result).toEqual({ available: true, instructions: ["Resolve the disputed shipment dates before the hearing."] })
  expect(observed.some((route) => route.includes("parentID=null") && route.includes("limit=32"))).toBe(true)
  expect(observed.some((route) => route.includes("/message?type=user&limit=6"))).toBe(true)
  expect(observed.some((route) => /\/(internal|child|lean)(?:\/|$)/.test(route))).toBe(false)
})

test("main-chat instruction changes enter relevance assessment content keys without importing assistant text", async () => {
  const dir = matter()
  let words = "Resolve disputed shipment dates before the hearing.",
    calls = 0
  const deps = fake((input) => {
    calls++
    expect(input.text).toContain(words)
    expect(input.text).not.toContain("Assistant guesses")
    return JSON.stringify({
      status: words.includes("disputed") ? "relevant" : "irrelevant",
      reason: "Assessment follows the latest recorded lawyer purpose.",
    })
  })
  deps.engine = { fetch: nativeChat(dir, () => words) }
  let state = await assessChronology(dir, deps)
  expect(state.suggestion.status).toBe("relevant")
  const oldKey = state.suggestion.assessedKey
  await assessChronology(dir, deps)
  expect(calls).toBe(1)
  words = "This matter is only a prospective standard form template; there is no historical event question."
  state = await assessChronology(dir, deps)
  expect(calls).toBe(2)
  expect(state.suggestion.assessedKey).not.toBe(oldKey)
  expect(state.suggestion.status).toBe("irrelevant")
})

test("a conversation moved outside the matter is rejected before any globally addressed history read", async () => {
  const dir = matter(),
    observed: string[] = []
  const result = await recentChronologyInstructions(dir, {
    fetch: nativeChat(dir, () => "Foreign matter instructions", observed, true),
  })
  expect(result).toEqual({ available: false, instructions: [] })
  expect(observed.some((route) => route.includes("/message"))).toBe(false)
})

test("source-explicit aliases cite verified passages while weak similarities stay uncertain and original names survive", async () => {
  const dir = matter()
  const quote =
    'Northmere Instrumentation Limited ("Northmere") received the device on 6 May 2025. Alex Carman and A. Carmen attended.'
  source(dir, "Names.md", quote)
  const deps = fake(() =>
    complete([
      event("Device received", quote, {
        people: ["Northmere", "Alex Carman", "A. Carmen"],
        nameAssociations: [
          { name: "Northmere", canonicalName: "Northmere Instrumentation Limited", status: "supported" },
          { name: "A. Carmen", canonicalName: "Alex Carman", status: "supported" },
          { name: "Northmere", canonicalName: "A fabricated organisation", status: "supported" },
        ],
      }),
    ]),
  )
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, deps)
  await awaitChronologyRun(dir)
  const account = readChronology(dir).events[0]!.accounts[0]!
  expect(account.people).toEqual(["Northmere", "Alex Carman", "A. Carmen"])
  expect(account.nameAssociations).toHaveLength(2)
  expect(account.nameAssociations![0]).toMatchObject({
    name: "Northmere",
    canonicalName: "Northmere Instrumentation Limited",
    status: "supported",
    evidenceIds: [account.evidence[0]!.id],
  })
  expect(account.nameAssociations![1]).toMatchObject({
    name: "A. Carmen",
    canonicalName: "Alex Carman",
    status: "uncertain",
  })
})

test("lawyer source dates anchor only their recorded source identity and byte version", () => {
  const makeDate = () =>
    parseChronologyDate({
      original: "3 calendar days after this note",
      precision: "relative",
      relative: {
        expression: "3 calendar days after",
        anchorDocument: "Note.md",
        offsetDays: 3,
        interpretation: "document-stated",
      },
    })
  const current = { documentName: "Note.md", sourceId: "current-source", version: "current-version" }
  const correction = {
    date: "2025-05-10",
    at: 1,
    actor: "local operator",
    sourceId: current.sourceId,
    version: current.version,
  }
  const bound = makeDate()
  resolveRelativeDate(bound, current, "Response is due 3 calendar days after this note.", { "Note.md": correction })
  expect(bound.start).toBe("2025-05-13")
  expect(bound.relative).toMatchObject({ anchorSourceId: current.sourceId, anchorVersion: current.version })
  for (const old of [
    { ...correction, sourceId: "replaced-source" },
    { ...correction, version: "old-version" },
    { date: correction.date, at: 1, actor: "legacy operator" },
  ]) {
    const unresolved = makeDate()
    unresolved.relative!.calculation = "An earlier calculation must not look current"
    unresolved.relative!.anchorSourceId = "old-source"
    unresolved.relative!.anchorVersion = "old-version"
    resolveRelativeDate(unresolved, current, "Response is due 3 calendar days after this note.", { "Note.md": old })
    expect(unresolved.start).toBeUndefined()
    expect(unresolved.relative!.calculation).toBeUndefined()
    expect(unresolved.relative!.anchorSourceId).toBeUndefined()
    expect(unresolved.relative!.anchorVersion).toBeUndefined()
  }
})

test("a changed source does not expose its stale lawyer-supplied date as a current extraction anchor", async () => {
  const dir = matter(),
    before = "A response is required 3 calendar days after this note."
  source(dir, "Note.md", before)
  let state = actOnChronology(dir, 0, {
    type: "source-date",
    documentName: "Note.md",
    date: "2025-05-10",
  }) as ReturnType<typeof readChronology>
  const boundDate = state.sourceDates["Note.md"]!
  const changed = before + " The source version was corrected without a stated source date."
  writeFileSync(path.join(dir, "Note.md"), changed)
  const deps = fake((input) => {
    expect(input.text).toContain('"suppliedSourceDate":null')
    return complete([
      event("Response required", changed, {
        status: "required",
        date: {
          original: "3 calendar days after this note",
          precision: "relative",
          relative: {
            expression: "3 calendar days after",
            anchorDocument: "Note.md",
            offsetDays: 3,
            interpretation: "document-stated",
          },
        },
      }),
    ])
  })
  await startChronologyRun(dir, { expectedRevision: state.revision, privilege: "mixed", mode: "initial" }, deps)
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(state.events[0]!.date.start).toBeUndefined()
  expect(state.sourceDates["Note.md"]).toEqual(boundDate)
})

test("clearly titled supplied chronologies cover case, separators, camelcase, dates and readable formats without admitting derived output", () => {
  const dir = matter()
  const supplied = [
    "2025-06-18_CLAIMANT_CHRONOLOGY.PDF",
    "laterChronologyRecord.docx",
    "Client-chronologies.xlsx",
    "CHRONOLOGY.markdown",
    "Chronology of Events.eml",
    "received_chronology.msg",
    "20250618chronology.md",
  ]
  const ordinary = [
    "correspondence.pdf",
    "ChronologicalIndex.md",
    "Chronology request.md",
    "ChronologyTemplates.docx",
    "ChronologyGuide.md",
  ]
  for (const name of [...supplied, ...ordinary]) source(dir, name, "A supplied document.")
  source(dir, "GeneratedChronology.md", "Derived work product.", undefined, "deliverable")
  const sources = chronologySources(dir)
  expect(sources).toHaveLength(supplied.length + ordinary.length)
  for (const name of supplied) expect(sources.find((s) => s.documentName === name)?.kind).toBe("supplied-chronology")
  for (const name of ordinary) expect(sources.find((s) => s.documentName === name)?.kind).toBe("source")
})

test("completed legacy evidence receives one grouped provenance proposal without replaying extraction or changing saved lawyer work", async () => {
  const dir = matter(),
    quote = "The instrument was collected on 6 May 2025."
  source(dir, "laterChronologyRecord.md", quote)
  source(dir, "Unchanged.md", "The courier returned an unsigned form.")
  const initial = fake((input) =>
    input.agent === "chronology-match"
      ? JSON.stringify({ eventId: null, relation: "independent" })
      : complete([
          event(
            sectionText(input.text).includes(quote) ? "Instrument collected" : "Form returned",
            sectionText(input.text).trim(),
          ),
        ]),
  )
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, initial)
  await awaitChronologyRun(dir)
  let state = readChronology(dir)
  const originalId = state.events.find((e) => e.title === "Instrument collected")!.id
  for (const action of [
    {
      type: "edit",
      eventId: originalId,
      title: "Lawyer collection wording",
      detail: "Corrected account",
      date: state.events.find((e) => e.id === originalId)!.date,
    },
    { type: "note", eventId: originalId, note: "Preserve my collection note" },
    { type: "key", eventId: originalId, key: true },
    { type: "review", eventId: originalId, reviewed: true },
    { type: "hide", eventId: originalId, hidden: true },
  ] as const)
    state = actOnChronology(dir, state.revision, action) as typeof state
  // A persisted pre-fix result had complete text coverage but ordinary-source metadata.
  state = mutateChronology(dir, state.revision, "legacy provenance fixture", (current) => {
    const evidence = current.events.find((e) => e.id === originalId)!.accounts[0]!.evidence[0]!
    evidence.provenance = "source"
    delete evidence.issue
    current.coverage.find((c) => c.documentName === "laterChronologyRecord.md")!.suppliedChronology = false
  })
  const savedBefore = structuredClone(state.events),
    keysBefore = state.coverage.map((c) => c.completedKeys)
  let jobs = 0
  const noReplay = fake(() => {
    jobs++
    throw new Error("Unchanged completed sections must not be replayed")
  })
  await startChronologyRun(dir, { expectedRevision: state.revision, privilege: "mixed", mode: "refresh" }, noReplay)
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(jobs).toBe(0)
  expect(state.events).toEqual(savedBefore)
  expect(state.coverage.map((c) => c.completedKeys)).toEqual(keysBefore)
  expect(state.coverage.find((c) => c.documentName === "laterChronologyRecord.md")!.suppliedChronology).toBe(true)
  expect(state.proposals).toHaveLength(1)
  const proposal = state.proposals[0]!,
    proposed = proposal.events[0]!
  expect(proposal.kind).toBe("amendment")
  expect(proposal.targetIds).toEqual([originalId])
  expect(proposed.accounts[0]!.id).toBe(savedBefore[0]!.accounts[0]!.id)
  expect(proposed.accounts[0]!.evidence[0]).toMatchObject({
    id: savedBefore[0]!.accounts[0]!.evidence[0]!.id,
    provenance: "supplied-chronology",
  })
  await startChronologyRun(
    dir,
    { expectedRevision: state.revision, privilege: "all-cleared", mode: "refresh" },
    noReplay,
  )
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(jobs).toBe(0)
  expect(state.proposals).toHaveLength(1)
  expect(state.proposals[0]!.id).toBe(proposal.id)
  state = decideChronologyProposal(dir, state.revision, proposal.id, "accept")
  const accepted = state.events.find((e) => e.id === originalId)!
  expect(accepted).toMatchObject({
    title: "Lawyer collection wording",
    detail: "Corrected account",
    hidden: true,
    key: true,
    note: "Preserve my collection note",
    reviewed: true,
  })
  expect(accepted.accounts).toHaveLength(1)
  expect(accepted.accounts[0]!.id).toBe(savedBefore[0]!.accounts[0]!.id)
  expect(accepted.accounts[0]!.evidence[0]!.provenance).toBe("supplied-chronology")
  await startChronologyRun(dir, { expectedRevision: state.revision, privilege: "mixed", mode: "refresh" }, noReplay)
  await awaitChronologyRun(dir)
  expect(readChronology(dir).proposals).toHaveLength(1)
  expect(jobs).toBe(0)
})

test("a provenance correction preserves unrelated evidence issues and stale confirmations cannot replace concurrent lawyer edits", async () => {
  const dir = matter(),
    quote = "The instrument was collected on 6 May 2025."
  source(dir, "Claimant chronology.md", quote)
  await startChronologyRun(
    dir,
    { expectedRevision: 0, privilege: "mixed", mode: "initial" },
    fake(() => complete([event("Instrument collected", quote)])),
  )
  await awaitChronologyRun(dir)
  let state = readChronology(dir),
    eventId = state.events[0]!.id
  state = mutateChronology(dir, state.revision, "legacy provenance fixture", (current) => {
    current.events[0]!.accounts[0]!.evidence[0]!.provenance = "source"
    current.events[0]!.accounts[0]!.evidence[0]!.issue = "A separate recorded evidence qualification."
  })
  await startChronologyRun(
    dir,
    { expectedRevision: state.revision, privilege: "mixed", mode: "refresh" },
    fake(() => {
      throw new Error("No replay")
    }),
  )
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(state.proposals[0]!.events[0]!.accounts[0]!.evidence[0]!.issue).toBe(
    "A separate recorded evidence qualification.",
  )
  const proposal = state.proposals[0]!
  state = actOnChronology(dir, state.revision, {
    type: "edit",
    eventId,
    title: "New lawyer wording",
    detail: "Concurrent correction",
    date: state.events[0]!.date,
  }) as typeof state
  expect(() => decideChronologyProposal(dir, state.revision, proposal.id, "accept")).toThrow("current event changed")
  expect(readChronology(dir).events[0]!.title).toBe("New lawyer wording")
})

test("calendar normalization preserves qualifiers and resolves only explicit event or scoped years", () => {
  expect(parseChronologyDate({ original: "by 17 February 2031", precision: "exact" })).toMatchObject({
    start: "2031-02-17",
    precision: "exact",
    qualifier: "by",
  })
  const date = parseChronologyDate({ original: "9 September", precision: "exact" })
  expect(date).toMatchObject({ precision: "ambiguous" })
  expect(date.start).toBeUndefined()
  const source =
    "# Reporting year: 2031\nThe seal was replaced on 9 September.\n# Events in 2032\nThe seal was replaced on 9 September."
  const first = normaliseChronologyDate(
    date,
    source,
    "The seal was replaced on 9 September.",
    source.indexOf("The seal"),
  )
  const last = normaliseChronologyDate(
    date,
    source,
    "The seal was replaced on 9 September.",
    source.lastIndexOf("The seal"),
  )
  expect(first.date.start).toBe("2031-09-09")
  expect(first.context?.quote).toBe("# Reporting year: 2031")
  expect(last.date.start).toBe("2032-09-09")
  expect(last.context?.quote).toBe("# Events in 2032")
  expect(
    normaliseChronologyDate(
      parseChronologyDate({ original: "12 August", precision: "ambiguous" }),
      "",
      "The board met on 12 August 2036.",
    ).date.start,
  ).toBe("2036-08-12")
})

test("document dates, mixed-year accounts, rollover and numeric ambiguity cannot invent a calendar year", () => {
  const date = parseChronologyDate({ original: "9 January", precision: "exact", start: "2031-01-09" })
  for (const [source, passage] of [
    ["Date: 4 October 2031\nThe inspection happened on 9 January.", "The inspection happened on 9 January."],
    ["# Year: 2031\nThe replacement is due on 9 January next year.", "The replacement is due on 9 January next year."],
    [
      "The earlier board met on 9 January 2030; the later board met on 9 January 2031.",
      "The earlier board met on 9 January 2030; the later board met on 9 January 2031.",
    ],
  ])
    expect(normaliseChronologyDate(date, source!, passage!, source!.indexOf(passage!)).date.start).toBeUndefined()
  expect(parseChronologyDate({ original: "04/05/31", precision: "exact", start: "2031-04-05" })).toMatchObject({
    precision: "ambiguous",
  })
  expect(parseChronologyDate({ original: "04/05/31", precision: "exact", start: "2031-04-05" }).start).toBeUndefined()
  expect(
    parseChronologyDate({ original: "Following the inspection", precision: "exact", start: "2031-01-01" }).start,
  ).toBeUndefined()
  expect(parseChronologyDate({ original: "17/02/2031", precision: "exact" }).start).toBe("2031-02-17")
  expect(parseChronologyDate({ original: "12–14 August 2036", precision: "exact" })).toMatchObject({
    precision: "ambiguous",
  })
  expect(parseChronologyDate({ original: "12–14 August 2036", precision: "exact" }).start).toBeUndefined()
  expect(parseChronologyDate({ original: "Undated", precision: "undated" })).toEqual({
    original: "Undated",
    precision: "undated",
  })
})

test("static scope and evidential qualifications do not become events while genuine undated occurrences remain", () => {
  const staticQuote = "The facility contains a heat exchanger and two pressure vessels."
  const qualification = "This note does not establish the cause of the interruption."
  const actual = "The overnight guard reported that someone opened a valve; no date was recorded."
  const parsed = parseChronologyExtraction(
    JSON.stringify({
      complete: true,
      events: [
        event("Equipment configuration", staticQuote, {
          entryKind: "context",
          date: { original: "Undated", precision: "undated" },
        }),
        event("Limits of note", qualification, {
          entryKind: "qualification",
          date: { original: "Undated", precision: "undated" },
        }),
        event("Valve opened", actual, { entryKind: "occurrence", date: { original: "Undated", precision: "undated" } }),
      ],
      gaps: [],
    }),
  )
  expect(parsed.events).toHaveLength(1)
  expect(parsed.events[0]!.title).toBe("Valve opened")
  expect(parsed.events[0]!.date.precision).toBe("undated")
  expect(parsed.nonEvents.map((item) => item.quote)).toEqual([staticQuote, qualification])
})

test("old completed contract reprocesses only affected sections and stages a display-date correction accepted through the real store", async () => {
  const dir = matter(),
    quote = "The technician replaced the pressure seal on 9 September."
  source(dir, "Inspection.md", "# Events in 2034\n" + quote)
  source(dir, "Other.md", "The overnight guard reported that someone opened a valve; no date was recorded.")
  const undatedQuote = "The overnight guard reported that someone opened a valve; no date was recorded."
  const model = fake((input) =>
    input.agent === "chronology-match"
      ? matchFirst(input.text)
      : complete([
          sectionText(input.text).includes(quote)
            ? event("Pressure seal replaced", quote, { date: { original: "9 September", precision: "exact" } })
            : event("Valve opened", undatedQuote, { date: { original: "Undated", precision: "undated" } }),
        ]),
  )
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, model)
  await awaitChronologyRun(dir)
  let state = readChronology(dir)
  const target = state.events.find((e) => e.title === "Pressure seal replaced")!,
    otherKey = state.coverage.find((c) => c.documentName === "Other.md")!.completedKeys[0]!
  state = mutateChronology(dir, state.revision, "legacy completed extraction fixture", (current) => {
    const old = current.events.find((e) => e.id === target.id)!
    old.date = { original: "9 September", precision: "ambiguous" }
    old.accounts[0]!.date = structuredClone(old.date)
    old.accounts[0]!.id = "legacy-ambiguous-account"
    old.accounts[0]!.evidence = old.accounts[0]!.evidence.slice(0, 1)
    old.note = "Lawyer note survives"
    old.key = true
    old.reviewed = true
    current.coverage.find((c) => c.documentName === "Inspection.md")!.completedKeys = [
      "previous-contract-completed-section",
    ]
  })
  const before = structuredClone(state.events)
  let extractionCalls = 0
  const refresh = fake((input) => {
    if (input.agent === "chronology-match") return matchFirst(input.text)
    extractionCalls++
    expect(input.text).toContain("Inspection.md")
    return complete([event("Pressure seal replaced", quote, { date: { original: "9 September", precision: "exact" } })])
  })
  await startChronologyRun(dir, { expectedRevision: state.revision, privilege: "mixed", mode: "refresh" }, refresh)
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(extractionCalls).toBe(1)
  expect(state.events).toEqual(before)
  expect(state.coverage.find((c) => c.documentName === "Other.md")!.completedKeys).toEqual([otherKey])
  const proposal = state.proposals.find((p) => p.kind === "amendment" && p.targetIds.includes(target.id))!
  expect(proposal.events[0]!.date.start).toBe("2034-09-09")
  expect(proposal.events[0]!.accounts.at(-1)!.evidence.some((e) => e.excerpt === "# Events in 2034")).toBe(true)
  state = decideChronologyProposal(dir, state.revision, proposal.id, "accept")
  const accepted = state.events.find((e) => e.id === target.id)!
  expect(accepted).toMatchObject({
    date: { start: "2034-09-09", precision: "exact" },
    note: "Lawyer note survives",
    key: true,
    reviewed: true,
  })
  expect(accepted.accounts.some((a) => a.id === "legacy-ambiguous-account" && a.date.precision === "ambiguous")).toBe(
    true,
  )
  let replayCalls = 0
  await startChronologyRun(
    dir,
    { expectedRevision: state.revision, privilege: "mixed", mode: "refresh" },
    fake(() => {
      replayCalls++
      throw new Error("Current complete section must be reused")
    }),
  )
  await awaitChronologyRun(dir)
  expect(replayCalls).toBe(0)
})

test("refresh proposes reversible static cleanup without hiding genuine undated background or saved work", async () => {
  const dir = matter(),
    staticQuote = "The facility contains a heat exchanger and two pressure vessels.",
    actual = "The overnight guard reported that someone opened a valve; no date was recorded."
  source(dir, "Operations.md", staticQuote + "\n" + actual)
  const oldModel = fake((input) =>
    input.agent === "chronology-match"
      ? JSON.stringify({ eventId: null, relation: "independent" })
      : complete([
          event("Equipment configuration", staticQuote, { date: { original: "Undated", precision: "undated" } }),
          event("Valve opened", actual, { date: { original: "Undated", precision: "undated" } }),
        ]),
  )
  await startChronologyRun(dir, { expectedRevision: 0, privilege: "mixed", mode: "initial" }, oldModel)
  await awaitChronologyRun(dir)
  let state = readChronology(dir)
  state = mutateChronology(dir, state.revision, "legacy extraction contract fixture", (current) => {
    current.coverage[0]!.completedKeys = ["old-contract-section"]
  })
  const before = structuredClone(state.events)
  const corrected = fake((input) =>
    input.agent === "chronology-match"
      ? JSON.stringify({ eventId: null, relation: "independent" })
      : JSON.stringify({
          complete: true,
          events: [event("Valve opened", actual, { date: { original: "Undated", precision: "undated" } })],
          nonEvents: [
            {
              quote: staticQuote,
              reason: "This describes the equipment configuration, rather than an act that occurred.",
            },
          ],
          gaps: [],
        }),
  )
  await startChronologyRun(dir, { expectedRevision: state.revision, privilege: "mixed", mode: "refresh" }, corrected)
  await awaitChronologyRun(dir)
  state = readChronology(dir)
  expect(state.events).toEqual(before)
  expect(state.proposals).toHaveLength(1)
  const removal = state.proposals[0]!,
    staticId = before.find((e) => e.title === "Equipment configuration")!.id
  expect(removal).toMatchObject({ kind: "removal", targetIds: [staticId] })
  state = decideChronologyProposal(dir, state.revision, removal.id, "accept")
  expect(state.events.find((e) => e.id === staticId)!.hidden).toBe(true)
  expect(state.events.find((e) => e.title === "Valve opened")).toMatchObject({
    hidden: false,
    date: { precision: "undated" },
  })
})
