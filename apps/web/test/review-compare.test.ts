import { expect, test } from "bun:test"
import { cellKey, hashQuestion, type Grid, type GridCellData } from "../src/api/ingest/grid"
import { matterClient } from "../src/api/opencode/client"
import { compareColumn, comparePrompt, comparisonInputs, MAX_COMPARE_INPUT_BYTES, readComparison, type ComparisonInput } from "../src/api/review-compare"

const term = "What is the initial term?"
const law = "Which law governs?"
const rows = ["Alpha.docx", "Beta.docx", "Gamma.docx", "Delta.docx", "Epsilon.docx", "Zeta.docx"].map(name => ({ name, docPath: `/matter/${name}` }))
const cited = (doc: string, answer: string, question = term, status: GridCellData["status"] = "filled"): GridCellData => ({
  answer, status, questionHash: hashQuestion(question),
  citation: { documentName: doc, docPath: `/matter/${doc}`, section: "2", excerpt: `${doc} clause 2 text.`, charStart: 10, charEnd: 30 },
})
const grid = (): Grid => ({
  columns: [{ id: "term", question: term }, { id: "law", question: law }],
  cells: {
    [cellKey("Alpha.docx", "term")]: cited("Alpha.docx", "Three years"),
    [cellKey("Beta.docx", "term")]: cited("Beta.docx", "Three years", term, "reviewed"),
    [cellKey("Gamma.docx", "term")]: cited("Beta.docx", "Five years"),
    [cellKey("Delta.docx", "term")]: { answer: "Not addressed", status: "filled", questionHash: hashQuestion(term) },
    [cellKey("Epsilon.docx", "term")]: cited("Epsilon.docx", "One year", "An earlier wording"),
    [cellKey("Zeta.docx", "term")]: cited("Zeta.docx", "Two years"),
    [cellKey("Alpha.docx", "law")]: cited("Alpha.docx", "Law of Ruritania", law),
    [cellKey("Beta.docx", "law")]: cited("Beta.docx", "Law of Elbonia", law),
  },
})
const inputFor = (value: Grid, column = "term") => comparisonInputs(value, column, rows) as ComparisonInput

test("compares only this column's current answers cited to their own row", () => {
  const input = inputFor(grid())
  expect(input.documents.map(d => d.name)).toEqual(["Alpha.docx", "Beta.docx", "Zeta.docx"])
  const prompt = comparePrompt(input)
  for (const text of ["Three years", "Two years", term]) expect(prompt).toContain(text)
  for (const text of ["Five years", "Gamma.docx", "Delta.docx", "One year", "Ruritania", law]) expect(prompt).not.toContain(text)
})

test("a column marked not to compare yields nothing to compare; other columns still compare", () => {
  const flagged = grid()
  flagged.columns[0] = { ...flagged.columns[0]!, compare: false }
  expect(comparisonInputs(flagged, "term", rows)).toBeUndefined()
  expect(inputFor(flagged, "law").documents).toHaveLength(2)
})

test("the signature follows this column's answers and citations only", () => {
  const before = inputFor(grid()).signature
  const otherColumn = grid()
  otherColumn.cells[cellKey("Alpha.docx", "law")]!.answer = "Law of Freedonia"
  expect(inputFor(otherColumn).signature).toBe(before)
  const answer = grid()
  answer.cells[cellKey("Zeta.docx", "term")]!.answer = "Four years"
  expect(inputFor(answer).signature).not.toBe(before)
  const moved = grid()
  moved.cells[cellKey("Zeta.docx", "term")]!.citation!.charStart = 40
  expect(inputFor(moved).signature).not.toBe(before)
  const qualified = grid()
  qualified.cells[cellKey("Zeta.docx", "term")]!.citation!.excerpt += " Except where the customer terminates earlier."
  expect(inputFor(qualified).signature).not.toBe(before)
  // A new version can move its document's row; the answers are unchanged.
  expect((comparisonInputs(grid(), "term", rows.toReversed()) as ComparisonInput).signature).toBe(before)
})

test("needs two cited documents, caps at twelve and preserves complete evidence", () => {
  expect(comparisonInputs(grid(), "law", rows.slice(0, 1))).toBeUndefined()
  expect(comparisonInputs(grid(), "missing", rows)).toBeUndefined()
  const many = Array.from({ length: 13 }, (_, i) => ({ name: `Doc${i}.docx`, docPath: `/matter/Doc${i}.docx` }))
  const wide: Grid = { columns: [{ id: "term", question: term }], cells: Object.fromEntries(many.map(r => [cellKey(r.name, "term"), cited(r.name, "x ".repeat(500))])) }
  expect(comparisonInputs(wide, "term", many)).toBe("too_many")
  const twelve = comparisonInputs(wide, "term", many.slice(0, 12)) as ComparisonInput
  expect(twelve.documents).toHaveLength(12)
  for (const doc of twelve.documents) expect(doc.answer).toBe("x ".repeat(500))
})

test("late qualifications in both answers and sources reach the comparison unchanged", () => {
  const value = grid()
  const prefix = "The supplier must deliver the contracted services. ".repeat(16)
  const exceptions = ["Except during a declared emergency.", "Including during a declared emergency."]
  for (const [index, name] of ["Alpha.docx", "Beta.docx"].entries()) {
    const cell = value.cells[cellKey(name, "term")]!
    cell.answer = `${prefix}\n\n${exceptions[index]}`
    cell.citation!.excerpt = `${prefix}\n${exceptions[index]}`
  }
  const input = inputFor(value)
  for (const doc of input.documents) {
    const cell = value.cells[cellKey(doc.name, "term")]!
    expect(doc.answer).toBe(cell.answer)
    expect(doc.excerpt).toBe(cell.citation!.excerpt)
    expect(comparePrompt(input)).toContain(JSON.stringify(doc.answer))
    expect(comparePrompt(input)).toContain(JSON.stringify(doc.excerpt))
  }
})

test("oversized complete evidence is not compared; the byte budget includes question and framing", () => {
  const value = grid()
  const bytes = new TextEncoder().encode(comparePrompt(inputFor(value))).byteLength
  value.cells[cellKey("Alpha.docx", "term")]!.answer += "x".repeat(MAX_COMPARE_INPUT_BYTES - bytes)
  const atLimit = comparisonInputs(value, "term", rows)
  expect(typeof atLimit).toBe("object")
  expect(new TextEncoder().encode(comparePrompt(atLimit as ComparisonInput)).byteLength).toBe(MAX_COMPARE_INPUT_BYTES)
  value.cells[cellKey("Alpha.docx", "term")]!.citation!.excerpt += "é"
  expect(comparisonInputs(value, "term", rows)).toBe("too_large")

  const longQuestion = grid()
  longQuestion.columns[0]!.question = "q".repeat(MAX_COMPARE_INPUT_BYTES)
  for (const cell of Object.values(longQuestion.cells)) cell.questionHash = hashQuestion(longQuestion.columns[0]!.question)
  expect(comparisonInputs(longQuestion, "term", rows)).toBe("too_large")
})

const docs = (...names: string[]) => ({ documents: names.map(name => ({ name, answer: "", section: "", excerpt: "" })) })
const reply = (groups: Array<[string[], string]>, points: string[] = []) =>
  JSON.stringify({ groups: groups.map(([documents, reason], index) => ({ documents, reason, point: points[index] ?? "the same term" })) })

test("labels: a unique largest group is consistent, one group is all consistent, a tie all differ", () => {
  const four = docs("A", "B", "C", "D")
  expect(readComparison(reply([[["A", "B", "C"], "Three years"], [["D"], "Two years"]], ["a three-year term", "a two-year term"]), four)).toEqual({
    summary: "3 of 4 answers state a three-year term; 1 states a two-year term.",
    cells: { A: { label: "consistent", reason: "Three years" }, B: { label: "consistent", reason: "Three years" }, C: { label: "consistent", reason: "Three years" }, D: { label: "differs", reason: "Two years" } },
  })
  expect(Object.values(readComparison(reply([[["A", "B", "C", "D"], "Three years"]]), four)!.cells).map(c => c.label)).toEqual(["consistent", "consistent", "consistent", "consistent"])
  expect(Object.values(readComparison(reply([[["A", "B"], "Three years"], [["C", "D"], "Two years"]]), four)!.cells).map(c => c.label)).toEqual(["differs", "differs", "differs", "differs"])
  expect(Object.values(readComparison(reply([[["A"], "One"], [["B"], "Two"], [["C"], "Three"], [["D"], "Four"]]), four)!.cells).every(c => c.label === "differs")).toBe(true)
})

test("the summary is built from group points, so its counts cannot be wrong", () => {
  const four = docs("A", "B", "C", "D")
  expect(readComparison(reply([[["A", "B", "C", "D"], "Three years"]], ["a three-year term"]), four)!.summary)
    .toBe("All 4 answers state a three-year term.")
  expect(readComparison(reply([[["A", "B"], "Three years"], [["C"], "Two years"], [["D"], "One year"]], ["a three-year term", "a two-year term", "a one-year term"]), four)!.summary)
    .toBe("2 of 4 answers state a three-year term; 1 states a two-year term; 1 states a one-year term.")
  expect(readComparison(reply([[["A"], "One"], [["B"], "Two"], [["C"], "Three"]], ["a one-year term", "a two-year term", "a three-year term"]), docs("A", "B", "C"))!.summary)
    .toBe("1 of 3 answers states a one-year term; 1 states a two-year term; 1 states a three-year term.")
})

test("a reason up to 240 characters is kept; a longer one is rejected", () => {
  const three = docs("A", "B", "C")
  const reason = "x".repeat(240)
  expect(readComparison(reply([[["A", "B", "C"], reason]]), three)!.cells.A!.reason).toBe(reason)
  expect(readComparison(reply([[["A", "B", "C"], `${reason}x`]]), three)).toBeUndefined()
})

test("rejects any reply that misplaces a document, quotes, or overruns its limits", () => {
  const three = docs("A", "B", "C")
  expect(readComparison("```json\n" + reply([[["A", "B", "C"], "Same"]]) + "\n```", three)).toBeDefined()
  for (const text of [
    reply([[["A", "B"], "Same"]]),
    reply([[["A", "B"], "Same"], [["B", "C"], "Other"]]),
    reply([[["A", "B", "C"], "Same"], [["D"], "Other"]]),
    reply([[["A", "B"], 'Says "three years"'], [["C"], "Other"]]),
    reply([[["A", "B"], "Says “three years”"], [["C"], "Other"]]),
    reply([[["A", "B"], "x".repeat(241)], [["C"], "Other"]]),
    reply([[["A", "B"], ""], [["C"], "Other"]]),
    reply([[["A", "B"], "Same"], [["C"], "Other"]], ["x".repeat(100), "y".repeat(100)]),
    reply([[["A", "B", "C"], "Same"]], ['All say "three years"']),
    reply([[["A", "B", "C"], "Same"]], ["All say «three years»"]),
    reply([[["A", "B", "C"], "Same"]], ["x".repeat(101)]),
    JSON.stringify({ groups: [{ documents: ["A", "B", "C"], reason: "Same" }] }),
    "The documents mostly agree.",
  ]) expect(readComparison(text, three)).toBeUndefined()
})

test("comparison reasons retain their wording even when a filename resembles a legal qualification", () => {
  const five = docs("Alpha.docx", "Beta.docx", "Gamma lease.pdf", "Delta.docx", "Epsilon.docx")
  const cells = readComparison(reply([
    [["Alpha.docx", "Beta.docx"], "Alpha.docx and Beta.docx: three years"],
    [["Gamma lease.pdf"], "Gamma lease — two years"],
    [["Delta.docx"], "Delta.docx has a one-year term"],
    [["Epsilon.docx"], "Alpha.docx term with Epsilon as agent"],
  ]), five)!.cells
  expect(Object.fromEntries(Object.entries(cells).map(([name, cell]) => [name, cell.reason]))).toEqual({
    "Alpha.docx": "Alpha.docx and Beta.docx: three years", "Beta.docx": "Alpha.docx and Beta.docx: three years", "Gamma lease.pdf": "Gamma lease — two years",
    "Delta.docx": "Delta.docx has a one-year term", "Epsilon.docx": "Alpha.docx term with Epsilon as agent",
  })
  const two = docs("Alpha.docx", "Beta.docx")
  expect(readComparison(reply([[["Alpha.docx"], "Alphabetical schedule"], [["Beta.docx"], "Other"]]), two)!.cells["Alpha.docx"]!.reason).toBe("Alphabetical schedule")
  expect(readComparison(reply([[["Alpha.docx"], `Alpha.docx: ${"x".repeat(229)}`], [["Beta.docx"], "Other"]]), two)).toBeUndefined()
  const qualified = readComparison(reply([[["No.docx"], "No liability cap applies."], [["Except.docx"], "Except for fraud, liability is capped."]]), docs("No.docx", "Except.docx"))!
  expect(qualified.cells["No.docx"]!.reason).toBe("No liability cap applies.")
  expect(qualified.cells["Except.docx"]!.reason).toBe("Except for fraud, liability is capped.")
})

async function engine(onWait: (request: Request) => Promise<Response>, response = reply([[["Alpha.docx", "Beta.docx"], "Three years"], [["Zeta.docx"], "Two years"]])) {
  const original = globalThis.fetch
  const requests: Request[] = []
  let created: Record<string, unknown> | undefined
  globalThis.fetch = (async (source: string | URL | Request, init?: RequestInit) => {
    const request = source instanceof Request ? source : new Request(source, init)
    requests.push(request.clone())
    const pathname = new URL(request.url).pathname
    if (pathname === "/engine/config") return Response.json({ model: "provider/primary", small_model: "provider/fast" })
    if (pathname === "/api/session" && request.method === "POST") {
      created = await request.json() as Record<string, unknown>
      return Response.json({ data: { id: created.id } })
    }
    if (pathname.endsWith("/prompt")) return Response.json({ data: { id: "msg_fixture" } })
    if (pathname.endsWith("/wait")) return onWait(request)
    if (pathname.endsWith("/message")) return Response.json({ data: [
      { type: "assistant", content: [{ type: "text", text: response }] },
      { type: "idle", outcome: "succeeded" },
    ], cursor: {} })
    if (pathname.endsWith("/interrupt")) return Response.json({ interrupted: true })
    if (request.method === "DELETE") return new Response(null, { status: 204 })
    throw new Error(`Unexpected fixture request: ${request.method} ${pathname}`)
  }) as typeof fetch
  return { requests, created: () => created, restore: () => { globalThis.fetch = original } }
}

test("runs one hidden review-compare session on the fast model and deletes it", async () => {
  const fake = await engine(async () => new Response(null, { status: 204 }))
  try {
    const result = await compareColumn(matterClient("/synthetic/matter"), inputFor(grid()), new AbortController().signal)
    expect(result).toMatchObject({ questionHash: hashQuestion(term), inputs: inputFor(grid()).signature, cells: { "Zeta.docx": { label: "differs" } } })
    expect(fake.created()).toMatchObject({ agent: "review-compare", model: { providerID: "provider", id: "fast" }, metadata: { dochaus_role: "review-compare" }, title: `doc.haus review compare: ${term}` })
    expect(fake.requests.some(r => r.method === "DELETE" && r.url.includes(String(fake.created()?.id)))).toBe(true)
    expect(fake.requests.some(r => r.url.endsWith("/interrupt"))).toBe(false)
  } finally { fake.restore() }
})

test("current cited inputs carry comparability instructions and an unavailable reply through the real producer", async () => {
  const value = grid()
  value.cells[cellKey("Beta.docx", "term")]!.answer = "The supplied excerpt does not establish the initial term."
  value.cells[cellKey("Alpha.docx", "term")]!.citation!.excerpt = "The initial term is three years."
  value.cells[cellKey("Beta.docx", "term")]!.citation!.excerpt = "The agreement takes effect on signature."
  const input = comparisonInputs(value, "term", rows.slice(0, 2)) as ComparisonInput
  const reason = "The initial term cannot be established for one document from its supplied evidence."
  const fake = await engine(async () => new Response(null, { status: 204 }), JSON.stringify({ status: "unavailable", reason }))
  try {
    const result = await compareColumn(matterClient("/synthetic/matter"), input, new AbortController().signal)
    const promptRequest = fake.requests.find(r => r.url.endsWith("/prompt"))!
    const body = await promptRequest.json() as { text: string; delivery: string }
    expect(body).toEqual({ text: comparePrompt(input), delivery: "queue" })
    const evidenceStart = body.text.indexOf(JSON.stringify(input.documents))
    expect(body.text.indexOf("Determine comparability before forming groups.")).toBeLessThan(evidenceStart)
    expect(body.text).toContain("Silence in the supplied evidence or a statement that the answer is unavailable does not establish that a term or obligation is absent.")
    expect(body.text).toContain('Unavailable result: {"status":"unavailable","reason":')
    expect(input.documents.map(d => [d.answer, d.excerpt])).toEqual(rows.slice(0, 2).map(r => {
      const cell = value.cells[cellKey(r.name, "term")]!
      return [cell.answer, cell.citation!.excerpt]
    }))
    expect(result).toMatchObject({ questionHash: hashQuestion(term), inputs: input.signature, status: "unavailable", summary: reason, cells: {} })
    expect(fake.requests.some(r => r.method === "DELETE" && r.url.includes(String(fake.created()?.id)))).toBe(true)
    expect(fake.requests.some(r => r.url.endsWith("/interrupt"))).toBe(false)
  } finally { fake.restore() }
})

test("an explicit supported absence remains comparable through the actual producer and admission", async () => {
  const question = "Is customer consent required for an assignment?"
  const value: Grid = { columns: [{ id: "consent", question }], cells: {} }
  const positions = ["Customer consent is required for assignment.", "No customer consent is required for assignment."]
  for (const [index, row] of rows.slice(0, 2).entries()) {
    const cell = cited(row.name, positions[index]!, question)
    cell.citation!.excerpt = positions[index]!
    value.cells[cellKey(row.name, "consent")] = cell
  }
  const input = comparisonInputs(value, "consent", rows.slice(0, 2)) as ComparisonInput
  const fake = await engine(async () => new Response(null, { status: 204 }), reply([
    [["Alpha.docx"], positions[0]!], [["Beta.docx"], positions[1]!],
  ], ["a requirement for customer consent", "no requirement for customer consent"]))
  try {
    const result = await compareColumn(matterClient("/synthetic/matter"), input, new AbortController().signal)
    const body = await fake.requests.find(r => r.url.endsWith("/prompt"))!.json() as { text: string }
    expect(body.text).toContain("An explicit supported absence of a requirement is a substantive answer and remains comparable.")
    expect(body.text).toContain(JSON.stringify(input.documents))
    expect(result?.status).toBeUndefined()
    expect(result?.cells).toEqual({
      "Alpha.docx": { label: "differs", reason: positions[0] },
      "Beta.docx": { label: "differs", reason: positions[1] },
    })
    expect(result).toMatchObject({ questionHash: hashQuestion(question), inputs: input.signature })
    expect(fake.requests.some(r => r.method === "DELETE")).toBe(true)
  } finally { fake.restore() }
})

test("an abort while waiting interrupts and still deletes the session", async () => {
  const controller = new AbortController()
  const fake = await engine(request => new Promise((_, reject) => {
    request.signal.addEventListener("abort", () => reject(request.signal.reason))
    controller.abort()
  }))
  try {
    await expect(compareColumn(matterClient("/synthetic/matter"), inputFor(grid()), controller.signal)).rejects.toThrow()
    const id = String(fake.created()?.id)
    expect(fake.requests.some(r => r.url.endsWith("/interrupt") && r.url.includes(id))).toBe(true)
    expect(fake.requests.some(r => r.method === "DELETE" && r.url.includes(id))).toBe(true)
  } finally { fake.restore() }
})


test("insufficient evidence is an explicit unlabelled result, never a difference", () => {
  const result = readComparison(JSON.stringify({ status: "unavailable", reason: "The exception referenced in one answer is not supplied." }), docs("A", "B"))
  expect(result).toEqual({ status: "unavailable", summary: "The exception referenced in one answer is not supplied.", cells: {} })
  expect(readComparison(JSON.stringify({ status: "unavailable", reason: "Missing evidence.", groups: [{ documents: ["A"], reason: "Unknown" }] }), docs("A", "B"))).toBeUndefined()
  expect(readComparison(JSON.stringify({ status: "unknown", reason: "Missing evidence." }), docs("A", "B"))).toBeUndefined()
})
