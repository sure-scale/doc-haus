import { expect, test } from "bun:test"
import { cellKey, hashQuestion, type Grid, type GridCellData } from "../src/api/ingest/grid"
import { matterClient } from "../src/api/opencode/client"
import { compareColumn, comparePrompt, comparisonInputs, readComparison, type ComparisonInput } from "../src/api/review-compare"

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
})

test("needs two cited documents, caps at twelve and bounds each document's text", () => {
  expect(comparisonInputs(grid(), "law", rows.slice(0, 1))).toBeUndefined()
  expect(comparisonInputs(grid(), "missing", rows)).toBeUndefined()
  const many = Array.from({ length: 13 }, (_, i) => ({ name: `Doc${i}.docx`, docPath: `/matter/Doc${i}.docx` }))
  const wide: Grid = { columns: [{ id: "term", question: term }], cells: Object.fromEntries(many.map(r => [cellKey(r.name, "term"), cited(r.name, "x ".repeat(500))])) }
  expect(comparisonInputs(wide, "term", many)).toBe("too_many")
  const twelve = comparisonInputs(wide, "term", many.slice(0, 12)) as ComparisonInput
  expect(twelve.documents).toHaveLength(12)
  for (const doc of twelve.documents) expect(doc.answer.length + doc.excerpt.length).toBeLessThanOrEqual(600)
})

const docs = (...names: string[]) => ({ documents: names.map(name => ({ name, answer: "", section: "", excerpt: "" })) })
const reply = (groups: Array<[string[], string]>, summary = "Most agree.") => JSON.stringify({ groups: groups.map(([documents, reason]) => ({ documents, reason })), summary })

test("labels: a unique largest group is consistent, one group is all consistent, a tie all differ", () => {
  const four = docs("A", "B", "C", "D")
  expect(readComparison(reply([[["A", "B", "C"], "Three years"], [["D"], "Two years"]]), four)).toEqual({
    summary: "Most agree.",
    cells: { A: { label: "consistent", reason: "Three years" }, B: { label: "consistent", reason: "Three years" }, C: { label: "consistent", reason: "Three years" }, D: { label: "differs", reason: "Two years" } },
  })
  expect(Object.values(readComparison(reply([[["A", "B", "C", "D"], "Three years"]]), four)!.cells).map(c => c.label)).toEqual(["consistent", "consistent", "consistent", "consistent"])
  expect(Object.values(readComparison(reply([[["A", "B"], "Three years"], [["C", "D"], "Two years"]]), four)!.cells).map(c => c.label)).toEqual(["differs", "differs", "differs", "differs"])
  expect(Object.values(readComparison(reply([[["A"], "One"], [["B"], "Two"], [["C"], "Three"], [["D"], "Four"]]), four)!.cells).every(c => c.label === "differs")).toBe(true)
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
    reply([[["A", "B"], "x".repeat(141)], [["C"], "Other"]]),
    reply([[["A", "B"], ""], [["C"], "Other"]]),
    reply([[["A", "B", "C"], "Same"]], "x".repeat(201)),
    reply([[["A", "B", "C"], "Same"]], 'All say "three years"'),
    reply([[["A", "B", "C"], "Same"]], "All say «three years»"),
    "The documents mostly agree.",
  ]) expect(readComparison(text, three)).toBeUndefined()
})

test("a reason that opens with its own documents' names loses them; the UI shows each name beside its reason", () => {
  const five = docs("Alpha.docx", "Beta.docx", "Gamma lease.pdf", "Delta.docx", "Epsilon.docx")
  const cells = readComparison(reply([
    [["Alpha.docx", "Beta.docx"], "Alpha.docx and Beta.docx: three years"],
    [["Gamma lease.pdf"], "Gamma lease — two years"],
    [["Delta.docx"], "Delta.docx has a one-year term"],
    [["Epsilon.docx"], "Alpha.docx term with Epsilon as agent"],
  ]), five)!.cells
  expect(Object.fromEntries(Object.entries(cells).map(([name, cell]) => [name, cell.reason]))).toEqual({
    "Alpha.docx": "Three years", "Beta.docx": "Three years", "Gamma lease.pdf": "Two years",
    "Delta.docx": "Has a one-year term", "Epsilon.docx": "Alpha.docx term with Epsilon as agent",
  })
  const two = docs("Alpha.docx", "Beta.docx")
  expect(readComparison(reply([[["Alpha.docx"], "Alphabetical schedule"], [["Beta.docx"], "Other"]]), two)!.cells["Alpha.docx"]!.reason).toBe("Alphabetical schedule")
  expect(readComparison(reply([[["Alpha.docx"], "Alpha.docx"], [["Beta.docx"], "Other"]]), two)).toBeUndefined()
  expect(readComparison(reply([[["Alpha.docx"], `Alpha.docx: ${"x".repeat(138)}`], [["Beta.docx"], "Other"]]), two)).toBeDefined()
})

async function engine(onWait: (request: Request) => Promise<Response>) {
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
      { type: "assistant", content: [{ type: "text", text: reply([[["Alpha.docx", "Beta.docx"], "Three years"], [["Zeta.docx"], "Two years"]]) }] },
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
