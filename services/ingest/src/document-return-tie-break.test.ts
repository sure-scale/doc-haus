import { expect, spyOn, test } from "bun:test"
import { Hono } from "hono"
import { mkdtempSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { documentExchangeRoutes } from "./document-exchange-routes"
import { checkReturnTieBreak } from "./document-return-tie-break"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { captureVersion, readAvailableDocumentVersions, readCurrentVersion, readDocumentVersions, readVersionBytes } from "./document-versions"
import { inspectDocumentReturn, stampDocumentReturnMarker } from "./document-return-matching"
import { recordDocumentReceived, relinkDocumentReceivedDocument } from "./document-exchanges"
import { openDb } from "./db"
import { openTelemetryDb, readGenerations } from "../../../dochaus/lib/telemetry"
import { readComparisonParagraphs } from "../../../dochaus/lib/version-compare"

const docx = (...text: string[]) => Packer.toBuffer(new Document({ sections: [{ children: text.map(item => new Paragraph(item)) }] }))
const incoming = [{ text: "Cedar transfers the Orchard site, except for the reserved access strip." }]
const candidates = [
  { versionId: "older", sha256: "a", current: false, sent: true, paragraphs: [{ scope: "body", text: "Cedar transfers the Orchard site, except for the reserved access strip." }] },
  { versionId: "related", sha256: "b", current: false, sent: true, paragraphs: [{ scope: "body", text: "Cedar transfers the Meadow site, including the access strip." }] },
  { versionId: "current", sha256: "c", current: true, sent: false, paragraphs: incoming.map(row => ({ scope: "body", ...row })) },
]
const supported = { choice: "candidate", versionId: "older", reason: "The transaction and reserved strip support this earlier base; the lawyer must confirm.", evidence: [{ versionId: "older", incomingQuote: incoming[0]!.text, candidateQuote: candidates[0]!.paragraphs[0]!.text }] }
const sourcePart = (prompt: string) => {
  const part = prompt.match(/<source-part encoding="json">([\s\S]*?)<\/source-part>/)
  return part ? JSON.parse(part[1]) as { source: string; part: number; parts: number; passages: { paragraph: number; scope: string; text: string; continuedBefore: boolean; continuedAfter: boolean }[] } : undefined
}
const sourceAnswer = (part: NonNullable<ReturnType<typeof sourcePart>>) => JSON.stringify({ quotes: part.passages.filter(passage => passage.text.trim().length >= 8).slice(0, 4).map(passage => ({ paragraph: passage.paragraph, text: passage.text.slice(0, 1200) })), limitations: [] })
const contrastPart = (prompt: string) => {
  const left = prompt.match(/<left-version encoding="json">([\s\S]*?)<\/left-version>/)
  if (!left) return
  return { left: JSON.parse(left[1]), right: JSON.parse(prompt.match(/<right-version encoding="json">([\s\S]*?)<\/right-version>/)![1]!), incoming: JSON.parse(prompt.match(/<incoming encoding="json">([\s\S]*?)<\/incoming>/)![1]!) }
}
const contrastAnswer = (part: NonNullable<ReturnType<typeof contrastPart>>, support: "left" | "right" | "neither" | "uncertain" = "neither") => JSON.stringify({ support, reason: "Synthetic contrast protocol only; no semantic quality claim.", limitations: [], evidence: { incomingQuote: part.incoming.text, leftQuote: part.left.passage.text, rightQuote: part.right.passage.text } })

test("matching validates closed earlier IDs and actual quotations, leaving uncertainty explicit", () => {
  expect(checkReturnTieBreak(JSON.stringify(supported), incoming, candidates).versionId).toBe("older")
  for (const versionId of ["current", "invented"]) expect(() => checkReturnTieBreak(JSON.stringify({ ...supported, versionId }), incoming, candidates)).toThrow("allowed earlier version")
  expect(() => checkReturnTieBreak(JSON.stringify({ ...supported, evidence: [{ ...supported.evidence[0], candidateQuote: "including the reserved access strip" }] }), incoming, candidates)).toThrow("quotation")
  expect(() => checkReturnTieBreak(JSON.stringify({ choice: "new-document", reason: "Different transaction.", evidence: supported.evidence }), incoming, candidates)).toThrow("every candidate")
  expect(checkReturnTieBreak(JSON.stringify({ choice: "uncertain", reason: "Shared boilerplate and conflicting exceptions cannot establish ancestry.", evidence: [] }), incoming, candidates).choice).toBe("uncertain")
})

test("actual tie-break HTTP uses only explicit ambiguous calls, complete evidence and fast role; it never saves", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-return-judgment-"))
  let writes = 0, calls = 0, response = "", invalidate = false
  const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async () => { writes++; return {} } }))
  const request = (bytes: Uint8Array, route = "tie-break") => {
    const form = new FormData(); form.append("file", new File([new Uint8Array(bytes)], "Agreement returned.docx"))
    return app.request(`/matters/one/documents/returns/${route}`, { method: "POST", body: form })
  }
  const fetch = spyOn(globalThis, "fetch").mockImplementation((async (url: Parameters<typeof globalThis.fetch>[0], init?: Parameters<typeof globalThis.fetch>[1]) => {
    expect(String(url).endsWith("/api/dochaus/generate")).toBe(true)
    const body = JSON.parse(String(init?.body))
    expect(body.modelRole).toBe("fast")
    expect((init?.headers as Record<string, string>)["x-opencode-directory"]).toBe(directory)
    calls++
    let text = response
    const reading = sourcePart(body.prompt)
    const candidate = body.prompt.match(/<candidate encoding="json">([\s\S]*?)<\/candidate>/), contrast = contrastPart(body.prompt)
    if (reading) text = sourceAnswer(reading)
    else if (contrast) text = contrastAnswer(contrast, contrast.incoming.text === contrast.left.passage.text ? "left" : contrast.incoming.text === contrast.right.passage.text ? "right" : "neither")
    else if (candidate) {
      const row = JSON.parse(candidate[1])
      const ownQuote = row.reading.parts.flatMap((part: { quotes: { text: string }[] }) => part.quotes).at(-1).text
      const incomingQuotes = JSON.parse(body.prompt.match(/<incoming encoding="json">([\s\S]*?)<\/incoming>/)[1]).reading.parts.flatMap((part: { quotes: { text: string }[] }) => part.quotes)
      expect(body.prompt).toContain("<incoming encoding=\"json\">")
      expect(body.prompt).toContain("except for the reserved access strip.")
      expect(row).not.toHaveProperty("paragraphs")
      text = JSON.stringify({ versionId: row.versionId, relation: ownQuote.includes("Different transaction") ? "different-document" : "same-document", reason: "Synthetic protocol judgment; semantic quality is not established.", evidence: [{ versionId: row.versionId, incomingQuote: incomingQuotes.at(-1).text, candidateQuote: ownQuote }] })
    } else {
      expect(body.prompt).toContain("<checked-pair-judgments")
      expect(body.prompt).not.toContain("<incoming")
      expect(body.prompt).toContain("Never choose or default to a current version")
      expect(Buffer.byteLength(body.prompt)).toBeLessThanOrEqual(64 * 1024)
      const summary = JSON.parse(body.prompt.match(/<checked-base-contrasts encoding="json">([\s\S]*?)<\/checked-base-contrasts>/)[1])
      expect(summary.representative).toHaveLength(1)
      expect(summary.contrasts.every((row: object) => !("parts" in row))).toBe(true)
      expect(summary.contrasts.every((row: object) => !("evidence" in row))).toBe(true)
    }
    if (invalidate && candidate && JSON.parse(candidate[1]).versionId === JSON.parse(response).versionId) {
      const db = openDb(directory)
      try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [JSON.parse(response).versionId]) } finally { db.close() }
    }
    return Response.json({ data: { text, model: { providerID: "protocol-fixture", id: "configured-fast" }, elapsed_ms: 1, usage: { tokens: { input: 12, output: 8, reasoning: 0, cache: { read: 0, write: 0 } }, cost: 0.001 } } })
  }) as typeof globalThis.fetch)
  try {
    const shared = "Delivery under this agreement requires written notice and transfer of the specified property by the agreed date."
    const first = await docx(shared, incoming[0]!.text, "Completion is Monday.")
    await writeDeliverable(directory, "Agreement.docx", first)
    const older = readCurrentVersion(directory, "Agreement.docx")!.current
    await writeDeliverableRevision(directory, "Agreement.docx", await docx(shared, "Cedar transfers the Orchard site, including the reserved access strip.", "Completion is Tuesday."))
    await writeDeliverable(directory, "Related.docx", await docx(shared, "Different transaction has Meadow property and includes the access strip."))
    const returned = await docx(shared, incoming[0]!.text, "Completion is Wednesday.")
    const source = readFileSync(path.join(directory, "Agreement.docx"))
    const versions = readDocumentVersions(directory, "Agreement.docx").versions.length
    expect((await request(returned, "inspect")).status).toBe(200)
    expect(calls).toBe(0)
    expect((await request(readVersionBytes(directory, older)!)).status).toBe(400)
    expect(calls).toBe(0)
    response = JSON.stringify({ choice: "candidate", versionId: older.id, reason: "The retained exclusion supports this base; confirm it.", evidence: [{ versionId: older.id, incomingQuote: incoming[0]!.text, candidateQuote: incoming[0]!.text }] })
    const suggested = await request(returned)
    expect(suggested.status).toBe(200)
    expect((await suggested.json()).advisory).toMatchObject({ choice: "candidate", versionId: older.id, model: { providerID: "protocol-fixture", id: "configured-fast" } })
    expect(calls).toBe(10) // Four source readings, three identity pairs, two contrasts, reconciliation.
    expect(writes).toBe(0)
    expect(readDocumentVersions(directory, "Agreement.docx").versions).toHaveLength(versions)
    expect(readFileSync(path.join(directory, "Agreement.docx"))).toEqual(source)
    const usage = openTelemetryDb(directory)
    try {
      const rows = readGenerations(usage)
      expect(rows).toHaveLength(10)
      expect(rows.every(row => row.agent === "returned-file-match" && row.model === "protocol-fixture/configured-fast" && row.cost === 0.001)).toBe(true)
    } finally { usage.close() }
    invalidate = true
    const invalidated = await request(returned)
    expect(invalidated.status).toBe(409)
    expect((await invalidated.json()).code).toBe("match_source_changed")
    expect(writes).toBe(0)
  } finally { fetch.mockRestore(); rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("long sources are fully read in bounded parts, incoming reading is reused, and pairs copy only checked quotations", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-return-part-reading-"))
  let fetch: ReturnType<typeof spyOn> | undefined
  try {
    const filler = "UNSELECTED_LONG_SOURCE " + "Ordinary repeated delivery provisions apply. ".repeat(250) + "💬 End of the long paragraph."
    const identity = "Identity: Cedar transfers the Orchard site, except for the reserved access strip."
    await writeDeliverable(directory, "Agreement.docx", await docx(filler, identity))
    const earlier = readCurrentVersion(directory, "Agreement.docx")!.current
    await writeDeliverableRevision(directory, "Agreement.docx", await docx(filler, identity + " Transfer is Monday."))
    const returned = await docx(filler, identity + " Transfer is Tuesday.")
    const sourceCalls: NonNullable<ReturnType<typeof sourcePart>>[] = []
    let pairCalls = 0, forge = false, forged = false
    const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async () => { throw new Error("The matching job must never save.") } }))
    fetch = spyOn(globalThis, "fetch").mockImplementation((async (_url: Parameters<typeof globalThis.fetch>[0], init?: Parameters<typeof globalThis.fetch>[1]) => {
      const { prompt, modelRole } = JSON.parse(String(init?.body))
      expect(modelRole).toBe("fast")
      const part = sourcePart(prompt), candidate = prompt.match(/<candidate encoding="json">([\s\S]*?)<\/candidate>/), contrast = contrastPart(prompt)
      let text: string
      if (part) {
        sourceCalls.push(part)
        expect(part.passages.reduce((count, passage) => count + passage.text.length, 0)).toBeLessThanOrEqual(4000)
        for (const passage of part.passages) { expect(/^[\uDC00-\uDFFF]|[\uD800-\uDBFF]$/.test(passage.text)).toBe(false) }
        const own = part.passages.filter(passage => passage.text.startsWith("Identity:"))
        const quotes = own.map(passage => ({ paragraph: passage.paragraph, text: passage.text }))
        if (forge && !forged && !own.length) { forged = true; quotes.push({ paragraph: 2, text: identity }) }
        text = JSON.stringify({ quotes, limitations: [] })
      } else if (contrast) text = contrastAnswer(contrast)
      else if (candidate) {
        pairCalls++
        expect(prompt).not.toContain("UNSELECTED_LONG_SOURCE")
        expect(prompt).toContain("model-selected")
        expect(Buffer.byteLength(prompt)).toBeLessThanOrEqual(64 * 1024)
        const row = JSON.parse(candidate[1]), incoming = JSON.parse(prompt.match(/<incoming encoding="json">([\s\S]*?)<\/incoming>/)[1])
        expect(row).not.toHaveProperty("paragraphs")
        expect(incoming.reading.coverage.partsProcessed).toBe(incoming.reading.coverage.totalParts)
        const selected = (reading: { parts: { quotes: { text: string }[] }[] }) => reading.parts.flatMap(part => part.quotes)[0]!.text
        text = JSON.stringify({ versionId: row.versionId, relation: row.versionId === earlier.id ? "same-document" : "different-document", reason: "Protocol fixture only.", evidence: [{ versionId: row.versionId, incomingQuote: selected(incoming.reading), candidateQuote: selected(row.reading) }] })
      } else text = JSON.stringify({ choice: "uncertain", reason: "Fresh wording does not identify either retained base.", evidence: [] })
      return Response.json({ data: { text, model: { providerID: "protocol-fixture", id: "configured-fast" }, usage: { tokens: { input: 12, output: 8, reasoning: 0, cache: { read: 0, write: 0 } }, cost: 0.001 } } })
    }) as typeof globalThis.fetch)
    const request = () => { const body = new FormData(); body.append("file", new File([new Uint8Array(returned)], "Returned.docx")); return app.request("/matters/one/documents/returns/tie-break", { method: "POST", body }) }
    const result = await request()
    expect(result.status).toBe(200)
    expect(pairCalls).toBe(2)
    const incomingParts = sourceCalls.filter(part => part.source === "incoming")
    expect(incomingParts.length).toBeGreaterThan(1)
    expect(new Set(incomingParts.map(part => part.part)).size).toBe(incomingParts.length)
    expect(incomingParts.length).toBe(incomingParts[0]!.parts)
    expect(new Set(sourceCalls.map(part => part.source)).size).toBe(3)
    const reconstructed = incomingParts.sort((left, right) => left.part - right.part).flatMap(part => part.passages).reduce((rows, passage) => {
      rows.set(passage.paragraph, (rows.get(passage.paragraph) ?? "") + passage.text); return rows
    }, new Map<number, string>())
    expect([...reconstructed.values()]).toEqual((await readComparisonParagraphs(returned)).map(paragraph => paragraph.text).filter(Boolean))
    forge = true
    const invalid = await request()
    expect(invalid.status).toBe(409)
    expect((await invalid.json()).code).toBe("match_judgment_unreadable")
    expect(forged).toBe(true)
    expect(pairCalls).toBe(2)
  } finally { fetch?.mockRestore(); rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("new independent candidates during pair work or reconciliation invalidate the complete matching set", async () => {
  for (const phase of ["pair", "reconciliation"] as const) {
    const directory = mkdtempSync(path.join(tmpdir(), "dochaus-return-candidate-drift-"))
    let writes = 0, mutated = false
    const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async () => { writes++; return {} } }))
    let fetch: ReturnType<typeof spyOn> | undefined
    try {
      const wording = "Delivery under this agreement requires written notice and transfer of the specified property by the agreed date."
      await writeDeliverable(directory, "Agreement.docx", await docx(wording, "The Orchard transfer is on Monday."))
      const earlier = readCurrentVersion(directory, "Agreement.docx")!.current
      await writeDeliverableRevision(directory, "Agreement.docx", await docx(wording, "The Orchard transfer is on Tuesday."))
      const returned = await docx(wording, "The Orchard transfer is on Wednesday.")
      const input = { name: "Agreement returned.docx", bytes: new Uint8Array(returned) }
      const before = await inspectDocumentReturn(directory, input)
      expect(before.kind).toBe("needs-confirmation")
      expect(before.candidates).toHaveLength(2)
      const originalIds = new Set(before.candidates.map(candidate => candidate.versionId))
      const originalRows = readAvailableDocumentVersions(directory).filter(row => originalIds.has(row.id)).map(row => [row.document_id, row.id, row.sha256, !!row.current])
      fetch = spyOn(globalThis, "fetch").mockImplementation((async (_url: Parameters<typeof globalThis.fetch>[0], init?: Parameters<typeof globalThis.fetch>[1]) => {
        const body = JSON.parse(String(init?.body))
        const reading = sourcePart(body.prompt)
        const candidate = body.prompt.match(/<candidate encoding="json">([\s\S]*?)<\/candidate>/), contrast = contrastPart(body.prompt)
        if (!reading && !mutated && (phase === "pair" ? Boolean(candidate) : body.prompt.includes("<checked-pair-judgments"))) {
          mutated = true
          await writeDeliverable(directory, "New independent draft.docx", await docx(wording, "The Orchard transfer is on Friday."))
          // In reconciliation, also change inspection kind to a newly retained
          // exact duplicate. Neither original candidate's bytes/current changes.
          if (phase === "reconciliation") await writeDeliverableRevision(directory, "New independent draft.docx", Buffer.from(returned))
        }
        const text = reading ? sourceAnswer(reading) : contrast ? contrastAnswer(contrast) : candidate ? (() => {
          const row = JSON.parse(candidate[1])
          return JSON.stringify({ versionId: row.versionId, relation: row.versionId === earlier.id ? "same-document" : "different-document", reason: "Protocol fixture only.", evidence: [{ versionId: row.versionId, incomingQuote: wording, candidateQuote: wording }] })
        })() : JSON.stringify({ choice: "candidate", versionId: earlier.id, reason: "Protocol fixture only; the lawyer must confirm.", evidence: [{ versionId: earlier.id, incomingQuote: wording, candidateQuote: wording }] })
        return Response.json({ data: { text, model: { providerID: "protocol-fixture", id: "configured-fast" }, usage: { tokens: { input: 12, output: 8, reasoning: 0, cache: { read: 0, write: 0 } }, cost: 0.001 } } })
      }) as typeof globalThis.fetch)
      const form = new FormData(); form.append("file", new File([new Uint8Array(returned)], input.name))
      const result = await app.request("/matters/one/documents/returns/tie-break", { method: "POST", body: form })
      expect(mutated).toBe(true)
      expect(readAvailableDocumentVersions(directory).filter(row => originalIds.has(row.id)).map(row => [row.document_id, row.id, row.sha256, !!row.current])).toEqual(originalRows)
      const after = await inspectDocumentReturn(directory, input)
      if (phase === "pair") { expect(after.kind).toBe(before.kind); expect(after.candidates).toHaveLength(3) }
      else expect(after.kind).toBe("duplicate")
      expect(result.status).toBe(409)
      const response = await result.json()
      expect(response.code).toBe("match_source_changed")
      expect(response).not.toHaveProperty("advisory")
      expect(writes).toBe(0)
    } finally { fetch?.mockRestore(); rmSync(directory, { recursive: true, force: true }) }
  }
}, 30_000)

test("common transaction identity cannot promote the sole non-current version to a base", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-return-identity-only-"))
  let fetch: ReturnType<typeof spyOn> | undefined
  let contrasts = 0, reconciliations = 0
  try {
    const identity = "Agreement for the Orchard transfer between Cedar and Riverstone."
    await writeDeliverable(directory, "Agreement.docx", await docx(identity, "The transfer excludes the northern service strip and retains access rights.", "Completion is Monday."))
    const older = readCurrentVersion(directory, "Agreement.docx")!.current
    await writeDeliverableRevision(directory, "Agreement.docx", await docx(identity, "The transfer includes the northern service strip and retains no access rights.", "Completion is Tuesday."))
    const returned = await docx(identity, "The transfer excludes the northern service strip and retains access rights.", "Completion is Wednesday.")
    fetch = spyOn(globalThis, "fetch").mockImplementation((async (_url, init) => {
      const { prompt } = JSON.parse(String(init?.body)), reading = sourcePart(prompt)
      const candidate = prompt.match(/<candidate encoding="json">([\s\S]*?)<\/candidate>/), contrast = contrastPart(prompt)
      const evidence = (versionId: string) => [{ versionId, incomingQuote: identity, candidateQuote: identity }]
      if (contrast) contrasts++
      if (prompt.includes("<checked-pair-judgments")) reconciliations++
      const text = reading ? JSON.stringify({ quotes: [{ paragraph: 1, text: identity }], limitations: [] }) : contrast ? contrastAnswer(contrast) : candidate
        ? JSON.stringify({ versionId: JSON.parse(candidate[1]).versionId, relation: "same-document", reason: "Common title and parties establish this transaction only.", evidence: evidence(JSON.parse(candidate[1]).versionId) })
        : JSON.stringify({ choice: "candidate", versionId: older.id, reason: "The only non-current same-document candidate.", evidence: evidence(older.id) })
      return Response.json({ data: { text, model: { providerID: "protocol-fixture", id: "configured-fast" }, usage: { tokens: { input: 1, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }, cost: 0.001 } } })
    }) as typeof globalThis.fetch)
    const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async () => { throw new Error("Advice must not save.") } }))
    const body = new FormData(); body.append("file", new File([new Uint8Array(returned)], "Return.docx"))
    const result = await app.request("/matters/one/documents/returns/tie-break", { method: "POST", body })
    expect(result.status).toBe(409)
    expect(contrasts).toBe(2)
    expect(reconciliations).toBe(1)
    const response = await result.json()
    expect(response).not.toHaveProperty("advisory")
    expect(response.error).toContain("reconciliation did not follow its checked candidate evidence")
  } finally { fetch?.mockRestore(); rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("an actual current version omitted by text retrieval remains an identified base rival", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-return-hidden-current-"))
  let fetch: ReturnType<typeof spyOn> | undefined
  try {
    const identity = "The Orchard transfer between Cedar and Riverstone excludes the northern access strip."
    const boilerplate = "Notice must be in writing. Delivery takes effect upon receipt at the agreed address."
    await writeDeliverable(directory, "Agreement.docx", await docx(boilerplate, identity, "Completion is Monday."))
    const older = readCurrentVersion(directory, "Agreement.docx")!.current
    await writeDeliverableRevision(directory, "Agreement.docx", await docx("Completely unrelated astronomy research measures stellar spectra using calibrated instruments."))
    const current = readCurrentVersion(directory, "Agreement.docx")!.current
    await writeDeliverable(directory, "Other.docx", await docx(boilerplate, "Separate Meadow transaction between different contracting companies."))
    const returned = await docx(boilerplate, identity, "Completion is Wednesday.")
    const input = { name: "Return.docx", bytes: new Uint8Array(returned) }
    const inspection = await inspectDocumentReturn(directory, input)
    expect(inspection.candidates).toHaveLength(2)
    expect(inspection.candidates.some(row => row.versionId === current.id)).toBe(false)
    const before = readDocumentVersions(directory, "Agreement.docx"), bytesBefore = readFileSync(path.join(directory, "Agreement.docx"))
    fetch = spyOn(globalThis, "fetch").mockImplementation((async (_url, init) => {
      const { prompt } = JSON.parse(String(init?.body)), reading = sourcePart(prompt), contrast = contrastPart(prompt)
      const candidate = prompt.match(/<candidate encoding="json">([\s\S]*?)<\/candidate>/)
      const text = reading ? sourceAnswer(reading) : contrast ? contrastAnswer(contrast, "uncertain") : candidate ? (() => {
        const row = JSON.parse(candidate[1]), own = row.reading.parts.flatMap((part: { quotes: { text: string }[] }) => part.quotes)[0].text
        return JSON.stringify({ versionId: row.versionId, relation: row.versionId === older.id ? "same-document" : "different-document", reason: "Protocol identity judgment only.", evidence: [{ versionId: row.versionId, incomingQuote: boilerplate, candidateQuote: own }] })
      })() : JSON.stringify({ choice: "uncertain", reason: "The current retained rival cannot be resolved as the same clause; choose the base explicitly.", evidence: [] })
      return Response.json({ data: { text, model: { providerID: "protocol-fixture", id: "configured-fast" }, usage: { tokens: { input: 1, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }, cost: 0.001 } } })
    }) as typeof globalThis.fetch)
    const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async () => { throw new Error("Read-only advice.") } }))
    const body = new FormData(); body.append("file", new File([new Uint8Array(returned)], input.name))
    const response = await app.request("/matters/one/documents/returns/tie-break", { method: "POST", body })
    expect(response.status).toBe(200)
    const result = await response.json()
    expect(result.advisory.choice).toBe("uncertain")
    expect(result.advisory.baseSources).toContainEqual(expect.objectContaining({ versionId: current.id, current: true, documentName: "Agreement.docx", versionNumber: 2, sha256: current.sha256 }))
    expect(result.advisory.baseContrasts).toHaveLength(1)
    expect(result.advisory.baseContrasts[0].unresolved.length).toBeGreaterThan(0)
    expect(readDocumentVersions(directory, "Agreement.docx")).toEqual(before)
    expect(readFileSync(path.join(directory, "Agreement.docx"))).toEqual(bytesBefore)
  } finally { fetch?.mockRestore(); rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("a new low-overlap same-document rival invalidates advice even when retrieval and original current flags stay unchanged", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-return-hidden-rival-drift-"))
  let fetch: ReturnType<typeof spyOn> | undefined, mutated = false, reconciliations = 0
  try {
    const identity = "Agreement between Cedar and Riverstone for the Orchard property transfer."
    const oldWords = "The transfer excludes the northern strip and retains access rights."
    await writeDeliverable(directory, "Agreement.docx", await docx(identity, oldWords, "Completion is Monday."))
    const older = readCurrentVersion(directory, "Agreement.docx")!.current
    await writeDeliverableRevision(directory, "Agreement.docx", await docx(identity, "The transfer includes the northern strip and retains no access rights.", "Completion is Tuesday."))
    const current = readCurrentVersion(directory, "Agreement.docx")!.current
    const returned = await docx(identity, oldWords, "Completion is Wednesday.")
    const hidden = await docx("Astronomical observations determine stellar distances through parallax measurements using calibrated telescopes.")
    const input = { name: "Return.docx", bytes: new Uint8Array(returned) }, before = await inspectDocumentReturn(directory, input)
    const rowSnapshot = () => readAvailableDocumentVersions(directory).filter(row => [older.id, current.id].includes(row.id)).map(row => [row.id, row.sha256, row.current])
    const beforeRows = rowSnapshot()
    fetch = spyOn(globalThis, "fetch").mockImplementation((async (_url, init) => {
      const { prompt } = JSON.parse(String(init?.body)), reading = sourcePart(prompt), contrast = contrastPart(prompt)
      const candidate = prompt.match(/<candidate encoding="json">([\s\S]*?)<\/candidate>/)
      if (contrast && !mutated) {
        mutated = true
        const db = openDb(directory)
        try { db.transaction(() => {
          captureVersion(db, directory, { name: "Agreement.docx", bytes: hidden, version: { origin: "upload" } })
          db.run("UPDATE versioned_documents SET current_version=? WHERE id=?", [current.id, current.document_id])
        })() } finally { db.close() }
      }
      if (prompt.includes("<checked-pair-judgments")) reconciliations++
      const text = reading ? sourceAnswer(reading) : contrast ? contrastAnswer(contrast) : candidate ? (() => {
        const row = JSON.parse(candidate[1])
        return JSON.stringify({ versionId: row.versionId, relation: "same-document", reason: "Protocol identity only.", evidence: [{ versionId: row.versionId, incomingQuote: identity, candidateQuote: identity }] })
      })() : JSON.stringify({ choice: "uncertain", reason: "No distinguishing support.", evidence: [] })
      return Response.json({ data: { text, model: { providerID: "protocol-fixture", id: "configured-fast" }, usage: { tokens: { input: 1, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }, cost: 0.001 } } })
    }) as typeof globalThis.fetch)
    const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async () => { throw new Error("Read-only advice.") } }))
    const body = new FormData(); body.append("file", new File([new Uint8Array(returned)], input.name))
    const response = await app.request("/matters/one/documents/returns/tie-break", { method: "POST", body })
    expect(mutated).toBe(true)
    expect(rowSnapshot()).toEqual(beforeRows)
    expect((await inspectDocumentReturn(directory, input)).candidates).toEqual(before.candidates)
    expect(readAvailableDocumentVersions(directory)).toHaveLength(3)
    expect(response.status).toBe(409)
    expect((await response.json()).code).toBe("match_source_changed")
    expect(reconciliations).toBe(0)
  } finally { fetch?.mockRestore(); rmSync(directory, { recursive: true, force: true }) }
}, 30_000)

test("a corrected received association cannot promote its non-current physical source over the qualified current target", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-return-corrected-owner-"))
  let fetch: ReturnType<typeof spyOn> | undefined, positive = 0
  try {
    const identity = "Agreement between Cedar and Riverstone for the Orchard property transfer."
    const excluded = "The transfer excludes the northern strip and retains access rights."
    const included = "The transfer includes the northern strip and retains no access rights."
    await writeDeliverable(directory, "Source.docx", await docx(identity, included, "Completion is Monday."))
    const wrongBase = readCurrentVersion(directory, "Source.docx")!.current
    await writeDeliverableRevision(directory, "Source.docx", await docx(identity, excluded, "Completion is Tuesday."))
    const physical = readCurrentVersion(directory, "Source.docx")!.current
    const receipt = recordDocumentReceived(directory, { versionId: physical.id, baseVersionId: wrongBase.id, requestId: "wrong-owner" })
    await writeDeliverable(directory, "Target.docx", await docx(identity, included, "Completion is Friday."))
    const target = readCurrentVersion(directory, "Target.docx")!.current
    relinkDocumentReceivedDocument(directory, { exchangeId: receipt.id, targetDocumentId: target.document_id, baseVersionId: target.id,
      expectedDocumentId: physical.document_id, expectedBaseVersionId: wrongBase.id, requestId: "correct-owner" })
    await writeDeliverableRevision(directory, "Source.docx", await docx(identity, included, "Completion is Thursday."))
    const returned = stampDocumentReturnMarker(await docx(identity, excluded, "Completion is Wednesday."), { documentId: physical.document_id, versionId: physical.id })
    const inspection = await inspectDocumentReturn(directory, { name: "Return.docx", bytes: returned })
    expect(inspection.kind).toBe("needs-confirmation")
    expect(inspection.candidates).toHaveLength(2)
    const beforeSource = readDocumentVersions(directory, "Source.docx"), beforeTarget = readDocumentVersions(directory, "Target.docx")
    fetch = spyOn(globalThis, "fetch").mockImplementation((async (_url, init) => {
      const { prompt } = JSON.parse(String(init?.body)), reading = sourcePart(prompt), contrast = contrastPart(prompt)
      const candidate = prompt.match(/<candidate encoding="json">([\s\S]*?)<\/candidate>/)
      let text: string
      if (reading) text = sourceAnswer(reading)
      else if (contrast) {
        const support = contrast.incoming.text === contrast.left.passage.text ? "left" : contrast.incoming.text === contrast.right.passage.text ? "right" : "neither"
        if (support !== "neither") positive++
        text = contrastAnswer(contrast, support)
      } else if (candidate) {
        const row = JSON.parse(candidate[1])
        text = JSON.stringify({ versionId: row.versionId, relation: "same-document", reason: "Protocol identity only.", evidence: [{ versionId: row.versionId, incomingQuote: identity, candidateQuote: identity }] })
      } else {
        const summary = JSON.parse(prompt.match(/<checked-base-contrasts encoding="json">([\s\S]*?)<\/checked-base-contrasts>/)[1])
        expect(summary.allowedBaseIds).toEqual([])
        expect(summary.representative).toEqual([])
        text = JSON.stringify({ choice: "uncertain", reason: "The source receipt was linked to another document; confirm its actual logical base.", evidence: [] })
      }
      return Response.json({ data: { text, model: { providerID: "protocol-fixture", id: "configured-fast" }, usage: { tokens: { input: 1, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }, cost: 0.001 } } })
    }) as typeof globalThis.fetch)
    const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async () => { throw new Error("Read-only advice.") } }))
    const body = new FormData(); body.append("file", new File([new Uint8Array(returned)], "Return.docx"))
    const response = await app.request("/matters/one/documents/returns/tie-break", { method: "POST", body })
    expect(response.status).toBe(200)
    const result = await response.json()
    expect(positive).toBe(2)
    expect(result.advisory.choice).toBe("uncertain")
    expect(result.advisory.baseSources).toContainEqual(expect.objectContaining({ versionId: physical.id, documentId: physical.document_id, current: false, associationCorrected: true,
      correctedAssociations: [{ documentId: target.document_id, baseVersionId: target.id }] }))
    expect(result.advisory.baseSources).toContainEqual(expect.objectContaining({ versionId: target.id, documentId: target.document_id, current: true }))
    expect(readDocumentVersions(directory, "Source.docx")).toEqual(beforeSource)
    expect(readDocumentVersions(directory, "Target.docx")).toEqual(beforeTarget)
  } finally { fetch?.mockRestore(); rmSync(directory, { recursive: true, force: true }) }
}, 30_000)
