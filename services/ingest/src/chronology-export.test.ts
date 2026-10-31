import { afterEach, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { Document, Packer, Paragraph } from "docx"
import type { ChronologyAccount, ChronologyDate } from "../../../dochaus/lib/chronology"
import {
  CHRONOLOGY_EXPORT_ORIGIN,
  CHRONOLOGY_EXPORT_PROPERTY,
  validChronologyExportRequest,
} from "../../../dochaus/lib/chronology-export"
import { chronologyExportSnapshot } from "./chronology-export"
import { chronologyRoutes } from "./chronology-routes"
import {
  currentChronologySources,
  mutateChronology,
  newChronologyEvent,
  readChronology,
  verifyChronologyProposals,
} from "./chronology-store"
import { chronologySources } from "./chronology-generation"
import { getDocumentSourceProvenance, listDocuments, openDb, upsertDocument } from "./db"
import {
  DerivedChronologyExportError,
  hasChronologyExportOrigin,
  recognizeChronologyExport,
} from "./chronology-export-origin"
import { ingestDocument } from "./ingest"
import * as embeddings from "./embed"
import { liveText } from "../../../dochaus/lib/extract"

const directories: string[] = []
afterEach(() => {
  for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true })
})
const date: ChronologyDate = {
  original: "About June 2025",
  precision: "approximate",
  start: "2025-06",
  qualifier: "about",
}
const quote = "The recorded delivery used a cobalt seal."

function setup() {
  const directory = mkdtempSync(path.join(tmpdir(), "chronology-export-"))
  directories.push(directory)
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ title: "Synthetic delivery review" }))
  const file = path.join(directory, "Delivery.md")
  writeFileSync(file, quote)
  const db = openDb(directory)
  try {
    upsertDocument(db, file, "Delivery.md", 1, null)
  } finally {
    db.close()
  }
  const source = currentChronologySources(directory)[0]!
  const account: ChronologyAccount = {
    id: "account-1",
    title: "Delivery reported",
    detail: "The sender reports delivery.",
    date,
    privilege: "uncertain",
    status: "reported",
    significance: "relevant",
    people: ["Synthetic sender"],
    attribution: "The sender",
    gaps: [],
    evidence: [
      {
        id: "evidence-1",
        documentName: source.documentName,
        sourceId: source.sourceId,
        version: source.version,
        excerpt: quote,
        section: "Delivery",
        charStart: 0,
        charEnd: quote.length,
        verified: true,
        freshness: "current",
        provenance: "source",
      },
    ],
  }
  const event = newChronologyEvent({
    id: "event-b",
    title: "Delivery reported",
    detail: "Delivery is disputed.",
    date,
    accounts: [account],
  })
  event.reviewed = true
  event.disputed = true
  event.dispute = "The accounts disagree."
  event.accounts.push({
    ...structuredClone(account),
    id: "account-2",
    title: "Delivery required",
    detail: "Delivery was required; completion is not established.",
    status: "required",
    attribution: "The recipient",
  })
  const earlier = newChronologyEvent({
    id: "event-a",
    title: "Earlier notice",
    detail: "A notice is recorded.",
    date: { original: "2024", precision: "year", start: "2024" },
    origin: "manual",
  })
  const hidden = newChronologyEvent({
    id: "hidden",
    title: "Hidden note",
    detail: "Excluded unless requested.",
    date: { original: "Unknown", precision: "undated" },
    origin: "manual",
  })
  hidden.hidden = true
  const old = newChronologyEvent({ id: "old", title: "Superseded account", detail: "Replaced in review.", date })
  old.supersededBy = [event.id]
  mutateChronology(directory, 0, "synthetic fixture", (state) => {
    state.created = true
    state.events = [event, hidden, old, earlier]
    state.proposals.push({
      id: "pending",
      kind: "addition",
      targetIds: [],
      baseEventHashes: {},
      events: [newChronologyEvent({ title: "Unaccepted proposal", detail: "Not a saved fact", date })],
      reason: "Pending",
      recommendation: "Review",
      status: "pending",
      origin: "refresh",
      createdAt: 1,
    })
  })
  const app = chronologyRoutes(() => directory)
  const request = {
    expectedRevision: readChronology(directory).revision,
    scope: "all-visible" as const,
    includeHidden: false,
  }
  return { directory, app, request, source, file }
}
const payload = (body: unknown) => ({
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(body),
})

test("chronology-wide gaps and failed processing stay visible in preview and Word", async () => {
  const { directory, app } = setup()
  const next = mutateChronology(directory, readChronology(directory).revision, "synthetic fixture", (state) => {
    state.gaps = [
      "The referenced inspection attachment is missing.",
      "The referenced inspection attachment is missing.",
    ]
    state.run = {
      id: "incomplete",
      mode: "initial",
      status: "failed",
      privilege: "mixed",
      startedAt: 1,
      sourceKeys: [],
      completedKeys: [],
      errors: ["One source section could not be read."],
    }
  })
  const input = { expectedRevision: next.revision, scope: "all-visible", includeHidden: false }
  const preview = await app.request("/matters/synthetic/chronology/export/preview", payload(input))
  const snapshot = await preview.json()
  expect(snapshot.warnings.join(" ")).toContain("inspection attachment is missing")
  expect(snapshot.warnings.join(" ")).toContain("building failed")
  expect(snapshot.warnings.join(" ")).toContain("source section could not be read")
  expect(snapshot.warnings.filter((warning: string) => warning.includes("inspection attachment"))).toHaveLength(1)
  const exported = await app.request("/matters/synthetic/chronology/export", payload(input))
  const xml = strFromU8(unzipSync(new Uint8Array(await exported.arrayBuffer()))["word/document.xml"]!)
  expect(xml).toContain("inspection attachment is missing")
  expect(xml).toContain("source section could not be read")
  expect(xml).toContain("building failed")
})

test("snapshot separates saved scope, hidden choices and pending proposals without changing the chronology", async () => {
  const { directory, app, request } = setup()
  const saved = readFileSync(path.join(directory, ".dochaus/chronology.json"), "utf8")
  const preview = await app.request("/matters/synthetic/chronology/export/preview", payload(request))
  expect(preview.status).toBe(200)
  const snapshot = await preview.json()
  expect(snapshot.events.map((event: { id: string }) => event.id)).toEqual(["event-a", "event-b"])
  expect(snapshot).toMatchObject({
    totalEvents: 3,
    excludedHidden: 1,
    pendingProposals: 1,
    title: "Synthetic delivery review",
  })
  expect(snapshot.events[1].accounts).toHaveLength(2)
  expect(snapshot.events[1].date).toEqual(date)
  expect(snapshot.warnings.join(" ")).toContain("privilege")
  const hidden = await app.request(
    "/matters/synthetic/chronology/export/preview",
    payload({ ...request, scope: "filtered", eventIds: ["hidden"] }),
  )
  expect(hidden.status).toBe(400)
  const explicit = await app.request(
    "/matters/synthetic/chronology/export/preview",
    payload({ ...request, scope: "filtered", eventIds: ["hidden", "event-b"], includeHidden: true }),
  )
  expect(explicit.status).toBe(200)
  expect((await explicit.json()).events.map((event: { id: string }) => event.id)).toEqual(["event-b", "hidden"])
  expect(readFileSync(path.join(directory, ".dochaus/chronology.json"), "utf8")).toBe(saved)
})

test("stale selections fail before writing and malformed scopes cannot select proposal or foreign events", async () => {
  const { directory, app, request } = setup()
  expect(validChronologyExportRequest({ ...request, eventIds: ["event-b"] })).toBe(false)
  expect(validChronologyExportRequest({ ...request, scope: "filtered", eventIds: ["event-b", "event-b"] })).toBe(false)
  for (const input of [
    { ...request, expectedRevision: request.expectedRevision - 1 },
    { ...request, scope: "filtered", eventIds: ["foreign-event"] },
    { ...request, scope: "filtered", eventIds: ["old"] },
  ]) {
    expect((await app.request("/matters/synthetic/chronology/export", payload(input))).status).toBe(409)
  }
  const db = openDb(directory)
  try {
    expect(listDocuments(db)).not.toContainEqual(expect.objectContaining({ role: "deliverable" }))
  } finally {
    db.close()
  }
})

test("fresh source projection suppresses cached redacted excerpts in preview and actual DOCX", async () => {
  const { directory, app, request, file } = setup()
  expect(chronologyExportSnapshot(directory, "synthetic", request).events[1].accounts[0].evidence[0].excerpt).toBe(
    quote,
  )
  writeFileSync(file, "A redacted replacement has removed the seal description.")
  const snapshot = chronologyExportSnapshot(directory, "synthetic", request)
  expect(snapshot.events[1].accounts[0].evidence[0]).toMatchObject({ excerpt: "", freshness: "stale" })
  const response = await app.request("/matters/synthetic/chronology/export", payload(request))
  expect(response.status).toBe(200)
  const parts = unzipSync(new Uint8Array(await response.arrayBuffer()))
  const xml = strFromU8(parts["word/document.xml"]!)
  expect(xml).not.toContain("cobalt seal")
  expect(xml).toContain("Historical excerpt omitted")
  expect(readChronology(directory).events[0].accounts[0].evidence[0].excerpt).toBe(quote)
  rmSync(file)
  expect(chronologyExportSnapshot(directory, "synthetic", request).events[1].accounts[0].evidence[0]).toMatchObject({
    excerpt: "",
    freshness: "unavailable",
  })
})

test("DOCX response is the exact registered derived version and keeps account context and origin through paragraph pointers", async () => {
  const { directory, app, request } = setup()
  const prior = readChronology(directory)
  const response = await app.request("/matters/synthetic/chronology/export", payload(request))
  expect(response.status).toBe(200)
  expect(response.headers.get("Content-Type")).toContain("wordprocessingml.document")
  const name = response.headers.get("X-DocHaus-Export-Name")!
  const bytes = Buffer.from(await response.arrayBuffer())
  expect(bytes.equals(readFileSync(path.join(directory, name)))).toBe(true)
  const db = openDb(directory)
  try {
    expect(listDocuments(db)).toContainEqual(expect.objectContaining({ name, role: "deliverable" }))
    expect(
      db.query("SELECT sha256,origin FROM document_versions WHERE label='Chronology export snapshot'").get(),
    ).toEqual({ sha256: createHash("sha256").update(bytes).digest("hex"), origin: "generated" })
    expect(db.query("SELECT COUNT(*) AS n FROM chunks WHERE doc_name=?").get(name)).toEqual({ n: 0 })
  } finally {
    db.close()
  }
  const parts = unzipSync(bytes)
  const xml = strFromU8(parts["word/document.xml"]!)
  for (const text of [
    "The sender",
    "The recipient",
    "completion is not established",
    "About June 2025",
    "Unreviewed",
    "uncertain",
    "Evidence notes",
  ])
    expect(xml).toContain(text)
  expect(xml).not.toContain("Unaccepted proposal")
  expect(xml).not.toContain("Hidden note")
  expect(hasChronologyExportOrigin(bytes)).toBe(true)
  expect(recognizeChronologyExport(directory, "Renamed.docx", bytes)).toEqual({ originalName: name })
  expect(chronologySources(directory).map((source) => source.documentName)).toEqual(["Delivery.md"])
  expect(readChronology(directory)).toEqual(prior)
})

test("origin recognition survives renamed ordinary re-saves but does not claim arbitrary stripped-copy detection", async () => {
  const { directory, app, request } = setup()
  const response = await app.request("/matters/synthetic/chronology/export", payload(request))
  const parts = unzipSync(new Uint8Array(await response.arrayBuffer()))
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!) + "\n")
  const resaved = zipSync(parts)
  expect(hasChronologyExportOrigin(resaved)).toBe(true)
  expect(recognizeChronologyExport(directory, "Lawyer-edited.docx", resaved)).toEqual({ originalName: undefined })
  delete parts["docProps/custom.xml"]
  expect(hasChronologyExportOrigin(zipSync(parts))).toBe(false)
  const unrelated = await Packer.toBuffer(
    new Document({ sections: [{ children: [new Paragraph("An unrelated source account.")] }] }),
  )
  expect(recognizeChronologyExport(directory, "Other.docx", unrelated)).toBeUndefined()
})

test("a recognized import first requires acknowledgement, then remains supplied chronology through replacement and verification", async () => {
  const { directory } = setup()
  const bytes = await Packer.toBuffer(
    new Document({
      customProperties: [{ name: CHRONOLOGY_EXPORT_PROPERTY, value: CHRONOLOGY_EXPORT_ORIGIN }],
      sections: [{ children: [new Paragraph("The lawyer corrects the receipt date to 7 July 2025.")] }],
    }),
  )
  await expect(ingestDocument(directory, "Lawyer corrections.docx", bytes)).rejects.toBeInstanceOf(
    DerivedChronologyExportError,
  )
  expect(existsSync(path.join(directory, "Lawyer corrections.docx"))).toBe(false)
  const embed = spyOn(embeddings, "embedChunk").mockResolvedValue(new Float32Array([1, 0, 0]))
  try {
    const text = "The lawyer corrects the receipt date to 7 July 2025."
    const accepted = await ingestDocument(directory, "Lawyer corrections.docx", bytes, {
      sourceProvenance: "supplied-chronology",
      extracted: { text, findings: [], quality: { status: "ok" } },
    })
    expect(accepted.source_provenance).toBe("supplied-chronology")
    const db = openDb(directory)
    try {
      expect(getDocumentSourceProvenance(db, accepted.docPath)).toBe("supplied-chronology")
      upsertDocument(db, accepted.docPath, accepted.name, 2, null)
      expect(getDocumentSourceProvenance(db, accepted.docPath)).toBe("supplied-chronology")
    } finally {
      db.close()
    }
    expect(chronologySources(directory).find((source) => source.documentName === accepted.name)?.kind).toBe(
      "supplied-chronology",
    )
    expect(
      currentChronologySources(directory).find((source) => source.documentName === accepted.name)?.provenance,
    ).toBe("supplied-chronology")
    const original = readChronology(directory)
    const event = structuredClone(original.events[0])
    const source = currentChronologySources(directory).find((source) => source.documentName === accepted.name)!
    event.accounts[0].evidence[0] = {
      ...event.accounts[0].evidence[0],
      sourceId: source.sourceId,
      documentName: accepted.name,
      version: source.version,
      provenance: "source",
    }
    const proposal = {
      id: "incorrect-origin",
      kind: "addition" as const,
      targetIds: [],
      baseEventHashes: {},
      events: [event],
      reason: "Incorrect corroboration",
      recommendation: "Review",
      status: "pending" as const,
      origin: "chat" as const,
      createdAt: 1,
    }
    await expect(verifyChronologyProposals(directory, [proposal])).rejects.toThrow("provenance")
    const canonical = await liveText(accepted.docPath)
    const start = canonical.indexOf(text)
    event.accounts[0].evidence[0] = {
      ...event.accounts[0].evidence[0],
      provenance: "supplied-chronology",
      excerpt: text,
      charStart: start,
      charEnd: start + text.length,
    }
    await expect(verifyChronologyProposals(directory, [proposal])).resolves.toBeUndefined()
    expect(readChronology(directory)).toEqual(original)
  } finally {
    embed.mockRestore()
  }
})
