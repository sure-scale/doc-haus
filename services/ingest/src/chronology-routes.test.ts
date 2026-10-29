import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { chronologyRoutes } from "./chronology-routes"
import { currentChronologySources, newChronologyEvent, readChronology } from "./chronology-store"
import { liveText } from "../../../dochaus/lib/extract"
import { openDb, upsertDocument } from "./db"
const dirs: string[] = []
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})
function setup() {
  const directory = mkdtempSync(path.join(tmpdir(), "chronology-route-"))
  dirs.push(directory)
  writeFileSync(path.join(directory, "matter.json"), "{}")
  const app = chronologyRoutes((id) => (id === "synthetic" ? directory : path.join(directory, "missing")))
  return { app, directory }
}
const payload = (value: unknown) => ({
  method: "POST",
  headers: { "content-type": "application/json" },
  body: JSON.stringify(value),
})
test("routes reject invalid matter ids, missing matters, stale revisions and foreign event ids", async () => {
  const { app } = setup()
  expect((await app.request("/matters/UPPER/chronology")).status).toBe(400)
  expect((await app.request("/matters/other/chronology")).status).toBe(404)
  const initial = await app.request("/matters/synthetic/chronology")
  expect(initial.status).toBe(200)
  const add = await app.request(
    "/matters/synthetic/chronology/actions",
    payload({
      expectedRevision: 0,
      action: {
        type: "add",
        title: "Manual entry",
        detail: "Source not supplied",
        date: { original: "Unknown", precision: "undated" },
      },
    }),
  )
  expect(add.status).toBe(200)
  const stale = await app.request(
    "/matters/synthetic/chronology/actions",
    payload({
      expectedRevision: 0,
      action: {
        type: "add",
        title: "Stale second entry",
        detail: "No source",
        date: { original: "Unknown", precision: "undated" },
      },
    }),
  )
  expect(stale.status).toBe(409)
  const foreign = await app.request(
    "/matters/synthetic/chronology/actions",
    payload({ expectedRevision: 1, action: { type: "review", eventId: "other-matter-event", reviewed: true } }),
  )
  expect(foreign.status).toBe(404)
})
test("run endpoint requires a fresh explicit privilege choice even for refresh and resume", async () => {
  const { app, directory } = setup()
  for (const mode of ["initial", "refresh", "resume"])
    expect(
      (await app.request("/matters/synthetic/chronology/run", payload({ expectedRevision: 0, mode }))).status,
    ).toBe(400)
  const stopped = await app.request(
    "/matters/synthetic/chronology/run",
    payload({ expectedRevision: 0, mode: "initial", privilege: "stop" }),
  )
  expect(stopped.status).toBe(200)
  expect(readChronology(directory).created).toBe(false)
})
test("chat may stage supported content but cannot forge a saved reviewed event", async () => {
  const { app, directory } = setup(),
    name = "Dispatch.md",
    quote = "The package was despatched on 2 June 2025."
  const docPath = path.join(directory, name)
  writeFileSync(docPath, quote)
  const db = openDb(directory)
  try {
    upsertDocument(db, docPath, name, 1, null)
  } finally {
    db.close()
  }
  const source = currentChronologySources(directory)[0]!,
    canonical = await liveText(docPath),
    at = canonical.indexOf(quote)
  const event = newChronologyEvent({
    title: "Package despatched",
    detail: "The dispatch note reports despatch.",
    date: { original: "2 June 2025", precision: "exact", start: "2025-06-02" },
    accounts: [
      {
        id: "dispatch-account",
        title: "Package despatched",
        detail: "Recorded dispatch",
        date: { original: "2 June 2025", precision: "exact", start: "2025-06-02" },
        privilege: "uncertain",
        people: [],
        status: "reported",
        significance: "background",
        gaps: [],
        evidence: [
          {
            id: "dispatch-evidence",
            documentName: name,
            sourceId: source.sourceId,
            version: source.version,
            excerpt: quote,
            section: "Dispatch",
            charStart: at,
            charEnd: at + quote.length,
            verified: true,
            freshness: "current",
            provenance: "source",
          },
        ],
      },
    ],
  })
  event.reviewed = true
  event.key = true
  event.hidden = true
  const staged = await app.request(
    "/matters/synthetic/chronology/proposals",
    payload({
      expectedRevision: 0,
      proposals: [
        {
          id: "chat-dispatch",
          kind: "addition",
          targetIds: [],
          events: [event],
          reason: "New source account.",
          recommendation: "Inspect the dispatch note.",
        },
      ],
    }),
  )
  expect(staged.status).toBe(200)
  const state = (await staged.json()) as any
  expect(state.events).toHaveLength(0)
  expect(state.proposals[0].events[0]).toMatchObject({ reviewed: false, key: false, hidden: false })
  const accept = await app.request(
    "/matters/synthetic/chronology/proposals/chat-dispatch/decision",
    payload({ expectedRevision: state.revision, decision: "accept" }),
  )
  expect(accept.status).toBe(200)
  const retry = await app.request(
    "/matters/synthetic/chronology/proposals/chat-dispatch/decision",
    payload({ expectedRevision: state.revision, decision: "accept" }),
  )
  expect(retry.status).toBe(200)
  expect(readChronology(directory).events).toHaveLength(1)
})
