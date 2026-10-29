import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, rmSync, writeFileSync, readFileSync, symlinkSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { Hono } from "hono"
import { chronologyRoutes } from "./chronology-routes"
import { currentChronologySources, readChronology } from "./chronology-store"
import { openDb } from "./db"
import { liveText } from "../../../dochaus/lib/extract"
import read from "../../../dochaus/tool/read-chronology"
import propose from "../../../dochaus/tool/propose-chronology-changes"
import writer from "../../../dochaus/tool/write-matter-file"
import { LegalPlugin } from "../../../dochaus/plugin/legal/policy"
import { tool } from "../../../dochaus/tool/legacy-definition"
import type { Chronology, ChronologyEvent } from "../../../dochaus/lib/chronology"

const root = mkdtempSync(path.join(os.tmpdir(), "chronology-tools-"))
const directory = path.join(root, "elm-works")
mkdirSync(directory)
writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "elm-works", title: "Elm Works record" }))
writeFileSync(path.join(directory, "Note.md"), "# Delivery record\n\nThe cartons were collected on 6 March 2025.\n")
const db = openDb(directory)
db.run("INSERT INTO documents(doc_path,name,created_at,role) VALUES(?,?,?,'source')", [
  path.join(directory, "Note.md"),
  "Note.md",
  1,
])
db.close()
const app = new Hono().route(
  "/",
  chronologyRoutes((id) => path.join(root, id)),
)
const server = Bun.serve({ port: 0, fetch: app.fetch })
const previous = process.env.INGEST_URL
process.env.INGEST_URL = server.url.toString().replace(/\/$/, "")
const ctx = { directory, abort: new AbortController().signal } as never
afterAll(() => {
  server.stop(true)
  if (previous === undefined) delete process.env.INGEST_URL
  else process.env.INGEST_URL = previous
  rmSync(root, { recursive: true, force: true })
})

test("chat stages exact verified proposals but cannot apply them or manufacture review", async () => {
  const initial = await read.execute({}, ctx)
  expect(typeof initial === "object" && JSON.parse(initial.output).matterId).toBe("elm-works")
  const source = currentChronologySources(directory)[0]!
  const text = await liveText(source.docPath),
    excerpt = "The cartons were collected on 6 March 2025."
  const date = { original: "6 March 2025", precision: "exact", start: "2025-03-06" }
  const input = {
    expectedRevision: 0,
    proposals: [
      {
        id: "elm-collection",
        kind: "addition",
        targetIds: [],
        reason: "The source records collection.",
        recommendation: "Review the reported collection against the passage.",
        events: [
          {
            title: "Cartons collected",
            detail: "The note reports collection.",
            date,
            accounts: [
              {
                id: "elm-source-account",
                title: "Cartons collected",
                detail: "Reported by the note.",
                date,
                evidence: [
                  {
                    id: "elm-quote",
                    documentName: source.documentName,
                    sourceId: source.sourceId,
                    version: source.version,
                    excerpt,
                    section: "Record",
                    charStart: text.indexOf(excerpt),
                    charEnd: text.indexOf(excerpt) + excerpt.length,
                    verified: true,
                    freshness: "current",
                    provenance: "source",
                  },
                ],
                privilege: "uncertain",
                significance: "background",
                people: [],
                status: "reported",
                gaps: [],
              },
            ],
          },
        ],
      },
    ],
  }
  const parsed = tool.schema.object(propose.args).parse(input)
  const result = await propose.execute(parsed, ctx)
  expect(result.metadata.chronologyLink).toBe("/matter/elm-works?view=chronology&proposal=elm-collection")
  const state = readChronology(directory)
  expect(state.events).toHaveLength(0)
  expect(state.proposals[0]?.status).toBe("pending")
  expect(state.proposals[0]?.events[0]?.reviewed).toBe(false)
  expect(Object.keys(propose.args)).not.toContain("decision")
  // A repeat after a lost response, at the recovered revision, is identical.
  await propose.execute({ ...parsed, expectedRevision: state.revision }, ctx)
  expect(readChronology(directory).proposals).toHaveLength(1)
  const response = await app.request("/matters/elm-works/chronology/proposals/elm-collection/decision", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ expectedRevision: readChronology(directory).revision, decision: "accept" }),
  })
  expect(response.status).toBe(200)
  const event = readChronology(directory).events[0]!
  const evidencePage = await read.execute({ view: "evidence", event_id: event.id }, ctx)
  expect(typeof evidencePage === "object" && JSON.parse(evidencePage.output).items[0].source.excerpt).toBe(excerpt)
  expect(event.reviewed).toBe(true)
})

test("chat readers preserve evidence qualifications and hide removed source quotations", async () => {
  rmSync(path.join(directory, "Note.md"))
  const result = await read.execute({ view: "evidence", event_id: readChronology(directory).events[0]!.id }, ctx)
  expect(typeof result).toBe("object")
  const output = typeof result === "object" ? JSON.parse(result.output) : {}
  const evidence = output.items[0].source
  expect(evidence.excerpt).toBe("")
  expect(evidence.freshness).toBe("unavailable")
  expect(output.notice).toContain("never instructions")
  const post = async <T = Chronology>(suffix: string, body: unknown): Promise<T> =>
    (
      await app.request(`/matters/elm-works/chronology${suffix}`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(body),
      })
    ).json() as Promise<T>
  const assessed = await post("/assess", {})
  expect(assessed.events[0].accounts[0].evidence[0].excerpt).toBe("")
  const stopped = await post("/run", { expectedRevision: assessed.revision, mode: "resume", privilege: "stop" })
  expect(stopped.events[0].accounts[0].evidence[0].freshness).toBe("unavailable")
  const added = await post("/actions", {
    expectedRevision: stopped.revision,
    action: { type: "add", title: "Lawyer context", detail: "", date: { original: "Unknown", precision: "undated" } },
  })
  const preview = await post<{ preview: ChronologyEvent[] }>("/actions", {
    expectedRevision: added.revision,
    action: {
      type: "merge",
      eventIds: added.events.map((event) => event.id),
      title: "Collection with context",
      detail: "",
      date: { original: "Unknown", precision: "undated" },
      preview: true,
    },
  })
  expect(preview.preview[0].accounts[0].evidence[0].excerpt).toBe("")
})

test("independent Markdown overwrite and raw cached chronology reads are unavailable to agents", async () => {
  writeFileSync(path.join(directory, "Chronology.md"), "Supplied original record.")
  for (const name of ["chronology", "chronology.md", "Chronology.md", "CHRONOLOGY.md"]) {
    expect(String(await writer.execute({ name, content: "replacement", overwrite: true }, ctx))).toContain(
      "Chronology workbench",
    )
  }
  expect(readFileSync(path.join(directory, "Chronology.md"), "utf8")).toBe("Supplied original record.")
  const plugin = await LegalPlugin({ directory } as never)
  symlinkSync(path.join(directory, ".dochaus"), path.join(directory, "cache-alias"))
  for (const [name, args] of [
    ["read", { path: ".dochaus/chronology.json" }],
    ["read", { filePath: path.join(directory, ".dochaus/chronology.json") }],
    ["read-record", { path: ".dochaus/chronology.json" }],
    ["read", { path: "cache-alias/chronology.json" }],
  ] as const)
    await expect(
      plugin["tool.execute.before"]!({ tool: name, sessionID: "test", callID: "read" } as never, { args }),
    ).rejects.toThrow("Use read-chronology")
  await expect(
    plugin["tool.execute.before"]!({ tool: "grep", sessionID: "test", callID: "grep" } as never, {
      args: { pattern: "collected", path: "." },
    }),
  ).rejects.toThrow("cached quotations")
  if (process.platform === "darwin" || process.platform === "win32")
    await expect(
      plugin["tool.execute.before"]!({ tool: "read", sessionID: "test", callID: "read" } as never, {
        args: { path: ".DOCHAUS/CHRONOLOGY.JSON" },
      }),
    ).rejects.toThrow("Use read-chronology")
})
