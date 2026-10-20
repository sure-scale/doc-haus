import { test, expect, afterAll } from "bun:test"
import { createHash } from "node:crypto"
import { appendFileSync, mkdtempSync, rmSync, existsSync, mkdirSync, readFileSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import type { Matter } from "./matter"

// Point WORKSPACE_ROOT at a throwaway dir before matter.ts captures it at module
// eval, so the CRUD tests never touch a real workspace. The import is dynamic for
// exactly this reason — a static import would hoist above the env assignment.
const root = mkdtempSync(path.join(tmpdir(), "dochaus-matter-"))
process.env.WORKSPACE_ROOT = root
const matter = await import("./matter")

afterAll(() => rmSync(root, { recursive: true, force: true }))

test("matterDir rejects path traversal and escaping ids", () => {
  // The fix for the traversal hole: ids arrive from a URL param, so anything that
  // could climb out of WORKSPACE_ROOT must throw before it reaches a filesystem path.
  for (const bad of ["..", "../evil", "a/b", "/etc/passwd", "foo/../bar", ".dochaus", "", "Foo", "a b"]) {
    expect(() => matter.matterDir(bad)).toThrow(/invalid matter id/)
  }
})

test("matterDir accepts the generated slug+uuid shape and stays under the root", () => {
  const id = "acme-merger-a1b2c3"
  expect(matter.matterDir(id)).toBe(path.join(root, id))
})

test("create, list, get, rename, delete round-trip", () => {
  const created = matter.createMatter("Acme / Merger 2026", "2026-0042", ["EW"])
  expect(created.id).toMatch(/^acme-merger-2026-[a-z0-9]{6}$/)
  expect(created.reference).toBe("2026-0042")
  expect(created.jurisdictions).toEqual(["EW"])
  expect(existsSync(path.join(root, created.id, "matter.json"))).toBe(true)

  expect(matter.listMatters().map((m) => m.id)).toContain(created.id)
  expect(matter.getMatter(created.id).title).toBe("Acme / Merger 2026")
  expect(matter.getMatter(created.id).jurisdictions).toEqual(["EW"])

  // Jurisdictions move on their own through the rename endpoint while title/ref hold.
  const renamed = matter.renameMatter(created.id, "Acme Acquisition", "2026-0099", ["US-NY", "EW"])
  expect(renamed.id).toBe(created.id) // id is stable across rename
  expect(renamed.title).toBe("Acme Acquisition")
  expect(matter.getMatter(created.id).reference).toBe("2026-0099")
  expect(matter.getMatter(created.id).jurisdictions).toEqual(["US-NY", "EW"])

  matter.deleteMatter(created.id)
  expect(existsSync(path.join(root, created.id))).toBe(false)
  expect(matter.listMatters().map((m) => m.id)).not.toContain(created.id)
})

test("create publishes an exact hash-only PERMANENT matter origin receipt", () => {
  const created = matter.createMatter("Privileged Falcon Instructions", "SECRET-0042", ["EW"])
  const content = readFileSync(path.join(created.dir, "matter.json"))
  const receipts = matterLineage(created.id)

  expect(receipts).toHaveLength(1)
  expect(receipts[0]).toMatchObject({
    tier: "PERMANENT",
    artifact: "matter.json",
    op: "write",
    stage: null,
    subagent: "ingest",
    content_sha256: sha256(content),
    bytes: content.byteLength,
    supersedes: null,
    provenance: {
      origin: "matter-metadata",
      mutation_kind: "create",
    },
  })
  expect(receipts[0]!.run).toBe(`ingest:${receipts[0]!.provenance!.correlation_id}`)
  expect(receipts[0]!.provenance!.matter_mutation_id).toBeTruthy()
  expect(JSON.stringify(receipts[0])).not.toContain("Privileged Falcon Instructions")
  expect(JSON.stringify(receipts[0])).not.toContain("SECRET-0042")

  matter.deleteMatter(created.id)
})

test("rename appends a superseding receipt without replacing source artifacts or their lineage", () => {
  const created = matter.createMatter("Preserved Sources")
  const source = path.join(created.dir, "source-index.json")
  const sourceBody = '{"documents":["Agreement.docx"]}'
  const sourceReceipt = {
    entry_id: "source-entry",
    tier: "PERMANENT",
    artifact: "source-index.json",
    content_sha256: sha256(Buffer.from(sourceBody)),
  }
  writeFileSync(source, sourceBody)
  appendFileSync(path.join(created.dir, ".dochaus", "lineage.jsonl"), `${JSON.stringify(sourceReceipt)}\n`)
  const first = matterLineage(created.id)[0]!

  matter.renameMatter(created.id, "Preserved Sources Renamed", "REF-2", ["US-NY"])
  const content = readFileSync(path.join(created.dir, "matter.json"))
  const receipts = matterLineage(created.id)

  expect(receipts).toHaveLength(2)
  expect(receipts[1]).toMatchObject({
    op: "edit",
    content_sha256: sha256(content),
    bytes: content.byteLength,
    supersedes: first.entry_id,
    provenance: { origin: "matter-metadata", mutation_kind: "rename" },
  })
  expect(readFileSync(source, "utf8")).toBe(sourceBody)
  expect(lineage(created.id)).toContainEqual(sourceReceipt)

  matter.deleteMatter(created.id)
})

test("every accepted conflict mutation appends its exact superseding receipt", () => {
  const created = matter.createMatter("Conflict Receipt Chain")
  const cleared = matter.recordConflicts(created.id, { status: "cleared", by: "private.clearer" })
  const clearedContent = readFileSync(path.join(cleared.dir, "matter.json"))
  const pending = matter.recordConflicts(created.id, { status: "pending", owner: "private.owner", due: "2026-07-21" })
  const pendingContent = readFileSync(path.join(pending.dir, "matter.json"))
  const bypassed = matter.recordConflicts(created.id, {
    status: "bypassed",
    by: "private.partner",
    reason: "privileged lateral-screen rationale",
  })
  const bypassedContent = readFileSync(path.join(bypassed.dir, "matter.json"))
  const receipts = matterLineage(created.id)

  expect(receipts).toHaveLength(4)
  expect(receipts.slice(1).map((row) => row.provenance!.mutation_kind)).toEqual(["conflicts", "conflicts", "conflicts"])
  expect(receipts[1]!.content_sha256).toBe(sha256(clearedContent))
  expect(receipts[2]!.content_sha256).toBe(sha256(pendingContent))
  expect(receipts[3]!.content_sha256).toBe(sha256(bypassedContent))
  expect(receipts.slice(1).map((row) => row.supersedes)).toEqual(receipts.slice(0, 3).map((row) => row.entry_id))
  expect(JSON.stringify(receipts)).not.toContain("private.clearer")
  expect(JSON.stringify(receipts)).not.toContain("private.owner")
  expect(JSON.stringify(receipts)).not.toContain("privileged lateral-screen rationale")

  matter.deleteMatter(created.id)
})

test("startup listing recovers a committed matter body from its durable origin intent", () => {
  const created = matter.createMatter("Before Restart")
  const current = matter.getMatter(created.id)
  const intent = stageInterruptedMatterMutation(created.id, { ...current, title: "After Restart" }, "rename", "restart")

  expect(matter.listMatters().find((entry) => entry.id === created.id)!.title).toBe("After Restart")
  const receipts = matterLineage(created.id)
  expect(receipts).toHaveLength(2)
  expect(receipts[1]).toMatchObject({
    entry_id: intent.entry_id,
    ts: intent.ts,
    content_sha256: intent.content_sha256,
    bytes: intent.bytes,
    supersedes: intent.supersedes,
    provenance: {
      matter_mutation_id: intent.mutation_id,
      correlation_id: intent.correlation_id,
      mutation_kind: "rename",
    },
  })
  expect(existsSync(path.join(created.dir, ".dochaus", "matter-origin-mutation.json"))).toBe(false)

  matter.deleteMatter(created.id)
})

test("restart recovery is idempotent when the exact receipt already settled", () => {
  const created = matter.createMatter("Idempotent Before")
  const intent = stageInterruptedMatterMutation(
    created.id,
    { ...matter.getMatter(created.id), title: "Idempotent After" },
    "rename",
    "idempotent",
  )

  expect(matter.getMatter(created.id).title).toBe("Idempotent After")
  expect(matterLineage(created.id)).toHaveLength(2)
  writeFileSync(
    path.join(created.dir, ".dochaus", "matter-origin-mutation.json"),
    `${JSON.stringify(intent, null, 2)}\n`,
  )
  expect(matter.getMatter(created.id).title).toBe("Idempotent After")
  expect(matterLineage(created.id)).toHaveLength(2)

  matter.deleteMatter(created.id)
})

test("listMatters isolates malformed metadata without hiding valid matters", () => {
  const created = matter.createMatter("Recovery remains visible")
  const malformed = path.join(root, "malformed-a1b2c3")
  const incomplete = path.join(root, "incomplete-a1b2c3")
  const dangling = path.join(root, "dangling-a1b2c3")
  mkdirSync(malformed)
  mkdirSync(incomplete)
  symlinkSync(path.join(root, "missing-target"), dangling)
  writeFileSync(path.join(malformed, "matter.json"), "{not-json")
  writeFileSync(path.join(incomplete, "matter.json"), JSON.stringify({ id: "incomplete-a1b2c3" }))

  expect(matter.listMatters().map((entry) => entry.id)).toContain(created.id)
  expect(matter.listMatters().map((entry) => entry.id)).not.toContain("malformed-a1b2c3")
  expect(matter.listMatters().map((entry) => entry.id)).not.toContain("incomplete-a1b2c3")
  expect(matter.listMatters().map((entry) => entry.id)).not.toContain("dangling-a1b2c3")

  matter.deleteMatter(created.id)
  rmSync(malformed, { recursive: true })
  rmSync(incomplete, { recursive: true })
  rmSync(dangling)
})

test("recordConflicts: the three valid exits persist via getMatter", () => {
  const created = matter.createMatter("Conflicts Exits")

  const cleared = matter.recordConflicts(created.id, { status: "cleared", by: "j.smith" })
  expect(cleared.conflicts!.status).toBe("cleared")
  expect(cleared.conflicts!.by).toBe("j.smith")
  expect(cleared.conflicts!.recorded_at).toBeGreaterThan(0)
  expect(matter.getMatter(created.id).conflicts!.status).toBe("cleared")

  const pending = matter.recordConflicts(created.id, { status: "pending", owner: "a.jones", due: "2026-07-01" })
  expect(pending.conflicts!.owner).toBe("a.jones")
  expect(pending.conflicts!.due).toBe("2026-07-01")
  expect(matter.getMatter(created.id).conflicts!.status).toBe("pending")

  const bypassed = matter.recordConflicts(created.id, {
    status: "bypassed",
    by: "partner.lee",
    reason: "lateral hire screen in place",
  })
  expect(bypassed.conflicts!.by).toBe("partner.lee")
  expect(bypassed.conflicts!.reason).toBe("lateral hire screen in place")
  expect(matter.getMatter(created.id).conflicts!.status).toBe("bypassed")

  // The rest of the matter survives the record untouched.
  expect(matter.getMatter(created.id).title).toBe("Conflicts Exits")

  matter.deleteMatter(created.id)
})

test("recordConflicts rejects half-recorded checks", () => {
  const created = matter.createMatter("Conflicts Rejections")

  expect(() => matter.recordConflicts(created.id, { status: "cleared" })).toThrow(/requires "by"/)
  expect(() => matter.recordConflicts(created.id, { status: "pending" })).toThrow(/requires "owner"/)
  expect(() => matter.recordConflicts(created.id, { status: "pending", owner: "a.jones", due: "soonish" })).toThrow(
    /ISO date/,
  )
  expect(() => matter.recordConflicts(created.id, { status: "bypassed", by: "partner.lee" })).toThrow(
    /requires "by" and "reason"/,
  )
  expect(() => matter.recordConflicts(created.id, { status: "bypassed", reason: "screen" })).toThrow(
    /requires "by" and "reason"/,
  )
  expect(() => matter.recordConflicts(created.id, { status: "waived" as never })).toThrow(/status must be/)

  // A rejected record never lands in matter.json.
  expect(matter.getMatter(created.id).conflicts).toBeUndefined()

  matter.deleteMatter(created.id)
})

test("client is stored as one normalised line and omitted when empty", () => {
  const created = matter.createMatter("Client Create", undefined, undefined, undefined, "  Northwind\tTraders\r\n\u0007Ltd  ")
  expect(created.client).toBe("Northwind Traders Ltd")
  expect(readMatterJson(created.id).client).toBe("Northwind Traders Ltd")
  expect(matter.listMatters().find((m) => m.id === created.id)!.client).toBe("Northwind Traders Ltd")

  const blank = matter.createMatter("Client Blank", undefined, undefined, undefined, " \n\t ")
  expect("client" in readMatterJson(blank.id)).toBe(false)
  const absent = matter.createMatter("Client Absent")
  expect("client" in readMatterJson(absent.id)).toBe(false)

  matter.deleteMatter(created.id)
  matter.deleteMatter(blank.id)
  matter.deleteMatter(absent.id)
})

test("client over 200 characters or not a string is refused, never truncated", () => {
  const before = matter.listMatters().length
  expect(() => matter.createMatter("Client Long", undefined, undefined, undefined, "N".repeat(201))).toThrow(
    matter.MatterValidationError,
  )
  expect(() => matter.createMatter("Client Number", undefined, undefined, undefined, 42 as unknown as string)).toThrow(
    /client must be a string/,
  )
  expect(matter.listMatters()).toHaveLength(before)
  // The cap applies after normalising, so padding does not count.
  expect(matter.createMatter("Client Cap", undefined, undefined, undefined, ` ${"N".repeat(200)}\n`).client).toHaveLength(200)

  const created = matter.createMatter("Client Rename Refusal", undefined, undefined, undefined, "Northwind Traders Ltd")
  expect(() => matter.renameMatter(created.id, "Client Rename Refusal", undefined, undefined, undefined, "N".repeat(201))).toThrow(
    /at most 200 characters/,
  )
  expect(() => matter.renameMatter(created.id, "Client Rename Refusal", undefined, undefined, undefined, null as unknown as string)).toThrow(
    matter.MatterValidationError,
  )
  expect(readMatterJson(created.id).client).toBe("Northwind Traders Ltd")

  matter.deleteMatter(created.id)
})

test("rename changes the client only when one is given", () => {
  const created = matter.createMatter("Client Rename", "2026-0007", ["EW"], undefined, "Northwind Traders Ltd")

  // Callers that do not edit the client (title or playbook edits) keep it.
  matter.renameMatter(created.id, "Client Rename Kept", "2026-0008", ["EW"], "playbook-nda")
  expect(readMatterJson(created.id)).toMatchObject({ title: "Client Rename Kept", reference: "2026-0008", client: "Northwind Traders Ltd" })

  const set = matter.renameMatter(created.id, "Client Rename Kept", "2026-0008", ["EW"], "playbook-nda", "Contoso\nHoldings plc")
  expect(set.client).toBe("Contoso Holdings plc")
  expect(readMatterJson(created.id).client).toBe("Contoso Holdings plc")

  matter.renameMatter(created.id, "Client Rename Kept", "2026-0008", ["EW"], "playbook-nda", "")
  const cleared = readMatterJson(created.id)
  expect("client" in cleared).toBe(false)
  expect(cleared).toMatchObject({ title: "Client Rename Kept", reference: "2026-0008", playbook: "playbook-nda" })

  matter.deleteMatter(created.id)
})

test("listJurisdictions surfaces the bundled EW pack from the dochaus config layer", () => {
  const ew = matter.listJurisdictions().find((j) => j.code === "EW")
  expect(ew).toBeDefined()
  expect(ew!.name).toBe("England & Wales")
  expect(ew!.citationStyle).toBe("OSCOLA")
})

type TestLineageRow = {
  entry_id: string
  ts?: string
  tier: string
  artifact: string
  op?: string
  run?: string
  stage?: string | null
  subagent?: string
  content_sha256: string
  bytes?: number
  supersedes?: string | null
  provenance?: {
    origin?: string
    matter_mutation_id?: string
    correlation_id?: string
    mutation_kind?: string
  }
}

function lineage(id: string) {
  return readFileSync(path.join(matter.matterDir(id), ".dochaus", "lineage.jsonl"), "utf8")
    .trim()
    .split("\n")
    .map((row) => JSON.parse(row) as TestLineageRow)
}

function readMatterJson(id: string) {
  return JSON.parse(readFileSync(path.join(matter.matterDir(id), "matter.json"), "utf8")) as Record<string, unknown>
}

function matterLineage(id: string) {
  return lineage(id).filter((row) => row.artifact === "matter.json")
}

function sha256(content: Buffer) {
  return createHash("sha256").update(content).digest("hex")
}

function stageInterruptedMatterMutation(
  id: string,
  body: Matter,
  mutationKind: "rename" | "conflicts",
  identity: string,
) {
  const content = Buffer.from(JSON.stringify(body, null, 2))
  const prior = matterLineage(id).at(-1)!
  const intent = {
    version: 1,
    artifact: "matter.json",
    mutation_id: `matter-mutation-${identity}`,
    correlation_id: `matter-correlation-${identity}`,
    mutation_kind: mutationKind,
    entry_id: `matter-entry-${identity}`,
    ts: "2026-07-19T00:00:00.000Z",
    content_sha256: sha256(content),
    bytes: content.byteLength,
    supersedes: prior.entry_id,
  }
  writeFileSync(path.join(matter.matterDir(id), "matter.json"), content)
  writeFileSync(
    path.join(matter.matterDir(id), ".dochaus", "matter-origin-mutation.json"),
    `${JSON.stringify(intent, null, 2)}\n`,
  )
  return intent
}
