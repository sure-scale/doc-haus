import { afterAll, beforeAll, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { realpathSync } from "node:fs"
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { Database } from "bun:sqlite"
import { tmpdir } from "node:os"
import path from "node:path"
import { pathToFileURL } from "node:url"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import * as listLabels from "../../../dochaus/lib/docx-list-labels"
import type { MarkupAuthor, MarkupComment, MarkupRevision } from "../../../dochaus/lib/docx-markup"
import { recordRedlineProposal } from "../../../dochaus/lib/redlines"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "../../../dochaus/test/support/counterparty-markup-fixture"
import { documentMarkupRoutes } from "./document-markup"
import { docxodus } from "./docxodus"
import { reconstructText } from "./structure"

// Counterparty markup triage (#200) through the real ingest routes of an
// isolated service, on the synthetic two-reviewer supply agreement: the lawyer's
// accept, reject and hold decisions become one new version of the source. The
// fixture's ten Docxodus units, in document order (docx-markup.test.ts):
//   0 Avery deletes "thirty"            5 Avery deletes the planted instruction
//   1 Avery inserts "forty-five"        6 Avery inserts two phrases (one unit)
//   2 Morgan inserts "the first "       7 Morgan deletes the second inside them
//   3 Avery moves the Notices clause    8 Morgan deletes clause 7's paragraph mark
//   4 Morgan bolds "indirect"           9 Morgan inserts the Waiver clause

type Decision = "accept" | "reject" | "hold"
type View = { name: string; version: { id: string | null; number: number | null; sha256: string }; current: boolean
  revisions: MarkupRevision[]; comments: MarkupComment[]; authors: MarkupAuthor[] }
type Row = { id: string; kind: string; family: string; author: string; date: string | null; decision: Decision; outcome: string }
type Added = { thread: number; id: number }
type Applied = { name: string; version: { id: string; number: number; origin: string; created: boolean }
  counts: { accepted: number; rejected: number; held: number; consumed: number }; decisions: Row[]; replies: Added[]; notice?: string }
type Finding = { rule: string; detail: string; snippet?: string; charStart?: number }
type History = { versions: Array<{ id: string; number: number; origin: string; actor: string; label: string | null }>
  events: Array<{ kind: string; version_id: string | null; actor: string; detail: { markup: Row[]; replies?: Added[] } | null }> }

const [AVERY, MORGAN] = COUNTERPARTY_MARKUP.reviewers as [string, string]
const { dates, plantedDeletion, plantedComment } = COUNTERPARTY_MARKUP
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const SUPPLY = (days: string) => `1. Supply. The Supplier shall deliver the Goods within ${days} days of each Order.`
const PRICE = "2. Price. The Price is fixed for twelve months."
const PRICE_FIRST = "2. Price. The Price is fixed for the first twelve months."
const NOTICES = "3. Notices. Notices must be in writing."
const LIABILITY = "4. Liability. Neither party is liable for indirect loss."
const TERM = "5. Term. This Agreement continues for two years."
const AUDIT = "6. Audit. The Buyer may audit the Supplier's records."
const AUDIT_NOTICE = "6. Audit. The Buyer may audit on ten Business Days' notice the Supplier's records."
const ASSIGNMENT = "7. Assignment. Neither party may assign"
const WAIVER = "8. Waiver. No waiver is effective unless in writing."
const LAW = "9. Governing law. This Agreement is governed by the law of the Province."

// The four decision sets, by unit, and the paragraphs each leaves. Mixed B holds
// four units, so its result is read as its accept-all and reject-all projections.
const SETS = {
  acceptAll: { decisions: Array<Decision>(10).fill("accept"), paragraphs: [
    SUPPLY("forty-five"), PRICE_FIRST, LIABILITY, NOTICES, TERM, AUDIT_NOTICE, `${ASSIGNMENT} without consent.`, WAIVER, LAW] },
  rejectAll: { decisions: Array<Decision>(10).fill("reject"), paragraphs: [
    SUPPLY("thirty"), PRICE, NOTICES, LIABILITY, `${TERM} ${plantedDeletion}`, AUDIT, ASSIGNMENT, " without consent.", LAW] },
  // Every unit decided; Morgan's deletion goes with Avery's rejected insertion.
  mixedA: { decisions: ["reject", "reject", "accept", "reject", "accept", "reject", "reject", "accept", "reject", "accept"] as Decision[], paragraphs: [
    SUPPLY("thirty"), PRICE_FIRST, NOTICES, LIABILITY, `${TERM} ${plantedDeletion}`, AUDIT, ASSIGNMENT, " without consent.", WAIVER, LAW] },
  // Avery's replace accepted, Morgan's other changes rejected; the move, the
  // planted deletion, Avery's insertions and the paragraph mark held, and
  // Morgan's deletion inside the held insertion accepted.
  mixedB: { decisions: ["accept", "accept", "reject", "hold", "reject", "hold", "hold", "accept", "hold", "reject"] as Decision[],
    accepted: [SUPPLY("forty-five"), PRICE, LIABILITY, NOTICES, TERM, AUDIT_NOTICE, `${ASSIGNMENT} without consent.`, LAW],
    rejected: [SUPPLY("forty-five"), PRICE, NOTICES, LIABILITY, `${TERM} ${plantedDeletion}`, AUDIT, ASSIGNMENT, " without consent.", LAW] },
} as const
type SetName = keyof typeof SETS

let temporary = ""
let service: ReturnType<typeof Bun.spawn> | undefined
let engine: ReturnType<typeof Bun.serve> | undefined
let base = ""
let matter = { id: "", dir: "" }

const sha256 = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
const request = (route: string, options?: RequestInit) => fetch(base + route, { ...options, signal: AbortSignal.timeout(60_000) })
const matterRoute = (route: string, options?: RequestInit) => request(`/matters/${matter.id}${route}`, options)
const query = (name: string, extra: Record<string, string> = {}) => new URLSearchParams({ name, ...extra })
const upload = async (name: string, bytes = counterpartyMarkupDocx()) => {
  const body = new FormData()
  body.append("file", new File([new Uint8Array(bytes)], name))
  const response = await matterRoute("/documents", { method: "POST", body })
  expect(response.status).toBe(200)
  return await response.json() as { injection: { findings: Finding[] } | null; notice?: string; tracked_changes?: { count: number; authors: string[] } }
}
const markup = (name: string, extra: Record<string, string> = {}) => matterRoute(`/documents/markup?${query(name, extra)}`)
const view = async (name: string) => {
  const response = await markup(name)
  expect(response.status).toBe(200)
  return await response.json() as View
}
const post = (name: string, body: unknown, headers: Record<string, string> = {}) =>
  matterRoute(`/documents/markup/decisions?${query(name)}`, { method: "POST", headers: { "content-type": "application/json", ...headers }, body: JSON.stringify(body) })
/** Decides every listed unit of `name`'s current version, unit i taking decisions[i]. */
const decide = async (name: string, decisions: readonly Decision[], headers: Record<string, string> = {}) => {
  const listed = await view(name)
  return post(name, { base_sha256: listed.version.sha256, decisions: listed.revisions.map((revision, at) => ({ id: revision.id, decision: decisions[at] })) }, headers)
}
const content = async (name: string) => new Uint8Array(await (await matterRoute(`/documents/content?${query(name)}`)).arrayBuffer())
const history = async (name: string) => await (await matterRoute(`/documents/versions?${query(name)}`)).json() as History
const working = (name: string) => readFile(path.join(matter.dir, name))
const saveAttorney = async (attorney: string) =>
  expect((await request("/preferences", { method: "PUT", headers: { "content-type": "application/json" }, body: JSON.stringify({ attorney }) })).status).toBe(200)
// A thread as the list shows it: the comments whose root is `root`, in list order.
const threadOf = (comments: MarkupComment[], root: number) => comments.filter(comment => comment.id === root || comment.parent === root)
  .map(({ id, author, initials, text, parent, resolved }) => ({ id, author, initials, text, parent, resolved }))

// Body paragraph text, as Docxodus reads a file with no pending change.
async function paragraphs(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try {
    return session.grep("^[\\s\\S]*$", { scope: 1 }).map(match => match.text)
  } finally {
    session.close()
  }
}
const projections = async (bytes: Uint8Array) => {
  const dx = await docxodus()
  return { accepted: await paragraphs(await dx.docxDiffAcceptRevisions(bytes)), rejected: await paragraphs(await dx.docxDiffRejectRevisions(bytes)) }
}

// Whether "indirect" is bold, and whether that is still a tracked formatting change.
function indirect(bytes: Uint8Array) {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const run = Array.from(doc.getElementsByTagNameNS(W, "r")).find(row =>
    Array.from(row.getElementsByTagNameNS(W, "t")).map(text => text.textContent).join("").includes("indirect"))!
  const properties = Array.from(run.childNodes).find(node => (node as Element).localName === "rPr") as Element | undefined
  const set = properties ? Array.from(properties.childNodes).filter(node => node.nodeType === 1) as Element[] : []
  return { bold: set.some(node => node.localName === "b" && !["0", "false"].includes(node.getAttributeNS(W, "val") ?? "")),
    tracked: set.some(node => node.localName === "rPrChange") }
}

// The source's indexed text, rebuilt from its chunks as citations read it.
function indexedText(name: string) {
  const db = new Database(path.join(matter.dir, ".dochaus", "legal.db"), { readonly: true })
  try {
    return reconstructText(db, path.join(realpathSync(matter.dir), name))
  } finally {
    db.close()
  }
}
const labels = (text: string) => [...text.matchAll(/\[Tracked (insertion|deletion|move from here|move to here) by ([^:\]]+): /g)].map(match => `${match[1]} by ${match[2]}`)
// Every change with wording, as the index labels it before any decision; the
// formatting change and the paragraph-mark deletion have none.
const UNDECIDED_LABELS = [
  `deletion by ${AVERY}`, `insertion by ${AVERY}`, `insertion by ${MORGAN}`, `move from here by ${AVERY}`, `move to here by ${AVERY}`,
  `deletion by ${AVERY}`, `insertion by ${AVERY}`, `deletion by ${MORGAN}`, `insertion by ${MORGAN}`,
]

// Each set applied once through the routes, to its own copy of the fixture.
const applied = new Map<SetName, Promise<{ name: string; listed: View; result: Applied; bytes: Uint8Array }>>()
function applySet(set: SetName) {
  if (!applied.has(set)) applied.set(set, (async () => {
    const name = `${set}.docx`
    await upload(name)
    const listed = await view(name)
    const response = await decide(name, SETS[set].decisions)
    expect(response.status).toBe(200)
    return { name, listed, result: await response.json() as Applied, bytes: await content(name) }
  })())
  return applied.get(set)!
}

beforeAll(async () => {
  temporary = await mkdtemp(path.join(tmpdir(), "dochaus-markup-http-"))
  const config = path.join(temporary, "config")
  await mkdir(path.join(temporary, "workspace"))
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
  engine = Bun.serve({ port: 0, fetch: () => new Response("No engine in this markup test", { status: 503 }) })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const log = path.join(temporary, "ingest.log")
  service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, WORKSPACE_ROOT: path.join(temporary, "workspace"), DOCHAUS_DIR: config,
      INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}` },
    stdout: Bun.file(log), stderr: Bun.file(log),
  })
  base = `http://127.0.0.1:${port}`
  const deadline = Date.now() + 20_000
  while (!await request("/matters").then(response => response.ok).catch(() => false)) {
    if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(log, "utf8")}`)
    await Bun.sleep(50)
  }
  const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Neutral markup matter" }) })
  expect(created.status).toBe(200)
  matter = await created.json() as { id: string; dir: string }
}, 90_000)

afterAll(async () => {
  service?.kill()
  await service?.exited
  engine?.stop(true)
  if (temporary) await rm(temporary, { recursive: true, force: true })
})

test("each decision set leaves the golden paragraphs, and formatting follows its decision", async () => {
  for (const set of ["acceptAll", "rejectAll", "mixedA"] as const) {
    const { listed, result, bytes } = await applySet(set)
    expect(listed).toMatchObject({ version: { number: 1 }, current: true })
    expect(listed.revisions).toHaveLength(10)
    expect(result.version).toMatchObject({ number: 2, origin: "accept", created: true })
    expect(await paragraphs(bytes)).toEqual([...SETS[set].paragraphs])
    // Accepting Morgan's change keeps "indirect" bold; rejecting it restores the old properties.
    expect(indirect(bytes)).toEqual({ bold: SETS[set].decisions[4] === "accept", tracked: false })
  }
  // Rejecting Avery's insertions takes Morgan's deletion inside them along.
  expect((await applySet("rejectAll")).result.counts).toEqual({ accepted: 0, rejected: 9, held: 0, consumed: 1 })
  expect((await applySet("mixedA")).result.counts).toEqual({ accepted: 3, rejected: 6, held: 0, consumed: 1 })

  const { result, bytes } = await applySet("mixedB")
  expect(result.counts).toEqual({ accepted: 3, rejected: 3, held: 4, consumed: 0 })
  expect(result.decisions.map(row => row.outcome)).toEqual(["applied", "applied", "applied", "held", "applied", "held", "held", "applied", "held", "applied"])
  expect(await projections(bytes)).toEqual({ accepted: [...SETS.mixedB.accepted], rejected: [...SETS.mixedB.rejected] })
  expect(indirect(bytes)).toEqual({ bold: false, tracked: false })
}, 180_000)

test("held changes stay pending with their own author and date, and only they stay labelled in the index", async () => {
  const { name, listed, result } = await applySet("mixedB")
  const after = await view(name)
  expect(after).toMatchObject({ version: { id: result.version.id, number: 2 }, current: true })
  expect(after.revisions.map(row => [row.kind, row.family, row.author, row.date, row.text, row.status])).toEqual([
    ["move", "move", AVERY, dates[AVERY], `${NOTICES}¶`, "supported"],
    ["delete", "content_delete", AVERY, dates[AVERY], ` ${plantedDeletion}`, "supported"],
    // Morgan's accepted deletion is gone from inside Avery's held insertion.
    ["insert", "content_insert", AVERY, dates[AVERY], "on ten Business Days' notice ", "supported"],
    ["delete", "paragraph_mark", MORGAN, dates[MORGAN], "¶", "supported"],
  ])
  // Comment threads, replies and resolved flags are carried into the new version.
  const thread = (comments: MarkupComment[]) => comments.map(({ id, author, text, parent, resolved }) => ({ id, author, text, parent, resolved }))
  expect(thread(after.comments)).toEqual(thread(listed.comments))

  // The index labels the held wording only; the decided changes read as agreed text or are gone.
  const text = indexedText(name)
  expect(labels(text)).toEqual([`move from here by ${AVERY}`, `move to here by ${AVERY}`, `deletion by ${AVERY}`, `insertion by ${AVERY}`])
  expect(text).toContain(SUPPLY("forty-five"))
  expect(text).toContain(`[Tracked insertion by ${AVERY}: on ten Business Days' notice ]the Supplier's records.`)
  expect(text).not.toContain("thirty")
  expect(text).not.toContain("the first")
  expect(text).not.toContain("Waiver")
  // Accepting every change leaves no label; before any decision, each change with wording had one.
  expect(labels(indexedText((await applySet("acceptAll")).name))).toEqual([])
  await upload("Undecided.docx")
  expect(labels(indexedText("Undecided.docx"))).toEqual(UNDECIDED_LABELS)
}, 180_000)

test("a stale base, an unknown change, pending doc.haus proposals and an all-hold set write nothing", async () => {
  const name = "Refusals.docx"
  await upload(name)
  const listed = await view(name)
  const before = await working(name)
  const holds = listed.revisions.map(revision => ({ id: revision.id, decision: "hold" }))
  const accept = [{ id: listed.revisions[0]!.id, decision: "accept" }, ...holds.slice(1)]
  const refused = async (response: Response, status: number, code: string) => {
    expect(response.status).toBe(status)
    expect((await response.json() as { code?: string }).code).toBe(code)
    expect(await working(name)).toEqual(before)
    const kept = await history(name)
    expect(kept.versions.map(version => version.origin)).toEqual(["upload"])
    expect(kept.events).toEqual([])
  }
  await refused(await post(name, { base_sha256: sha256(new TextEncoder().encode("another version")), decisions: accept }), 409, "base_changed")
  await refused(await post(name, { base_sha256: listed.version.sha256, decisions: [...accept, { id: "rev2-0000000000000000000", decision: "reject" }] }), 409, "unknown_revision")
  await refused(await post(name, { base_sha256: listed.version.sha256, decisions: holds }), 422, "nothing_decided")
  expect((await post(name, { base_sha256: listed.version.sha256, decisions: [accept[0], accept[0]] })).status).toBe(400)
  recordRedlineProposal(matter.dir, {
    docPath: path.join(matter.dir, name), docName: name, scope: "phrase", findText: "twelve months", oldText: "twelve months",
    newText: "eighteen months", author: "Test reviewer", reviewComment: "Review this proposed wording.", anchorId: "price",
  })
  await refused(await post(name, { base_sha256: listed.version.sha256, decisions: accept }), 409, "proposals_pending")
}, 120_000)

test("the version records the accept, who decided and an event that names each change without its text", async () => {
  const name = "Actors.docx"
  await upload(name)
  const original = await view(name)
  const decisions = (accepting: number[], rejecting: number[] = []) => original.revisions.map((_, at) =>
    accepting.includes(at) ? "accept" : rejecting.includes(at) ? "reject" : "hold") as Decision[]
  // No header and no saved attorney: unknown. Avery's replace is accepted.
  expect((await decide(name, decisions([0, 1]))).status).toBe(200)
  // The saved attorney when the request names nobody. The remaining units shift up by two.
  expect((await request("/preferences", { method: "PUT", headers: { "content-type": "application/json" }, body: JSON.stringify({ attorney: "Sam Rivera" }) })).status).toBe(200)
  expect((await decide(name, decisions([0]))).status).toBe(200)
  // The header wins over the saved attorney. Morgan's Waiver clause is rejected.
  const last = await decide(name, decisions([], [6]), { "x-dochaus-actor": "Riley Chen" })
  expect(last.status).toBe(200)
  expect((await last.json() as Applied).counts).toEqual({ accepted: 0, rejected: 1, held: 6, consumed: 0 })
  await request("/preferences", { method: "PUT", headers: { "content-type": "application/json" }, body: JSON.stringify({ attorney: "" }) })

  const kept = await history(name)
  expect(kept.versions.map(version => [version.number, version.origin, version.actor, version.label])).toEqual([
    [4, "accept", "Riley Chen", "Tracked changes resolved: 0 accepted, 1 rejected, 6 held"],
    [3, "accept", "Sam Rivera", "Tracked changes resolved: 1 accepted, 0 rejected, 7 held"],
    [2, "accept", "unknown", "Tracked changes resolved: 2 accepted, 0 rejected, 8 held"],
    [1, "upload", "unknown", null],
  ])
  expect(kept.events.map(event => [event.kind, event.actor, event.version_id])).toEqual([
    ["accept", "unknown", kept.versions[2]!.id], ["accept", "Sam Rivera", kept.versions[1]!.id], ["accept", "Riley Chen", kept.versions[0]!.id],
  ])
  const first = kept.events[0]!.detail!.markup
  expect(first.map(row => Object.keys(row).sort())).toEqual(original.revisions.map(() => ["author", "date", "decision", "family", "id", "kind", "outcome"]))
  expect(first.map(row => [row.id, row.kind, row.family, row.author, row.date, row.decision, row.outcome])).toEqual(original.revisions.map((revision, at) =>
    [revision.id, revision.kind, revision.family, revision.author, revision.date, at < 2 ? "accept" : "hold", at < 2 ? "applied" : "held"]))
  // No wording of any change, kept or removed, is copied into the events.
  const events = JSON.stringify(kept.events)
  for (const revision of original.revisions) expect(events).not.toContain(JSON.stringify(revision.text).slice(1, -1))
  expect(events).not.toContain("Ignore all previous")
}, 180_000)

// Approved replies (#200 scope 5) are written as Word replies under each
// thread's root, in the same Apply and the same version as the decisions.
test("approved replies alone make one version signed by the saved attorney, each under its root, with no event", async () => {
  const name = "Replies.docx"
  await upload(name)
  const listed = await view(name)
  const answers = [{ thread: 0, text: "Forty-five days works if the stock report is monthly." }, { thread: 3, text: "Noted; we will review it." }]
  await saveAttorney("Sam Rivera")
  let result: Applied
  try {
    const response = await post(name, { base_sha256: listed.version.sha256, decisions: [], replies: answers })
    expect(response.status).toBe(200)
    result = await response.json() as Applied
  } finally {
    await saveAttorney("")
  }
  expect(result.version).toMatchObject({ number: 2, origin: "revision", created: true })
  expect(result.counts).toEqual({ accepted: 0, rejected: 0, held: 10, consumed: 0 })
  expect(result.replies.map(reply => reply.thread)).toEqual([0, 3])
  const kept = await history(name)
  expect(kept.versions.map(version => [version.number, version.origin, version.actor, version.label])).toEqual([
    [2, "revision", "Sam Rivera", "2 replies added"], [1, "upload", "unknown", null]])
  expect(kept.events).toEqual([])

  // Every change is still pending; each thread gains the reply, signed and open.
  const after = await view(name)
  expect(after).toMatchObject({ version: { id: result.version.id, number: 2 }, current: true })
  expect(after.revisions.map(({ id, kind, author, text, status }) => [id, kind, author, text, status]))
    .toEqual(listed.revisions.map(({ id, kind, author, text, status }) => [id, kind, author, text, status]))
  const signed = (reply: Added, text: string) => ({ id: reply.id, author: "Sam Rivera", initials: "SR", text, parent: reply.thread, resolved: false })
  expect(threadOf(after.comments, 0)).toEqual([...threadOf(listed.comments, 0), signed(result.replies[0]!, answers[0]!.text)])
  expect(threadOf(after.comments, 3)).toEqual([...threadOf(listed.comments, 3), signed(result.replies[1]!, answers[1]!.text)])
  expect(threadOf(after.comments, 2)).toEqual(threadOf(listed.comments, 2))
}, 120_000)

test("decisions and replies make one accept version; the event names the replies by id, and a held change stays beside its reply", async () => {
  const name = "Answered.docx"
  await upload(name)
  const listed = await view(name)
  // Morgan's "the first " accepted and Waiver clause rejected; Avery's "forty-five", under thread 0, held.
  const decisions = listed.revisions.map((revision, at) => ({ id: revision.id, decision: (at === 2 ? "accept" : at === 9 ? "reject" : "hold") as Decision }))
  const text = "We can live with forty-five days; confirm the stock report."
  const response = await post(name, { base_sha256: listed.version.sha256, decisions, replies: [{ thread: 0, text }] }, { "x-dochaus-actor": "Riley Chen" })
  expect(response.status).toBe(200)
  const result = await response.json() as Applied
  expect(result.version).toMatchObject({ number: 2, origin: "accept", created: true })
  expect(result.counts).toEqual({ accepted: 1, rejected: 1, held: 8, consumed: 0 })
  expect(result.replies).toEqual([{ thread: 0, id: expect.any(Number) }])

  const kept = await history(name)
  expect(kept.versions.map(version => [version.number, version.origin, version.actor, version.label])).toEqual([
    [2, "accept", "Riley Chen", "Tracked changes resolved: 1 accepted, 1 rejected, 8 held; 1 reply added"], [1, "upload", "unknown", null]])
  expect(kept.events.map(event => [event.kind, event.actor, event.version_id])).toEqual([["accept", "Riley Chen", kept.versions[0]!.id]])
  expect(kept.events[0]!.detail!.replies).toEqual(result.replies)
  expect(kept.events[0]!.detail!.markup.map(row => row.decision)).toEqual(decisions.map(row => row.decision))
  expect(JSON.stringify(kept.events)).not.toContain(text)

  const after = await view(name)
  const commented = listed.revisions[1]!
  expect(after.revisions.find(revision => revision.id === commented.id)).toMatchObject({ kind: "insert", author: AVERY, date: dates[AVERY], text: commented.text, status: "supported" })
  expect(after.revisions).toHaveLength(8)
  expect(threadOf(after.comments, 0)).toEqual([...threadOf(listed.comments, 0),
    { id: result.replies[0]!.id, author: "Riley Chen", initials: "RC", text, parent: 0, resolved: false }])
  expect((await projections(await content(name))).accepted.slice(0, 2)).toEqual([SUPPLY("forty-five"), PRICE_FIRST])
}, 120_000)

test("a reply to a resolved thread, to a reply or to no comment, an unsigned or empty reply, a stale base and an empty Apply write nothing", async () => {
  const name = "Reply refusals.docx"
  await upload(name)
  const listed = await view(name)
  const before = await working(name)
  const refused = async (response: Response, status: number, code: string) => {
    expect(response.status).toBe(status)
    expect((await response.json() as { code?: string }).code).toBe(code)
    expect(await working(name)).toEqual(before)
    const kept = await history(name)
    expect(kept.versions.map(version => version.origin)).toEqual(["upload"])
    expect(kept.events).toEqual([])
  }
  const holds = listed.revisions.map(revision => ({ id: revision.id, decision: "hold" }))
  const reply = (thread: number, text = "Agreed.") => ({ base_sha256: listed.version.sha256, decisions: holds, replies: [{ thread, text }] })
  const signed = { "x-dochaus-actor": "Riley Chen" }
  // No name sent and none saved: a Word comment is never signed "unknown".
  await refused(await post(name, reply(0)), 422, "author_required")
  await refused(await post(name, reply(2), signed), 409, "thread_resolved")
  await refused(await post(name, reply(1), signed), 409, "thread_not_root")
  await refused(await post(name, reply(99), signed), 409, "thread_not_found")
  await refused(await post(name, reply(0, " \u200B\n "), signed), 422, "empty_reply")
  await refused(await post(name, { ...reply(0), base_sha256: sha256(new TextEncoder().encode("another version")) }, signed), 409, "base_changed")
  await refused(await post(name, { ...reply(0), replies: [] }, signed), 422, "nothing_decided")
  expect((await post(name, { ...reply(0), replies: [{ thread: 0, text: "One." }, { thread: 0, text: "Two." }] }, signed)).status).toBe(400)
  expect((await post(name, reply(0, "x".repeat(4_001)), signed)).status).toBe(400)
}, 120_000)

test("a deletion inside a rejected insertion is consumed whichever order the decisions come in, and so is a held one", async () => {
  const nested = (inner: Decision) => (_: MarkupRevision, at: number) => at === 6 ? "reject" : at === 7 ? inner : "hold"
  const outcomes = new Map<string, Row[]>()
  for (const [name, inner, reversed] of [["Nested.docx", "accept", false], ["Nested reversed.docx", "accept", true], ["Nested held.docx", "hold", false]] as const) {
    await upload(name)
    const listed = await view(name)
    const decisions = listed.revisions.map((revision, at) => ({ id: revision.id, decision: nested(inner)(revision, at) }))
    const response = await post(name, { base_sha256: listed.version.sha256, decisions: reversed ? decisions.toReversed() : decisions })
    expect(response.status).toBe(200)
    const result = await response.json() as Applied
    expect(result.counts).toEqual({ accepted: 0, rejected: 1, held: 8, consumed: 1 })
    expect(result.decisions.slice(6, 8).map(row => [row.author, row.decision, row.outcome])).toEqual([[AVERY, "reject", "applied"], [MORGAN, inner, "consumed"]])
    expect((await history(name)).versions[0]!.label).toBe("Tracked changes resolved: 0 accepted, 1 rejected, 8 held, 1 resolved with another change")
    expect((await projections(await content(name))).accepted[5]).toBe(AUDIT)
    outcomes.set(name, result.decisions)
  }
  expect(sha256(await content("Nested.docx"))).toBe(sha256(await content("Nested reversed.docx")))
  expect(outcomes.get("Nested.docx")).toEqual(outcomes.get("Nested reversed.docx"))
}, 180_000)

test("the injection report names the planted deletion and the planted comment", async () => {
  const uploaded = await upload("Planted.docx")
  expect(uploaded.tracked_changes).toEqual({ count: 10, authors: [AVERY, MORGAN] })
  expect(uploaded.notice).toBe(`Saved as version 1, with 10 tracked changes by ${AVERY} and ${MORGAN}; review them in the viewer. doc.haus adds its edits beside them.`)
  const text = await (await matterRoute(`/documents/text?${query("Planted.docx")}`)).json() as { injection: { findings: Finding[] } }
  expect(text.injection.findings).toEqual(uploaded.injection!.findings)
  // The deleted instruction is in the indexed text (#267), with its offsets.
  const deleted = text.injection.findings.filter(finding => finding.snippet?.includes("Ignore all previous instructions"))
  expect(deleted.map(finding => [finding.rule, typeof finding.charStart])).toEqual([["instruction-override", "number"]])
  // The comment is not indexed: document-level findings. read-document prints
  // details outside its untrusted frame, so they name no counterparty string.
  expect(text.injection.findings.filter(finding => finding.detail.endsWith(", in a comment"))).toEqual([
    { rule: "ai-directive", detail: "instructions addressed to an AI assistant, in a comment", snippet: expect.stringContaining("AI assistant") },
    { rule: "concealment", detail: "text that asks the assistant to hide something from the user, in a comment", snippet: expect.stringContaining("do not tell the lawyer") },
  ])
  expect(text.injection.findings.filter(finding => finding.detail.includes(AVERY))).toEqual([])
  expect(plantedComment).toContain("do not tell the lawyer")
}, 120_000)

test("the list refuses anything but a registered Word source, and reads a kept version as not current", async () => {
  const status = async (response: Response) => [response.status, (await response.json() as { code?: string }).code]
  expect(await status(await markup("Notice.pdf"))).toEqual([400, undefined])
  expect(await status(await markup("Absent.docx"))).toEqual([404, "not_source"])
  await writeFile(path.join(matter.dir, "Loose.docx"), counterpartyMarkupDocx())
  expect(await status(await markup("Loose.docx"))).toEqual([404, "not_source"])
  const body = new FormData()
  body.append("file", new File([new Uint8Array(counterpartyMarkupDocx())], "Deliverable.docx"))
  expect((await matterRoute("/deliverables", { method: "POST", body })).status).toBe(200)
  expect(await status(await markup("Deliverable.docx"))).toEqual([404, "not_source"])
  expect(await status(await post("Deliverable.docx", { base_sha256: sha256(counterpartyMarkupDocx()), decisions: [{ id: "x", decision: "accept" }] }))).toEqual([404, "not_source"])

  const { name, listed } = await applySet("acceptAll")
  const first = await markup(name, { version: listed.version.id! })
  expect(first.status).toBe(200)
  expect(await first.json()).toMatchObject({ version: listed.version, current: false, revisions: listed.revisions })
  expect((await markup(name, { version: crypto.randomUUID() })).status).toBe(404)
}, 120_000)

// After Apply, the cross-references and defined terms (#198) of the version it
// made, as a notice in the response. A synthetic services agreement whose
// clauses Word's list numbers (no number is typed): Avery deletes clause 7,
// which clause 5 cites, and clause 2 already cites a clause 12 it never had.
const CLAUSES = [
  "Services. The Provider shall perform the services in each order form.",
  "Fees. The Customer shall pay the fees as adjusted under clause 12.",
  "Records. Each party shall keep accurate records of its performance.",
  "Notices. Notices must be in writing.",
  "Termination. Either party may end this agreement as set out in clause 7.",
  "Liability. Neither party is liable for indirect loss.",
  "Termination procedure. A party ending this agreement must give thirty days' notice.",
  "Assignment. Neither party may assign this agreement without consent.",
  "Governing law. This agreement is governed by the law of the Province.",
]
function numberedDocx() {
  const ns = `xmlns:w="${W}"`, part = "application/vnd.openxmlformats-officedocument.wordprocessingml"
  const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const by = `w:author="${AVERY}" w:date="${dates[AVERY]}"`
  const list = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
  const body = CLAUSES.map((text, at) => at === 6
    ? `<w:p><w:pPr>${list}<w:rPr><w:del w:id="2" ${by}/></w:rPr></w:pPr><w:del w:id="1" ${by}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del></w:p>`
    : `<w:p><w:pPr>${list}</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join("")
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${part}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${part}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${part}.numbering+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="${rel}/styles" Target="styles.xml"/><Relationship Id="rIdN" Type="${rel}/numbering" Target="numbering.xml"/></Relationships>`),
    "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${ns}><w:body>${body}</w:body></w:document>`),
    "word/styles.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${ns}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/numbering.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${ns}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`),
  })
}
type Consistency = { status: "checked"; findings: { detail: string }[] } | { status: "not_checked" }
/** Uploads the numbered agreement as `name` and applies `decision` to each of its listed changes. */
async function decideNumbered(name: string, decision: Decision) {
  await upload(name, numberedDocx())
  const response = await decide(name, (await view(name)).revisions.map(() => decision))
  expect(response.status).toBe(200)
  return (await response.json() as Applied & { consistency: Consistency }).consistency
}

test("after Apply, the reference an accepted deletion broke is reported; with the deletion rejected, nothing is", async () => {
  // Under Word's numbering, clause 8 becomes clause 7, and clause 5 now cites it.
  expect(await decideNumbered("Renumbered.docx", "accept")).toEqual({ status: "checked", findings: [{ detail: expect.stringMatching(/^Clause 5 refers to clause 7\b.*which was clause 8 before/) }] })
  // The same file with the deletion rejected: clause 2's reference to clause 12
  // was there before the decisions, so it is not one to report.
  expect(await decideNumbered("Unchanged.docx", "reject")).toEqual({ status: "checked", findings: [] })
}, 120_000)

test("a check that cannot read the clause numbers is reported as not checked, never as clean", async () => {
  const name = "Unnumbered.docx"
  await upload(name, numberedDocx())
  const listed = await view(name)
  // The route in this process over the service's matter, so the labeller can
  // fail; the stub writer keeps the file as it was.
  const routes = documentMarkupRoutes({ directory: () => matter.dir, locks: (_, file) => [path.basename(file)], write: async (_, written) => ({ name: written }) })
  const unread = spyOn(listLabels, "writeListLabels").mockImplementation(async (buffer, unread) => { unread?.(); return buffer })
  try {
    const response = await routes.request(`/${matter.id}/documents/markup/decisions?${query(name)}`, {
      method: "POST", headers: { "content-type": "application/json", "x-dochaus-actor": "Riley Chen" },
      body: JSON.stringify({ base_sha256: listed.version.sha256, decisions: listed.revisions.map(({ id }) => ({ id, decision: "accept" })) }) })
    expect(response.status).toBe(200)
    expect((await response.json() as { consistency: Consistency }).consistency).toEqual({ status: "not_checked" })
    expect(unread).toHaveBeenCalled()
  } finally {
    unread.mockRestore()
  }
}, 120_000)

// Live check against the installed LibreOffice, skipped where it is absent: the
// same four sets applied there by an in-process macro in a private profile give
// the paragraphs the route gave. LibreOffice lists Docxodus units in pieces (a
// move's two sides, a delete-on-insert stacked on the insertion it exposes when
// rejected), so each unit's decision goes to every piece in its paragraphs with
// its author and kind.
test.skipIf(!Bun.which("soffice"))("LibreOffice applies the same decision sets to the same paragraphs", async () => {
  const work = await mkdtemp(path.join(tmpdir(), "dochaus-markup-soffice-"))
  try {
    const profile = path.join(work, "lo-profile")
    await mkdir(path.join(profile, "user", "Scripts", "python"), { recursive: true })
    await writeFile(path.join(profile, "user", "Scripts", "python", "oracle.py"), ORACLE)
    const input = path.join(work, "fixture.docx")
    await writeFile(input, counterpartyMarkupDocx())
    const listed = (await applySet("acceptAll")).listed.revisions
    const types: Record<MarkupRevision["kind"], string[]> = { insert: ["Insert"], delete: ["Delete"], move: ["Insert", "Delete"], format: ["Format"], structure: [] }
    const sets = Object.keys(SETS).map(name => ({ name, rules: listed.map((revision, at) => ({
      author: revision.author, types: types[revision.kind], paragraphs: revision.locations.map(location => location.paragraph), decision: SETS[name as SetName].decisions[at],
    })) }))
    const job = path.join(work, "job.json")
    await writeFile(job, JSON.stringify({ input, sets }))
    const soffice = Bun.spawn([Bun.which("soffice")!, "--headless", "--norestore", `-env:UserInstallation=${pathToFileURL(profile).href}`,
      "vnd.sun.star.script:oracle.py$run?language=Python&location=user"], { env: { ...process.env, DOCHAUS_ORACLE_JOB: job }, stdout: "pipe", stderr: "pipe" })
    const timer = setTimeout(() => void stop(soffice.pid), 150_000)
    const [exit, stdout, stderr] = await Promise.all([soffice.exited, new Response(soffice.stdout).text(), new Response(soffice.stderr).text()])
    clearTimeout(timer)
    const output = Bun.file(`${job}.out`)
    if (!await output.exists()) throw new Error(`soffice exited ${exit} without a result:\n${stdout}\n${stderr}`)
    const oracle = JSON.parse(await output.text()) as Record<string, { error?: string; paragraphs: Array<{ text: string; bold: string[] }>; pending: number; steps: number[]
      accepted?: Array<{ text: string }>; rejected?: Array<{ text: string }> }> & { error?: string }
    if (oracle.error) throw new Error(`The LibreOffice macro failed:\n${oracle.error}`)
    const text = (rows: Array<{ text: string }>) => rows.map(row => row.text)
    for (const set of Object.keys(SETS) as SetName[]) {
      const lo = oracle[set]!
      const { bytes } = await applySet(set)
      // A selection resolves only the piece it covers.
      expect(lo.steps.every(step => step <= 1)).toBe(true)
      if (set === "mixedB") {
        expect(lo.pending).toBeGreaterThan(0)
        const route = await projections(bytes)
        expect({ accepted: text(lo.accepted!), rejected: text(lo.rejected!) }).toEqual(route)
      } else {
        expect(lo.pending).toBe(0)
        expect(text(lo.paragraphs)).toEqual(await paragraphs(bytes))
        expect(lo.paragraphs.flatMap(row => row.bold)).toEqual(indirect(bytes).bold ? ["indirect"] : [])
      }
    }
  } finally {
    await rm(work, { recursive: true, force: true })
  }
}, 240_000)

/** Stops the running soffice wrapper `pid` and the LibreOffice it started, by
 * pid. The wrapper waits for LibreOffice, so once it exits nothing is left. */
async function stop(pid: number) {
  const children = (await new Response(Bun.spawn(["pgrep", "-P", String(pid)], { stdout: "pipe" }).stdout).text()).split("\n").filter(Boolean).map(Number)
  for (const child of [...children, pid]) try { process.kill(child, "SIGKILL") } catch {}
}

// The in-process LibreOffice macro (Python inside soffice; never LibreOffice's
// own python binary). It reads the job named by DOCHAUS_ORACLE_JOB, applies each
// set to a fresh load of the input and writes <job>.out, then ends LibreOffice.
// A held set is also read with every remaining change accepted, and rejected,
// in the same session (a LibreOffice save and reload moves a comment end into
// the planted deletion, which accepting it then empties clause 5).
const ORACLE = String.raw`import json, os, traceback
import uno
from com.sun.star.beans import PropertyValue

FORMATS = ("Format", "ParagraphFormat", "Attributes")

def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p

def paragraphs(doc):
    e = doc.Text.createEnumeration()
    out = []
    while e.hasMoreElements():
        p = e.nextElement()
        if p.supportsService("com.sun.star.text.Paragraph"):
            out.append(p)
    return out

# 1-based, counting deleted and moved-from paragraphs, as the markup reader does.
def ordinal(doc, ps, start):
    found = None
    for i, p in enumerate(ps):
        if doc.Text.compareRegionStarts(p.getStart(), start) < 0:
            break
        found = i + 1
    return found

def redlines(doc):
    ps = paragraphs(doc)
    out = []
    e = doc.getRedlines().createEnumeration()
    while e.hasMoreElements():
        r = e.nextElement()
        succ = r.getPropertyValue("RedlineSuccessorData")
        succ = {p.Name: p.Value for p in succ} if succ else None
        out.append({"id": r.getPropertyValue("RedlineIdentifier"), "author": r.getPropertyValue("RedlineAuthor"),
            "type": r.getPropertyValue("RedlineType"), "paragraph": ordinal(doc, ps, r.getPropertyValue("RedlineStart")), "object": r,
            "successor": {"author": succ["RedlineAuthor"], "type": succ["RedlineType"]} if succ else None})
    return out

def rule(rules, author, kind, paragraph):
    found = [r for r in rules if r["author"] == author and kind in r["types"] and paragraph in r["paragraphs"]]
    if len(found) != 1:
        raise Exception("no single unit for %s %s in paragraph %s" % (author, kind, paragraph))
    return found[0]["decision"]

def resolve(doc, dispatcher, frame, redline, decision):
    cursor = doc.Text.createTextCursorByRange(redline.getPropertyValue("RedlineStart"))
    cursor.gotoRange(redline.getPropertyValue("RedlineEnd"), True)
    doc.getCurrentController().select(cursor)
    command = ".uno:AcceptTrackedChange" if decision == "accept" else ".uno:RejectTrackedChange"
    dispatcher.executeDispatch(frame, command, "", 0, ())

def bold(p):
    out = []
    e = p.createEnumeration()
    while e.hasMoreElements():
        portion = e.nextElement()
        if portion.TextPortionType == "Text" and portion.getString() and portion.CharWeight >= 150:
            out.append(portion.getString())
    return out

# One set on a fresh load. Rejecting a delete-on-insert leaves the insertion
# under it, under the same identifier; it takes its own unit's decision.
# steps: how many pieces each decision resolved. then: a command run after.
def apply(desktop, dispatcher, path, decided, then=None):
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(path), "_blank", 0, (prop("Hidden", True),))
    try:
        frame = doc.getCurrentController().getFrame()
        listed = redlines(doc)
        for r in listed:
            r["decision"] = rule(decided["rules"], r["author"], "Format" if r["type"] in FORMATS else r["type"], r["paragraph"])
            if r["successor"]:
                r["exposed"] = rule(decided["rules"], r["successor"]["author"], r["successor"]["type"], r["paragraph"])
        steps = []
        for r in listed:
            if r["decision"] == "hold":
                continue
            current = redlines(doc)
            before = set(x["id"] for x in current)
            live = [x for x in current if x["id"] == r["id"]]
            if not live:
                steps.append(0)
                continue
            resolve(doc, dispatcher, frame, live[0]["object"], r["decision"])
            if r["successor"] and r["decision"] == "reject" and r["exposed"] != "hold":
                for exposed in [x for x in redlines(doc) if x["id"] == r["id"] or x["id"] not in before]:
                    resolve(doc, dispatcher, frame, exposed["object"], r["exposed"])
            steps.append(len(before - set(x["id"] for x in redlines(doc))))
        pending = len(redlines(doc))
        if then:
            dispatcher.executeDispatch(frame, then, "", 0, ())
        return {"paragraphs": [{"text": p.getString(), "bold": bold(p)} for p in paragraphs(doc)], "pending": pending, "steps": steps}
    finally:
        doc.close(True)

def run(*args):
    desktop = XSCRIPTCONTEXT.getDesktop()
    ctx = XSCRIPTCONTEXT.getComponentContext()
    dispatcher = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    path = os.environ["DOCHAUS_ORACLE_JOB"]
    result = {}
    try:
        job = json.load(open(path))
        for decided in job["sets"]:
            out = apply(desktop, dispatcher, job["input"], decided)
            if out["pending"]:
                out["accepted"] = apply(desktop, dispatcher, job["input"], decided, ".uno:AcceptAllTrackedChanges")["paragraphs"]
                out["rejected"] = apply(desktop, dispatcher, job["input"], decided, ".uno:RejectAllTrackedChanges")["paragraphs"]
            result[decided["name"]] = out
    except Exception:
        result["error"] = traceback.format_exc()
    finally:
        with open(path + ".out", "w") as f:
            json.dump(result, f)
        desktop.terminate()
`
