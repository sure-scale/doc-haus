import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import type { ChronologyAccount, ChronologyDate, ChronologyProposal } from "../../../dochaus/lib/chronology"
import { chronologyCounters } from "../../../dochaus/lib/chronology"
import { openDb, upsertDocument } from "./db"
import { liveText } from "../../../dochaus/lib/extract"
import {
  actOnChronology,
  batchChronologyProposals,
  chronologyEventHash,
  chronologyProposalDrafts,
  currentChronologySources,
  decideChronologyProposal,
  mutateChronology,
  newChronologyEvent,
  projectChronologySources,
  readChronology,
  stageChronologyProposals,
  verifyChronologyProposals,
  validChronologyDate,
} from "./chronology-store"
const dirs: string[] = []
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})
const date: ChronologyDate = { original: "6 May 2025", precision: "exact", start: "2025-05-06" }
function matter() {
  const dir = mkdtempSync(path.join(tmpdir(), "chronology-store-"))
  dirs.push(dir)
  writeFileSync(path.join(dir, "matter.json"), "{}")
  return dir
}
function addSource(
  dir: string,
  name = "Note.md",
  body = "The parcel was collected on 6 May 2025.\nThe recipient signed on 8 May 2025.",
) {
  const file = path.join(dir, name)
  writeFileSync(file, body)
  const db = openDb(dir)
  try {
    upsertDocument(db, file, name, Date.now(), null)
  } finally {
    db.close()
  }
  return currentChronologySources(dir).find((s) => s.documentName === name)!
}
async function account(dir: string): Promise<ChronologyAccount> {
  const source = currentChronologySources(dir).find((s) => s.documentName === "Note.md")!,
    canonical = await liveText(source.docPath),
    quote = "The parcel was collected on 6 May 2025.",
    start = canonical.indexOf(quote)
  return {
    id: "account-collection",
    title: "Parcel collected",
    detail: "The sender records collection.",
    date,
    evidence: [
      {
        id: "evidence-collection",
        documentName: "Note.md",
        sourceId: source.sourceId,
        version: source.version,
        excerpt: quote,
        section: "Note",
        charStart: start,
        charEnd: start + quote.length,
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
  }
}
function proposal(
  id: string,
  event: ReturnType<typeof newChronologyEvent>,
  target?: ReturnType<typeof newChronologyEvent>,
): ChronologyProposal {
  return {
    id,
    kind: target ? "amendment" : "addition",
    targetIds: target ? [target.id] : [],
    baseEventHashes: target ? { [target.id]: chronologyEventHash(target) } : {},
    events: [event],
    reason: "Additional source account.",
    recommendation: "Inspect its evidence.",
    status: "pending",
    origin: "refresh",
    createdAt: 1,
  }
}

test("opening clears New independently of review; lawyer choices survive reopening and stale writes fail", () => {
  const dir = matter()
  let state = mutateChronology(dir, 0, "seed", (s) => {
    s.created = true
    s.events.push(
      newChronologyEvent({ id: "event", title: "Manual event", detail: "No source supplied", date, origin: "manual" }),
    )
  })
  expect(chronologyCounters(state)).toMatchObject({ new: 1, unreviewed: 1 })
  state = actOnChronology(dir, state.revision, { type: "open", eventId: "event" }) as typeof state
  expect(chronologyCounters(state)).toMatchObject({ new: 0, unreviewed: 1 })
  for (const action of [
    { type: "note", eventId: "event", note: "Check envelope" },
    { type: "key", eventId: "event", key: true },
    { type: "dispute", eventId: "event", disputed: true, dispute: "Occurrence disputed" },
    { type: "review", eventId: "event", reviewed: true },
    { type: "hide", eventId: "event", hidden: true },
  ] as const)
    state = actOnChronology(dir, state.revision, action) as typeof state
  expect(readChronology(dir).events[0]).toMatchObject({
    reviewed: true,
    hidden: true,
    key: true,
    disputed: true,
    note: "Check envelope",
    origin: "manual",
    accounts: [],
  })
  expect(() => actOnChronology(dir, 0, { type: "note", eventId: "event", note: "Stale overwrite" })).toThrow("Reload")
})
test("invalid stored state stays intact instead of being replaced by an empty chronology", () => {
  const dir = matter()
  mutateChronology(dir, 0, "create", (s) => {
    s.created = true
  })
  const file = path.join(dir, ".dochaus", "chronology.json"),
    invalid = '{"schemaVersion":1,"revision":2,"events":null}'
  writeFileSync(file, invalid)
  expect(() => readChronology(dir)).toThrow("preserved")
  expect(() =>
    mutateChronology(dir, 2, "overwrite", (s) => {
      s.created = true
    }),
  ).toThrow("preserved")
  expect(readFileSync(file, "utf8")).toBe(invalid)
})
test("verified proposals cannot forge signoff; replay of stage and accept is idempotent", async () => {
  const dir = matter()
  addSource(dir)
  const a = await account(dir),
    event = newChronologyEvent({ title: a.title, detail: a.detail, date, accounts: [a] })
  const draft = {
    ...proposal("add-collection", event),
    events: [
      {
        title: event.title,
        detail: event.detail,
        date,
        accounts: [a],
        reviewed: true,
        hidden: true,
        key: true,
        note: "Forged signoff",
      },
    ],
  }
  let drafts = chronologyProposalDrafts([draft])
  await verifyChronologyProposals(dir, drafts)
  let state = stageChronologyProposals(dir, 0, drafts)
  expect(state.events).toHaveLength(0)
  expect(state.proposals[0]!.events[0]).toMatchObject({
    reviewed: false,
    hidden: false,
    key: false,
    note: "",
    lawyerEdited: false,
  })
  drafts = chronologyProposalDrafts([draft])
  await verifyChronologyProposals(dir, drafts)
  const before = state.revision
  state = stageChronologyProposals(dir, state.revision, drafts)
  expect(state.revision).toBe(before)
  state = decideChronologyProposal(dir, state.revision, "add-collection", "accept")
  expect(state.events[0]!.reviewed).toBe(true)
  expect(decideChronologyProposal(dir, 0, "add-collection", "accept").revision).toBe(state.revision)
  expect(readChronology(dir).events).toHaveLength(1)
})
test("freshness preserves historical review while unavailable cached evidence is redacted", async () => {
  const dir = matter(),
    source = addSource(dir),
    a = await account(dir)
  const originalBytes = readFileSync(source.docPath)
  mutateChronology(dir, 0, "seed", (s) => {
    s.events.push({ ...newChronologyEvent({ title: a.title, detail: a.detail, date, accounts: [a] }), reviewed: true })
  })
  writeFileSync(source.docPath, "Different source text")
  let projected = projectChronologySources(dir)
  expect(projected.events[0]!.reviewed).toBe(true)
  expect(projected.events[0]!.accounts[0]!.evidence[0]!.freshness).toBe("stale")
  rmSync(source.docPath)
  projected = projectChronologySources(dir)
  expect(projected.events[0]!.accounts[0]!.evidence[0]).toMatchObject({ freshness: "unavailable", excerpt: "" })
  expect(readChronology(dir).events[0]!.accounts[0]!.evidence[0]!.excerpt).toContain("parcel")
  mutateChronology(dir, readChronology(dir).revision, "unavailable after refresh", (state) => {
    state.events[0]!.accounts[0]!.evidence[0]!.freshness = "unavailable"
  })
  writeFileSync(source.docPath, originalBytes)
  projected = projectChronologySources(dir)
  expect(projected.events[0]!.accounts[0]!.evidence[0]!.freshness).toBe("current")
  expect(projected.events[0]!.reviewed).toBe(true)
})
test("stale event bases and changed source versions cannot apply; keep and defer preserve lawyer wording", async () => {
  const dir = matter(),
    source = addSource(dir),
    a = await account(dir)
  let state = mutateChronology(dir, 0, "seed", (s) => {
    s.events.push(newChronologyEvent({ id: "saved", title: a.title, detail: a.detail, date, accounts: [a] }))
  })
  const original = state.events[0]!,
    changed = { ...structuredClone(original), title: "AI correction" }
  state = stageChronologyProposals(dir, state.revision, [proposal("amend", changed, original)])
  state = actOnChronology(dir, state.revision, {
    type: "edit",
    eventId: "saved",
    title: "Lawyer wording",
    detail: "Corrected",
    date,
  }) as typeof state
  expect(() => decideChronologyProposal(dir, state.revision, "amend", "accept")).toThrow("current event changed")
  state = decideChronologyProposal(dir, state.revision, "amend", "defer")
  expect(state.events[0]!.title).toBe("Lawyer wording")
  state = decideChronologyProposal(dir, state.revision, "amend", "keep")
  state = stageChronologyProposals(dir, state.revision, [
    proposal("new", newChronologyEvent({ id: "new", title: "Another", detail: a.detail, date, accounts: [a] })),
  ])
  writeFileSync(source.docPath, "Changed before decision")
  expect(() => decideChronologyProposal(dir, state.revision, "new", "accept")).toThrow("evidence changed")
  expect(readChronology(dir).events[0]!.title).toBe("Lawyer wording")
})
test("batch additions stay unreviewed and competing accounts survive individual amendment", async () => {
  const dir = matter()
  addSource(dir)
  const a = await account(dir)
  let state = stageChronologyProposals(dir, 0, [
    proposal("batch", newChronologyEvent({ id: "saved", title: a.title, detail: a.detail, date, accounts: [a] })),
  ])
  state = batchChronologyProposals(dir, state.revision, ["batch"])
  expect(state.events[0]!.reviewed).toBe(false)
  const next = newChronologyEvent({
    id: "proposed",
    title: "Changed account",
    detail: "Alternative",
    date,
    accounts: [{ ...a, id: "other-account" }],
  })
  state = stageChronologyProposals(dir, state.revision, [proposal("amend", next, state.events[0]!)])
  expect(() => batchChronologyProposals(dir, state.revision, ["amend"])).toThrow("straightforward")
  state = decideChronologyProposal(dir, state.revision, "amend", "accept")
  expect(state.events[0]!.accounts).toHaveLength(2)
  const removal = { ...proposal("remove", next, state.events[0]!), kind: "removal" as const, events: [] }
  state = stageChronologyProposals(dir, state.revision, [removal])
  state = decideChronologyProposal(dir, state.revision, "remove", "accept")
  expect(state.events[0]!.hidden).toBe(true)
  expect(state.events).toHaveLength(1)
})
test("merge and split preview without writing, preserve accounts and notes and reset review", async () => {
  const dir = matter()
  addSource(dir)
  const a = await account(dir)
  let state = mutateChronology(dir, 0, "seed", (s) => {
    for (const id of ["one", "two"])
      s.events.push({
        ...newChronologyEvent({
          id,
          title: "Collection reported",
          detail: a.detail,
          date,
          accounts: [{ ...a, id: `a-${id}` }],
        }),
        reviewed: true,
        note: `Note ${id}`,
      })
  })
  const merge = { type: "merge" as const, eventIds: ["one", "two"], title: "Collection", detail: "Two accounts", date },
    preview = actOnChronology(dir, state.revision, { ...merge, preview: true })
  expect("preview" in preview && preview.preview[0]!.accounts).toHaveLength(2)
  expect(readChronology(dir).revision).toBe(state.revision)
  state = actOnChronology(dir, state.revision, merge) as typeof state
  const merged = state.events.find((e) => !e.hidden)!
  expect(merged).toMatchObject({ reviewed: false, note: "Note one\n\nNote two", previousIds: ["one", "two"] })
  state = actOnChronology(dir, state.revision, {
    type: "split",
    eventId: merged.id,
    parts: [
      { title: "One account", detail: "One", date, accountIds: ["a-one"] },
      { title: "Other account", detail: "Two", date, accountIds: ["a-two"] },
    ],
  }) as typeof state
  expect(state.events.filter((e) => !e.hidden)).toHaveLength(2)
  expect(state.events.filter((e) => !e.hidden).every((e) => !e.reviewed && e.note.includes("Note one"))).toBe(true)
})
test("source date corrections propose only affected relative dates without changing sources or saved events", async () => {
  const dir = matter()
  addSource(dir)
  const a = await account(dir),
    relative: ChronologyDate = {
      original: "three days after this note",
      precision: "relative",
      relative: {
        expression: "three days after",
        anchorDocument: "Note.md",
        offsetDays: 3,
        interpretation: "document-stated",
      },
    }
  let state = mutateChronology(dir, 0, "seed", (s) => {
    s.events.push(
      newChronologyEvent({
        id: "relative",
        title: "Response required",
        detail: "Scheduled",
        date: relative,
        accounts: [a],
      }),
      newChronologyEvent({ id: "exact", title: "Collection", detail: "Actual", date, accounts: [a] }),
    )
    s.events[0]!.reviewed = true
  })
  state = actOnChronology(dir, state.revision, {
    type: "source-date",
    documentName: "Note.md",
    date: "2025-05-06",
  }) as typeof state
  expect(state.proposals).toHaveLength(1)
  expect(state.proposals[0]!.events[0]!.date).toMatchObject({
    original: "three days after this note",
    start: "2025-05-09",
  })
  expect(state.events[0]!.date.start).toBeUndefined()
  expect(state.events[0]!.reviewed).toBe(true)
  expect(state.proposals[0]!.events[0]!.reviewed).toBe(false)
  const firstProposal = state.proposals[0]!.id
  state = actOnChronology(dir, state.revision, {
    type: "source-date",
    documentName: "Note.md",
    date: "2025-05-07",
  }) as typeof state
  expect(() => decideChronologyProposal(dir, state.revision, firstProposal, "accept")).toThrow("source date")
  expect(readChronology(dir).events[0]!.date.start).toBeUndefined()
  state = decideChronologyProposal(dir, state.revision, state.proposals[1]!.id, "accept")
  expect(state.events[0]!.date.start).toBe("2025-05-10")
  expect(readFileSync(path.join(dir, "Note.md"), "utf8")).toContain("parcel")
})
test("supplied chronology remains a source; derived deliverables and symlink aliases cannot corroborate", () => {
  const dir = matter()
  writeFileSync(path.join(dir, "Chronology.md"), "A source-only account.")
  let sources = currentChronologySources(dir)
  expect(sources[0]!.sourceId).toBe("supplied:Chronology.md")
  const outside = matter()
  writeFileSync(path.join(outside, "outside.md"), "Unauthorised")
  symlinkSync(path.join(outside, "outside.md"), path.join(dir, "outside.md"))
  const db = openDb(dir)
  try {
    db.run("INSERT INTO documents(doc_path,name,created_at,role) VALUES(?,?,?,'deliverable')", [
      path.join(dir, "Chronology.md"),
      "Chronology.md",
      1,
    ])
    db.run("INSERT INTO documents(doc_path,name,created_at,role) VALUES(?,?,?,'source')", [
      path.join(dir, "outside.md"),
      "outside.md",
      1,
    ])
  } finally {
    db.close()
  }
  sources = currentChronologySources(dir)
  expect(sources).toHaveLength(0)
  expect(readFileSync(path.join(dir, "Chronology.md"), "utf8")).toContain("source-only")
})
test("cross-matter source names and fabricated offsets fail exact verification", async () => {
  const dir = matter()
  addSource(dir)
  const a = await account(dir),
    p = proposal("unsafe", newChronologyEvent({ title: a.title, detail: a.detail, date, accounts: [a] }))
  p.events[0]!.accounts[0]!.evidence[0]!.documentName = "Other.md"
  await expect(verifyChronologyProposals(dir, [p])).rejects.toThrow("unavailable")
  p.events[0]!.accounts[0]!.evidence[0]!.documentName = "Note.md"
  p.events[0]!.accounts[0]!.evidence[0]!.charStart = 10
  await expect(verifyChronologyProposals(dir, [p])).rejects.toThrow("exact source passage")
  expect(readChronology(dir).events).toHaveLength(0)
})

test("refresh may preserve stale historical accounts while accepting newly verified evidence", async () => {
  const dir = matter(),
    source = addSource(dir),
    old = await account(dir)
  let state = mutateChronology(dir, 0, "seed", (s) => {
    s.events.push(newChronologyEvent({ id: "saved", title: old.title, detail: old.detail, date, accounts: [old] }))
  })
  writeFileSync(source.docPath, "The parcel was collected on 6 May 2025.\nAdditional testimony was recorded later.")
  const fresh = await account(dir)
  fresh.id = "new-source-version"
  const next = { ...structuredClone(state.events[0]!), accounts: [old, fresh] }
  const p = proposal("changed-source", next, state.events[0]!)
  await verifyChronologyProposals(dir, [p])
  state = stageChronologyProposals(dir, state.revision, [p])
  state = decideChronologyProposal(dir, state.revision, p.id, "accept")
  const projected = projectChronologySources(dir, state)
  expect(projected.events[0]!.accounts).toHaveLength(2)
  expect(projected.events[0]!.accounts[0]!.evidence[0]!.freshness).toBe("stale")
  expect(projected.events[0]!.accounts[1]!.evidence[0]!.freshness).toBe("current")
  expect(projected.events[0]!.reviewed).toBe(true)
})

test("removal decisions expire if a source is restored or added after the proposal", async () => {
  const dir = matter(),
    source = addSource(dir),
    a = await account(dir)
  let state = mutateChronology(dir, 0, "seed", (s) => {
    s.events.push(newChronologyEvent({ id: "saved", title: a.title, detail: a.detail, date, accounts: [a] }))
  })
  rmSync(source.docPath)
  const removal = {
    ...proposal("remove-withdrawn", state.events[0]!, state.events[0]!),
    kind: "removal" as const,
    events: [],
  }
  state = stageChronologyProposals(dir, state.revision, [removal])
  writeFileSync(source.docPath, "Restored source material")
  expect(() => decideChronologyProposal(dir, state.revision, removal.id, "accept")).toThrow("source set changed")
  expect(readChronology(dir).events[0]!.hidden).toBe(false)
})

test("a source date correction does not convert business-day or guessed relative formulas", async () => {
  const dir = matter()
  addSource(dir)
  const a = await account(dir)
  let state = mutateChronology(dir, 0, "seed", (s) => {
    for (const [id, expression] of [
      ["business", "3 business days after this note"],
      ["month", "one month after this note"],
    ])
      s.events.push(
        newChronologyEvent({
          id,
          title: "Response required",
          detail: "Scheduled",
          date: {
            original: expression,
            precision: "relative",
            relative: { expression, anchorDocument: "Note.md", offsetDays: 3, interpretation: "document-stated" },
          },
          accounts: [a],
        }),
      )
  })
  state = actOnChronology(dir, state.revision, {
    type: "source-date",
    documentName: "Note.md",
    date: "2025-05-06",
  }) as typeof state
  expect(state.proposals).toHaveLength(0)
})

test("chat account references preserve unavailable historical evidence without exposing or forging it", async () => {
  const dir = matter(),
    source = addSource(dir),
    a = await account(dir)
  let state = mutateChronology(dir, 0, "seed", (s) => {
    s.events.push(
      newChronologyEvent({ id: "target", title: a.title, detail: a.detail, date, accounts: [a] }),
      newChronologyEvent({
        id: "unrelated",
        title: "Other event",
        detail: a.detail,
        date,
        accounts: [{ ...a, id: "unrelated-account" }],
      }),
    )
  })
  rmSync(source.docPath)
  const content = {
    id: "split-by-reference",
    kind: "split",
    targetIds: ["target"],
    reason: "Separate two reported aspects for review.",
    recommendation: "Check the preserved historical support.",
    events: [
      { title: "First aspect", detail: "Historical account, source unavailable.", date, accountIds: [a.id] },
      {
        title: "Second aspect",
        detail: "Alternative aspect of the same historical account.",
        date,
        accountIds: [a.id],
      },
    ],
  }
  const proposals = chronologyProposalDrafts([content], state)
  await verifyChronologyProposals(dir, proposals)
  state = stageChronologyProposals(dir, state.revision, proposals)
  expect(projectChronologySources(dir, state).proposals[0]!.events[0]!.accounts[0]!.evidence[0]).toMatchObject({
    freshness: "unavailable",
    excerpt: "",
  })
  expect(state.proposals[0]!.events[0]!.accounts[0]!.evidence[0]!.excerpt).toContain("parcel")
  expect(() =>
    chronologyProposalDrafts(
      [{ ...content, events: [{ ...content.events[0], accountIds: ["unrelated-account"] }] }],
      state,
    ),
  ).toThrow("targeted events")
  expect(() =>
    chronologyProposalDrafts(
      [{ ...content, events: [{ ...content.events[0], accountIds: ["foreign-matter-account"] }] }],
      state,
    ),
  ).toThrow("targeted events")
  const forged = structuredClone(a)
  forged.id = "forged"
  forged.detail = "Changed historical account"
  forged.evidence[0]!.excerpt = "Invented source passage"
  const forgedProposal = chronologyProposalDrafts(
    [{ ...content, kind: "amendment", events: [{ ...content.events[0], accountIds: [], accounts: [forged] }] }],
    state,
  )
  await expect(verifyChronologyProposals(dir, forgedProposal)).rejects.toThrow("unavailable")
  state = decideChronologyProposal(dir, state.revision, proposals[0]!.id, "accept")
  expect(state.events.filter((e) => e.previousIds.includes("target"))).toHaveLength(2)
  expect(
    projectChronologySources(dir, state)
      .events.filter((e) => e.previousIds.includes("target"))
      .every((e) => e.accounts[0]!.evidence[0]!.excerpt === ""),
  ).toBe(true)
})

test("stored and proposed dates retain real precision and distinguish intervals from durations", () => {
  expect(validChronologyDate({ original: "May 2025", precision: "month", start: "2025-05" })).toBe(true)
  expect(validChronologyDate({ original: "2025-02-29", precision: "exact", start: "2025-02-29" })).toBe(false)
  expect(validChronologyDate({ original: "May 2025", precision: "month", start: "2025-05-01" })).toBe(false)
  expect(validChronologyDate({ original: "03/04/25", precision: "ambiguous", start: "2025-03-04" })).toBe(false)
  expect(validChronologyDate({ original: "Unknown", precision: "undated", start: "2025" })).toBe(false)
  expect(validChronologyDate({ original: "2025-13", precision: "month", start: "2025-13" })).toBe(false)
  expect(validChronologyDate({ original: "May to June", precision: "month", start: "2025-05", end: "2025-06" })).toBe(
    false,
  )
  expect(
    validChronologyDate({
      original: "May to June",
      precision: "month",
      start: "2025-05",
      end: "2025-06",
      rangeKind: "duration",
    }),
  ).toBe(true)
  expect(
    validChronologyDate({
      original: "Approximate May interval",
      precision: "approximate",
      start: "2025-05-06",
      end: "2025-05-03",
      rangeKind: "uncertainty",
    }),
  ).toBe(false)
})
