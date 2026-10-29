import { referenceEvent } from "./chronology-reference.fixture"
import { expect, test } from "bun:test"
import type { ChronologyProposal } from "../src/api/ingest/chronology"
import {
  chronologyChatContext,
  chronologyChatReference,
  addChronologyChatReference,
  withChronologyChatReferences,
  currentChronologyChatReferences,
} from "../src/components/chronology/model"

test("reference snapshots keep saved and proposed evidence with review and provenance qualifications", () => {
  const proposed = structuredClone(referenceEvent)
  proposed.title = "Proposed arrival date"
  proposed.accounts[0]!.evidence[0]!.documentName = "Carrier email.eml"
  const proposal: ChronologyProposal = {
    id: "date-proposal",
    kind: "amendment",
    targetIds: ["delivery"],
    baseEventHashes: {},
    events: [proposed],
    reason: "Check the newer account.",
    recommendation: "Inspect both accounts.",
    status: "pending",
    origin: "refresh",
    createdAt: 1,
  }
  const ref = chronologyChatReference("matter", referenceEvent, proposal)
  expect(ref.text).toBe(chronologyChatContext("matter", referenceEvent, proposal))
  expect(ref.sources).toHaveLength(2)
  expect(ref).toMatchObject({ reviewed: false, hidden: true, proposalId: "date-proposal" })
  expect(ref.qualifications).toContain("Conflict / dispute")
  expect(ref.sources[0]).toMatchObject({
    freshness: "stale",
    verified: true,
    provenance: "supplied-chronology",
    scope: "saved",
  })
  expect(ref.sources[1]!.scope).toBe("proposal")
  expect(JSON.stringify(ref.sources)).not.toContain("private-internal-version")
  referenceEvent.accounts[0]!.evidence[0]!.excerpt += " changed"
  expect(ref.sources[0]!.excerpt).not.toContain("changed")
  referenceEvent.accounts[0]!.evidence[0]!.excerpt = ref.sources[0]!.excerpt
})
test("references deduplicate by matter/event/proposal and model payload cannot include another matter", () => {
  const ref = chronologyChatReference("matter", referenceEvent),
    latest = { ...ref, title: "Latest snapshot", text: ref.text + "\nUpdated qualification." }
  const refs = addChronologyChatReference([ref], latest, "matter")
  expect(refs).toHaveLength(1)
  expect(refs[0]!.title).toBe("Latest snapshot")
  expect(addChronologyChatReference(refs, chronologyChatReference("elsewhere", referenceEvent), "matter")).toEqual(refs)
  const message = withChronologyChatReferences(
    "My specific question",
    [...refs, chronologyChatReference("elsewhere", referenceEvent)],
    "matter",
  )
  expect(message).toContain("My specific question")
  expect(message).toContain("Supplied chronology.md")
  expect(message).toContain("stale")
  expect(message).not.toContain("/matter/elsewhere")
})

test("current access refresh strips unavailable passages, drops obsolete references and retires decided proposals", () => {
  const saved = chronologyChatReference("matter", referenceEvent)
  const inaccessible = structuredClone(referenceEvent)
  inaccessible.accounts[0]!.evidence[0]!.excerpt = ""
  inaccessible.accounts[0]!.evidence[0]!.freshness = "unavailable"
  const refreshed = currentChronologyChatReferences([saved], { events: [inaccessible], proposals: [] }, "matter")
  expect(refreshed.changed).toBe(true)
  expect(refreshed.references[0]!.text).not.toContain("The carrier recorded delivery at the west gate.")
  expect(refreshed.references[0]!.sources[0]!.freshness).toBe("unavailable")
  expect(currentChronologyChatReferences([saved], { events: [], proposals: [] }, "matter")).toMatchObject({
    changed: true,
    removed: true,
    references: [],
  })
  const proposal: ChronologyProposal = {
    id: "proposal",
    kind: "amendment",
    targetIds: [referenceEvent.id],
    baseEventHashes: {},
    events: [referenceEvent],
    reason: "Proposed date",
    recommendation: "Inspect",
    status: "pending",
    origin: "refresh",
    createdAt: 1,
  }
  const proposed = chronologyChatReference("matter", referenceEvent, proposal)
  const decided = currentChronologyChatReferences(
    [proposed],
    { events: [referenceEvent], proposals: [{ ...proposal, status: "accepted" }] },
    "matter",
  )
  expect(decided).toMatchObject({ changed: true, proposalChanged: true })
  expect(decided.references[0]!.proposalId).toBeUndefined()
  expect(decided.references[0]!.text).not.toContain("Pending proposal")
})
