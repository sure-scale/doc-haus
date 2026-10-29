import { describe, expect, test } from "bun:test"
import type { ChronologyEvent, ChronologyProposal } from "../src/api/ingest/chronology"
import {
  chronologyChatContext,
  chronologyPeople,
  dateLabel,
  emptyFilters,
  matchesChronology,
  simpleAddition,
} from "../src/components/chronology/model"
const event: ChronologyEvent = {
  id: "collection",
  title: "Equipment collected",
  detail: "A carrier collected equipment; receipt is a separate occurrence.",
  date: { original: "May 2026", precision: "month", start: "2026-05" },
  accounts: [
    {
      id: "account",
      title: "Collection reported",
      detail: "",
      date: { original: "May 2026", precision: "month", start: "2026-05" },
      evidence: [
        {
          id: "evidence",
          documentName: "Carrier note.md",
          sourceId: "source-1",
          version: "v1",
          excerpt: "The sealed cases left the depot with the carrier.",
          section: "Collection",
          charStart: 21,
          charEnd: 74,
          verified: true,
          freshness: "current",
          provenance: "source",
        },
      ],
      privilege: "uncertain",
      significance: "background",
      people: ["Carrier"],
      status: "reported",
      gaps: [],
    },
  ],
  origin: "generated",
  reviewed: false,
  opened: false,
  hidden: false,
  key: false,
  disputed: false,
  note: "",
  lawyerEdited: false,
  createdAt: 1,
  updatedAt: 1,
  previousIds: [],
}
describe("chronology lawyer view", () => {
  test("date labels retain deadline qualifiers stored separately from the original date words", () => {
    expect(dateLabel({ original: "12 May 2026", precision: "exact", start: "2026-05-12", qualifier: "before" })).toBe(
      "before 12 May 2026",
    )
    expect(dateLabel({ original: "by 12 May 2026", precision: "exact", start: "2026-05-12", qualifier: "by" })).toBe(
      "by 12 May 2026",
    )
  })
  test("person filters join only supported current aliases and keep ambiguous identities separate", () => {
    const named = {
      ...event,
      accounts: [
        {
          ...event.accounts[0]!,
          people: ["M. Chen"],
          nameAssociations: [
            {
              name: "M. Chen",
              canonicalName: "Maya Chen",
              status: "supported" as const,
              evidenceIds: ["evidence"],
            },
          ],
        },
      ],
    }
    const other = { ...event, id: "other", accounts: [{ ...event.accounts[0]!, people: ["Maya Chen"] }] }
    const linked = chronologyPeople([named, other])
    expect(linked.options).toEqual(["Maya Chen"])
    expect(matchesChronology(named, "", { ...emptyFilters, person: "Maya Chen" }, linked.canonical)).toBe(true)
    const stale = {
      ...named,
      accounts: [
        { ...named.accounts[0]!, evidence: [{ ...event.accounts[0]!.evidence[0]!, freshness: "stale" as const }] },
      ],
    }
    expect(chronologyPeople([stale]).canonical("M. Chen")).toBe("M. Chen")
    const uncertain = {
      ...named,
      accounts: [
        {
          ...named.accounts[0]!,
          nameAssociations: [{ ...named.accounts[0]!.nameAssociations[0]!, status: "uncertain" as const }],
        },
      ],
    }
    expect(chronologyPeople([uncertain]).canonical("M. Chen")).toBe("M. Chen")
    const collision = {
      ...named,
      id: "collision",
      accounts: [
        {
          ...named.accounts[0]!,
          nameAssociations: [{ ...named.accounts[0]!.nameAssociations[0]!, canonicalName: "Morgan Chen" }],
        },
      ],
    }
    expect(chronologyPeople([named, collision]).canonical("M. Chen")).toBe("M. Chen")
  })
  test("date filters overlap partial dates and known periods without inventing event precision", () => {
    expect(matchesChronology(event, "", { ...emptyFilters, from: "2026-05-15", to: "2026-05-20" })).toBe(true)
    expect(matchesChronology(event, "", { ...emptyFilters, from: "2026-06-01" })).toBe(false)
    expect(
      matchesChronology(
        {
          ...event,
          date: {
            original: "April to June 2026",
            precision: "month",
            start: "2026-04",
            end: "2026-06",
            rangeKind: "duration",
          },
        },
        "",
        { ...emptyFilters, from: "2026-05-15", to: "2026-05-20" },
      ),
    ).toBe(true)
    expect(event.date.start).toBe("2026-05")
  })
  test("search includes supporting quotations and filters keep hidden status explicit", () => {
    expect(matchesChronology(event, "sealed cases", emptyFilters)).toBe(true)
    expect(matchesChronology({ ...event, hidden: true }, "sealed cases", emptyFilters)).toBe(false)
    expect(matchesChronology({ ...event, hidden: true }, "sealed cases", { ...emptyFilters, hidden: "include" })).toBe(
      true,
    )
    expect(matchesChronology(event, "", { ...emptyFilters, person: "Receiver" })).toBe(false)
    expect(matchesChronology(event, "", { ...emptyFilters, review: "new" })).toBe(true)
    expect(matchesChronology({ ...event, opened: true }, "", { ...emptyFilters, review: "unreviewed" })).toBe(true)
  })
  test("batch accepts only independent additions with current support", () => {
    const proposal: ChronologyProposal = {
      id: "new",
      kind: "addition",
      targetIds: [],
      baseEventHashes: {},
      events: [event],
      reason: "New source",
      recommendation: "Keep for review",
      status: "pending",
      origin: "refresh",
      createdAt: 1,
    }
    expect(simpleAddition(proposal)).toBe(true)
    expect(simpleAddition({ ...proposal, targetIds: ["saved"] })).toBe(false)
    expect(simpleAddition({ ...proposal, events: [{ ...event, disputed: true }] })).toBe(false)
    expect(simpleAddition({ ...proposal, events: [{ ...event, lawyerEdited: true }] })).toBe(false)
    expect(
      simpleAddition({
        ...proposal,
        events: [
          {
            ...event,
            accounts: [{ ...event.accounts[0], evidence: [{ ...event.accounts[0].evidence[0], freshness: "stale" }] }],
          },
        ],
      }),
    ).toBe(false)
  })
  test("chat context carries hidden/review/freshness and return link together", () => {
    const context = chronologyChatContext("matter-one", {
      ...event,
      hidden: true,
      reviewed: true,
      accounts: [
        {
          ...event.accounts[0],
          evidence: [{ ...event.accounts[0].evidence[0], freshness: "stale", provenance: "supplied-chronology" }],
        },
      ],
    })
    expect(context).toContain("reviewed; hidden")
    expect(context).toContain("stale; supplied-chronology")
    expect(context).toContain("view=chronology&event=collection")
    expect(context).toContain("The sealed cases left the depot")
    expect(context).toContain("do not apply changes directly")
  })
  test("labels preserve partial and ambiguous date wording", () => {
    expect(dateLabel(event.date)).toBe("May 2026")
    expect(dateLabel({ original: "06/05/26", precision: "ambiguous" })).toBe("06/05/26 (ambiguous date)")
    expect(
      dateLabel({ original: "at 3pm local time", precision: "exact", time: "15:00", timezone: "Europe/Lisbon" }),
    ).toContain("15:00 Europe/Lisbon")
  })
})
